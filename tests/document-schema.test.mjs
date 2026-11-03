import test from 'node:test';
import assert from 'node:assert/strict';
import { readdirSync, readFileSync } from 'node:fs';
import { validateDocument, validateLegacyDocument, assertValidDocument, formatIssues } from '../packages/document-schema/dist/index.js';

const validThesis = () => ({
  type: 'thesis',
  meta: { title: 'A thesis' },
  cover: [{ type: 'centered_text', text: 'Cover' }],
  sections: [{
    id: 'intro', type: 'chapter', title: 'Introduction',
    content: [
      { type: 'paragraph', text: 'Body.' },
      { type: 'table', caption: 'Results', headers: ['A', 'B'], data: [['1', '2']] },
    ],
    subsections: [{ id: 'sub', type: 'chapter', title: 'Sub', content: [] }],
  }],
});

const validJournal = () => ({
  type: 'journal',
  meta: {
    title: 'A paper',
    authors: [{ name: 'A', affiliations: [{ institution: 'Somewhere' }] }],
    keywords: ['one'],
  },
  sections: [{ id: 'intro', type: 'introduction', title: 'Intro', content: [] }],
  backMatter: { references: [{ id: '1', text: 'A reference.' }] },
});

const validOfficial = () => ({
  type: 'official',
  meta: {
    title: 'A notice', issuingAuthority: 'Somewhere', documentNumber: 'X〔2026〕1号',
    documentCategory: '通知', primaryRecipients: ['各单位'], date: '2026年1月1日',
  },
  body: [{ type: 'paragraph', text: 'Body.' }],
});

test('accepts well-formed documents of every type', () => {
  assert.deepEqual(validateDocument(validThesis()), []);
  assert.deepEqual(validateDocument(validJournal()), []);
  assert.deepEqual(validateDocument(validOfficial()), []);
});

/** Examples the documentation points a reader at. They must stay loadable. */
const SHIPPED_EXAMPLES = ['sample-thesis.json', 'test-image.json'];

test('every shipped example validates under its own validator', () => {
  // Nothing in the repository referenced examples/test-image.json, so when the
  // validator started requiring `section.type` the file went on failing
  // `build` with no test to notice. Walk the directory instead of naming one
  // file, so a new example is covered the moment it lands.
  const present = readdirSync('examples').filter(name => name.endsWith('.json'));
  for (const name of SHIPPED_EXAMPLES) {
    assert.ok(present.includes(name), `examples/${name} is a documented example and must exist`);
  }

  for (const name of present) {
    const example = JSON.parse(readFileSync(`examples/${name}`, 'utf8'));
    const legacy = !('type' in example);
    const issues = legacy ? validateLegacyDocument(example) : validateDocument(example);
    assert.deepEqual(issues, [], `examples/${name} must validate as ${legacy ? 'legacy content' : example.type}`);
  }
});

test('validates the legacy format too', () => {
  assert.match(validateLegacyDocument(null)[0].message, /expected a document object/);
  assert.equal(validateLegacyDocument({ cover_blocks: [], body_blocks: [] })[0].path, 'title');

  const broken = {
    title: 'T',
    cover_blocks: [],
    body_blocks: [{ type: 'table', caption: 'C', headers: ['A'], data: [['1', '2']] }],
  };
  assert.match(validateLegacyDocument(broken)[0].message, /has 2 cells but the table declares 1 column/);
});

test('rejects values that are not documents', () => {
  assert.match(validateDocument(null)[0].message, /expected a document object/);
  assert.match(validateDocument('nope')[0].message, /expected a document object/);
  assert.equal(validateDocument({ type: 'slides' })[0].path, 'type');
});

test('reports every problem at once, each with its path', () => {
  const broken = validThesis();
  broken.meta.title = '';
  broken.sections[0].content[0].text = '';
  broken.sections[0].id = '';

  const issues = validateDocument(broken);
  const paths = issues.map(issue => issue.path);
  // A single-shot validator would stop at the first problem; the point is that
  // the author sees the whole list.
  assert.deepEqual(paths.sort(), ['meta.title', 'sections[0].content[0].text', 'sections[0].id']);
  assert.match(issues.find(i => i.path === 'meta.title').message, /an empty string/);
});

test('rejects an unknown block type and names the valid ones', () => {
  const broken = validThesis();
  broken.sections[0].content.push({ type: 'sparkle', text: 'x' });
  const issue = validateDocument(broken).find(i => i.path.endsWith('.type'));
  assert.match(issue.path, /content\[2\]\.type/);
  assert.match(issue.message, /received "sparkle"/);
  assert.match(issue.message, /"paragraph"/);
});

test('catches table rows that do not match the declared columns', () => {
  const broken = validThesis();
  broken.sections[0].content[1].data = [['1'], ['1', '2', '3']];
  const issues = validateDocument(broken);
  // The renderer pads and truncates silently, so this used to lose data quietly.
  assert.equal(issues.length, 2);
  assert.match(issues[0].message, /has 1 cell but the table declares 2 columns/);
  assert.match(issues[1].message, /has 3 cells but the table declares 2 columns/);
});

test('catches a table with no columns and mismatched column widths', () => {
  const noHeaders = validThesis();
  noHeaders.sections[0].content[1].headers = [];
  assert.match(validateDocument(noHeaders)[0].message, /at least one column/);

  const badWidths = validThesis();
  badWidths.sections[0].content[1].columnWidths = [100];
  assert.match(validateDocument(badWidths)[0].message, /1 entries but the table declares 2 columns/);

  const negativeWidth = validThesis();
  negativeWidth.sections[0].content[1].columnWidths = [100, -5];
  assert.match(validateDocument(negativeWidth)[0].message, /positive number/);
});

test('enforces the GB/T 9704 required fields on official documents', () => {
  const broken = validOfficial();
  broken.meta.documentCategory = '总结';
  broken.meta.primaryRecipients = [];
  const paths = validateDocument(broken).map(issue => issue.path);
  assert.ok(paths.includes('meta.documentCategory'));
  assert.ok(paths.includes('meta.primaryRecipients'));

  const missing = validOfficial();
  delete missing.meta.issuingAuthority;
  assert.match(validateDocument(missing)[0].message, /must be a non-empty string/);
});

test('validates journal authors and section types', () => {
  const noAuthors = validJournal();
  noAuthors.meta.authors = [];
  assert.match(validateDocument(noAuthors)[0].message, /at least one author/);

  const badAffiliation = validJournal();
  badAffiliation.meta.authors[0].affiliations = [{}];
  assert.match(validateDocument(badAffiliation)[0].path, /affiliations\[0\]\.institution/);

  const badSection = validJournal();
  badSection.sections[0].type = 'chapter';
  assert.match(validateDocument(badSection)[0].message, /received "chapter"/);
});

test('walks nested subsections', () => {
  const broken = validThesis();
  broken.sections[0].subsections[0].id = '';
  const issue = validateDocument(broken)[0];
  assert.equal(issue.path, 'sections[0].subsections[0].id');
});

test('assertValidDocument throws one message listing every issue', () => {
  const broken = validThesis();
  broken.meta.title = '';
  broken.sections[0].id = '';

  assert.throws(() => assertValidDocument(broken), error => {
    assert.match(error.message, /2 problems/);
    assert.match(error.message, /meta\.title/);
    assert.match(error.message, /sections\[0\]\.id/);
    return true;
  });

  // A valid document passes through untouched.
  assert.doesNotThrow(() => assertValidDocument(validThesis()));
});

test('formatIssues renders paths for a CLI error message', () => {
  const text = formatIssues([{ path: 'sections[0].id', message: 'must be a non-empty string' }]);
  assert.equal(text, '  sections[0].id: must be a non-empty string');
});

test('rejects a degree outside the schema vocabulary', () => {
  // `degree: PhD` used to reach the JSON as `undefined`, which JSON.stringify
  // then removed — the key vanished rather than being reported.
  const broken = validThesis();
  broken.meta.degree = 'PhD';
  assert.equal(validateDocument(broken)[0].path, 'meta.degree');
  assert.match(validateDocument(broken)[0].message, /received "PhD"/);

  const good = validThesis();
  good.meta.degree = 'doctor';
  assert.deepEqual(validateDocument(good), []);
});
