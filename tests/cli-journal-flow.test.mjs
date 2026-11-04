/**
 * The whole submission flow, end to end, through the CLI.
 *
 * The library-level tests check pieces. This one checks the thing a person
 * actually does: parse a publisher's template, write content, build, and get a
 * document back. If that chain breaks, every unit test can still be green.
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';

const CLI = resolve('packages/cli/dist/index.js');
const TEMPLATE = resolve('tests/fixtures/templates/acm-journal.docx');

/** Run the CLI and return its combined output, throwing on a non-zero exit. */
function cli(args, cwd) {
  return execFileSync(process.execPath, [CLI, ...args], { cwd, encoding: 'utf8' });
}

test('parse a publisher template, then build a submission from it', () => {
  const work = mkdtempSync(join(tmpdir(), 'openthesis-journal-flow-'));
  // `parse` writes the template JSON beside the source, named after it.
  writeFileSync(join(work, 'acm.docx'), readFileSync(TEMPLATE));

  const parseOutput = cli(['parse', 'acm.docx', '--type', 'journal', '--org', 'ACM'], work);
  assert.match(parseOutput, /acm\.template\.json/);

  const templatePath = join(work, 'acm.template.json');
  assert.ok(existsSync(templatePath), 'parse should write a .template.json');

  const template = JSON.parse(readFileSync(templatePath, 'utf8'));
  assert.equal(template.meta.documentType, 'journal');
  // The ACM template declares named styles, and they must bind.
  assert.ok(template.roleWinners, 'the template should record a style for each role');
  for (const role of ['journal_title', 'journal_author', 'journal_abstract', 'journal_keywords', 'figure_caption']) {
    assert.ok(template.roleWinners[role], `role ${role} should bind to a template style`);
  }

  const content = {
    type: 'journal',
    meta: {
      title: 'A template-driven pipeline',
      authors: [{ name: 'San Zhang', affiliations: [{ institution: 'Example University' }] }],
      language: 'en',
      abstract: 'An abstract.',
      keywords: ['template', 'automation', 'typesetting'],
    },
    sections: [
      { id: 'intro', type: 'introduction', title: 'Introduction', content: [
        { type: 'paragraph', text: 'Text [[引用#r1]].' },
      ] },
      { id: 'method', type: 'methods', title: 'Method', content: [
        { type: 'table', caption: 'Stages', headers: ['Stage', 'Output'], data: [['Parse', 'JSON']] },
        { type: 'equation_numbered', latex: 'P = \\frac{a}{b}' },
      ] },
    ],
    backMatter: { references: [
      { id: 'r1', type: 'journal', language: 'en', authors: ['Mary Smith'], title: 'T', container: 'J', year: 2024 },
    ] },
  };
  writeFileSync(join(work, 'manuscript.json'), JSON.stringify(content), 'utf8');

  const buildOutput = cli(['build', 'manuscript.json', '-t', 'acm.template.json', '-o', 'manuscript.docx'], work);
  assert.match(buildOutput, /journal built successfully/);

  const docx = join(work, 'manuscript.docx');
  assert.ok(existsSync(docx), 'build should write the document');
  assert.ok(readFileSync(docx).length > 5000, 'the document should have content in it');
});

test('the CLI refuses an invalid submission before it builds anything', () => {
  // The same validator the library uses, with the same path-shaped message.
  const work = mkdtempSync(join(tmpdir(), 'openthesis-journal-bad-'));
  writeFileSync(join(work, 'manuscript.json'), JSON.stringify({
    type: 'journal',
    meta: { title: 'T', authors: [{ name: 'A' }] },
    sections: [],
  }), 'utf8');

  let failed = false;
  let output = '';
  try {
    cli(['build', 'manuscript.json', '-o', 'out.docx'], work);
  } catch (error) {
    failed = true;
    output = String(error.stdout ?? '') + String(error.stderr ?? '');
  }
  assert.ok(failed, 'an author with no affiliations should be refused');
  assert.match(output, /meta\.authors\[0\]\.affiliations/, 'the message should name the field');
});
