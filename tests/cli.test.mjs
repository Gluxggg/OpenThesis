import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync, writeFileSync, existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';

const cli = resolve('packages/cli/dist/index.js');

function run(args, cwd = process.cwd()) {
  return spawnSync(process.execPath, [cli, ...args], { cwd, encoding: 'utf8' });
}

test('help exits successfully', () => {
  const result = run(['--help']);
  assert.equal(result.status, 0);
  assert.match(result.stdout, /OpenThesis/);
});

test('accepts the documented `thesis <command>` prefix', () => {
  // The README and help text document `thesis build …` (the bin's name), while
  // the agent skill invokes the CLI directly. Both must work.
  const prefixed = run(['thesis', '--help']);
  assert.equal(prefixed.status, 0);
  assert.match(prefixed.stdout, /OpenThesis/);

  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const input = join(dir, 'paper.md');
  const output = join(dir, 'paper.json');
  writeFileSync(input, '# Title\n\nBody.');
  const result = run(['thesis', 'import', input, '-o', output]);
  assert.equal(result.status, 0, result.stderr);
  assert.equal(JSON.parse(readFileSync(output, 'utf8')).type, 'thesis');
});

test('parse rejects non-DOCX input before it can be overwritten', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const input = join(dir, 'template.txt');
  writeFileSync(input, 'keep me');
  const result = run(['parse', input]);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /must be a \.docx file/);
  assert.equal(readFileSync(input, 'utf8'), 'keep me');
});

test('build appends .docx for an extensionless content file', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const input = join(dir, 'content');
  const content = {
    type: 'thesis', meta: { title: 'CLI test' }, cover: [],
    sections: [{ id: 'intro', type: 'chapter', title: 'Intro', content: [] }],
  };
  writeFileSync(input, JSON.stringify(content));
  const result = run(['build', input]);
  assert.equal(result.status, 0, result.stderr);
  assert.doesNotThrow(() => JSON.parse(readFileSync(input, 'utf8')));
  assert.equal(existsSync(`${input}.docx`), true);
});

test('missing option values produce a clear error', () => {
  const result = run(['build', 'content.json', '-o']);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /Missing value for -o/);
});

test('import converts Markdown into structured JSON', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const input = join(dir, 'paper.md');
  const output = join(dir, 'paper.json');
  writeFileSync(input, `---\ntype: journal\ntitle: CLI Markdown\nauthors: [Agent]\n---\n## Methods\nReproducible method.`);
  const result = run(['import', input, '-o', output]);
  assert.equal(result.status, 0, result.stderr);
  const document = JSON.parse(readFileSync(output, 'utf8'));
  assert.equal(document.type, 'journal');
  assert.equal(document.sections[0].title, 'Methods');
});

test('build renders Markdown directly to DOCX', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const input = join(dir, 'thesis.md');
  const output = join(dir, 'thesis.docx');
  writeFileSync(input, `---\ntitle: Markdown Build\nauthor: Test Agent\n---\n## Introduction\nDirect build works.`);
  const result = run(['build', input, '-o', output]);
  assert.equal(result.status, 0, result.stderr);
  assert.equal(existsSync(output), true);
  assert.equal(readFileSync(output).subarray(0, 2).toString('ascii'), 'PK');
});

test('accepts the --flag=value form instead of silently ignoring it', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const input = join(dir, 'paper.md');
  const output = join(dir, 'paper.json');
  writeFileSync(input, '# Title\n\nBody text.');
  const result = run(['import', input, '--type=official', '-o', output]);
  assert.equal(result.status, 0, result.stderr);
  assert.equal(JSON.parse(readFileSync(output, 'utf8')).type, 'official');
});

test('refuses an output path that differs from the input only by case', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const input = join(dir, 'paper.md');
  const source = '# Title\n\nBody text.';
  writeFileSync(input, source);
  // Windows and macOS filesystems are case-insensitive, so this used to
  // overwrite the Markdown source with the JSON it had just produced.
  const result = run(['import', input, '-o', join(dir, 'paper.MD')]);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /must be different from the Markdown input path/);
  assert.equal(readFileSync(input, 'utf8'), source);
});

test('rejects unknown options instead of ignoring them', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const input = join(dir, 'content.json');
  writeFileSync(input, JSON.stringify({
    type: 'thesis', meta: { title: 't' }, cover: [],
    sections: [{ id: 'a', type: 'chapter', title: 'A', content: [] }],
  }));
  const result = run(['build', input, '--out', 'elsewhere.docx']);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /Unknown option: --out/);
  assert.equal(existsSync(join(dir, 'content.docx')), false);
});

test('reads a BOM-prefixed JSON content file', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const input = join(dir, 'content.json');
  const content = {
    type: 'thesis', meta: { title: 'BOM' }, cover: [],
    sections: [{ id: 'a', type: 'chapter', title: 'A', content: [] }],
  };
  writeFileSync(input, `\uFEFF${JSON.stringify(content)}`, 'utf8');
  const result = run(['build', input, '-o', join(dir, 'out.docx')]);
  assert.equal(result.status, 0, result.stderr);
  assert.equal(existsSync(join(dir, 'out.docx')), true);
});

test('init honours the --type= flag form', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const result = run(['init', '--type=journal'], dir);
  assert.equal(result.status, 0, result.stderr);
  assert.equal(existsSync(join(dir, 'journal-content.json')), true);
  assert.equal(existsSync(join(dir, 'thesis-content.json')), false);
});

test('build rejects invalid structured content, naming the offending path', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const input = join(dir, 'content.json');
  const output = join(dir, 'out.docx');
  // `ThesisSection.type` is required by the schema the docs point at. The
  // renderer ignores it and used to render this happily, so nothing caught it.
  writeFileSync(input, JSON.stringify({
    type: 'thesis', meta: { title: 'Bad' }, cover: [],
    sections: [{ id: 'a', title: 'A', content: [] }],
  }));

  const result = run(['build', input, '-o', output]);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /Invalid document content — 1 problem/);
  assert.match(result.stderr, /sections\[0\]\.type/);
  // Validation runs before rendering, so a rejected document leaves nothing behind.
  assert.equal(existsSync(output), false);
});

test('build rejects invalid legacy content too', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const input = join(dir, 'legacy.json');
  const output = join(dir, 'out.docx');
  // The renderer padded and truncated mismatched rows silently, which lost data.
  writeFileSync(input, JSON.stringify({
    title: 'Legacy',
    cover_blocks: [],
    body_blocks: [{ type: 'table', caption: 'C', headers: ['A'], data: [['1', '2']] }],
  }));

  const result = run(['build', input, '-o', output]);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /body_blocks\[0\]\.data\[0\]/);
  assert.match(result.stderr, /has 2 cells but the table declares 1 column/);
  assert.equal(existsSync(output), false);
});

/** Same as `run`, but keeps stdout as bytes: a DOCX must not be decoded. */
function runBuffer(args, input) {
  return spawnSync(process.execPath, [cli, ...args], {
    cwd: process.cwd(),
    ...(input === undefined ? {} : { input }),
  });
}

const thesisContent = title => JSON.stringify({
  type: 'thesis', meta: { title }, cover: [],
  sections: [{ id: 'a', type: 'chapter', title: 'A', content: [{ type: 'paragraph', text: 'Body.' }] }],
});

test('--version prints the packaged version', () => {
  for (const flag of ['--version', '-v']) {
    const result = run([flag]);
    assert.equal(result.status, 0, result.stderr);
    assert.match(result.stdout.trim(), /^\d+\.\d+\.\d+$/);
  }
});

test('creates a missing output directory instead of failing with ENOENT', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const input = join(dir, 'content.json');
  writeFileSync(input, thesisContent('Deep'));
  const output = join(dir, 'deeply', 'nested', 'out.docx');

  const result = run(['build', input, '-o', output]);
  assert.equal(result.status, 0, result.stderr);
  assert.equal(existsSync(output), true);
});

test('refuses to overwrite an existing output without --force', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const input = join(dir, 'content.json');
  writeFileSync(input, thesisContent('Once'));
  const output = join(dir, 'out.docx');

  assert.equal(run(['build', input, '-o', output]).status, 0, 'first build');
  const first = readFileSync(output);

  // Running the same command twice used to replace the document in silence.
  const second = run(['build', input, '-o', output]);
  assert.equal(second.status, 1);
  assert.match(second.stderr, /Output already exists/);
  assert.match(second.stderr, /--force/);
  assert.deepEqual(readFileSync(output), first, 'the first document must survive');

  assert.equal(run(['build', input, '-o', output, '--force']).status, 0);
});

test('reads the input from stdin', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const output = join(dir, 'piped.docx');
  const result = runBuffer(['build', '-', '--type', 'thesis', '-o', output],
    Buffer.from('---\ntitle: Piped\n---\n## Section\nBody text.'));

  assert.equal(result.status, 0, result.stderr.toString());
  assert.equal(existsSync(output), true);
});

test('writes the document to stdout without corrupting it', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const input = join(dir, 'content.json');
  writeFileSync(input, thesisContent('Pipe'));

  const result = runBuffer(['build', input, '-o', '-']);
  assert.equal(result.status, 0, result.stderr.toString());
  // A DOCX is a ZIP: one progress line on stdout would move the signature.
  assert.equal(result.stdout.subarray(0, 2).toString('ascii'), 'PK');
  assert.ok(result.stdout.length > 5000, `expected a document, received ${result.stdout.length} bytes`);
  assert.match(result.stderr.toString(), /Output:/, 'progress must move to stderr');
});
