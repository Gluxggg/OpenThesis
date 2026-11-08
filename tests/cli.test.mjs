import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync, writeFileSync, existsSync, mkdirSync, chmodSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve, delimiter } from 'node:path';
import { spawnSync } from 'node:child_process';
import { buildTestPdf } from './helpers/pdf.mjs';

const cli = resolve('packages/cli/dist/index.js');

function run(args, cwd = process.cwd(), env = process.env) {
  return spawnSync(process.execPath, [cli, ...args], { cwd, encoding: 'utf8', env });
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
  assert.match(result.stderr, /must be a \.docx or \.pdf file/);
  assert.equal(readFileSync(input, 'utf8'), 'keep me');
});

test('parse infers a template JSON file from a text-based PDF', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const input = join(dir, 'template.pdf');
  const output = join(dir, 'template.template.json');
  writeFileSync(input, buildTestPdf());
  const result = run(['parse', input, '--org', 'PDF University']);
  assert.equal(result.status, 0, result.stderr);
  const template = JSON.parse(readFileSync(output, 'utf8'));
  assert.equal(template.meta.sourceFormat, 'pdf');
  assert.equal(template.meta.organization, 'PDF University');
  assert.ok(Object.keys(template.styles).length >= 2);

  const content = join(dir, 'content.json');
  const document = join(dir, 'rendered.docx');
  writeFileSync(content, JSON.stringify({
    type: 'thesis', meta: { title: 'PDF template test' }, cover: [],
    sections: [{ id: 'intro', type: 'chapter', title: 'Introduction', content: [
      { type: 'paragraph', text: 'Rendered with a PDF-derived template.' },
    ] }],
  }));
  const build = run(['build', content, '-t', output, '-o', document]);
  assert.equal(build.status, 0, build.stderr);
  assert.equal(readFileSync(document).subarray(0, 2).toString('ascii'), 'PK');
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

test('init rejects the flags it does not read instead of writing a sample anyway', () => {
  // This is the audit's repro: the flag table was global, so `init` accepted
  // `--org` and `-t`, ignored both, exited 0 and wrote thesis-content.json.
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const result = run(['init', '--org', 'Acme', '-t', 'missing.json'], dir);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /Unknown option: --org/);
  // The error names the command and what it does take, so the fix is obvious.
  assert.match(result.stderr, /`thesis init` accepts: --type, --force, -y/);
  assert.equal(existsSync(join(dir, 'thesis-content.json')), false);
});

test('each command rejects a flag that belongs to another command', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  writeFileSync(join(dir, 'paper.md'), '# Title\n\nBody.');
  writeFileSync(join(dir, 'content.json'), thesisContent('Cross'));
  const cases = [
    { args: ['build', 'content.json', '--org', 'Acme'], flag: '--org', writes: 'content.docx' },
    { args: ['import', 'paper.md', '-t', 'template.json'], flag: '-t', writes: 'paper.json' },
    { args: ['parse', 'template.docx', '-o', 'out.json'], flag: '-o', writes: 'template.template.json' },
  ];
  for (const { args, flag, writes } of cases) {
    const result = run(args, dir);
    assert.equal(result.status, 1, `${args[0]} must reject ${flag}: ${result.stderr}`);
    assert.match(result.stderr, new RegExp(`Unknown option: ${flag}`));
    assert.match(result.stderr, new RegExp(`\`thesis ${args[0]}\` accepts:`));
    assert.equal(existsSync(join(dir, writes)), false, `${args[0]} wrote ${writes} despite the rejection`);
  }
});

test('the global options work for commands other than build', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));

  // `--force`/`-y` used to be documented under "Options (build)", but every
  // command runs the same overwrite guard.
  assert.equal(run(['init'], dir).status, 0);
  const refused = run(['init'], dir);
  assert.equal(refused.status, 1);
  assert.match(refused.stderr, /Output already exists/);
  assert.equal(run(['init', '--force'], dir).status, 0);
  assert.equal(run(['init', '-y'], dir).status, 0);

  const md = join(dir, 'paper.md');
  const json = join(dir, 'paper.json');
  writeFileSync(md, '# Title\n\nBody.');
  assert.equal(run(['import', md, '-o', json], dir).status, 0);
  assert.equal(run(['import', md, '-o', json], dir).status, 1);
  assert.equal(run(['import', md, '-o', json, '--force'], dir).status, 0);
});

test('help lists --force as a global option, not a build-only one', () => {
  const result = run(['--help']);
  assert.equal(result.status, 0);
  const help = result.stdout;
  const globalStart = help.indexOf('Global options');
  const perCommandStart = help.indexOf('Options by command');
  assert.notEqual(globalStart, -1, 'help must have a global options section');
  assert.notEqual(perCommandStart, -1, 'help must list options per command');
  const force = help.indexOf('--force, -y');
  assert.ok(force > globalStart && force < perCommandStart, '--force belongs to the global section');
});

test('--help is honoured after a command name', () => {
  // `build --help` used to treat the flag as a content path.
  const result = run(['build', '--help']);
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /Options by command/);
});

test('build reports the type the content declares', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const input = join(dir, 'journal.json');
  writeFileSync(input, JSON.stringify({
    type: 'journal',
    meta: { title: 'Label', authors: [{ name: 'A', affiliations: [{ institution: 'U' }] }] },
    sections: [{
      id: 'intro', type: 'introduction', title: '1 Introduction',
      content: [{ type: 'paragraph', text: 'Body.' }],
    }],
  }));

  const result = run(['build', input, '-o', join(dir, 'out.docx')]);
  assert.equal(result.status, 0, result.stderr);
  // The removed ternary chain relabelled 'journal' as 'journal article'.
  assert.match(result.stdout, /✅ journal built successfully!/);
});

test('build does not guess a legacy document type from its title', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const input = join(dir, 'legacy.json');
  // Legacy JSON carries no `type` field; the removed chain read 期刊 out of the
  // title and reported the document as a "journal article".
  writeFileSync(input, JSON.stringify({ title: '期刊投稿说明', cover_blocks: [], body_blocks: [] }));

  const result = run(['build', input, '-o', join(dir, 'out.docx')]);
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /document built successfully!/);
  assert.doesNotMatch(result.stdout, /journal article/);
});

function ocrInstalled() {
  const lookup = process.platform === 'win32' ? 'where' : 'which';
  return spawnSync(lookup, ['ocrmypdf'], { encoding: 'utf8' }).status === 0;
}

test('a scanned PDF explains how to run OCR when no OCR tool is installed', (t) => {
  if (ocrInstalled()) {
    t.skip('ocrmypdf is installed on this machine');
    return;
  }
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const input = join(dir, 'scanned.pdf');
  writeFileSync(input, buildTestPdf({ text: false }));

  const result = run(['parse', input]);
  assert.equal(result.status, 1);
  assert.match(result.stderr, /looks scanned/);
  assert.match(result.stderr, /OCRmyPDF/);
});

test('a scanned PDF is OCRed with ocrmypdf when the tool is on the path', () => {
  const dir = mkdtempSync(join(tmpdir(), 'openthesis-cli-'));
  const binDir = join(dir, 'bin');
  mkdirSync(binDir);
  const fake = join(dir, 'fake-ocr.mjs');
  // The fake stands in for ocrmypdf: `--version` succeeds, and a real run
  // writes a PDF with a text layer to the output path (its last argument).
  writeFileSync(fake, [
    "import { writeFileSync } from 'node:fs';",
    'const args = process.argv.slice(2);',
    "if (!args.includes('--skip-text')) process.exit(0);",
    'const out = args[args.length - 1];',
    `writeFileSync(out, Buffer.from('${buildTestPdf().toString('base64')}', 'base64'));`,
    '',
  ].join('\n'));
  if (process.platform === 'win32') {
    writeFileSync(join(binDir, 'ocrmypdf.cmd'), `@echo off\r\n"${process.execPath}" "${fake}" %*\r\n`);
  } else {
    const script = join(binDir, 'ocrmypdf');
    writeFileSync(script, `#!/bin/sh\nexec "${process.execPath}" "${fake}" "$@"\n`);
    chmodSync(script, 0o755);
  }

  const input = join(dir, 'scanned.pdf');
  writeFileSync(input, buildTestPdf({ text: false }));
  const env = { ...process.env, PATH: `${binDir}${delimiter}${process.env.PATH ?? ''}` };
  const result = run(['parse', input], dir, env);
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /running local OCR/);
  const template = JSON.parse(readFileSync(join(dir, 'scanned.template.json'), 'utf8'));
  assert.equal(template.meta.sourceFormat, 'pdf');
});
