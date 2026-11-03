#!/usr/bin/env node
// ============================================================
// @openthesis/cli — OpenThesis Command Line Interface
// ============================================================
// Usage:
//   thesis parse <template.docx> [--type thesis|journal|official]
//   thesis import <content.md> [--type thesis|journal|official] [-o <content.json>]
//   thesis build <content.json|content.md> [-t <template.json>] [-o <output.docx>]
//   thesis init [--type thesis|journal|official]
// ============================================================

import { readFileSync, writeFileSync, realpathSync, mkdirSync, existsSync } from 'fs';
import { resolve, dirname, extname, basename } from 'path';
import { parseTemplate } from '@openthesis/template-parser';
import { parseMarkdown } from '@openthesis/markdown-parser';
import { renderLegacy, createUSTBTemplate, renderDocument } from '@openthesis/docx-renderer';
import { validateDocument, validateLegacyDocument, formatIssues } from '@openthesis/document-schema';
import type { DocumentTemplate, DocumentType, LegacyDocumentJSON, ThesisDocument, JournalArticle, OfficialDocument } from '@openthesis/document-schema';

// The help text and CLAUDE.md document `thesis build …` (the bin is named
// `thesis`), while the agent skill invokes the CLI directly as `build …`.
// Both spellings are accepted so the documented form actually works.
/** Read from the package manifest so `--version` cannot drift from the release. */
const VERSION = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf-8')).version as string;

const rawArgs = process.argv.slice(2);
const args = rawArgs[0] === 'thesis' ? rawArgs.slice(1) : rawArgs;
const command = args[0];

async function main() {
  switch (command) {
    case 'parse':   await cmdParse(args.slice(1)); break;
    case 'import':  cmdImport(args.slice(1)); break;
    case 'build':   await cmdBuild(args.slice(1)); break;
    case 'init':    cmdInit(args.slice(1)); break;
    case '--version':
    case '-v':      console.log(VERSION); break;
    case '--help':
    case '-h':
    case undefined: printHelp(); break;
    default:
      console.error(`Unknown command: ${command}`);
      printHelp();
      process.exit(1);
  }
}

// ── thesis import <content.md> ────────────────────────────

function cmdImport(args: string[]) {
  rejectUnknownFlags(args);
  if (args.length < 1) {
    console.error('Usage: thesis import <content.md> [--type thesis|journal|official] [-o <content.json>]');
    process.exit(1);
  }

  const fromStdin = args[0] === '-';
  const inputPath = fromStdin ? '-' : resolve(args[0]);
  if (!fromStdin && !isMarkdownPath(inputPath)) {
    console.error('Markdown input must use a .md or .markdown extension.');
    process.exit(1);
  }
  const typeOption = readOption(args, '--type');
  const outputOption = readOption(args, '-o');
  if (fromStdin && !outputOption) {
    console.error('Reading Markdown from stdin requires -o <content.json>.');
    process.exit(1);
  }
  const outputPath = outputOption
    ? resolve(outputOption)
    : inputPath.slice(0, -extname(inputPath).length) + '.json';

  if (isSamePath(outputPath, inputPath)) {
    console.error('Output path must be different from the Markdown input path.');
    process.exit(1);
  }
  ensureOverwritable(outputPath, args);
  ensureOutputDir(outputPath);

  try {
    const document = parseMarkdown(readTextFile(inputPath), {
      documentType: typeOption ? validateDocType(typeOption) : undefined,
      sourceName: basename(inputPath, extname(inputPath)),
    });
    writeFileSync(outputPath, JSON.stringify(document, null, 2) + '\n', 'utf-8');
    console.log('✅ Markdown imported successfully!');
    console.log(`   Type:   ${document.type}`);
    console.log(`   Output: ${outputPath}`);
  } catch (err: unknown) {
    console.error(`\n❌ Failed to import Markdown: ${errorMessage(err)}`);
    process.exit(1);
  }
}

// ── thesis parse <template.docx> ──────────────────────────

async function cmdParse(args: string[]) {
  rejectUnknownFlags(args);
  if (args.length < 1) {
    console.error('Usage: thesis parse <template.docx> [--type thesis|journal|official] [--org <name>]');
    process.exit(1);
  }

  const templatePath = resolve(args[0]);
  if (extname(templatePath).toLowerCase() !== '.docx') {
    console.error('Template input must be a .docx file.');
    process.exit(1);
  }
  const docType = validateDocType(readOption(args, '--type') ?? 'thesis');
  const orgName = readOption(args, '--org');

  console.log(`Parsing template: ${templatePath}`);
  console.log(`Document type:   ${docType}`);

  try {
    const buffer = readFileSync(templatePath);
    const template = await parseTemplate(buffer, {
      organization: orgName,
      documentType: docType,
      sourceFile: templatePath,
    });

    const outPath = templatePath.slice(0, -extname(templatePath).length) + '.template.json';
    ensureOverwritable(outPath, args);
    ensureOutputDir(outPath);
    writeFileSync(outPath, JSON.stringify(template, null, 2), 'utf-8');

    console.log(`\n✅ Template parsed successfully!`);
    console.log(`   Output:    ${outPath}`);
    console.log(`   Org:       ${template.meta.organization}`);
    console.log(`   Type:      ${template.meta.documentType}`);
    console.log(`   Page:      ${Math.round(template.page.width / 20)}pt × ${Math.round(template.page.height / 20)}pt`);
    if (template.page.columns && template.page.columns > 1) {
      console.log(`   Columns:   ${template.page.columns}`);
    }
    console.log(`   Styles:    ${Object.keys(template.styles).length} found, ${Object.keys(template.styleRoles).length} roles detected`);
    console.log(`\n   Style roles:`);
    for (const [styleId, role] of Object.entries(template.styleRoles)) {
      const style = template.styles[styleId];
      // A style only declares the properties its template actually specifies,
      // so every field here is optional.
      const fontInfo = style
        ? [
            style.font.eastAsia || style.font.name || '(default font)',
            style.font.size ? `${style.font.size / 2}pt` : '(default size)',
          ].join(' ')
        : '—';
      console.log(`     ${styleId} → ${role} (${fontInfo})`);
    }

    if (template.warnings && template.warnings.length > 0) {
      console.log(`\n   ⚠️  ${template.warnings.length} warning${template.warnings.length === 1 ? '' : 's'}:`);
      for (const warning of template.warnings) {
        console.log(`     - ${warning}`);
      }
    }
  } catch (err: unknown) {
    console.error(`\n❌ Failed to parse template: ${errorMessage(err)}`);
    process.exit(1);
  }
}
async function cmdBuild(args: string[]) {
  rejectUnknownFlags(args);
  if (args.length < 1) {
    console.error('Usage: thesis build <content.json|content.md> [-t <template.json>] [-o <output.docx>] [--type thesis|journal|official]');
    process.exit(1);
  }

  const fromStdin = args[0] === '-';
  const contentPath = fromStdin ? '-' : resolve(args[0]);
  const templateOption = readOption(args, '-t');
  const outputOption = readOption(args, '-o');
  const typeOption = readOption(args, '--type');
  const templatePath = templateOption ? resolve(templateOption) : null;
  const toStdout = outputOption !== undefined && isStdout(outputOption);
  const outputPath = outputOption
    ? (toStdout ? '-' : resolve(outputOption))
    : ['.json', '.md', '.markdown'].includes(extname(contentPath).toLowerCase())
      ? contentPath.slice(0, -extname(contentPath).length) + '.docx'
      : contentPath + '.docx';

  if (toStdout) {
    // The document is binary: one progress line on stdout would corrupt it, so
    // everything this command prints moves to stderr.
    console.log = (...parts: unknown[]) => console.error(...parts);
    if (process.stdout.isTTY) {
      console.error('Refusing to write a binary DOCX to a terminal.');
      console.error('  Redirect it (`-o - > out.docx`) or pass a file path.');
      process.exit(1);
    }
  }
  if (!toStdout && isSamePath(outputPath, contentPath)) {
    console.error('Output path must be different from the content input path.');
    process.exit(1);
  }
  if (!toStdout && templatePath && isSamePath(outputPath, templatePath)) {
    console.error('Output path must be different from the template path.');
    process.exit(1);
  }
  if (!toStdout) {
    ensureOverwritable(outputPath, args);
    ensureOutputDir(outputPath);
  }

  console.log(`Content:  ${contentPath}`);
  console.log(`Template: ${templatePath || '(built-in USTB)'}`);
  console.log(`Output:   ${outputPath}`);

  try {
    const source = readTextFile(contentPath);
    // stdin carries no extension, so a leading `{` is the only honest signal.
    const isMarkdown = fromStdin ? !source.trimStart().startsWith('{') : isMarkdownPath(contentPath);
    const contentJson = isMarkdown
      ? parseMarkdown(source, {
          documentType: typeOption ? validateDocType(typeOption) : undefined,
          sourceName: basename(contentPath, extname(contentPath)),
        })
      : JSON.parse(source);
    const contentDir = dirname(contentPath);

    let template: DocumentTemplate;
    if (templatePath) {
      template = JSON.parse(readTextFile(templatePath));
      // A parsed template carries the problems found while parsing it.
      if (template.warnings && template.warnings.length > 0) {
        for (const warning of template.warnings) {
          console.warn(`  ⚠️  Template: ${warning}`);
        }
      }
    } else {
      console.log('  Using built-in USTB template...');
      template = createUSTBTemplate();
    }

    let docType = 'document';
    if (isMarkdown) console.log(`  Markdown imported as: ${contentJson.type}`);
    if (contentJson && typeof contentJson === 'object' && 'type' in contentJson) {
      console.log(`  Structured document detected: ${contentJson.type}`);

      // Validate before rendering: the renderer's own checks are few, and the
      // errors it can raise name neither the offending block nor its path.
      const issues = validateDocument(contentJson);
      if (issues.length > 0) {
        console.error(`\n❌ Invalid document content — ${issues.length} problem${issues.length === 1 ? '' : 's'}:`);
        console.error(formatIssues(issues));
        process.exit(1);
      }

      docType = contentJson.type === 'thesis' ? 'thesis'
        : contentJson.type === 'journal' ? 'journal article'
        : 'official document';
        
      const buffer = await renderDocument({
        outputPath: toStdout ? '' : outputPath,
        template,
        document: contentJson,
        contentDir,
      });
      if (toStdout) process.stdout.write(buffer);
    } else {
      console.log('  Legacy document format detected, falling back to renderLegacy...');

      const legacyIssues = validateLegacyDocument(contentJson);
      if (legacyIssues.length > 0) {
        console.error(`\n❌ Invalid document content — ${legacyIssues.length} problem${legacyIssues.length === 1 ? '' : 's'}:`);
        console.error(formatIssues(legacyIssues));
        process.exit(1);
      }

      docType = contentJson.title?.includes('期刊') ? 'journal article'
        : contentJson.title?.includes('公文') || contentJson.title?.includes('通知') ? 'official document'
        : 'document';
        
      const buffer = await renderLegacy(contentJson as LegacyDocumentJSON, template, toStdout ? '' : outputPath, contentDir);
      if (toStdout) process.stdout.write(buffer);
    }

    console.log(`\n✅ ${docType} built successfully!`);
    console.log(`   Output: ${outputPath}`);
  } catch (err: unknown) {
    console.error(`\n❌ Failed to build: ${errorMessage(err)}`);
    process.exit(1);
  }
}

// ── thesis init [--type thesis|journal|official] ───────────

function cmdInit(args: string[]) {
  rejectUnknownFlags(args);
  // Shared with every other command so `init --type=journal` works too; the
  // old `args.indexOf('--type')` lookup silently produced a thesis sample.
  const docType = validateDocType(readOption(args, '--type') ?? 'thesis');
  // Overwriting a sample the author has since edited would cost them their work.
  ensureOverwritable(`${docType}-content.json`, args);

  switch (docType) {
    case 'thesis':   createThesisSample(); break;
    case 'journal':  createJournalSample(); break;
    case 'official': createOfficialSample(); break;
    default:         createThesisSample(); break;
  }
}

function createThesisSample() {
  const content: ThesisDocument = {
    type: 'thesis',
    meta: {
      title: '示例学位论文标题',
      degree: 'master',
      date: '2026年6月',
    },
    cover: [
      { type: 'spacer', lines: 6 },
      { type: 'centered_text', text: 'XX大学', font_size_pt: 26, bold: true },
      { type: 'centered_text', text: 'University Name (English)', font_size_pt: 16, bold: false },
      { type: 'spacer', lines: 2 },
      { type: 'centered_text', text: '示例硕士学位论文', font_size_pt: 22, bold: true },
      { type: 'spacer', lines: 4 },
      { type: 'centered_text', text: '学科专业：XXX', font_size_pt: 14, bold: false },
      { type: 'centered_text', text: '研究方向：XXX', font_size_pt: 14, bold: false },
      { type: 'centered_text', text: '日    期：2026年6月', font_size_pt: 14, bold: false },
      { type: 'page_break' },
    ],
    sections: [
      {
        id: 'intro',
        type: 'chapter',
        title: '第一章  绪论',
        content: [
          { type: 'heading2', text: '1.1  研究背景' },
          { type: 'paragraph', text: '正文示例——宋体小四号字，1.5倍行距，首行缩进2字符。西文使用Times New Roman。' },
          { type: 'heading2', text: '1.2  研究目的与意义' },
          { type: 'paragraph', text: '在此撰写研究目的...' },
        ],
      },
      {
        id: 'theory',
        type: 'chapter',
        title: '第二章  理论基础',
        content: [
          { type: 'heading2', text: '2.1  基本方程' },
          { type: 'paragraph', text: '控制方程为：' },
          { type: 'equation', latex: 'E = mc^2' },
        ],
      },
      {
        id: 'conclusion',
        type: 'chapter',
        title: '第三章  结论',
        content: [
          { type: 'paragraph', text: '研究成果总结...' },
        ],
      },
    ],
    backMatter: {
      references: [
        { id: '1', text: '作者一, 作者二. 论文标题. 期刊名称, 2026.' }
      ]
    }
  };

  writeFileSync('thesis-content.json', JSON.stringify(content, null, 2), 'utf-8');
  console.log('✅ Created thesis-content.json (学位论文示例)');
  printInitNext('thesis');
}

function createJournalSample() {
  const content: JournalArticle = {
    type: 'journal',
    meta: {
      title: '示例期刊论文标题',
      authors: [
        { name: '作者一', affiliations: [{ institution: 'XX大学' }] },
        { name: '作者二', affiliations: [{ institution: 'XX研究所' }] },
        { name: '通讯作者', isCorresponding: true, affiliations: [{ institution: 'XX大学' }] }
      ],
      abstract: '本文研究了XXX问题，采用XXX方法，得出XXX结论。',
      keywords: ['关键词一', '关键词二', '关键词三']
    },
    sections: [
      {
        id: 'intro',
        type: 'introduction',
        title: '1  Introduction',
        content: [
          { type: 'paragraph', text: 'Introduction content here. This is a sample journal article body text.' },
          { type: 'heading2', text: '1.1  Background' },
          { type: 'paragraph', text: 'Background and literature review content...' },
        ]
      },
      {
        id: 'methods',
        type: 'methods',
        title: '2  Methods',
        content: [
          { type: 'paragraph', text: 'Methodology description...' },
        ]
      },
      {
        id: 'results',
        type: 'results',
        title: '3  Results and Discussion',
        content: [
          { type: 'paragraph', text: 'Results analysis...' },
          { type: 'table', caption: 'Table 1  Experimental results', headers: ['Parameter', 'Value', 'Error'], data: [['p1', '0.123', '±0.001'], ['p2', '0.456', '±0.002']] },
        ]
      },
      {
        id: 'conclusion',
        type: 'conclusion',
        title: '4  Conclusion',
        content: [
          { type: 'paragraph', text: 'Concluding remarks...' },
        ]
      }
    ],
    backMatter: {
      references: [
        { id: '1', text: 'Author A, Author B. Title of paper. Journal Name, 2026, 100(1): 1-10.' },
        { id: '2', text: 'Author C. Title of book. Publisher, 2025.' }
      ]
    }
  };

  writeFileSync('journal-content.json', JSON.stringify(content, null, 2), 'utf-8');
  console.log('✅ Created journal-content.json (期刊论文示例)');
  printInitNext('journal');
}

function createOfficialSample() {
  const content: OfficialDocument = {
    type: 'official',
    meta: {
      title: '关于做好2026年防汛工作的通知',
      issuingAuthority: 'XX省人民政府',
      documentNumber: 'X政发〔2026〕1号',
      documentCategory: '通知',
      primaryRecipients: ['各市、州人民政府', '省政府各部门'],
      ccRecipients: ['省委各部门', '省人大常委会办公厅', '省政协办公厅'],
      date: '2026年6月13日',
    },
    body: [
      { type: 'recipient_line', text: '各市、州人民政府，省政府各部门：', recipientType: 'primary' },
      { type: 'paragraph', text: '为切实做好2026年防汛工作，保障人民群众生命财产安全，现将有关事项通知如下：' },
      { type: 'heading1', text: '一、提高思想认识，压实防汛责任' },
      { type: 'paragraph', text: '各级各部门要充分认识当前防汛形势的严峻性，坚决克服麻痹思想和侥幸心理。要严格落实以行政首长负责制为核心的各项防汛责任制，确保责任到人、措施到位。' },
      { type: 'heading1', text: '二、加强监测预警，做好应急准备' },
      { type: 'paragraph', text: '气象、水文部门要加强监测预报，及时发布预警信息。各地要修订完善防汛应急预案，充实抢险救援队伍，储备充足的防汛物资。' },
    ],
    attachments: [
      {
        order: 1,
        title: '2026年防汛重点区域清单',
        content: [
          { type: 'paragraph', text: '1. XX水库周边区域' },
          { type: 'paragraph', text: '2. XX河流中下游低洼地带' }
        ]
      }
    ]
  };

  writeFileSync('official-content.json', JSON.stringify(content, null, 2), 'utf-8');
  console.log('✅ Created official-content.json (公文示例)');
  printInitNext('official');
}

function printInitNext(type: DocumentType) {
  console.log('   Next steps:');
  console.log(`     1. Edit the generated JSON with your content`);
  console.log(`     2. thesis parse <your-template.docx> --type ${type}`);
  console.log(`     3. thesis build <content.json> -t <template.json>`);
}

const VALID_DOC_TYPES: DocumentType[] = ['thesis', 'journal', 'official'];

/** Options the CLI understands; anything else on the command line is a typo. */
const VALUE_FLAGS = ['--type', '-t', '-o', '--org'];
const SET_FLAGS = ['--force', '-y'];
const KNOWN_FLAGS = new Set([...VALUE_FLAGS, ...SET_FLAGS, '--help', '-h', '--version', '-v']);

/**
 * Read an option written either as `--flag value` or `--flag=value`.
 *
 * The `=` form used to be ignored outright: `build x.md --type=official` fell
 * back to the front-matter type and exited 0, so the flag silently did nothing.
 */
function readOption(args: string[], flag: string): string | undefined {
  const inlinePrefix = `${flag}=`;
  const inline = args.find(arg => arg.startsWith(inlinePrefix));
  if (inline !== undefined) {
    const value = inline.slice(inlinePrefix.length);
    if (!value) {
      console.error(`Missing value for ${flag}.`);
      process.exit(1);
    }
    return value;
  }

  const index = args.indexOf(flag);
  if (index < 0) return undefined;
  const value = args[index + 1];
  // `-` is a value (stdin/stdout), not a stray flag.
  if (!value || (value.startsWith('-') && value !== '-')) {
    console.error(`Missing value for ${flag}.`);
    process.exit(1);
  }
  return value;
}

/** Fail loudly on unknown `-`/`--` tokens instead of quietly ignoring them. */
function rejectUnknownFlags(args: string[]): void {
  for (const arg of args) {
    if (!arg.startsWith('-') || arg === '-') continue;
    const name = arg.includes('=') ? arg.slice(0, arg.indexOf('=')) : arg;
    if (!KNOWN_FLAGS.has(name)) {
      console.error(`Unknown option: ${name}`);
      console.error(`Known options: ${[...KNOWN_FLAGS].join(', ')}`);
      process.exit(1);
    }
  }
}

/**
 * True when two paths refer to the same file.
 *
 * Windows and macOS filesystems are case-insensitive, so the previous plain
 * string comparison let `import a.md -o a.MD` overwrite the Markdown source
 * with the JSON it had just produced.
 */
function isSamePath(a: string, b: string): boolean {
  const normalize = (p: string) => {
    const resolved = resolve(p);
    try {
      return realpathSync.native(resolved).toLowerCase();
    } catch {
      // Output paths usually do not exist yet.
      return resolved.toLowerCase();
    }
  };
  return normalize(a) === normalize(b);
}

/** Read UTF-8 text, dropping a BOM that would otherwise break JSON.parse. */
function readTextFile(path: string): string {
  // `-` means stdin, so a manuscript can be piped in without a temp file.
  if (path === '-') return readFileSync(0, 'utf-8').replace(/^\uFEFF/, '');
  return readFileSync(path, 'utf-8').replace(/^\uFEFF/, '');
}

/** Create the output's parent directory so `-o a/b/c.docx` works like `mkdir -p`. */
function ensureOutputDir(path: string): void {
  mkdirSync(dirname(resolve(path)), { recursive: true });
}

/**
 * Refuse to write over an existing file unless `--force` was given.
 *
 * Every output used to be overwritten in silence, so a second run — or a typo in
 * `-o` — destroyed the previous document without a word.
 */
function ensureOverwritable(path: string, args: string[]): void {
  if (!existsSync(path)) return;
  if (args.includes('--force') || args.includes('-y')) return;
  console.error(`Output already exists: ${path}`);
  console.error('  Pass --force to overwrite it.');
  process.exit(1);
}

/** True when the caller asked to write the document to stdout. */
function isStdout(path: string): boolean {
  return path === '-';
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function isMarkdownPath(path: string): boolean {
  return ['.md', '.markdown'].includes(extname(path).toLowerCase());
}

function validateDocType(value: string | undefined): DocumentType {
  if (!value || !VALID_DOC_TYPES.includes(value as DocumentType)) {
    console.error(`\n\u274c Invalid document type: "${value || ''}"`);
    console.error(`   Valid types: ${VALID_DOC_TYPES.join(', ')}`);
    process.exit(1);
  }
  return value as DocumentType;
}

// ── Help ──────────────────────────────────────────────────

function printHelp() {
  console.log(`
  OpenThesis — AI-powered Document Template Engine
  ─────────────────────────────────────────────────

  Parse any .docx template. Write in structured JSON.
  Output submission-ready DOCX. For thesis, journal
  articles, and official documents (公文).

  Commands:
    thesis parse <template.docx>     Parse template → JSON style DSL
    thesis import <content.md>        Convert Markdown → structured JSON
    thesis build <content.json|md>    Render document → .docx
    thesis init [--type <t>]          Create sample content file

  Document types (--type):
    thesis   — 学位论文 (default)
    journal  — 期刊论文
    official — 公文 (GB/T 9704)

  Examples:
    thesis parse 清华博士模板.docx --type thesis --org "清华大学"
    thesis parse elsevier-template.docx --type journal
    thesis parse 公文模板.docx --type official --org "XX省人民政府"

    thesis import manuscript.md --type thesis
    thesis build manuscript.md -t template.json -o output.docx
    thesis build thesis-content.json -t template.json -o output.docx

  Options (build):
    -t <template.json>    Template file (default: built-in USTB)
    -o <output.docx>      Output path (default: input + .docx, "-" for stdout)
    --type <t>            Markdown document type (overrides front matter)
    --force, -y           Overwrite an existing output file

  Other:
    --version, -v         Print the version
    -                     Read the input from stdin (build, import)

  Examples:
    thesis import - --type journal -o paper.json < manuscript.md
    thesis build - -t template.json -o - > output.docx
  `);
}

main().catch(e => { console.error(e); process.exit(1); });
