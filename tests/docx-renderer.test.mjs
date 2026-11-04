import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createUSTBTemplate, renderDocument } from '../packages/docx-renderer/dist/index.js';
import { parseTemplate } from '../packages/template-parser/dist/index.js';

/** The instruction text of every complex field in the document. */
function fieldInstrs(xml) {
  return [...xml.matchAll(/<w:instrText[^>]*>([^<]*)<\/w:instrText>/g)].map(m => m[1].trim());
}

/** The bookmark a `REF` field points at, if there is one. */
function refTarget(xml) {
  return fieldInstrs(xml).map(i => /^REF\s+(\S+)/.exec(i)?.[1]).find(Boolean);
}


/** Every `w:t` in a paragraph, joined — how a caption reads once fields are in it. */
function paragraphText(xml, anchor) {
  const at = xml.indexOf(anchor);
  if (at < 0) return '';
  const open = xml.lastIndexOf('<w:p>', at);
  const close = xml.indexOf('</w:p>', at);
  return [...xml.slice(open, close).matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(m => m[1]).join('');
}


const template = createUSTBTemplate();

async function documentXml(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  return { zip, xml: await zip.file('word/document.xml').async('string') };
}

test('renders real list, code, quote and figure blocks with a caption', async () => {
  const contentDir = mkdtempSync(join(tmpdir(), 'openthesis-renderer-'));
  writeFileSync(
    join(contentDir, 'figure.png'),
    Buffer.from('/9j/4AAQSkZJRgABAQAAAAAAAAD/2wBDAAMCAgICAgMCAgIDAwMDBAYEBAQEBAgGBgUGCQgKCgkICQkKDA8MCgsOCwkJDRENDg8QEBEQCgwSExIQEw8QEBD/wAALCAACAAIBAREA/8QAFAABAAAAAAAAAAAAAAAAAAAACf/EABQQAQAAAAAAAAAAAAAAAAAAAAD/2gAIAQEAAD8AVN//2Q==', 'base64'),
  );
  const buffer = await renderDocument({
    outputPath: '',
    template,
    contentDir,
    document: {
      type: 'thesis',
      meta: { title: 'Renderer test' },
      cover: [],
      sections: [{
        id: 'blocks', type: 'chapter', title: 'Blocks',
        content: [
          { type: 'list_item', text: 'First item' },
          { type: 'code_block', text: 'const x = 1;\nconsole.log(x);' },
          { type: 'blockquote', text: 'Quoted text' },
          { type: 'figure', path: 'figure.png', caption: 'Figure caption', number: '1' },
        ],
      }],
    },
  });
  const { zip, xml } = await documentXml(buffer);
  assert.match(xml, /First item/);
  assert.match(xml, /const x = 1;/);
  assert.match(xml, /Quoted text/);
  // One space between the number and the title; the old renderer emitted two.
  // The number is its own run now, wrapped in a bookmark so a `[[图1]]`
  // reference in the body can point at it — the text still reads the same.
  assert.match(xml, /<w:t xml:space="preserve">1<\/w:t>/);
  assert.match(xml, /<w:t xml:space="preserve"> Figure caption<\/w:t>/);
  assert.match(xml, /<w:bookmarkStart[^>]*w:name="_Ref1"/);
  assert.doesNotMatch(xml, /1 {2}Figure caption/);
  assert.doesNotMatch(xml, /\[list_item:|\[code_block:|\[blockquote:/);
  const relationships = await zip.file('word/_rels/document.xml.rels').async('string');
  if (/relationships\/theme/.test(relationships)) {
    assert.ok(zip.file('word/theme/theme1.xml'), 'a theme relationship must keep its target part');
  }

  const mediaName = Object.keys(zip.files).find(name => name.startsWith('word/media/') && !name.endsWith('/'));
  assert.ok(mediaName, 'figure image should be embedded');
  const media = await zip.file(mediaName).async('nodebuffer');
  assert.equal(media.readUInt16BE(0), 0xffd8, 'JPEG bytes should be detected despite the .png filename');
});

test('renders the CC label exactly once in official documents', async () => {
  const buffer = await renderDocument({
    outputPath: '',
    template,
    document: {
      type: 'official',
      meta: {
        title: '测试通知', issuingAuthority: '测试机关', documentNumber: '测〔2026〕1号',
        documentCategory: '通知', primaryRecipients: ['主送单位'], ccRecipients: ['抄送单位'], date: '2026年8月30日',
      },
      body: [{ type: 'paragraph', text: '正文。' }],
    },
  });
  const { xml } = await documentXml(buffer);
  assert.match(xml, /抄送：抄送单位。/);
  assert.doesNotMatch(xml, /抄送：抄送：/);
});

test('rejects tables without header columns', async () => {
  await assert.rejects(() => renderDocument({
    outputPath: '',
    template,
    document: {
      type: 'thesis', meta: { title: 'Bad table' }, cover: [],
      sections: [{
        id: 'table', type: 'chapter', title: 'Table',
        content: [{ type: 'table', caption: 'Empty', headers: [], data: [] }],
      }],
    },
  }), /must define at least one header column/);
});

test('rejects table column width counts that do not match headers', async () => {
  await assert.rejects(() => renderDocument({
    outputPath: '', template,
    document: {
      type: 'thesis', meta: { title: 'Bad widths' }, cover: [],
      sections: [{
        id: 'table', type: 'chapter', title: 'Table',
        content: [{ type: 'table', caption: 'Widths', headers: ['A', 'B'], data: [], columnWidths: [1000] }],
      }],
    },
  }), /2 columns but 1 column widths/);
});

test('writes configured multi-column sections to OOXML', async () => {
  const twoColumnTemplate = structuredClone(template);
  twoColumnTemplate.page.columns = 2;
  twoColumnTemplate.page.columnGutter = 720;
  const buffer = await renderDocument({
    outputPath: '', template: twoColumnTemplate,
    document: {
      type: 'thesis', meta: { title: 'Columns' }, cover: [],
      sections: [{ id: 'body', type: 'chapter', title: 'Body', content: [] }],
    },
  });
  const { xml } = await documentXml(buffer);
  assert.match(xml, /<w:cols[^>]*w:space="720"[^>]*w:num="2"|<w:cols[^>]*w:num="2"[^>]*w:space="720"/);
});

test('carries an exact line-spacing rule into the OOXML', async () => {
  const exactTemplate = structuredClone(template);
  // A role is resolved through `roleWinners`; "the first style carrying the
  // role" is a different style, so mutating that one changed nothing.
  const paragraphStyleId = exactTemplate.roleWinners?.paragraph
    ?? Object.entries(exactTemplate.styleRoles).find(([, role]) => role === 'paragraph')[0];
  exactTemplate.styles[paragraphStyleId] = {
    ...exactTemplate.styles[paragraphStyleId],
    lineSpacing: 360,
    lineSpacingRule: 'exact',
  };

  const buffer = await renderDocument({
    outputPath: '', template: exactTemplate,
    document: {
      type: 'thesis', meta: { title: 'Spacing' }, cover: [],
      sections: [{
        id: 'body', type: 'chapter', title: 'Body',
        content: [{ type: 'paragraph', text: 'Body text.' }],
      }],
    },
  });
  const { xml } = await documentXml(buffer);
  // Without the rule Word reads 360 as 1.5 lines instead of 18 points.
  assert.match(
    xml,
    /<w:spacing[^>]*w:line="360"[^>]*w:lineRule="exact"|<w:spacing[^>]*w:lineRule="exact"[^>]*w:line="360"/,
  );
});

test('finds a style role without scanning the whole table per block', async () => {
  // `Object.entries(styleRoles).find(...)` materialised all 72 pairs before
  // matching, which cost 9-11µs per block — about a fifth of render time.
  let reads = 0;
  const countingTemplate = structuredClone(template);
  countingTemplate.styleRoles = new Proxy(countingTemplate.styleRoles, {
    get(target, key, receiver) {
      reads += 1;
      return Reflect.get(target, key, receiver);
    },
  });

  const blocks = Array.from({ length: 40 }, (_, index) => ({ type: 'paragraph', text: `Paragraph ${index}.` }));
  await renderDocument({
    outputPath: '', template: countingTemplate,
    document: {
      type: 'thesis', meta: { title: 'Role lookup' }, cover: [],
      sections: [{ id: 'body', type: 'chapter', title: 'Body', content: blocks }],
    },
  });

  // A per-block scan reads every role for every block (≈2900 here); stopping at
  // the first match reads a handful. Counting property reads keeps the test
  // independent of machine speed.
  assert.ok(reads < blocks.length * 5, `${reads} role reads for ${blocks.length} blocks — the table is scanned per block`);
});

test('the built-in USTB template carries the university formatting', () => {
  // The asset was generated by a parser that invented one signature for all 72
  // styles, so every heading rendered like body text. These are the numbers the
  // 《北京科技大学研究生学位论文书写指南》 asks for.
  const asset = JSON.parse(readFileSync('assets/ustb-thesis-template.json', 'utf8'));
  const style = role => asset.styles[asset.roleWinners[role]];

  assert.equal(style('heading1').font.eastAsia, '黑体');
  assert.equal(style('heading1').font.size, 30);          // 小三
  assert.equal(style('heading1').paragraph.alignment, 'center');
  assert.equal(style('heading2').font.size, 28);          // 四号
  assert.equal(style('heading3').font.size, 28);
  assert.equal(style('paragraph').font.eastAsia, '宋体');
  assert.equal(style('paragraph').font.size, 24);         // 小四
  assert.equal(style('paragraph').paragraph.firstLineIndent, 480);   // 2 characters
});

test('writes supported LaTeX as native Office Math elements', async () => {
  const buffer = await renderDocument({
    outputPath: '', template,
    document: {
      type: 'thesis', meta: { title: 'Native math' }, cover: [],
      sections: [{
        id: 'math', type: 'chapter', title: 'Math',
        content: [{ type: 'equation_numbered', latex: '\\frac{x_1}{\\sqrt{y^2}}', number: '1' }],
      }],
    },
  });
  const { xml } = await documentXml(buffer);
  assert.match(xml, /<m:oMath>/);
  assert.match(xml, /<m:f>/);
  assert.match(xml, /<m:rad>/);
  assert.match(xml, /<m:sSub>/);
  assert.match(xml, /<m:sSup>/);
  assert.match(xml, /\(1\)/);
});

test('writes advanced LaTeX macros and environments as native OMML', async () => {
  const buffer = await renderDocument({
    outputPath: '', template,
    document: {
      type: 'thesis', meta: { title: 'Advanced native math' }, cover: [],
      sections: [{
        id: 'math', type: 'chapter', title: 'Math',
        content: [
          { type: 'equation', latex: '\\hat{x}+\\lim_{t\\to0}\\sin(t)+\\binom{n}{k}' },
          { type: 'equation', latex: '\\overline{AB}+\\prod_{i=1}^{n}x_i+\\overset{*}{=}+\\left\\langle x,y\\right\\rangle' },
          { type: 'equation', latex: '\\begin{pmatrix}a & b \\\\ c & d\\end{pmatrix}' },
          { type: 'equation', latex: 'f(x)=\\begin{cases}x^2 & x\\geq0 \\\\ -x & x<0\\end{cases}' },
        ],
      }],
    },
  });
  const { xml } = await documentXml(buffer);
  assert.match(xml, /<m:acc>/);
  assert.match(xml, /<m:bar>/);
  assert.match(xml, /<m:limLow>/);
  assert.match(xml, /<m:limUpp>/);
  assert.match(xml, /<m:func>/);
  assert.match(xml, /<m:nary>/);
  assert.match(xml, /<m:chr m:val="∏"\/>/);
  assert.match(xml, /<m:type m:val="noBar"\/>/);
  assert.match(xml, /<m:m>/);
  assert.match(xml, /<m:mr>/);
  assert.match(xml, /<m:begChr m:val="\("\/>/);
  assert.match(xml, /<m:begChr m:val="\{"\/>/);
  assert.match(xml, /<m:begChr m:val="⟨"\/>/);
  assert.match(xml, /<m:endChr m:val=""\/>/);
});

/** The `<w:rPr>` of the run containing `needle`. */
function runProperties(xml, needle) {
  const index = xml.indexOf(needle);
  assert.notEqual(index, -1, `expected to find ${needle}`);
  const start = xml.lastIndexOf('<w:r>', index);
  return xml.slice(start, xml.indexOf('</w:r>', index) + 6);
}

test('renders a fourth-level heading as a heading, not as body text', async () => {
  const buffer = await renderDocument({
    outputPath: '', template,
    document: {
      type: 'thesis', meta: { title: 'Nesting' }, cover: [],
      sections: [{
        id: 'l1', type: 'chapter', title: 'Level One', content: [],
        subsections: [{
          id: 'l2', type: 'section', title: 'Level Two', content: [],
          subsections: [{
            id: 'l3', type: 'section', title: 'Level Three', content: [],
            subsections: [{ id: 'l4', type: 'section', title: 'Level Four', content: [] }],
          }],
        }],
      }],
    },
  });
  const { xml } = await documentXml(buffer);
  // `heading4` was missing from the fallback table, so it resolved to the body
  // paragraph style and looked identical to surrounding text.
  const body = runProperties(xml, 'Level One');
  const fourth = runProperties(xml, 'Level Four');
  assert.match(fourth, /<w:b\/>/);
  assert.doesNotMatch(fourth, /w:eastAsia="仿宋"/);
  assert.match(fourth, /w:eastAsia="黑体"/);
  assert.notEqual(fourth, body);
});

test('fits a table to the printable page width instead of a fixed A4 constant', async () => {
  const narrow = structuredClone(template);
  narrow.page.width = 8391;
  narrow.page.height = 11906;
  narrow.page.margins = { top: 1701, bottom: 1701, left: 1701, right: 1701 };
  const printable = narrow.page.width - narrow.page.margins.left - narrow.page.margins.right;

  const buffer = await renderDocument({
    outputPath: '', template: narrow,
    document: {
      type: 'thesis', meta: { title: 'Table width' }, cover: [],
      sections: [{
        id: 't', type: 'chapter', title: 'T',
        content: [{ type: 'table', caption: 'Cap', headers: ['A', 'B'], data: [['1', '2']] }],
      }],
    },
  });
  const { xml } = await documentXml(buffer);
  // Cell widths always summed to 8504 twips, overflowing a narrower page.
  const cellWidths = [...xml.matchAll(/<w:tcW w:type="dxa" w:w="(\d+)"/g)].map(m => Number(m[1]));
  const rowWidth = cellWidths.slice(0, 2).reduce((sum, width) => sum + width, 0);
  assert.ok(rowWidth <= printable, `row width ${rowWidth} must fit printable ${printable}`);
  assert.match(xml, new RegExp(`<w:tblW w:type="dxa" w:w="${rowWidth}"`));
  assert.doesNotMatch(xml, /<w:tblW w:type="pct"/);
});

test('takes table typography from the template', async () => {
  const styled = structuredClone(template);
  styled.styles = {
    ...styled.styles,
    MyTableText: { font: { name: 'Georgia', eastAsia: '楷体', size: 36 }, paragraph: { alignment: 'left' } },
    MyTableHead: { font: { name: 'Georgia', eastAsia: '黑体', size: 36, bold: true }, paragraph: { alignment: 'left' } },
  };
  styled.styleRoles = { ...styled.styleRoles, MyTableText: 'table', MyTableHead: 'table_header' };

  const buffer = await renderDocument({
    outputPath: '', template: styled,
    document: {
      type: 'thesis', meta: { title: 'Table style' }, cover: [],
      sections: [{
        id: 't', type: 'chapter', title: 'T',
        content: [{ type: 'table', caption: 'Cap', headers: ['HDR'], data: [['DAT']] }],
      }],
    },
  });
  const { xml } = await documentXml(buffer);
  // Tables used to hardcode Times New Roman / 黑体 / 宋体 and ignore the template.
  assert.match(runProperties(xml, 'HDR'), /w:ascii="Georgia"/);
  assert.match(runProperties(xml, 'HDR'), /w:eastAsia="黑体"/);
  assert.match(runProperties(xml, 'HDR'), /<w:sz w:val="36"\/>/);
  assert.match(runProperties(xml, 'DAT'), /w:eastAsia="楷体"/);
  assert.match(runProperties(xml, 'DAT'), /<w:sz w:val="36"\/>/);
});

test('falls back to domain-specific typography when the template is silent', async () => {
  const bare = structuredClone(template);
  bare.styles = {};
  bare.styleRoles = {};

  const thesisBuffer = await renderDocument({
    outputPath: '', template: bare,
    document: {
      type: 'thesis', meta: { title: 'Defaults' }, cover: [],
      sections: [{ id: 'p', type: 'chapter', title: 'Chapter', content: [{ type: 'paragraph', text: 'BodyText' }] }],
    },
  });
  const officialBuffer = await renderDocument({
    outputPath: '', template: bare,
    document: {
      type: 'official',
      meta: {
        title: '公文', issuingAuthority: '某机关', documentNumber: 'X〔2026〕1号',
        documentCategory: '通知', primaryRecipients: ['各单位'], date: '2026年1月1日',
      },
      body: [{ type: 'paragraph', text: 'BodyText' }],
    },
  });

  const thesisXml = (await documentXml(thesisBuffer)).xml;
  const officialXml = (await documentXml(officialBuffer)).xml;

  // 学位论文 body is 宋体 小四; GB/T 9704 公文 body is 仿宋 三号.
  assert.match(runProperties(thesisXml, 'BodyText'), /w:eastAsia="宋体"/);
  assert.match(runProperties(thesisXml, 'BodyText'), /<w:sz w:val="24"\/>/);
  assert.match(runProperties(officialXml, 'BodyText'), /w:eastAsia="仿宋"/);
  assert.match(runProperties(officialXml, 'BodyText'), /<w:sz w:val="32"\/>/);
  assert.match(runProperties(thesisXml, 'Chapter'), /w:eastAsia="黑体"/);
});

test('emits real Word list numbering instead of literal markers', async () => {
  const buffer = await renderDocument({
    outputPath: '', template,
    document: {
      type: 'thesis', meta: { title: 'Lists' }, cover: [],
      sections: [{
        id: 'l', type: 'chapter', title: 'Lists',
        content: [
          { type: 'list_item', text: 'bullet one' },
          { type: 'list_item', text: 'bullet two', level: 1 },
          { type: 'paragraph', text: 'interrupt' },
          { type: 'list_item', text: 'second bullet list' },
          { type: 'paragraph', text: 'interrupt again' },
          { type: 'list_item', text: 'ordered one', ordered: true },
          { type: 'list_item', text: 'ordered two', ordered: true },
        ],
      }],
    },
  });

  const { zip, xml } = await documentXml(buffer);
  const numbering = await zip.file('word/numbering.xml')?.async('string');

  assert.ok(numbering, 'numbering.xml must be written');
  assert.match(numbering, /<w:abstractNum/);
  assert.match(numbering, /w:val="•"/);

  const numPrs = [...xml.matchAll(/<w:numPr>[\s\S]*?<\/w:numPr>/g)];
  assert.equal(numPrs.length, 5);

  // Each separate list restarts its own counter, so the two bullet lists must
  // not share a numId.
  const numIds = [...xml.matchAll(/<w:numId w:val="(\d+)"\/>/g)].map(m => m[1]);
  assert.ok(new Set(numIds).size >= 3, `expected distinct numIds, got ${JSON.stringify(numIds)}`);

  const levels = [...xml.matchAll(/<w:ilvl w:val="(\d+)"\/>/g)].map(m => m[1]);
  assert.ok(levels.includes('1'), 'nested level must be written');

  // The old renderer prefixed the text with a literal marker.
  assert.doesNotMatch(xml, /<w:t[^>]*>•\s/);
  assert.match(xml, /bullet one/);
  assert.match(xml, /ordered two/);
});

test('applies the guide formatting the built-in template carries', async () => {
  // One document that exercises the properties added for the university guide,
  // asserted on the produced OOXML rather than on the template object.
  const buffer = await renderDocument({
    outputPath: '', template,
    document: {
      type: 'thesis', meta: { title: 'Guide check' }, cover: [],
      sections: [
        { id: 'abs', type: 'abstract', title: '摘要', content: [{ type: 'paragraph', text: '摘要正文。' }] },
        { id: 'ch', type: 'chapter', title: '第一章', content: [{ type: 'paragraph', text: '正文。' }] },
      ],
      backMatter: { references: [{ id: '1', text: '一条参考文献.' }] },
    },
  });
  const { zip, xml } = await documentXml(buffer);

  // 对称页边距 + 装订线 1cm (567 twips) for a bound thesis. Mirror margins are a
  // document setting (settings.xml), not a section property — writing it into
  // w:sectPr produced a part Word ignores.
  const settings = await zip.file('word/settings.xml').async('string');
  assert.match(settings, /<w:mirrorMargins\/>/);
  assert.doesNotMatch(xml, /w:mirrorMargins/);
  assert.match(xml, /w:gutter="567"/);
  // A section heading is not a chapter heading: 段前17磅/段后16.5磅/2.41倍行距(579),
  // where heading1 is 340/340 and 1.3倍(312).
  assert.match(xml, /w:line="579"/);
  // A hanging indent belongs to a flush-left heading (level 2 = 1cm) and to a
  // reference entry (1cm, measured in the real thesis). It must NOT sit on a
  // centred heading: the guide's "0.75cm" cannot survive contact with a centred
  // line, and the accepted thesis has every chapter heading exactly centred.
  assert.match(xml, /w:hanging="567"/);
  assert.doesNotMatch(xml, /w:hanging="425"/);
  // Body: 宋体 小四 with a two-character first-line indent at 1.3 line spacing.
  assert.match(xml, /w:firstLine="480"/);
  assert.match(xml, /w:line="312"/);
});

test('styles the cover from the template, one tier at a time', async () => {
  // The guide specifies the cover box by box: 校名行 小二 18pt bold centred,
  // 研究生/指导教师 四号 14pt bold on justified rows, 中图分类号 五号. The renderer
  // used to hardcode centre and one size for every cover line.
  const buffer = await renderDocument({
    outputPath: '', template,
    document: {
      type: 'thesis', meta: { title: 'Cover' },
      cover: [
        { type: 'centered_text', text: '北京科技大学博士学位论文', styleRole: 'cover_title' },
        { type: 'centered_text', text: '研究生  张三', styleRole: 'cover_line' },
        { type: 'centered_text', text: '中图分类号：TU3', styleRole: 'cover_meta' },
      ],
      sections: [{ id: 'a', type: 'chapter', title: '第一章', content: [] }],
    },
  });
  const { xml } = await documentXml(buffer);

  const para = text => xml.split(/<\/w:p>/).find(chunk => chunk.includes(text)) ?? '';
  const title = para('北京科技大学博士学位论文');
  assert.match(title, /<w:sz w:val="36"\/>/);       // 小二
  assert.match(title, /<w:b\/>/);
  assert.match(title, /<w:jc w:val="center"\/>/);

  const line = para('研究生  张三');
  assert.match(line, /<w:sz w:val="28"\/>/);        // 四号
  assert.match(line, /<w:jc w:val="center"\/>/);    // centered cover metadata row

  const meta = para('中图分类号：TU3');
  assert.match(meta, /<w:sz w:val="21"\/>/);        // 五号
  assert.doesNotMatch(meta, /<w:b\/>/);
});

test('a 目录 section lists the headings, with a page reference on each line', async () => {
  // A TOC field is empty until Word updates it, and most people never press F9.
  // The entries are built at render time; each carries a PAGEREF field so Word
  // still refreshes the page numbers when fields update.
  const buffer = await renderDocument({
    outputPath: '', template,
    document: {
      type: 'thesis', meta: { title: 'TOC' }, cover: [],
      sections: [
        { id: 'toc', type: 'toc', title: '目录', content: [] },
        { id: 'c1', type: 'chapter', title: '第一章 绪论', content: [{ type: 'paragraph', text: '正文。' }] },
      ],
    },
  });
  const { xml } = await documentXml(buffer);
  assert.match(xml, /目录/);
  assert.match(xml, /第一章 绪论/);
  // One PAGEREF per entry, pointing at a bookmark the heading actually carries.
  const refs = fieldInstrs(xml).filter(i => i.startsWith('PAGEREF'));
  assert.equal(refs.length, 1, `expected one PAGEREF, got ${JSON.stringify(refs)}`);
  const target = /PAGEREF\s+(\S+)/.exec(refs[0])?.[1];
  assert.ok(target, 'the PAGEREF must name a bookmark');
  assert.match(xml, new RegExp(`w:name="${target}"`), 'the heading must carry that bookmark');
});

/** Every `w:sectPr` in document order: the section breaks of the document. */
function sectionProperties(xml) {
  return [...xml.matchAll(/<w:sectPr[\s\S]*?<\/w:sectPr>/g)].map(match => match[0]);
}

/** The text of every header (or footer) part in the package. */
async function runningHeadTexts(zip, kind) {
  const names = Object.keys(zip.files)
    .filter(name => new RegExp(`^word/${kind}\\d+\\.xml$`).test(name))
    .sort();
  const parts = [];
  for (const name of names) {
    const xml = await zip.file(name).async('string');
    parts.push({
      name,
      xml,
      text: [...xml.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(match => match[1]).join(''),
    });
  }
  return parts;
}

/**
 * The header parts the document references as `w:type="even"`, resolved through
 * the package relationships — a part's *name* does not say which slot it fills,
 * so the only way to know which header is the even one is to follow the r:id.
 */
async function evenHeaderParts(zip, xml) {
  const rels = await zip.file('word/_rels/document.xml.rels').async('string');
  const relationshipTags = [...rels.matchAll(/<Relationship\b[^>]*\/>/g)].map(match => match[0]);
  const refs = [...xml.matchAll(/<w:headerReference\b[^>]*>/g)]
    .map(match => match[0])
    .filter(tag => /w:type="even"/.test(tag));

  const parts = [];
  for (const ref of refs) {
    const id = (ref.match(/r:id="([^"]+)"/) ?? [])[1];
    const target = (relationshipTags.find(tag => tag.includes(`Id="${id}"`))?.match(/Target="([^"]+)"/) ?? [])[1];
    assert.ok(target, `no relationship target for ${id}`);
    const name = `word/${target.replace(/^\/?word\//, '')}`;
    const partXml = await zip.file(name).async('string');
    parts.push({
      name,
      xml: partXml,
      text: [...partXml.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(match => match[1]).join(''),
    });
  }
  return parts;
}

test('lays a thesis out as a cover, a front matter and a body section', async () => {
  // A thesis is not one run of pages. The cover carries no running head and no
  // page number, the front matter is numbered I, II, III, and the body restarts
  // at 1 — three page setups, so three sections.
  const buffer = await renderDocument({
    outputPath: '', template,
    document: {
      type: 'thesis',
      meta: { title: '节理岩体隧道纵向地震易损性分析' },
      cover: [
        { type: 'centered_text', text: '北京科技大学硕士学位论文', styleRole: 'cover_title' },
        { type: 'page_break' },
      ],
      sections: [
        { id: 'abs', type: 'abstract', title: '摘要', content: [{ type: 'paragraph', text: '摘要正文。' }] },
        { id: 'toc', type: 'toc', title: '目录', content: [] },
        {
          id: 'c1', type: 'chapter', number: '1', title: '绪论',
          content: [{ type: 'paragraph', text: '正文。' }],
          subsections: [{
            id: 'c1-1', type: 'chapter', number: '1.1', title: '背景',
            content: [{ type: 'paragraph', text: '小节正文。' }],
          }],
        },
      ],
      backMatter: { references: [{ id: '1', text: '一条参考文献.' }] },
    },
  });
  const { zip, xml } = await documentXml(buffer);
  const sections = sectionProperties(xml);

  assert.equal(sections.length, 3, 'cover + front matter + body');
  assert.deepEqual(
    sections.map(section => /<w:headerReference/.test(section)),
    [false, true, true],
  );

  // 1. The cover carries no running head and no page number at all — the
  //    template keeps it head-less by referencing *empty* parts, and there is no
  //    `w:titlePg` anywhere in it.
  assert.doesNotMatch(sections[0], /<w:(header|footer)Reference/);
  assert.match(sections[0], /<w:pgMar[^>]*w:gutter="567"/);
  // Its own footer distance, straight from the template's 封一 (851, not 850).
  assert.match(sections[0], /<w:pgMar[^>]*w:footer="851"/);

  // 2. Front matter: Roman numerals restarting at I, and it carries the head and
  //    the newer footer distance of the 摘要 pages.
  assert.match(sections[1], /<w:pgNumType w:start="1" w:fmt="upperRoman"\/>/);
  assert.match(sections[1], /<w:pgMar[^>]*w:header="851"[^>]*w:footer="850"/);
  assert.match(sections[1], /<w:headerReference w:type="even"/);

  // 3. Body: decimal, restarting at 1 on the 引言 page.
  assert.match(sections[2], /<w:pgNumType w:start="1" w:fmt="decimal"\/>/);
  assert.match(sections[2], /<w:footerReference w:type="even"/);

  // The running heads come from the template's header parts, not from a
  // generated `学位论文` line: odd pages read the university's line, even pages
  // the thesis's own title — a template can only carry the *sample* thesis's
  // title in that slot, so it names what belongs there rather than what to
  // print. Both keep the template's 篇眉 rule.
  const headers = await runningHeadTexts(zip, 'header');
  const odd = headers.find(part => part.text === '北京科技大学硕士学位论文');
  assert.ok(odd, `expected the template's odd-page head, got ${JSON.stringify(headers.map(h => h.text))}`);
  assert.match(odd.xml, /<w:pBdr>[\s\S]*<w:bottom /, '篇眉 rule');

  const evenParts = await evenHeaderParts(zip, xml);
  assert.ok(evenParts.length > 0, 'the sections must reference an even-page head');
  for (const part of evenParts) {
    assert.equal(part.text, '节理岩体隧道纵向地震易损性分析', `${part.name} must print the document's title`);
    // The template's *formatting* for the slot survives the substitution:
    // centred, 五号 (21 half-points), with the 篇眉 rule.
    assert.match(part.xml, /<w:jc w:val="center"\/>/, 'centred');
    assert.match(part.xml, /<w:sz w:val="21"\/>/, '五号');
    assert.match(part.xml, /<w:pBdr>[\s\S]*<w:bottom /, '篇眉 rule');
  }

  // The sample thesis's title must not reach the package at all.
  const placeholder = '现代绿色化学中的物理有机问题';
  for (const name of Object.keys(zip.files).filter(n => !zip.files[n].dir && /\.(xml|rels)$/.test(n))) {
    const partXml = await zip.file(name).async('string');
    assert.ok(
      !partXml.includes(placeholder),
      `${name} still carries the template's sample title`,
    );
  }

  // The page-number parts print a PAGE field. The digits the template's own
  // footers still carry are the cached result of Word's last layout, not content.
  const footers = await runningHeadTexts(zip, 'footer');
  assert.ok(footers.length > 0);
  assert.ok(footers.every(part => /PAGE/.test(part.xml)));
  assert.ok(
    footers.every(part => part.text === ''),
    `a cached page number leaked into a footer: ${JSON.stringify(footers.map(f => f.text))}`,
  );

  // 偶数页页眉 + 对称页边距 are document settings, and the 目录 asks Word to
  // refresh the field on open.
  const settings = await zip.file('word/settings.xml').async('string');
  assert.match(settings, /<w:evenAndOddHeaders\/>/);
  assert.doesNotMatch(settings, /<w:evenAndOddHeaders w:val="false"/);
  assert.equal((settings.match(/<w:evenAndOddHeaders/g) ?? []).length, 1, 'written exactly once');
  assert.match(settings, /<w:mirrorMargins\/>/);
  assert.match(settings, /<w:updateFields w:val="true"\/>/);
  assert.doesNotMatch(xml, /mirrorMargins|evenAndOddHeaders/);

  // The guide's typography is untouched by any of this.
  assert.match(xml, /w:line="579"/);
  assert.match(xml, /w:hanging="567"/);
  assert.match(xml, /w:firstLine="480"/);
  assert.match(xml, /w:line="312"/);
});

test('leaves an even head alone when the template asks for something or for nothing', async () => {
  // The even slot is the one place the renderer overrides the template's text —
  // because a template can only carry its author's *sample* title there. Two
  // things the template means literally still win: a `{...}` request, which says
  // what it wants, and an empty part, which suppresses the head entirely.
  const thesis = title => ({
    outputPath: '', template: title,
    document: {
      type: 'thesis', meta: { title: 'Real thesis title' }, cover: [],
      sections: [{ id: 'ch', type: 'chapter', title: '第一章', content: [{ type: 'paragraph', text: '正文。' }] }],
    },
  });

  // The body part takes its geometry from `template.page`, the cover and front
  // matter from `pageSections`, so a template's even slot has to be changed in
  // both places to see the rule.
  const withEvenText = text => {
    const clone = structuredClone(template);
    if (clone.page.headers?.even) clone.page.headers.even.text = text;
    for (const section of clone.pageSections ?? []) {
      if (section.headers?.even) section.headers.even.text = text;
    }
    return clone;
  };

  const { zip: askingZip, xml: askingXml } = await documentXml(await renderDocument(thesis(withEvenText('{organization}'))));
  const asked = await evenHeaderParts(askingZip, askingXml);
  assert.ok(asked.length > 0, 'the even head must still be written');
  assert.ok(
    asked.every(part => part.text === '北京科技大学'),
    `the template asked for {organization}, got ${JSON.stringify(asked.map(p => p.text))}`,
  );

  const { zip: suppressedZip, xml: suppressedXml } = await documentXml(await renderDocument(thesis(withEvenText(''))));
  assert.equal((await evenHeaderParts(suppressedZip, suppressedXml)).length, 0, 'an empty part must stay empty');
  for (const part of await runningHeadTexts(suppressedZip, 'header')) {
    assert.notEqual(part.text, 'Real thesis title', `${part.name} revived a suppressed head`);
  }
});

test('keeps one section when the template declares no sections', async () => {
  // Backwards compatibility: a template JSON without `pageSections` — an older
  // parse, or one written by hand — still produces the single section it always
  // did, with the generated running head.
  const flat = structuredClone(template);
  delete flat.pageSections;

  const buffer = await renderDocument({
    outputPath: '', template: flat,
    document: {
      type: 'thesis', meta: { title: 'One section' },
      cover: [{ type: 'centered_text', text: '封面' }],
      sections: [
        { id: 'abs', type: 'abstract', title: '摘要', content: [] },
        { id: 'c1', type: 'chapter', title: '第一章', content: [] },
      ],
    },
  });
  const { zip, xml } = await documentXml(buffer);

  assert.equal(sectionProperties(xml).length, 1);
  const headers = await runningHeadTexts(zip, 'header');
  assert.ok(headers.some(part => part.text === '北京科技大学学位论文'), 'the generated head is kept');

  // Without any even-page head, `w:evenAndOddHeaders` must NOT be set: the flag
  // demotes `w:type="default"` to odd pages, so it would blank the even pages.
  const settings = await zip.file('word/settings.xml').async('string');
  assert.doesNotMatch(settings, /<w:evenAndOddHeaders\/>/);
});

test('writes an even-page head, and the flag it needs, from the template', async () => {
  // 偶数页页眉: the guide puts the thesis title on the even pages. A template
  // says so with a `w:type="even"` reference, and `{title}` stands for the title
  // of the document being rendered — a template cannot know it.
  const withEvenHead = structuredClone(template);
  withEvenHead.page = {
    ...withEvenHead.page,
    headers: { default: { text: '北京科技大学硕士学位论文', rule: true }, even: { text: '{title}' } },
    footers: { default: { pageNumber: true } },
  };
  withEvenHead.pageSections = [
    { width: 11906, height: 16838, margins: { top: 1701, bottom: 1134, left: 1701, right: 1701 }, columns: 1 },
    { ...withEvenHead.page },
  ];
  // The template asks for it; the flag must reach settings.xml — as `true`, not
  // as the `w:val="false"` the writer library always emits.
  withEvenHead.evenAndOddHeaders = true;

  const buffer = await renderDocument({
    outputPath: '', template: withEvenHead,
    document: {
      type: 'thesis', meta: { title: '节理岩体隧道' }, cover: [{ type: 'centered_text', text: '封面' }],
      sections: [{ id: 'c1', type: 'chapter', title: '第一章', content: [] }],
    },
  });
  const { zip, xml } = await documentXml(buffer);

  assert.match(xml, /<w:headerReference w:type="even"/);
  const headers = await runningHeadTexts(zip, 'header');
  assert.ok(headers.some(part => part.text === '节理岩体隧道'), "the even head carries this document's title");
  assert.ok(headers.some(part => part.text === '北京科技大学硕士学位论文'));

  const settings = await zip.file('word/settings.xml').async('string');
  assert.match(settings, /<w:evenAndOddHeaders\/>/);
  assert.doesNotMatch(settings, /<w:evenAndOddHeaders w:val="false"/);
  assert.equal((settings.match(/<w:evenAndOddHeaders/g) ?? []).length, 1);
});

test('drops the explicit page break a part ends with', async () => {
  // The cover in the sample content ends with a `page_break`, left over from when
  // the whole document was one section. With a section break after it that would
  // print an empty page.
  const buffer = await renderDocument({
    outputPath: '', template,
    document: {
      type: 'thesis', meta: { title: 'Break' },
      cover: [{ type: 'centered_text', text: '封面行' }, { type: 'page_break' }],
      sections: [{ id: 'abs', type: 'abstract', title: '摘要', content: [{ type: 'paragraph', text: '摘要。' }] }],
    },
  });
  const { xml } = await documentXml(buffer);
  const sections = sectionProperties(xml);

  assert.equal(sections.length, 2);
  // One break — the section break. The cover's own `w:br w:type="page"` is gone.
  assert.equal((xml.match(/<w:br w:type="page"\/>/g) ?? []).length, 0);
  assert.doesNotMatch(sections[0], /<w:(header|footer)Reference/);
  assert.match(sections[1], /<w:headerReference w:type="default"/);
});

test('turns on titlePg for a template that references a first-page head', async () => {
  // A `w:type="first"` reference is inert without `w:titlePg`, so the section has
  // to ask for it — the writer library exposes it as `titlePage`.
  const withFirstHead = structuredClone(template);
  withFirstHead.page = {
    ...withFirstHead.page,
    headers: { first: { text: '内部资料' }, default: { text: '北京科技大学硕士学位论文' } },
    footers: { default: { pageNumber: true }, first: { pageNumber: false, text: '' } },
  };
  withFirstHead.pageSections = [
    { width: 11906, height: 16838, margins: { top: 1701, bottom: 1134, left: 1701, right: 1701 }, columns: 1 },
    { ...withFirstHead.page },
  ];

  const buffer = await renderDocument({
    outputPath: '', template: withFirstHead,
    document: {
      type: 'thesis', meta: { title: 'First page' }, cover: [{ type: 'centered_text', text: '封面' }],
      sections: [{ id: 'c1', type: 'chapter', title: '第一章', content: [] }],
    },
  });
  const { zip, xml } = await documentXml(buffer);
  const sections = sectionProperties(xml);

  assert.match(xml, /<w:titlePg\/>/);
  assert.match(xml, /<w:headerReference w:type="first"/);
  // An *empty* part stays empty: no reference, so the first page keeps the head
  // it inherits rather than printing a blank one.
  assert.doesNotMatch(xml, /<w:footerReference w:type="first"/);
  assert.doesNotMatch(sections[0], /<w:titlePg\/>/, 'the cover is not a title page');
  const headers = await runningHeadTexts(zip, 'header');
  assert.ok(headers.some(part => part.text === '内部资料'));
});

test('academic tables default to standard 三线表 (three-line tables)', async () => {
  const buffer = await renderDocument({
    outputPath: '', template,
    document: {
      type: 'thesis', meta: { title: 'Three line table' }, cover: [],
      sections: [{
        id: 'c1', type: 'chapter', title: 'Chapter',
        content: [{
          type: 'table', caption: 'Test Table',
          headers: ['A', 'B'],
          data: [['1', '2'], ['3', '4']],
        }],
      }],
    },
  });
  const { xml } = await documentXml(buffer);

  // Table borders: top 1.5pt (12), bottom 1.5pt (12), no left/right/inside verticals
  assert.match(xml, /<w:tblBorders><w:top w:val="single" w:color="000000" w:sz="12"\/>/);
  // An absent rule is a white hairline: nil/none make LibreOffice draw its
  // default grid instead of nothing.
  assert.match(xml, /<w:left w:val="single" w:color="FFFFFF" w:sz="2"\/>/);
  assert.match(xml, /<w:insideV w:val="single" w:color="FFFFFF" w:sz="2"\/>/);
  assert.match(xml, /<w:bottom w:val="single" w:color="000000" w:sz="12"\/>/);
  // An absent rule is a white hairline: nil/none make LibreOffice draw its
  // default grid instead of nothing.
  assert.match(xml, /<w:left w:val="single" w:color="FFFFFF" w:sz="2"\/>/);
  assert.match(xml, /<w:insideV w:val="single" w:color="FFFFFF" w:sz="2"\/>/);
  // An absent rule is a white hairline: nil/none make LibreOffice draw its
  // default grid instead of nothing.
  assert.match(xml, /<w:left w:val="single" w:color="FFFFFF" w:sz="2"\/>/);
  assert.match(xml, /<w:insideV w:val="single" w:color="FFFFFF" w:sz="2"\/>/);

  // Header bottom border: 0.75pt (6)
  assert.match(xml, /<w:bottom w:val="single" w:color="000000" w:sz="6"\/>/);

  // No gray header shading by default
  assert.doesNotMatch(xml, /w:fill="D9D9D9"/);
});

test('numbered equations place equation in center and number right-aligned via tab stops', async () => {
  const buffer = await renderDocument({
    outputPath: '', template,
    document: {
      type: 'thesis', meta: { title: 'Equation Tab Stops' }, cover: [],
      sections: [{
        id: 'c1', type: 'chapter', title: 'Chapter',
        content: [{
          type: 'equation_numbered',
          latex: 'E = mc^2',
          number: '1-1',
        }],
      }],
    },
  });
  const { xml } = await documentXml(buffer);

  // Tab stops: center tab at page middle, right tab at page right margin
  assert.match(xml, /<w:tabs><w:tab w:val="center" w:pos="\d+"\s*\/><w:tab w:val="right" w:pos="\d+"\s*\/><\/w:tabs>/);
  // Left alignment so tabs function properly
  assert.match(xml, /<w:jc w:val="left"\/>/);
  // Tab-separated equation and right-aligned number
  assert.match(xml, /<w:t xml:space="preserve">\s*<\/w:t><\/w:r><m:oMath>/);
  assert.match(xml, /<w:t xml:space="preserve">\t\(1-1\)<\/w:t>/);
});

test('renders declaration and authorBiography in backMatter', async () => {
  const buffer = await renderDocument({
    outputPath: '', template,
    document: {
      type: 'thesis', meta: { title: 'BackMatter' }, cover: [],
      sections: [{ id: 'c1', type: 'chapter', title: 'Chapter', content: [] }],
      backMatter: {
        declaration: '本人郑重声明，所呈交的学位论文是本人在导师指导下进行的研究工作及取得的研究成果。',
        authorBiography: '张三，男，1998年生。在学期间发表论文1篇。',
      },
    },
  });
  const { xml } = await documentXml(buffer);

  assert.match(xml, /独创性说明/);
  assert.match(xml, /本人郑重声明/);
  assert.match(xml, /作者简历及在学期间取得的成果/);
  assert.match(xml, /在学期间发表论文/);
});

test('renders inline LaTeX math as native Office Math (OMML) within paragraphs', async () => {
  const buffer = await renderDocument({
    outputPath: '', template,
    document: {
      type: 'thesis', meta: { title: 'Inline math' }, cover: [],
      sections: [{
        id: 'c1', type: 'chapter', title: 'Chapter',
        content: [{
          type: 'paragraph',
          text: '时间步长 $\\Delta t_{\\text{cr}}$ 与矩阵 $\\mathbf{M}$，以及式中 $K_{ij}$ 为节点 $i$。',
        }],
      }],
    },
  });
  const { xml } = await documentXml(buffer);

  // Contains both text runs and inline Office Math
  assert.match(xml, /<w:t xml:space="preserve">时间步长 <\/w:t>/);
  assert.match(xml, /<m:oMath>/);
  assert.match(xml, /<m:t>Δ<\/m:t>/);
  assert.match(xml, /<m:t>cr<\/m:t>/);
  assert.match(xml, /<m:t>M<\/m:t>/);
  assert.match(xml, /<m:t>ij<\/m:t>/);
  // Ensure the raw dollar signs and raw LaTeX code are NOT emitted as literal text
  assert.doesNotMatch(xml, /\\Delta/);
  assert.doesNotMatch(xml, /\\mathbf/);
  assert.doesNotMatch(xml, /\$K_\{ij\}\$/);
});

// ── Regression: the fixes from the follow-up audit ───────────

test('a table with no caption renders instead of throwing', async () => {
  // `captionText` used `.trim()` on an undefined caption, so a document that
  // skipped `validateDocument` killed the build with
  // `Cannot read properties of undefined (reading 'trim')`.
  const buffer = await renderDocument({
    template,
    document: {
      type: 'thesis', meta: { title: 'No caption' }, cover: [],
      sections: [{
        id: 'c1', type: 'chapter', title: 'Chapter',
        content: [{ type: 'table', headers: ['A', 'B'], data: [['1', '2']] }],
      }],
    },
  });
  const { xml } = await documentXml(buffer);
  assert.match(xml, /<w:t[^>]*>1<\/w:t>/);
});

test('the emitted package is deflate-compressed, not stored', async () => {
  // The renderer used to finish with a second `generateAsync` that passed no
  // `compression` option, i.e. STORE. Every output was the uncompressed sum of
  // its parts — 188.9 KB where 136.8 KB was achievable.
  const buffer = await renderDocument({
    template,
    document: {
      type: 'thesis', meta: { title: 'Compression' }, cover: [],
      sections: [{
        id: 'c1', type: 'chapter', title: 'Chapter',
        content: [{ type: 'paragraph', text: '正文段落 '.repeat(200) }],
      }],
    },
  });
  const zip = await JSZip.loadAsync(buffer);
  const stored = await zip.generateAsync({ type: 'nodebuffer', compression: 'STORE' });
  assert.ok(
    buffer.length < stored.length * 0.9,
    `expected a compressed package, got ${buffer.length} bytes against ${stored.length} stored`,
  );
});

test('a template mutated between renders is read fresh', async () => {
  // `resolveStyle` is cached per template object. The cache is dropped at the
  // start of every render, so a caller that edits its template in place — as
  // the template-authoring flow does — still sees the edit.
  const doc = {
    type: 'thesis', meta: { title: 'Cache' }, cover: [],
    sections: [{
      id: 'c1', type: 'chapter', title: 'Chapter',
      content: [{ type: 'paragraph', text: 'Body' }],
    }],
  };
  const mutable = structuredClone(template);
  const before = await documentXml(await renderDocument({ template: mutable, document: doc }));
  mutable.styles[mutable.roleWinners.paragraph].font.size = 44;
  const after = await documentXml(await renderDocument({ template: mutable, document: doc }));

  assert.match(before.xml, /<w:sz w:val="24"\/>/);
  assert.match(after.xml, /<w:sz w:val="44"\/>/);
});

test('nested subsections keep their heading level', async () => {
  // `flattenSections` was written twice — once for theses, once for journals —
  // character for character identical. Both now go through one helper, and this
  // pins the depth-to-heading mapping it must preserve.
  //
  // The renderer bakes each role's formatting into the paragraph instead of
  // emitting a `w:pStyle` reference, so the level is observed through the font
  // size the role resolves to. The shipped template sizes heading2 and heading3
  // identically, so this clone gives every level its own size.
  const tpl = structuredClone(template);
  const sizes = { heading1: 30, heading2: 26, heading3: 22 };
  for (const [role, size] of Object.entries(sizes)) {
    tpl.styles[tpl.roleWinners[role]].font.size = size;
  }

  const buffer = await renderDocument({
    template: tpl,
    document: {
      type: 'thesis', meta: { title: 'Nesting' }, cover: [],
      sections: [{
        id: 'c1', type: 'chapter', title: 'Level one',
        content: [],
        subsections: [{
          id: 'c2', type: 'chapter', title: 'Level two',
          content: [],
          subsections: [{
            id: 'c3', type: 'chapter', title: 'Level three',
            content: [{ type: 'paragraph', text: 'Leaf' }],
          }],
        }],
      }],
    },
  });
  const { xml } = await documentXml(buffer);

  const sizeOfHeading = (title) => {
    const at = xml.indexOf(title);
    assert.ok(at >= 0, `expected "${title}" in the output`);
    const lastSize = [...xml.slice(0, at).matchAll(/<w:sz w:val="(\d+)"\/>/g)].pop();
    return Number(lastSize?.[1]);
  };

  assert.equal(sizeOfHeading('Level one'), 30);
  assert.equal(sizeOfHeading('Level two'), 26);
  assert.equal(sizeOfHeading('Level three'), 22);
});

test('a journal body takes its typography from the template, not from the thesis defaults', async () => {
  // The journal path called `processBlocks` with `docType: undefined`, so
  // `resolveStyle` never consulted the template and a journal rendered from a
  // journal template came out in the thesis fallback styles.
  const tpl = structuredClone(template);
  const headingId = tpl.roleWinners.heading1;
  tpl.styles[headingId].font.size = 37;

  const buffer = await renderDocument({
    template: tpl,
    document: {
      type: 'journal',
      meta: {
        title: 'A paper',
        authors: [{ name: 'A', affiliations: [{ institution: 'Somewhere' }] }],
        keywords: ['one'],
      },
      sections: [{
        id: 'intro', type: 'introduction', title: 'Introduction',
        content: [{ type: 'paragraph', text: 'Body.' }],
      }],
    },
  });
  const { xml } = await documentXml(buffer);

  const at = xml.indexOf('Introduction');
  assert.ok(at >= 0, 'expected the section heading in the output');
  const lastSize = [...xml.slice(0, at).matchAll(/<w:sz w:val="(\d+)"\/>/g)].pop();
  assert.equal(Number(lastSize?.[1]), 37);
});

test('a heading that declares w:keepNext carries it into the document', async () => {
  // `keepNext` / `keepLines` were parsed by the template parser and merged into
  // every resolved style, but `ParagraphFormatting` had no field for them, so
  // they were dropped before the renderer could act — and a heading left alone
  // at the foot of a page is exactly what `keepNext` prevents.
  const tpl = structuredClone(template);
  const headingId = tpl.roleWinners.heading1;
  tpl.styles[headingId].paragraph.keepNext = true;
  tpl.styles[headingId].paragraph.keepLines = true;

  const buffer = await renderDocument({
    template: tpl,
    document: {
      type: 'thesis', meta: { title: 'Keep' }, cover: [],
      sections: [{
        id: 'c1', type: 'chapter', title: 'Kept heading',
        content: [{ type: 'paragraph', text: 'Body.' }],
      }],
    },
  });
  const { xml } = await documentXml(buffer);
  assert.match(xml, /<w:keepNext\/>/);
  assert.match(xml, /<w:keepLines\/>/);
});

test('a table cell keeps the template line-spacing rule', async () => {
  // Table cells passed `line` without `lineRule`, so an `exact` twips value was
  // read as a multiple of the line height — the only place the template's own
  // spacing was discarded instead of honoured.
  const tpl = structuredClone(template);
  const tableId = tpl.roleWinners.table ?? Object.keys(tpl.styles)[0];
  tpl.styleRoles[tableId] = 'table';
  tpl.roleWinners.table = tableId;
  tpl.styles[tableId].lineSpacing = 300;
  tpl.styles[tableId].lineSpacingRule = 'exact';

  const buffer = await renderDocument({
    template: tpl,
    document: {
      type: 'thesis', meta: { title: 'Cells' }, cover: [],
      sections: [{
        id: 'c1', type: 'chapter', title: 'Chapter',
        content: [{ type: 'table', caption: 'T', headers: ['A'], data: [['1']] }],
      }],
    },
  });
  const { xml } = await documentXml(buffer);
  assert.match(xml, /w:line="300" w:lineRule="exact"/);
});

test('a caller-supplied figure width is clamped to the printable page', async () => {
  // Only the no-size branch clamped; a `width` from the content JSON reached the
  // document unclamped and ran off the right margin.
  const contentDir = mkdtempSync(join(tmpdir(), 'openthesis-figure-'));
  writeFileSync(
    join(contentDir, 'wide.png'),
    Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAFUlEQVR42mNk+M9Qz0AEYBxVSF+FAP5FDvcfRYWgAAAAAElFTkSuQmCC', 'base64'),
  );

  const buffer = await renderDocument({
    template,
    contentDir,
    document: {
      type: 'thesis', meta: { title: 'Figure' }, cover: [],
      sections: [{
        id: 'c1', type: 'chapter', title: 'Chapter',
        content: [{ type: 'figure', path: 'wide.png', caption: 'Wide', width: 4000 }],
      }],
    },
  });
  const { xml } = await documentXml(buffer);
  // A4 minus the template's 3 cm margins is 851 twips short of the page width;
  // 1 px = 9525 EMU, so the printable width in px is about 600.
  const extent = xml.match(/<wp:extent cx="(\d+)" cy="(\d+)"\/>/);
  assert.ok(extent, 'expected a drawing extent');
  const widthPx = Number(extent[1]) / 9525;
  assert.ok(widthPx <= 620, `expected the width to be clamped, got ${widthPx.toFixed(0)} px`);
  assert.ok(Number(extent[2]) > 0, 'expected a positive height');
});

test('the package carries the template TOC entry styles', async () => {
  // The renderer emits a real TOC field; Word builds its entries with the
  // TOC1/TOC2/TOC3 styles. Those ids did not exist in the generated package, so
  // the template's own directory typography was parsed into `toc_entry` and then
  // never applied.
  const tpl = structuredClone(template);
  tpl.styles.TOC1.font.size = 27;
  tpl.styles.TOC1.paragraph.alignment = 'left';

  const buffer = await renderDocument({
    template: tpl,
    document: {
      type: 'thesis', meta: { title: 'Contents' }, cover: [],
      sections: [
        { id: 'toc', type: 'toc', title: '目录', content: [] },
        { id: 'c1', type: 'chapter', title: 'Chapter', content: [{ type: 'paragraph', text: 'Body' }] },
      ],
    },
  });
  const zip = await JSZip.loadAsync(buffer);
  const stylesXml = await zip.file('word/styles.xml').async('string');

  const toc1 = stylesXml.match(/<w:style[^>]*w:styleId="TOC1"[\s\S]*?<\/w:style>/)?.[0];
  assert.ok(toc1, 'expected a TOC1 paragraph style in the package');
  assert.match(toc1, /<w:sz w:val="27"\/>/);
  assert.match(stylesXml, /w:styleId="TOC2"/);
  assert.match(stylesXml, /w:styleId="TOC3"/);
});

test('the template document grid replaces the writer library default', async () => {
  // The writer always emits its own `linePitch="360"`. A template that snaps text
  // to a different grid gets its own values written into the generated section
  // properties, or a page breaks lines where the template would not.
  const tpl = structuredClone(template);
  tpl.page.docGrid = { type: 'linesAndChars', linePitch: 312, charSpace: 0 };

  const buffer = await renderDocument({
    template: tpl,
    document: {
      type: 'thesis', meta: { title: 'Grid' }, cover: [],
      sections: [{ id: 'c1', type: 'chapter', title: 'Chapter', content: [{ type: 'paragraph', text: 'Body' }] }],
    },
  });
  const { xml } = await documentXml(buffer);
  assert.match(xml, /<w:docGrid w:type="linesAndChars" w:linePitch="312" w:charSpace="0"\/>/);
  assert.doesNotMatch(xml, /w:linePitch="360"/);
});

test('a journal front matter resolves through the template instead of hardcoded fonts', async () => {
  // The title, authors, affiliations, abstract and keywords were built from
  // literal font/size/spacing values, so a journal rendered from a journal
  // template ignored that template entirely.
  const tpl = structuredClone(template);
  // Give each journal role a unique size through the style the role resolves to.
  tpl.styles.JournalTitle = {
    font: { name: 'Times New Roman', eastAsia: '黑体', size: 41, bold: true },
    paragraph: { alignment: 'left', spaceBefore: 0, spaceAfter: 0 },
    lineSpacing: 300,
  };
  tpl.styleRoles.JournalTitle = 'journal_title';
  tpl.roleWinners.journal_title = 'JournalTitle';

  const buffer = await renderDocument({
    template: tpl,
    document: {
      type: 'journal',
      meta: {
        title: 'A title',
        authors: [{ name: 'A. Author', affiliations: [{ institution: 'Somewhere' }] }],
        keywords: ['one', 'two'],
      },
      sections: [{ id: 'intro', type: 'introduction', title: 'Introduction', content: [] }],
    },
  });
  const { xml } = await documentXml(buffer);

  const at = xml.indexOf('A title');
  assert.ok(at >= 0, 'expected the article title in the output');
  const lastSize = [...xml.slice(0, at).matchAll(/<w:sz w:val="(\d+)"\/>/g)].pop();
  assert.equal(Number(lastSize?.[1]), 41, 'the title should carry the template role size');
});

test('the font guard leaves fonts the template declares alone', async () => {
  // The guard used to strip every `*Theme` attribute and rewrite every Latin
  // `w:eastAsia` to 仿宋 across styles.xml and document.xml. Harmless for a
  // template that says nothing, hostile to one that says exactly what it wants.
  // It is now limited to `w:docDefaults`, where the library bug actually lives.
  const tpl = structuredClone(template);
  tpl.styles.TOC1.font.eastAsia = 'Arial';
  tpl.styles.TOC1.font.size = 24;

  const buffer = await renderDocument({
    template: tpl,
    document: {
      type: 'thesis', meta: { title: 'Fonts' }, cover: [],
      sections: [
        { id: 'toc', type: 'toc', title: '目录', content: [] },
        { id: 'c1', type: 'chapter', title: 'Chapter', content: [{ type: 'paragraph', text: '正文' }] },
      ],
    },
  });
  const zip = await JSZip.loadAsync(buffer);
  const stylesXml = await zip.file('word/styles.xml').async('string');
  const toc1 = stylesXml.match(/<w:style[^>]*w:styleId="TOC1"[\s\S]*?<\/w:style>/)?.[0];
  assert.ok(toc1, 'expected a TOC1 style');
  // Arial is what the template asked for; the guard must not turn it into 仿宋.
  assert.match(toc1, /w:eastAsia="Arial"/);
});

test('a table row with the wrong cell count warns instead of padding in silence', async () => {
  // `validateDocument` reports the mismatch, but a caller using the renderer
  // directly got a silently padded or truncated row.
  const warnings = [];
  const original = console.warn;
  console.warn = (message) => warnings.push(String(message));
  try {
    await renderDocument({
      template,
      document: {
        type: 'thesis', meta: { title: 'Rows' }, cover: [],
        sections: [{
          id: 'c1', type: 'chapter', title: 'Chapter',
          content: [{
            type: 'table', caption: '短行', headers: ['A', 'B'],
            data: [['only one'], ['1', '2', 'extra']],
          }],
        }],
      },
    });
  } finally {
    console.warn = original;
  }

  const short = warnings.find(w => w.includes('row 1') && w.includes('Missing cells'));
  const long = warnings.find(w => w.includes('row 2') && w.includes('Extra cells'));
  assert.ok(short, `expected a missing-cells warning, got: ${JSON.stringify(warnings)}`);
  assert.ok(long, `expected a dropped-cells warning, got: ${JSON.stringify(warnings)}`);
});

test('a spacer can be an exact height, so a cover gap is a measurement', async () => {
  // The guide fixes 5 cm between the cover tiers. Five blank lines is only as
  // tall as the line height in force, which is how the spacing got lost.
  const buffer = await renderDocument({
    template,
    document: {
      type: 'thesis', meta: { title: 'Spacer' },
      cover: [
        { type: 'centered_text', text: '校名' },
        { type: 'spacer', cm: 5 },
        { type: 'centered_text', text: '标题' },
      ],
      sections: [{ id: 'c1', type: 'chapter', title: 'Chapter', content: [] }],
    },
  });
  const { xml } = await documentXml(buffer);
  // 5 cm = 50 mm; at 1440 twips per inch that is 2834 twips (the library rounds
  // down).
  assert.match(xml, /w:line="2834" w:lineRule="exact"/);
});

test('a line-count spacer still renders one blank paragraph per line', async () => {
  const buffer = await renderDocument({
    template,
    document: {
      type: 'thesis', meta: { title: 'Spacer lines' },
      cover: [{ type: 'spacer', lines: 3 }],
      sections: [{ id: 'c1', type: 'chapter', title: 'Chapter', content: [] }],
    },
  });
  const { xml } = await documentXml(buffer);
  const blank = [...xml.matchAll(/<w:spacing w:line="312"\/>/g)].length;
  assert.equal(blank, 3);
});

test('a heading bound to template numbering is numbered, and a written number wins', async () => {
  // The rule: a `number` in the content file is the author's exact label and is
  // kept; otherwise the template's own numbering applies.
  const tpl = structuredClone(template);
  tpl.styles.NumHeading = {
    font: { name: 'Times New Roman', eastAsia: '黑体', size: 30, bold: true },
    paragraph: { alignment: 'left', numbering: { numId: 7, ilvl: 0 } },
    lineSpacing: 312,
  };
  tpl.styleRoles.NumHeading = 'heading1';
  tpl.roleWinners.heading1 = 'NumHeading';
  tpl.numbering = [{
    numId: 7,
    abstractId: 3,
    levels: [{ ilvl: 0, format: 'decimal', text: '第%1章', start: 1, pStyle: 'NumHeading' }],
  }];

  const buffer = await renderDocument({
    template: tpl,
    document: {
      type: 'thesis', meta: { title: 'Numbering' }, cover: [],
      sections: [{
        id: 'c1', type: 'chapter', title: 'Auto numbered', content: [
          { type: 'heading2', text: 'Manual', number: '2.9' },
        ],
      }],
    },
  });
  const zip = await JSZip.loadAsync(buffer);
  const doc = await zip.file('word/document.xml').async('string');
  const numberingXml = await zip.file('word/numbering.xml').async('string');

  // The template's level text reaches the package.
  assert.match(numberingXml, /第%1章/);

  const headingAt = doc.indexOf('Auto numbered');
  const paragraphBefore = doc.slice(doc.lastIndexOf('<w:p>', headingAt), headingAt);
  assert.match(paragraphBefore, /<w:numPr>/, 'the bound heading should carry a numbering reference');

  // The manual number is rendered as text and does not also take auto numbering.
  const manualAt = doc.indexOf('2.9');
  assert.ok(manualAt >= 0, 'the written number should be in the text');
  const manualParagraph = doc.slice(doc.lastIndexOf('<w:p>', manualAt), doc.indexOf('</w:p>', manualAt));
  assert.doesNotMatch(manualParagraph, /<w:numPr>/);
});

test('a section break inside the body starts a new section with its own page setup', async () => {
  // A chapter that turns one passage landscape, or sets it in two columns, is a
  // section change inside the body. It used to be flattened into whichever part
  // it fell in, so the passage came out portrait and single-column.
  const buffer = await renderDocument({
    template,
    document: {
      type: 'thesis', meta: { title: 'Layout' }, cover: [{ type: 'centered_text', text: '封面' }],
      sections: [{
        id: 'c1', type: 'chapter', title: 'Chapter One',
        content: [
          { type: 'paragraph', text: 'portrait body' },
          { type: 'section_break', page: { orientation: 'landscape', columns: 2, columnGutter: 425 } },
          { type: 'paragraph', text: 'landscape two columns' },
        ],
      }],
    },
  });
  const { xml } = await documentXml(buffer);
  const sections = [...xml.matchAll(/<w:sectPr[\s\S]*?<\/w:sectPr>/g)].map(m => m[0]);

  const landscape = sections.find(s => /w:orient="landscape"/.test(s));
  assert.ok(landscape, 'expected a landscape section');
  // The writer swaps the dimensions for landscape; the template's A4 becomes
  // 16838 x 11906.
  assert.match(landscape, /<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"\/>/);
  assert.match(landscape, /<w:cols w:space="425" w:num="2"\/>/);

  // The passage after the break is in that section; the one before it is not.
  const landscapeAt = xml.indexOf('landscape two columns');
  const portraitAt = xml.indexOf('portrait body');
  assert.ok(portraitAt >= 0 && landscapeAt > portraitAt);
  const sectionStart = xml.lastIndexOf('<w:sectPr', landscapeAt);
  assert.ok(sectionStart > portraitAt, 'the landscape section starts after the portrait passage');
});

// ── Real-world template rendering ────────────────────────────

test('renders a document with each real-world template', async () => {
  // Parsing a template is only half the promise. These four were downloaded
  // from their publishers, so this exercises the renderer against geometry,
  // fonts and style tables it did not author.
  const fixtureDir = join(dirname(fileURLToPath(import.meta.url)), 'fixtures', 'templates');
  const doc = {
    type: 'thesis', meta: { title: 'Compat', degree: 'master', date: '2026' }, cover: [],
    sections: [{
      id: 'c1', type: 'chapter', title: 'Chapter One',
      content: [
        { type: 'heading2', text: 'Section A' },
        { type: 'paragraph', text: 'Body paragraph.' },
        { type: 'table', caption: 'Table 1', headers: ['A', 'B'], data: [['1', '2']] },
        { type: 'equation', latex: 'E = mc^2', number: '1-1' },
      ],
    }],
  };

  for (const name of ['acm-journal.docx', 'cn-cjstp.docx', 'cn-dlmu-thesis.docx', 'elsevier-hardwarex.docx']) {
    const template = await parseTemplate(readFileSync(join(fixtureDir, name)), { documentType: 'thesis' });
    const buffer = await renderDocument({ template, document: doc });
    const { xml } = await documentXml(buffer);

    assert.match(xml, /<w:t[^>]*>Body paragraph\.<\/w:t>/, `${name}: body text missing`);
    assert.match(xml, /<m:oMath>/, `${name}: equation missing`);
    // A page's width comes from the template, so the table has to fit inside it
    // rather than the old hardcoded 8504 twips.
    // The writer emits the attributes in this order: type first, then width.
    const tableWidth = Number(xml.match(/<w:tblW w:type="dxa" w:w="(\d+)"\/>/)?.[1] ?? 0);
    const pageWidth = template.page.width - template.page.margins.left - template.page.margins.right;
    assert.ok(tableWidth > 0 && tableWidth <= pageWidth + 2, `${name}: table width ${tableWidth} exceeds the printable ${pageWidth}`);
  }
});

// ── Template table/image format and caption cross-references ──

test('a table takes its borders from the template that declares them', async () => {
  // The renderer hardcoded one border shape, so every template produced the same
  // table. These two fixtures genuinely disagree: 大连海事 draws 1.5 pt rules,
  // 中国科技期刊研究 draws 0.5 pt, Elsevier draws a full grid.
  const fixtureDir = join(dirname(fileURLToPath(import.meta.url)), 'fixtures', 'templates');
  const doc = {
    type: 'thesis', meta: { title: 'Borders', degree: 'master', date: '2026' }, cover: [],
    sections: [{ id: 'c1', type: 'chapter', title: 'C', content: [
      { type: 'table', caption: 'T', headers: ['A', 'B'], data: [['1', '2']] },
    ] }],
  };

  const expected = {
    'cn-dlmu-thesis.docx': { top: '12', grid: false },
    'cn-cjstp.docx': { top: '4', grid: false },
    'elsevier-hardwarex.docx': { top: '4', grid: true },
  };

  for (const [name, want] of Object.entries(expected)) {
    const template = await parseTemplate(readFileSync(join(fixtureDir, name)), { documentType: 'thesis' });
    const buffer = await renderDocument({ template, document: doc });
    const { xml } = await documentXml(buffer);
    const borders = xml.match(/<w:tblBorders>[\s\S]*?<\/w:tblBorders>/)?.[0] ?? '';

    assert.match(borders, new RegExp(`<w:top w:val="single" w:color="000000" w:sz="${want.top}"`), `${name}: top rule`);
    if (want.grid) {
      assert.match(borders, /<w:insideV w:val="single" w:color="000000"/, `${name}: expected a full grid`);
    } else {
      assert.doesNotMatch(borders, /<w:insideV w:val="single" w:color="000000"/, `${name}: expected no vertical rules`);
      assert.match(borders, /<w:insideV w:val="single" w:color="FFFFFF"/, `${name}: missing rules are white hairlines`);
    }
  }
});

test('the parser records the table and figure format a template declares', async () => {
  const fixtureDir = join(dirname(fileURLToPath(import.meta.url)), 'fixtures', 'templates');
  const cjstp = await parseTemplate(readFileSync(join(fixtureDir, 'cn-cjstp.docx')), { documentType: 'thesis' });
  assert.equal(cjstp.tableFormat.borders.top.size, 4);
  assert.equal(cjstp.tableFormat.alignment, 'center');
  assert.deepEqual(cjstp.tableFormat.columnWidths?.slice(0, 2), [1704, 1704]);
  assert.equal(cjstp.imageFormat.placement, 'inline');

  const hardwarex = await parseTemplate(readFileSync(join(fixtureDir, 'elsevier-hardwarex.docx')), { documentType: 'thesis' });
  assert.equal(hardwarex.tableFormat.borders.insideV.style, 'single');
});

test('a [[表3-1]] in the body becomes a REF field pointing at the caption bookmark', async () => {
  // Typing 如表3-1所示 makes the number plain text: insert a table and the
  // sentence is wrong. The standard mechanism is a REF field over a bookmark on
  // the caption, which Word recomputes.
  const buffer = await renderDocument({
    template,
    document: {
      type: 'thesis', meta: { title: 'Refs', degree: 'master', date: '2026' }, cover: [],
      sections: [{ id: 'c1', type: 'chapter', title: 'C', content: [
        { type: 'paragraph', text: '如[[表3-1]]所示。' },
        { type: 'table', caption: '参数一览', number: '表3-1', headers: ['A', 'B'], data: [['1', '2']] },
      ] }],
    },
  });
  const { xml } = await documentXml(buffer);

  const bookmark = xml.match(/<w:bookmarkStart[^>]*w:name="(_Ref[^"]+)"/)?.[1];
  const target = refTarget(xml);
  assert.ok(bookmark, 'expected a caption bookmark');
  assert.ok(target, 'expected a REF field');
  assert.equal(bookmark, target, 'the field must point at the bookmark the caption created');

  // A viewer that does not evaluate fields still shows the number: the field
  // carries its result between `separate` and `end`.
  const result = /<w:fldChar w:fldCharType="separate"[\s\S]{0,24}?<w:t[^>]*>([^<]*)<\/w:t>/.exec(xml)?.[1];
  assert.equal(result, '表3-1');
});

test('a longhand caption number is left as text, not converted to a field', async () => {
  // A manuscript that writes 如表3-1所示 itself keeps working exactly as before.
  const buffer = await renderDocument({
    template,
    document: {
      type: 'thesis', meta: { title: 'Plain', degree: 'master', date: '2026' }, cover: [],
      sections: [{ id: 'c1', type: 'chapter', title: 'C', content: [
        { type: 'paragraph', text: '如表3-1所示。' },
      ] }],
    },
  });
  const { xml } = await documentXml(buffer);
  assert.equal(refTarget(xml), undefined, 'no REF field should be emitted');
  assert.match(xml, /如表3-1所示。/);
});

// ── Table and figure formatting, round two ───────────────────

test('a table uses the template column widths when the column count matches', async () => {
  // The template's own grid only means something at the same column count: its
  // 12-column example says nothing about a 2-column table.
  const tpl = structuredClone(template);
  tpl.tableFormat = {
    alignment: 'center',
    layout: 'fixed',
    columnWidths: [3000, 2000, 1500],
  };

  const buffer = await renderDocument({
    template: tpl,
    document: {
      type: 'thesis', meta: { title: 'Widths', degree: 'master', date: '2026' }, cover: [],
      sections: [{ id: 'c1', type: 'chapter', title: 'C', content: [
        { type: 'table', caption: '三列', headers: ['A', 'B', 'C'], data: [['1', '2', '3']] },
        { type: 'table', caption: '两列', headers: ['A', 'B'], data: [['1', '2']] },
      ] }],
    },
  });
  const { xml } = await documentXml(buffer);
  const grids = [...xml.matchAll(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/g)]
    .map(m => [...m[0].matchAll(/w:w="(\d+)"/g)].map(x => Number(x[1])));

  assert.deepEqual(grids[0], [3000, 2000, 1500], 'the matching table takes the template grid');
  assert.equal(grids[1].length, 2);
  assert.ok(Math.abs(grids[1][0] - grids[1][1]) <= 1, 'an unmatched table splits the printable width evenly');

  // Alignment and layout reach `w:tblPr`.
  assert.match(xml, /<w:jc w:val="center"\/>/);
  assert.match(xml, /<w:tblLayout w:type="fixed"\/>/);
});

test('a figure takes its alignment from the template', async () => {
  const contentDir = mkdtempSync(join(tmpdir(), 'openthesis-fig-align-'));
  writeFileSync(
    join(contentDir, 'pic.png'),
    Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAFUlEQVR42mNk+M9Qz0AEYBxVSF+FAP5FDvcfRYWgAAAAAElFTkSuQmCC', 'base64'),
  );

  const tpl = structuredClone(template);
  tpl.imageFormat = { alignment: 'left', placement: 'inline' };

  const buffer = await renderDocument({
    template: tpl,
    contentDir,
    document: {
      type: 'thesis', meta: { title: 'Fig', degree: 'master', date: '2026' }, cover: [],
      sections: [{ id: 'c1', type: 'chapter', title: 'C', content: [
        { type: 'figure', path: 'pic.png', caption: '图', number: '图1' },
      ] }],
    },
  });
  const { xml } = await documentXml(buffer);
  const drawingAt = xml.indexOf('<w:drawing>');
  assert.ok(drawingAt > 0);
  const paragraph = xml.slice(xml.lastIndexOf('<w:p>', drawingAt), drawingAt);
  assert.match(paragraph, /<w:jc w:val="left"\/>/, 'the figure paragraph should be left-aligned');
});

test('captions without a number are numbered per chapter', async () => {
  // A caption with no `number` used to render with no number at all, which made
  // it unreachable from a cross-reference.
  const buffer = await renderDocument({
    template,
    document: {
      type: 'thesis', meta: { title: 'Auto', degree: 'master', date: '2026' }, cover: [],
      sections: [
        { id: 'abs', type: 'abstract', title: '摘要', content: [
          { type: 'table', caption: '前置表', headers: ['A'], data: [['1']] },
        ] },
        { id: 'c1', type: 'chapter', title: '第一章', content: [
          { type: 'table', caption: '第一个表', headers: ['A'], data: [['1']] },
          { type: 'table', caption: '第二个表', headers: ['B'], data: [['2']] },
        ] },
        { id: 'c2', type: 'chapter', title: '第二章', content: [
          { type: 'figure', path: 'missing.png', caption: '一个图' },
        ] },
      ],
    },
  });
  const { xml } = await documentXml(buffer);

  // Front matter is outside a chapter, so its caption is 表1; chapters restart.
  // The number is not one run any more: it is the label run, a STYLEREF field and
  // a SEQ field. Join everything the paragraph prints and drop the title.
  const captionText = (title) => {
    const at = xml.indexOf(title);
    assert.ok(at > 0, `expected the caption ${title}`);
    const open = xml.lastIndexOf('<w:p>', at);
    const close = xml.indexOf('</w:p>', at);
    const printed = [...xml.slice(open, close).matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(m => m[1]).join('');
    return printed.slice(0, printed.indexOf(title)).trim();
  };
  assert.equal(captionText('前置表'), '表1');
  assert.equal(captionText('第一个表'), '表1-1');
  assert.equal(captionText('第二个表'), '表1-2');
  assert.equal(captionText('一个图'), '图2-1');
});

test('a reference to a caption that does not exist warns instead of emitting a broken field', async () => {
  // A REF field pointing at a missing bookmark makes Word print
  // "Error! Reference source not found." — and nothing said so before.
  const warnings = [];
  const original = console.warn;
  console.warn = (message) => warnings.push(String(message));
  try {
    const buffer = await renderDocument({
      template,
      document: {
        type: 'thesis', meta: { title: 'Dangling', degree: 'master', date: '2026' }, cover: [],
        sections: [{ id: 'c1', type: 'chapter', title: 'C', content: [
          { type: 'paragraph', text: '见[[表9-9]]。' },
        ] }],
      },
    });
    const { xml } = await documentXml(buffer);
    assert.equal(refTarget(xml), undefined, 'no REF field should be emitted');
    assert.match(xml, /\[\[表9-9\]\]/, 'the reference is left as text');
  } finally {
    console.warn = original;
  }
  assert.ok(warnings.some(w => w.includes('表9-9')), `expected a warning, got: ${JSON.stringify(warnings)}`);
});

test('a reference can name a block by id, which survives automatic numbering', async () => {
  // The author cannot know the table will become 表3-1, so `[[表#tab2]]` resolves
  // through the block id instead.
  const buffer = await renderDocument({
    template,
    document: {
      type: 'thesis', meta: { title: 'By id', degree: 'master', date: '2026' }, cover: [],
      sections: [{ id: 'c1', type: 'chapter', title: 'C', content: [
        { type: 'paragraph', text: '见[[表#tab2]]。' },
        { type: 'table', id: 'tab2', caption: '第二表', headers: ['A'], data: [['1']] },
      ] }],
    },
  });
  const { xml } = await documentXml(buffer);
  const bookmark = xml.match(/<w:bookmarkStart[^>]*w:name="(_Ref[^"]+)"/)?.[1];
  const target = refTarget(xml);
  assert.ok(bookmark, 'expected the caption bookmark');
  assert.equal(target, bookmark, 'the field must point at that caption');
});

test('a group figure lays its panels out in a borderless table with (a), (b) labels', async () => {
  // 图1(a), 图1(b): the common academic shape for one experiment compared across
  // cases. Word itself lays these out in a table, which keeps the labels aligned
  // under their images.
  const contentDir = mkdtempSync(join(tmpdir(), 'openthesis-subfig-'));
  const png = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAFUlEQVR42mNk+M9Qz0AEYBxVSF+FAP5FDvcfRYWgAAAAAElFTkSuQmCC',
    'base64',
  );
  writeFileSync(join(contentDir, 'a.png'), png);
  writeFileSync(join(contentDir, 'b.png'), png);
  writeFileSync(join(contentDir, 'c.png'), png);

  const buffer = await renderDocument({
    template,
    contentDir,
    document: {
      type: 'thesis', meta: { title: 'Panels', degree: 'master', date: '2026' }, cover: [],
      sections: [{ id: 'c1', type: 'chapter', title: 'C', content: [{
        type: 'figure', caption: '三种工况对比', number: '图1-1',
        items: [
          { path: 'a.png', caption: '工况一' },
          { path: 'b.png', caption: '工况二' },
          { path: 'c.png', caption: '工况三' },
        ],
      }] }],
    },
  });
  const { xml } = await documentXml(buffer);

  // Three panels, three labels, one group caption.
  assert.match(xml, /\(a\)/);
  assert.match(xml, /\(b\)/);
  assert.match(xml, /\(c\)/);
  assert.match(xml, /工况一/);
  assert.match(xml, /工况二/);
  assert.match(xml, /工况三/);
  assert.match(xml, /图1-1/);
  assert.match(xml, /三种工况对比/);

  // The layout table draws nothing itself.
  const layoutTable = xml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/)?.[0] ?? '';
  assert.ok(layoutTable, 'expected a layout table');
  assert.match(layoutTable, /<w:tblBorders>/);
  assert.doesNotMatch(layoutTable, /w:color="000000"/, 'the layout table must draw no visible rule');
});

test('a group figure numbers once and can be referenced by that number', async () => {
  const contentDir = mkdtempSync(join(tmpdir(), 'openthesis-subfig2-'));
  writeFileSync(
    join(contentDir, 'a.png'),
    Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAFUlEQVR42mNk+M9Qz0AEYBxVSF+FAP5FDvcfRYWgAAAAAElFTkSuQmCC', 'base64'),
  );

  const buffer = await renderDocument({
    template,
    contentDir,
    document: {
      type: 'thesis', meta: { title: 'One number', degree: 'master', date: '2026' }, cover: [],
      sections: [{ id: 'c1', type: 'chapter', title: 'C', content: [
        { type: 'paragraph', text: '如图1-1(a)所示。' },
        { type: 'figure', caption: '面板', items: [{ path: 'a.png', label: '(a)' }] },
      ] }],
    },
  });
  const { xml } = await documentXml(buffer);
  // Exactly one number for the group, and the panel label is not a second number.
  const numbers = [...xml.matchAll(/图1-1/g)].length;
  assert.ok(numbers >= 1);
  assert.equal((xml.match(/_Ref图11/g) ?? []).length >= 1, true, 'the group caption should be bookmarked');
});

// ── Captions as Word captions ──────────────────────────────

test('a caption number is a SEQ field, which is what makes it a Word caption', async () => {
  // Word's Cross-reference dialog lists captions by SEQ identifier. A number
  // written as plain text looks identical and is invisible to it.
  const buffer = await renderDocument({
    template,
    document: {
      type: 'thesis', meta: { title: 'SEQ', degree: 'master', date: '2026' }, cover: [],
      sections: [{ id: 'c1', type: 'chapter', title: '第一章', content: [
        { type: 'table', caption: '参数', headers: ['A'], data: [['1']] },
        { type: 'figure', path: 'missing.png', caption: '示意' },
      ] }],
    },
  });
  const { xml } = await documentXml(buffer);
  const instrs = fieldInstrs(xml);

  assert.ok(
    instrs.some(i => i.startsWith('SEQ Table') && i.includes('\\* ARABIC') && i.includes('\\s 1')),
    `a table caption must carry SEQ Table with \\s 1, got ${JSON.stringify(instrs)}`,
  );
  assert.ok(
    instrs.some(i => i.startsWith('SEQ Figure') && i.includes('\\s 1')),
    `a figure caption must carry SEQ Figure with \\s 1, got ${JSON.stringify(instrs)}`,
  );
  // No STYLEREF: the chapter titles carry their number in their text, so a
  // STYLEREF there renders as "Error: Reference source not found".
  assert.ok(!instrs.some(i => i.startsWith('STYLEREF')), 'no STYLEREF should be emitted');
});

test('a caption\'s SEQ field sits inside its bookmark, so a REF returns the number', async () => {
  const buffer = await renderDocument({
    template,
    document: {
      type: 'thesis', meta: { title: 'Bounds', degree: 'master', date: '2026' }, cover: [],
      sections: [{ id: 'c1', type: 'chapter', title: 'C', content: [
        { type: 'table', caption: '参数', headers: ['A'], data: [['1']] },
      ] }],
    },
  });
  const { xml } = await documentXml(buffer);
  // The caption's own bookmark, not the first one in the document: headings
  // carry `_Toc…` bookmarks of their own now.
  const nameAt = xml.indexOf('_Ref');
  const start = xml.lastIndexOf('<w:bookmarkStart', nameAt);
  const end = xml.indexOf('<w:bookmarkEnd', nameAt);
  const between = xml.slice(start, end);
  assert.match(between, /SEQ Table/, 'the SEQ field must be inside the bookmark');
});

test('chapter headings carry an outline level, without which \\s 1 cannot restart', async () => {
  const buffer = await renderDocument({
    template,
    document: {
      type: 'thesis', meta: { title: 'Outline', degree: 'master', date: '2026' }, cover: [],
      sections: [{ id: 'c1', type: 'chapter', title: '第一章', content: [{ type: 'paragraph', text: 'x' }] }],
    },
  });
  const { xml } = await documentXml(buffer);
  assert.match(xml, /<w:outlineLvl w:val="0"\/>/, 'a chapter is outline level 1, which is w:val="0"');
});

test('a hand-written caption number stays literal text and is not renumbered', async () => {
  // The author asked for that exact label; a SEQ field would renumber it out
  // from under them.
  const buffer = await renderDocument({
    template,
    document: {
      type: 'thesis', meta: { title: 'Manual', degree: 'master', date: '2026' }, cover: [],
      sections: [{ id: 'c1', type: 'chapter', title: 'C', content: [
        { type: 'table', caption: '参数', number: '表3-1', headers: ['A'], data: [['1']] },
      ] }],
    },
  });
  const { xml } = await documentXml(buffer);
  assert.ok(!fieldInstrs(xml).some(i => i.startsWith('SEQ')), 'no SEQ field for a manual number');
  assert.match(xml, /<w:t[^>]*>表3-1<\/w:t>/, 'the manual label is written as text');
});

// ── Comparing the output against the 书写指南 ───────────────

test('references are numbered in print order, not by their internal id', async () => {
  // The id is a handle for cross-referencing in the content file. Printing it as
  // the number put `[r1]` where `[1]` belongs.
  const buffer = await renderDocument({
    template,
    document: {
      type: 'thesis', meta: { title: 'Refs', degree: 'master', date: '2026' }, cover: [],
      sections: [{ id: 'c1', type: 'chapter', title: 'C', content: [{ type: 'paragraph', text: 'x' }] }],
      backMatter: { references: [
        { id: 'alpha', text: '第一篇文献.' },
        { id: 'beta', text: '第二篇文献.' },
      ] },
    },
  });
  const { xml } = await documentXml(buffer);
  assert.match(xml, /\[1\] 第一篇文献/);
  assert.match(xml, /\[2\] 第二篇文献/);
  assert.doesNotMatch(xml, /\[alpha\]/);
});

test('each front-matter part starts on its own page', async () => {
  // 摘要, Abstract, 目录 and 图表清单 are separate pages in the guide; they used to
  // run into one another because they share a section.
  const buffer = await renderDocument({
    template,
    document: {
      type: 'thesis', meta: { title: 'Pages', degree: 'master', date: '2026' }, cover: [],
      sections: [
        { id: 'abs', type: 'abstract', title: '摘要', content: [{ type: 'paragraph', text: 'a' }] },
        { id: 'absEn', type: 'abstract', title: 'Abstract', content: [{ type: 'paragraph', text: 'b' }] },
        { id: 'toc', type: 'toc', title: '目录', content: [] },
      ],
    },
  });
  const { xml } = await documentXml(buffer);
  // One break before each part after the first.
  assert.equal((xml.match(/<w:pageBreakBefore\/>/g) ?? []).length >= 2, true, 'expected a break before each later part');
});

test('a 图表清单 lists the captions the document actually has', async () => {
  const contentDir = mkdtempSync(join(tmpdir(), 'openthesis-lof-'));
  writeFileSync(
    join(contentDir, 'a.png'),
    Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAFUlEQVR42mNk+M9Qz0AEYBxVSF+FAP5FDvcfRYWgAAAAAElFTkSuQmCC', 'base64'),
  );
  const buffer = await renderDocument({
    template, contentDir,
    document: {
      type: 'thesis', meta: { title: 'LOF', degree: 'master', date: '2026' }, cover: [],
      sections: [
        { id: 'lof', type: 'list_of_figures', title: '插图清单', content: [] },
        { id: 'c1', type: 'chapter', title: 'C', content: [
          { type: 'figure', path: 'a.png', caption: '系统结构' },
          { type: 'table', caption: '参数一览', headers: ['A'], data: [['1']] },
        ] },
      ],
    },
  });
  const { xml } = await documentXml(buffer);
  // The list prints the caption and a page reference; it must not print the
  // table, which belongs to 表格清单.
  assert.match(xml, /图1-1 {2}系统结构/);
  assert.doesNotMatch(xml, /表1-1 {2}参数一览/);
  assert.ok(fieldInstrs(xml).some(i => i.startsWith('PAGEREF')), 'each line needs a page reference');
});

test('a 图表清单 with nothing to list says so instead of printing an empty page', async () => {
  const buffer = await renderDocument({
    template,
    document: {
      type: 'thesis', meta: { title: 'Empty', degree: 'master', date: '2026' }, cover: [],
      sections: [{ id: 'lof', type: 'list_of_figures', title: '插图清单', content: [] }],
    },
  });
  const { xml } = await documentXml(buffer);
  assert.match(xml, /本文无此类内容/);
});

test('the back matter parts each start on their own page', async () => {
  const buffer = await renderDocument({
    template,
    document: {
      type: 'thesis', meta: { title: 'Back', degree: 'master', date: '2026' }, cover: [],
      sections: [{ id: 'c1', type: 'chapter', title: 'C', content: [{ type: 'paragraph', text: 'x' }] }],
      backMatter: {
        references: [{ id: 'r1', text: '文献.' }],
        authorBiography: '简介.',
        declaration: '声明.',
      },
    },
  });
  const { xml } = await documentXml(buffer);
  // 参考文献, 作者简历 and 独创性说明 each carry a break before their heading.
  assert.equal((xml.match(/<w:pageBreakBefore\/>/g) ?? []).length >= 3, true);
});
