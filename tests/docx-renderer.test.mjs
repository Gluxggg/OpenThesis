import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createUSTBTemplate, renderDocument } from '../packages/docx-renderer/dist/index.js';

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
  assert.match(xml, /1 Figure caption/);
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

test('emits a real TOC field for a 目录 section', async () => {
  // A table of contents is a field, not text: Word fills in the entries and
  // their page numbers when the field is updated. No generator can compute page
  // numbers, so it emits the field rather than faking a list.
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
  // The field's quotes are XML-escaped in the part.
  assert.match(xml, /<w:instrText[^>]*>TOC \\h \\o &quot;1-3&quot;<\/w:instrText>/);
  assert.match(xml, /目录/);
  assert.match(xml, /第一章 绪论/);
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
  assert.match(xml, /<w:left w:val="none" w:color="FFFFFF" w:sz="0"\/>/);
  assert.match(xml, /<w:bottom w:val="single" w:color="000000" w:sz="12"\/>/);
  assert.match(xml, /<w:right w:val="none" w:color="FFFFFF" w:sz="0"\/>/);
  assert.match(xml, /<w:insideV w:val="none" w:color="FFFFFF" w:sz="0"\/>/);

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
