import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { parsePdfTemplate, parseTemplate } from '../packages/template-parser/dist/index.js';
import { buildLinedPdf, buildTestPdf } from './helpers/pdf.mjs';

async function createTemplate() {
  const zip = new JSZip();
  zip.file('word/styles.xml', `<?xml version="1.0"?>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:eastAsia="宋体"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>
      <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:jc w:val="both"/></w:pPr></w:style>
      <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
    </w:styles>`);
  zip.file('word/document.xml', `<?xml version="1.0"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1800" w:right="1800"/><w:cols w:num="2" w:space="720"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

test('parses page settings and resolves inherited styles', async () => {
  const template = await parseTemplate(await createTemplate(), {
    organization: 'Test University', documentType: 'thesis',
  });
  assert.equal(template.meta.organization, 'Test University');
  assert.equal(template.page.columns, 2);
  assert.equal(template.page.columnGutter, 720);
  assert.equal(template.styles.Heading1.font.name, 'Arial');
  assert.equal(template.styles.Heading1.font.size, 32);
  assert.equal(template.styles.Heading1.font.bold, true);
  assert.equal(template.styleRoles.Heading1, 'heading1');
  assert.equal(template.styleInheritance.Heading1, 'Normal');
});

test('rejects a ZIP that is not a Word template', async () => {
  const zip = new JSZip();
  zip.file('readme.txt', 'not a docx');
  const buffer = await zip.generateAsync({ type: 'nodebuffer' });
  await assert.rejects(() => parseTemplate(buffer), /No styles\.xml found/);
});

test('keeps metadata defaults when optional overrides are undefined', async () => {
  const template = await parseTemplate(await createTemplate(), {
    organization: undefined,
    name: undefined,
  });
  assert.equal(template.meta.organization, 'Unknown Organization');
  assert.equal(template.meta.name, 'Untitled Template');
  assert.equal(template.meta.parserVersion, '0.2.0');
});

/**
 * Build a template whose heading carries its font in `w:pPr/w:rPr` (the shape
 * Chinese Word/WPS templates emit) and whose indents use the `*Chars` units.
 */
async function createChineseTemplate() {
  const zip = new JSZip();
  zip.file('word/styles.xml', `<?xml version="1.0"?>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:docDefaults><w:pPrDefault/></w:docDefaults>
      <w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:jc w:val="both"/><w:ind w:firstLineChars="200"/></w:pPr></w:style>
      <w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/><w:basedOn w:val="a"/>
        <w:pPr><w:outlineLvl w:val="0"/><w:jc w:val="center"/><w:rPr><w:rFonts w:eastAsia="黑体"/><w:b/><w:sz w:val="32"/></w:rPr></w:pPr>
      </w:style>
      <w:style w:type="character" w:styleId="uChar"><w:name w:val="char"/><w:rPr><w:color w:val="FF0000"/></w:rPr></w:style>
      <w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style>
      <w:style w:type="paragraph" w:styleId="MTConvertedEquation"><w:name w:val="MT Converted Equation"/><w:basedOn w:val="a"/><w:rPr><w:i/><w:sz w:val="28"/></w:rPr></w:style>
    </w:styles>`);
  zip.file('word/document.xml', `<?xml version="1.0"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1797" w:right="1797"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

test('reads run properties nested under w:pPr/w:rPr', async () => {
  const template = await parseTemplate(await createChineseTemplate(), { documentType: 'thesis' });
  // Without this, every style looked empty and headings rendered as body text.
  assert.equal(template.styles['1'].font.eastAsia, '黑体');
  assert.equal(template.styles['1'].font.size, 32);
  assert.equal(template.styles['1'].font.bold, true);
  assert.equal(template.styles['1'].paragraph.alignment, 'center');
});

test('converts w:firstLineChars into twips instead of storing the raw count', async () => {
  const template = await parseTemplate(await createChineseTemplate(), { documentType: 'thesis' });
  // 200 hundredths of a character at 12pt = 2 chars × 240 twips = 480 twips.
  assert.equal(template.styles.a.paragraph.firstLineIndent, 480);
});

test('only paragraph styles receive block roles', async () => {
  const template = await parseTemplate(await createChineseTemplate(), { documentType: 'thesis' });
  // Character and table styles are not paragraph-level blocks; assigning them a
  // role let `resolveStyle` pick e.g. a character style for body text.
  assert.equal(template.styleRoles.uChar, undefined);
  assert.equal(template.styleRoles.TableGrid, undefined);
  // `Table Grid` used to become heading3 via an unanchored /^(表|Table|题注)/.
  assert.notEqual(template.styleRoles.TableGrid, 'heading3');
});

test('detects a template equation style', async () => {
  const template = await parseTemplate(await createChineseTemplate(), { documentType: 'thesis' });
  assert.equal(template.styleRoles.MTConvertedEquation, 'equation');
});

test('omits properties a style does not declare', async () => {
  const template = await parseTemplate(await createChineseTemplate(), { documentType: 'thesis' });
  // `Normal` declares no run properties, so none may be invented for it —
  // otherwise "the template is silent" is indistinguishable from an explicit
  // choice, and the renderer's role defaults can never take effect.
  assert.equal(template.styles.a.font.size, undefined);
  assert.equal(template.styles.a.font.eastAsia, undefined);
  assert.equal(template.styles.a.font.name, undefined);
  // Properties it does declare are kept.
  assert.equal(template.styles.a.paragraph.alignment, 'justified');
  // Nothing is invented for absent spacing either.
  assert.equal(template.styles['1'].paragraph.spaceBefore, undefined);
});

/** Build a template whose sections and style quality we control. */
async function createTemplateWith({ styles, body }) {
  const zip = new JSZip();
  zip.file('word/styles.xml', `<?xml version="1.0"?>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${styles}</w:styles>`);
  zip.file('word/document.xml', `<?xml version="1.0"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

const BODY_SECTION = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1797" w:right="1797"/></w:sectPr>';

test('reads section properties from paragraph-level section breaks', async () => {
  // A cover section (landscape) followed by the body section. Only the
  // body-level sectPr used to be read, so the cover geometry vanished.
  const buffer = await createTemplateWith({
    styles: '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:rPr><w:sz w:val="24"/></w:rPr></w:style>',
    body: `<w:p><w:pPr><w:sectPr><w:pgSz w:w="16838" w:h="11906"/><w:pgMar w:top="1000" w:bottom="1000" w:left="1000" w:right="1000"/></w:sectPr></w:pPr></w:p>${BODY_SECTION}`,
  });
  const template = await parseTemplate(buffer, { documentType: 'thesis' });

  assert.equal(template.pageSections.length, 2);
  // Document order: cover first, body last, and `page` is the body.
  assert.equal(template.pageSections[0].width, 16838, 'cover section is landscape');
  assert.equal(template.page.width, 11906);
  assert.equal(template.page.margins.left, 1797);
  assert.ok(template.warnings.some(w => /declares 2 sections/.test(w)));
});

test('warns when a template carries no formatting at all', async () => {
  // This is exactly the shape that made headings render as body text.
  const buffer = await createTemplateWith({
    styles: '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style>',
    body: BODY_SECTION,
  });
  const template = await parseTemplate(buffer, { documentType: 'thesis' });
  assert.ok(template.warnings.some(w => /carries no typography/.test(w)), JSON.stringify(template.warnings));
});

test('keeps the line spacing rule apart from the distance', async () => {
  // `w:lineRule` decides whether `w:line` counts 240ths of a line (auto) or
  // twips (exact). Dropping it reinterpreted every exact spacing as a multiple.
  const exact = await parseTemplate(await createTemplateWith({
    styles: '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
      + '<w:pPr><w:spacing w:line="360" w:lineRule="exact"/></w:pPr></w:style>',
    body: BODY_SECTION,
  }), { documentType: 'thesis' });
  assert.equal(exact.styles.a.lineSpacing, 360);
  assert.equal(exact.styles.a.lineSpacingRule, 'exact');

  // A style that declares no spacing declares no rule either.
  const bare = await parseTemplate(await createTemplateWith({
    styles: '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:rPr><w:sz w:val="24"/></w:rPr></w:style>',
    body: BODY_SECTION,
  }), { documentType: 'thesis' });
  assert.equal(bare.styles.a.lineSpacingRule, undefined);
});

test('breaks a basedOn loop at a fixed point and reports it', async () => {
  // a → b → a. The merge used to depend on which style was resolved first, and
  // `a` inherited `b`'s bold back through the loop.
  const buffer = await createTemplateWith({
    styles: '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
      + '<w:basedOn w:val="b"/><w:rPr><w:sz w:val="24"/></w:rPr></w:style>'
      + '<w:style w:type="paragraph" w:styleId="b"><w:name w:val="B"/><w:basedOn w:val="a"/>'
      + '<w:rPr><w:b/></w:rPr></w:style>',
    body: BODY_SECTION,
  });
  const template = await parseTemplate(buffer, { documentType: 'thesis' });

  assert.ok(template.warnings.some(w => /loop/.test(w)), JSON.stringify(template.warnings));
  // The loop is cut at `a`, the earliest-declared member, so `a` keeps only what
  // it declares while `b` still inherits from it.
  assert.equal(template.styles.a.font.bold, undefined);
  assert.equal(template.styles.a.font.size, 24);
  assert.equal(template.styles.b.font.bold, true);
  assert.equal(template.styles.b.font.size, 24);
});

test('warns when a style inherits from an undefined style', async () => {
  const buffer = await createTemplateWith({
    styles: '<w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/><w:basedOn w:val="Missing"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>',
    body: BODY_SECTION,
  });
  const template = await parseTemplate(buffer, { documentType: 'thesis' });
  assert.ok(template.warnings.some(w => /does not define/.test(w)), JSON.stringify(template.warnings));
});

test('stays quiet about a template that does declare formatting', async () => {
  const template = await parseTemplate(await createChineseTemplate(), { documentType: 'thesis' });
  assert.equal(template.warnings, undefined);
  assert.equal(template.pageSections.length, 1);
  assert.equal(template.page.columns, 1);
});

test('reads a template that binds WordprocessingML to another prefix', async () => {
  // The prefix is declared per document; nothing requires it to be `w`. Keying
  // on the literal `w:` parsed such a template to zero styles.
  const zip = new JSZip();
  zip.file('word/styles.xml', `<?xml version="1.0"?>
    <x:styles xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <x:style x:type="paragraph" x:default="1" x:styleId="a"><x:name x:val="Normal"/>
        <x:pPr><x:spacing x:line="360" x:lineRule="exact"/></x:pPr>
        <x:rPr><x:sz x:val="24"/><x:rFonts x:eastAsia="宋体"/></x:rPr></x:style>
    </x:styles>`);
  zip.file('word/document.xml', `<?xml version="1.0"?>
    <x:document xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <x:body>${BODY_SECTION.replace(/w:/g, 'x:')}</x:body></x:document>`);

  const template = await parseTemplate(await zip.generateAsync({ type: 'nodebuffer' }), { documentType: 'thesis' });

  assert.equal(template.styles.a.font.size, 24);
  assert.equal(template.styles.a.font.eastAsia, '宋体');
  assert.equal(template.styles.a.lineSpacing, 360);
  assert.equal(template.styles.a.lineSpacingRule, 'exact');
  assert.equal(template.page.width, 11906);
});

test('a style the document uses beats an unused stock style of the same role', async () => {
  // Word's own `heading 2` carries the role by its name, but the template never
  // applies it: every heading is set in the author's `u2级标题`. A plain object
  // enumerates integer-like keys first, so "first entry carrying this role" won
  // the stock style and rendered level-2 headings at its size instead.
  const buffer = await createTemplateWith({
    styles: '<w:style w:type="paragraph" w:styleId="2"><w:name w:val="heading 2"/>'
      + '<w:rPr><w:sz w:val="32"/><w:rFonts w:eastAsia="宋体"/></w:rPr></w:style>'
      + '<w:style w:type="paragraph" w:customStyle="1" w:styleId="u2级标题"><w:name w:val="u2级标题"/>'
      + '<w:rPr><w:sz w:val="28"/><w:rFonts w:eastAsia="黑体"/></w:rPr></w:style>',
    body: '<w:p><w:pPr><w:pStyle w:val="u2级标题"/></w:pPr><w:r><w:t>标题</w:t></w:r></w:p>' + BODY_SECTION,
  });
  const template = await parseTemplate(buffer, { documentType: 'thesis' });

  // Both styles keep the role; the winner is recorded separately.
  assert.equal(template.styleRoles['2'], 'heading2');
  assert.equal(template.styleRoles['u2级标题'], 'heading2');
  assert.equal(template.roleWinners.heading2, 'u2级标题');
  assert.equal(template.styles[template.roleWinners.heading2].font.size, 28);
});

test('w:firstLineChars wins over the twips value Word stores beside it', async () => {
  // Word writes both, and the twips attribute is a stale copy: two characters
  // at 12pt is 480 twips, not the 200 left behind when the style was made.
  const buffer = await createTemplateWith({
    styles: '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
      + '<w:pPr><w:ind w:firstLineChars="200" w:firstLine="200"/></w:pPr>'
      + '<w:rPr><w:sz w:val="24"/></w:rPr></w:style>',
    body: BODY_SECTION,
  });
  const template = await parseTemplate(buffer, { documentType: 'thesis' });
  assert.equal(template.styles.a.paragraph.firstLineIndent, 480);
});

/**
 * A two-section template shaped like a real thesis: an empty 篇眉 part on the
 * cover, the university's line and a PAGE footer on the body, `upperRoman` on
 * the first section, and the document-level flags in settings.xml.
 */
async function createSectionedTemplate() {
  const zip = new JSZip();
  zip.file('word/styles.xml', `<?xml version="1.0"?>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:rPr><w:sz w:val="24"/></w:rPr></w:style>
    </w:styles>`);
  zip.file('word/settings.xml', `<?xml version="1.0"?>
    <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:mirrorMargins/><w:evenAndOddHeaders/>
    </w:settings>`);
  zip.file('word/_rels/document.xml.rels', `<?xml version="1.0"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>
      <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
    </Relationships>`);
  // The cover's even head is an *empty* paragraph: that is how the university
  // template suppresses it, not `w:titlePg`.
  zip.file('word/header1.xml', `<?xml version="1.0"?>
    <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p/></w:hdr>`);
  zip.file('word/header2.xml', `<?xml version="1.0"?>
    <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr>
      <w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr>
      <w:jc w:val="center"/><w:rPr><w:sz w:val="21"/></w:rPr></w:pPr>
      <w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>北京科技大学硕士学位论文</w:t></w:r></w:p></w:hdr>`);
  // A PAGE field plus the cached number Word last laid out — "7" is not content.
  zip.file('word/footer1.xml', `<?xml version="1.0"?>
    <w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText>PAGE   \\* MERGEFORMAT</w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>7</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>`);
  zip.file('word/document.xml', `<?xml version="1.0"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>
      <w:p><w:r><w:t>封面</w:t></w:r></w:p>
      <w:p><w:pPr><w:sectPr>
        <w:headerReference w:type="even" r:id="rId1"/>
        <w:pgSz w:w="11906" w:h="16838"/>
        <w:pgMar w:top="1701" w:right="1701" w:bottom="1134" w:left="1701" w:header="851" w:footer="851" w:gutter="567"/>
        <w:pgNumType w:fmt="upperRoman"/>
      </w:sectPr></w:pPr></w:p>
      <w:p><w:r><w:t>摘要</w:t></w:r></w:p>
      <w:sectPr>
        <w:headerReference w:type="default" r:id="rId2"/>
        <w:footerReference w:type="default" r:id="rId3"/>
        <w:pgSz w:w="11906" w:h="16838"/>
        <w:pgMar w:top="1701" w:right="1701" w:bottom="1134" w:left="1701" w:header="851" w:footer="850" w:gutter="567"/>
        <w:pgNumType w:start="1"/>
      </w:sectPr>
    </w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

test('reads each section\'s running heads and page numbering from the parts', async () => {
  const template = await parseTemplate(await createSectionedTemplate(), { documentType: 'thesis' });
  const [cover, body] = template.pageSections;

  // 页码 comes from `w:pgNumType`, never from the footer that prints it.
  assert.equal(cover.pageNumberFormat, 'upperRoman');
  assert.equal(cover.pageNumberStart, undefined);
  assert.equal(body.pageNumberStart, 1);
  assert.equal(body.pageNumberFormat, undefined, 'the format is inherited, then decimal on a restart');

  // An empty part is a suppression, not a missing head.
  assert.equal(cover.headers.even.text, '');
  assert.equal(cover.headers.even.pageNumber, false);

  // The head itself, with the 篇眉 rule that draws its 0.5 pt line.
  assert.equal(body.headers.default.text, '北京科技大学硕士学位论文');
  assert.equal(body.headers.default.rule, true);
  assert.equal(body.headers.default.pageNumber, false);

  // OOXML inherits each slot independently: the body declares no even head, yet
  // it prints the covers's — which is what makes the body look head-less when
  // only the sections that declare references are read.
  assert.equal(body.headers.even.text, '');

  // A PAGE field is the page number; the "7" beside it is a cached layout.
  assert.equal(body.footers.default.pageNumber, true);
  assert.equal(body.footers.default.text, '7');

  // Both document settings live in settings.xml, and only there.
  assert.equal(template.evenAndOddHeaders, true);
  assert.equal(template.page.mirrorMargins, true);
  assert.equal(cover.mirrorMargins, true);
  assert.equal(template.page.headerDistance, 851);
  assert.equal(template.page.footerDistance, 850);
});

test('leaves evenAndOddHeaders unset when settings.xml does not ask for it', async () => {
  const zip = await JSZip.loadAsync(await createSectionedTemplate());
  zip.file('word/settings.xml', `<?xml version="1.0"?>
    <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:mirrorMargins/></w:settings>`);
  const template = await parseTemplate(await zip.generateAsync({ type: 'nodebuffer' }), { documentType: 'thesis' });

  assert.equal(template.evenAndOddHeaders, undefined);
  // `w:val="false"` is how the writer library says "off"; it must not read as on.
  const off = await JSZip.loadAsync(await createSectionedTemplate());
  off.file('word/settings.xml', `<?xml version="1.0"?>
    <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:evenAndOddHeaders w:val="false"/></w:settings>`);
  const parsedOff = await parseTemplate(await off.generateAsync({ type: 'nodebuffer' }), { documentType: 'thesis' });
  assert.equal(parsedOff.evenAndOddHeaders, undefined);
});

test('counts style usage from the parsed document when it binds another prefix', async () => {
  // Two neutral-named styles, both outline level 0: neither a name pattern nor
  // the outline level decides between them, only how often the document applies
  // them. The document binds WordprocessingML to `wp`, and the usage counter used
  // to scan the raw XML for the literal `<w:pStyle`, find nothing, count zero for
  // every style, and let declaration order hand heading1 to the unused `a1`.
  const styles = '<wp:style wp:type="paragraph" wp:styleId="a1"><wp:name wp:val="N1"/>'
    + '<wp:pPr><wp:outlineLvl wp:val="0"/></wp:pPr><wp:rPr><wp:sz wp:val="32"/></wp:rPr></wp:style>'
    + '<wp:style wp:type="paragraph" wp:styleId="a2"><wp:name wp:val="N2"/>'
    + '<wp:pPr><wp:outlineLvl wp:val="0"/></wp:pPr><wp:rPr><wp:sz wp:val="28"/></wp:rPr></wp:style>';
  const zip = new JSZip();
  zip.file('word/styles.xml', `<?xml version="1.0"?>
    <wp:styles xmlns:wp="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${styles}</wp:styles>`);
  zip.file('word/document.xml', `<?xml version="1.0"?>
    <wp:document xmlns:wp="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><wp:body>
      <wp:p><wp:pPr><wp:pStyle wp:val="a2"/></wp:pPr><wp:r><wp:t>标题</wp:t></wp:r></wp:p>
      <wp:sectPr><wp:pgSz wp:w="11906" wp:h="16838"/><wp:pgMar wp:top="1440" wp:bottom="1440" wp:left="1797" wp:right="1797"/></wp:sectPr>
    </wp:body></wp:document>`);

  const template = await parseTemplate(
    await zip.generateAsync({ type: 'nodebuffer' }),
    { documentType: 'thesis' },
  );

  // a2 is the style the document actually applies; a1 is declared but unused.
  assert.equal(template.roleWinners.heading1, 'a2');
  assert.equal(template.styles[template.roleWinners.heading1].font.size, 28);
});

test('converts w:beforeLines / w:afterLines into twips', async () => {
  // Chinese Word/WPS templates write 段前0.5行 as `w:beforeLines="50"`, and reading
  // only `w:before` / `w:after` parsed such a style as declaring no spacing at
  // all. One line is one font size: 50 hundredths at 12 pt is 120 twips.
  const lines = await parseTemplate(await createTemplateWith({
    styles: '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
      + '<w:pPr><w:spacing w:beforeLines="50" w:afterLines="30"/></w:pPr>'
      + '<w:rPr><w:sz w:val="24"/></w:rPr></w:style>',
    body: BODY_SECTION,
  }), { documentType: 'thesis' });
  assert.equal(lines.styles.a.paragraph.spaceBefore, 120);
  assert.equal(lines.styles.a.paragraph.spaceAfter, 72);

  // A template that writes plain twips keeps working the way it always did.
  const twips = await parseTemplate(await createTemplateWith({
    styles: '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
      + '<w:pPr><w:spacing w:before="300" w:after="150"/></w:pPr></w:style>',
    body: BODY_SECTION,
  }), { documentType: 'thesis' });
  assert.equal(twips.styles.a.paragraph.spaceBefore, 300);
  assert.equal(twips.styles.a.paragraph.spaceAfter, 150);
});

test('a line-based spacing value wins over the twips beside it and follows the font size', async () => {
  // Word writes both, and the twips copy is stale in the same way as
  // `w:firstLine`: the line-based value is the one that holds. A 16 pt style's
  // 0.5 line is 160 twips, not the 300 left behind at 10 pt.
  const template = await parseTemplate(await createTemplateWith({
    styles: '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
      + '<w:pPr><w:spacing w:beforeLines="50" w:before="300" w:afterLines="100" w:after="999"/></w:pPr>'
      + '<w:rPr><w:sz w:val="32"/></w:rPr></w:style>',
    body: BODY_SECTION,
  }), { documentType: 'thesis' });
  assert.equal(template.styles.a.paragraph.spaceBefore, 160);
  assert.equal(template.styles.a.paragraph.spaceAfter, 320);
});

test('warns about a template whose only formatting declaration is an empty w:docDefaults', async () => {
  // Word writes `<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault></w:docDefaults>`
  // into every file it saves. The inheritance merge spread two absent
  // run-property bags into `rPr: {}`, so a template that declares nothing counted
  // as declaring typography and this warning could never fire.
  const buffer = await createTemplateWith({
    styles: '<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault></w:docDefaults>'
      + '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/></w:style>',
    body: BODY_SECTION,
  });
  const template = await parseTemplate(buffer, { documentType: 'thesis' });
  assert.ok(template.warnings.some(w => /carries no typography/.test(w)), JSON.stringify(template.warnings));

  // The same empty defaults must not mask a style that does declare formatting.
  const formatted = await parseTemplate(await createTemplateWith({
    styles: '<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault></w:docDefaults>'
      + '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/>'
      + '<w:rPr><w:sz w:val="24"/></w:rPr></w:style>',
    body: BODY_SECTION,
  }), { documentType: 'thesis' });
  assert.equal(formatted.warnings, undefined);
});

test('the shipped USTB template declares formatting and carries no typography warning', async () => {
  // Acceptance for the empty-rPr fix has two halves: a silent template must warn,
  // and a real one must stay quiet. Stating the second half against the shipped
  // asset is what used to require regenerating it to notice.
  const { readFile } = await import('node:fs/promises');
  const asset = JSON.parse(await readFile(new URL('../assets/ustb-thesis-template.json', import.meta.url), 'utf8'));

  assert.ok(
    !(asset.warnings ?? []).some(w => /carries no typography/.test(w)),
    JSON.stringify(asset.warnings),
  );
  // The premise the warning checks, stated against the asset itself: every style
  // declares something.
  const silent = Object.entries(asset.styles).filter(([, style]) => (
    Object.keys(style.font ?? {}).length === 0
    && Object.keys(style.paragraph ?? {}).length === 0
    && style.lineSpacing === undefined
    && style.lineSpacingRule === undefined
  ));
  assert.deepEqual(silent, []);
});

// ── PDF templates (feat: infer templates from text-based PDFs) ──

test('infers a reusable template from a text-based PDF', async () => {
  const template = await parsePdfTemplate(buildTestPdf(), {
    organization: 'PDF University', documentType: 'thesis', sourceFile: 'guide.pdf',
  });
  assert.equal(template.meta.sourceFormat, 'pdf');
  assert.equal(template.meta.pageCount, 1);
  assert.equal(template.page.width, 12240);
  assert.equal(template.page.height, 15840);
  assert.ok(Object.values(template.styleRoles).includes('paragraph'));
  assert.ok(Object.values(template.styleRoles).includes('centered_text'));
  assert.ok(template.meta.warnings.some(warning => /inferred/.test(warning)));
});

test('asks for OCR when a PDF has no extractable text', async () => {
  await assert.rejects(() => parsePdfTemplate(buildTestPdf({ text: false })), /run OCR first/);
});

test('maps the CJK face a PDF embeds onto the family Word names', async () => {
  // PDF.js reports `sans-serif` as the CSS family of an embedded font and keeps
  // `SimSun` on the font object. Reading the CSS family alone turned every
  // Chinese face in the inferred template into Arial — 宋体 body text became
  // "Arial", and the rendered document fell back to whatever Word substitutes.
  // The fixture is a subsetted 宋体/黑体 PDF (37 KB), so the names stay real.
  const template = await parsePdfTemplate(
    readFileSync(new URL('./fixtures/pdfs/cjk-template.pdf', import.meta.url)),
    { documentType: 'thesis', organization: 'Test University' },
  );

  const styles = Object.values(template.styles);
  const body = styles.find(style => style.font?.size === 24);
  const heading = styles.find(style => style.font?.size === 30);
  assert.equal(body?.font?.eastAsia, '宋体');
  assert.equal(body?.font?.name, 'Times New Roman');
  assert.equal(heading?.font?.eastAsia, '黑体');
  assert.equal(template.styleRoles['pdf_style_01'], 'heading1');
  assert.equal(template.styleRoles['pdf_style_02'], 'paragraph');
});

test('reads a bold face from the PDF font objects, not the CSS fallback', async () => {
  // The fallback family is the same `sans-serif` for every face, so bold used
  // to be undetectable on a PDF-derived template. The base-14 font object is
  // named `Helvetica-Bold`, which is where the weight is read from.
  const template = await parsePdfTemplate(buildLinedPdf({
    pages: [[
      { text: 'A Bold Heading', x: 72, y: 700, size: 15, font: 'bold' },
      { text: 'The body text carries the bulk of the sampled characters.', x: 72, y: 660, size: 12 },
      { text: 'so the dominant cluster is the paragraph style.', x: 72, y: 640, size: 12 },
      { text: 'A third line makes the body cluster large enough to win.', x: 72, y: 620, size: 12 },
      { text: 'A fourth line keeps the inferred margins stable.', x: 72, y: 600, size: 12 },
    ]],
  }), { documentType: 'thesis' });

  const styles = Object.values(template.styles);
  const heading = styles.find(style => style.font?.size === 30);
  const body = styles.find(style => style.font?.size === 24);
  assert.equal(heading?.font?.bold, true);
  assert.equal(heading?.font?.name, 'Arial');
  assert.equal(body?.font?.bold, undefined);
});

test('measures the left margin from the text column, not from an indented first line', async () => {
  // The old 20th-percentile left edge landed on the indent whenever the sampled
  // paragraphs were short: a 2-character indent became part of the left margin.
  // The leftmost line is the column, and the indent is the distance from it.
  const template = await parsePdfTemplate(buildLinedPdf({
    pages: [[
      { text: 'A paragraph opens with a two-character first-line indent.', x: 96, y: 700, size: 12 },
      { text: 'Its continuation line returns to the text column.', x: 72, y: 680, size: 12 },
      { text: 'Another paragraph opens with the same first-line indent.', x: 96, y: 660, size: 12 },
      { text: 'A third paragraph opens the same way.', x: 96, y: 640, size: 12 },
      { text: 'A fourth paragraph opens the same way.', x: 96, y: 620, size: 12 },
      { text: 'A fifth paragraph opens the same way.', x: 96, y: 600, size: 12 },
    ]],
  }), { documentType: 'thesis' });

  assert.equal(template.page.margins.left, 1440);
  const body = Object.values(template.styles).find(style => style.font?.size === 24);
  assert.equal(body?.paragraph?.firstLineIndent, 480);
});

test('a running head does not set the top margin', async () => {
  // The header sits in the top band and used to be the topmost line of the
  // page, so the inferred top margin measured the header: 1.5 cm for a page
  // whose text column starts at 2.5 cm. Only the text column counts now.
  const template = await parsePdfTemplate(buildLinedPdf({
    pages: [[
      { text: 'University of Somewhere Thesis', x: 200, y: 780, size: 9 },
      { text: 'The body starts at the top margin of the page.', x: 72, y: 700, size: 12 },
      { text: 'A second line continues the body text.', x: 72, y: 680, size: 12 },
      { text: 'A third line carries more characters than the header.', x: 72, y: 660, size: 12 },
      { text: 'A fourth line keeps the body cluster dominant.', x: 72, y: 640, size: 12 },
    ]],
  }), { documentType: 'thesis' });

  assert.equal(template.page.margins.top, 1600);
});

test('a table-of-contents line is not read as a heading', async () => {
  // A TOC entry is "1 Introduction ........ 3": it opens like a numbered
  // heading but is set in the body size. The old rule matched the number and
  // promoted a 10.5 pt TOC entry to heading1.
  const template = await parsePdfTemplate(buildLinedPdf({
    pages: [[
      { text: '1 Introduction ........ 3', x: 72, y: 700, size: 10.5 },
      { text: '2 Method', x: 72, y: 660, size: 15, font: 'bold' },
      { text: 'The body text carries the bulk of the sampled characters.', x: 72, y: 620, size: 12 },
      { text: 'so the dominant cluster is the paragraph style.', x: 72, y: 600, size: 12 },
      { text: 'A third line makes the body cluster large enough to win.', x: 72, y: 580, size: 12 },
      { text: 'A fourth line keeps the inferred margins stable.', x: 72, y: 560, size: 12 },
    ]],
  }), { documentType: 'thesis' });

  const styles = Object.entries(template.styles);
  const toc = styles.find(([, style]) => style.font?.size === 21);
  const heading = styles.find(([, style]) => style.font?.size === 30);
  assert.equal(toc && template.styleRoles[toc[0]], 'paragraph_no_indent');
  assert.equal(heading && template.styleRoles[heading[0]], 'heading1');
});

test('reads the running head and the footer page number out of the bands', async () => {
  // The head and the number are printed outside the text column, so they used
  // to be excluded from the margins and then dropped. Here they become the
  // template's own header/footer slots.
  const page = (head, number) => [
    { text: head, x: 200, y: 760, size: 9 },
    { text: 'The body text carries the bulk of the sampled characters.', x: 72, y: 700, size: 12 },
    { text: 'so the dominant cluster is the paragraph style.', x: 72, y: 680, size: 12 },
    { text: 'A third line makes the body cluster large enough to win.', x: 72, y: 660, size: 12 },
    { text: 'A fourth line keeps the inferred margins stable.', x: 72, y: 640, size: 12 },
    { text: number, x: 300, y: 40, size: 9 },
  ];
  const template = await parsePdfTemplate(buildLinedPdf({
    pages: [page('Odd Head', '1'), page('Even Head', '2')],
  }), { documentType: 'thesis' });

  assert.equal(template.page.headers?.default?.text, 'Odd Head');
  assert.equal(template.page.headers?.even?.text, 'Even Head');
  assert.equal(template.page.footers?.default?.pageNumber, true);
  assert.equal(template.page.headers?.default?.text?.includes('Odd Head'), true);
});

test('reads a three-line table’s rules and header fill out of the PDF', async () => {
  // A Chinese thesis table is three rules and a shaded header band; the
  // renderer needs the rule weights (w:sz) and the fill (w:shd).
  const template = await parsePdfTemplate(buildLinedPdf({
    pages: [[
      { text: 'The body text carries the bulk of the sampled characters.', x: 100, y: 640, size: 12 },
      { text: 'so the dominant cluster is the paragraph style.', x: 100, y: 620, size: 12 },
      { text: 'A third line makes the body cluster large enough to win.', x: 100, y: 600, size: 12 },
    ]],
    shapes: [[
      { type: 'line', x1: 100, y1: 700, x2: 500, y2: 700, width: 1.5 },
      { type: 'line', x1: 100, y1: 680, x2: 500, y2: 680, width: 0.5 },
      { type: 'line', x1: 100, y1: 650, x2: 500, y2: 650, width: 1.5 },
      { type: 'rect', x: 100, y: 680, w: 400, h: 20, color: '#D9D9D9' },
    ]],
  }), { documentType: 'thesis' });

  assert.equal(template.tableFormat?.borders?.top?.size, 12);
  assert.equal(template.tableFormat?.borders?.bottom?.size, 12);
  assert.equal(template.tableFormat?.headerFill, 'D9D9D9');
  // Three rules with no verticals is a three-line table, not a grid.
  assert.equal(template.tableFormat?.borders?.insideV, undefined);
});

test('reads a full-grid table’s vertical rules', async () => {
  const template = await parsePdfTemplate(buildLinedPdf({
    pages: [[
      { text: 'The body text carries the bulk of the sampled characters.', x: 100, y: 600, size: 12 },
      { text: 'so the dominant cluster is the paragraph style.', x: 100, y: 580, size: 12 },
      { text: 'A third line makes the body cluster large enough to win.', x: 100, y: 560, size: 12 },
    ]],
    shapes: [[
      { type: 'line', x1: 100, y1: 700, x2: 500, y2: 700, width: 0.5 },
      { type: 'line', x1: 100, y1: 680, x2: 500, y2: 680, width: 0.5 },
      { type: 'line', x1: 100, y1: 660, x2: 500, y2: 660, width: 0.5 },
      { type: 'line', x1: 100, y1: 640, x2: 500, y2: 640, width: 0.5 },
      { type: 'line', x1: 100, y1: 640, x2: 100, y2: 700, width: 0.5 },
      { type: 'line', x1: 300, y1: 640, x2: 300, y2: 700, width: 0.5 },
      { type: 'line', x1: 500, y1: 640, x2: 500, y2: 700, width: 0.5 },
    ]],
  }), { documentType: 'thesis' });

  assert.equal(template.tableFormat?.borders?.insideV?.size, 4);
  assert.equal(template.tableFormat?.borders?.insideH?.size, 4);
  // Four horizontal rules means every row is ruled, not just the header.
  assert.equal(template.tableFormat?.borders?.top?.size, 4);
});

test('carries w:keepNext and w:keepLines through to the style', async () => {
  // Both flags were parsed into RawParagraphProps and merged into every resolved
  // style, but `ParagraphFormatting` had no field to hold them, so they were
  // dropped before the renderer — and a heading that strands itself at the foot
  // of a page is exactly what `keepNext` exists to prevent.
  const zip = new JSZip();
  zip.file('word/styles.xml', `<?xml version="1.0"?>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>
      <w:style w:type="paragraph" w:styleId="h"><w:name w:val="heading 1"/>
        <w:pPr><w:outlineLvl w:val="0"/><w:keepNext/><w:keepLines/></w:pPr>
      </w:style>
    </w:styles>`);
  zip.file('word/document.xml', `<?xml version="1.0"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:pPr><w:pStyle w:val="h"/></w:pPr><w:r><w:t>H</w:t></w:r></w:p>
    </w:body></w:document>`);

  const template = await parseTemplate(await zip.generateAsync({ type: 'nodebuffer' }), { documentType: 'thesis' });
  assert.equal(template.styles.h.paragraph.keepNext, true);
  assert.equal(template.styles.h.paragraph.keepLines, true);
});

test('converts character units with the inherited font size, not the own style', async () => {
  // `*Chars` / `*Lines` are hundredths of a character or a line, and a character
  // is as wide as the style's *effective* size. The conversion used to run while
  // the style bag was being read — before the `basedOn` chain was resolved — so
  // a style that inherits 10.5pt converted with the 12pt fallback: 480 twips
  // where 420 is right.
  const zip = new JSZip();
  zip.file('word/styles.xml', `<?xml version="1.0"?>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>
      <w:style w:type="paragraph" w:default="1" w:styleId="base">
        <w:name w:val="Normal"/><w:rPr><w:sz w:val="21"/></w:rPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="child">
        <w:name w:val="Body Text"/><w:basedOn w:val="base"/>
        <w:pPr><w:ind w:firstLineChars="200"/></w:pPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="own">
        <w:name w:val="Own Size"/><w:rPr><w:sz w:val="21"/></w:rPr>
        <w:pPr><w:ind w:firstLineChars="200"/></w:pPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="explicit">
        <w:name w:val="Explicit Twips"/><w:basedOn w:val="child"/>
        <w:pPr><w:ind w:firstLine="999"/></w:pPr>
      </w:style>
    </w:styles>`);
  zip.file('word/document.xml', `<?xml version="1.0"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:pPr><w:pStyle w:val="child"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>
    </w:body></w:document>`);

  const template = await parseTemplate(await zip.generateAsync({ type: 'nodebuffer' }), { documentType: 'thesis' });

  // 10.5pt (21 half-points) × two characters = 420 twips. The old code produced
  // 480 because it converted with the 12pt fallback.
  assert.equal(template.styles.child.font.size, 21);
  assert.equal(template.styles.child.paragraph.firstLineIndent, 420);

  // A style that declares the size itself converts identically.
  assert.equal(template.styles.own.paragraph.firstLineIndent, 420);

  // An explicit twips value overrides the inherited character form, rather than
  // losing to it in the later conversion.
  assert.equal(template.styles.explicit.paragraph.firstLineIndent, 999);
});

test('reads w:docGrid from the section properties', async () => {
  // The writer library always emits its own line grid, so the template's grid has
  // to survive parsing for the renderer to have anything to write back.
  const zip = new JSZip();
  zip.file('word/styles.xml', `<?xml version="1.0"?>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>
      <w:style w:type="paragraph" w:styleId="a"><w:name w:val="Normal"/><w:rPr><w:sz w:val="24"/></w:rPr></w:style>
    </w:styles>`);
  zip.file('word/document.xml', `<?xml version="1.0"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:sectPr><w:pgSz w:w="11906" w:h="16838"/>
        <w:pgMar w:top="1701" w:bottom="1134" w:left="1701" w:right="1701"/>
        <w:docGrid w:type="linesAndChars" w:linePitch="312" w:charSpace="0"/>
      </w:sectPr>
    </w:body></w:document>`);

  const template = await parseTemplate(await zip.generateAsync({ type: 'nodebuffer' }), { documentType: 'thesis' });
  assert.deepEqual(template.page.docGrid, { type: 'linesAndChars', linePitch: 312, charSpace: 0 });
});

test('the parser package exposes only its two entry points', async () => {
  // Five internal helpers used to be re-exported, but their signatures name
  // `RawStyle` / `RawParagraphProps`, which are not exported — an external caller
  // could not even write the parameter type, and nothing in this repository
  // imported them.
  const parser = await import('../packages/template-parser/dist/index.js');
  assert.deepEqual(Object.keys(parser).sort(), ['parsePdfTemplate', 'parseTemplate']);
});

test('a long basedOn chain still resolves every inherited property', async () => {
  // The chain used to be walked from scratch for every inheriting style. It is
  // now memoised; this pins that the shared result is still correct — the root
  // contributes its font, each link contributes its own override, and the leaf
  // carries both.
  const styles = ['<w:style w:type="paragraph" w:default="1" w:styleId="s0"><w:name w:val="Normal"/><w:rPr><w:sz w:val="24"/><w:rFonts w:eastAsia="宋体"/></w:rPr></w:style>'];
  for (let i = 1; i < 40; i++) {
    styles.push(
      `<w:style w:type="paragraph" w:styleId="s${i}"><w:name w:val="S${i}"/><w:basedOn w:val="s${i - 1}"/>`
      + `<w:rPr><w:b/></w:rPr><w:pPr><w:ind w:left="${i * 10}"/></w:pPr></w:style>`,
    );
  }
  const zip = new JSZip();
  zip.file('word/styles.xml', `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${styles.join('')}</w:styles>`);
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>`);

  const template = await parseTemplate(await zip.generateAsync({ type: 'nodebuffer' }), { documentType: 'thesis' });
  const leaf = template.styles['s39'];
  assert.equal(leaf.font.size, 24, 'root font size reaches the leaf');
  assert.equal(leaf.font.eastAsia, '宋体', 'root eastAsia reaches the leaf');
  assert.equal(leaf.font.bold, true, 'the leaf own override applies');
  assert.equal(leaf.paragraph.leftIndent, 390, 'the nearest indent wins');
});

test('reads numbering.xml and the styles it binds', async () => {
  // A template numbers its chapters by binding styles to w:num instances through
  // w:numPr. None of that was read, so a heading number could only come from the
  // content file.
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const zip = new JSZip();
  zip.file('word/styles.xml', `<?xml version="1.0"?><w:styles xmlns:w="${W}">
    <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="24"/></w:rPr></w:style>
    <w:style w:type="paragraph" w:styleId="ch"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>
      <w:pPr><w:outlineLvl w:val="0"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style>
  </w:styles>`);
  zip.file('word/numbering.xml', `<?xml version="1.0"?><w:numbering xmlns:w="${W}">
    <w:abstractNum w:abstractNumId="0">
      <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="ch"/><w:lvlText w:val="第%1章"/></w:lvl>
    </w:abstractNum>
    <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  </w:numbering>`);
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document xmlns:w="${W}"><w:body>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`);

  const template = await parseTemplate(await zip.generateAsync({ type: 'nodebuffer' }), { documentType: 'thesis' });

  assert.equal(template.numbering.length, 1);
  assert.equal(template.numbering[0].numId, 1);
  assert.deepEqual(
    template.numbering[0].levels.map(l => ({ ilvl: l.ilvl, text: l.text, format: l.format, pStyle: l.pStyle })),
    [{ ilvl: 0, text: '第%1章', format: 'decimal', pStyle: 'ch' }],
  );
  assert.deepEqual(template.styles.ch.paragraph.numbering, { numId: 1, ilvl: 0 });
});

// ── Real-world template compatibility ────────────────────────
//
// These four are unmodified .docx files downloaded from their publishers, kept
// as fixtures so a parser change is checked against something it did not
// generate. Their provenance:
//
//   acm-journal.docx      ACM interim Word template (acm.org)
//   cn-cjstp.docx         中国科技期刊研究 论文模板 (cjstp.cn)
//   cn-dlmu-thesis.docx   《大连海事大学学报》论文模板 (cnki.net)
//   elsevier-hardwarex.docx  HardwareX (Elsevier) article template (zenodo.org)

const FIXTURE_DIR = join(dirname(fileURLToPath(import.meta.url)), 'fixtures', 'templates');
const FIXTURES = new URL('./fixtures/templates/', import.meta.url);
const fixture = (name) => readFileSync(join(FIXTURE_DIR, name));

test('parses every real-world template without throwing', async () => {
  for (const name of ['acm-journal.docx', 'cn-cjstp.docx', 'cn-dlmu-thesis.docx', 'elsevier-hardwarex.docx']) {
    const template = await parseTemplate(fixture(name), { documentType: 'thesis', organization: 'Fixture' });
    assert.ok(Object.keys(template.styles).length > 0, `${name}: expected at least one style`);
    assert.ok(template.page.width > 0 && template.page.height > 0, `${name}: expected page geometry`);
  }
});

test('picks the body style a real template actually uses', async () => {
  // `cn-dlmu-thesis.docx` calls its body style 硕士正文样式 and applies it to its
  // body paragraphs; the old pattern list only knew the stock names and settled
  // for `Normal`. `elsevier-hardwarex.docx` calls it `Standard`.
  const dlmu = await parseTemplate(fixture('cn-dlmu-thesis.docx'), { documentType: 'thesis' });
  const dlmuBody = dlmu.styles[dlmu.roleWinners.paragraph];
  assert.equal(dlmuBody.font.size, 24, 'expected the 12pt 硕士正文样式, not Normal at 21');
  assert.equal(dlmuBody.font.eastAsia, '宋体');

  const hardwarex = await parseTemplate(fixture('elsevier-hardwarex.docx'), { documentType: 'thesis' });
  assert.equal(hardwarex.roleWinners.paragraph, 'Standard');
});

test('a template that defines no heading styles still parses and still warns', async () => {
  // Journals routinely format headings directly rather than through named styles.
  // That is not a parser failure — the renderer falls back to the role defaults —
  // but the template must not pretend it declared heading formatting.
  const cjstp = await parseTemplate(fixture('cn-cjstp.docx'), { documentType: 'thesis' });
  const headingIds = Object.keys(cjstp.styleRoles).filter(id => cjstp.styleRoles[id].startsWith('heading'));
  assert.equal(headingIds.length, 0, 'this fixture genuinely declares no heading style');
  // Its body style is still found, so a document built from it is not left unstyled.
  assert.ok(cjstp.roleWinners.paragraph, 'expected a body style');
});

test('a heading name a real template uses is recognised', async () => {
  // ACM and the Zenodo journal template both name their heading styles `Head1`
  // / `Heading1` rather than Word's stock `Heading 1`.
  const acm = await parseTemplate(fixture('acm-journal.docx'), { documentType: 'thesis' });
  const headingRoles = Object.entries(acm.styleRoles)
    .filter(([, role]) => role === 'heading1')
    .map(([id]) => id);
  assert.ok(
    headingRoles.some(id => /^Head1|^Head\s*1|^Heading1/i.test(id)),
    `expected a level-1 heading among: ${headingRoles.slice(0, 10).join(', ')}`,
  );
});
