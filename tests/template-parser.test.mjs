import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { parseTemplate } from '../packages/template-parser/dist/index.js';

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
