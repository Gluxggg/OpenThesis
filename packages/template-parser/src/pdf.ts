import type {
  BlockType,
  DocumentTemplate,
  PageSettings,
  ParagraphFormatting,
  ParagraphStyle,
  TemplateMeta,
} from '@openthesis/document-schema';

interface PdfTextItem {
  str: string;
  transform: number[];
  width: number;
  height: number;
  fontName: string;
}

interface PdfTextStyle { fontFamily?: string }

interface PdfLine {
  page: number;
  pageWidth: number;
  pageHeight: number;
  text: string;
  x: number;
  y: number;
  right: number;
  fontFamily: string;
  fontSize: number;
  bold: boolean;
  italic: boolean;
}

interface StyleCluster {
  id: string;
  lines: PdfLine[];
  fontFamily: string;
  fontSize: number;
  bold: boolean;
  italic: boolean;
  alignment: ParagraphFormatting['alignment'];
  weight: number;
}

const POINTS_TO_TWIPS = 20;
const MAX_SAMPLED_PAGES = 50;

/**
 * CSS families PDF.js falls back to when a font carries no usable PostScript
 * name. A PDF template never says "sans-serif"; it says Arial. These mappings
 * turn the fallback back into a family Word knows.
 */
const GENERIC_FONT_FAMILIES: Record<string, string> = {
  'sans-serif': 'Arial',
  serif: 'Times New Roman',
  monospace: 'Courier New',
};

/**
 * The CJK faces a Chinese writing guide actually names. PDF.js reports the
 * PostScript name (`BCDEEE+SimSun`), and Word wants `宋体`, so each alias maps
 * to the family the template's own styles use.
 */
const CJK_FONT_FAMILIES: Array<[RegExp, string]> = [
  [/simsun|songti|stsong|nsimsun|宋体/i, '宋体'],
  [/simhei|heiti|sthei|黑体/i, '黑体'],
  [/simkai|kaiti|stkaiti|楷体/i, '楷体'],
  [/simfang|fangsong|stfangsong|仿宋/i, '仿宋'],
  [/microsoft ?yahei|yahei|微软雅黑/i, '微软雅黑'],
  [/dengxian|等线/i, '等线'],
];

/** The Latin half of the font pair, mapped from PostScript names to Word ones. */
const LATIN_FONT_FAMILIES: Array<[RegExp, string]> = [
  [/times/i, 'Times New Roman'],
  [/arial|helvetica/i, 'Arial'],
  [/courier/i, 'Courier New'],
  [/cambria/i, 'Cambria'],
  [/calibri/i, 'Calibri'],
  [/georgia/i, 'Georgia'],
  [/verdana/i, 'Verdana'],
  [/palatino/i, 'Palatino Linotype'],
];

/**
 * Where a running head or a page number sits, as a fraction of the page.
 *
 * These lines are printed outside the text column, so measuring the margins
 * from every line on the page measured the header instead of the body: the
 * inferred top margin came out at 1.5 cm for a template whose text starts at
 * 2.5 cm.
 */
const HEADER_BAND = 0.9;
const FOOTER_BAND = 0.08;

/** The operator PDF.js uses for `setFont`, and the page type it is read from. */
interface FontNameSource {
  getOperatorList(): Promise<{ fnArray: number[]; argsArray: unknown[] }>;
  commonObjs: { get(id: string): unknown };
  objs: { get(id: string): unknown };
}

/** Infer a reusable OpenThesis style template from a text-based PDF. */
export async function parsePdfTemplate(
  buffer: Buffer | ArrayBuffer,
  metaOverrides?: Partial<TemplateMeta>,
): Promise<DocumentTemplate> {
  const { getDocument, OPS } = await import('pdfjs-dist/legacy/build/pdf.mjs');
  const data = buffer instanceof ArrayBuffer
    ? new Uint8Array(buffer.slice(0))
    : Uint8Array.from(buffer);
  const loadingTask = getDocument({ data, useSystemFonts: true });
  const pdf = await loadingTask.promise;
  const sampledPages = Math.min(pdf.numPages, MAX_SAMPLED_PAGES);
  const lines: PdfLine[] = [];
  const pageSizes: Array<{ width: number; height: number }> = [];

  try {
    for (let pageNumber = 1; pageNumber <= sampledPages; pageNumber += 1) {
      const page = await pdf.getPage(pageNumber);
      const viewport = page.getViewport({ scale: 1 });
      pageSizes.push({ width: viewport.width, height: viewport.height });
      // `content.styles` only carries a fallback CSS family ("sans-serif"),
      // which used to become Arial for every Chinese font. The real name lives
      // on the page's font objects, so it is read from there first.
      const fontNames = await readFontNames(page as unknown as FontNameSource, OPS.setFont);
      const content = await page.getTextContent();
      const items = content.items
        .filter(item => 'str' in item && Boolean(item.str.trim()))
        .map(item => item as unknown as PdfTextItem);
      lines.push(...groupItemsIntoLines(
        items,
        content.styles as Record<string, PdfTextStyle>,
        fontNames,
        pageNumber,
        viewport.width,
        viewport.height,
      ));
      page.cleanup();
    }
  } finally {
    await loadingTask.destroy();
  }

  if (!lines.some(line => line.text.trim())) {
    throw new Error('No extractable PDF text found. The PDF may be scanned; run OCR first and retry.');
  }

  const clusters = buildStyleClusters(lines);
  const body = clusters.reduce((largest, current) => current.weight > largest.weight ? current : largest);
  const page = inferPageSettings(pageSizes[0], lines, body);
  const styles: Record<string, ParagraphStyle> = {};
  const styleRoles: Record<string, BlockType> = {};
  const headingClusters = clusters
    .filter(cluster => cluster.id !== body.id && cluster.fontSize > body.fontSize * 1.1)
    .sort((a, b) => b.fontSize - a.fontSize);

  for (const cluster of clusters) {
    styles[cluster.id] = clusterToParagraphStyle(cluster, body);
    styleRoles[cluster.id] = inferRole(cluster, body, headingClusters);
  }

  const warnings = [
    'PDF styles are inferred from visible text geometry; review the generated template JSON before production use.',
  ];
  if (pdf.numPages > sampledPages) {
    // ponytail: 50 pages bound runtime; add a configurable sampler if long standards need full-document analysis.
    warnings.push(`Only the first ${sampledPages} of ${pdf.numPages} pages were sampled.`);
  }
  if (pageSizes.some(size => !samePageSize(size, pageSizes[0]))) {
    warnings.push('The PDF contains mixed page sizes; the first page size is used for the template.');
  }
  if (lines.reduce((sum, line) => sum + line.text.length, 0) < sampledPages * 20) {
    warnings.push('Very little text was extracted; OCR or layout inference may be incomplete.');
  }

  return {
    meta: {
      ...metaOverrides,
      organization: metaOverrides?.organization?.trim() || 'Unknown Organization',
      name: metaOverrides?.name?.trim() || 'PDF-derived Template',
      documentType: metaOverrides?.documentType || 'thesis',
      sourceFormat: 'pdf',
      pageCount: pdf.numPages,
      warnings,
      parserVersion: '0.3.0',
      parsedAt: new Date().toISOString(),
    },
    page,
    styles,
    styleRoles,
  };
}

/**
 * The real font name behind each text-content font id.
 *
 * PDF.js reports `sans-serif` / `serif` / `monospace` as the CSS family of an
 * embedded font and keeps the PostScript name on the font object instead. This
 * reads the operator list once per page, which is what loads those objects, and
 * maps each id to `BCDEEE+SimSun` and friends. A page whose operators cannot be
 * read still yields its text; the caller falls back to the CSS family then.
 */
async function readFontNames(
  page: FontNameSource,
  setFontOperator: number,
): Promise<Map<string, string>> {
  const names = new Map<string, string>();
  let operators: { fnArray: number[]; argsArray: unknown[] };
  try {
    operators = await page.getOperatorList();
  } catch {
    return names;
  }

  const ids = new Set<string>();
  for (let index = 0; index < operators.fnArray.length; index += 1) {
    if (operators.fnArray[index] !== setFontOperator) continue;
    const args = operators.argsArray[index] as unknown[] | undefined;
    const id = args?.[0];
    if (typeof id === 'string' && id) ids.add(id);
  }
  for (const id of ids) {
    const name = fontObjectName(page, id);
    if (name) names.set(id, name);
  }
  return names;
}

function fontObjectName(page: FontNameSource, id: string): string {
  for (const store of [page.commonObjs, page.objs]) {
    try {
      const font = store?.get(id) as { name?: unknown } | null | undefined;
      if (font && typeof font.name === 'string' && font.name) return font.name;
    } catch {
      // The object belongs to the other store; try that one.
    }
  }
  return '';
}

function groupItemsIntoLines(
  items: PdfTextItem[],
  styles: Record<string, PdfTextStyle>,
  fontNames: Map<string, string>,
  page: number,
  pageWidth: number,
  pageHeight: number,
): PdfLine[] {
  const groups = new Map<number, PdfTextItem[]>();
  for (const item of items) {
    const y = Number(item.transform[5] ?? 0);
    const key = Math.round(y * 2) / 2;
    const existingKey = [...groups.keys()].find(candidate => Math.abs(candidate - key) <= 1);
    groups.set(existingKey ?? key, [...(groups.get(existingKey ?? key) ?? []), item]);
  }

  return [...groups.entries()].map(([y, lineItems]) => {
    lineItems.sort((a, b) => Number(a.transform[4] ?? 0) - Number(b.transform[4] ?? 0));
    const weightedStyles = new Map<string, number>();
    for (const item of lineItems) {
      weightedStyles.set(item.fontName, (weightedStyles.get(item.fontName) ?? 0) + item.str.length);
    }
    const dominantFont = [...weightedStyles.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? '';
    const dominantItem = lineItems.find(item => item.fontName === dominantFont) ?? lineItems[0];
    const fontSize = inferFontSize(dominantItem);
    // The font object's name ("BCDEEE+SimSun") beats the CSS fallback
    // ("sans-serif"), which would otherwise turn every Chinese face into Arial.
    const resolvedFont = fontNames.get(dominantFont) || styles[dominantFont]?.fontFamily || dominantFont;
    const fontFamily = normalizeFontFamily(resolvedFont);
    const text = joinTextItems(lineItems, fontSize);
    const x = Math.min(...lineItems.map(item => Number(item.transform[4] ?? 0)));
    const right = Math.max(...lineItems.map(item => Number(item.transform[4] ?? 0) + item.width));
    const fontDescriptor = `${fontFamily} ${resolvedFont}`;
    return {
      page, pageWidth, pageHeight, text, x, y, right, fontFamily, fontSize,
      bold: /bold|black|heavy|semibold|demi/i.test(fontDescriptor),
      italic: /italic|oblique/i.test(fontDescriptor),
    };
  });
}

function joinTextItems(items: PdfTextItem[], fontSize: number): string {
  let text = '';
  let right: number | undefined;
  for (const item of items) {
    const x = Number(item.transform[4] ?? 0);
    if (right !== undefined && x - right > fontSize * 0.25 && !text.endsWith(' ')) text += ' ';
    text += item.str;
    right = x + item.width;
  }
  return text.trim();
}

function inferFontSize(item: PdfTextItem): number {
  const [a = 0, b = 0, c = 0, d = 0] = item.transform;
  return Math.max(Math.hypot(a, b), Math.hypot(c, d), item.height || 0, 1);
}

function normalizeFontFamily(font: string): string {
  const cleaned = font.replace(/^['"]|['"]$/g, '').replace(/^[A-Z]{6}\+/, '').trim();
  const generic = GENERIC_FONT_FAMILIES[cleaned.toLowerCase()];
  if (generic) return generic;
  return cleaned && !/^g_d\d+_f\d+$/i.test(cleaned) ? cleaned : 'Times New Roman';
}

/** `SimSun` → `宋体`, so the template names the face the guide names. */
function cjkFontFamily(family: string): string | undefined {
  return CJK_FONT_FAMILIES.find(([pattern]) => pattern.test(family))?.[1];
}

/** `TimesNewRomanPS-BoldMT` → `Times New Roman`, a name Word resolves. */
function latinFontFamily(family: string): string | undefined {
  return LATIN_FONT_FAMILIES.find(([pattern]) => pattern.test(family))?.[1];
}

function buildStyleClusters(lines: PdfLine[]): StyleCluster[] {
  const grouped = new Map<string, PdfLine[]>();
  for (const line of lines) {
    const halfPoints = Math.max(2, Math.round(line.fontSize * 2));
    const key = `${line.fontFamily}|${halfPoints}|${line.bold}|${line.italic}`;
    grouped.set(key, [...(grouped.get(key) ?? []), line]);
  }

  return [...grouped.values()]
    .sort((a, b) => {
      const sizeDifference = b[0].fontSize - a[0].fontSize;
      return sizeDifference || a[0].fontFamily.localeCompare(b[0].fontFamily);
    })
    .map((clusterLines, index) => ({
      id: `pdf_style_${String(index + 1).padStart(2, '0')}`,
      lines: clusterLines,
      fontFamily: clusterLines[0].fontFamily,
      fontSize: Math.max(1, Math.round(clusterLines[0].fontSize * 2) / 2),
      bold: clusterLines[0].bold,
      italic: clusterLines[0].italic,
      alignment: inferAlignment(clusterLines),
      weight: clusterLines.reduce((sum, line) => sum + line.text.length, 0),
    }));
}

function inferAlignment(lines: PdfLine[]): ParagraphFormatting['alignment'] {
  const centered = lines.filter(line => {
    const left = line.x;
    const right = line.pageWidth - line.right;
    return Math.abs(left - right) <= Math.max(12, line.pageWidth * 0.04);
  }).length;
  const rightAligned = lines.filter(line => line.pageWidth - line.right < line.pageWidth * 0.08).length;
  if (centered > lines.length / 2) return 'center';
  if (rightAligned > lines.length / 2) return 'right';
  const wide = lines.filter(line => (line.right - line.x) / line.pageWidth > 0.6).length;
  return wide > lines.length / 2 ? 'justified' : 'left';
}

function clusterToParagraphStyle(
  cluster: StyleCluster,
  body: StyleCluster,
): ParagraphStyle {
  const hasCjk = cluster.lines.some(line => /[\u3400-\u9fff]/u.test(line.text));
  const cjk = cjkFontFamily(cluster.fontFamily);
  const latin = latinFontFamily(cluster.fontFamily);
  const lineSpacing = cluster.id === body.id ? inferLineSpacing(cluster.lines) : undefined;
  return {
    font: {
      // One PDF font stands for both halves of the pair: a CJK face names the
      // eastAsia side and takes Times New Roman as its Latin partner, while an
      // English face keeps its own family on the Latin side.
      name: cjk || hasCjk ? 'Times New Roman' : latin ?? cluster.fontFamily,
      eastAsia: cjk ?? '宋体',
      size: Math.round(cluster.fontSize * 2),
      bold: cluster.bold || undefined,
      italic: cluster.italic || undefined,
    },
    paragraph: {
      alignment: cluster.alignment,
      firstLineIndent: inferFirstLineIndent(cluster, body),
    },
    lineSpacing,
  };
}

function inferFirstLineIndent(cluster: StyleCluster, body: StyleCluster): number {
  if (cluster.id !== body.id || cluster.lines.length < 4) return 0;
  const starts = cluster.lines.map(line => line.x).sort((a, b) => a - b);
  // The column's left edge is its leftmost line. Using the 20th percentile put
  // the measured left margin inside the column whenever a paragraph's indented
  // first line was at the percentile, which measured the indent as margin.
  const left = quantile(starts, 0.05);
  const indented = starts.filter(start => start - left >= body.fontSize * 1.5 && start - left <= body.fontSize * 3);
  if (indented.length < Math.max(2, starts.length * 0.1)) return 0;
  const indent = median(indented.map(start => start - left));
  return Math.round(indent * POINTS_TO_TWIPS);
}

function inferLineSpacing(lines: PdfLine[]): number | undefined {
  const gaps: number[] = [];
  const byPage = new Map<number, PdfLine[]>();
  for (const line of lines) byPage.set(line.page, [...(byPage.get(line.page) ?? []), line]);
  for (const pageLines of byPage.values()) {
    pageLines.sort((a, b) => b.y - a.y);
    for (let i = 1; i < pageLines.length; i += 1) {
      const gap = pageLines[i - 1].y - pageLines[i].y;
      if (gap >= pageLines[i].fontSize * 0.8 && gap <= pageLines[i].fontSize * 3) gaps.push(gap);
    }
  }
  return gaps.length ? Math.round(median(gaps) * POINTS_TO_TWIPS) : undefined;
}

function inferRole(
  cluster: StyleCluster,
  body: StyleCluster,
  headings: StyleCluster[],
): BlockType {
  if (cluster.id === body.id) return 'paragraph';
  const sample = cluster.lines.map(line => line.text).join('\n');
  // A table-of-contents line reads "1 Introduction ........ 3": it starts like
  // a numbered heading but is set in the body size. Requiring a heading to be
  // larger than the body keeps a TOC entry — or a numbered paragraph — out of
  // the heading roles, which is what used to make the 10.5 pt TOC an h1.
  const dotLeader = /(?:\.{3,}|…{2,})\s*\d+\s*$/m.test(sample);
  const largerThanBody = cluster.fontSize > body.fontSize * 1.05;
  const numberedHeading = !dotLeader && largerThanBody
    ? sample.match(/^(?:第.{1,12}[章节]|chapter\s+\d+|\d+(?:\.\d+){0,3}\s+\S)/im)?.[0]
    : undefined;
  if (numberedHeading) {
    const dots = (numberedHeading.match(/\./g) ?? []).length;
    return `heading${Math.min(4, dots + 1)}` as BlockType;
  }
  if (cluster.alignment === 'center' && cluster.fontSize > body.fontSize * 1.2) return 'centered_text';
  const headingIndex = headings.findIndex(heading => heading.id === cluster.id);
  if (headingIndex >= 0) return `heading${Math.min(4, headingIndex + 1)}` as BlockType;
  return cluster.alignment === 'center' ? 'centered_text' : 'paragraph_no_indent';
}

function inferPageSettings(
  firstPage: { width: number; height: number },
  lines: PdfLine[],
  body: StyleCluster,
): PageSettings {
  const bodyLines = body.lines.length ? body.lines : lines;
  // A running head and a page number are printed outside the text column, so
  // they are not part of the page's margins. Measuring every line measured the
  // header instead: the built-in USTB template's 2.5 cm top margin came out as
  // 1.5 cm, because its running head sits there.
  const contentLines = lines.filter(line => !isMarginalLine(line, body));
  const measured = contentLines.length ? contentLines : lines;
  // The left edge is the column's leftmost line: a paragraph's indented first
  // line sits inside the column, and the 20th percentile measured it as margin.
  const leftPoints = quantile(bodyLines.map(line => line.x), 0.05);
  const rightPoints = quantile(bodyLines.map(line => line.pageWidth - line.right), 0.05);
  const pageGroups = new Map<number, PdfLine[]>();
  for (const line of measured) pageGroups.set(line.page, [...(pageGroups.get(line.page) ?? []), line]);
  // The margin is where text can start on the page, so the pages whose first
  // element sits highest drive it. A chapter heading carries 段前 spacing and
  // puts its glyphs below the margin, so the median measured that spacing as
  // margin (3.6 cm for a section whose text starts at 2.5 cm).
  const topPoints = quantile([...pageGroups.values()].map(pageLines => {
    const topLine = pageLines.reduce((top, line) => line.y > top.y ? line : top);
    return Math.max(0, topLine.pageHeight - topLine.y - topLine.fontSize);
  }), 0.25);
  const bottomPoints = median([...pageGroups.values()].map(pageLines => {
    const bottomLine = pageLines.reduce((bottom, line) => line.y < bottom.y ? line : bottom);
    return Math.max(0, bottomLine.y - bottomLine.fontSize * 0.25);
  }));
  const columns = inferColumns(bodyLines, firstPage.width);
  return {
    width: Math.round(firstPage.width * POINTS_TO_TWIPS),
    height: Math.round(firstPage.height * POINTS_TO_TWIPS),
    margins: {
      top: toMarginTwips(topPoints, firstPage.height),
      bottom: toMarginTwips(bottomPoints, firstPage.height),
      left: toMarginTwips(leftPoints, firstPage.width),
      right: toMarginTwips(rightPoints, firstPage.width),
    },
    columns,
    ...(columns > 1 ? { columnGutter: inferColumnGutter(bodyLines, firstPage.width) } : {}),
  };
}

/** A small line printed in the page's running-head or footer band. */
function isMarginalLine(line: PdfLine, body: StyleCluster): boolean {
  const small = line.fontSize < body.fontSize * 0.95;
  const inHeaderBand = line.y > line.pageHeight * HEADER_BAND;
  const inFooterBand = line.y < line.pageHeight * FOOTER_BAND;
  return small && (inHeaderBand || inFooterBand);
}

function inferColumns(lines: PdfLine[], pageWidth: number): number {
  if (lines.length < 6) return 1;
  const starts = lines.map(line => line.x);
  const first = quantile(starts, 0.2);
  const second = starts.filter(start => start - first > pageWidth * 0.25);
  return second.length >= Math.max(3, lines.length * 0.2) ? 2 : 1;
}

function inferColumnGutter(lines: PdfLine[], pageWidth: number): number {
  const starts = lines.map(line => line.x).sort((a, b) => a - b);
  const first = quantile(starts, 0.2);
  const second = quantile(starts.filter(start => start - first > pageWidth * 0.25), 0.2);
  const firstColumnRights = lines.filter(line => line.x < (first + second) / 2).map(line => line.right);
  const gutter = Math.max(12, second - quantile(firstColumnRights, 0.8));
  return Math.round(gutter * POINTS_TO_TWIPS);
}

function toMarginTwips(points: number, pageDimension: number): number {
  const inferred = points > pageDimension * 0.2 ? 72 : points;
  const clamped = Math.min(pageDimension * 0.2, Math.max(18, inferred));
  return Math.round(clamped * POINTS_TO_TWIPS);
}

function samePageSize(
  a: { width: number; height: number },
  b: { width: number; height: number },
): boolean {
  return Math.abs(a.width - b.width) < 1 && Math.abs(a.height - b.height) < 1;
}

function median(values: number[]): number {
  return quantile(values, 0.5);
}

function quantile(values: number[], fraction: number): number {
  if (!values.length) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  return sorted[Math.min(sorted.length - 1, Math.max(0, Math.floor((sorted.length - 1) * fraction)))];
}
