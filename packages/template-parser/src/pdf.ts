import type {
  BlockType,
  BorderSpec,
  DocumentTemplate,
  PageSettings,
  ParagraphFormatting,
  ParagraphStyle,
  RunningHeadSlots,
  TableFormat,
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
  const rules: PdfRule[] = [];
  const filledRects: PdfFilledRect[] = [];

  try {
    for (let pageNumber = 1; pageNumber <= sampledPages; pageNumber += 1) {
      const page = await pdf.getPage(pageNumber);
      const viewport = page.getViewport({ scale: 1 });
      pageSizes.push({ width: viewport.width, height: viewport.height });
      // `content.styles` only carries a fallback CSS family ("sans-serif"),
      // which used to become Arial for every Chinese font. The real name lives
      // on the page's font objects, so it is read from there first. The same
      // operator list also carries the drawn rules a table is made of.
      const graphics = await readPageGraphics(page as unknown as FontNameSource, OPS as unknown as OpCodes);
      rules.push(...graphics.rules.map(rule => ({ ...rule, page: pageNumber, pageHeight: viewport.height })));
      filledRects.push(...graphics.fills.map(rect => ({ ...rect, page: pageNumber })));
      const fontNames = graphics.fontNames;
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
  const page: PageSettings = {
    ...inferPageSettings(pageSizes[0], lines, body),
    ...inferRunningHeads(lines, body),
  };
  const tableFormat = inferTableFormat(rules, filledRects);
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
    // A head that differs by parity only prints correctly when Word is told the
    // document uses odd and even heads.
    evenAndOddHeaders: Boolean(page.headers?.even || page.footers?.even),
    ...(tableFormat ? { tableFormat } : {}),
    styles,
    styleRoles,
  };
}

/**
 * The operator codes this file needs from `pdfjs-dist`.
 *
 * The values are stable and come from the module itself; declaring the shape
 * here keeps the dynamic import in one place.
 */
interface OpCodes {
  setFont: number;
  setLineWidth: number;
  setStrokeRGBColor: number;
  setFillRGBColor: number;
  constructPath: number;
  stroke: number;
  closeStroke: number;
  fill: number;
  eoFill: number;
  fillStroke: number;
  eoFillStroke: number;
  closeFillStroke: number;
  closeEOFillStroke: number;
}

/** A drawn rule: one stroke segment, or one very thin filled rectangle. */
interface PdfRule {
  page: number;
  /** Height of the page the rule was drawn on, for the header/footer bands. */
  pageHeight?: number;
  orientation: 'horizontal' | 'vertical';
  /** Baseline coordinate: y for a horizontal rule, x for a vertical one. */
  position: number;
  start: number;
  end: number;
  /** Rendered thickness, in points. */
  width: number;
  color: string;
}

/** A filled rectangle large enough to be a cell shading rather than a rule. */
interface PdfFilledRect {
  page: number;
  x0: number;
  y0: number;
  x1: number;
  y1: number;
  color: string;
}

interface PageGraphics {
  fontNames: Map<string, string>;
  rules: PdfRule[];
  fills: PdfFilledRect[];
}

/** Segments shorter than this are underlines or decoration, not table rules. */
const MIN_RULE_LENGTH = 24;
/** A path this close to horizontal/vertical is a rule. */
const RULE_TOLERANCE = 1;
/** A filled rectangle thinner than this is a rule, not a shaded cell. */
const THIN_FILL = 2.5;

/**
 * Read one page's operator list once for the two things it is needed for: the
 * real font names behind the text, and the drawn rules a table is made of.
 *
 * A page whose operators cannot be read still yields its text; the caller falls
 * back to the CSS family then.
 */
async function readPageGraphics(page: FontNameSource, ops: OpCodes): Promise<PageGraphics> {
  const fontNames = new Map<string, string>();
  const rules: PdfRule[] = [];
  const fills: PdfFilledRect[] = [];
  let operators: { fnArray: number[]; argsArray: unknown[] };
  try {
    operators = await page.getOperatorList();
  } catch {
    return { fontNames, rules, fills };
  }

  const fontIds = new Set<string>();
  let lineWidth = 1;
  let strokeColor = '#000000';
  let fillColor = '#000000';

  for (let index = 0; index < operators.fnArray.length; index += 1) {
    const fn = operators.fnArray[index];
    const args = operators.argsArray[index] as unknown[] | undefined;
    if (fn === ops.setFont) {
      const id = args?.[0];
      if (typeof id === 'string' && id) fontIds.add(id);
      continue;
    }
    if (fn === ops.setLineWidth) {
      const value = Number(args?.[0]);
      if (Number.isFinite(value) && value > 0) lineWidth = value;
      continue;
    }
    if (fn === ops.setStrokeRGBColor) {
      strokeColor = String(args?.[0] ?? strokeColor);
      continue;
    }
    if (fn === ops.setFillRGBColor) {
      fillColor = String(args?.[0] ?? fillColor);
      continue;
    }
    if (fn !== ops.constructPath) continue;
    collectPath(args, ops, lineWidth, strokeColor, fillColor, rules, fills);
  }

  for (const id of fontIds) {
    const name = fontObjectName(page, id);
    if (name) fontNames.set(id, name);
  }
  return { fontNames, rules, fills };
}

function collectPath(
  args: unknown[] | undefined,
  ops: OpCodes,
  lineWidth: number,
  strokeColor: string,
  fillColor: string,
  rules: PdfRule[],
  fills: PdfFilledRect[],
): void {
  const paint = Number(args?.[0]);
  const pathList = arrayLike(args?.[1]);
  if (!pathList) return;
  const paintsStroke = paint === ops.stroke || paint === ops.closeStroke
    || paint === ops.fillStroke || paint === ops.eoFillStroke
    || paint === ops.closeFillStroke || paint === ops.closeEOFillStroke;
  const paintsFill = paint === ops.fill || paint === ops.eoFill
    || paint === ops.fillStroke || paint === ops.eoFillStroke
    || paint === ops.closeFillStroke || paint === ops.closeEOFillStroke;
  if (!paintsStroke && !paintsFill) return;

  for (const raw of pathList) {
    // pdf.js hands the segment list back as a plain object keyed `0,1,2,…`
    // rather than as a real array in the legacy build.
    const path = toNumberArray(raw);
    if (!path) continue;
    if (paintsStroke) {
      for (const segment of pathSegments(path)) {
        const dx = Math.abs(segment.x1 - segment.x0);
        const dy = Math.abs(segment.y1 - segment.y0);
        if (dy <= RULE_TOLERANCE && dx >= MIN_RULE_LENGTH) {
          rules.push({
            page: 0,
            orientation: 'horizontal',
            position: (segment.y0 + segment.y1) / 2,
            start: Math.min(segment.x0, segment.x1),
            end: Math.max(segment.x0, segment.x1),
            width: lineWidth,
            color: strokeColor,
          });
        } else if (dx <= RULE_TOLERANCE && dy >= MIN_RULE_LENGTH) {
          rules.push({
            page: 0,
            orientation: 'vertical',
            position: (segment.x0 + segment.x1) / 2,
            start: Math.min(segment.y0, segment.y1),
            end: Math.max(segment.y0, segment.y1),
            width: lineWidth,
            color: strokeColor,
          });
        }
      }
    }

    if (!paintsFill) continue;
    const box = pathBounds(path);
    if (!box) continue;
    const width = box.x1 - box.x0;
    const height = box.y1 - box.y0;
    if (height <= THIN_FILL && width >= MIN_RULE_LENGTH) {
      rules.push({
        page: 0,
        orientation: 'horizontal',
        position: (box.y0 + box.y1) / 2,
        start: box.x0,
        end: box.x1,
        width: Math.max(height, 0.3),
        color: fillColor,
      });
    } else if (width <= THIN_FILL && height >= MIN_RULE_LENGTH) {
      rules.push({
        page: 0,
        orientation: 'vertical',
        position: (box.x0 + box.x1) / 2,
        start: box.y0,
        end: box.y1,
        width: Math.max(width, 0.3),
        color: fillColor,
      });
    } else if (height > THIN_FILL) {
      fills.push({ page: 0, x0: box.x0, y0: box.y0, x1: box.x1, y1: box.y1, color: fillColor });
    }
  }
}

/** pdf.js sometimes hands back a plain object keyed `0,1,2,…` for an array. */
function arrayLike(value: unknown): unknown[] | null {
  if (Array.isArray(value)) return value;
  if (!value || typeof value !== 'object') return null;
  const record = value as Record<string, unknown>;
  const keys = Object.keys(record)
    .filter(key => /^\d+$/.test(key))
    .sort((a, b) => Number(a) - Number(b));
  return keys.length > 0 ? keys.map(key => record[key]) : null;
}

function toNumberArray(value: unknown): number[] | null {
  const list = arrayLike(value);
  if (!list) return null;
  const numbers = list.map(Number);
  return numbers.every(Number.isFinite) ? numbers : null;
}

function pathSegments(path: number[]): Array<{ x0: number; y0: number; x1: number; y1: number }> {
  const segments: Array<{ x0: number; y0: number; x1: number; y1: number }> = [];
  let x = 0;
  let y = 0;
  let startX = 0;
  let startY = 0;
  for (let index = 0; index < path.length;) {
    const command = path[index];
    index += 1;
    if (command === 0) {
      x = path[index];
      y = path[index + 1];
      index += 2;
      startX = x;
      startY = y;
    } else if (command === 1) {
      const nextX = path[index];
      const nextY = path[index + 1];
      index += 2;
      segments.push({ x0: x, y0: y, x1: nextX, y1: nextY });
      x = nextX;
      y = nextY;
    } else if (command === 2) {
      index += 4;
      const nextX = path[index];
      const nextY = path[index + 1];
      index += 2;
      segments.push({ x0: x, y0: y, x1: nextX, y1: nextY });
      x = nextX;
      y = nextY;
    } else if (command === 3) {
      segments.push({ x0: x, y0: y, x1: startX, y1: startY });
      x = startX;
      y = startY;
    } else {
      break;
    }
  }
  return segments;
}

function pathBounds(path: number[]): { x0: number; y0: number; x1: number; y1: number } | null {
  const points: Array<[number, number]> = [];
  for (let index = 0; index < path.length;) {
    const command = path[index];
    index += 1;
    if (command === 0 || command === 1) {
      points.push([path[index], path[index + 1]]);
      index += 2;
    } else if (command === 2) {
      for (let controls = 0; controls < 3; controls += 1) {
        points.push([path[index], path[index + 1]]);
        index += 2;
      }
    } else if (command === 3) {
      continue;
    } else {
      return null;
    }
  }
  if (points.length === 0) return null;
  const xs = points.map(point => point[0]);
  const ys = points.map(point => point[1]);
  return { x0: Math.min(...xs), y0: Math.min(...ys), x1: Math.max(...xs), y1: Math.max(...ys) };
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

/**
 * The text a PDF prints outside the text column on every page.
 *
 * Those lines are excluded from the margin measurement; here they are read for
 * what they say. A head that differs between odd and even pages becomes the two
 * OOXML slots the renderer understands, a footer that is only a page number
 * becomes `pageNumber`, and the distance Word measures each from the page edge
 * is estimated from where the line sits.
 */
function inferRunningHeads(
  lines: PdfLine[],
  body: StyleCluster,
): Pick<PageSettings, 'headers' | 'footers' | 'headerDistance' | 'footerDistance'> {
  const marginal = lines.filter(line => isMarginalLine(line, body));
  const headers = marginal.filter(line => line.y > line.pageHeight * HEADER_BAND);
  const footers = marginal.filter(line => line.y < line.pageHeight * FOOTER_BAND);
  const inferred: Pick<PageSettings, 'headers' | 'footers' | 'headerDistance' | 'footerDistance'> = {};

  const headerSlots = runningHeadSlots(headers);
  if (headerSlots) inferred.headers = headerSlots;
  const footerSlots = runningHeadSlots(footers);
  if (footerSlots) inferred.footers = footerSlots;

  if (headers.length > 0) {
    inferred.headerDistance = Math.round(median(headers.map(
      line => Math.max(0, line.pageHeight - line.y - line.fontSize),
    )) * POINTS_TO_TWIPS);
  }
  if (footers.length > 0) {
    inferred.footerDistance = Math.round(median(footers.map(
      line => Math.max(0, line.y - line.fontSize * 0.25),
    )) * POINTS_TO_TWIPS);
  }
  return inferred;
}

function runningHeadSlots(lines: PdfLine[]): RunningHeadSlots | undefined {
  if (lines.length === 0) return undefined;
  const oddText = mostCommonText(lines.filter(line => line.page % 2 === 1));
  const evenText = mostCommonText(lines.filter(line => line.page % 2 === 0));

  if (!oddText && !evenText) {
    // Nothing but page numbers: the slot still has to draw the number.
    return lines.some(line => isPageNumberText(cleanRunningText(line.text)))
      ? { default: { pageNumber: true } }
      : undefined;
  }

  const slots: RunningHeadSlots = {};
  if (oddText) slots.default = { text: oddText };
  if (evenText && evenText !== oddText) slots.even = { text: evenText };
  return slots;
}

/** The text that repeats across the pages in this band, ignoring page numbers. */
function mostCommonText(lines: PdfLine[]): string | undefined {
  const counts = new Map<string, number>();
  for (const line of lines) {
    const text = cleanRunningText(line.text);
    if (!text || isPageNumberText(text)) continue;
    counts.set(text, (counts.get(text) ?? 0) + 1);
  }
  const best = [...counts.entries()].sort((a, b) => b[1] - a[1])[0];
  // One occurrence is enough: a short sample may carry the even-page head on a
  // single sampled page, and a head repeated on two pages still wins by count.
  return best && best[0].length >= 2 ? best[0] : undefined;
}

function cleanRunningText(text: string): string {
  return text.replace(/\s+/g, ' ').trim();
}

/** `12`, `IV`, `iv` — a page number rather than a running head. */
function isPageNumberText(text: string): boolean {
  return /^(?:\d{1,4}|[ivxlcdm]{1,7})$/i.test(text);
}

/** The `w:sz` values OOXML accepts, in eighths of a point. */
const BORDER_SIZES = [2, 4, 6, 8, 12, 18, 24];

/**
 * The table the PDF's pages draw, read off its rules.
 *
 * A Chinese thesis draws three rules — top, under the header, bottom — with no
 * verticals; a journal draws a full grid. The rules are grouped into rows and
 * rows into a candidate table; the strongest candidate becomes the template's
 * `tableFormat`, so rule weights carry into `w:sz` and a shaded header band
 * becomes `headerFill`. One shape is kept because a template repeats it.
 */
function inferTableFormat(rules: PdfRule[], fills: PdfFilledRect[]): TableFormat | undefined {
  // The 篇眉 rule under a running head is horizontal and wide, but it is not a
  // table rule: left in, it merged with a three-line table as a fourth row.
  const contentRules = rules.filter(rule => !inHeaderFooterBand(rule));
  const horizontals = contentRules.filter(rule => rule.orientation === 'horizontal');
  const verticals = contentRules.filter(rule => rule.orientation === 'vertical');
  if (horizontals.length < 2) return undefined;

  const rows = mergeRuleRows(horizontals);
  // A border can be drawn per cell, and Word also emits hairline rules the eye
  // does not read as part of the table (0.24 pt beside a 1.44 pt border). Rules
  // much thinner than the table's strongest one are dropped before grouping, so
  // a three-line table is not mistaken for a fully ruled one.
  const strongest = Math.max(...rows.map(rule => rule.width));
  const significant = rows.filter(rule => rule.width >= Math.max(0.2, strongest * 0.25));
  const groups = groupRuleRows(significant);

  let best: { rows: PdfRule[]; x0: number; x1: number; score: number } | undefined;
  for (const group of groups) {
    if (group.length < 2) continue;
    const x0 = Math.min(...group.map(rule => rule.start));
    const x1 = Math.max(...group.map(rule => rule.end));
    const width = x1 - x0;
    if (width < 80) continue;
    const score = width * (Math.abs(group[0].position - group[group.length - 1].position) + 20);
    if (!best || score > best.score) best = { rows: group, x0, x1, score };
  }
  if (!best) return undefined;

  const ordered = [...best.rows].sort((a, b) => b.position - a.position);
  const top = ordered[0];
  const bottom = ordered[ordered.length - 1];
  const borders: NonNullable<TableFormat['borders']> = {
    top: ruleBorder(top),
    bottom: ruleBorder(bottom),
  };

  const gridColumns = verticals.filter(column =>
    column.start <= bottom.position + 2 && column.end >= top.position - 2
  );
  if (gridColumns.length >= 2) borders.insideV = ruleBorder(gridColumns[0]);
  // An interior rule beyond the header one means the table draws every row.
  if (ordered.length > 3) borders.insideH = ruleBorder(ordered[1]);

  const format: TableFormat = { borders };
  const headerRule = ordered.length >= 2 ? ordered[1] : undefined;
  if (headerRule) {
    const fill = fills.find(rect =>
      rect.x0 <= best.x0 + 6 && rect.x1 >= best.x1 - 6
      && rect.y0 <= headerRule.position + 3 && rect.y1 >= top.position - 3
      && !isWhiteColor(rect.color)
    );
    if (fill) format.headerFill = normalizeColor(fill.color);
  }
  return format;
}

/** Collapse rules that share a row position — a stroke and a fill, say. */
function mergeRuleRows(rules: PdfRule[]): PdfRule[] {
  const sorted = [...rules].sort((a, b) => b.position - a.position);
  const rows: PdfRule[] = [];
  for (const rule of sorted) {
    const current = rows[rows.length - 1];
    if (current && Math.abs(current.position - rule.position) <= 2) {
      current.position = (current.position + rule.position) / 2;
      current.start = Math.min(current.start, rule.start);
      current.end = Math.max(current.end, rule.end);
      current.width = Math.max(current.width, rule.width);
      continue;
    }
    rows.push({ ...rule });
  }
  return rows;
}

/** Group rows that overlap horizontally and sit close enough to be one table. */
function groupRuleRows(rows: PdfRule[]): PdfRule[][] {
  const groups: PdfRule[][] = [];
  for (const row of rows) {
    const group = groups[groups.length - 1];
    if (group) {
      const previous = group[group.length - 1];
      const gap = Math.abs(previous.position - row.position);
      const overlap = Math.min(previous.end, row.end) - Math.max(previous.start, row.start);
      const minWidth = Math.min(previous.end - previous.start, row.end - row.start);
      if (gap <= 260 && overlap >= minWidth * 0.5) {
        group.push(row);
        continue;
      }
    }
    groups.push([row]);
  }
  return groups;
}

function ruleBorder(rule: PdfRule): BorderSpec {
  return {
    style: 'single',
    size: ooxmlBorderSize(rule.width),
    color: normalizeColor(rule.color),
  };
}

function ooxmlBorderSize(points: number): number {
  const eighths = points * 8;
  let best = BORDER_SIZES[0];
  for (const size of BORDER_SIZES) {
    if (Math.abs(size - eighths) < Math.abs(best - eighths)) best = size;
  }
  return best;
}

/** `#RRGGBB` / `#RGB` → the `w:color` form, or black when the value is unknown. */
function normalizeColor(value: string): string {
  if (/^#[0-9a-f]{6}$/i.test(value)) return value.slice(1).toUpperCase();
  if (/^#[0-9a-f]{3}$/i.test(value)) {
    return value.slice(1).split('').map(char => char + char).join('').toUpperCase();
  }
  return '000000';
}

function isWhiteColor(value: string): boolean {
  return normalizeColor(value) === 'FFFFFF';
}

/** A rule printed in the page's running-head or footer band. */
function inHeaderFooterBand(rule: PdfRule): boolean {
  const height = rule.pageHeight ?? 0;
  if (height <= 0) return false;
  const margin = height * 0.08;
  if (rule.orientation === 'horizontal') {
    return rule.position > height - margin || rule.position < margin;
  }
  return rule.end > height - margin || rule.start < margin;
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
