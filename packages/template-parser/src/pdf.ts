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

/** Infer a reusable OpenThesis style template from a text-based PDF. */
export async function parsePdfTemplate(
  buffer: Buffer | ArrayBuffer,
  metaOverrides?: Partial<TemplateMeta>,
): Promise<DocumentTemplate> {
  const { getDocument } = await import('pdfjs-dist/legacy/build/pdf.mjs');
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
      const content = await page.getTextContent();
      const items = content.items
        .filter(item => 'str' in item && Boolean(item.str.trim()))
        .map(item => item as unknown as PdfTextItem);
      lines.push(...groupItemsIntoLines(
        items,
        content.styles as Record<string, PdfTextStyle>,
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

function groupItemsIntoLines(
  items: PdfTextItem[],
  styles: Record<string, PdfTextStyle>,
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
    const fontFamily = normalizeFontFamily(styles[dominantFont]?.fontFamily || dominantFont);
    const text = joinTextItems(lineItems, fontSize);
    const x = Math.min(...lineItems.map(item => Number(item.transform[4] ?? 0)));
    const right = Math.max(...lineItems.map(item => Number(item.transform[4] ?? 0) + item.width));
    const fontDescriptor = `${fontFamily} ${dominantFont}`;
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
  if (/^sans-serif$/i.test(cleaned)) return 'Arial';
  if (/^serif$/i.test(cleaned)) return 'Times New Roman';
  if (/^monospace$/i.test(cleaned)) return 'Courier New';
  return cleaned && !/^g_d\d+_f\d+$/i.test(cleaned) ? cleaned : 'Times New Roman';
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
  const lineSpacing = cluster.id === body.id ? inferLineSpacing(cluster.lines) : undefined;
  return {
    font: {
      name: hasCjk ? 'Times New Roman' : cluster.fontFamily,
      eastAsia: hasCjk ? cluster.fontFamily : '宋体',
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
  const left = quantile(starts, 0.2);
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
  const numberedHeading = sample.match(/^(?:第.{1,12}[章节]|chapter\s+\d+|\d+(?:\.\d+){0,3}\s+\S)/im)?.[0];
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
  const leftPoints = quantile(bodyLines.map(line => line.x), 0.2);
  const rightPoints = quantile(bodyLines.map(line => line.pageWidth - line.right), 0.2);
  const pageGroups = new Map<number, PdfLine[]>();
  for (const line of lines) pageGroups.set(line.page, [...(pageGroups.get(line.page) ?? []), line]);
  const topPoints = median([...pageGroups.values()].map(pageLines => {
    const topLine = pageLines.reduce((top, line) => line.y > top.y ? line : top);
    return Math.max(0, topLine.pageHeight - topLine.y - topLine.fontSize);
  }));
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
