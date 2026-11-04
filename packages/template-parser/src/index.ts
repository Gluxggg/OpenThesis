// ============================================================
// @openthesis/template-parser — DOCX Template Style Extractor
// ============================================================
// Parses a .docx thesis template and extracts its formatting
// rules into a structured JSON DSL (DocumentTemplate).
// ============================================================

import JSZip from 'jszip';
import { readJournalRequirements } from './requirements.js';
import { XMLParser } from 'fast-xml-parser';
import type {
  DocumentTemplate,
  TemplateMeta,
  PageSettings,
  PageNumberFormat,
  ParagraphStyle,
  ParagraphFormatting,
  LineSpacingRule,
  RunningHeadSlots,
  StyleRole,
  BorderSpec,
  TableFormat,
  ImageFormat,
  TemplateNumbering,
  TemplateNumberingLevel,
  BlockType,
} from '@openthesis/document-schema';

const WORDPROCESSINGML_NAMESPACE = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/**
 * Rewrite the document's own WordprocessingML prefix to `w:`.
 *
 * Nothing in OOXML requires the prefix to be `w` — it is declared per document.
 * Every lookup in this file is written against the literal `w:`, so a
 * styles.xml binding the namespace to another prefix parsed to zero styles and
 * the template came out empty. The binding is declared on the root element, so
 * it is read from there and normalised before parsing.
 */
function normalizeWordPrefix(xml: string): string {
  const declaration = /xmlns:([\w.-]+)\s*=\s*["']([^"']*)["']/g;
  let prefix: string | undefined;
  for (let match = declaration.exec(xml); match; match = declaration.exec(xml)) {
    if (match[2] === WORDPROCESSINGML_NAMESPACE) { prefix = match[1]; break; }
  }
  if (!prefix || prefix === 'w') return xml;

  const escaped = prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return xml
    .replace(new RegExp(`<${escaped}:`, 'g'), '<w:')
    .replace(new RegExp(`</${escaped}:`, 'g'), '</w:')
    .replace(new RegExp(`(\\s)${escaped}:`, 'g'), '$1w:');
}

// Create XML parser with namespace-aware settings
function createParser(): XMLParser {
  return new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    isArray: (name: string) =>
      ['w:p', 'w:r', 'w:t', 'w:tab', 'w:pPr', 'w:rPr', 'w:style',
       'w:tbl', 'w:tr', 'w:tc', 'w:sectPr'].includes(name),
    removeNSPrefix: false,
  });
}

// ── Main Entry Point ──────────────────────────────────────

/**
 * Parse a .docx file and extract the thesis template definition.
 * @param buffer — Raw .docx file content (as Buffer)
 * @param meta — Template metadata (university, degree, etc.)
 * @returns DocumentTemplate — structured style DSL
 */
export async function parseTemplate(
  buffer: Buffer | ArrayBuffer,
  metaOverrides?: Partial<TemplateMeta>,
): Promise<DocumentTemplate> {
  const zip = await JSZip.loadAsync(buffer);
  const parser = createParser();

  // 1. Parse styles.xml
  const stylesXml = await zip.file('word/styles.xml')?.async('string');
  if (!stylesXml) throw new Error('No styles.xml found in template — is this a valid .docx?');

  const stylesParsed = parser.parse(normalizeWordPrefix(stylesXml));
  const rawStyles = extractRawStyles(stylesParsed);

  // 2. Resolve style inheritance chain. A `basedOn` loop is malformed; break it
  // first so the merged result does not depend on where the walk started.
  const { styles: acyclicStyles, cycles } = breakStyleCycles(rawStyles);
  const resolvedStyles = resolveStyleInheritance(acyclicStyles);

  // 3. Parse document.xml for page settings. Every section's running heads come
  // from the header/footer parts it references, so those are read first and
  // resolved in: without them a template's 篇眉 is invisible to the renderer,
  // which then invents its own. 对称页边距 and 偶数页页眉 come from settings.xml.
  const docXml = await zip.file('word/document.xml')?.async('string');
  const documentSettings = await readDocumentSettings(zip, parser);
  const runningHeads = await readRunningHeadParts(zip, parser);
  // Normalise the prefix once and share the tree: a document binds the
  // WordprocessingML namespace to whatever prefix it likes, and the usage
  // counter used to scan the *raw* string for `<w:pStyle`, which found nothing
  // in a document bound to another prefix. Both readers now see one tree.
  const parsedDocument = docXml ? parser.parse(normalizeWordPrefix(docXml)) : undefined;
  const format = readTableAndImageFormat(parsedDocument);

  const pageSections = parsedDocument
    ? extractPageSections(parsedDocument, runningHeads, documentSettings)
    : [];
  const pageSettings = pageSections.length > 0
    ? pageSections[pageSections.length - 1]
    : getDefaultPageSettings();

  // 4. Map styles to semantic roles (heuristic detection), and record which
  // style wins each role. See `pickRoleWinners`.
  const usage = countStyleUsageFromTree(parsedDocument);
  const { roles: styleRoles, explicit } = detectStyleRoles(resolvedStyles, Object.keys(rawStyles));
  const roleWinners = pickRoleWinners(styleRoles, usage, explicit);

  const evenAndOddHeaders = documentSettings.evenAndOddHeaders;

  // 6. Build template
  const meta: TemplateMeta = {
    ...metaOverrides,
    organization: metaOverrides?.organization?.trim() || 'Unknown Organization',
    name: metaOverrides?.name?.trim() || 'Untitled Template',
    documentType: metaOverrides?.documentType || 'thesis',
    sourceFormat: 'docx',
    parserVersion: '0.2.0',
    parsedAt: new Date().toISOString(),
  };

  // Convert to our style format
  const styles: Record<string, ParagraphStyle> = {};
  for (const [name, raw] of Object.entries(resolvedStyles)) {
    styles[name] = rawToParagraphStyle(raw);
  }

  const warnings = buildWarnings(rawStyles, resolvedStyles, pageSections, cycles);

  const numbering = await readNumbering(zip, parser);

  // A journal template states its requirements in its own text (Title: [Max. 20
  // words], Keywords: [At least 3], the sections it expects). Nothing about those
  // is layout, so they are read separately and reported.
  const documentTexts = docXml ? documentParagraphTexts(docXml, parser) : [];
  const requirements = readJournalRequirements(documentTexts, meta.documentType);

  return {
    meta,
    page: pageSettings,
    ...(pageSections.length > 0 ? { pageSections } : {}),
    ...(evenAndOddHeaders ? { evenAndOddHeaders } : {}),
    styles,
    styleRoles,
    ...(Object.keys(roleWinners).length > 0 ? { roleWinners } : {}),
    styleInheritance: buildInheritanceMap(rawStyles),
    ...(warnings.length > 0 ? { warnings } : {}),
    ...(numbering.length > 0 ? { numbering } : {}),
    ...(format.tableFormat ? { tableFormat: format.tableFormat } : {}),
    ...(format.imageFormat ? { imageFormat: format.imageFormat } : {}),
    ...(requirements ? { requirements } : {}),
  };
}

/**
 * The table and figure formatting a template actually declares.
 *
 * Templates disagree here far more than they do about styles: a Chinese
 * journal draws three rules at 0.5 pt, a university thesis draws 1.5 pt top and
 * bottom with a 0.75 pt header rule, an Elsevier template draws a full grid.
 * None of it was read before — the renderer hardcoded one shape — so every
 * template produced the same table.
 *
 * Only the *first* table and the *first* drawing are read. A template's later
 * tables are variations (a matrix, a parameter list) rather than the format a
 * document's own tables should inherit.
 */
function readTableAndImageFormat(
  document: any,
): { tableFormat?: TableFormat; imageFormat?: ImageFormat } {
  const result: { tableFormat?: TableFormat; imageFormat?: ImageFormat } = {};

  const body = firstElement(document?.['w:document']?.['w:body']);
  if (!body) return result;

  // ── Tables ──
  const firstTable = asArray(body['w:tbl'])[0];
  if (firstTable) {
    const tblPr = firstElement(firstTable['w:tblPr']);
    const format: TableFormat = {};

    const borders = firstElement(tblPr?.['w:tblBorders']);
    if (borders) {
      const spec = (value: any): BorderSpec | undefined => {
        const element = firstElement(value);
        if (!element) return undefined;
        const size = parseInt(element['@_w:sz']);
        return {
          style: String(element['@_w:val'] ?? 'single'),
          size: isNaN(size) ? 4 : size,
          color: String(element['@_w:color'] ?? 'auto'),
        };
      };
      format.borders = {
        top: spec(borders['w:top']),
        bottom: spec(borders['w:bottom']),
        left: spec(borders['w:left']),
        right: spec(borders['w:right']),
        insideH: spec(borders['w:insideH']),
        insideV: spec(borders['w:insideV']),
      };
    }

    const jc = tblPr?.['w:jc']?.['@_w:val'];
    if (jc === 'center' || jc === 'left' || jc === 'right') format.alignment = jc;
    const layout = tblPr?.['w:tblLayout']?.['@_w:type'];
    if (layout) format.layout = String(layout);

    const grid = firstElement(firstTable['w:tblGrid']);
    if (grid) {
      const widths = asArray(grid['w:gridCol'])
        .map(column => parseInt(column?.['@_w:w']))
        .filter((width): width is number => !isNaN(width));
      if (widths.length > 0) format.columnWidths = widths;
    }

    // A header cell's own shading and run size, from the first row.
    const firstRow = firstElement(firstTable['w:tr']);
    const firstCell = firstElement(firstRow?.['w:tc']);
    const cellPr = firstElement(firstCell?.['w:tcPr']);
    const fill = cellPr?.['w:shd']?.['@_w:fill'];
    if (fill && fill !== 'auto') format.headerFill = String(fill);
    const cellRun = firstElement(cellPr?.['w:rPr']) ?? firstElement(firstCell?.['w:p']?.['w:r']?.['w:rPr']);
    const size = parseInt(cellRun?.['w:sz']?.['@_w:val']);
    if (!isNaN(size)) format.headerSize = size;
    if (cellRun?.['w:b'] !== undefined) format.headerBold = flagOn(cellRun['w:b']);

    if (Object.keys(format).length > 0) result.tableFormat = format;
  }

  // ── Figures ──
  const drawing = findFirst(body, 'w:drawing');
  if (drawing) {
    const inline = findFirst(drawing, 'wp:inline');
    const anchor = findFirst(drawing, 'wp:anchor');
    const extent = findFirst(inline ?? anchor, 'wp:extent');
    const cx = parseInt(extent?.['@_cx']);
    const cy = parseInt(extent?.['@_cy']);
    const format: ImageFormat = {
      placement: anchor ? 'anchor' : 'inline',
      ...(isNaN(cx) ? {} : { widthEmu: cx }),
      ...(isNaN(cy) ? {} : { heightEmu: cy }),
    };
    result.imageFormat = format;
  }

  return result;
}

/** The first node stored under one tag, searched depth-first. */
function findFirst(node: any, tag: string): any {
  if (!node || typeof node !== 'object') return undefined;
  if (Array.isArray(node)) {
    for (const item of node) {
      const found = findFirst(item, tag);
      if (found !== undefined) return found;
    }
    return undefined;
  }
  for (const [key, value] of Object.entries(node)) {
    if (key === tag) return firstElement(value);
    const found = findFirst(value, tag);
    if (found !== undefined) return found;
  }
  return undefined;
}

/**
 * `word/numbering.xml`, resolved into the shapes a heading needs.
 *
 * A template numbers its chapters by binding styles to `w:num` instances
 * through `w:numPr`, and each instance points at an `w:abstractNum` whose
 * levels say what the number looks like (`%1`, `%1.%2`, `附录 %1`). Nothing
 * here was read before, so a chapter's number could only ever come from the
 * content file — and the guide's own numbering was invisible.
 *
 * Not read: `w:lvlOverride` on a `w:num`. A template that overrides a level
 * inline keeps the abstract definition instead.
 */
async function readNumbering(zip: JSZip, parser: XMLParser): Promise<TemplateNumbering[]> {
  const file = zip.file('word/numbering.xml');
  if (!file) return [];

  const parsed = parser.parse(normalizeWordPrefix(await file.async('string')));
  const root = parsed?.['w:numbering'];
  if (!root) return [];

  const abstracts = new Map<number, TemplateNumberingLevel[]>();
  for (const abstract of asArray(root['w:abstractNum']) as any[]) {
    const id = parseInt(abstract?.['@_w:abstractNumId']);
    if (isNaN(id)) continue;
    const levels: TemplateNumberingLevel[] = [];
    for (const lvl of asArray(abstract['w:lvl']) as any[]) {
      const ilvl = parseInt(lvl?.['@_w:ilvl']);
      if (isNaN(ilvl)) continue;
      const start = parseInt(lvl['w:start']?.['@_w:val']);
      levels.push({
        ilvl,
        format: lvl['w:numFmt']?.['@_w:val'],
        text: lvl['w:lvlText']?.['@_w:val'],
        start: isNaN(start) ? undefined : start,
        pStyle: lvl['w:pStyle']?.['@_w:val'],
        suff: lvl['w:suff']?.['@_w:val'],
      });
    }
    abstracts.set(id, levels.sort((a, b) => a.ilvl - b.ilvl));
  }

  const numbering: TemplateNumbering[] = [];
  for (const num of asArray(root['w:num']) as any[]) {
    const numId = parseInt(num?.['@_w:numId']);
    const abstractId = parseInt(num?.['w:abstractNumId']?.['@_w:val']);
    if (isNaN(numId) || isNaN(abstractId)) continue;
    const levels = abstracts.get(abstractId);
    if (!levels) continue;
    numbering.push({ numId, abstractId, levels });
  }
  return numbering;
}

/**
 * The text of every paragraph in `word/document.xml`, in document order.
 *
 * A journal template is an instruction sheet, so its body text *is* the spec.
 * Reading it is how the requirements get out.
 */
function documentParagraphTexts(xml: string, parser: XMLParser): string[] {
  const parsed = parser.parse(xml) as {
    'w:document'?: { 'w:body'?: { 'w:p'?: unknown } };
  };
  const body = parsed['w:document']?.['w:body'];
  if (!body) return [];
  const raw = body['w:p'];
  const paragraphs = Array.isArray(raw) ? raw : raw ? [raw] : [];
  return paragraphs
    .map(paragraphText)
    .filter((text): text is string => Boolean(text && text.trim()));
}

/** Every `w:t` inside one paragraph, joined. */
function paragraphText(paragraph: unknown): string {
  const texts: string[] = [];
  const collect = (node: unknown): void => {
    if (!node || typeof node !== 'object') return;
    const record = node as Record<string, unknown>;
    const text = record['w:t'];
    if (typeof text === 'string') {
      texts.push(text);
    } else if (Array.isArray(text)) {
      // fast-xml-parser emits repeated `w:t` elements as an array.
      for (const piece of text) {
        if (typeof piece === 'string') texts.push(piece);
        else if (piece && typeof piece === 'object' && '#text' in piece) {
          texts.push(String((piece as { '#text': unknown })['#text']));
        }
      }
    } else if (text && typeof text === 'object' && '#text' in text) {
      texts.push(String((text as { '#text': unknown })['#text']));
    }
    for (const [key, value] of Object.entries(record)) {
      if (key === 'w:t') continue;
      if (Array.isArray(value)) value.forEach(collect);
      else if (value && typeof value === 'object') collect(value);
    }
  };
  collect(paragraph);
  return texts.join('');
}

/**
 * Non-fatal problems worth telling the caller about.
 *
 * These exist because a template that carries no formatting used to parse
 * "successfully" into a DocumentTemplate whose every style was identical, and
 * the first symptom was a document with headings that looked like body text.
 */
function buildWarnings(
  rawStyles: Record<string, RawStyle>,
  resolvedStyles: Record<string, RawStyle>,
  pageSections: PageSettings[],
  cycles: string[] = [],
): string[] {
  const warnings: string[] = [];
  const styleIds = Object.keys(resolvedStyles);

  if (styleIds.length === 0) {
    warnings.push('styles.xml defines no paragraph styles; every block will use the built-in defaults.');
    return warnings;
  }

  const missingParents = new Set<string>();
  for (const style of Object.values(rawStyles)) {
    if (style.basedOn && !rawStyles[style.basedOn]) {
      missingParents.add(`${style.styleId} → ${style.basedOn}`);
    }
  }
  if (missingParents.size > 0) {
    const shown = [...missingParents].slice(0, 5).join(', ');
    const more = missingParents.size > 5 ? ` (+${missingParents.size - 5} more)` : '';
    warnings.push(`styles inherit from styles this template does not define: ${shown}${more}.`);
  }

  if (cycles.length > 0) {
    const shown = cycles.slice(0, 3).join('; ');
    const more = cycles.length > 3 ? ` (+${cycles.length - 3} more)` : '';
    warnings.push(
      `styles inherit from each other in a loop (${shown}${more}); inheritance stops at the first style `
      + 'in each loop, so the formatting may not match the source document.',
    );
  }

  const declaresFormatting = styleIds.some(id => {
    const style = resolvedStyles[id];
    return style.rPr !== undefined || style.pPr !== undefined;
  });
  if (!declaresFormatting) {
    warnings.push(
      'No style declares any font or paragraph formatting, so this template carries no typography; '
      + 'every block will fall back to the built-in defaults. Re-check the source .docx, or author the '
      + 'styles explicitly.',
    );
  }

  if (pageSections.length > 1) {
    warnings.push(
      `The document declares ${pageSections.length} sections; the generator reproduces the cover, the `
      + 'front matter and the body, so a section that only changes the setup mid-chapter is not '
      + 'reproduced.',
    );
  }

  return warnings;
}

// ── Style Extraction ──────────────────────────────────────

interface RawStyle {
  styleId: string;
  name: string;
  type: 'paragraph' | 'character' | 'table' | 'numbering';
  /** True for the `w:default="1"` style of its type (usually Normal/正文). */
  default?: boolean;
  basedOn?: string;
  /** Paragraph properties */
  pPr?: RawParagraphProps;
  /** Run (character) properties */
  rPr?: RawRunProps;
}

interface RawParagraphProps {
  alignment?: string;
  /**
   * Indent, in twips, alongside the raw character-unit counts Word writes next
   * to them.
   *
   * `w:firstLineChars` / `w:leftChars` / `w:rightChars` are hundredths of a
   * *character*, and a character is as wide as the style's **effective** font
   * size — the inherited one, not the one the style happens to declare. The
   * conversion therefore cannot happen while this bag is being read: the
   * inheritance chain is resolved afterwards. The raw counts are carried here
   * and converted in `rawToParagraphStyle`, where the merged run size is known.
   */
  indent?: {
    firstLine?: number; left?: number; right?: number; hanging?: number;
    firstLineChars?: number; leftChars?: number; rightChars?: number;
  };
  spacing?: {
    before?: number; after?: number; line?: number; lineRule?: string;
    /** `w:beforeLines` / `w:afterLines`: hundredths of a line, converted later. */
    beforeLines?: number; afterLines?: number;
  };
  outlineLvl?: number;
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  /** `w:numPr`: the numbering this style is bound to. `numId: 0` means "none". */
  numbering?: { numId: number; ilvl: number };
}

interface RawRunProps {
  fontName?: string;
  eastAsiaFont?: string;
  fontSize?: number;        // in half-points
  fontSizeCs?: number;      // complex script font size
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
}

function extractRawStyles(parsed: any): Record<string, RawStyle> {
  const styles: Record<string, RawStyle> = {};
  const styleElements = parsed?.['w:styles']?.['w:style'] || [];

  for (const el of styleElements) {
    const styleId = el['@_w:styleId'];
    if (!styleId) continue;

    const type = el['@_w:type'] || 'paragraph';
    const basedOn = el['w:basedOn']?.['@_w:val'];
    const pPrEl = firstElement(el['w:pPr']);

    // Chinese Word/WPS templates routinely put run properties (黑体, 三号, …)
    // inside `w:pPr/w:rPr` instead of a style-level `w:rPr`. Reading only the
    // direct child silently dropped them, which made every style look empty.
    // A style-level `w:rPr` wins when both are present.
    const nestedRPr = extractRunProps(pPrEl?.['w:rPr']);
    const directRPr = extractRunProps(el['w:rPr']);
    const rPr = mergeRunProps(nestedRPr, directRPr);

    const raw: RawStyle = {
      styleId,
      name: el['w:name']?.['@_w:val'] || styleId,
      type,
      default: el['@_w:default'] === '1' || el['@_w:default'] === 'true',
      basedOn,
      pPr: extractParagraphProps(el['w:pPr'], rPr?.fontSize),
      rPr,
    };

    styles[styleId] = raw;
  }

  // Also extract document defaults (DocDefaults)
  const docDefaults = parsed?.['w:styles']?.['w:docDefaults'];
  if (docDefaults) {
    const defaultRPr = extractRunProps(docDefaults['w:rPrDefault']?.['w:rPr']);
    const defaultPPr = extractParagraphProps(docDefaults['w:pPrDefault']?.['w:pPr'], defaultRPr?.fontSize);
    styles['_defaults'] = {
      styleId: '_defaults',
      name: 'Document Defaults',
      type: 'paragraph',
      pPr: defaultPPr,
      rPr: defaultRPr,
    };
  }

  return styles;
}

/** Merge two run-property bags; `override` wins. Keys are only present when set. */
function mergeRunProps(base?: RawRunProps, override?: RawRunProps): RawRunProps | undefined {
  if (!base && !override) return undefined;
  return { ...base, ...override };
}

function extractParagraphProps(pPr: any, runSize?: number): RawParagraphProps | undefined {
  pPr = firstElement(pPr);
  if (!pPr) return undefined;

  const props: RawParagraphProps = {};

  // `*Chars` (indent) and `*Lines` (spacing) are hundredths of a unit, not
  // `*Chars` (indent) and `*Lines` (spacing) are hundredths of a unit, not
  // twips, and they can only be converted once the style's **inherited** font
  // size is known — `runSize` is only what this one style declares. The raw
  // counts are stored here and converted in `rawToParagraphStyle`.
  void runSize;

  // Alignment
  const jc = pPr['w:jc']?.['@_w:val'];
  if (jc) props.alignment = jc;

  // Indentation. `w:firstLineChars` / `w:leftChars` / `w:rightChars` are
  // hundredths of a *character*, not twips, and Chinese templates emit them
  // constantly.
  const ind = pPr['w:ind'];
  if (ind) {
    props.indent = {};

    const firstLine = parseInt(ind['@_w:firstLine']);
    const firstLineChars = parseInt(ind['@_w:firstLineChars']);
    const left = parseInt(ind['@_w:left']);
    const leftChars = parseInt(ind['@_w:leftChars']);
    const right = parseInt(ind['@_w:right']);
    const rightChars = parseInt(ind['@_w:rightChars']);

    // `w:firstLineChars` wins when present: Word writes both attributes, and the
    // twips value beside it is derived from whatever size the paragraph had at
    // the time, so it goes stale — two characters at 12pt is 480 twips, not the
    // 200 Word left behind. `w:leftChars` / `w:rightChars` behave the same way.
    if (!isNaN(firstLineChars)) props.indent.firstLineChars = firstLineChars;
    else if (!isNaN(firstLine)) props.indent.firstLine = firstLine;

    if (!isNaN(leftChars)) props.indent.leftChars = leftChars;
    else if (!isNaN(left)) props.indent.left = left;

    if (!isNaN(rightChars)) props.indent.rightChars = rightChars;
    else if (!isNaN(right)) props.indent.right = right;

    // A hanging indent is separate from a first-line indent and was dropped
    // entirely; the guide asks for it on every heading level.
    const hanging = parseInt(ind['@_w:hanging']);
    if (!isNaN(hanging)) props.indent.hanging = hanging;
  }

  // Spacing. `w:beforeLines` / `w:afterLines` are hundredths of a line and are
  // what Chinese Word/WPS templates write for 段前 / 段后; reading only
  // `w:before` / `w:after` dropped them, so a style that declared 段前0.5行
  // parsed as declaring no spacing at all. As with the `*Chars` attributes, the
  // line-based value is the live one and wins over the twips beside it.
  const spacing = pPr['w:spacing'];
  if (spacing) {
    props.spacing = {};
    const before = parseInt(spacing['@_w:before']);
    const beforeLines = parseInt(spacing['@_w:beforeLines']);
    const after = parseInt(spacing['@_w:after']);
    const afterLines = parseInt(spacing['@_w:afterLines']);
    const line = parseInt(spacing['@_w:line']);
    if (!isNaN(beforeLines)) props.spacing.beforeLines = beforeLines;
    else if (!isNaN(before)) props.spacing.before = before;
    if (!isNaN(afterLines)) props.spacing.afterLines = afterLines;
    else if (!isNaN(after)) props.spacing.after = after;
    if (!isNaN(line)) {
      props.spacing.line = line;
      props.spacing.lineRule = spacing['@_w:lineRule'] || 'auto';
    }
  }

  // Outline level
  const outlineLvl = pPr['w:outlineLvl']?.['@_w:val'];
  if (outlineLvl !== undefined) props.outlineLvl = parseInt(outlineLvl);

  // Keep with next / keep lines / page break.
  //
  // All three are OOXML booleans, which Word writes as *empty* elements
  // (`<w:keepNext/>`). fast-xml-parser turns those into the empty string, which
  // is falsy — so `if (pPr['w:keepNext'])` never fired for a real Word file.
  // The shipped asset only carries `pageBreakBefore` because it was written into
  // Numbering binding. A template that numbers its chapters binds a style to a
  // `w:num` instance here; without it the number can only come from the content.
  const numPr = firstElement(pPr['w:numPr']);
  if (numPr) {
    const numId = parseInt(numPr['w:numId']?.['@_w:val']);
    const ilvl = parseInt(numPr['w:ilvl']?.['@_w:val']);
    if (!isNaN(numId)) {
      props.numbering = { numId, ilvl: isNaN(ilvl) ? 0 : ilvl };
    }
  }

  // the JSON by hand from the university guide, not because parsing produced it.
  if (flagOn(pPr['w:keepNext'])) props.keepNext = true;
  if (flagOn(pPr['w:keepLines'])) props.keepLines = true;
  if (flagOn(pPr['w:pageBreakBefore'])) props.pageBreakBefore = true;
  return Object.keys(props).length > 0 ? props : undefined;
}

function extractRunProps(rPr: any): RawRunProps | undefined {
  rPr = firstElement(rPr);
  if (!rPr) return undefined;

  const props: RawRunProps = {};

  // Font names
  const rFonts = rPr['w:rFonts'];
  if (rFonts) {
    if (rFonts['@_w:ascii']) props.fontName = rFonts['@_w:ascii'];
    if (rFonts['@_w:hAnsi']) props.fontName = props.fontName || rFonts['@_w:hAnsi'];
    if (rFonts['@_w:eastAsia']) props.eastAsiaFont = rFonts['@_w:eastAsia'];
  }

  // Font size (half-points)
  const sz = rPr['w:sz']?.['@_w:val'];
  if (sz) props.fontSize = parseInt(sz);
  const szCs = rPr['w:szCs']?.['@_w:val'];
  if (szCs) props.fontSizeCs = parseInt(szCs);

  // Bold
  if (rPr['w:b'] !== undefined) {
    const value = rPr['w:b']?.['@_w:val'];
    props.bold = value !== '0' && value !== 'false';
  }
  // Italic
  if (rPr['w:i'] !== undefined) {
    const value = rPr['w:i']?.['@_w:val'];
    props.italic = value !== '0' && value !== 'false';
  }
  // Underline
  if (rPr['w:u'] !== undefined) {
    const value = rPr['w:u']?.['@_w:val'];
    props.underline = value !== 'none' && value !== '0' && value !== 'false';
  }
  // Color
  if (rPr['w:color']) props.color = rPr['w:color']['@_w:val'];

  return Object.keys(props).length > 0 ? props : undefined;
}

function firstElement<T>(value: T | T[] | undefined): T | undefined {
  return Array.isArray(value) ? value[0] : value;
}

// ── Style Inheritance Resolution ──────────────────────────

/**
 * Cut every `basedOn` loop at a canonical point.
 *
 * A style whose chain returns to itself is malformed, and the merge used to
 * depend on which of its styles happened to be resolved first. Dropping the
 * link that closes the loop — always at the earliest-declared member — makes
 * the result the same whichever entry point the walk uses, and the loop is
 * reported through the template's warnings.
 */
function breakStyleCycles(
  styles: Record<string, RawStyle>,
): { styles: Record<string, RawStyle>; cycles: string[] } {
  const ids = Object.keys(styles);
  const cut = new Set<string>();
  const cycles: string[] = [];

  for (const start of ids) {
    const path: string[] = [];
    let current: string | undefined = start;
    while (current && styles[current] && !path.includes(current) && !cut.has(current)) {
      path.push(current);
      current = styles[current].basedOn;
    }
    if (!current || !path.includes(current)) continue;

    const loop = path.slice(path.indexOf(current));
    cut.add(loop.reduce((a, b) => (ids.indexOf(a) <= ids.indexOf(b) ? a : b)));
    cycles.push(`${loop.join(' → ')} → ${loop[0]}`);
  }

  if (cut.size === 0) return { styles, cycles };

  const broken: Record<string, RawStyle> = {};
  for (const [id, style] of Object.entries(styles)) {
    broken[id] = cut.has(id) ? { ...style, basedOn: undefined } : style;
  }
  return { styles: broken, cycles };
}

/**
 * Resolve the style inheritance chain.
 * In DOCX, styles can be basedOn other styles and only define
 * overrides. We recursively merge all inherited properties.
 */
function resolveStyleInheritance(raw: Record<string, RawStyle>): Record<string, RawStyle> {
  const defaults = raw['_defaults'];
  const resolved: Record<string, RawStyle> = {};

  // Memoised: a `basedOn` chain used to be walked from scratch for every style
  // that inherited it, so a template whose whole style table hangs off `Normal`
  // re-merged that chain n times — O(styles × depth) property-bag spreads. Each
  // style resolves once and its descendants reuse the result.
  const cache = new Map<string, RawStyle>();
  // Path-local, so a loop is cut rather than mistaken for a cache hit. Cycles are
  // already broken by `breakStyleCycles`; this is the belt-and-braces guard.
  const resolving = new Set<string>();

  const resolveOne = (id: string): RawStyle => {
    const cached = cache.get(id);
    if (cached) return cached;
    const style = raw[id];
    if (!style) return style;
    if (resolving.has(id)) return style;

    resolving.add(id);
    const parent = style.basedOn ? raw[style.basedOn] : undefined;
    const result = !style.basedOn || !parent
      ? (defaults ? mergeStyles(defaults, style) : style)
      : mergeStyles(resolveOne(style.basedOn), style);
    resolving.delete(id);

    cache.set(id, result);
    return result;
  };

  for (const id of Object.keys(raw)) {
    if (id === '_defaults') continue;
    resolved[id] = resolveOne(id);
  }

  return resolved;
}

function mergeStyles(base: RawStyle, override: RawStyle): RawStyle {
  return {
    ...base,
    ...override,
    styleId: override.styleId,  // keep the override's ID
    name: override.name || base.name,
    pPr: mergeParagraphProps(base.pPr, override.pPr),
    // Two absent run-property bags stay absent. Spreading them unconditionally
    // produced `rPr: {}` on every style, which made a template that declares no
    // typography look like one that declares some — Word writes an empty
    // `w:docDefaults/w:rPrDefault/w:rPr` into every file, so the "carries no
    // typography" warning could never fire.
    rPr: mergeRunProps(base.rPr, override.rPr),
  };
}

/** Deep-merge paragraph properties so nested indent/spacing are not lost */
function mergeParagraphProps(
  base?: RawParagraphProps,
  override?: RawParagraphProps,
): RawParagraphProps | undefined {
  if (!base && !override) return undefined;
  if (!base) return override;
  if (!override) return base;

  const indent = (base.indent || override.indent)
    ? { ...base.indent, ...override.indent }
    : undefined;
  const spacing = (base.spacing || override.spacing)
    ? { ...base.spacing, ...override.spacing }
    : undefined;

  // A measurement has two forms — twips, and hundredths of a character or a
  // line. A style that declares the twips form explicitly is overriding that
  // measurement, not only the twips half of it, so the inherited character
  // form must not survive to win the later conversion.
  dropOverriddenUnits(indent, override.indent, [
    ['firstLine', 'firstLineChars'],
    ['left', 'leftChars'],
    ['right', 'rightChars'],
  ]);
  dropOverriddenUnits(spacing, override.spacing, [
    ['before', 'beforeLines'],
    ['after', 'afterLines'],
  ]);

  return { ...base, ...override, indent, spacing };
}

/** Delete the unit-form key when the override declares the twips form instead. */
function dropOverriddenUnits(
  merged: Record<string, unknown> | undefined,
  override: Record<string, unknown> | undefined,
  pairs: readonly (readonly [string, string])[],
): void {
  if (!merged || !override) return;
  for (const [twipsKey, unitsKey] of pairs) {
    if (override[twipsKey] !== undefined && override[unitsKey] === undefined) {
      delete merged[unitsKey];
    }
  }
}

function buildInheritanceMap(raw: Record<string, RawStyle>): Record<string, string> {
  const map: Record<string, string> = {};
  for (const [id, style] of Object.entries(raw)) {
    if (style.basedOn) {
      map[id] = style.basedOn;
    }
  }
  return map;
}

// ── Style Conversion ──────────────────────────────────────

/** Drop keys whose value is `undefined` so the JSON stays honest and compact. */
function definedOnly<T extends object>(value: T): T {
  return Object.fromEntries(
    Object.entries(value).filter(([, v]) => v !== undefined),
  ) as T;
}

/**
 * Convert resolved OOXML style properties into the public ParagraphStyle shape.
 *
 * Only properties the template actually declares are emitted. Inventing
 * fallbacks here (the previous `|| 24` / `|| '宋体'` / `?? 0`) made a silent
 * template indistinguishable from one that explicitly asked for 12pt 宋体,
 * which is why every style in a formatting-free template looked identical and
 * headings ended up rendered as body text.
 */
function rawToParagraphStyle(raw: RawStyle): ParagraphStyle {
  // `*Chars` / `*Lines` are hundredths of a character or a line. A character
  // is as wide as the style's **effective** size — the inherited one — which is
  // only known here, after the `basedOn` chain and the docDefaults merge have
  // run. Using the style's own size produced 480 twips for a 10.5pt style that
  // inherits its size, where 420 is right.
  const effectiveSize = raw.rPr?.fontSize && raw.rPr.fontSize > 0 ? raw.rPr.fontSize : 24;
  const unitsToTwips = (units: number) => Math.round((units * effectiveSize) / 10);
  const indentFirstLine = raw.pPr?.indent?.firstLineChars !== undefined
    ? unitsToTwips(raw.pPr.indent.firstLineChars)
    : raw.pPr?.indent?.firstLine;
  const indentLeft = raw.pPr?.indent?.leftChars !== undefined
    ? unitsToTwips(raw.pPr.indent.leftChars)
    : raw.pPr?.indent?.left;
  const indentRight = raw.pPr?.indent?.rightChars !== undefined
    ? unitsToTwips(raw.pPr.indent.rightChars)
    : raw.pPr?.indent?.right;
  const spaceBefore = raw.pPr?.spacing?.beforeLines !== undefined
    ? unitsToTwips(raw.pPr.spacing.beforeLines)
    : raw.pPr?.spacing?.before;
  const spaceAfter = raw.pPr?.spacing?.afterLines !== undefined
    ? unitsToTwips(raw.pPr.spacing.afterLines)
    : raw.pPr?.spacing?.after;

  return definedOnly({
    font: definedOnly({
      name: raw.rPr?.fontName,
      eastAsia: raw.rPr?.eastAsiaFont,
      size: raw.rPr?.fontSize,
      bold: raw.rPr?.bold,
      italic: raw.rPr?.italic,
      underline: raw.rPr?.underline,
      color: raw.rPr?.color,
    }),
    paragraph: definedOnly({
      alignment: mapAlignment(raw.pPr?.alignment),
      // A chapter starts on a new page; the flag was parsed and then dropped
      // here, so every chapter ran on from the previous one.
      pageBreakBefore: raw.pPr?.pageBreakBefore,
      // `w:keepNext` / `w:keepLines` were parsed and merged into every resolved
      // style, then dropped here because the schema had no field for them.
      keepNext: raw.pPr?.keepNext,
      keepLines: raw.pPr?.keepLines,
      numbering: raw.pPr?.numbering,
      firstLineIndent: indentFirstLine,
      hangingIndent: raw.pPr?.indent?.hanging,
      leftIndent: indentLeft,
      rightIndent: indentRight,
      spaceBefore,
      spaceAfter,
      outlineLevel: raw.pPr?.outlineLvl,
    }),
    lineSpacing: raw.pPr?.spacing?.line,
    lineSpacingRule: toLineSpacingRule(raw.pPr?.spacing?.lineRule),
  });
}

/**
 * OOXML's `w:lineRule`: `auto` measures the line in 240ths, `exact` and
 * `atLeast` in twips. A value outside that vocabulary is dropped rather than
 * guessed at.
 */
function toLineSpacingRule(value?: string): LineSpacingRule | undefined {
  return value === 'auto' || value === 'exact' || value === 'atLeast' ? value : undefined;
}

function mapAlignment(align?: string): ParagraphFormatting['alignment'] | undefined {
  switch (align) {
    case 'left': return 'left';
    case 'center': return 'center';
    case 'right': return 'right';
    case 'both': return 'justified';
    case 'distribute': return 'distribute';
    default: return undefined;
  }
}

// ── Page Settings Extraction ───────────────────────────────

/**
 * What one header or footer part prints, keyed by the relationship id that
 * points at it.
 */
interface RunningHeadPart {
  text: string;
  pageNumber: boolean;
  rule: boolean;
}

type RunningHeadPartIndex = Map<string, RunningHeadPart>;

/**
 * Read every header/footer part `document.xml` relates to, so a section's
 * `w:headerReference` / `w:footerReference` can be resolved to what it prints.
 *
 * The parts are the only place the 篇眉 lives: the university template's odd
 * pages read 北京科技大学硕士学位论文 and its even pages carry the thesis title,
 * both centred 五号 with a 0.5 pt rule. Until this ran, the renderer invented a
 * head of its own and every template looked the same.
 *
 * `w:type="even"` is inert unless `word/settings.xml` sets
 * `w:evenAndOddHeaders`, which is read alongside this.
 */
async function readRunningHeadParts(zip: JSZip, parser: XMLParser): Promise<RunningHeadPartIndex> {
  const relsFile = zip.file('word/_rels/document.xml.rels');
  if (!relsFile) return new Map();
  const rels = parser.parse(await relsFile.async('string'));

  // Read every part concurrently. The reads used to be awaited one at a time
  // inside a serial loop; a template with 11 sections declares roughly 10–20
  // header/footer parts, and those round trips were most of `thesis parse`'s
  // wall time.
  const targets = (asArray(rels?.['Relationships']?.['Relationship']) as any[])
    .map(relationship => {
      const type = String(relationship?.['@_Type'] ?? '');
      const isHeadOrFoot = type.endsWith('/header') || type.endsWith('/footer');
      const id = relationship?.['@_Id'];
      const target = relationship?.['@_Target'];
      if (!isHeadOrFoot || id === undefined || typeof target !== 'string') return undefined;
      return { id: String(id), part: resolvePartPath(target) };
    })
    .filter((entry): entry is { id: string; part: string } => entry !== undefined);

  const parts = await Promise.all(
    targets.map(async ({ id, part }) => {
      const xml = await zip.file(part)?.async('string');
      // A relationship whose part is missing still occupies the slot: the
      // template asked for something here, and an empty head is the reading
      // that cannot invent text.
      return [id, xml ? readRunningHeadPart(parser, xml) : { text: '', pageNumber: false, rule: false }] as const;
    }),
  );

  return new Map(parts);
}

/** A relationship target is relative to `word/`; a leading `/` makes it absolute. */
function resolvePartPath(target: string): string {
  return target.startsWith('/') ? target.slice(1) : `word/${target.replace(/^\.\//, '')}`;
}

/**
 * What one header/footer part prints.
 *
 * The text is every `w:t` in the part, joined: a running head is one line, and
 * joining is the only reading that survives Word splitting it across runs. A
 * PAGE field marks the part as the page number, `w:pBdr/w:bottom` as the 篇眉
 * rule. An *empty* part is a suppression rather than an absence — the
 * university template keeps the cover head-less by pointing at one, since it
 * has no `w:titlePg` anywhere.
 */
function readRunningHeadPart(parser: XMLParser, xml: string): RunningHeadPart {
  const parsed = parser.parse(normalizeWordPrefix(xml));
  return {
    text: collectTagText(parsed, 'w:t').join(''),
    pageNumber: collectTagText(parsed, 'w:instrText').some(instruction => /\bPAGE\b/.test(instruction)),
    rule: hasBottomRule(parsed),
  };
}

/** Every value of one tag in a parsed part, in document order. */
function collectTagText(node: unknown, tag: string, out: string[] = []): string[] {
  if (!node || typeof node !== 'object') return out;
  if (Array.isArray(node)) {
    for (const item of node) collectTagText(item, tag, out);
    return out;
  }
  for (const [key, value] of Object.entries(node as Record<string, unknown>)) {
    if (key === tag) {
      const text = tagText(value);
      if (text) out.push(text);
    } else {
      collectTagText(value, tag, out);
    }
  }
  return out;
}

/** A leaf's character data, whether the parser kept it as a string or a `#text`. */
function tagText(value: unknown): string {
  const one = (item: unknown): string => {
    if (typeof item === 'string') return item;
    if (typeof item === 'number') return String(item);
    if (item && typeof item === 'object') {
      const text = (item as Record<string, unknown>)['#text'];
      if (typeof text === 'string' || typeof text === 'number') return String(text);
    }
    return '';
  };
  return Array.isArray(value) ? value.map(one).join('') : one(value);
}

/**
 * The 篇眉 rule: a `w:pBdr/w:bottom` that actually draws, i.e. one whose
 * `w:val` is not `none`. A paragraph that switches the border off explicitly
 * still carries the element.
 */
function hasBottomRule(node: unknown): boolean {
  if (!node || typeof node !== 'object') return false;
  if (Array.isArray(node)) return node.some(item => hasBottomRule(item));

  for (const [key, value] of Object.entries(node as Record<string, unknown>)) {
    if (key === 'w:pBdr') {
      for (const borders of asArray(value) as any[]) {
        const bottom = borders?.['w:bottom'];
        if (bottom === undefined) continue;
        const val = bottom?.['@_w:val'];
        if (val !== 'none' && val !== 'nil') return true;
      }
    } else if (hasBottomRule(value)) {
      return true;
    }
  }
  return false;
}

function asArray<T>(value: T | T[] | undefined): T[] {
  if (value === undefined || value === null) return [];
  return Array.isArray(value) ? value : [value];
}

/** Document-wide settings from `word/settings.xml` that the renderer must honour. */
interface DocumentSettings {
  /** `<w:evenAndOddHeaders/>`: odd and even pages may use different heads. */
  evenAndOddHeaders: boolean;
  /** `<w:mirrorMargins/>`: 对称页边距 for a bound thesis. */
  mirrorMargins: boolean;
}

/**
 * Read the document settings that change how pages are laid out.
 *
 * Both of these are `w:settings` children and are *invalid* anywhere else, so
 * this is the only place they can come from: the university template sets
 * 对称页边距 and 偶数页页眉 here, and looking for them inside `w:sectPr` — as the
 * page geometry reader does — reports `false` for a template that asks for
 * both. `<w:evenAndOddHeaders/>` is what makes a `w:type="even"` reference do
 * anything, and while it is on `w:type="default"` means odd pages only.
 */
async function readDocumentSettings(zip: JSZip, parser: XMLParser): Promise<DocumentSettings> {
  const settings: DocumentSettings = { evenAndOddHeaders: false, mirrorMargins: false };

  const xml = await zip.file('word/settings.xml')?.async('string');
  if (!xml) return settings;

  const parsed = parser.parse(normalizeWordPrefix(xml));
  const root = parsed?.['w:settings'];
  settings.evenAndOddHeaders = root?.['w:evenAndOddHeaders'] === undefined
    ? false
    : isOn(root['w:evenAndOddHeaders']);
  settings.mirrorMargins = root?.['w:mirrorMargins'] === undefined
    ? false
    : isOn(root['w:mirrorMargins']);

  return settings;
}

/** An OOXML on/off element is on when present, unless `w:val` says otherwise. */
function isOn(element: any): boolean {
  const value = element?.['@_w:val'];
  if (value === undefined) return true;
  const text = String(value).toLowerCase();
  return text !== '0' && text !== 'false' && text !== 'off';
}

/**
 * The same test, for a value that may be *absent*.
 *
 * `isOn` answers "is this present element on?"; the callers above have to ask
 * "is this element there at all, and if so is it on?" — and `isOn(undefined)`
 * returns `true`, so it cannot be used unguarded. An empty element parses to
 * `''`, which is present and on but falsy: that is the case `if (element)` gets
 * wrong.
 */
function flagOn(element: unknown): boolean {
  if (element === undefined || element === null) return false;
  return isOn(element);
}

/**
 * Every section geometry in the document, in document order.
 *
 * In OOXML every section but the last keeps its properties in
 * `w:p/w:pPr/w:sectPr`; the final section's properties are a direct child of
 * `w:body`. Reading only the body-level element meant a template with a
 * distinct cover or a landscape appendix reported a single geometry.
 */
function extractPageSections(
  parsed: any,
  runningHeads: RunningHeadPartIndex = new Map(),
  documentSettings: DocumentSettings = { evenAndOddHeaders: false, mirrorMargins: false },
): PageSettings[] {
  const body = parsed?.['w:document']?.['w:body'];
  if (!body) return [];

  const sections: PageSettings[] = [];

  for (const paragraph of asArray(body['w:p'])) {
    const pPr = firstElement(paragraph?.['w:pPr']);
    const sectPr = firstElement(pPr?.['w:sectPr']);
    if (sectPr) sections.push(sectPrToPageSettings(sectPr, runningHeads, documentSettings));
  }

  const bodySectPr = firstElement(body['w:sectPr']);
  if (bodySectPr) sections.push(sectPrToPageSettings(bodySectPr, runningHeads, documentSettings));

  return resolveRunningHeadInheritance(sections);
}

/**
 * Carry each of the six slots forward, the way OOXML does.
 *
 * A section that declares no reference of its own still prints its
 * predecessor's head. The university template's body relies on exactly that:
 * it declares no `w:headerReference`, yet it prints 北京科技大学硕士学位论文 on odd
 * pages and the thesis title on even ones, inherited from sections 8 and 5.
 * Reading only the sections that declare references makes the body look
 * head-less.
 */
function resolveRunningHeadInheritance(sections: PageSettings[]): PageSettings[] {
  let carriedHeaders: RunningHeadSlots | undefined;
  let carriedFooters: RunningHeadSlots | undefined;

  for (const section of sections) {
    carriedHeaders = mergeSlots(carriedHeaders, section.headers);
    carriedFooters = mergeSlots(carriedFooters, section.footers);

    if (carriedHeaders) section.headers = carriedHeaders;
    else delete section.headers;
    if (carriedFooters) section.footers = carriedFooters;
    else delete section.footers;
  }

  return sections;
}

/**
 * One slot at a time: a declared reference replaces that slot and leaves the
 * other two as they were carried in.
 */
function mergeSlots(
  carried: RunningHeadSlots | undefined,
  declared: RunningHeadSlots | undefined,
): RunningHeadSlots | undefined {
  if (!declared) return carried;
  const merged: RunningHeadSlots = { ...carried, ...declared };
  return Object.keys(merged).length > 0 ? merged : undefined;
}

/** The slots one section declares itself, before inheritance. */
function readReferences(refs: unknown, runningHeads: RunningHeadPartIndex): RunningHeadSlots | undefined {
  const slots: RunningHeadSlots = {};

  for (const ref of asArray(refs) as any[]) {
    const type = ref?.['@_w:type'];
    const slot: keyof RunningHeadSlots = type === 'even' ? 'even' : type === 'first' ? 'first' : 'default';
    const part = runningHeads.get(String(ref?.['@_r:id']));
    slots[slot] = {
      text: part?.text ?? '',
      pageNumber: part?.pageNumber ?? false,
      ...(part?.rule ? { rule: true } : {}),
    };
  }

  return Object.keys(slots).length > 0 ? slots : undefined;
}

const PAGE_NUMBER_FORMATS: PageNumberFormat[] = [
  'decimal', 'upperRoman', 'lowerRoman', 'upperLetter', 'lowerLetter',
  'chineseCounting', 'chineseCountingThousand', 'none',
];

/** `w:pgNumType/@w:fmt` in our vocabulary; an unknown format is dropped. */
function toPageNumberFormat(value?: string): PageNumberFormat | undefined {
  return PAGE_NUMBER_FORMATS.find(format => format === value);
}

function sectPrToPageSettings(
  sectPr: any,
  runningHeads: RunningHeadPartIndex = new Map(),
  documentSettings: DocumentSettings = { evenAndOddHeaders: false, mirrorMargins: false },
): PageSettings {
  const pgSz = sectPr['w:pgSz'];
  const pgMar = sectPr['w:pgMar'];
  // 对称页边距 is a document setting; a section that declares it anyway is
  // malformed but is still honoured rather than dropped.
  const mirrorMargins = documentSettings.mirrorMargins || sectPr['w:mirrorMargins'] !== undefined;

  let width = 11906;   // A4 default in twips
  let height = 16838;
  let orientation: 'portrait' | 'landscape' | undefined;
  if (pgSz) {
    const w = parseInt(pgSz['@_w:w']);
    const h = parseInt(pgSz['@_w:h']);
    if (!isNaN(w)) width = w;
    if (!isNaN(h)) height = h;
    const orient = pgSz['@_w:orient'];
    if (orient === 'portrait' || orient === 'landscape') orientation = orient;
  }

  const margins = {
    top: 1440,     // 1 inch default
    bottom: 1440,
    left: 1800,    // 1.25 inch default
    right: 1800,
  };
  if (pgMar) {
    const t = parseInt(pgMar['@_w:top']);
    const b = parseInt(pgMar['@_w:bottom']);
    const l = parseInt(pgMar['@_w:left']);
    const r = parseInt(pgMar['@_w:right']);
    if (!isNaN(t)) margins.top = t;
    if (!isNaN(b)) margins.bottom = b;
    if (!isNaN(l)) margins.left = l;
    if (!isNaN(r)) margins.right = r;
  }

  // Extract header/footer distances from pgMar attributes
  let headerDistance: number | undefined;
  let footerDistance: number | undefined;
  let gutter: number | undefined;
  if (pgMar) {
    const hd = parseInt(pgMar['@_w:header']);
    if (!isNaN(hd)) headerDistance = hd;
    const fd = parseInt(pgMar['@_w:footer']);
    if (!isNaN(fd)) footerDistance = fd;
    // The binding edge: the guide asks for 1 cm on the left.
    const gu = parseInt(pgMar['@_w:gutter']);
    if (!isNaN(gu)) gutter = gu;
  }

  const cols = sectPr['w:cols'];
  let columns = 1;
  let columnGutter: number | undefined;
  if (cols) {
    const num = parseInt(cols['@_w:num']);
    if (!isNaN(num)) columns = num;
    const space = parseInt(cols['@_w:space']);
    if (!isNaN(space)) columnGutter = space;
  }

  // 页码: the format this section numbers in and where it restarts. The
  // university template numbers the front matter `upperRoman` from I and
  // restarts the body at arabic 1 — which is a `w:pgNumType` on the section,
  // never on the footer that prints the number.
  // 文档网格: the line grid the template snaps text to. The writer library
  // always emits its own `linePitch="360"`, so the renderer rewrites this into
  // the generated section properties — without it a page breaks lines at
  // different places than the template does.
  const docGridElement = firstElement(sectPr["w:docGrid"]);
  const docGridPitch = docGridElement ? parseInt(docGridElement["@_w:linePitch"]) : NaN;
  const docGridCharSpace = docGridElement ? parseInt(docGridElement["@_w:charSpace"]) : NaN;
  const docGridType = docGridElement?.["@_w:type"] as string | undefined;
  const docGrid = docGridElement && (docGridType || !isNaN(docGridPitch) || !isNaN(docGridCharSpace))
    ? {
        ...(docGridType ? { type: docGridType } : {}),
        ...(!isNaN(docGridPitch) ? { linePitch: docGridPitch } : {}),
        ...(!isNaN(docGridCharSpace) ? { charSpace: docGridCharSpace } : {}),
      }
    : undefined;
  const pgNumType = firstElement(sectPr['w:pgNumType']);
  const pageNumberFormat = toPageNumberFormat(pgNumType?.['@_w:fmt']);
  const pageNumberStart = parseInt(pgNumType?.['@_w:start']);

  // 篇眉 / 页码 parts, as this section writes them. Danger: OOXML inherits each
  // slot independently, so what this section *prints* is resolved afterwards.
  const headers = readReferences(sectPr['w:headerReference'], runningHeads);
  const footers = readReferences(sectPr['w:footerReference'], runningHeads);

  return {
    width, height, margins, headerDistance, footerDistance, gutter, mirrorMargins, columns, columnGutter,
    ...(orientation ? { orientation } : {}),
    ...(docGrid && (docGrid.type || docGrid.linePitch !== undefined || docGrid.charSpace !== undefined) ? { docGrid } : {}),
    ...(pageNumberFormat ? { pageNumberFormat } : {}),
    ...(Number.isFinite(pageNumberStart) ? { pageNumberStart } : {}),
    ...(headers ? { headers } : {}),
    ...(footers ? { footers } : {}),
  };
}

function getDefaultPageSettings(): PageSettings {
  return {
    width: 11906,   // A4
    height: 16838,
    margins: { top: 1440, bottom: 1440, left: 1800, right: 1800 },
    columns: 1,
  };
}

// ── Semantic Role Detection ───────────────────────────────

/**
 * Heuristically detect the semantic role of each style.
 * This uses style name matching (Chinese + English) to guess
 * what each style is used for.
 *
 * In V4, this would be replaced by ML/layout-based detection.
 * For V1, we use heuristics based on common Chinese thesis patterns.
 */
function detectStyleRoles(
  styles: Record<string, RawStyle>,
  styleNames: string[],
): { roles: Record<string, StyleRole>; explicit: Set<string> } {
  const roles: Record<string, StyleRole> = {};
  // Styles a name pattern recognised, as opposed to ones guessed from their
  // outline level. A named match is the stronger signal of intent.
  const explicit = new Set<string>();

  // Heuristic patterns for common Chinese thesis style naming.
  // Patterns are anchored on purpose: the previous unanchored `/^(表|Table|题注)/`
  // also matched Word's built-in `TableGrid` ("Table Grid") and
  // "Table of Contents", turning table and TOC styles into level-3 headings.
  const patterns: Array<{ regex: RegExp; role: StyleRole }> = [
    // Chapter titles. `u1级标题` is how Chinese Word templates name the author's
    // own level-1 style; a bare `u标题` deliberately does not match, because the
    // USTB template uses it for front-matter headings (摘要, ABSTRACT, 序) with
    // different spacing from a chapter title.
    { regex: /^(标题\s*1|Heading\s*1|Chapter|第.*章|h1|标题1|章标题|u?1级标题|一级标题|Head1|Head\s*1|MainHeading)$/i, role: 'heading1' },
    // Section titles
    { regex: /^(标题\s*2|Heading\s*2|Section|h2|标题2|节标题|u?2级标题|二级标题|Head2|Head\s*2)$/i, role: 'heading2' },
    // Subsection titles
    { regex: /^(标题\s*3|Heading\s*3|Subsection|h3|标题3|小节标题|u?3级标题|三级标题|Head3|Head\s*3)$/i, role: 'heading3' },
    { regex: /^(标题\s*4|Heading\s*4|h4|标题4|u?4级标题)$/i, role: 'heading4' },
{ regex: /^(标题\s*5|Heading\s*5|h5|标题5|u?5级标题)$/i, role: 'heading5' },
{ regex: /^(标题\s*6|Heading\s*6|h6|标题6|u?6级标题)$/i, role: 'heading6' },
    // Body text
    // Body text. The stock names come first; the rest were added after parsing
    // real templates, where a university calls its body style `硕士正文样式` and an
    // Elsevier template calls it `Standard`.
    {
      regex: /^(正文|u正文|Normal|Body|body\s*text|正文文本|普通|硕士正文样式|博士正文样式|学位论文正文|Standard|Body\s*Text\s*Indent|正文样式|Body\s*Text|First\s*Paragraph)$/i,
      role: 'paragraph',
    },
    // Journal front matter, by the names Pandoc's reference document uses and
    // that a journal template built from one carries: Title, Subtitle, Author,
    // Date, Abstract. Without these a journal template recognised nothing but
    // `paragraph`, and the renderer fell back to its own defaults for the rest.
    { regex: /^(Title|论文题目|标题)$/i, role: 'journal_title' },
    { regex: /^(Subtitle|副标题)$/i, role: 'journal_subtitle' },
    { regex: /^(Author|Authors|作者)$/i, role: 'journal_author' },
    { regex: /^(Affiliation|Affiliations|作者单位|单位)$/i, role: 'journal_affiliation' },
    { regex: /^(Date|日期)$/i, role: 'journal_date' },
    { regex: /^(Abstract|摘要)$/i, role: 'journal_abstract' },
    { regex: /^(Keywords?|Key\s*words?|关键词)$/i, role: 'journal_keywords' },

    // Figure and table captions are separate roles: the guide spaces a table
    // caption (段前1行) differently from a figure caption (段前0.1行).
    { regex: /^(u?图标题|图题|Figure\s*Caption|Caption|Figure|Figure\s*With\s*Caption|Image\s*Caption)$/i, role: 'figure_caption' },
    { regex: /^(u?表标题|表题|Table\s*Caption)$/i, role: 'table_caption' },
    // Headings of the front/back matter (摘要, 目录, 序, 附录, 致谢, 参考文献). The
    // guide sets these apart from a chapter heading — 段后16.5磅 and 2.41倍行距
    // against 段后17磅 and 1.3倍 — and the USTB template has a style for them
    // (`u标题`, `u附录标题`) that used to be read as just another heading1.
    { regex: /^(u?标题|u?附录标题|.*标题\s*不入目录|摘要|目录|致谢|参考文献|序)$/i, role: 'section_heading' },
    // The 摘要 body and its 关键词 line. A template that carries these has its own
    // rule for them (the USTB guide sets the label 黑体 and the text 宋体), and a
    // template without them falls back to the body style.
    { regex: /^(摘要正文|中文摘要|Abstract摘要|u?摘要内容)$/i, role: 'thesis_abstract' },
    { regex: /^(英文摘要|Abstract\s*(body|text)?|English\s*Abstract)$/i, role: 'thesis_abstract_en' },
    { regex: /^(关键词|u?关键字|Keywords?)$/i, role: 'thesis_keywords' },
    { regex: /^(Key\s*words?|English\s*Keywords?)$/i, role: 'thesis_keywords_en' },
    // Reference entries under that heading.
    { regex: /^(u?参考文献条目.*|Bibliography)$/i, role: 'reference_item' },
    // Table-of-contents entries. Word writes them with these styles when the TOC
    // field is updated; they were being treated as body paragraphs.
    { regex: /^TOC\s*\d+$/i, role: 'toc_entry' },
    // Cover title
    { regex: /^(封面|Cover|Title|论文题目|题目)$/i, role: 'centered_text' },
    // Equation styles — MathType emits "MT Converted Equation", WPS emits "公式".
    // Without this, no template style ever drove `equation` and every formula
    // silently fell back to the built-in defaults.
    { regex: /^(公式|Equation|MT\s*Converted\s*Equation|Math\s*Equation)$/i, role: 'equation' },
  ];

  // Unmatched paragraph styles still act as body text, but they are appended
  // after every explicit match so that the renderer's first-match lookup picks
  // a style that actually declares a body-text role.
  const defaultParagraph: string[] = [];
  const otherParagraph: string[] = [];

  for (const name of styleNames) {
    const raw = styles[name];
    if (!raw) continue;

    // Character, table and numbering styles are not paragraph-level blocks.
    // Assigning them a block role made `resolveStyle` pick whichever style
    // happened to appear first in the XML — e.g. `uChar` (a character style)
    // or `TOC3` supplying the body-text formatting.
    if (raw.type !== 'paragraph') continue;

    let matched = false;
    for (const { regex, role } of patterns) {
      if (regex.test(name) || regex.test(raw.name)) {
        roles[name] = role;
        explicit.add(name);
        matched = true;
        break;
      }
    }
    if (matched) continue;

    // Fallback: use outline level
    if (raw.pPr?.outlineLvl !== undefined) {
      const lvl = raw.pPr.outlineLvl;
      if (lvl >= 0 && lvl <= 3) {
        roles[name] = `heading${lvl + 1}` as BlockType;
        continue;
      }
    }

    (raw.default ? defaultParagraph : otherParagraph).push(name);
  }

  for (const name of defaultParagraph) roles[name] = 'paragraph';
  for (const name of otherParagraph) roles[name] = 'paragraph';

  return { roles, explicit };
}

/**
 * How often each paragraph style is applied in a `document.xml`, given the
 * already-parsed part.
 *
 * Read from the tree, not from the raw XML. The prefix a document binds
 * WordprocessingML to is arbitrary, so a scan for the literal `<w:pStyle`
 * found zero uses in a document that binds another prefix — and every role then
 * tied at zero, with the winner decided by declaration order instead of by
 * what the document actually applies.
 */
function countStyleUsageFromTree(document?: unknown): Map<string, number> {
  const usage = new Map<string, number>();
  if (!document) return usage;

  for (const value of collectTagNodes(document, 'w:pStyle')) {
    for (const node of asArray(value as any) as any[]) {
      const id = node?.['@_w:val'];
      if (id === undefined || id === null) continue;
      const key = String(id);
      usage.set(key, (usage.get(key) ?? 0) + 1);
    }
  }
  return usage;
}

/** Every node stored under one tag in a parsed part, in document order. */
function collectTagNodes(node: unknown, tag: string, out: unknown[] = []): unknown[] {
  if (!node || typeof node !== 'object') return out;
  if (Array.isArray(node)) {
    for (const item of node) collectTagNodes(item, tag, out);
    return out;
  }
  for (const [key, value] of Object.entries(node as Record<string, unknown>)) {
    if (key === tag) out.push(value);
    else collectTagNodes(value, tag, out);
  }
  return out;
}

/**
 * Pick one style per role, preferring the styles the document actually uses.
 *
 * Word's stock `heading 2` / `heading 3` carry a heading role by name whether or
 * not the template uses them. In the USTB template they are unused: every
 * heading is set in the author's own `u2级标题` / `u3级标题` (黑体 四号, exactly
 * what the university's writing guide requires). The winners cannot simply be
 * written first in `styleRoles` either — a plain object enumerates integer-like
 * keys first, and Word names those stock styles `1`, `2`, `3`, so they win any
 * ordering that relies on key position. Hence a separate, explicit record.
 */
function pickRoleWinners(
  roles: Record<string, StyleRole>,
  usage: Map<string, number>,
  explicit: ReadonlySet<string> = new Set(),
): Partial<Record<StyleRole, string>> {
  interface Candidate { id: string; explicit: boolean; count: number }
  const best = new Map<StyleRole, Candidate>();

  for (const [id, role] of Object.entries(roles)) {
    const candidate: Candidate = { id, explicit: explicit.has(id), count: usage.get(id) ?? 0 };
    const current = best.get(role);
    // A style a name pattern recognised beats one guessed from its outline
    // level; between equals, the one the document applies more often wins.
    // Ties keep the first style seen, which is the order callers already had.
    const better = !current
      || (candidate.explicit && !current.explicit)
      || (candidate.explicit === current.explicit && candidate.count > current.count);
    if (better) best.set(role, candidate);
  }

  const winners: Partial<Record<StyleRole, string>> = {};
  for (const [role, { id }] of best) winners[role] = id;
  return winners;
}

// ── Utility Exports ───────────────────────────────────────

// Public surface: `parseTemplate` (defined above) and `parsePdfTemplate` (below).
// `extractRawStyles`, `resolveStyleInheritance`, `detectStyleRoles`,
// `countStyleUsageFromTree` and `pickRoleWinners` are implementation detail — they
// were exported, but their signatures name `RawStyle` / `RawParagraphProps`, which
// are not exported, so no external caller could even write the parameter type.
// Nothing in this repository imports them.
export { parsePdfTemplate } from './pdf.js';
