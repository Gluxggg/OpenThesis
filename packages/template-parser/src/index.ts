// ============================================================
// @openthesis/template-parser — DOCX Template Style Extractor
// ============================================================
// Parses a .docx thesis template and extracts its formatting
// rules into a structured JSON DSL (DocumentTemplate).
// ============================================================

import JSZip from 'jszip';
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
  const pageSections = docXml
    ? extractPageSections(parser.parse(normalizeWordPrefix(docXml)), runningHeads, documentSettings)
    : [];
  const pageSettings = pageSections.length > 0
    ? pageSections[pageSections.length - 1]
    : getDefaultPageSettings();

  // 4. Map styles to semantic roles (heuristic detection), and record which
  // style wins each role. See `pickRoleWinners`.
  const usage = countStyleUsage(docXml);
  const { roles: styleRoles, explicit } = detectStyleRoles(resolvedStyles, Object.keys(rawStyles));
  const roleWinners = pickRoleWinners(styleRoles, usage, explicit);

  const evenAndOddHeaders = documentSettings.evenAndOddHeaders;

  // 6. Build template
  const meta: TemplateMeta = {
    ...metaOverrides,
    organization: metaOverrides?.organization?.trim() || 'Unknown Organization',
    name: metaOverrides?.name?.trim() || 'Untitled Template',
    documentType: metaOverrides?.documentType || 'thesis',
    parserVersion: '0.2.0',
    parsedAt: new Date().toISOString(),
  };

  // Convert to our style format
  const styles: Record<string, ParagraphStyle> = {};
  for (const [name, raw] of Object.entries(resolvedStyles)) {
    styles[name] = rawToParagraphStyle(raw);
  }

  const warnings = buildWarnings(rawStyles, resolvedStyles, pageSections, cycles);

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
  };
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
  indent?: { firstLine?: number; left?: number; right?: number; hanging?: number };
  spacing?: { before?: number; after?: number; line?: number; lineRule?: string };
  outlineLvl?: number;
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
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

  // Alignment
  const jc = pPr['w:jc']?.['@_w:val'];
  if (jc) props.alignment = jc;

  // Indentation. `w:firstLineChars` / `w:leftChars` / `w:rightChars` are
  // hundredths of a *character*, not twips, and Chinese templates emit them
  // constantly. Converting needs the run size: twips = chars × halfPoints / 10.
  const ind = pPr['w:ind'];
  if (ind) {
    props.indent = {};
    const effectiveSize = runSize && runSize > 0 ? runSize : 24; // 12pt fallback
    const charsToTwips = (chars: number) => Math.round((chars * effectiveSize) / 10);

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
    if (!isNaN(firstLineChars)) props.indent.firstLine = charsToTwips(firstLineChars);
    else if (!isNaN(firstLine)) props.indent.firstLine = firstLine;

    if (!isNaN(leftChars)) props.indent.left = charsToTwips(leftChars);
    else if (!isNaN(left)) props.indent.left = left;

    if (!isNaN(rightChars)) props.indent.right = charsToTwips(rightChars);
    else if (!isNaN(right)) props.indent.right = right;

    // A hanging indent is separate from a first-line indent and was dropped
    // entirely; the guide asks for it on every heading level.
    const hanging = parseInt(ind['@_w:hanging']);
    if (!isNaN(hanging)) props.indent.hanging = hanging;
  }

  // Spacing
  const spacing = pPr['w:spacing'];
  if (spacing) {
    props.spacing = {};
    const before = parseInt(spacing['@_w:before']);
    const after = parseInt(spacing['@_w:after']);
    const line = parseInt(spacing['@_w:line']);
    if (!isNaN(before)) props.spacing.before = before;
    if (!isNaN(after)) props.spacing.after = after;
    if (!isNaN(line)) {
      props.spacing.line = line;
      props.spacing.lineRule = spacing['@_w:lineRule'] || 'auto';
    }
  }

  // Outline level
  const outlineLvl = pPr['w:outlineLvl']?.['@_w:val'];
  if (outlineLvl !== undefined) props.outlineLvl = parseInt(outlineLvl);

  // Keep with next / keep lines
  if (pPr['w:keepNext']) props.keepNext = true;
  if (pPr['w:keepLines']) props.keepLines = true;
  if (pPr['w:pageBreakBefore']) props.pageBreakBefore = true;
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
  const resolved: Record<string, RawStyle> = {};
  const defaults = raw['_defaults'];

  for (const [id, style] of Object.entries(raw)) {
    if (id === '_defaults') continue;
    resolved[id] = resolveSingleStyle(id, style, raw, new Set(), defaults);
  }

  return resolved;
}

function resolveSingleStyle(
  id: string,
  style: RawStyle,
  all: Record<string, RawStyle>,
  visited: Set<string>,
  defaults?: RawStyle,
): RawStyle {
  // Prevent infinite loops
  if (visited.has(id)) return style;
  visited.add(id);

  // No inheritance
  if (!style.basedOn || !all[style.basedOn]) {
    // Apply defaults at the bottom of the chain
    if (defaults) {
      return mergeStyles(defaults, style);
    }
    return style;
  }

  // Recursively resolve parent
  const parent = resolveSingleStyle(style.basedOn, all[style.basedOn], all, visited, defaults);
  return mergeStyles(parent, style);
}

function mergeStyles(base: RawStyle, override: RawStyle): RawStyle {
  return {
    ...base,
    ...override,
    styleId: override.styleId,  // keep the override's ID
    name: override.name || base.name,
    pPr: mergeParagraphProps(base.pPr, override.pPr),
    rPr: { ...base.rPr, ...override.rPr },
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
  return {
    ...base,
    ...override,
    indent: (base.indent || override.indent)
      ? { ...base.indent, ...override.indent }
      : undefined,
    spacing: (base.spacing || override.spacing)
      ? { ...base.spacing, ...override.spacing }
      : undefined,
  };
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
      firstLineIndent: raw.pPr?.indent?.firstLine,
      hangingIndent: raw.pPr?.indent?.hanging,
      leftIndent: raw.pPr?.indent?.left,
      rightIndent: raw.pPr?.indent?.right,
      spaceBefore: raw.pPr?.spacing?.before,
      spaceAfter: raw.pPr?.spacing?.after,
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
  const index: RunningHeadPartIndex = new Map();

  const relsFile = zip.file('word/_rels/document.xml.rels');
  if (!relsFile) return index;
  const rels = parser.parse(await relsFile.async('string'));

  for (const relationship of asArray(rels?.['Relationships']?.['Relationship']) as any[]) {
    const type = String(relationship?.['@_Type'] ?? '');
    const isHeader = type.endsWith('/header');
    if (!isHeader && !type.endsWith('/footer')) continue;

    const id = relationship?.['@_Id'];
    const target = relationship?.['@_Target'];
    if (id === undefined || typeof target !== 'string') continue;

    const xml = await zip.file(resolvePartPath(target))?.async('string');
    // A relationship whose part is missing still occupies the slot: the template
    // asked for something here, and an empty head is the reading that cannot
    // invent text.
    index.set(String(id), xml ? readRunningHeadPart(parser, xml) : { text: '', pageNumber: false, rule: false });
  }

  return index;
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
  if (pgSz) {
    const w = parseInt(pgSz['@_w:w']);
    const h = parseInt(pgSz['@_w:h']);
    if (!isNaN(w)) width = w;
    if (!isNaN(h)) height = h;
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
  const pgNumType = firstElement(sectPr['w:pgNumType']);
  const pageNumberFormat = toPageNumberFormat(pgNumType?.['@_w:fmt']);
  const pageNumberStart = parseInt(pgNumType?.['@_w:start']);

  // 篇眉 / 页码 parts, as this section writes them. Danger: OOXML inherits each
  // slot independently, so what this section *prints* is resolved afterwards.
  const headers = readReferences(sectPr['w:headerReference'], runningHeads);
  const footers = readReferences(sectPr['w:footerReference'], runningHeads);

  return {
    width, height, margins, headerDistance, footerDistance, gutter, mirrorMargins, columns, columnGutter,
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
    { regex: /^(标题\s*1|Heading\s*1|Chapter|第.*章|h1|标题1|章标题|u?1级标题)$/i, role: 'heading1' },
    // Section titles
    { regex: /^(标题\s*2|Heading\s*2|Section|h2|标题2|节标题|u?2级标题)$/i, role: 'heading2' },
    // Subsection titles
    { regex: /^(标题\s*3|Heading\s*3|Subsection|h3|标题3|小节标题|u?3级标题)$/i, role: 'heading3' },
    { regex: /^(标题\s*4|Heading\s*4|h4|标题4|u?4级标题)$/i, role: 'heading4' },
    // Body text
    { regex: /^(正文|u正文|Normal|Body|body\s*text|正文文本|普通)$/i, role: 'paragraph' },
    // Figure and table captions are separate roles: the guide spaces a table
    // caption (段前1行) differently from a figure caption (段前0.1行).
    { regex: /^(u?图标题|图题|Figure\s*Caption)$/i, role: 'figure_caption' },
    { regex: /^(u?表标题|表题|Table\s*Caption)$/i, role: 'table_caption' },
    // Headings of the front/back matter (摘要, 目录, 序, 附录, 致谢, 参考文献). The
    // guide sets these apart from a chapter heading — 段后16.5磅 and 2.41倍行距
    // against 段后17磅 and 1.3倍 — and the USTB template has a style for them
    // (`u标题`, `u附录标题`) that used to be read as just another heading1.
    { regex: /^(u?标题|u?附录标题|.*标题\s*不入目录|摘要|目录|致谢|参考文献|序)$/i, role: 'section_heading' },
    // Reference entries under that heading.
    { regex: /^u?参考文献条目.*$/i, role: 'reference_item' },
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

/** How often each paragraph style is applied in `document.xml`. */
function countStyleUsage(docXml?: string): Map<string, number> {
  const usage = new Map<string, number>();
  if (!docXml) return usage;
  for (const match of docXml.matchAll(/<w:pStyle[^>]*w:val="([^"]*)"/g)) {
    usage.set(match[1], (usage.get(match[1]) ?? 0) + 1);
  }
  return usage;
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

export { extractRawStyles, resolveStyleInheritance, detectStyleRoles, countStyleUsage, pickRoleWinners };
