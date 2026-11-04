// ============================================================
// @openthesis/docx-renderer — Template-Driven DOCX Renderer
// ============================================================
// Converts ThesisDocument content into a standards-compliant
// .docx file, using a parsed DocumentTemplate for all formatting.
// ============================================================

import {
  Document, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, PageNumber, PageBreak,
  AlignmentType,
  BorderStyle, WidthType, ShadingType, LevelFormat, TableLayoutType,
  XmlComponent,
  convertMillimetersToTwip, Packer, LineRuleType, PageOrientation,
  ImageRun,
  Math as DocxMath, MathRun, MathFraction, MathRadical, MathSuperScript,
  MathSubScript, MathSubSuperScript,
  MathFunction, MathLimitLower, MathLimitUpper, BuilderElement,
  TableOfContents, TabStopType, LeaderType,
} from 'docx';
import type {
  ISectionOptions, ISectionPropertiesOptions, IPageNumberTypeAttributes,
  MathComponent, INumberingOptions, IParagraphStyleOptions,
} from 'docx';
import type {
  DocumentTemplate,
  ThesisDocument,
  ThesisSection,
  JournalArticle,
  OfficialDocument,
  ContentBlock,
  ParagraphStyle,
  PageSettings,
  PageNumberFormat,
  RunningHead,
  RunningHeadSlots,
  LineSpacingRule,
  StyleRole,
  LegacyDocumentJSON,
  SpacerBlock,
  BorderSpec,
  FigureItem,
  OpenThesisDocument,
  HeadingBlock,
  ParagraphBlock,
  CenteredTextBlock,
  EquationBlock,
  FigureBlock,
  TableBlock,
  ListItemBlock,
  CodeBlockBlock,
  BlockquoteBlock,
  RedHeaderBlock,
  DocumentNumberBlock,
  RecipientLineBlock,
  SignatureBlockBlock,
  AttachmentNoteBlock,
} from '@openthesis/document-schema';
import { writeFileSync, readFileSync, existsSync, mkdirSync } from 'fs';
import { resolve, extname, isAbsolute, dirname } from 'path';
import { fileURLToPath } from 'url';
import JSZip from 'jszip';
import { latexToMathAst, latexToPlainText } from '@openthesis/equation-engine';
import type { LatexMathNode } from '@openthesis/equation-engine';

// ESM-compatible __dirname
const _filename = fileURLToPath(import.meta.url);
const _dirname = dirname(_filename);

// ── Renderer Config ───────────────────────────────────────

export interface RenderOptions {
  /**
   * Path to write the output .docx.
   *
   * Optional: omitting it (or passing an empty string) renders the buffer and
   * writes nothing, which is how the tests and the library API use it. The type
   * used to declare it required while every caller treated it as optional and
   * the implementation guarded it with `if (options.outputPath)`.
   */
  outputPath?: string;
  /** Parsed template (from template-parser) */
  template: DocumentTemplate;
  /** Content to render */
  document: OpenThesisDocument;
  /** Header text (defaults to template.meta.organization + '学位论文') */
  headerText?: string;
  /** Show page numbers in footer */
  showPageNumbers?: boolean;
  /** Base directory of content file for relative paths (e.g. image paths) */
  contentDir?: string;
  /**
   * Restart each part's page numbering at 1 (default true).
   *
   * Turn it off only to work around a viewer. LibreOffice — and any viewer built
   * on its DOCX import, which may include WPS — drops a section's tables entirely
   * when that section declares `w:pgNumType w:start="1"` and is not the first
   * section. The attribute is what makes the body start at page 1, so it stays on
   * by default; this switch trades that numbering for tables that render there.
   */
  pageNumberRestart?: boolean;
}

// ── Fallback Styles ───────────────────────────────────────
//
// These are consulted only for properties a parsed template does not declare.
// The two domains genuinely differ — a 学位论文 body is 宋体 小四 with a 2-character
// first-line indent, while a 公文 body is 仿宋 三号 per GB/T 9704-2012 — so a single
// shared table made thesis output inherit official-document typography.

const THESIS_STYLES: Record<string, ParagraphStyle> = {
  heading1: {
    font: { name: 'Times New Roman', eastAsia: '黑体', size: 32, bold: true },
    paragraph: { alignment: 'center', spaceBefore: 240, spaceAfter: 240 },
    lineSpacing: 360,
  },
  heading2: {
    font: { name: 'Times New Roman', eastAsia: '黑体', size: 28, bold: true },
    paragraph: { alignment: 'left', spaceBefore: 240, spaceAfter: 120 },
    lineSpacing: 360,
  },
  heading3: {
    font: { name: 'Times New Roman', eastAsia: '黑体', size: 24, bold: true },
    paragraph: { alignment: 'left', spaceBefore: 180, spaceAfter: 120 },
    lineSpacing: 360,
  },
  heading4: {
    font: { name: 'Times New Roman', eastAsia: '黑体', size: 24, bold: true },
    paragraph: { alignment: 'left', spaceBefore: 120, spaceAfter: 60 },
    lineSpacing: 360,
  },
  paragraph: {
    font: { name: 'Times New Roman', eastAsia: '宋体', size: 24 },
    paragraph: { alignment: 'justified', firstLineIndent: 480 },
    lineSpacing: 360,
  },
  paragraph_no_indent: {
    font: { name: 'Times New Roman', eastAsia: '宋体', size: 24 },
    paragraph: { alignment: 'justified' },
    lineSpacing: 360,
  },
  centered_text: {
    font: { name: 'Times New Roman', eastAsia: '黑体', size: 30, bold: true },
    paragraph: { alignment: 'center' },
    lineSpacing: 360,
  },
  equation: {
    font: { name: 'Times New Roman', eastAsia: '宋体', size: 24 },
    paragraph: { alignment: 'center' },
    lineSpacing: 360,
  },
  table: {
    font: { name: 'Times New Roman', eastAsia: '宋体', size: 21 },
    paragraph: { alignment: 'center' },
    lineSpacing: 276,
  },
  table_header: {
    font: { name: 'Times New Roman', eastAsia: '黑体', size: 21, bold: true },
    paragraph: { alignment: 'center' },
    lineSpacing: 276,
  },
  // 表题 and 图题 are separate roles because the guide spaces them differently:
  // 段前1行 for a table caption against 段前0.1行 for a figure caption. The numbers
  // are the template's own (`ub` and `u4`), so a build that supplies no template
  // lands on the same twips the university template declares.
  table_caption: {
    font: { name: 'Times New Roman', eastAsia: '黑体', size: 21, bold: true },
    paragraph: { alignment: 'center', spaceBefore: 150, spaceAfter: 50 },
    lineSpacing: 312,
  },
  figure_caption: {
    font: { name: 'Times New Roman', eastAsia: '黑体', size: 21, bold: true },
    paragraph: { alignment: 'center', spaceBefore: 10, spaceAfter: 10 },
    lineSpacing: 312,
  },
  // Journal front matter. A journal article carries no template of its own in
  // most workflows, so these are the defaults a bare `--type journal` build gets.
  journal_title: {
    font: { name: 'Times New Roman', eastAsia: '黑体', size: 32, bold: true },
    paragraph: { alignment: 'center', spaceBefore: 240, spaceAfter: 120 },
    lineSpacing: 312,
  },
  journal_author: {
    font: { name: 'Times New Roman', eastAsia: '宋体', size: 21 },
    paragraph: { alignment: 'center', spaceBefore: 120, spaceAfter: 60 },
    lineSpacing: 312,
  },
  journal_affiliation: {
    font: { name: 'Times New Roman', eastAsia: '宋体', size: 18, italic: true },
    paragraph: { alignment: 'center', spaceBefore: 60, spaceAfter: 120 },
    lineSpacing: 312,
  },
  journal_abstract: {
    font: { name: 'Times New Roman', eastAsia: '宋体', size: 20 },
    paragraph: { alignment: 'justified', leftIndent: 720, rightIndent: 720, spaceBefore: 120, spaceAfter: 60 },
    lineSpacing: 240,
  },
  journal_keywords: {
    font: { name: 'Times New Roman', eastAsia: '宋体', size: 20 },
    paragraph: { alignment: 'justified', leftIndent: 720, rightIndent: 720, spaceBefore: 60, spaceAfter: 120 },
    lineSpacing: 240,
  },
};

const OFFICIAL_STYLES: Record<string, ParagraphStyle> = {
  // GB/T 9704-2012《党政机关公文格式》
  heading1: {
    // 一级标题用3号黑体（16pt）
    font: { name: 'Times New Roman', eastAsia: '黑体', size: 32, bold: true },
    paragraph: { alignment: 'left', spaceBefore: 340, spaceAfter: 340 },
    lineSpacing: 312,
  },
  heading2: {
    // 二级标题用3号楷体（16pt）
    font: { name: 'Times New Roman', eastAsia: '楷体', size: 32, bold: true },
    paragraph: { alignment: 'left', spaceBefore: 260, spaceAfter: 260 },
    lineSpacing: 312,
  },
  heading3: {
    font: { name: 'Times New Roman', eastAsia: '黑体', size: 28, bold: true },
    paragraph: { alignment: 'left', spaceBefore: 200, spaceAfter: 200 },
    lineSpacing: 312,
  },
  heading4: {
    font: { name: 'Times New Roman', eastAsia: '楷体', size: 28, bold: true },
    paragraph: { alignment: 'left', spaceBefore: 160, spaceAfter: 160 },
    lineSpacing: 312,
  },
  paragraph: {
    // 正文用3号仿宋体（16pt）
    font: { name: 'Times New Roman', eastAsia: '仿宋', size: 32 },
    paragraph: { alignment: 'justified', firstLineIndent: convertMillimetersToTwip(7.4) },
    lineSpacing: 312,
  },
  paragraph_no_indent: {
    font: { name: 'Times New Roman', eastAsia: '仿宋', size: 32 },
    paragraph: { alignment: 'left' },
    lineSpacing: 312,
  },
  centered_text: {
    font: { name: 'Times New Roman', eastAsia: '黑体', size: 30, bold: true },
    paragraph: { alignment: 'center' },
    lineSpacing: 312,
  },
  equation: {
    font: { name: 'Times New Roman', eastAsia: '仿宋', size: 32, italic: true },
    paragraph: { alignment: 'center' },
    lineSpacing: 312,
  },
  table: {
    font: { name: 'Times New Roman', eastAsia: '仿宋', size: 24 },
    paragraph: { alignment: 'center' },
    lineSpacing: 276,
  },
  table_header: {
    font: { name: 'Times New Roman', eastAsia: '黑体', size: 24, bold: true },
    paragraph: { alignment: 'center' },
    lineSpacing: 276,
  },
  // GB/T 9704-2012 says nothing about captions, so 公文 keeps its table text
  // typography for both roles instead of falling back to indented body text.
  table_caption: {
    font: { name: 'Times New Roman', eastAsia: '仿宋', size: 24 },
    paragraph: { alignment: 'center' },
    lineSpacing: 276,
  },
  figure_caption: {
    font: { name: 'Times New Roman', eastAsia: '仿宋', size: 24 },
    paragraph: { alignment: 'center' },
    lineSpacing: 276,
  },
};

// ── Style Resolver ────────────────────────────────────────

function fallbackStyles(docType?: string): Record<string, ParagraphStyle> {
  return docType === 'official' ? OFFICIAL_STYLES : THESIS_STYLES;
}

/** Drop keys whose value is `undefined` so they cannot shadow a fallback. */
function definedOnly<T extends object>(value: T | undefined): Partial<T> {
  if (!value) return {};
  return Object.fromEntries(
    Object.entries(value).filter(([, v]) => v !== undefined),
  ) as Partial<T>;
}

/**
 * Resolved-style cache, keyed by the template object it was derived from.
 *
 * `resolveStyle` runs at least once per block and three times per table, and it
 * used to derive the same `ParagraphStyle` every time: a `for…in` over the
 * template's role table for every role without a recorded winner (the shipped
 * template records 12 of 48), then two `Object.entries`/`fromEntries` merges and
 * four spreads. The result depends only on (template, role, docType), and all
 * three are fixed for the duration of a render, so it is cached.
 *
 * A `WeakMap` keeps entries from outliving their template, and `renderDocument`
 * drops its template's entry before it starts, so a template mutated between
 * renders is still read fresh.
 */
const STYLE_CACHE = new WeakMap<DocumentTemplate, Map<string, ParagraphStyle>>();

/**
 * Resolve the effective ParagraphStyle for a semantic role.
 *
 * The template wins for every property it actually declares; properties it
 * leaves unset fall back to the domain defaults. Merging per property rather
 * than per object is what lets a template that carries no formatting still
 * produce correctly styled headings instead of body text.
 */
function resolveStyle(template: DocumentTemplate, role: StyleRole, docType?: string): ParagraphStyle {
  let byRole = STYLE_CACHE.get(template);
  if (!byRole) {
    byRole = new Map();
    STYLE_CACHE.set(template, byRole);
  }
  const key = `${docType ?? 'thesis'}|${role}`;
  const cached = byRole.get(key);
  if (cached) return cached;
  const resolved = resolveStyleUncached(template, role, docType);
  byRole.set(key, resolved);
  return resolved;
}

function resolveStyleUncached(template: DocumentTemplate, role: StyleRole, docType?: string): ParagraphStyle {
  const defaults = fallbackStyles(docType);
  const fallback = defaults[role] || defaults['paragraph'];

  // 公文 follows GB/T 9704-2012 regardless of the supplied thesis template.
  if (docType === 'official') return fallback;

  // A parsed template records which style wins each role, because a plain
  // object enumerates integer-like keys first and Word names its stock styles
  // `1`, `2`, `3` — "the first entry carrying this role" is not the author's
  // choice. Templates written before that field existed still resolve by scan,
  // which also stops at the first match rather than building every pair.
  const roles = template.styleRoles;
  let styleId: string | undefined = template.roleWinners?.[role];
  if (!styleId) {
    for (const id in roles) {
      if (Object.prototype.hasOwnProperty.call(roles, id) && roles[id] === role) {
        styleId = id;
        break;
      }
    }
  }
  const fromTemplate = styleId ? template.styles?.[styleId] : undefined;

  if (!fromTemplate) return fallback;

  // A template line spacing carries its own rule; the fallback's only applies
  // when the template is silent about the distance itself.
  const templateLine = fromTemplate.lineSpacing;

  return {
    font: { ...fallback.font, ...definedOnly(fromTemplate.font) },
    paragraph: { ...fallback.paragraph, ...definedOnly(fromTemplate.paragraph) },
    lineSpacing: templateLine ?? fallback.lineSpacing,
    lineSpacingRule: templateLine === undefined ? fallback.lineSpacingRule : fromTemplate.lineSpacingRule,
  };
}

/**
 * `lineSpacing` plus its rule, ready to spread into a `Paragraph`'s `spacing`.
 *
 * The rule cannot be omitted when it is not `auto`: an `exact` value is in
 * twips and would otherwise be read as a multiple of the line height.
 */
function lineStyle(style: ParagraphStyle, fallbackLine?: number): { line?: number; lineRule?: LineSpacingRule } {
  const line = style.lineSpacing ?? fallbackLine;
  if (line === undefined) return {};
  return style.lineSpacingRule ? { line, lineRule: style.lineSpacingRule } : { line };
}

/**
 * The paragraph indent, as OOXML wants it.
 *
 * A hanging indent and a first-line indent are mutually exclusive, so whichever
 * the template declares wins. Headings in the guide carry a hanging indent
 * (0.75 / 1 / 1.25 cm), which used to be dropped on the floor.
 */
function indentOf(style: ParagraphStyle): { hanging?: number; left?: number; firstLine?: number } | undefined {
  const { hangingIndent, leftIndent, firstLineIndent } = style.paragraph;
  if (hangingIndent) return { hanging: hangingIndent, ...(leftIndent ? { left: leftIndent } : {}) };
  if (firstLineIndent) return { firstLine: firstLineIndent };
  return undefined;
}

// ── Block Renderers ───────────────────────────────────────

/** Outline level for a heading kind: chapter is level 1, so it is level 0 here. */
function headingOutlineLevel(type: string): number {
  switch (type) {
    case 'chapter': return 0;
    case 'section': return 1;
    case 'subsection': return 2;
    case 'subsubsection': return 3;
    default: return 0;
  }
}

function renderHeading(
  template: DocumentTemplate,
  block: HeadingBlock,
  docType?: string,
): Paragraph {
  // `styleRole` lets a front/back-matter heading use its own format.
  const style = resolveStyle(template, block.styleRole ?? block.type, docType);
  // The template binds a heading style to a numbering instance through
  // `w:numPr`. A number written in the content file wins — the author asked for
  // that exact label — and the template's own numbering fills in otherwise.
  const binding = style.paragraph.numbering;
  const autoNumbering = binding && binding.numId !== 0
    ? { reference: templateNumberingReference(binding.numId), level: binding.ilvl }
    : undefined;

  const numberPrefix = block.number ? `${block.number}  ` : '';
  // A 目录 entry points at a bookmark on its heading, so the heading carries one.
  const listed = headingBookmarks.get(block.text);
  const tocBookmarkId = listed ? nextBookmarkId() : 0;
  return new Paragraph({
    // `STYLEREF` and a `SEQ … \s` field look the chapter number up by outline
    // level, so a heading that declares none cannot be found by either.
    outlineLevel: headingOutlineLevel(block.type),
    // A number written in the content file wins: the author asked for that
    // exact label. Otherwise the template's own numbering takes over, which is
    // how a guide that numbers chapters `1`, `1.1`, `2.3.3` gets honoured.
    numbering: block.number ? undefined : autoNumbering,
    children: [
      ...(listed ? [captionBookmarkStart(listed.bookmark, tocBookmarkId)] : []),
      ...styledRuns(numberPrefix + block.text, style, style.font.bold),
      ...(listed ? [captionBookmarkEnd(tocBookmarkId)] : []),
    ],
    alignment: mapAlignment(style.paragraph.alignment),
    // Headings carry the guide's hanging indent (0.75 / 1 / 1.25 cm); this
    // paragraph never emitted an indent at all, so it was dropped silently.
    indent: indentOf(style),
    // The guide wants a page break between chapters, and the chapter style
    // declares one.
    pageBreakBefore: style.paragraph.pageBreakBefore,
    // A heading that declares `w:keepNext` must not be left alone at the foot of
    // a page. The parser read the flag but the schema had no field to carry it,
    // so it never reached the document.
    keepNext: style.paragraph.keepNext,
    keepLines: style.paragraph.keepLines,
    spacing: {
      before: style.paragraph.spaceBefore,
      after: style.paragraph.spaceAfter,
      ...lineStyle(style),
    },
  });
}

function renderParagraph(
  template: DocumentTemplate,
  block: ParagraphBlock,
  docType?: string,
  role: StyleRole = block.type,
): Paragraph {
  const style = resolveStyle(template, role, docType);
  return new Paragraph({
    children: styledRuns(block.text, style),
    alignment: mapAlignment(style.paragraph.alignment),
    indent: indentOf(style),
    spacing: {
      before: style.paragraph.spaceBefore ?? 12,
      after: style.paragraph.spaceAfter ?? 12,
      ...lineStyle(style),
    },
  });
}

function renderCenteredText(
  template: DocumentTemplate,
  block: CenteredTextBlock,
  docType?: string,
): Paragraph {
  const style = resolveStyle(template, block.styleRole ?? 'centered_text', docType);
  const actualSize = block.font_size_pt
    ? block.font_size_pt * 2  // convert pt → half-pt
    : style.font.size;
  const actualBold = block.bold !== undefined ? block.bold : style.font.bold;
  return new Paragraph({
    children: styledRuns(block.text, style, actualBold, actualSize),
    // The cover's own styles decide the alignment: 校名行 is centred while the
    // 研究生/指导教师 rows are justified rows. Hardcoding centre ignored them.
    //
    // `mapAlignment` never returns undefined — it falls back to JUSTIFIED — so
    // the `?? AlignmentType.CENTER` that used to sit here could not fire. The
    // centred fallback the comment describes comes from `centered_text`'s own
    // style, which declares `alignment: 'center'`.
    alignment: mapAlignment(style.paragraph.alignment),
    spacing: { ...lineStyle(style) },
  });
}

function renderEquation(
  template: DocumentTemplate,
  block: EquationBlock,
  docType?: string,
): Paragraph {
  // V3: emit native Office Math (OMML) for the supported LaTeX subset.
  const style = resolveStyle(template, 'equation', docType);
  let equationChildren: MathComponent[];
  try {
    equationChildren = mathComponents(latexToMathAst(block.latex));
  } catch (error: unknown) {
    // Falling back to Unicode is fine, but silently swallowing the reason made
    // unsupported LaTeX indistinguishable from a rendering bug.
    const message = error instanceof Error ? error.message : String(error);
    console.warn(
      `Falling back to plain text for equation "${block.latex}": ${message}`,
    );
    equationChildren = [new MathRun(latexToPlainText(block.latex))];
  }

  if (block.number) {
    const page = template.page;
    const printableWidth = Math.max(
      1200,
      (page?.width ?? 11906) - (page?.margins?.left ?? 1701) - (page?.margins?.right ?? 1701) - (page?.gutter ?? 0),
    );
    const centerPos = Math.round(printableWidth / 2);
    const rightPos = printableWidth;

    return new Paragraph({
      alignment: AlignmentType.LEFT,
      tabStops: [
        { type: TabStopType.CENTER, position: centerPos },
        { type: TabStopType.RIGHT, position: rightPos },
      ],
      children: [
        new TextRun('\t'),
        new DocxMath({ children: equationChildren }),
        new TextRun({
          text: `\t(${block.number})`,
          size: style.font.size ?? 24,
          font: {
            ascii: style.font.name ?? 'Times New Roman',
            hAnsi: style.font.name ?? 'Times New Roman',
            eastAsia: style.font.eastAsia ?? '宋体',
            cs: style.font.name ?? 'Times New Roman',
          },
        }),
      ],
      spacing: { before: 120, after: 120, ...lineStyle(style, 312) },
    });
  }

  return new Paragraph({
    children: [new DocxMath({ children: equationChildren })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 120, ...lineStyle(style, 312) },
  });
}

function mathComponents(nodes: LatexMathNode[]): MathComponent[] {
  return nodes.map(node => {
    switch (node.type) {
      case 'run':
        return new MathRun(node.text);
      case 'fraction':
        return node.bar === false
          ? mathNoBarFraction(node.numerator, node.denominator)
          : new MathFraction({
              numerator: mathComponents(node.numerator),
              denominator: mathComponents(node.denominator),
            });
      case 'radical':
        return new MathRadical({
          children: mathComponents(node.children),
          degree: node.degree ? mathComponents(node.degree) : undefined,
        });
      case 'script': {
        const base = mathComponents(node.base);
        if (node.subScript && node.superScript) {
          return new MathSubSuperScript({
            children: base,
            subScript: mathComponents(node.subScript),
            superScript: mathComponents(node.superScript),
          });
        }
        if (node.subScript) {
          return new MathSubScript({ children: base, subScript: mathComponents(node.subScript) });
        }
        return new MathSuperScript({ children: base, superScript: mathComponents(node.superScript ?? []) });
      }
      case 'sum':
        return (!node.subScript && !node.superScript)
          ? new MathRun('∑')
          : mathNary('∑', node.subScript, node.superScript);
      case 'integral':
        return (!node.subScript && !node.superScript)
          ? new MathRun('∫')
          : mathNary('∫', node.subScript, node.superScript);
      case 'nary':
        return (!node.subScript && !node.superScript)
          ? new MathRun(node.operator)
          : mathNary(node.operator, node.subScript, node.superScript);
      case 'accent':
        return mathAccent(node.accent, node.children);
      case 'bar':
        return mathBar(node.position, node.children);
      case 'delimiter':
        return mathDelimiter(node.opening, node.closing, mathComponents(node.children));
      case 'function':
        return new MathFunction({
          name: [new MathRun(node.name)],
          children: mathComponents(node.children),
        });
      case 'limit': {
        let component: MathComponent = new MathRun(node.name);
        if (node.subScript) {
          component = asMathComponent(new MathLimitLower({
            children: [component], limit: mathComponents(node.subScript),
          }));
        }
        if (node.superScript) {
          component = asMathComponent(new MathLimitUpper({
            children: [component], limit: mathComponents(node.superScript),
          }));
        }
        return component;
      }
      case 'overUnder': {
        let components = mathComponents(node.base);
        if (node.under) {
          components = [asMathComponent(new MathLimitLower({
            children: components, limit: mathComponents(node.under),
          }))];
        }
        if (node.over) {
          components = [asMathComponent(new MathLimitUpper({
            children: components, limit: mathComponents(node.over),
          }))];
        }
        return components[0] ?? new MathRun('');
      }
      case 'matrix': {
        const matrix = mathMatrix(node.rows);
        return node.opening || node.closing
          ? mathDelimiter(node.opening, node.closing, [matrix])
          : matrix;
      }
      default: {
        // Without this the switch fell through to `undefined`, which then
        // landed in the OMML children array and produced invalid Office Math
        // instead of an error the caller could see.
        const unknown = node as { type?: unknown };
        throw new Error(`Unsupported math node type: "${String(unknown.type)}"`);
      }
    }
  });
}

function asMathComponent(element: BuilderElement | MathLimitLower | MathLimitUpper): MathComponent {
  return element as unknown as MathComponent;
}

function mathProperty(name: string, value: string): BuilderElement {
  return new BuilderElement({
    name,
    attributes: { value: { key: 'm:val', value } },
  });
}

function mathArgument(name: string, children: readonly MathComponent[]): BuilderElement {
  return new BuilderElement({ name, children });
}

function mathAccent(accent: string, children: LatexMathNode[]): MathComponent {
  const properties = new BuilderElement({
    name: 'm:accPr',
    children: [mathProperty('m:chr', accent)],
  });
  return asMathComponent(new BuilderElement({
    name: 'm:acc',
    children: [properties, mathArgument('m:e', mathComponents(children))],
  }));
}

function mathBar(position: 'top' | 'bottom', children: LatexMathNode[]): MathComponent {
  const properties = new BuilderElement({
    name: 'm:barPr',
    children: [mathProperty('m:pos', position)],
  });
  return asMathComponent(new BuilderElement({
    name: 'm:bar',
    children: [properties, mathArgument('m:e', mathComponents(children))],
  }));
}

function mathDelimiter(opening: string, closing: string, children: MathComponent[]): MathComponent {
  const properties = new BuilderElement({
    name: 'm:dPr',
    children: [mathProperty('m:begChr', opening), mathProperty('m:endChr', closing)],
  });
  return asMathComponent(new BuilderElement({
    name: 'm:d',
    children: [properties, mathArgument('m:e', children)],
  }));
}

function mathMatrix(rows: LatexMathNode[][][]): MathComponent {
  const width = Math.max(1, ...rows.map(row => row.length));
  const matrixRows = rows.map(row => {
    const cells = Array.from({ length: width }, (_, index) => row[index] ?? []);
    return new BuilderElement({
      name: 'm:mr',
      children: cells.map(cell => mathArgument('m:e', mathComponents(cell))),
    });
  });
  return asMathComponent(new BuilderElement({
    name: 'm:m',
    children: matrixRows,
  }));
}

function mathNoBarFraction(numerator: LatexMathNode[], denominator: LatexMathNode[]): MathComponent {
  const properties = new BuilderElement({
    name: 'm:fPr',
    children: [mathProperty('m:type', 'noBar')],
  });
  return asMathComponent(new BuilderElement({
    name: 'm:f',
    children: [
      properties,
      mathArgument('m:num', mathComponents(numerator)),
      mathArgument('m:den', mathComponents(denominator)),
    ],
  }));
}

function mathNary(operator: string, subScript?: LatexMathNode[], superScript?: LatexMathNode[]): MathComponent {
  const properties = new BuilderElement({
    name: 'm:naryPr',
    children: [
      mathProperty('m:chr', operator),
      mathProperty('m:limLoc', 'undOvr'),
      mathProperty('m:subHide', subScript ? '0' : '1'),
      mathProperty('m:supHide', superScript ? '0' : '1'),
    ],
  });
  return asMathComponent(new BuilderElement({
    name: 'm:nary',
    children: [
      properties,
      mathArgument('m:sub', subScript ? mathComponents(subScript) : []),
      mathArgument('m:sup', superScript ? mathComponents(superScript) : []),
      mathArgument('m:e', []),
    ],
  }));
}

/** Build a run that inherits the resolved style's font settings. */
function styledRun(
  text: string,
  style: ParagraphStyle,
  forceBold?: boolean,
  customSize?: number,
): TextRun {
  return new TextRun({
    text,
    bold: forceBold ?? style.font.bold,
    italics: style.font.italic,
    color: style.font.color,
    size: customSize ?? style.font.size,
    font: {
      ascii: style.font.name,
      hAnsi: style.font.name,
      eastAsia: style.font.eastAsia,
      cs: style.font.name,
    },
  });
}

/**
 * Parse text into runs, converting inline LaTeX math (`$ ... $`)
 * into native Office Math (`DocxMath`) elements.
 */
function styledRuns(
  text: string,
  style: ParagraphStyle,
  forceBold?: boolean,
  customSize?: number,
): (TextRun | DocxMath | XmlComponent)[] {
  const hasMath = text.includes('$');
  CROSS_REF.lastIndex = 0;
  const hasRef = CROSS_REF.test(text);
  CROSS_REF.lastIndex = 0;
  if (!hasMath && !hasRef) {
    return [styledRun(text, style, forceBold, customSize)];
  }

  // One pass, cutting at whichever marker comes first: an inline `$…$` equation
  // or a `[[表3-1]]` cross-reference. A reference inside a sentence must not
  // disturb the runs around it, so both are handled in the same scan.
  const runs: (TextRun | DocxMath | XmlComponent)[] = [];
  const mathSource = '\\$(?!\\s)([^$\\n]+?)(?<![\\s\\\\])\\$';
  const pattern = new RegExp(`${mathSource}|\\[\\[(表|图)\\s*([^\\[\\]]+?)\\]\\]`, 'g');
  let lastIndex = 0;
  let match: RegExpExecArray | null;

  while ((match = pattern.exec(text)) !== null) {
    if (match.index > lastIndex) {
      runs.push(styledRun(text.slice(lastIndex, match.index), style, forceBold, customSize));
    }

    if (match[2]) {
      const resolved = resolveReference(`${match[2]}${match[3].trim()}`);
      runs.push(resolved === undefined ? styledRun(`[[${match[2]}${match[3].trim()}]]`, style, forceBold, customSize) : crossReferenceField(resolved.bookmark, resolved.result));
    } else {
      const mathLatex = match[1].trim();
      try {
        const ast = latexToMathAst(mathLatex);
        runs.push(new DocxMath({ children: mathComponents(ast) }));
      } catch {
        runs.push(styledRun(`$${mathLatex}$`, style, forceBold, customSize));
      }
    }
    lastIndex = pattern.lastIndex;
  }

  if (lastIndex < text.length) {
    runs.push(styledRun(text.slice(lastIndex), style, forceBold, customSize));
  }

  return runs.length > 0 ? runs : [styledRun(text, style, forceBold, customSize)];
}

/**
 * A caption's number and title, separated by exactly one space.
 *
 * The renderer used to emit `${number}  ${caption}` — two spaces, which Word
 * keeps. A content file may write the label into either half (`number: '表3-1'`
 * with a bare title, or a number with the label already inside the caption), so
 * both sides are normalised and the two are joined once instead of blindly
 * prefixed.
 */
function captionText(number: string | undefined, caption: string | undefined): string {
  // A table written without a caption used to reach this function and throw
  // `Cannot read properties of undefined (reading 'trim')` from inside the
  // renderer — after `validateDocument` had already passed the document, since
  // the validator only checked a table's headers and data. Both ends are fixed;
  // this one keeps a malformed-but-loadable document from crashing the build.
  const title = (caption ?? '').trim();
  const label = number?.trim();
  if (!label) return title;
  if (title.startsWith(label)) {
    const rest = title.slice(label.length).replace(/^[\s\u3000]+/, '');
    return rest ? `${label} ${rest}` : label;
  }
  return title ? `${label} ${title}` : label;
}

function renderTable(
  template: DocumentTemplate,
  block: TableBlock,
  docType?: string,
): (Paragraph | Table)[] {
  const colCount = block.headers.length;
  if (colCount === 0) {
    throw new Error(`Table "${block.caption}" must define at least one header column.`);
  }
  if (block.columnWidths && block.columnWidths.length !== colCount) {
    throw new Error(`Table "${block.caption}" has ${colCount} columns but ${block.columnWidths.length} column widths.`);
  }
  if (block.columnWidths?.some(width => !Number.isFinite(width) || width <= 0)) {
    throw new Error(`Table "${block.caption}" column widths must be positive numbers.`);
  }

  // Fit the table to the template's printable width rather than a hardcoded A4
  // constant: on a narrower page (or with wider margins) every table used to
  // overflow the right margin, because the cell widths always summed to 8504.
  const page = template.page;
  const printableWidth = Math.max(
    1200,
    (page?.width ?? 11906) - (page?.margins?.left ?? 0) - (page?.margins?.right ?? 0),
  );
  // Column widths, in order of authority: the content file, the template's own
  // table when it has the same number of columns, then an equal split of the
  // printable width. The template's widths only mean something at the same
  // column count — its 12-column example says nothing about a 2-column table.
  const templateColumns = template.tableFormat?.columnWidths;
  const colWidths = block.columnWidths
    ?? (templateColumns && templateColumns.length === colCount
      ? templateColumns
      : Array.from({ length: colCount }, () => Math.floor(printableWidth / colCount)));
  const tableWidth = colWidths.reduce((sum, width) => sum + width, 0);

  // Table typography comes from the template like every other block, so a
  // template that styles its tables is no longer silently ignored.
  // A table caption is spaced differently from a figure caption: the guide asks
  // for 段前1行 on a table and 段前0.1行 on a figure.
  const captionStyle = resolveStyle(template, 'table_caption', docType);
  const headerStyle = resolveStyle(template, 'table_header', docType);
  const bodyStyle = resolveStyle(template, 'table', docType);

  const isAcademic = docType === 'thesis' || docType === 'journal' || template.meta?.documentType === 'thesis' || template.meta?.documentType === 'journal';
  // Standard academic thesis (GB/T 7713 / USTB guide) requires 三线表:
  // 1.5pt top/bottom border, 0.75pt header bottom border, no vertical borders, no header shading.
  const isThreeLine = isAcademic
    ? block.showGridlines !== true
    : block.showGridlines === false;

  const showGrid = !isThreeLine && block.showGridlines !== false;
  const showShading = isThreeLine ? block.headerShading === true : block.headerShading !== false;

  // The template's own strokes, when it declared any. `w:sz` is in eighths of a
  // point, which is what the writer library's `size` means too, so the value
  // passes through unchanged. A template that declares nothing keeps the old
  // shape: 1.5 pt outer rules, 0.75 pt header rule, nothing inside.
  const declared = template.tableFormat?.borders;
  type BorderShape = { style: (typeof BorderStyle)[keyof typeof BorderStyle]; size: number; color: string };
  const toBorder = (spec: BorderSpec | undefined, fallback: BorderShape): BorderShape =>
    spec && spec.style !== 'none' && spec.style !== 'nil'
      ? { style: borderStyleOf(spec.style), size: spec.size, color: spec.color === 'auto' ? '000000' : spec.color }
      : fallback;
  const BO = toBorder(declared?.top, { style: BorderStyle.SINGLE, size: 12, color: '000000' });
  const BI = toBorder(declared?.insideH, { style: BorderStyle.SINGLE, size: 6, color: '000000' });
  const BLeft = toBorder(declared?.left, { style: BorderStyle.SINGLE, size: 2, color: 'FFFFFF' });
  const BRight = toBorder(declared?.right, { style: BorderStyle.SINGLE, size: 2, color: 'FFFFFF' });
  const BBottom = toBorder(declared?.bottom, BO);
  const BInsideV = toBorder(declared?.insideV, { style: BorderStyle.SINGLE, size: 2, color: 'FFFFFF' });
  // `nil` is the spec-correct value for "no border" and Word honours it, but
  // LibreOffice — and any viewer built on it, including WPS — falls back to its
  // default table grid instead, which ruled every row and column of a three-line
  // table. A white hairline draws nothing on paper and is honoured everywhere.
  // Verified by rendering one table three ways and comparing the images.
  const NO = { style: BorderStyle.SINGLE, size: 2, color: 'FFFFFF' };

  // When the template draws a full grid, that is the format — the academic
  // three-line default only applies to a template that did not say otherwise.
  const templateDrawsGrid = Boolean(declared?.insideV && declared.insideV.style !== 'none' && declared.insideV.style !== 'nil');

  // Caption. The role's own 段前/段后 are what reach `w:spacing`: hardcoding
  // 120/60 here overwrote the template's `ub` (150/50), which is why the
  // university's 段前1行 never appeared in the produced document.xml. A table with
  // no caption gets no empty paragraph carrying that 段前 spacing.
  const captionValue = captionText(block.number, block.caption);
  // A caption's number is bookmarked, so a `[[表3-1]]` in the body becomes a
  // REF field that Word keeps in step with the caption instead of a hand-typed
  // number that goes stale the moment a table is inserted.
  const captionBookmarkId = nextBookmarkId();
  const captionPara = captionValue
    ? new Paragraph({
        children: [
          ...(block.number
            ? [
                ...captionNumberRuns(captionIndex.entries.get(block.number)!, captionStyle, captionBookmarkId),
                styledRun(captionValue.slice(block.number.length), captionStyle, true),
              ]
            : styledRuns(captionValue, captionStyle, true)),
        ],
        alignment: mapAlignment(captionStyle.paragraph.alignment),
        spacing: {
          before: captionStyle.paragraph.spaceBefore,
          after: captionStyle.paragraph.spaceAfter,
          ...lineStyle(captionStyle, 312),
        },
      })
    : undefined;

  // Header row
  const headerRow = new TableRow({
    tableHeader: true,
    children: block.headers.map((h, ci) =>
      new TableCell({
        width: { size: colWidths[ci], type: WidthType.DXA },
        // Only the horizontal rules live on the cell: a vertical rule declared
        // here paints over the black rule at the column join, which is what put
        // notches in the header rule. The table-level borders supply verticals.
        borders: {
          top: BO,
          bottom: templateDrawsGrid ? BBottom : BI,
        },
        shading: showShading
          ? {
              fill: template.tableFormat?.headerFill ?? 'D9D9D9',
              type: ShadingType.SOLID,
              color: 'auto',
            }
          : undefined,
        children: [
          new Paragraph({
            children: styledRuns(
              h,
              headerStyle,
              template.tableFormat?.headerBold ?? true,
              template.tableFormat?.headerSize,
            ),
            alignment: mapAlignment(headerStyle.paragraph.alignment),
            // `lineStyle` carries the template's `lineRule`. Passing `line`
            // alone dropped it, so an `exact` twips value was read as a multiple
            // of the line height — the one place the template's spacing was
            // discarded rather than honoured.
            spacing: {
              before: headerStyle.paragraph.spaceBefore ?? 40,
              after: headerStyle.paragraph.spaceAfter ?? 40,
              ...lineStyle(headerStyle),
            },
          }),
        ],
      }),
    ),
  });

  // Data rows
  const dataRows = block.data.map((row, ri) => {
    // `validateDocument` reports a row whose cell count is wrong, but a caller that
    // uses the renderer directly got a silently padded or truncated row. Say so.
    if (row.length !== colCount) {
      console.warn(
        `Table "${block.caption}" row ${ri + 1} has ${row.length} cell(s); the table declares ${colCount} column(s). ` +
        `${row.length < colCount ? 'Missing cells are rendered empty' : 'Extra cells are dropped'}.`,
      );
    }
    const isLast = ri === block.data.length - 1;
    return new TableRow({
      children: block.headers.map((_, ci) =>
        new TableCell({
          width: { size: colWidths[ci], type: WidthType.DXA },
          // Horizontal only, for the same reason as the header row.
          borders: {
            top: NO,
            bottom: isLast ? BBottom : NO,
          },
          children: [
            new Paragraph({
              children: styledRuns(row[ci] ?? '', bodyStyle),
              alignment: mapAlignment(bodyStyle.paragraph.alignment),
              // Same as the header row: keep the template's rule and its
              // 段前/段后, and only fall back to the old tight values when the
              // template does not declare them.
              spacing: {
                before: bodyStyle.paragraph.spaceBefore ?? 20,
                after: bodyStyle.paragraph.spaceAfter ?? 20,
                ...lineStyle(bodyStyle),
              },
            }),
          ],
        }),
      ),
    });
  });

  const table = new Table({
    // `w:jc` and `w:tblLayout` from the template. A template that centres its
    // tables and fixes its column widths said so; the renderer used to drop both.
    ...(template.tableFormat?.alignment
      ? { alignment: mapAlignment(template.tableFormat.alignment) }
      : {}),
    ...(template.tableFormat?.layout === 'fixed'
      ? { layout: TableLayoutType.FIXED }
      : template.tableFormat?.layout === 'autofit'
        ? { layout: TableLayoutType.AUTOFIT }
        : {}),
    // Emitting the grid explicitly keeps `w:gridCol` in step with the cell
    // widths; without it the two disagree and Word re-lays the table out.
    columnWidths: colWidths,
    // Declare the same total the cells declare; a 100% table width alongside
    // absolute cell widths is contradictory OOXML.
    width: { size: tableWidth, type: WidthType.DXA },
    // The table-level borders are what a table with no per-cell override draws,
    // so they have to agree with the cells: a template that draws a full grid
    // gets its inside rules here too, not the academic three-line default.
    borders: templateDrawsGrid
      ? {
          top: BO,
          bottom: BBottom,
          left: BLeft,
          right: BRight,
          insideHorizontal: BI,
          insideVertical: BInsideV,
        }
      : isThreeLine
        ? {
            top: BO, bottom: BBottom, left: NO, right: NO,
            insideHorizontal: NO, insideVertical: NO,
          }
        : (showGrid
          ? {
              top: BO, bottom: BBottom, left: BO, right: BO,
              insideHorizontal: BI, insideVertical: BI,
            }
          : {
              top: NO, bottom: NO, left: NO, right: NO,
              insideHorizontal: NO, insideVertical: NO,
            }),
    rows: [headerRow, ...dataRows],
  });

  // Spacer after table
  const spacer = new Paragraph({ spacing: { line: 312 }, children: [] });

  // The caption goes ahead of the table: 表题 sits above the table it names.
  return captionPara ? [captionPara, table, spacer] : [table, spacer];
}

function getImageSize(buffer: Buffer): { width: number; height: number } | null {
  // Check PNG signature
  if (buffer.length >= 24 && buffer.readUInt32BE(0) === 0x89504E47 && buffer.readUInt32BE(4) === 0x0D0A1A0A) {
    return {
      width: buffer.readUInt32BE(16),
      height: buffer.readUInt32BE(20)
    };
  }
  
  // Check JPEG signature
  if (buffer.length >= 4 && buffer.readUInt16BE(0) === 0xFFD8) {
    let offset = 2;
    while (offset + 2 < buffer.length) {
      const marker = buffer.readUInt16BE(offset);
      offset += 2;
      
      if (marker >= 0xFFC0 && marker <= 0xFFCF && marker !== 0xFFC4 && marker !== 0xFFC8 && marker !== 0xFFCC) {
        // SOF segment: need at least 7 bytes (length(2) + precision(1) + height(2) + width(2))
        if (offset + 7 > buffer.length) break;
        return {
          height: buffer.readUInt16BE(offset + 3),
          width: buffer.readUInt16BE(offset + 5)
        };
      }
      
      if (offset + 2 > buffer.length) break;
      const segLength = buffer.readUInt16BE(offset);
      if (segLength < 2) break; // invalid segment length
      offset += segLength;
    }
  }
  
  // Check GIF signature
  if (buffer.length >= 10 && buffer.readUInt32BE(0) === 0x47494638 && (buffer.readUInt16BE(4) === 0x3761 || buffer.readUInt16BE(4) === 0x3961)) {
    return {
      width: buffer.readUInt16LE(6),
      height: buffer.readUInt16LE(8)
    };
  }

  return null;
}

function detectImageType(buffer: Buffer): 'png' | 'jpg' | 'gif' | 'bmp' | null {
  if (buffer.length >= 8 && buffer.readUInt32BE(0) === 0x89504E47 && buffer.readUInt32BE(4) === 0x0D0A1A0A) return 'png';
  if (buffer.length >= 2 && buffer.readUInt16BE(0) === 0xFFD8) return 'jpg';
  if (buffer.length >= 6 && buffer.toString('ascii', 0, 3) === 'GIF') return 'gif';
  if (buffer.length >= 2 && buffer.toString('ascii', 0, 2) === 'BM') return 'bmp';
  return null;
}

/**
 * One panel image, sized to fit the cell it sits in.
 *
 * A missing or unreadable file still produces a paragraph, so the panel keeps its
 * label and the layout keeps its shape instead of collapsing.
 */
function panelImage(item: FigureItem, contentDir: string | undefined, maxWidthPx: number): Paragraph {
  let imgPath = item.path;
  if (contentDir && !isAbsolute(imgPath)) imgPath = resolve(contentDir, imgPath);

  let width = item.width;
  let height = item.height;
  if (existsSync(imgPath)) {
    try {
      const buffer = readFileSync(imgPath);
      const size = getImageSize(buffer);
      if (size && size.width > 0 && size.height > 0) {
        const ratio = size.width / size.height;
        if (!width && !height) { width = Math.min(size.width, maxWidthPx); height = width / ratio; }
        else if (width && !height) height = width / ratio;
        else if (!width && height) width = height * ratio;
      } else {
        width = width || Math.min(300, maxWidthPx);
        height = height || 200;
      }
      if (width! > maxWidthPx) { const shrink = maxWidthPx / width!; width = maxWidthPx; height = height! * shrink; }
      return new Paragraph({
        children: [new ImageRun({ data: buffer, transformation: { width: width!, height: height! }, type: detectImageType(buffer) ?? 'png' })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 40, after: 40 },
      });
    } catch {
      // fall through to the placeholder
    }
  }
  return new Paragraph({
    children: [new TextRun({ text: `[未找到图片: ${item.path}]`, size: 18, italics: true })],
    alignment: AlignmentType.CENTER,
    spacing: { before: 40, after: 40 },
  });
}

/**
 * A group figure — 图1(a), (b), (c) — laid out as a borderless table.
 *
 * The layout table is how Word itself does side-by-side panels, and it keeps the
 * labels aligned under their images. The group is numbered once; the letters
 * label the panels.
 */
function renderFigureGroup(
  template: DocumentTemplate,
  block: FigureBlock,
  contentDir?: string,
  docType?: string,
): (Paragraph | Table)[] {
  const captionStyle = resolveStyle(template, 'figure_caption', docType);
  const captionSpacing = lineStyle(captionStyle, 312);
  const page = template.page;
  const printableTwips = Math.max(1200, (page?.width ?? 11906) - (page?.margins?.left ?? 0) - (page?.margins?.right ?? 0));

  const items = block.items ?? [];
  const columns = Math.max(1, Math.min(block.columns ?? items.length, items.length));
  const cellTwips = Math.floor(printableTwips / columns);
  const maxWidthPx = Math.floor(cellTwips / 15);
  const noBorder = { top: NO_BORDER, bottom: NO_BORDER, left: NO_BORDER, right: NO_BORDER };

  const rows: TableRow[] = [];
  for (let start = 0; start < items.length; start += columns) {
    const slice = items.slice(start, start + columns);
    rows.push(new TableRow({
      children: Array.from({ length: columns }, (_, column) => {
        const item = slice[column];
        const children: Paragraph[] = [];
        if (item) {
          children.push(panelImage(item, contentDir, maxWidthPx));
          const label = item.label ?? `(${String.fromCharCode(97 + start + column)})`;
          children.push(new Paragraph({
            children: styledRuns(label, captionStyle),
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: item.caption ? 0 : 60 },
          }));
          if (item.caption) {
            children.push(new Paragraph({
              children: styledRuns(item.caption, captionStyle),
              alignment: AlignmentType.CENTER,
              spacing: { before: 0, after: 60 },
            }));
          }
        }
        return new TableCell({ width: { size: cellTwips, type: WidthType.DXA }, borders: noBorder, children });
      }),
    }));
  }

  const table = new Table({
    width: { size: cellTwips * columns, type: WidthType.DXA },
    borders: { ...noBorder, insideHorizontal: NO_BORDER, insideVertical: NO_BORDER },
    rows,
  });

  // The group caption, numbered once and bookmarked like any other.
  const caption = new Paragraph({
    children: [
      ...(block.number
        ? [
            ...captionNumberRuns(captionIndex.entries.get(block.number)!, captionStyle, nextBookmarkId()),
            styledRun(captionText(block.number, block.caption).slice(block.number.length), captionStyle),
          ]
        : styledRuns(captionText(block.number, block.caption), captionStyle)),
    ],
    alignment: mapAlignment(captionStyle.paragraph.alignment),
    spacing: { before: captionStyle.paragraph.spaceBefore, after: captionStyle.paragraph.spaceAfter, ...captionSpacing },
  });

  return [table, caption];
}

function renderFigure(
  template: DocumentTemplate,
  block: FigureBlock,
  contentDir?: string,
  docType?: string,
): (Paragraph | Table)[] {
  if (block.items && block.items.length > 0) {
    return renderFigureGroup(template, block, contentDir, docType);
  }
  // A figure with neither panels nor a path has nothing to draw.
  if (!block.path) return [];
  const ext = extname(block.path).toLowerCase();
  let type: 'png' | 'jpg' | 'gif' | 'bmp' = 'png';
  if (ext === '.jpg' || ext === '.jpeg') type = 'jpg';
  else if (ext === '.gif') type = 'gif';
  else if (ext === '.bmp') type = 'bmp';

  let imgPath = block.path;
  if (contentDir && !isAbsolute(imgPath)) {
    imgPath = resolve(contentDir, imgPath);
  }

  const captionStyle = resolveStyle(template, 'figure_caption', docType);
  // Where the template puts its figures. The renderer always centred them; a
  // template whose drawings sit elsewhere said so through its own paragraph.
  const figureAlignment = mapAlignment(template.imageFormat?.alignment) ?? AlignmentType.CENTER;
  const captionSpacing = lineStyle(captionStyle, 312);

  // ImageRun sizes are CSS pixels while the printable width is in twips
  // (1pt = 20 twips, 1px = 0.75pt), so px = twips / 15.
  const page = template.page;
  const printableTwips = Math.max(
    1200,
    (page?.width ?? 11906) - (page?.margins?.left ?? 0) - (page?.margins?.right ?? 0),
  );
  const maxWidthPx = Math.floor(printableTwips / 15);

  if (existsSync(imgPath)) {
    try {
      const buffer = readFileSync(imgPath);
      type = detectImageType(buffer) ?? type;
      const size = getImageSize(buffer);
      
      let width = block.width;
      let height = block.height;

      if (size && size.width > 0 && size.height > 0) {
        // A corrupt header can report a zero dimension; dividing by it put
        // `Infinity` into `w:extent` and produced a package Word refuses to open.
        const ratio = size.width / size.height;
        if (!width && !height) {
          // Fit the printable page width while keeping the aspect ratio.
          width = Math.min(size.width, maxWidthPx);
          height = width / ratio;
        } else if (width && !height) {
          height = width / ratio;
        } else if (!width && height) {
          width = height * ratio;
        }
      } else {
        width = width || Math.min(400, maxWidthPx);
        height = height || 300;
      }

      // A caller-supplied `width` used to reach the document unclamped, so
      // `width: 1200` on a 600 px printable area ran off the right margin. The
      // no-size branch already clamped; every branch does now, and the aspect
      // ratio is preserved when the width is reduced.
      if (width! > maxWidthPx) {
        const shrink = maxWidthPx / width!;
        width = maxWidthPx;
        height = height! * shrink;
      }

      const image = new Paragraph({
        children: [
          new ImageRun({
            data: buffer,
            transformation: {
              width: width!,
              height: height!,
            },
            type,
          }),
        ],
        alignment: figureAlignment,
        spacing: { before: 120, after: 120, line: 312 },
      });
      // 图题 follows the figure and carries the role's own 段前/段后 (`u4`: 10/10
      // for 0.1 行). The run used to be `${number}  ${caption}` with two spaces and
      // a hardcoded 段后 120.
      const caption = new Paragraph({
children: [
          ...(block.number
            ? [
                ...captionNumberRuns(captionIndex.entries.get(block.number)!, captionStyle, nextBookmarkId()),
                styledRun(captionText(block.number, block.caption).slice(block.number.length), captionStyle),
              ]
            : styledRuns(captionText(block.number, block.caption), captionStyle)),
        ],
        alignment: mapAlignment(captionStyle.paragraph.alignment),
        spacing: {
          before: captionStyle.paragraph.spaceBefore,
          after: captionStyle.paragraph.spaceAfter,
          ...captionSpacing,
        },
      });
      return [image, caption];
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : String(error);
      console.warn(`Failed to embed image at ${imgPath}: ${message}`);
    }
  }

  // The placeholder used to say 未找到图片 in every failure case, including the
  // one where the file *was* read and simply could not be embedded (an SVG or
  // WebP reaches here because `detectImageType` returns null and the extension
  // wins). Two different problems deserve two different messages.
  const reason = existsSync(imgPath) ? '无法嵌入图片' : '未找到图片';
  const placeholder = new Paragraph({
    children: [
      new TextRun({
        text: `[${reason}: ${block.path}]`,
        size: captionStyle.font.size,
        font: {
          ascii: captionStyle.font.name,
          hAnsi: captionStyle.font.name,
          eastAsia: captionStyle.font.eastAsia,
          cs: captionStyle.font.name,
        },
        italics: true,
      }),
    ],
    alignment: figureAlignment,
    spacing: { before: 120, after: 120, ...captionSpacing },
  });

  // The caption is emitted even when the image is not, so the figure keeps its
  // number and its bookmark and a `[[图2-1]]` reference still resolves.
  const caption = new Paragraph({
    children: [
      ...(block.number
        ? [
            ...captionNumberRuns(captionIndex.entries.get(block.number)!, captionStyle, nextBookmarkId()),
            styledRun(captionText(block.number, block.caption).slice(block.number.length), captionStyle),
          ]
        : styledRuns(captionText(block.number, block.caption), captionStyle)),
    ],
    alignment: mapAlignment(captionStyle.paragraph.alignment),
    spacing: {
      before: captionStyle.paragraph.spaceBefore,
      after: captionStyle.paragraph.spaceAfter,
      ...captionSpacing,
    },
  });

  return [placeholder, caption];
}

function renderSpacer(block: SpacerBlock): Paragraph[] {
  // An exact height is one empty paragraph whose line box *is* that height.
  // A cover's 5 cm gap is a measurement; five blank lines is only as tall as
  // whatever line height is in force, which is how the guide's spacing got
  // lost in the first place.
  const exact = block.twips ?? (block.cm === undefined ? undefined : convertMillimetersToTwip(block.cm * 10));
  if (exact !== undefined) {
    return [new Paragraph({
      children: [],
      spacing: { line: exact, lineRule: LineRuleType.EXACT },
    })];
  }
  const count = Math.max(0, Math.floor(block.lines ?? 0));
  return Array.from({ length: count }, () =>
    new Paragraph({ spacing: { line: 312 }, children: [] }),
  );
}

function renderPageBreak(): Paragraph {
  return new Paragraph({ children: [new PageBreak()] });
}

function renderHorizontalRule(): Paragraph {
  // Word has no native horizontal-rule paragraph, so use a bottom border.
  return new Paragraph({
    children: [],
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000' } },
  });
}

// ── Official Document Block Renderers ──────────────────────

function renderRedHeader(
  block: RedHeaderBlock,
): Paragraph {
  // GB/T 9704-2012: 发文机关标志用方正小标宋简体，红色，字号自行酌定（以不大于上级机关为原则）
  const size = (block.font_size_pt || 22) * 2; // pt → half-pt，默认22pt
  return new Paragraph({
    children: [
      new TextRun({
        text: block.text,
        size,
        bold: true,
        color: 'FF0000',
        font: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '方正小标宋简体', cs: 'Times New Roman' },
      }),
    ],
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 120, line: 312 },
  });
}

function renderDocumentNumber(
  block: DocumentNumberBlock,
): Paragraph {
  // GB/T 9704-2012: 发文字号用3号仿宋体（16pt）
  return new Paragraph({
    children: [
      new TextRun({
        text: block.text,
        size: 32, // 16pt = 三号
        font: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '仿宋', cs: 'Times New Roman' },
      }),
    ],
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 60, line: 312 },
  });
}

function renderRecipientLine(
  block: RecipientLineBlock,
): Paragraph {
  // GB/T 9704-2012: 主送机关用3号仿宋体（16pt）
  const prefix = block.recipientType === 'cc' ? '抄送：' : '';
  return new Paragraph({
    children: [
      new TextRun({
        text: prefix + block.text,
        size: 32, // 16pt = 三号
        font: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '仿宋', cs: 'Times New Roman' },
      }),
    ],
    alignment: AlignmentType.LEFT,
    spacing: { before: 60, after: 60, line: 312 },
  });
}

function renderSignatureBlock(
  block: SignatureBlockBlock,
): [Paragraph, Paragraph] {
  // GB/T 9704-2012: 发文机关署名和成文日期用3号仿宋体（16pt）
  return [
    new Paragraph({
      children: [
        new TextRun({
          text: block.authority,
          size: 32, // 16pt = 三号
          font: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '仿宋', cs: 'Times New Roman' },
        }),
      ],
      alignment: AlignmentType.RIGHT,
      indent: { right: 720 }, // ~1cm from right
      spacing: { before: 60, after: 0, line: 312 },
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: block.date,
          size: 32, // 16pt = 三号
          font: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '仿宋', cs: 'Times New Roman' },
        }),
      ],
      alignment: AlignmentType.RIGHT,
      indent: { right: 720 }, // ~1cm from right
      spacing: { before: 0, after: 60, line: 312 },
    }),
  ];
}

function renderAttachmentNote(
  block: AttachmentNoteBlock,
): Paragraph[] {
  // GB/T 9704-2012: 附件说明用3号仿宋体（16pt）
  const prefix = '附件：';
  const text = prefix + block.attachments.map((a, i) => `${i + 1}. ${a}`).join('  ');
  return [
    new Paragraph({
      children: [
        new TextRun({
          text,
          size: 32, // 16pt = 三号
          font: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '仿宋', cs: 'Times New Roman' },
        }),
      ],
      alignment: AlignmentType.LEFT,
      spacing: { before: 60, after: 60, line: 312 },
    }),
  ];
}

// ── List numbering ────────────────────────────────────────
//
// Lists used to be plain paragraphs with a literal `•` / `1.` prefix, so Word
// could not renumber them and every ordered item carried the parser's marker.
// These definitions hand numbering back to Word.

const BULLET_NUMBERING_REFERENCE = 'openthesis-bullet';
const ORDERED_NUMBERING_REFERENCE = 'openthesis-ordered';

/** Levels beyond this reuse the deepest definition. */
const MAX_LIST_LEVEL = 2;

const LIST_LEVEL_INDENT = (level: number) => ({
  left: 720 * (level + 1),
  hanging: 360,
});

const LIST_NUMBERING: INumberingOptions['config'] = [
  {
    reference: BULLET_NUMBERING_REFERENCE,
    levels: [
      { level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: LIST_LEVEL_INDENT(0) } } },
      { level: 1, format: LevelFormat.BULLET, text: '○', alignment: AlignmentType.LEFT, style: { paragraph: { indent: LIST_LEVEL_INDENT(1) } } },
      { level: 2, format: LevelFormat.BULLET, text: '▪', alignment: AlignmentType.LEFT, style: { paragraph: { indent: LIST_LEVEL_INDENT(2) } } },
    ],
  },
  {
    reference: ORDERED_NUMBERING_REFERENCE,
    levels: [
      { level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: LIST_LEVEL_INDENT(0) } } },
      { level: 1, format: LevelFormat.LOWER_LETTER, text: '%2.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: LIST_LEVEL_INDENT(1) } } },
      { level: 2, format: LevelFormat.LOWER_ROMAN, text: '%3.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: LIST_LEVEL_INDENT(2) } } },
    ],
  },
];

/** The heading level a section at `depth` renders as. */
const HEADING_TYPES = ['heading1', 'heading2', 'heading3', 'heading4'] as const;

/**
 * The shape `ThesisSection` and `JournalSection` both present to the flattener.
 *
 * The two differ only in their `type` vocabulary, which this helper never reads.
 */
interface NestableSection {
  id: string;
  title: string;
  number?: string;
  content: ContentBlock[];
  subsections?: NestableSection[];
}

/**
 * Flatten a section tree into a flat block list, one heading per section.
 *
 * Written once. It used to exist twice — `flattenSections` and
 * `flattenJournalSections` — character for character identical apart from the
 * section type and the name, each with its own copy of `HEADING_TYPES`.
 */
function flattenSections(sections: readonly NestableSection[], depth = 0): ContentBlock[] {
  const result: ContentBlock[] = [];
  const headingType = HEADING_TYPES[Math.min(depth, HEADING_TYPES.length - 1)];
  for (const section of sections) {
    result.push({
      type: headingType,
      text: section.title,
      number: section.number,
      id: section.id,
    } as ContentBlock);
    result.push(...section.content);
    if (section.subsections && section.subsections.length > 0) {
      result.push(...flattenSections(section.subsections, depth + 1));
    }
  }
  return result;
}

/**
 * Tracks contiguous list runs so each one restarts its own numbering. Without
 * an instance, two separate ordered lists would continue 1,2,3 → 4,5,6.
 */
interface RenderState {
  listInstance: number;
  inList: boolean;
}

function createRenderState(): RenderState {
  return { listInstance: 0, inList: false };
}

function renderListItem(
  block: ListItemBlock,
  instance: number,
  style: ParagraphStyle,
): Paragraph {
  const level = Math.min(Math.max(0, block.level ?? 0), MAX_LIST_LEVEL);
  return new Paragraph({
    children: styledRuns(block.text, style),
    // The marker is generated by Word, so `block.marker` is only a hint for
    // text-only consumers and is intentionally not rendered here.
    numbering: {
      reference: block.ordered ? ORDERED_NUMBERING_REFERENCE : BULLET_NUMBERING_REFERENCE,
      level,
      instance,
    },
    spacing: { after: 60, ...lineStyle(style, 312) },
  });
}

function renderCodeBlock(block: CodeBlockBlock): Paragraph {
  const lines = block.text.split(/\r?\n/);
  return new Paragraph({
    children: lines.map((line, index) => new TextRun({
      text: line || ' ',
      break: index === 0 ? undefined : 1,
      size: 20,
      font: { ascii: 'Courier New', hAnsi: 'Courier New', eastAsia: '等线', cs: 'Courier New' },
    })),
    indent: { left: 360, right: 360 },
    border: {
      top: { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' },
      left: { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' },
      right: { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' },
    },
    spacing: { before: 120, after: 120, line: 276 },
  });
}

function renderBlockquote(block: BlockquoteBlock): Paragraph {
  const quoteStyle: ParagraphStyle = {
    font: { name: 'Times New Roman', eastAsia: '宋体', size: 24, italic: true },
    paragraph: { alignment: 'left' },
    lineSpacing: 312,
  };
  return new Paragraph({
    children: styledRuns(block.text, quoteStyle),
    indent: { left: 480, right: 240 },
    border: { left: { style: BorderStyle.SINGLE, size: 12, color: '808080' } },
    spacing: { before: 120, after: 120, line: 312 },
  });
}

// ── Block Dispatcher ──────────────────────────────────────

/**
 * The core block dispatcher — same pattern as your generate_docx.mjs.
 * Routes each ContentBlock to its renderer based on `type`.
 */
function processBlock(
  block: ContentBlock,
  template: DocumentTemplate,
  contentDir?: string,
  docType?: string,
  state: RenderState = createRenderState(),
): (Paragraph | Table)[] {
  switch (block.type) {
    case 'heading1':
    case 'heading2':
    case 'heading3':
    case 'heading4':
      return [renderHeading(template, block, docType)];

    case 'paragraph':
    case 'paragraph_no_indent':
      return [renderParagraph(template, block, docType)];

    case 'centered_text':
      return [renderCenteredText(template, block, docType)];

    case 'equation':
    case 'equation_numbered':
      return [renderEquation(template, block, docType)];

    case 'table':
      return renderTable(template, block, docType);

    case 'figure':
      return renderFigure(template, block, contentDir, docType);

    case 'spacer':
      return renderSpacer(block);

    case 'section_break':
      // Nothing to draw: `renderThesisParts` already cut the block list here and
      // started a new section. Reaching this means a block slipped past that split.
      return [];

    case 'page_break':
      return [renderPageBreak()];

    case 'horizontal_rule':
      return [renderHorizontalRule()];

    case 'red_header':
      return [renderRedHeader(block)];

    case 'document_number':
      return [renderDocumentNumber(block)];

    case 'recipient_line':
      return [renderRecipientLine(block)];

    case 'signature_block':
      return renderSignatureBlock(block);

    case 'attachment_note':
      return renderAttachmentNote(block);

    case 'list_item':
      return [renderListItem(block, Math.max(1, state.listInstance), resolveStyle(template, 'paragraph', docType))];

    case 'code_block':
      return [renderCodeBlock(block)];

    case 'blockquote':
      return [renderBlockquote(block)];

    default:
      console.warn(`Unknown block type: ${String((block as { type?: unknown }).type)}`);
      return [];
  }
}

// ── Structured Renderers ──────────────────────────────────

/**
 * The parts of a thesis that get a page setup of their own.
 *
 * A thesis is not one run of pages: the cover carries no running head and no
 * page number, the front matter is numbered in Roman numerals and the body
 * restarts at arabic 1. OOXML expresses each of those as a *section*, so the
 * renderer groups the content into these three before laying anything out.
 */
export type DocumentPartRole = 'cover' | 'front_matter' | 'body';

interface RenderedPart {
  role: DocumentPartRole;
  children: (Paragraph | Table)[];
  /** Page setup for this part, when a `section_break` asked for one. */
  page?: { orientation?: 'portrait' | 'landscape'; columns?: number; columnGutter?: number };
}

/**
 * Section types that belong to the front matter, i.e. the pages numbered in
 * Roman numerals before the body restarts at 1: 摘要, ABSTRACT, 目录, 插图和附表
 * 清单, 符号清单. 致谢 and 附录 are *not* here — the university guide numbers them
 * with the body, after 参考文献.
 */
const FRONT_MATTER_TYPES: ReadonlySet<ThesisSection['type']> = new Set([
  'abstract', 'toc', 'list_of_figures', 'list_of_tables',
]);

/** A 目录 is a field, not text — see `tocBlock`. */
/**
 * A paragraph whose only job is to start a new page.
 *
 * The front matter and the back matter are lists of parts that each belong on
 * their own page, but they share one Word section, so the break has to be
 * explicit. The line box is one twip tall, so it costs nothing visible.
 */
function pageBreakParagraph(): Paragraph {
  return new Paragraph({
    children: [],
    pageBreakBefore: true,
    spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT },
  });
}

/** One heading the 目录 must list, in reading order. */
interface HeadingEntry {
  level: number;
  text: string;
  bookmark: string;
}

let headingIndex: HeadingEntry[] = [];
let headingBookmarks = new Map<string, HeadingEntry>();

function resetHeadingIndex(): void {
  headingIndex = [];
  headingBookmarks = new Map();
}

/** Register one heading, returning its entry. */
function registerHeading(level: number, text: string): HeadingEntry {
  const existing = headingBookmarks.get(text);
  if (existing) return existing;
  const entry: HeadingEntry = { level, text, bookmark: `_Toc${90000000 + headingIndex.length + 1}` };
  headingIndex.push(entry);
  headingBookmarks.set(text, entry);
  return entry;
}

/**
 * Walk the document in reading order and register the headings a 目录 lists.
 *
 * A TOC field shows nothing until Word updates it and most people never press
 * F9, so the entries are built at render time. Each entry carries a PAGEREF
 * field, so Word still refreshes the page numbers.
 */
function collectHeadingIndex(doc: OpenThesisDocument): void {
  const skipped = new Set(['toc']);
  const visit = (sections: readonly unknown[]): void => {
    for (const raw of sections) {
      const section = raw as { type?: string; title?: string; subsections?: readonly unknown[] };
      if (!section || !section.type || !section.title) continue;
      if (!skipped.has(section.type)) {
        const level = section.type === 'chapter' ? 1 : 2;
        registerHeading(level, section.title);
      }
      if (section.subsections) visit(section.subsections);
    }
  };
  if (doc.type !== 'thesis') return;
  visit(doc.sections ?? []);
  // The back matter is part of the 目录 too. Its headings are the fixed ones the
  // renderer prints, so they can be listed without looking at any content.
  const back = doc.backMatter;
  if (!back) return;
  if (back.references && back.references.length > 0) registerHeading(1, '参考文献');
  for (const appendix of back.appendices ?? []) {
    if (appendix.title) registerHeading(1, appendix.title);
  }
  if (back.authorBiography) registerHeading(1, '作者简历及在学期间取得的成果');
  if (back.declaration) registerHeading(1, '独创性说明');
}

/** Where a 目录 line's page number sits: the right edge of the text column. */
function tocRightTab(template: DocumentTemplate): number {
  const page = template.page;
  const width = page?.width ?? 11906;
  const left = page?.margins?.left ?? 1701;
  const right = page?.margins?.right ?? 1701;
  return Math.max(1200, width - left - right);
}

/** A page reference to a bookmark: Word fills the number in on update. */
function pageRefField(bookmark: string): XmlComponent {
  return new RawElement('w:r', {}, undefined, [
    new RawElement('w:fldChar', { 'w:fldCharType': 'begin' }),
    new RawElement('w:instrText', { 'xml:space': 'preserve' }, ` PAGEREF ${bookmark} \\h `),
    new RawElement('w:fldChar', { 'w:fldCharType': 'separate' }),
    new RawElement('w:fldChar', { 'w:fldCharType': 'end' }),
  ]);
}

/** One line of a 目录: the heading, a tab with a dot leader, and its page. */
function tocEntryParagraph(entry: HeadingEntry, template: DocumentTemplate): Paragraph {
  const style = resolveStyle(template, 'toc_entry');
  return new Paragraph({
    children: [styledRun(entry.text, style), styledRun('	', style), pageRefField(entry.bookmark)],
    tabStops: [{ type: TabStopType.RIGHT, position: tocRightTab(template), leader: LeaderType.DOT }],
    indent: entry.level > 1 ? { left: (entry.level - 1) * 480 } : undefined,
    spacing: { before: 0, after: 0 },
  });
}

/** One line of a 图表清单: the caption, a tab, and the page it is on. */
function listEntryParagraph(entry: CaptionEntry, template: DocumentTemplate): Paragraph {
  const style = resolveStyle(template, 'toc_entry');
  return new Paragraph({
    children: [
      styledRun(`${entry.label}  ${entry.title}`, style),
      styledRun('	', style),
      pageRefField(captionBookmark(entry.label)),
    ],
    tabStops: [{ type: TabStopType.RIGHT, position: tocRightTab(template), leader: LeaderType.DOT }],
    spacing: { before: 0, after: 0 },
  });
}
function tocBlock(title: string): Paragraph {
  return new TableOfContents(title, { hyperlink: true, headingStyleRange: '1-3' }) as unknown as Paragraph;
}

/**
 * Group a thesis into the parts a page setup can be attached to.
 *
 * The grouping comes from the content, not from a flag, so a content file
 * written before sections existed still lays out correctly: a document with no
 * cover blocks simply has no cover part, and one with no 摘要/目录 has no front
 * matter.
 */
function renderThesisParts(
  doc: ThesisDocument,
  template: DocumentTemplate,
  contentDir?: string,
  state: RenderState = createRenderState(),
): RenderedPart[] {
  const parts: RenderedPart[] = [];
  // Body Sections — flattened by the shared `flattenSections` helper.

  const sections = doc.sections ?? [];
  const frontMatter = sections.filter(section => FRONT_MATTER_TYPES.has(section.type));
  const body = sections.filter(section => !FRONT_MATTER_TYPES.has(section.type));

  // Cover Page. A trailing explicit break is dropped: the section break that
  // follows the cover already starts a new page.
  if (doc.cover && doc.cover.length > 0) {
    parts.push({
      role: 'cover',
      children: processBlocks(trimTrailingPageBreak(doc.cover), template, contentDir, undefined, state),
    });
  }

  if (frontMatter.length > 0) {
    // Document order inside the front matter: the 目录 field keeps its place
    // (摘要, ABSTRACT, 目录 …) rather than being hoisted. Every front-matter
    // section is laid out before the body part anyway, so the 目录 still cannot
    // end up after a chapter.
    const children: (Paragraph | Table)[] = [];
    for (const section of frontMatter) {
      // Each front-matter part begins on its own page: 摘要, Abstract, 目录 and
      // 图表清单 are separate pages in the guide, not one continuous flow.
      if (children.length > 0) children.push(pageBreakParagraph());
      if (section.type === 'toc') {
        children.push(renderHeading(template, { type: 'heading1', text: section.title, styleRole: 'section_heading' }));
        if (headingIndex.length === 0) {
          // Nothing to list: leave the field for Word to fill rather than
          // printing an empty page.
          children.push(tocBlock(section.title));
        } else {
          for (const heading of headingIndex) children.push(tocEntryParagraph(heading, template));
        }
        continue;
      }
      if (section.type === 'list_of_figures' || section.type === 'list_of_tables') {
        const kind = section.type === 'list_of_figures' ? '图' : '表';
        children.push(renderHeading(template, { type: 'heading1', text: section.title, styleRole: 'section_heading' }));
        // A 图表清单 lists the captions, and the caption index already knows every
        // one of them - so build the list instead of printing an empty heading.
        const entries = [...captionIndex.entries.values()].filter(entry => entry.kind === kind);
        if (entries.length === 0) {
          children.push(renderParagraph(template, { type: 'paragraph', text: '（本文无此类内容）' }));
        }
        for (const entry of entries) children.push(listEntryParagraph(entry, template));
        continue;
      }
      children.push(...processBlocks(trimTrailingPageBreak(flattenSections([section])), template, contentDir, undefined, state));
    }
    parts.push({ role: 'front_matter', children });
  }

  // A chapter that turns one passage landscape, or sets it in two columns, is a
  // section boundary inside the body. Each chunk is laid out as its own section
  // with the override its `section_break` declared; the back matter below belongs
  // to the last one, because the guide numbers it with the body.
  // A break outside the body has nothing to cut: the cover, the front matter and
  // the back matter are laid out as their own sections already. Say so rather
  // than dropping it in silence.
  const partsOutsideBody = [
    ...(doc.cover ?? []),
    ...doc.sections.filter(section => !body.includes(section)).flatMap(section => section.content),
    ...(doc.backMatter?.appendices ?? []).flatMap(section => section.content),
  ];
  if (partsOutsideBody.some(block => block.type === 'section_break')) {
    console.warn(
      'A section_break outside the thesis body is ignored: the cover, the front ' +
      'matter and the back matter are already separate sections.',
    );
  }

  const bodyChunks = splitAtSectionBreaks(body.length > 0 ? flattenSections(body) : []);
  const bodyParts: RenderedPart[] = bodyChunks.map(chunk => ({
    role: 'body' as const,
    ...(chunk.page ? { page: chunk.page } : {}),
    children: processBlocks(chunk.blocks, template, contentDir, undefined, state),
  }));
  if (bodyParts.length === 0) bodyParts.push({ role: 'body', children: [] });
  const bodyChildren = bodyParts[bodyParts.length - 1].children;

  // Back Matter. 参考文献 and everything after it belong to the body section:
  // the guide numbers them with the body, not with the front matter.
  if (doc.backMatter?.references && doc.backMatter.references.length > 0) {
    bodyChildren.push(
      pageBreakParagraph(),
      renderHeading(template, { type: 'heading1', text: '参考文献', styleRole: 'section_heading' }),
      ...doc.backMatter.references.map((ref, index) =>
        renderParagraph(template, {
          type: 'paragraph_no_indent',
          text: `[${index + 1}] ${ref.text}`,
        }, undefined, 'reference_item'),
      ),
    );
  }

  if (doc.backMatter?.appendices && doc.backMatter.appendices.length > 0) {
    bodyChildren.push(
      pageBreakParagraph(),
      ...processBlocks(flattenSections(doc.backMatter.appendices), template, contentDir, undefined, state),
    );
  }

  if (doc.backMatter?.authorBiography) {
    bodyChildren.push(
      pageBreakParagraph(),
      renderHeading(template, { type: 'heading1', text: '作者简历及在学期间取得的成果', styleRole: 'section_heading' }),
      renderParagraph(template, { type: 'paragraph', text: doc.backMatter.authorBiography }),
    );
  }

  if (doc.backMatter?.declaration) {
    bodyChildren.push(
      pageBreakParagraph(),
      renderHeading(template, { type: 'heading1', text: '独创性说明', styleRole: 'section_heading' }),
      renderParagraph(template, { type: 'paragraph', text: doc.backMatter.declaration }),
    );
  }

  const nonEmptyBodyParts = bodyParts.filter(part => part.children.length > 0);
  if (nonEmptyBodyParts.length > 0) {
    parts.push(...nonEmptyBodyParts);
  } else if (parts.length === 0) {
    parts.push({ role: 'body', children: [] });
  }

  return parts;
}

/**
 * Cut a block list at its `section_break` markers.
 *
 * Each chunk becomes a section of its own, laid out with the override the
 * break declared. A list with no breaks comes back as one unnamed chunk, so
 * the common case is unchanged.
 */
function splitAtSectionBreaks(
  blocks: ContentBlock[],
): { blocks: ContentBlock[]; page?: RenderedPart['page'] }[] {
  const chunks: { blocks: ContentBlock[]; page?: RenderedPart['page'] }[] = [];
  let current: ContentBlock[] = [];
  let page: RenderedPart['page'];
  for (const block of blocks) {
    if (block.type === 'section_break') {
      chunks.push({ blocks: current, ...(page ? { page } : {}) });
      current = [];
      page = (block as { page?: RenderedPart['page'] }).page;
      continue;
    }
    current.push(block);
  }
  chunks.push({ blocks: current, ...(page ? { page } : {}) });
  return chunks;
}

/**
 * Drop the explicit page break a part ends with.
 *
 * The section break already starts a new page, so a trailing `page_break`
 * prints an empty one between the two — the sample content's cover ends with
 * exactly that break, left over from when the whole document was one section.
 */
function trimTrailingPageBreak(blocks: ContentBlock[]): ContentBlock[] {
  const trimmed = [...blocks];
  while (trimmed.length > 0 && trimmed[trimmed.length - 1].type === 'page_break') {
    trimmed.pop();
  }
  return trimmed;
}

function renderJournalArticle(
  doc: JournalArticle,
  template: DocumentTemplate,
  contentDir?: string,
  state: RenderState = createRenderState(),
): (Paragraph | Table)[] {
  const children: (Paragraph | Table)[] = [];

  // Front matter. Every piece resolves through `resolveStyle` against the
  // template like any other block: this section used to bake its fonts, sizes
  // and spacing in as literals, so a journal rendered from a journal template
  // ignored that template completely.
  const titleStyle = resolveStyle(template, 'journal_title', 'journal');
  children.push(new Paragraph({
    children: styledRuns(doc.meta.title, titleStyle, titleStyle.font.bold),
    alignment: mapAlignment(titleStyle.paragraph.alignment),
    spacing: {
      before: titleStyle.paragraph.spaceBefore,
      after: titleStyle.paragraph.spaceAfter,
      ...lineStyle(titleStyle),
    },
  }));

  // Authors — one run per name, separated by a gap at the role's own size.
  const authorStyle = resolveStyle(template, 'journal_author', 'journal');
  if (doc.meta.authors && doc.meta.authors.length > 0) {
    const authorRuns: TextRun[] = [];
    doc.meta.authors.forEach((author, index) => {
      authorRuns.push(new TextRun({
        text: author.name + (author.isCorresponding ? '*' : ''),
        size: authorStyle.font.size,
        bold: authorStyle.font.bold,
        italics: authorStyle.font.italic,
        font: {
          ascii: authorStyle.font.name,
          hAnsi: authorStyle.font.name,
          eastAsia: authorStyle.font.eastAsia,
          cs: authorStyle.font.name,
        },
      }));
      if (index < doc.meta.authors.length - 1) {
        authorRuns.push(new TextRun({ text: '  ', size: authorStyle.font.size }));
      }
    });
    children.push(new Paragraph({
      children: authorRuns,
      alignment: mapAlignment(authorStyle.paragraph.alignment),
      spacing: {
        before: authorStyle.paragraph.spaceBefore,
        after: authorStyle.paragraph.spaceAfter,
        ...lineStyle(authorStyle),
      },
    }));
  }

  // Affiliations
  const affiliationStyle = resolveStyle(template, 'journal_affiliation', 'journal');
  const allAffiliations: string[] = [];
  doc.meta.authors.forEach(author => {
    author.affiliations.forEach(aff => {
      const affStr = `${aff.institution}${aff.department ? ', ' + aff.department : ''}${aff.city ? ', ' + aff.city : ''}`;
      if (!allAffiliations.includes(affStr)) allAffiliations.push(affStr);
    });
  });
  if (allAffiliations.length > 0) {
    const affText = allAffiliations.map((aff, i) => `(${i + 1}) ${aff}`).join('  ');
    children.push(new Paragraph({
      children: styledRuns(affText, affiliationStyle),
      alignment: mapAlignment(affiliationStyle.paragraph.alignment),
      spacing: {
        before: affiliationStyle.paragraph.spaceBefore,
        after: affiliationStyle.paragraph.spaceAfter,
        ...lineStyle(affiliationStyle),
      },
    }));
  }

  // 摘要 / 关键词. The bold label is part of the block: it takes the role's font
  // but keeps its own weight, as the hardcoded version did.
  const abstractStyle = resolveStyle(template, 'journal_abstract', 'journal');
  if (doc.meta.abstract) {
    children.push(new Paragraph({
      children: [
        ...styledRuns('摘  要：', abstractStyle, true),
        ...styledRuns(doc.meta.abstract, abstractStyle),
      ],
      indent: indentOf(abstractStyle),
      alignment: mapAlignment(abstractStyle.paragraph.alignment),
      spacing: {
        before: abstractStyle.paragraph.spaceBefore,
        after: abstractStyle.paragraph.spaceAfter,
        ...lineStyle(abstractStyle),
      },
    }));
  }

  const keywordsStyle = resolveStyle(template, 'journal_keywords', 'journal');
  if (doc.meta.keywords && doc.meta.keywords.length > 0) {
    children.push(new Paragraph({
      children: [
        ...styledRuns('关键词：', keywordsStyle, true),
        ...styledRuns(doc.meta.keywords.join('；'), keywordsStyle),
      ],
      indent: indentOf(keywordsStyle),
      alignment: mapAlignment(keywordsStyle.paragraph.alignment),
      spacing: {
        before: keywordsStyle.paragraph.spaceBefore,
        after: keywordsStyle.paragraph.spaceAfter,
        ...lineStyle(keywordsStyle),
      },
    }));
  }

  // Divider line or spacer before body
  children.push(new Paragraph({ children: [], spacing: { after: 120 } }));

  // Sections — flattened by the shared `flattenSections` helper.
  //
  // `docType` is passed as 'journal' so the template's own styles reach the body.
  // It used to be `undefined`, which made `resolveStyle` fall back to the thesis
  // defaults and left a journal rendered from a journal template wearing the
  // thesis typography. (Only 公文 opts out of the template, and it says so in
  // `resolveStyle`: GB/T 9704-2012 fixes its typography.)
  children.push(...processBlocks(flattenSections(doc.sections), template, contentDir, 'journal', state));

  // Back Matter
  if (doc.backMatter) {
    if (doc.backMatter.references && doc.backMatter.references.length > 0) {
      children.push(
        renderHeading(template, { type: 'heading1', text: 'References' }),
        ...doc.backMatter.references.map((ref, index) =>
          renderParagraph(template, {
            type: 'paragraph_no_indent',
            text: `[${index + 1}] ${ref.text}`,
          })
        )
      );
    }
    
    if (doc.backMatter.acknowledgments) {
      children.push(
        renderHeading(template, { type: 'heading1', text: 'Acknowledgments' }),
        renderParagraph(template, { type: 'paragraph', text: doc.backMatter.acknowledgments })
      );
    }
  }

  return children;
}

function renderOfficialDocument(
  doc: OfficialDocument,
  template: DocumentTemplate,
  contentDir?: string,
  state: RenderState = createRenderState(),
): (Paragraph | Table)[] {
  const children: (Paragraph | Table)[] = [];

  // Urgency
  if (doc.meta.urgency && doc.meta.urgency !== 'normal') {
    const urgencyText = doc.meta.urgency === 'urgent' ? '急件' : '特急';
    children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: urgencyText,
            size: 24,
            bold: true,
            color: 'FF0000',
            font: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '黑体', cs: 'Times New Roman' },
          }),
        ],
        alignment: AlignmentType.LEFT,
        spacing: { before: 120, after: 120 },
      })
    );
  }

  // Red Header
  children.push(
    renderRedHeader({
      type: 'red_header',
      text: doc.meta.issuingAuthority + '文件',
    })
  );

  // Document Number
  children.push(
    renderDocumentNumber({
      type: 'document_number',
      text: doc.meta.documentNumber,
    })
  );

  // Red horizontal line
  children.push(
    new Paragraph({
      children: [],
      border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: 'FF0000' } },
      spacing: { after: 240 },
    })
  );

  // Title
  // GB/T 9704-2012: 标题用2号方正小标宋简体（22pt）
  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: doc.meta.title,
          size: 44, // 22pt = 二号
          bold: true,
          font: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '方正小标宋简体', cs: 'Times New Roman' },
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 240, after: 240, line: 360 },
    })
  );

  // Recipients
  if (doc.meta.primaryRecipients && doc.meta.primaryRecipients.length > 0) {
    children.push(
      renderRecipientLine({
        type: 'recipient_line',
        text: doc.meta.primaryRecipients.join('、') + '：',
        recipientType: 'primary',
      })
    );
  }

  // Body
  children.push(...processBlocks(doc.body, template, contentDir, 'official', state));

  // Signature Block
  children.push(
    ...renderSignatureBlock({
      type: 'signature_block',
      authority: doc.meta.signatureAuthority || doc.meta.issuingAuthority,
      date: doc.meta.date,
    })
  );

  // Attachments
  if (doc.attachments && doc.attachments.length > 0) {
    children.push(
      ...renderAttachmentNote({
        type: 'attachment_note',
        attachments: doc.attachments.map(att => att.title),
      })
    );
    
    for (const att of doc.attachments) {
      if (att.content && att.content.length > 0) {
        children.push(renderPageBreak());
        // GB/T 9704-2012: 附件标题用3号黑体（16pt）
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: `附件${att.order}：${att.title}`,
                size: 32, // 16pt = 三号
                bold: true,
                font: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '黑体', cs: 'Times New Roman' },
              }),
            ],
            spacing: { before: 240, after: 240, line: 312 },
          })
        );
        children.push(...processBlocks(att.content, template, contentDir, 'official', state));
      }
    }
  }

  // CC Recipients
  if (doc.meta.ccRecipients && doc.meta.ccRecipients.length > 0) {
    children.push(
      new Paragraph({
        children: [],
        border: { top: { style: BorderStyle.SINGLE, size: 6, color: '000000' } },
        spacing: { before: 240 },
      })
    );
    children.push(
      renderRecipientLine({
        type: 'recipient_line',
        text: doc.meta.ccRecipients.join('、') + '。',
        recipientType: 'cc',
      })
    );
  }

  // Publishing Info (版记)
  // GB/T 9704-2012: 版记部分用4号仿宋体（14pt）
  const issuingDept = doc.meta.issuingDepartment || doc.meta.issuingAuthority;
  const issueDate = doc.meta.issueDate || doc.meta.date;
  children.push(
    new Paragraph({
      children: [],
      border: { top: { style: BorderStyle.SINGLE, size: 6, color: '000000' } },
      spacing: { before: 60 },
    })
  );
  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: `${issuingDept}办公厅`,
          size: 28, // 14pt = 四号
          font: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '仿宋', cs: 'Times New Roman' },
        }),
        new TextRun({
          text: `\t\t\t\t\t\t${issueDate}印发`,
          size: 28, // 14pt = 四号
          font: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '仿宋', cs: 'Times New Roman' },
        }),
      ],
      spacing: { before: 60, after: 120 },
    })
  );

  return children;
}

// ── Sections ──────────────────────────────────────────────

/** 五号 (10.5 pt) — the size the university template sets its 篇眉 in. */
const RUNNING_HEAD_SIZE = 21;
const RUNNING_HEAD_FONT = {
  ascii: 'Times New Roman',
  hAnsi: 'Times New Roman',
  eastAsia: '宋体',
  cs: 'Times New Roman',
};

/** The three slots OOXML lets one section reference. */
interface RunningHeadGroup<T> { default?: T; even?: T; first?: T }

/**
 * The template page section that governs a part.
 *
 * - the cover uses the template's first section (its 封一);
 * - the front matter uses the last section that restarts page numbering in
 *   Roman numerals — the 摘要 page. A template usually carries a *continuous*
 *   marker section just before it, which occupies no page of its own, so the
 *   restart that matters is the last one, not the first;
 * - the body uses the template's final section, which is the geometry
 *   `template.page` already exposes.
 *
 * Every part falls back to `template.page`: the university template's eleven
 * sections share one geometry, and a template that declares only a cover
 * section still has only one setup to give.
 */
function partSettings(
  role: DocumentPartRole,
  sections: PageSettings[],
  template: DocumentTemplate,
): PageSettings {
  if (sections.length === 0) return template.page;
  if (role === 'cover') return sections[0];
  if (role === 'body') return template.page;

  const restartsRoman = sections.filter(
    section => section.pageNumberStart === 1 && isRomanFormat(section.pageNumberFormat),
  );
  return restartsRoman[restartsRoman.length - 1]
    ?? sections.find((section, index) => index > 0 && isRomanFormat(section.pageNumberFormat))
    ?? sections[1]
    ?? template.page;
}

function isRomanFormat(format?: PageNumberFormat): boolean {
  return format === 'upperRoman' || format === 'lowerRoman';
}

/**
 * `w:pgNumType` for a part.
 *
 * The template's own declaration wins where there is one — the university
 * template numbers the front matter `upperRoman` and restarts it at I. Where
 * the template is silent, the requirement the guide states applies: the front
 * matter is Roman from I, and the body restarts at 1 in decimal.
 */
function partPageNumbers(
  role: DocumentPartRole,
  settings: PageSettings,
  restart: boolean,
): IPageNumberTypeAttributes {
  if (role === 'cover') {
    // Nothing prints on the cover; the format only decides what the *next*
    // section continues from, so an undeclared one is left undeclared.
    return { ...(settings.pageNumberFormat ? { formatType: settings.pageNumberFormat } : {}) };
  }
  if (role === 'front_matter') {
    return {
      formatType: settings.pageNumberFormat ?? 'upperRoman',
      ...(restart ? { start: settings.pageNumberStart ?? 1 } : {}),
    };
  }
  return {
    formatType: settings.pageNumberFormat ?? 'decimal',
      ...(restart ? { start: settings.pageNumberStart ?? 1 } : {}),
  };
}

/**
 * Size, margins, head/footer distances, the binding gutter and the columns of
 * one part, plus its page numbering. Everything comes from the template; none
 * of it is invented here.
 */
function pageProperties(
  settings: PageSettings,
  pageNumbers: IPageNumberTypeAttributes = {},
): ISectionPropertiesOptions {
  return {
    page: {
      size: {
        // `w:orient` tells Word which way the content reads; a section that
        // swapped its width and height for landscape must say so, or Word keeps
        // printing it as a portrait page with the wrong dimensions.
        orientation: (settings.orientation ?? (settings.width > settings.height ? 'landscape' : 'portrait')) === 'landscape'
          ? PageOrientation.LANDSCAPE
          : PageOrientation.PORTRAIT,
        width: settings.width,
        height: settings.height,
      },
      margin: {
        top: settings.margins.top,
        bottom: settings.margins.bottom,
        left: settings.margins.left,
        right: settings.margins.right,
        // The binding edge: the guide asks for 1 cm on the left.
        ...(settings.gutter ? { gutter: settings.gutter } : {}),
        // How far the head and the page number sit from the page edge.
        ...(settings.headerDistance !== undefined ? { header: settings.headerDistance } : {}),
        ...(settings.footerDistance !== undefined ? { footer: settings.footerDistance } : {}),
      },
      ...(Object.keys(pageNumbers).length > 0 ? { pageNumbers } : {}),
    },
    ...(settings.columns && settings.columns > 1
      ? {
          column: {
            count: settings.columns,
            space: settings.columnGutter ?? 708,
          },
        }
      : {}),
  };
}

/**
 * Fill `{title}`, `{organization}` and `{author}` in a running head.
 *
 * The guide puts the thesis title on the even pages, and a *template* cannot
 * know what that title will be — the university template carries the sample
 * thesis's own title in that part. A template that writes `{title}` there gets
 * the document's title instead; a template that writes a literal gets the
 * literal. Nothing is substituted that was not asked for.
 */
function fillRunningHead(text: string | undefined, variables: Record<string, string | undefined>): string | undefined {
  if (!text || !text.includes('{')) return text;
  return text.replace(/\{(title|organization|author)\}/g, (match, name: string) => variables[name] ?? match);
}

/** One header or footer paragraph: the 篇眉 text, the page number, or both. */
function runningHeadParagraph(spec: RunningHead): Paragraph {
  const runs: TextRun[] = [];

  // A part that draws the page number prints a PAGE field. Any digits beside it
  // in the source part are whatever Word last laid out — the template's footers
  // all carry such a cached number — so they are not content and are not copied.
  if (!spec.pageNumber && spec.text) {
    runs.push(new TextRun({ text: spec.text, size: RUNNING_HEAD_SIZE, font: RUNNING_HEAD_FONT }));
  }
  if (spec.pageNumber) {
    runs.push(new TextRun({ children: [PageNumber.CURRENT], size: RUNNING_HEAD_SIZE, font: RUNNING_HEAD_FONT }));
  }

  return new Paragraph({
    children: runs,
    alignment: AlignmentType.CENTER,
    // The 篇眉 rule: the template's heads draw a 0.5 pt line under the text.
    ...(spec.rule
      ? { border: { bottom: { style: BorderStyle.SINGLE, size: 4, space: 1, color: 'auto' } } }
      : {}),
  });
}

/** The paragraph above, wrapped in the part it belongs to. */
function runningHeadElement<T extends Header | Footer>(kind: 'header' | 'footer', spec: RunningHead): T {
  const paragraph = runningHeadParagraph(spec);
  return (kind === 'header'
    ? new Header({ children: [paragraph] })
    : new Footer({ children: [paragraph] })) as T;
}

/** How one part's running heads are resolved, slot by slot. */
interface RunningHeadOptions {
  /** The generated head/foot a template without header parts falls back to. */
  fallbackText?: string;
  fallbackPageNumber?: boolean;
  /** `RenderOptions.headerText`, which overrides the default slot's text. */
  textOverride?: string;
  /** The document's own title, which overrides the even slot's text. */
  evenTextOverride?: string;
  /** `{title}` and friends, filled from the document being rendered. */
  variables?: Record<string, string | undefined>;
  showPageNumbers: boolean;
}

/**
 * Whether the template asks the renderer to fill a slot in.
 *
 * `{title}`, `{organization}` and `{author}` are requests. Any other text is the
 * *sample* the template was authored with — a university template can only carry
 * the thesis its author happened to be writing — which is evidence of what
 * belongs in the slot, not text to print.
 */
function asksForSubstitution(text: string | undefined): boolean {
  return typeof text === 'string' && text.includes('{');
}

/**
 * The text one slot starts from, before `{...}` is filled.
 *
 * - `default` — the odd pages, once `w:evenAndOddHeaders` is set: the caller's
 *   `headerText` where there is one, else the template's own text.
 * - `even` — the document's own title. The guide puts the thesis title on the
 *   even pages, and the text a template carries there is the *sample* thesis's
 *   title, so it is replaced rather than printed. Two things the template means
 *   literally still win: an empty part is a suppression, and a `{...}` request
 *   already says what the template wants.
 * - `first` — always the template's own text; nothing generates one.
 */
function slotText(
  slot: keyof RunningHeadSlots,
  spec: RunningHead,
  options: RunningHeadOptions,
): string | undefined {
  if (slot === 'default') return options.textOverride ?? spec.text;
  if (slot === 'even' && options.evenTextOverride && spec.text && !asksForSubstitution(spec.text)) {
    return options.evenTextOverride;
  }
  return spec.text;
}

/**
 * The header (or footer) group for one part.
 *
 * Where the template carries header/footer parts, they decide: an *empty* part
 * is a suppression rather than a missing head, which is how the university
 * template keeps the cover head-less (it has no `w:titlePg` anywhere), and an
 * `even` slot is what puts the thesis title on the even pages. That slot's text
 * comes from the document, not from the template — see `slotText`; its
 * formatting (centred, 五号, the 篇眉 rule) still comes from the template.
 *
 * A template that carries no such parts keeps the running head and page number
 * the renderer has always generated — except on the cover, which never gets
 * either, because that is what a cover is for.
 */
function partRunningHeads<T extends Header | Footer>(
  slots: RunningHeadSlots | undefined,
  kind: 'header' | 'footer',
  role: DocumentPartRole,
  options: RunningHeadOptions,
): RunningHeadGroup<T> | undefined {
  if (!slots) {
    if (role === 'cover') return undefined;
    const text = fillRunningHead(options.textOverride ?? options.fallbackText, options.variables ?? {});
    const pageNumber = Boolean(options.fallbackPageNumber) && options.showPageNumbers;
    if (!text && !pageNumber) return undefined;
    return { default: runningHeadElement<T>(kind, { text, pageNumber }) };
  }

  const group: RunningHeadGroup<T> = {};
  for (const slot of ['default', 'even', 'first'] as const) {
    const spec = slots[slot];
    if (!spec) continue;

    const text = fillRunningHead(slotText(slot, spec, options), options.variables ?? {});
    const pageNumber = Boolean(spec.pageNumber) && (kind === 'header' || options.showPageNumbers);
    if (!text && !pageNumber) continue;

    group[slot] = runningHeadElement<T>(kind, { text, pageNumber, rule: spec.rule });
  }

  return Object.keys(group).length > 0 ? group : undefined;
}

/** One `w:sectPr` assembled from a part, plus whether it asks for an even head. */
/**
 * Apply a `section_break` override on top of the geometry a part inherits.
 *
 * Landscape swaps the page's width and height, because that is what the
 * attribute means: `w:orient` only tells Word which way the content reads.
 */
function applyPageOverride(
  settings: PageSettings,
  override?: RenderedPart['page'],
): PageSettings {
  if (!override) return settings;
  const next: PageSettings = { ...settings, margins: { ...settings.margins } };
  // Do not swap width and height here: the writer library does that itself when
  // it is told the page is landscape, and swapping twice cancels out.
  if (override.orientation) next.orientation = override.orientation;
  if (override.columns !== undefined) next.columns = override.columns;
  if (override.columnGutter !== undefined) next.columnGutter = override.columnGutter;
  return next;
}

function buildPartSection(
  part: RenderedPart,
  template: DocumentTemplate,
  sections: PageSettings[],
  options: {
    headerText?: string;
    /** The document's own title: what the even-page running head must read. */
    evenHeaderText?: string;
    variables?: Record<string, string | undefined>;
    showPageNumbers: boolean;
    pageNumberRestart: boolean;
  },
): { section: ISectionOptions; evenReference: boolean } {
  const settings = applyPageOverride(partSettings(part.role, sections, template), part.page);
  const shared = {
    textOverride: options.headerText,
    variables: options.variables,
    showPageNumbers: options.showPageNumbers,
  };

  const headers = partRunningHeads<Header>(settings.headers, 'header', part.role, {
    ...shared,
    evenTextOverride: options.evenHeaderText,
    fallbackText: `${template.meta.organization}学位论文`,
  });
  const footers = partRunningHeads<Footer>(settings.footers, 'footer', part.role, {
    ...shared,
    fallbackPageNumber: true,
  });

  const section: ISectionOptions = {
    properties: {
      ...pageProperties(settings, partPageNumbers(part.role, settings, options.pageNumberRestart)),
      // A `w:type="first"` reference does nothing without `w:titlePg`.
      ...(headers?.first || footers?.first ? { titlePage: true } : {}),
    },
    headers,
    footers,
    children: part.children,
  };

  return { section, evenReference: Boolean(headers?.even || footers?.even) };
}

/**
 * The single-section layout: every block in one section, with the template's
 * final geometry and the generated running head and page number. This is what a
 * template without section information still produces, and what journal and
 * official documents have always used.
 */
function buildSingleSection(
  children: (Paragraph | Table)[],
  template: DocumentTemplate,
  options: { generatedHeader: string; official: boolean; showPageNumbers: boolean },
): ISectionOptions {
  const headers = options.official ? undefined : {
    default: runningHeadElement<Header>('header', { text: options.generatedHeader }),
  };
  const footers = options.showPageNumbers ? {
    default: runningHeadElement<Footer>('footer', { pageNumber: true }),
  } : undefined;

  return {
    properties: pageProperties(template.page),
    headers,
    footers,
    children,
  };
}

// ── Main Render Function ──────────────────────────────────

/**
 * Render any structured OpenThesisDocument to a .docx file using the provided DocumentTemplate.
 * This is the main entry point.
 */
export async function renderDocument(options: RenderOptions): Promise<Buffer> {
  const { template, document: doc, headerText, showPageNumbers = true, contentDir } = options;
  const pageNumberRestart = options.pageNumberRestart ?? true;

  // Start every render from a cold style cache: a caller that mutates the
  // template between renders must see its own edit, and the cache is otherwise
  // keyed only by the template's identity.
  STYLE_CACHE.delete(template);
  resetBookmarks();
  resetCaptionIndex();

  // Shared across the whole document so list numbering instances stay unique.
  const state = createRenderState();

  // The running head a template without header parts falls back to.
  // Captions are numbered before anything renders: a `[[表#id]]` reference needs
  // the number that the caption will carry, and it may appear above the caption.
  numberCaptionsInDocument(doc);
  // The 目录 is built before the body renders for the same reason: its entries
  // are the headings, and the headings need their bookmarks assigned up front.
  resetHeadingIndex();
  collectHeadingIndex(doc);

  const generatedHeader = headerText
    || (doc.type === 'official' ? '' : `${template.meta.organization}${doc.type === 'journal' ? '学术论文' : '学位论文'}`);

  const templateSections = template.pageSections ?? [];
  const sections: ISectionOptions[] = [];
  let evenReference = false;

  // What a `{title}`-style placeholder in a template's running head resolves to.
  const headVariables: Record<string, string | undefined> = {
    title: doc.meta.title,
    organization: template.meta.organization,
    author: 'author' in doc.meta ? (doc.meta as { author?: string }).author : undefined,
  };

  if (doc.type === 'thesis') {
    const parts = renderThesisParts(doc, template, contentDir, state);

    // A template that declares several sections is what the cover / front matter
    // / body split is laid out against. One that declares none — an older parsed
    // JSON, or a hand-written template — keeps the single section it always had.
    // A part that carries its own page setup forces per-part sections even when
    // the template declares only one: that is what a mid-chapter `section_break`
    // is asking for.
    if (templateSections.length > 1 || parts.some(part => part.page)) {
      for (const part of parts) {
        const built = buildPartSection(part, template, templateSections, {
          headerText,
          // The even-page running head is the thesis's own title. A template's
          // part can only carry its author's sample title, which must not print.
          evenHeaderText: doc.meta.title,
          variables: headVariables,
          pageNumberRestart,
          showPageNumbers,
        });
        sections.push(built.section);
        evenReference = evenReference || built.evenReference;
      }
    } else {
      sections.push(buildSingleSection(
        parts.flatMap(part => part.children),
        template,
        { generatedHeader, official: false, showPageNumbers },
      ));
    }
  } else if (doc.type === 'journal') {
    sections.push(buildSingleSection(
      renderJournalArticle(doc, template, contentDir, state),
      template,
      { generatedHeader, official: false, showPageNumbers },
    ));
  } else if (doc.type === 'official') {
    sections.push(buildSingleSection(
      renderOfficialDocument(doc, template, contentDir, state),
      template,
      { generatedHeader, official: true, showPageNumbers },
    ));
  } else {
    const type = (doc as { type?: unknown }).type;
    throw new Error(`Unsupported document type: "${String(type)}". Expected 'thesis', 'journal', or 'official'.`);
  }

  const wordDoc = new Document({
    numbering: { config: [...LIST_NUMBERING, ...templateNumberingConfig(template)] },
    // The 目录 is a field; Word builds its entries with the TOC1/TOC2/TOC3
    // styles, so those must exist in the package or the template's own directory
    // typography is silently ignored.
    styles: { paragraphStyles: tocEntryStyles(template) },
    sections,
  });

  let buffer = await Packer.toBuffer(wordDoc);

  // ── Post-process: fix Chinese fonts (WPS compatibility) ──────
  // The docx library has a bug where font.eastAsia in styles.default
  // gets overwritten with the ascii font name. WPS then shows all text
  // in Times New Roman. We fix this by post-processing the zip:
  // 1. Replace all *Theme attributes in styles.xml with explicit fonts
  // 2. Fix docDefaults eastAsia font
  // 偶数页页眉 is written only when the output really asks for an even-page
  // head: the flag also demotes `w:type="default"` to odd pages, so setting it
  // while emitting no even reference would leave the even pages' head blank.
  const templateCarriesHeads = templateSections.some(section => section.headers || section.footers);
  buffer = await postProcess(buffer, {
    mirrorMargins: Boolean(template.page.mirrorMargins),
    evenAndOddHeaders: evenReference || (Boolean(template.evenAndOddHeaders) && templateCarriesHeads),
    docGrid: template.page.docGrid,
  });

  if (options.outputPath) {
    // `-o a/b/c.docx` should behave like `mkdir -p` rather than fail with an
    // ENOENT raised deep inside the writer.
    mkdirSync(dirname(options.outputPath), { recursive: true });
    writeFileSync(options.outputPath, buffer);
  }

  return buffer;
}

/**
 * The single post-process pass over the packed package.
 *
 * Two things have to happen after `Packer.toBuffer`, and they used to be two
 * separate functions that each unzipped and re-zipped the whole document. The
 * second pass generated without a `compression` option, so the package that
 * reached disk was the stored one — 188.9 KB for 136.8 KB of content. Loading
 * once, applying every patch, and generating once with DEFLATE removes the
 * redundant pass and the size regression.
 *
 * 1. Chinese fonts (WPS compatibility). The `docx` library writes
 *    `w:eastAsia="Times New Roman"` into docDefaults when it was given
 *    `font: { ascii: 'Times New Roman', eastAsia: '宋体' }` — it uses the ascii
 *    name for all four slots. WPS honours docDefaults, so Chinese text
 *    inheriting from it renders in Times New Roman. Stripping every `*Theme`
 *    attribute and rewriting a Latin `w:eastAsia` to 仿宋 fixes the inherited
 *    runs; run-level eastAsia values (宋体/黑体/仿宋) are left alone and now take
 *    effect. The theme part itself stays, so its relationship remains valid.
 *
 * 2. Document settings the writer library cannot emit: `w:mirrorMargins`,
 *    `w:evenAndOddHeaders` and `w:updateFields`. All three are children of
 *    `w:settings` in `word/settings.xml` — none is a legal child of `w:sectPr`,
 *    so writing them into document.xml produces a part Word ignores or repairs,
 *    and a check that only greps for the string cannot tell the difference.
 *    `updateFields` matters because the 目录 is a TOC field whose page numbers
 *    only exist once Word evaluates it.
 *
 * Idempotent: a package that needs no patch is returned as the original buffer.
 */
async function postProcess(
  buffer: Buffer,
  options: {
    mirrorMargins: boolean;
    evenAndOddHeaders: boolean;
    /** `w:docGrid` from the template's section properties, when it declares one. */
    docGrid?: { type?: string; linePitch?: number; charSpace?: number };
  },
): Promise<Buffer> {
  const zip = await JSZip.loadAsync(buffer);
  let changed = false;

  // 1. Chinese fonts. The writer library writes the ascii font name into
  //    `w:docDefaults/w:rPrDefault/w:rPr/w:rFonts/@w:eastAsia` — the same value as
  //    `w:ascii` — and WPS honours docDefaults, so any run that inherits from it
  //    renders Chinese in Times New Roman.
  //
  //    Status as of docx 9.7.1: the emitted docDefaults is
  //    `<w:docDefaults><w:rPrDefault/><w:pPrDefault/></w:docDefaults>` — no rFonts at
  //    all, because this renderer never asks the library for a default document run.
  //    The patch is therefore inert today; it stays as a guard because the symptom
  //    it addresses was real and a library bump could reintroduce it, and because
  //    staying narrow costs nothing.
  //
  //    Only that element is rewritten. Scoping the patch to docDefaults matters:
  //    the previous version stripped every `*Theme` attribute and rewrote every
  //    Latin `w:eastAsia` in both styles.xml and document.xml, which is harmless
  //    for a template that says nothing but overrides one that says exactly what
  //    it wants. A template's own run-level fonts now survive untouched.
  const stylesPart = zip.file('word/styles.xml');
  if (stylesPart) {
    const original = await stylesPart.async('string');
    const patched = original.replace(
      /(<w:docDefaults>)[\s\S]*?(<\/w:docDefaults>)/,
      (_match, open: string, close: string) => {
        const body = _match.slice(open.length, _match.length - close.length);
        const cleaned = body
          .replace(/\s+w:(asciiTheme|hAnsiTheme|eastAsiaTheme|cstheme)="[^"]*"/g, '')
          .replace(/\sw:eastAsia="(Times New Roman|Arial|Calibri|Cambria)"/g, ' w:eastAsia="仿宋"');
        return open + cleaned + close;
      },
    );
    if (patched !== original) {
      zip.file('word/styles.xml', patched);
      changed = true;
    }
  }

  // 2. Document settings.
  const settings = zip.file('word/settings.xml');
  if (settings) {
    const original = await settings.async('string');
    let xml = original;

    // Ask Word to refresh fields when the document opens, so the 目录 and its
    // page numbers are right on the first look rather than after a manual F9.
    if (!xml.includes('<w:updateFields')) {
      xml = xml.replace(/(<w:settings[^>]*>)/, '$1<w:updateFields w:val="true"/>');
    }

    if (options.mirrorMargins && !xml.includes('<w:mirrorMargins')) {
      xml = xml.replace(/(<w:settings[^>]*>)/, '$1<w:mirrorMargins/>');
    }

    // 偶数页页眉. The writer library always emits this element — with
    // `w:val="false"` when it was not asked for — so a document that wants it
    // has to rewrite the value. Inserting a second element instead would leave
    // Word reading the first one, which is `false`.
    if (options.evenAndOddHeaders) {
      xml = xml.includes('<w:evenAndOddHeaders')
        ? xml.replace(/<w:evenAndOddHeaders[^>]*\/>/, '<w:evenAndOddHeaders/>')
        : xml.replace(/(<w:settings[^>]*>)/, '$1<w:evenAndOddHeaders/>');
    }

    if (xml !== original) {
      zip.file('word/settings.xml', xml);
      changed = true;
    }
  }

  // 3. 文档网格. The writer library always emits its own `linePitch="360"`.
  // A template that snaps text to a different grid needs its own values in the
  // generated section properties, or a page breaks lines where the template
  // would not.
  if (options.docGrid && (options.docGrid.type || options.docGrid.linePitch !== undefined || options.docGrid.charSpace !== undefined)) {
    const document = zip.file('word/document.xml');
    if (document) {
      const original = await document.async('string');
      const gridAttributes = [
        options.docGrid.type ? `w:type="${options.docGrid.type}"` : '',
        options.docGrid.linePitch !== undefined ? `w:linePitch="${options.docGrid.linePitch}"` : '',
        options.docGrid.charSpace !== undefined ? `w:charSpace="${options.docGrid.charSpace}"` : '',
      ].filter(Boolean).join(' ');
      const patched = original.includes('<w:docGrid')
        ? original.replace(/<w:docGrid[^>]*\/>/g, `<w:docGrid ${gridAttributes}/>`)
        : original;
      if (patched !== original) {
        zip.file('word/document.xml', patched);
        changed = true;
      }
    }
  }

  if (!changed) return buffer;
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

// ── Legacy API (backward compat with existing JSON format) ──

/**
 * Render using the legacy flat JSON format (backward compatible).
 * This wraps the old cover_blocks/body_blocks into a ThesisDocument.
 */
export async function renderLegacy(
  legacy: LegacyDocumentJSON,
  template: DocumentTemplate,
  outputPath: string,
  contentDir: string = dirname(outputPath),
): Promise<Buffer> {
  const doc: ThesisDocument = {
    type: 'thesis',
    meta: { title: legacy.title },
    cover: legacy.cover_blocks,
    sections: [
      {
        id: 'body',
        type: 'chapter',
        title: legacy.title,
        content: legacy.body_blocks,
      },
    ],
  };

  return renderDocument({
    outputPath,
    template,
    document: doc,
    headerText: legacy.report_header,
    contentDir,
  });
}

/**
 * Paragraph styles for the table-of-contents entries Word generates.
 *
 * The renderer emits a real `TOC` field, and Word fills it with one paragraph
 * per heading, styled `TOC1` / `TOC2` / `TOC3`. The generated package carried
 * none of those style ids, so the entries fell back to Word's own defaults even
 * when the template had parsed its `TOC1` style and mapped it to the `toc_entry`
 * role — the role was parsed and then never requested.
 */
function tocEntryStyles(template: DocumentTemplate): IParagraphStyleOptions[] {
  const styles: IParagraphStyleOptions[] = [];
  for (const level of [1, 2, 3] as const) {
    const id = `TOC${level}`;
    const source = template.styles?.[id];
    if (!source) continue;
    const line = lineStyle(source);
    styles.push({
      id,
      name: `TOC ${level}`,
      quickFormat: true,
      run: {
        size: source.font.size,
        bold: source.font.bold,
        italics: source.font.italic,
        font: {
          ascii: source.font.name,
          hAnsi: source.font.name,
          eastAsia: source.font.eastAsia,
          cs: source.font.name,
        },
      },
      paragraph: {
        alignment: mapAlignment(source.paragraph.alignment),
        indent: indentOf(source),
        spacing: {
          before: source.paragraph.spaceBefore,
          after: source.paragraph.spaceAfter,
          ...line,
        },
      },
    });
  }
  return styles;
}

/**
 * The template's own heading numbering, as the writer library wants it.
 *
 * The template binds its heading styles to `w:num` instances through
 * `w:numPr`, and each instance points at an `w:abstractNum` whose `w:lvlText`
 * says what the number looks like (`%1`, `%1.%2`, `附录 %1`). None of that was
 * read before, so a heading's number could only come from the content file.
 */
const LEVEL_FORMATS: Record<string, (typeof LevelFormat)[keyof typeof LevelFormat]> = {
  decimal: LevelFormat.DECIMAL,
  upperRoman: LevelFormat.UPPER_ROMAN,
  lowerRoman: LevelFormat.LOWER_ROMAN,
  upperLetter: LevelFormat.UPPER_LETTER,
  lowerLetter: LevelFormat.LOWER_LETTER,
  chineseCounting: LevelFormat.CHINESE_COUNTING,
  chineseCountingThousand: LevelFormat.CHINESE_COUNTING_THOUSAND,
  decimalEnclosedCircle: LevelFormat.DECIMAL_ENCLOSED_CIRCLE,
  none: LevelFormat.NONE,
};

function mapLevelFormat(format?: string): (typeof LevelFormat)[keyof typeof LevelFormat] {
  return (format && LEVEL_FORMATS[format]) || LevelFormat.DECIMAL;
}

/** The numbering reference a template `numId` is emitted under. */
function templateNumberingReference(numId: number): string {
  return `template-num-${numId}`;
}

function templateNumberingConfig(template: DocumentTemplate): INumberingOptions['config'][number][] {
  const config: INumberingOptions['config'][number][] = [];
  for (const definition of template.numbering ?? []) {
    // `w:numId="0"` is OOXML for "no numbering", not an instance.
    if (definition.numId === 0 || definition.levels.length === 0) continue;
    config.push({
      reference: templateNumberingReference(definition.numId),
      levels: definition.levels.map(level => ({
        level: level.ilvl,
        format: mapLevelFormat(level.format),
        text: level.text ?? `%${level.ilvl + 1}.`,
        start: level.start ?? 1,
        alignment: AlignmentType.LEFT,
      })),
    });
  }
  return config;
}

/**
 * Map an OOXML `w:val` border style to the writer library's enum.
 *
 * The two vocabularies are nearly identical; an unknown value falls back to a
 * single rule rather than dropping the border, because a table that declared a
 * border and got none reads as a broken template.
 */
function borderStyleOf(value: string): (typeof BorderStyle)[keyof typeof BorderStyle] {
  const known = BorderStyle as unknown as Record<string, (typeof BorderStyle)[keyof typeof BorderStyle]>;
  const candidate = value.charAt(0).toUpperCase() + value.slice(1);
  return known[candidate] ?? BorderStyle.SINGLE;
}

/**
/**
 * A cross-reference to a caption, as Word understands one.
 *
 * Writing `如表3-1所示` into a paragraph makes the number plain text: renumber
 * the table and the sentence is wrong. The standard mechanism is a `REF` field
 * pointing at a bookmark around the caption's number, so Word recomputes it.
 * The bookmark sits on the caption; the sentence carries the field.
 *
 * Markdown reaches this through `[[表3-1]]` / `[[图2-1]]`, which `styledRuns`
 * turns into the field. A manuscript that writes the number out longhand keeps
 * working exactly as before — its text is left alone.
 */
const CROSS_REF = /\[\[(表|图)\s*([^\[\]]+?)\]\]/g;

/** The bookmark name a caption's number lives under. */
function captionBookmark(label: string): string {
  return `_Ref${label.replace(/[^0-9A-Za-z\u4e00-\u9fa5]/g, '')}`;
}

/**
 * A raw OOXML element the writer library has no wrapper for.
 *
 * `XmlComponent`'s `root` follows the library's own convention: an optional
 * prefix, then an attribute bag, then either a text payload or child elements.
 * A string in the third slot is the element's text.
 */
class RawElement extends XmlComponent {
  // The base class keeps `root` protected; this subclass assigns it.
  declare protected root: (XmlComponent | string | any)[];

  constructor(name: string, attributes: Record<string, string>, text?: string, children?: XmlComponent[]) {
    super(name);
    // Three cases, and the distinction matters: an element with children, an
    // element with text, and a self-closing element. Passing `''` for the last
    // one emits `<w:t></w:t>` — a text element with no text — which is not what
    // a bookmark marker is.
    this.root = children
      ? (['', { _attr: attributes }, ...children] as (XmlComponent | string | any)[])
      : text === undefined
        ? (['', { _attr: attributes }] as (XmlComponent | string | any)[])
        : (['', { _attr: attributes }, text] as (XmlComponent | string | any)[]);
  }

}
/**
 * The caption label Word knows this figure or table by.
 *
 * A Word caption is a `SEQ` field whose identifier names the label — `SEQ Figure`
 * for 图, `SEQ Table` for 表. Word's Insert Caption and Cross-reference dialogs
 * only recognise a caption label that matches a SEQ identifier, so writing the
 * number as a plain string (which is what this renderer did) produces text that
 * looks right and cannot be referenced, renumbered, or listed.
 */
function captionSeqLabel(kind: '表' | '图'): string {
  return kind === '表' ? 'Table' : 'Figure';
}

/** A bookmark around a caption's number, so a REF field can point at it. */
function captionBookmarkStart(name: string, id: number): XmlComponent {
  return new RawElement('w:bookmarkStart', { 'w:id': String(id), 'w:name': name });
}

function captionBookmarkEnd(id: number): XmlComponent {
  return new RawElement('w:bookmarkEnd', { 'w:id': String(id) });
}

/**
 * A `SEQ` field: Word's own caption numbering.
 *
 * `\s 1` restarts the sequence at each heading of the given outline level, which
 * is what makes 图 1-1 / 图 2-1 work without any chapter number written into the
 * text. `\* ARABIC` asks for plain digits.
 *
 * The field carries a **materialised result**: viewers that do not evaluate
 * fields (LibreOffice's renderer, some WPS builds) then show the number the
 * renderer computed instead of a blank. Word recalculates it on update.
 */
function captionSeqField(kind: '表' | '图', result: string, restartAtHeading: boolean): XmlComponent {
  const seq = captionSeqLabel(kind);
  const instr = ` SEQ ${seq} \\* ARABIC${restartAtHeading ? ' \\s 1' : ''} `;
  return new RawElement('w:r', {}, undefined, [
    new RawElement('w:fldChar', { 'w:fldCharType': 'begin' }),
    new RawElement('w:instrText', { 'xml:space': 'preserve' }, instr),
    new RawElement('w:fldChar', { 'w:fldCharType': 'separate' }),
    new RawElement('w:t', { 'xml:space': 'preserve' }, result),
    new RawElement('w:fldChar', { 'w:fldCharType': 'end' }),
  ]);
}

/**
 * The runs that make up a caption's number, as Word builds one.
 *
 * `表` and `-` are literal text, the chapter number is literal, and the sequence is
 * a `SEQ Table` field — the part Word actually renumbers. A caption the author
 * numbered by hand keeps its literal text instead.
 */
function captionNumberRuns(
  entry: CaptionEntry,
  style: ParagraphStyle,
  bookmarkId: number,
): XmlComponent[] {
  const runs: XmlComponent[] = [captionBookmarkStart(captionBookmark(entry.label), bookmarkId)];

  if (entry.manual) {
    // The author's own label, kept as text: a SEQ field here would renumber it.
    runs.push(styledRun(entry.label, style, true));
    runs.push(captionBookmarkEnd(bookmarkId));
    return runs;
  }

  runs.push(styledRun(entry.kind, style, true));
  if (entry.chapter > 0) {
    // Literal, not a STYLEREF field: a heading's number exists as a field only
    // when the heading is auto-numbered, and these titles carry theirs in their
    // text, so a STYLEREF here renders as "Error: Reference source not found".
    // The chapter number is computed at build time and is right for the document
    // as written; the sequence within the chapter is what Word renumbers.
    runs.push(styledRun(String(entry.chapter), style, true));
    runs.push(styledRun('-', style, true));
  }
  runs.push(captionSeqField(entry.kind, String(entry.seq), entry.chapter > 0));
  runs.push(captionBookmarkEnd(bookmarkId));
  return runs;
}

/**
 * A `REF` field that Word evaluates to the caption's number.
 *
 * `\h` makes the reference a hyperlink to the bookmark. The field carries the
 * number as its result, so a viewer that does not evaluate fields still shows it
 * rather than an empty gap.
 */
/**
 * Turn a reference token into the caption bookmark it points at, plus the
 * number to show as the field's result.
 *
 * `表#tab1` names a block by id, which is the form that survives automatic
 * numbering. A literal `表3-1` is taken at face value. An unknown label is
 * reported: a REF field pointing at a bookmark that does not exist makes Word
 * print "Error! Reference source not found."
 */
function resolveReference(raw: string): { bookmark: string; result: string } | undefined {
  const label = raw.trim();
  const entry = label.includes('#')
    ? captionIndex.byId.get(label)
    : captionIndex.entries.get(label);
  if (!entry) {
    console.warn(`Cross-reference to "${label}" matches no caption in this document; the reference will be left as text.`);
    return undefined;
  }
  return { bookmark: captionBookmark(entry.label), result: entry.label };
}

function crossReferenceField(bookmark: string, result: string): XmlComponent {
  return new RawElement('w:r', {}, undefined, [
    new RawElement('w:fldChar', { 'w:fldCharType': 'begin' }),
    new RawElement('w:instrText', { 'xml:space': 'preserve' }, ` REF ${bookmark} \\h `),
    new RawElement('w:fldChar', { 'w:fldCharType': 'separate' }),
    new RawElement('w:t', { 'xml:space': 'preserve' }, result),
    new RawElement('w:fldChar', { 'w:fldCharType': 'end' }),
  ]);
}

let bookmarkCounter = 0;

/** A fresh bookmark id. Word only needs them unique within the document. */
function nextBookmarkId(): number {
  bookmarkCounter += 1;
  return bookmarkCounter;
}

/** Reset between renders, so two documents do not share a numbering. */
function resetBookmarks(): void {
  bookmarkCounter = 0;
}

/**
 * Every caption in the document, and the numbers they ended up with.
 *
 * A caption with no `number` used to render with no number at all, which made it
 * unreachable from a cross-reference. Captions are numbered here instead, per
 * chapter, in document order — the standard 表1-1 / 图2-3 form — and a reference
 * can name either that label (`[[表1-1]]`) or the block's id (`[[表#tab1]]`).
 */
/**
 * Every caption in the document, and the structure Word needs to number it.
 *
 * A Word caption is not a string: it is `表` + a `STYLEREF` field for the chapter
 * number + `-` + a `SEQ Table` field. That is what makes 图 2-1 both display
 * correctly and stay updatable — and what makes it appear in Word's
 * Cross-reference dialog at all. So the renderer records the parts, not just the
 * finished label.
 */
interface CaptionEntry {
  kind: '表' | '图';
  chapter: number;
  seq: number;
  /** The label as it reads, e.g. `表2-1`. */
  label: string;
  /** The caption's own words, without the number. */
  title: string;

  /** True when the author wrote the number, so it is literal text, not a field. */
  manual: boolean;
}

interface CaptionIndex {
  entries: Map<string, CaptionEntry>;
  byId: Map<string, CaptionEntry>;
}

let captionIndex: CaptionIndex = { entries: new Map(), byId: new Map() };



function resetCaptionIndex(): void {
  captionIndex = { entries: new Map(), byId: new Map() };
}

type NumberedBlock = { type?: string; number?: string; id?: string; caption?: string };

/** Number every table and figure in one block list that brought no number. */
/** Number every table and figure in one block list that brought no number. */
function numberCaptions(blocks: NumberedBlock[], chapter: number, counters: Map<string, number>): void {
  for (const block of blocks) {
    const kind = block.type === 'table' ? '表' : block.type === 'figure' ? '图' : undefined;
    if (!kind) continue;

    const key = `${chapter}|${kind}`;
    const seq = (counters.get(key) ?? 0) + 1;
    counters.set(key, seq);

    // A number the author wrote wins and keeps its literal text: it is outside
    // the sequence Word manages, so giving it a SEQ field would renumber it out
    // from under them. Everything else is Word's to number.
    const written = block.number;
    const entry: CaptionEntry = written
      ? { kind, chapter, seq, label: written, manual: true, title: block.caption ?? '' }
      : { kind, chapter, seq, label: chapter > 0 ? `${kind}${chapter}-${seq}` : `${kind}${seq}`, manual: false, title: block.caption ?? '' };
    block.number = entry.label;

    captionIndex.entries.set(entry.label, entry);
    if (block.id) captionIndex.byId.set(`${kind}#${block.id}`, entry);
  }
}

/**
 * Walk the whole document in reading order and number its captions.
 *
 * Only sections that are chapters advance the chapter number; a 摘要 or 目录
 * section between them must not turn 表1-1 into 表2-1.
 */
function numberCaptionsInDocument(doc: OpenThesisDocument): void {
  const counters = new Map<string, number>();
  if (doc.type === 'thesis') {
    numberCaptions((doc.cover ?? []) as NumberedBlock[], 0, counters);
    let chapter = 0;
    for (const section of doc.sections ?? []) {
      if (section.type === 'chapter') chapter += 1;
      numberCaptions(section.content as NumberedBlock[], chapter, counters);
      for (const sub of section.subsections ?? []) {
        numberCaptions(sub.content as NumberedBlock[], chapter, counters);
      }
    }
    for (const appendix of doc.backMatter?.appendices ?? []) {
      numberCaptions(appendix.content as NumberedBlock[], chapter, counters);
    }
    return;
  }
  if (doc.type === 'journal') {
    let chapter = 0;
    for (const section of doc.sections ?? []) {
      chapter += 1;
      numberCaptions(section.content as NumberedBlock[], chapter, counters);
    }
    return;
  }
  numberCaptions((doc as { body?: NumberedBlock[] }).body ?? [], 0, counters);
}


/**
 * The border value that draws nothing.
 *
 * `nil` is spec-correct and Word honours it, but LibreOffice — and viewers built
 * on it — falls back to its default table grid, ruling every row and column of a
 * three-line table. A white hairline draws nothing on paper and works everywhere.
 */
const NO_BORDER = { style: BorderStyle.SINGLE, size: 2, color: 'FFFFFF' };

// ── Utility ───────────────────────────────────────────────

/**
 * A schema alignment to a writer-library one, or `undefined` for "not declared".
 *
 * Returning `JUSTIFIED` by default made a paragraph whose style declares no
 * alignment render justified, even when Word's own default for it would be left
 * — a heading, for instance. `undefined` lets the library omit `w:jc` and Word
 * apply its default.
 */
function mapAlignment(align?: string): (typeof AlignmentType)[keyof typeof AlignmentType] | undefined {
  switch (align) {
    case 'center': return AlignmentType.CENTER;
    case 'left': return AlignmentType.LEFT;
    case 'right': return AlignmentType.RIGHT;
    case 'justified': return AlignmentType.JUSTIFIED;
    case 'distribute': return AlignmentType.DISTRIBUTE;
    default: return undefined;
  }
}

/**
 * Process an array of blocks (for flattening section trees)
 */
function processBlocks(
  blocks: ContentBlock[],
  template: DocumentTemplate,
  contentDir?: string,
  docType?: string,
  state: RenderState = createRenderState(),
): (Paragraph | Table)[] {
  const result: (Paragraph | Table)[] = [];
  for (const block of blocks) {
    // A run of consecutive list items is one list; anything else ends it, so the
    // next list starts a fresh numbering instance.
    if (block.type === 'list_item') {
      if (!state.inList) {
        state.listInstance += 1;
        state.inList = true;
      }
    } else {
      state.inList = false;
    }
    result.push(...processBlock(block, template, contentDir, docType, state));
  }
  return result;
}

// ── Quick-start helpers ───────────────────────────────────

/** Path to the parsed USTB template JSON (relative to project root). */
const USTB_TEMPLATE_JSON = resolve(_dirname, '..', '..', '..', 'assets', 'ustb-thesis-template.json');

/** The parsed built-in template, loaded at most once per process. */
let USTB_TEMPLATE_CACHE: DocumentTemplate | undefined;

/**
 * Create a USTB thesis template, loading from the pre-parsed .docx template.
 * Falls back to a minimal hardcoded template if the JSON file is unavailable.
 */
export function createUSTBTemplate(): DocumentTemplate {
  // Parsed once per process: callers build the same template repeatedly, and it
  // used to re-read and re-parse a 1,300-line JSON on every call.
  if (USTB_TEMPLATE_CACHE) return USTB_TEMPLATE_CACHE;

  try {
    if (existsSync(USTB_TEMPLATE_JSON)) {
      const raw = JSON.parse(readFileSync(USTB_TEMPLATE_JSON, 'utf-8'));
      USTB_TEMPLATE_CACHE = raw as DocumentTemplate;
      return USTB_TEMPLATE_CACHE;
    }
  } catch (error: unknown) {
    // A moved or corrupt asset used to fall through silently, producing a
    // document styled by the minimal hand-written template with no hint that
    // anything was wrong. Say so, then fall back.
    const message = error instanceof Error ? error.message : String(error);
    console.warn(
      `Could not read the built-in USTB template at ${USTB_TEMPLATE_JSON} (${message}); ` +
      'falling back to the minimal built-in styles.',
    );
  }

  // Minimal hardcoded fallback (used when the parsed JSON is not available)
  USTB_TEMPLATE_CACHE = {
    meta: {
      organization: '北京科技大学',
      name: '博士/硕士学位论文模板',
      documentType: 'thesis' as const,
      parserVersion: '0.1.0',
      parsedAt: new Date().toISOString(),
    },
    page: {
      width: convertMillimetersToTwip(210),
      height: convertMillimetersToTwip(297),
      margins: {
        top: convertMillimetersToTwip(30),
        bottom: convertMillimetersToTwip(20),
        left: convertMillimetersToTwip(30),
        right: convertMillimetersToTwip(30),
      },
    },
    styles: {
      ...THESIS_STYLES,
    },
    styleRoles: {
      'heading1': 'heading1',
      'heading2': 'heading2',
      'heading3': 'heading3',
      'heading4': 'heading4',
      'Normal': 'paragraph',
      'CoverTitle': 'centered_text',
      'Equation': 'equation',
      'TableText': 'table',
      'TableHeader': 'table_header',
      'FigureCaption': 'figure_caption',
    },
  };

  return USTB_TEMPLATE_CACHE;
}
