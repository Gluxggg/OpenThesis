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
  BorderStyle, WidthType, ShadingType, LevelFormat,
  convertMillimetersToTwip, Packer,
  ImageRun,
  Math as DocxMath, MathRun, MathFraction, MathRadical, MathSuperScript,
  MathSubScript, MathSubSuperScript,
  MathFunction, MathLimitLower, MathLimitUpper, BuilderElement,
  TableOfContents, TabStopType,
} from 'docx';
import type {
  ISectionOptions, ISectionPropertiesOptions, IPageNumberTypeAttributes,
  MathComponent, INumberingOptions,
} from 'docx';
import type {
  DocumentTemplate,
  ThesisDocument,
  ThesisSection,
  JournalArticle,
  JournalSection,
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
  /** Path to write the output .docx */
  outputPath: string;
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
 * Resolve the effective ParagraphStyle for a semantic role.
 *
 * The template wins for every property it actually declares; properties it
 * leaves unset fall back to the domain defaults. Merging per property rather
 * than per object is what lets a template that carries no formatting still
 * produce correctly styled headings instead of body text.
 */
function resolveStyle(template: DocumentTemplate, role: StyleRole, docType?: string): ParagraphStyle {
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

function renderHeading(
  template: DocumentTemplate,
  block: HeadingBlock,
  docType?: string,
): Paragraph {
  // `styleRole` lets a front/back-matter heading use its own format.
  const style = resolveStyle(template, block.styleRole ?? block.type, docType);
  const numberPrefix = block.number ? `${block.number}  ` : '';
  return new Paragraph({
    children: styledRuns(numberPrefix + block.text, style, style.font.bold),
    alignment: mapAlignment(style.paragraph.alignment),
    // Headings carry the guide's hanging indent (0.75 / 1 / 1.25 cm); this
    // paragraph never emitted an indent at all, so it was dropped silently.
    indent: indentOf(style),
    // The guide wants a page break between chapters, and the chapter style
    // declares one.
    pageBreakBefore: style.paragraph.pageBreakBefore,
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
    alignment: mapAlignment(style.paragraph.alignment) ?? AlignmentType.CENTER,
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
): (TextRun | DocxMath)[] {
  if (!text.includes('$')) {
    return [styledRun(text, style, forceBold, customSize)];
  }

  const runs: (TextRun | DocxMath)[] = [];
  // Match inline math $...$ that does not span newlines and is not escaped by \$
  const regex = /(?<!\\)\$(?!\s)([^$\n]+?)(?<![\s\\])\$/g;
  let lastIndex = 0;
  let match: RegExpExecArray | null;

  while ((match = regex.exec(text)) !== null) {
    if (match.index > lastIndex) {
      runs.push(styledRun(text.slice(lastIndex, match.index), style, forceBold, customSize));
    }
    const mathLatex = match[1].trim();
    try {
      const ast = latexToMathAst(mathLatex);
      runs.push(new DocxMath({ children: mathComponents(ast) }));
    } catch {
      runs.push(styledRun(`$${mathLatex}$`, style, forceBold, customSize));
    }
    lastIndex = regex.lastIndex;
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
function captionText(number: string | undefined, caption: string): string {
  const title = caption.trim();
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
  const colWidths = block.columnWidths ?? Array.from(
    { length: colCount },
    () => Math.floor(printableWidth / colCount),
  );
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

  const BO = { style: BorderStyle.SINGLE, size: 12, color: '000000' };
  const BI = { style: BorderStyle.SINGLE, size: 6, color: '000000' };
  const NO = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };

  // Caption. The role's own 段前/段后 are what reach `w:spacing`: hardcoding
  // 120/60 here overwrote the template's `ub` (150/50), which is why the
  // university's 段前1行 never appeared in the produced document.xml. A table with
  // no caption gets no empty paragraph carrying that 段前 spacing.
  const captionValue = captionText(block.number, block.caption);
  const captionPara = captionValue
    ? new Paragraph({
        children: styledRuns(captionValue, captionStyle, true),
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
        borders: isThreeLine ? {
          top: BO,
          bottom: BI,
          left: NO,
          right: NO,
        } : (showGrid ? {
          top: BO,
          bottom: BO,
          left: ci === 0 ? BO : BI,
          right: ci === colCount - 1 ? BO : BI,
        } : {
          top: NO, bottom: NO, left: NO, right: NO,
        }),
        shading: showShading ? { fill: 'D9D9D9', type: ShadingType.SOLID, color: 'auto' } : undefined,
        children: [
          new Paragraph({
            children: styledRuns(h, headerStyle, true),
            alignment: mapAlignment(headerStyle.paragraph.alignment),
            spacing: { before: 40, after: 40, line: headerStyle.lineSpacing },
          }),
        ],
      }),
    ),
  });

  // Data rows
  const dataRows = block.data.map((row, ri) => {
    const isLast = ri === block.data.length - 1;
    return new TableRow({
      children: block.headers.map((_, ci) =>
        new TableCell({
          width: { size: colWidths[ci], type: WidthType.DXA },
          borders: isThreeLine ? {
            top: NO,
            bottom: isLast ? BO : NO,
            left: NO,
            right: NO,
          } : (showGrid ? {
            top: BI,
            bottom: isLast ? BO : BI,
            left: ci === 0 ? BO : BI,
            right: ci === colCount - 1 ? BO : BI,
          } : {
            top: NO, bottom: NO, left: NO, right: NO,
          }),
          children: [
            new Paragraph({
              children: styledRuns(row[ci] ?? '', bodyStyle),
              alignment: mapAlignment(bodyStyle.paragraph.alignment),
              spacing: { before: 20, after: 20, line: bodyStyle.lineSpacing },
            }),
          ],
        }),
      ),
    });
  });

  const table = new Table({
    // Declare the same total the cells declare; a 100% table width alongside
    // absolute cell widths is contradictory OOXML.
    width: { size: tableWidth, type: WidthType.DXA },
    borders: isThreeLine ? {
      top: BO, bottom: BO, left: NO, right: NO,
      insideHorizontal: NO, insideVertical: NO,
    } : (showGrid ? {
      top: BO, bottom: BO, left: BO, right: BO,
      insideHorizontal: BI, insideVertical: BI,
    } : {
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

function renderFigure(
  template: DocumentTemplate,
  block: FigureBlock,
  contentDir?: string,
  docType?: string,
): Paragraph[] {
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
      
      if (size) {
        if (!width && !height) {
          // Fit the printable page width while keeping the aspect ratio
          const targetWidth = Math.min(size.width, maxWidthPx);
          const ratio = size.width / size.height;
          width = targetWidth;
          height = targetWidth / ratio;
        } else if (width && !height) {
          const ratio = size.width / size.height;
          height = width / ratio;
        } else if (!width && height) {
          const ratio = size.width / size.height;
          width = height * ratio;
        }
      } else {
        width = width || Math.min(400, maxWidthPx);
        height = height || 300;
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
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 120, line: 312 },
      });
      // 图题 follows the figure and carries the role's own 段前/段后 (`u4`: 10/10
      // for 0.1 行). The run used to be `${number}  ${caption}` with two spaces and
      // a hardcoded 段后 120.
      const caption = new Paragraph({
        children: styledRuns(captionText(block.number, block.caption), captionStyle),
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

  return [new Paragraph({
    children: [
      new TextRun({
        text: `[图: ${block.caption} (未找到图片: ${block.path})]`,
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
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 120, ...captionSpacing },
  })];
}

function renderSpacer(_template: DocumentTemplate, lines: number): Paragraph[] {
  return Array.from({ length: lines }, () =>
    new Paragraph({
      spacing: { line: 312 },
      children: [],
    }),
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
      return renderSpacer(template, block.lines);

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
  const HEADING_TYPES = ['heading1', 'heading2', 'heading3', 'heading4'] as const;

  // Body Sections — recursively flatten to support arbitrary nesting depth
  function flattenSections(sections: ThesisSection[], depth: number = 0): ContentBlock[] {
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
      if (section.type === 'toc') {
        // Word fills the entries and their page numbers when the field is
        // updated; no generator can compute page numbers, so the honest output
        // is the field plus `w:updateFields` asking Word to refresh it.
        children.push(renderHeading(template, { type: 'heading1', text: section.title, styleRole: 'section_heading' }));
        children.push(tocBlock(section.title));
        continue;
      }
      children.push(...processBlocks(trimTrailingPageBreak(flattenSections([section])), template, contentDir, undefined, state));
    }
    parts.push({ role: 'front_matter', children });
  }

  const bodyChildren: (Paragraph | Table)[] = [];
  if (body.length > 0) {
    bodyChildren.push(...processBlocks(flattenSections(body), template, contentDir, undefined, state));
  }

  // Back Matter. 参考文献 and everything after it belong to the body section:
  // the guide numbers them with the body, not with the front matter.
  if (doc.backMatter?.references && doc.backMatter.references.length > 0) {
    bodyChildren.push(
      renderHeading(template, { type: 'heading1', text: '参考文献', styleRole: 'section_heading' }),
      ...doc.backMatter.references.map(ref =>
        renderParagraph(template, {
          type: 'paragraph_no_indent',
          text: `[${ref.id}] ${ref.text}`,
        }, undefined, 'reference_item'),
      ),
    );
  }

  if (doc.backMatter?.appendices && doc.backMatter.appendices.length > 0) {
    bodyChildren.push(...processBlocks(flattenSections(doc.backMatter.appendices), template, contentDir, undefined, state));
  }

  if (doc.backMatter?.authorBiography) {
    bodyChildren.push(
      renderHeading(template, { type: 'heading1', text: '作者简历及在学期间取得的成果', styleRole: 'section_heading' }),
      renderParagraph(template, { type: 'paragraph', text: doc.backMatter.authorBiography }),
    );
  }

  if (doc.backMatter?.declaration) {
    bodyChildren.push(
      renderHeading(template, { type: 'heading1', text: '独创性说明', styleRole: 'section_heading' }),
      renderParagraph(template, { type: 'paragraph', text: doc.backMatter.declaration }),
    );
  }

  if (bodyChildren.length > 0 || parts.length === 0) {
    parts.push({ role: 'body', children: bodyChildren });
  }

  return parts;
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

  // Title
  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: doc.meta.title,
          bold: true,
          size: 32,
          font: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '黑体', cs: 'Times New Roman' },
        }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 240, after: 120, line: 312 },
    })
  );

  // Authors
  if (doc.meta.authors && doc.meta.authors.length > 0) {
    const authorRuns: TextRun[] = [];
    doc.meta.authors.forEach((author, index) => {
      authorRuns.push(
        new TextRun({
          text: author.name + (author.isCorresponding ? '*' : ''),
          size: 21,
          font: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '宋体', cs: 'Times New Roman' },
        })
      );
      
      if (index < doc.meta.authors.length - 1) {
        authorRuns.push(new TextRun({ text: '  ', size: 21 }));
      }
    });

    children.push(
      new Paragraph({
        children: authorRuns,
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 60, line: 312 },
      })
    );
  }

  // Affiliations
  const allAffiliations: string[] = [];
  doc.meta.authors.forEach(author => {
    author.affiliations.forEach(aff => {
      const affStr = `${aff.institution}${aff.department ? ', ' + aff.department : ''}${aff.city ? ', ' + aff.city : ''}`;
      if (!allAffiliations.includes(affStr)) {
        allAffiliations.push(affStr);
      }
    });
  });

  if (allAffiliations.length > 0) {
    const affText = allAffiliations.map((aff, i) => `(${i + 1}) ${aff}`).join('  ');
    children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: affText,
            size: 18,
            italics: true,
            font: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '宋体', cs: 'Times New Roman' },
          }),
        ],
        alignment: AlignmentType.CENTER,
        spacing: { before: 60, after: 120, line: 312 },
      })
    );
  }

  // Abstract Block
  if (doc.meta.abstract) {
    children.push(
      new Paragraph({
        children: [
          new TextRun({ text: '摘  要：', bold: true, size: 20, font: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '黑体', cs: 'Times New Roman' } }),
          new TextRun({ text: doc.meta.abstract, size: 20, font: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '宋体', cs: 'Times New Roman' } }),
        ],
        indent: { left: 720, right: 720 },
        spacing: { before: 120, after: 60, line: 240 },
        alignment: AlignmentType.JUSTIFIED,
      })
    );
  }

  // Keywords
  if (doc.meta.keywords && doc.meta.keywords.length > 0) {
    children.push(
      new Paragraph({
        children: [
          new TextRun({ text: '关键词：', bold: true, size: 20, font: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '黑体', cs: 'Times New Roman' } }),
          new TextRun({ text: doc.meta.keywords.join('；'), size: 20, font: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '宋体', cs: 'Times New Roman' } }),
        ],
        indent: { left: 720, right: 720 },
        spacing: { before: 60, after: 240, line: 240 },
        alignment: AlignmentType.LEFT,
      })
    );
  }

  // Divider line or spacer before body
  children.push(new Paragraph({ children: [], spacing: { after: 120 } }));

  const HEADING_TYPES = ['heading1', 'heading2', 'heading3', 'heading4'] as const;

  // Sections — recursively flatten to support arbitrary nesting depth
  function flattenJournalSections(sections: JournalSection[], depth: number = 0): ContentBlock[] {
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
        result.push(...flattenJournalSections(section.subsections, depth + 1));
      }
    }
    return result;
  }

  children.push(...processBlocks(flattenJournalSections(doc.sections), template, contentDir, undefined, state));

  // Back Matter
  if (doc.backMatter) {
    if (doc.backMatter.references && doc.backMatter.references.length > 0) {
      children.push(
        renderHeading(template, { type: 'heading1', text: 'References' }),
        ...doc.backMatter.references.map(ref =>
          renderParagraph(template, {
            type: 'paragraph_no_indent',
            text: `[${ref.id}] ${ref.text}`,
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
function partPageNumbers(role: DocumentPartRole, settings: PageSettings): IPageNumberTypeAttributes {
  if (role === 'cover') {
    // Nothing prints on the cover; the format only decides what the *next*
    // section continues from, so an undeclared one is left undeclared.
    return { ...(settings.pageNumberFormat ? { formatType: settings.pageNumberFormat } : {}) };
  }
  if (role === 'front_matter') {
    return {
      formatType: settings.pageNumberFormat ?? 'upperRoman',
      start: settings.pageNumberStart ?? 1,
    };
  }
  return {
    formatType: settings.pageNumberFormat ?? 'decimal',
    start: settings.pageNumberStart ?? 1,
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
  },
): { section: ISectionOptions; evenReference: boolean } {
  const settings = partSettings(part.role, sections, template);
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
      ...pageProperties(settings, partPageNumbers(part.role, settings)),
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

  // Shared across the whole document so list numbering instances stay unique.
  const state = createRenderState();

  // The running head a template without header parts falls back to.
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
    if (templateSections.length > 1) {
      for (const part of parts) {
        const built = buildPartSection(part, template, templateSections, {
          headerText,
          // The even-page running head is the thesis's own title. A template's
          // part can only carry its author's sample title, which must not print.
          evenHeaderText: doc.meta.title,
          variables: headVariables,
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
    numbering: { config: LIST_NUMBERING },
    sections,
  });

  let buffer = await Packer.toBuffer(wordDoc);

  // ── Post-process: fix Chinese fonts (WPS compatibility) ──────
  // The docx library has a bug where font.eastAsia in styles.default
  // gets overwritten with the ascii font name. WPS then shows all text
  // in Times New Roman. We fix this by post-processing the zip:
  // 1. Replace all *Theme attributes in styles.xml with explicit fonts
  // 2. Fix docDefaults eastAsia font
  buffer = await fixChineseFonts(buffer);

  // ── Post-process: document settings ─────────────────────────
  // 对称页边距, 偶数页页眉 and "refresh fields on open" are all `w:settings`
  // children, not section properties, so they are written into settings.xml
  // after packing. `updateFields` matters because the 目录 is a TOC field whose
  // page numbers only exist once Word evaluates it — and the source template
  // does not set it, so a generated thesis would otherwise open with an empty
  // table of contents.
  //
  // 偶数页页眉 is written only when the output really asks for an even-page
  // head: the flag also demotes `w:type="default"` to odd pages, so setting it
  // while emitting no even reference would leave the even pages' head blank.
  const templateCarriesHeads = templateSections.some(section => section.headers || section.footers);
  buffer = await patchSettings(buffer, {
    mirrorMargins: Boolean(template.page.mirrorMargins),
    evenAndOddHeaders: evenReference || (Boolean(template.evenAndOddHeaders) && templateCarriesHeads),
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
 * Post-process the docx buffer to fix Chinese font rendering in WPS.
 *
 * Root cause: The `docx` npm library (v9.x) writes `w:eastAsia="Times New Roman"`
 * in docDefaults when you set `font: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: '宋体', cs: 'Times New Roman' }`
 * in styles.default.document.run — it ignores eastAsia and uses `name` for all
 * four font slots. WPS honors the docDefaults eastAsia value, so all Chinese
 * text renders in Times New Roman (which falls back to a default CJK font that
 * looks wrong). Individual TextRun-level font.eastAsia DOES work, but only for
 * runs that explicitly set it — any run inheriting from defaults is broken.
 *
 * Fix: After packing, unzip the buffer, strip all *Theme attributes from the
 * generated XML, and set docDefaults eastAsia to 仿宋 (the
 * standard Chinese official document body font). Individual run-level eastAsia
 * values (宋体/黑体/仿宋) are preserved and now take effect. The theme part is
 * retained so its package relationship remains valid.
 */
/**
 * 对称页边距 and 偶数页页眉 — document settings the writer library cannot emit.
 *
 * `ISectionPropertiesOptions` in the writer library has no mirror-margins
 * option, so the flag is written into the package after it is built. Idempotent:
 * a document that already declares it is returned untouched.
 */
async function patchSettings(
  buffer: Buffer,
  options: { mirrorMargins: boolean; evenAndOddHeaders: boolean },
): Promise<Buffer> {
  // `w:mirrorMargins`, `w:evenAndOddHeaders` and `w:updateFields` are *document*
  // settings — children of `w:settings` in `word/settings.xml`. None of them is a
  // legal child of `w:sectPr`, so writing them into document.xml (as the first
  // attempt did) produces a part Word ignores or repairs, and a verification that
  // only greps for the string cannot tell the difference.
  const zip = await JSZip.loadAsync(buffer);
  const entry = zip.file('word/settings.xml');
  if (!entry) return buffer;

  let xml = await entry.async('string');

  // Ask Word to refresh fields when the document opens, so the 目录 and its page
  // numbers are right on the first look rather than after a manual F9.
  if (!xml.includes('<w:updateFields')) {
    xml = xml.replace(/(<w:settings[^>]*>)/, '$1<w:updateFields w:val="true"/>');
  }

  if (options.mirrorMargins && !xml.includes('<w:mirrorMargins')) {
    xml = xml.replace(/(<w:settings[^>]*>)/, '$1<w:mirrorMargins/>');
  }

  // 偶数页页眉. The writer library always emits this element — with
  // `w:val="false"` when it was not asked for — so a document that wants it has
  // to rewrite the value. Inserting a second element instead would leave Word
  // reading the first one, which is `false`.
  if (options.evenAndOddHeaders) {
    xml = xml.includes('<w:evenAndOddHeaders')
      ? xml.replace(/<w:evenAndOddHeaders[^>]*\/>/, '<w:evenAndOddHeaders/>')
      : xml.replace(/(<w:settings[^>]*>)/, '$1<w:evenAndOddHeaders/>');
  }

  if (xml === await entry.async('string')) return buffer;
  zip.file('word/settings.xml', xml);
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function fixChineseFonts(buffer: Buffer): Promise<Buffer> {
  const zip = await JSZip.loadAsync(buffer);

  // 1. Fix styles.xml
  const stylesFile = zip.file('word/styles.xml');
  if (stylesFile) {
    let stylesXml = await stylesFile.async('string');
    // Remove all *Theme attributes
    stylesXml = stylesXml.replace(
      /\s+w:(asciiTheme|hAnsiTheme|eastAsiaTheme|cstheme)="[^"]*"/g,
      '',
    );
    // Fix any rFonts where eastAsia is a Latin font → 仿宋
    stylesXml = stylesXml.replace(
      /\sw:eastAsia="(Times New Roman|Arial|Calibri|Cambria)"/g,
      ' w:eastAsia="仿宋"',
    );
    zip.file('word/styles.xml', stylesXml);
  }

  // 2. Fix document.xml
  const docFile = zip.file('word/document.xml');
  if (docFile) {
    let docXml = await docFile.async('string');
    // Remove all *Theme attributes
    docXml = docXml.replace(
      /\s+w:(asciiTheme|hAnsiTheme|eastAsiaTheme|cstheme)="[^"]*"/g,
      '',
    );
    // Fix any rFonts where eastAsia is a Latin font → 仿宋
    docXml = docXml.replace(
      /\sw:eastAsia="(Times New Roman|Arial|Calibri|Cambria)"/g,
      ' w:eastAsia="仿宋"',
    );
    zip.file('word/document.xml', docXml);
  }

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

// ── Utility ───────────────────────────────────────────────

function mapAlignment(align?: string): (typeof AlignmentType)[keyof typeof AlignmentType] {
  switch (align) {
    case 'center': return AlignmentType.CENTER;
    case 'left': return AlignmentType.LEFT;
    case 'right': return AlignmentType.RIGHT;
    case 'justified': return AlignmentType.JUSTIFIED;
    case 'distribute': return AlignmentType.DISTRIBUTE;
    default: return AlignmentType.JUSTIFIED;
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

/**
 * Create a USTB thesis template, loading from the pre-parsed .docx template.
 * Falls back to a minimal hardcoded template if the JSON file is unavailable.
 */
export function createUSTBTemplate(): DocumentTemplate {
  try {
    if (existsSync(USTB_TEMPLATE_JSON)) {
      const raw = JSON.parse(readFileSync(USTB_TEMPLATE_JSON, 'utf-8'));
      return raw as DocumentTemplate;
    }
  } catch {
    // Fall through to hardcoded fallback
  }

  // Minimal hardcoded fallback (used when the parsed JSON is not available)
  return {
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
}
