// ============================================================
// @openthesis/document-schema — Document Domain Models
// ============================================================
// Defines WHAT a document IS — independent of any file format,
// rendering engine, or template system.
//
// Three domain models, one style system:
//   1. Thesis (学位论文)             — 学士/硕士/博士论文
//   2. JournalArticle (期刊论文)      — 学术期刊投稿
//   3. OfficialDocument (公文)       — 党政机关公文 (GB/T 9704)
// ============================================================

// ============================================================
// ── Document Type Enum ─────────────────────────────────────
// ============================================================

export type DocumentType = 'thesis' | 'journal' | 'official';

// ============================================================
// ── TEMPLATE (parsed from .docx) ───────────────────────────
// ============================================================

export interface DocumentTemplate {
  meta: TemplateMeta;
  page: PageSettings;
  /**
   * Every section geometry found, in document order (cover, body, appendix…).
   * `page` is the last one — the body — which is what a single-section render
   * targets. Present only when the document actually declares sections.
   */
  pageSections?: PageSettings[];
  /**
   * `<w:evenAndOddHeaders/>` in `word/settings.xml`: the document tells Word that
   * odd and even pages use different running heads, so a `w:type="even"`
   * reference is honoured and `w:type="default"` means *odd* pages only.
   *
   * This is a document setting, not a section property — writing it into
   * `w:sectPr` produces a part Word ignores.
   */
  evenAndOddHeaders?: boolean;
  styles: Record<string, ParagraphStyle>;
  styleRoles: Record<string, StyleRole>;
  /**
   * The style that should drive each role.
   *
   * `styleRoles` is an id→role map, and a plain object always enumerates
   * integer-like keys first — Word commonly names styles `1`, `2`, `3`, so "the
   * first entry carrying this role" may not be the one the template was built
   * around. This records the parser's decision explicitly, leaving `styleRoles`
   * as the complete mapping.
   */
  roleWinners?: Partial<Record<StyleRole, string>>;
  styleInheritance?: Record<string, string>;
  documentType?: DocumentType;
  /**
   * Non-fatal problems found while parsing — a style based on a style the
   * template does not define, a template that declares no formatting at all,
   * or several sections where only one can be rendered. Callers should surface
   * these rather than letting a template silently produce wrong output.
   */
  warnings?: string[];
}

export interface TemplateMeta {
  organization: string;
  name: string;
  documentType: DocumentType;
  subType?: string;
  sourceFile?: string;
  parserVersion: string;
  parsedAt: string;
}

export interface PageSettings {
  width: number;
  height: number;
  margins: {
    top: number;
    bottom: number;
    left: number;
    right: number;
  };
  headerDistance?: number;
  footerDistance?: number;
  /**
   * Twips reserved on the binding edge (`w:pgMar/@w:gutter`). The university
   * guide asks for 1 cm on the left.
   */
  gutter?: number;
  /**
   * 对称页边距: mirror the left/right margins on facing pages, as the guide asks
   * for a bound thesis. The writer library has no option for it, so the renderer
   * sets the OOXML flag itself.
   */
  mirrorMargins?: boolean;
  columns?: number;
  columnGutter?: number;
  /**
   * `w:pgNumType/@w:fmt`: the format this section numbers its pages in. The
   * template declares `upperRoman` for the front matter and nothing for the
   * body, and a section that restarts without naming a format is decimal.
   */
  pageNumberFormat?: PageNumberFormat;
  /**
   * `w:pgNumType/@w:start`: restart the count at this number here. The template
   * restarts the front matter at I and the body at 1.
   */
  pageNumberStart?: number;
  /**
   * The running heads (`w:headerReference`) this section ends up with, one per
   * OOXML slot, after inheritance from the preceding sections is resolved.
   *
   * OOXML inherits each of the six slots independently, so a section that
   * declares nothing still prints its predecessor's head — the body of the
   * university template declares no `w:headerReference` of its own yet does
   * print 北京科技大学硕士学位论文 on odd pages and the thesis title on even ones.
   */
  headers?: RunningHeadSlots;
  /** The footers (`w:footerReference`) this section ends up with, same rules. */
  footers?: RunningHeadSlots;
}

/** Page-number formats the generator understands (`w:pgNumType/@w:fmt`). */
export type PageNumberFormat =
  | 'decimal' | 'upperRoman' | 'lowerRoman' | 'upperLetter' | 'lowerLetter'
  | 'chineseCounting' | 'chineseCountingThousand' | 'none';

/**
 * One running-head/footer slot, as the template's own part defines it. A slot is
 * absent when no section references a part for it; it is present with `text: ''`
 * when the template points at an *empty* part, which is how the university
 * template suppresses the head on the cover.
 */
export interface RunningHead {
  /** Literal text of the referenced part. `''` = the part is empty. */
  text?: string;
  /** The part draws a PAGE field, so this slot shows the page number. */
  pageNumber?: boolean;
  /** 篇眉: the part asks for a rule under the text (`w:pBdr/w:bottom`). */
  rule?: boolean;
}

/**
 * The three header (or footer) slots a section can carry. `default` is the odd
 * page's head once `<w:evenAndOddHeaders/>` is set, `first` needs `w:titlePg`.
 */
export interface RunningHeadSlots {
  default?: RunningHead;
  even?: RunningHead;
  first?: RunningHead;
}

// ── Styles (shared) ────────────────────────────────────────

export interface ParagraphStyle {
  font: FontSettings;
  paragraph: ParagraphFormatting;
  lineSpacing?: number;
  /**
   * How `lineSpacing` must be read. `auto` counts 240ths of a line, while
   * `exact` and `atLeast` are twips. OOXML keeps both numbers in the same
   * `w:spacing/@w:line` attribute and puts the difference in `w:lineRule`, so
   * dropping the rule silently reinterpreted every `exact` value as a multiple
   * of the line height.
   */
  lineSpacingRule?: LineSpacingRule;
}

export type LineSpacingRule = 'auto' | 'exact' | 'atLeast';

/**
 * Run-level formatting. Every field is optional: a parsed template only
 * declares what it actually specifies. The renderer fills the gaps from the
 * semantic-role defaults, so "the template is silent" stays distinguishable
 * from "the template explicitly asks for 12pt".
 */
export interface FontSettings {
  name?: string;
  eastAsia?: string;
  size?: number;        // half-points (24 = 12pt)
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
}

export interface ParagraphFormatting {
  alignment?: 'left' | 'center' | 'right' | 'justified' | 'distribute';
  firstLineIndent?: number;
  leftIndent?: number;
  rightIndent?: number;
  /**
   * Twips of hanging indent: the first line starts this far left of the rest.
   * The guide asks for it on every heading level (0.75 / 1 / 1.25 cm), which is
   * a different property from `firstLineIndent` and was being thrown away.
   */
  hangingIndent?: number;
  spaceBefore?: number;
  spaceAfter?: number;
  outlineLevel?: number;
  /**
   * Start this paragraph on a fresh page. The guide requires a page break
   * between chapters, and the template declares it on the chapter style —
   * parsing kept the flag but the style mapping dropped it, so every chapter
   * ran on from the previous one.
   */
  pageBreakBefore?: boolean;
}

// ════════════════════════════════════════════════════════════
// ── CONTENT BLOCKS ─────────────────────────────────────────
// ════════════════════════════════════════════════════════════

export type BlockType =
  | 'heading1' | 'heading2' | 'heading3' | 'heading4'
  | 'paragraph' | 'paragraph_no_indent'
  | 'centered_text'
  | 'equation' | 'equation_numbered'
  | 'figure' | 'table'
  | 'list_item' | 'code_block' | 'blockquote'
  | 'horizontal_rule' | 'spacer' | 'page_break'
  // Official document blocks
  | 'red_header' | 'document_number' | 'recipient_line'
  | 'signature_block' | 'attachment_note';

export interface BaseBlock { type: BlockType; id?: string; }

/**
 * Semantic style roles. Every content block type doubles as a role, plus the
 * parts of a document that carry no block of their own: table headers, figure
 * and table captions, the headings of the front/back-matter sections (摘要,
 * 目录, 附录, 致谢, 参考文献 — the guide sets these apart from a chapter heading),
 * and the reference entries under one.
 */
export type StyleRole = BlockType
  | 'table_header' | 'figure_caption' | 'table_caption'
  | 'section_heading' | 'reference_item'
  | 'cover_title' | 'cover_line' | 'cover_meta'
  | 'toc_entry';

export interface HeadingBlock extends BaseBlock {
  type: 'heading1' | 'heading2' | 'heading3' | 'heading4';
  text: string; number?: string;
  /**
   * Render this heading with another role's style.
   *
   * The guide sets the headings of the front and back matter (摘要, 目录, 附录,
   * 致谢, 参考文献) apart from a chapter heading — 段后16.5磅 and 2.41倍行距
   * against 段后17磅 and 1.3倍 — so a level-1 heading is not always a chapter.
   */
  styleRole?: StyleRole;
}

export interface ParagraphBlock extends BaseBlock {
  type: 'paragraph' | 'paragraph_no_indent';
  text: string;
}

export interface CenteredTextBlock extends BaseBlock {
  type: 'centered_text';
  text: string; font_size_pt?: number; bold?: boolean;
  /**
   * Use the template's style for this line instead of the block-level defaults.
   *
   * The cover page is specified box by box (校名行 小二 18pt bold, 研究生/指导教师
   * 四号 bold at 1.5 line spacing, 中图分类号 五号), so the three tiers are roles
   * of their own rather than one `centered_text` style.
   */
  styleRole?: 'cover_title' | 'cover_line' | 'cover_meta';
}

export interface EquationBlock extends BaseBlock {
  type: 'equation' | 'equation_numbered';
  latex: string; number?: string;
}

export interface FigureBlock extends BaseBlock {
  type: 'figure';
  path: string; caption: string; number?: string;
  width?: number; height?: number;
}

export interface TableBlock extends BaseBlock {
  type: 'table';
  caption: string; headers: string[]; data: string[][];
  number?: string; columnWidths?: number[];
  showGridlines?: boolean; headerShading?: boolean;
}

export interface ListItemBlock extends BaseBlock {
  type: 'list_item';
  text: string; level?: number; ordered?: boolean; marker?: string;
}

export interface CodeBlockBlock extends BaseBlock {
  type: 'code_block'; text: string; language?: string;
}

export interface BlockquoteBlock extends BaseBlock {
  type: 'blockquote'; text: string;
}

export interface SpacerBlock extends BaseBlock { type: 'spacer'; lines: number; }
export interface PageBreakBlock extends BaseBlock { type: 'page_break'; }
export interface HorizontalRuleBlock extends BaseBlock { type: 'horizontal_rule'; }

// Official document blocks
export interface RedHeaderBlock extends BaseBlock {
  type: 'red_header'; text: string; font_size_pt?: number;
}
export interface DocumentNumberBlock extends BaseBlock {
  type: 'document_number'; text: string;
}
export interface RecipientLineBlock extends BaseBlock {
  type: 'recipient_line'; text: string; recipientType: 'primary' | 'cc';
}
export interface SignatureBlockBlock extends BaseBlock {
  type: 'signature_block'; authority: string; date: string;
}
export interface AttachmentNoteBlock extends BaseBlock {
  type: 'attachment_note'; attachments: string[];
}

export type ContentBlock =
  | HeadingBlock | ParagraphBlock | CenteredTextBlock
  | EquationBlock | FigureBlock | TableBlock
  | ListItemBlock | CodeBlockBlock | BlockquoteBlock
  | SpacerBlock | PageBreakBlock | HorizontalRuleBlock
  | RedHeaderBlock | DocumentNumberBlock | RecipientLineBlock
  | SignatureBlockBlock | AttachmentNoteBlock;

// ════════════════════════════════════════════════════════════
// ── 1. THESIS (学位论文) ────────────────────────────────────
// ════════════════════════════════════════════════════════════

export interface ThesisDocument {
  type: 'thesis';
  meta: ThesisMeta;
  cover: ContentBlock[];
  sections: ThesisSection[];
  backMatter?: ThesisBackMatter;
}

export interface ThesisMeta {
  title: string;
  titleEn?: string;
  author?: string;
  supervisor?: string;
  department?: string;
  major?: string;
  degree?: 'bachelor' | 'master' | 'doctor';
  date?: string;
  keywords?: string[];
  abstract?: string;
  abstractEn?: string;
  classificationNumber?: string;
  studentId?: string;
}

export interface ThesisSection {
  id: string;
  type: 'chapter' | 'abstract' | 'acknowledgement' | 'toc'
      | 'list_of_figures' | 'list_of_tables' | 'references' | 'appendix';
  title: string;
  number?: string;
  content: ContentBlock[];
  subsections?: ThesisSection[];
}

export interface ThesisBackMatter {
  references?: Reference[];
  appendices?: ThesisSection[];
  authorBiography?: string;
  declaration?: string;
  datasetInfo?: Record<string, string>;
}

// ════════════════════════════════════════════════════════════
// ── 2. JOURNAL ARTICLE (期刊论文) ──────────────────────────
// ════════════════════════════════════════════════════════════

export interface JournalArticle {
  type: 'journal';
  meta: JournalMeta;
  sections: JournalSection[];
  backMatter?: JournalBackMatter;
}

export interface JournalMeta {
  title: string;
  runningTitle?: string;
  authors: JournalAuthor[];
  journalName?: string;
  articleType?: 'research' | 'review' | 'short_communication' | 'case_study' | 'letter';
  abstract?: string;
  keywords?: string[];
  funding?: FundingInfo[];
  submittedAt?: string;
  acceptedAt?: string;
  doi?: string;
}

export interface JournalAuthor {
  name: string;
  nameLocal?: string;
  affiliations: Affiliation[];
  isCorresponding?: boolean;
  email?: string;
  orcid?: string;
  equalContribution?: boolean;
}

export interface Affiliation {
  institution: string;
  department?: string;
  city?: string;
  country?: string;
  postalCode?: string;
}

export interface FundingInfo {
  agency: string;
  grantNumber?: string;
}

export interface JournalSection {
  id: string;
  type: 'introduction' | 'methods' | 'results' | 'discussion'
      | 'conclusion' | 'materials' | 'background' | 'related_work'
      | 'experiment' | 'analysis' | 'appendix' | 'supplementary';
  title: string;
  number?: string;
  content: ContentBlock[];
  subsections?: JournalSection[];
}

export interface JournalBackMatter {
  references?: Reference[];
  supplementary?: string;
  authorContributions?: string;
  conflictOfInterest?: string;
  dataAvailability?: string;
  acknowledgments?: string;
}

export interface Reference {
  id: string;
  text: string;
  type?: 'journal' | 'book' | 'conference' | 'thesis' | 'web' | 'standard' | 'other';
  doi?: string;
  url?: string;
}

// ════════════════════════════════════════════════════════════
// ── 3. OFFICIAL DOCUMENT (公文 / GB/T 9704-2012) ───────────
// ════════════════════════════════════════════════════════════

export interface OfficialDocument {
  type: 'official';
  meta: OfficialMeta;
  body: ContentBlock[];
  attachments?: OfficialAttachment[];
}

export interface OfficialMeta {
  title: string;
  issuingAuthority: string;
  documentNumber: string;
  signer?: string;
  documentCategory: OfficialDocCategory;
  primaryRecipients: string[];
  ccRecipients?: string[];
  signatureAuthority?: string;
  date: string;
  notes?: string;
  issuingDepartment?: string;
  issueDate?: string;
  urgency?: 'normal' | 'urgent' | 'most_urgent';
  confidentiality?: 'unclassified' | 'secret' | 'confidential' | 'top_secret';
  distributionNumber?: string;
}

export type OfficialDocCategory =
  | '通知' | '通告' | '报告' | '请示' | '批复'
  | '函'   | '纪要' | '决定' | '意见' | '通报'
  | '议案' | '公告' | '命令' | '决议' | '公报';

export interface OfficialAttachment {
  order: number;
  title: string;
  content?: ContentBlock[];
}

// ════════════════════════════════════════════════════════════
// ── LEGACY FORMAT ──────────────────────────────────────────
// ════════════════════════════════════════════════════════════

export interface LegacyDocumentJSON {
  title: string;
  report_header?: string;
  cover_blocks: ContentBlock[];
  body_blocks: ContentBlock[];
}

export type OpenThesisDocument = ThesisDocument | JournalArticle | OfficialDocument;

// ════════════════════════════════════════════════════════════
// ── VALIDATION ─────────────────────────────────────────────
// ════════════════════════════════════════════════════════════

export { validateDocument, validateLegacyDocument, assertValidDocument, formatIssues } from './validate.js';
export type { ValidationIssue } from './validate.js';

