import type {
  ContentBlock,
  DocumentType,
  JournalArticle,
  JournalAuthor,
  JournalSection,
  OfficialDocCategory,
  OfficialDocument,
  OpenThesisDocument,
  ThesisDocument,
  ThesisSection,
} from '@openthesis/document-schema';

type FrontMatterValue = string | string[] | number | boolean;
export type MarkdownFrontMatter = Record<string, FrontMatterValue>;

export interface MarkdownParseOptions {
  documentType?: DocumentType;
  sourceName?: string;
  metadata?: MarkdownFrontMatter;
}

export interface MarkdownParseResult {
  document: OpenThesisDocument;
  frontMatter: MarkdownFrontMatter;
}

interface HeadingItem {
  kind: 'heading';
  level: number;
  text: string;
}

interface BlockItem {
  kind: 'block';
  block: ContentBlock;
}

type MarkdownItem = HeadingItem | BlockItem;

interface SectionNode {
  level: number;
  title: string;
  content: ContentBlock[];
  children: SectionNode[];
}

const DOCUMENT_TYPES: DocumentType[] = ['thesis', 'journal', 'official'];
const OFFICIAL_CATEGORIES: OfficialDocCategory[] = [
  '通知', '通告', '报告', '请示', '批复', '函', '纪要', '决定',
  '意见', '通报', '议案', '公告', '命令', '决议', '公报',
];

export function parseMarkdown(markdown: string, options: MarkdownParseOptions = {}): OpenThesisDocument {
  return parseMarkdownWithMetadata(markdown, options).document;
}

export function parseMarkdownWithMetadata(
  markdown: string,
  options: MarkdownParseOptions = {},
): MarkdownParseResult {
  const { body, frontMatter: parsedFrontMatter } = extractFrontMatter(markdown);
  const frontMatter = { ...parsedFrontMatter, ...options.metadata };
  const documentType = resolveDocumentType(options.documentType, frontMatter.type);
  const items = tokenize(body);
  const firstHeadingIndex = items.findIndex(item => item.kind === 'heading' && item.level === 1);
  const headingTitle = firstHeadingIndex >= 0
    ? (items[firstHeadingIndex] as HeadingItem).text
    : undefined;
  if (firstHeadingIndex >= 0) items.splice(firstHeadingIndex, 1);

  const title = stringValue(frontMatter, 'title')
    ?? headingTitle
    ?? options.sourceName
    ?? 'Untitled Document';

  const document = documentType === 'journal'
    ? createJournalDocument(title, items, frontMatter)
    : documentType === 'official'
      ? createOfficialDocument(title, items, frontMatter)
      : createThesisDocument(title, items, frontMatter);

  return { document, frontMatter };
}

export function parseMarkdownBlocks(markdown: string): ContentBlock[] {
  const { body } = extractFrontMatter(markdown);
  return tokenize(body).map(item => item.kind === 'heading'
    ? ({
        type: `heading${Math.min(item.level, 4)}`,
        text: item.text,
      } as ContentBlock)
    : item.block);
}

function extractFrontMatter(markdown: string): { body: string; frontMatter: MarkdownFrontMatter } {
  const normalized = markdown.replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n');
  const lines = normalized.split('\n');
  if (lines[0]?.trim() !== '---') return { body: normalized, frontMatter: {} };

  const closingIndex = lines.findIndex((line, index) => index > 0 && line.trim() === '---');
  if (closingIndex < 0) return { body: normalized, frontMatter: {} };

  // `---` is also a thematic break. Treating any later `---` as the end of a
  // front-matter block deleted everything between the two whenever the leading
  // line was not front matter at all — headings and prose included, silently.
  if (!looksLikeFrontMatter(lines.slice(1, closingIndex))) {
    return { body: normalized, frontMatter: {} };
  }

  const frontMatter: MarkdownFrontMatter = {};
  for (let index = 1; index < closingIndex; index += 1) {
    const match = /^([A-Za-z][\w-]*):\s*(.*)$/.exec(lines[index]);
    if (!match) continue;
    const key = match[1];
    const rawValue = match[2].trim();
    if (rawValue === '|' || rawValue === '>') {
      const parts: string[] = [];
      while (index + 1 < closingIndex && /^\s+/.test(lines[index + 1])) {
        index += 1;
        parts.push(lines[index].trim());
      }
      frontMatter[key] = rawValue === '|' ? parts.join('\n') : parts.join(' ');
    } else if (rawValue === '') {
      // Block sequence:
      //   keywords:
      //     - alpha
      //     - beta
      // The key pattern only matches unindented `key:` lines, so these items
      // used to be skipped entirely and the key silently became "".
      const items: string[] = [];
      let cursor = index + 1;
      while (cursor < closingIndex) {
        const item = /^\s+-\s+(.*)$/.exec(lines[cursor]);
        if (!item) break;
        const parsed = parseFrontMatterValue(item[1]);
        items.push(typeof parsed === 'string' ? parsed : String(parsed));
        cursor += 1;
      }
      if (items.length > 0) {
        frontMatter[key] = items;
        index = cursor - 1;
      } else {
        frontMatter[key] = '';
      }
    } else {
      frontMatter[key] = parseFrontMatterValue(rawValue);
    }
  }

  return { body: lines.slice(closingIndex + 1).join('\n'), frontMatter };
}

/**
 * Whether a `---`-delimited block really is front matter.
 *
 * It must hold at least one `key:` line and nothing that could not belong to
 * one, its block-sequence items and indented scalar continuations aside. A
 * `# comment` line is deliberately *not* accepted: in Markdown it is far more
 * likely to be a heading between two thematic breaks, and accepting it would
 * put the silent deletion back.
 */
function looksLikeFrontMatter(interior: string[]): boolean {
  const meaningful = interior.filter(line => line.trim() !== '');
  // `---\n---` is a deliberate empty block: there is nothing to lose.
  if (meaningful.length === 0) return true;

  let sawKey = false;
  for (const line of meaningful) {
    if (/^[A-Za-z][\w-]*:(\s|$)/.test(line)) { sawKey = true; continue; }
    if (/^\s+-(\s|$)/.test(line)) continue;   // block sequence item
    if (/^\s+\S/.test(line)) continue;        // folded or literal scalar continuation
    return false;
  }
  return sawKey;
}

function parseFrontMatterValue(value: string): FrontMatterValue {
  const trimmed = value.trim();
  const unquoted = stripQuotes(trimmed);
  if (unquoted !== trimmed) return unquoted;
  if (/^\[.*\]$/.test(unquoted)) {
    const inner = unquoted.slice(1, -1).trim();
    return inner ? splitList(inner) : [];
  }
  if (/^(true|false)$/i.test(unquoted)) return unquoted.toLowerCase() === 'true';
  if (/^-?\d+(?:\.\d+)?$/.test(unquoted)) return Number(unquoted);
  return unquoted;
}

function tokenize(markdown: string): MarkdownItem[] {
  const lines = markdown.replace(/\r\n?/g, '\n').split('\n');
  const items: MarkdownItem[] = [];
  let index = 0;

  while (index < lines.length) {
    const line = lines[index];
    const trimmed = line.trim();
    if (!trimmed) {
      index += 1;
      continue;
    }

    // CommonMark only treats a trailing `#` run as a closing sequence when a
    // space separates it from the text, so `# C#` keeps its `#`.
    const heading = /^ {0,3}(#{1,6})\s+(.+?)(?:\s+#+)?\s*$/.exec(line);
    if (heading) {
      items.push({ kind: 'heading', level: heading[1].length, text: cleanInline(heading[2]) });
      index += 1;
      continue;
    }

    // The info string may carry more than a language token (`js title=x`);
    // the old `[\w.+-]*` failed to match, so the opener became a paragraph and
    // the closing fence then swallowed the rest of the document.
    const fence = /^\s*(`{3,}|~{3,})[ \t]*([^\n]*)$/.exec(line);
    if (fence) {
      const code: string[] = [];
      const fenceMarker = fence[1][0];
      index += 1;
      while (index < lines.length && !new RegExp(`^\\s*${escapeRegExp(fenceMarker)}{${fence[1].length},}\\s*$`).test(lines[index])) {
        code.push(lines[index]);
        index += 1;
      }
      if (index < lines.length) index += 1;
      items.push({ kind: 'block', block: { type: 'code_block', text: code.join('\n'), language: fence[2] || undefined } });
      continue;
    }

    // An opener carrying no second `$$` starts a multi-line block. `$$a=b$$
    // trailing` is a complete equation, and reading it as an opener made the
    // scan below swallow every following line up to the next `$$`.
    if (trimmed === '$$' || (trimmed.startsWith('$$') && !trimmed.slice(2).includes('$$'))) {
      const equation: string[] = [];
      if (trimmed !== '$$') equation.push(trimmed.slice(2));
      index += 1;
      while (index < lines.length && !lines[index].trim().endsWith('$$')) {
        equation.push(lines[index].trim());
        index += 1;
      }
      if (index < lines.length) {
        equation.push(lines[index].trim().replace(/\$\$$/, ''));
        index += 1;
      }
      items.push({ kind: 'block', block: { type: 'equation', latex: equation.join(' ').trim() } });
      continue;
    }
    const displayMath = /^\$\$([\s\S]*?)\$\$(.*)$/.exec(trimmed);
    if (displayMath) {
      items.push({ kind: 'block', block: { type: 'equation', latex: displayMath[1].trim() } });
      // Text after the closing `$$` is body text belonging to the paragraph that
      // follows, not part of the equation.
      const trailing = displayMath[2].trim();
      if (trailing) {
        items.push({ kind: 'block', block: { type: 'paragraph', text: cleanInline(trailing) } });
      }
      index += 1;
      continue;
    }

    if (/^<!--\s*(?:page[-_ ]?break|newpage)\s*-->$/i.test(trimmed)) {
      items.push({ kind: 'block', block: { type: 'page_break' } });
      index += 1;
      continue;
    }

    if (/^ {0,3}((\*|-|_)\s*){3,}$/.test(line)) {
      items.push({ kind: 'block', block: { type: 'horizontal_rule' } });
      index += 1;
      continue;
    }

    const image = /^!\[([^\]]*)\]\((.+?)(?:\s+["']([^"']*)["'])?\)\s*$/.exec(trimmed);
    if (image) {
      const path = stripAngleBrackets(image[2].trim());
      const fallbackCaption = path.split(/[\\/]/).pop() ?? 'Figure';
      items.push({
        kind: 'block',
        block: { type: 'figure', path, caption: cleanInline(image[1] || image[3] || fallbackCaption) },
      });
      index += 1;
      continue;
    }

    if (isTableStart(lines, index)) {
      const headers = splitTableRow(lines[index]).map(cleanInline);
      const data: string[][] = [];
      index += 2;
      while (index < lines.length && lines[index].includes('|') && lines[index].trim()) {
        const row = splitTableRow(lines[index]).map(cleanInline);
        data.push(headers.map((_, cellIndex) => row[cellIndex] ?? ''));
        index += 1;
      }
      // A caption line sits immediately above its table in most manuscripts.
      // GFM has no caption syntax, so that paragraph *is* the caption — it used
      // to be emitted as body text, which is why `表 3-1` came out 宋体 小四
      // justified instead of the template's 黑体 五号 centred with 段前1行.
      let caption = '';
      const previous = items[items.length - 1];
      if (previous?.kind === 'block' && previous.block.type === 'paragraph'
          && /^(?:表|图)\s*[\dA-Za-z]/.test(previous.block.text)) {
        // One space between the number and the title, however many the source had.
        caption = previous.block.text.replace(/^((?:表|图)\s*[\dA-Za-z][\d.\-–—]*)\s+/, '$1 ');
        items.pop();
      }
      items.push({ kind: 'block', block: { type: 'table', caption, headers, data } });
      continue;
    }

    if (/^\s*>/.test(line)) {
      const quote: string[] = [];
      while (index < lines.length && /^\s*>/.test(lines[index])) {
        quote.push(lines[index].replace(/^\s*>\s?/, ''));
        index += 1;
      }
      items.push({ kind: 'block', block: { type: 'blockquote', text: cleanInline(joinSoftLines(quote)) } });
      continue;
    }

    const listMatch = /^(\s*)([-+*]|\d+[.)])\s+(.+)$/.exec(line);
    if (listMatch) {
      // Indent column of each open list level. Nesting is derived from where a
      // marker actually sits, not from `floor(indent / 2)`: Markdown nests at
      // 2 or 4 spaces and both used to land on the wrong level.
      const openIndents: number[] = [];

      while (index < lines.length) {
        const item = /^(\s*)([-+*]|\d+[.)])\s+(.+)$/.exec(lines[index]);
        if (!item) break;
        const indentation = item[1].replace(/\t/g, '    ').length;

        while (openIndents.length > 0 && indentation < openIndents[openIndents.length - 1]) {
          openIndents.pop();
        }
        if (openIndents.length === 0 || indentation > openIndents[openIndents.length - 1]) {
          openIndents.push(indentation);
        }

        items.push({
          kind: 'block',
          block: {
            type: 'list_item',
            text: cleanInline(item[3]),
            level: openIndents.length - 1,
            ordered: /^\d/.test(item[2]),
            marker: item[2],
          },
        });
        index += 1;
      }
      continue;
    }

    const paragraph: string[] = [];
    while (index < lines.length && lines[index].trim() && !startsSpecialBlock(lines, index)) {
      paragraph.push(lines[index].trim());
      index += 1;
    }
    if (paragraph.length === 0) {
      paragraph.push(trimmed);
      index += 1;
    }
    items.push({ kind: 'block', block: { type: 'paragraph', text: cleanInline(joinSoftLines(paragraph)) } });
  }

  attachCaptionLines(items);

  return items;
}

/**
 * A `表 N 标题` / `图 N 标题` line written as its own paragraph.
 *
 * The number and the title must be separated by one space, so a line that
 * carries two (`表 2  40 m资格模型结果`) is normalised. Both 表3-1 and 表 3-1 are
 * numbering an author chose, so the label keeps the shape it was written in. A
 * body sentence that merely starts with the character (`图1表明…`) has no space
 * after its number and is not a caption.
 */
const CAPTION_LINE = /^\s*([表图])([\s\u3000]*)([\dA-Za-z][\d.\-–—]*)[\s\u3000]+(\S[\s\S]*)$/;

function captionLine(line: string): string | null {
  const match = CAPTION_LINE.exec(line);
  if (!match) return null;
  const label = match[2] ? `${match[1]} ${match[3]}` : `${match[1]}${match[3]}`;
  return `${label} ${match[4].trim()}`;
}

/**
 * GFM has no caption syntax, so a manuscript writes 表题/图题 as a paragraph of
 * their own. Attach each to the block it names — the table it precedes, the
 * figure it follows — so the caption reaches its own style role, with the
 * template's 段前/段后 and 黑体 五号 centred, instead of being rendered as body text.
 *
 * A table caption the tokenizer already absorbed is left alone, which keeps this
 * pass idempotent.
 */
function attachCaptionLines(items: MarkdownItem[]): void {
  for (let index = 0; index < items.length; index += 1) {
    const item = items[index];
    if (item.kind !== 'block') continue;

    if (item.block.type === 'table' && !item.block.caption) {
      const previous = items[index - 1];
      const title = previous?.kind === 'block' && previous.block.type === 'paragraph'
        ? captionLine(previous.block.text)
        : null;
      if (title?.startsWith('表')) {
        item.block.caption = title;
        items.splice(index - 1, 1);
        index -= 1;
      }
    } else if (item.block.type === 'figure') {
      // An explicit 图题 line wins over the alt text, which is only a fallback.
      const next = items[index + 1];
      const title = next?.kind === 'block' && next.block.type === 'paragraph'
        ? captionLine(next.block.text)
        : null;
      if (title?.startsWith('图')) {
        item.block.caption = title;
        items.splice(index + 1, 1);
      }
    }
  }
}

function startsSpecialBlock(lines: string[], index: number): boolean {
  const line = lines[index];
  const trimmed = line.trim();
  return /^ {0,3}#{1,6}\s+/.test(line)
    || /^\s*(```+|~~~+)/.test(line)
    || trimmed.startsWith('$$')
    || /^<!--\s*(?:page[-_ ]?break|newpage)\s*-->$/i.test(trimmed)
    || /^ {0,3}((\*|-|_)\s*){3,}$/.test(line)
    || /^!\[[^\]]*\]\(.+\)\s*$/.test(trimmed)
    || /^\s*>/.test(line)
    || /^(\s*)([-+*]|\d+[.)])\s+/.test(line)
    || isTableStart(lines, index);
}

function isTableStart(lines: string[], index: number): boolean {
  if (index + 1 >= lines.length || !lines[index].includes('|')) return false;
  const separator = lines[index + 1].trim();
  // GFM allows single-column tables, so the repeated group is optional. The
  // separator must still contain a pipe, otherwise `a | b` followed by `---`
  // is a setext heading rather than a table.
  if (!separator.includes('|')) return false;
  return /^\|?\s*:?-{3,}:?\s*(?:\|\s*:?-{3,}:?\s*)*\|?$/.test(separator);
}

function splitTableRow(line: string): string[] {
  const trimmed = line.trim().replace(/^\|/, '').replace(/\|$/, '');
  const cells: string[] = [];
  let current = '';
  let escaped = false;
  for (const character of trimmed) {
    if (escaped) {
      current += character;
      escaped = false;
    } else if (character === '\\') {
      escaped = true;
    } else if (character === '|') {
      cells.push(current.trim());
      current = '';
    } else {
      current += character;
    }
  }
  cells.push(current.trim());
  return cells;
}

function createThesisDocument(
  title: string,
  items: MarkdownItem[],
  metadata: MarkdownFrontMatter,
): ThesisDocument {
  const author = stringValue(metadata, 'author');
  const supervisor = stringValue(metadata, 'supervisor');
  const department = stringValue(metadata, 'department', 'institution');
  const major = stringValue(metadata, 'major');
  const studentId = stringValue(metadata, 'studentId', 'student_id');
  const degree = degreeValue(metadata.degree);
  const date = stringValue(metadata, 'date');
  const titleEn = stringValue(metadata, 'titleEn', 'title_en');
  const classificationNumber = stringValue(metadata, 'classificationNumber', 'classification_number');

  const cover: ContentBlock[] = [];
  if (classificationNumber) {
    cover.push({ type: 'centered_text', text: `中图分类号：${classificationNumber}`, styleRole: 'cover_meta' });
  }
  const degreeText = degree === 'doctor' ? '北京科技大学博士学位论文' : '北京科技大学硕士学位论文';
  cover.push(
    { type: 'spacer', lines: 3 },
    { type: 'centered_text', text: degreeText, styleRole: 'cover_title' },
    { type: 'spacer', lines: 2 },
    { type: 'centered_text', text: title, styleRole: 'cover_title' },
  );
  if (titleEn) {
    cover.push({ type: 'centered_text', text: titleEn, styleRole: 'cover_line' });
  }
  cover.push({ type: 'spacer', lines: 3 });
  if (studentId) cover.push({ type: 'centered_text', text: `学　　号：${studentId}`, styleRole: 'cover_line' });
  if (author) cover.push({ type: 'centered_text', text: `研 究 生：${author}`, styleRole: 'cover_line' });
  if (supervisor) cover.push({ type: 'centered_text', text: `指导教师：${supervisor}`, styleRole: 'cover_line' });
  if (department) cover.push({ type: 'centered_text', text: `培养学院：${department}`, styleRole: 'cover_line' });
  if (major) cover.push({ type: 'centered_text', text: `专　　业：${major}`, styleRole: 'cover_line' });
  if (date) cover.push({ type: 'centered_text', text: `日　　期：${date}`, styleRole: 'cover_line' });
  cover.push({ type: 'page_break' });

  return {
    type: 'thesis',
    meta: {
      title,
      titleEn: stringValue(metadata, 'titleEn', 'title_en'),
      author,
      supervisor: stringValue(metadata, 'supervisor'),
      department,
      major: stringValue(metadata, 'major'),
      degree: degreeValue(metadata.degree),
      date,
      keywords: arrayValue(metadata, 'keywords'),
      abstract: stringValue(metadata, 'abstract'),
      abstractEn: stringValue(metadata, 'abstractEn', 'abstract_en'),
      studentId: stringValue(metadata, 'studentId', 'student_id'),
    },
    cover,
    sections: buildSectionTree(items).map((node, index) => thesisSection(node, `${index + 1}`)),
  };
}

function createJournalDocument(
  title: string,
  items: MarkdownItem[],
  metadata: MarkdownFrontMatter,
): JournalArticle {
  const authorNames = arrayValue(metadata, 'authors', 'author') ?? [];
  const affiliations = arrayValue(metadata, 'affiliations', 'affiliation', 'institution') ?? [];
  const authors: JournalAuthor[] = authorNames.map(name => ({
    name,
    affiliations: affiliations.map(institution => ({ institution })),
  }));

  return {
    type: 'journal',
    meta: {
      title,
      runningTitle: stringValue(metadata, 'runningTitle', 'running_title'),
      authors,
      journalName: stringValue(metadata, 'journalName', 'journal'),
      articleType: articleTypeValue(metadata.articleType ?? metadata.article_type),
      abstract: stringValue(metadata, 'abstract'),
      keywords: arrayValue(metadata, 'keywords'),
      doi: stringValue(metadata, 'doi'),
    },
    sections: buildSectionTree(items).map((node, index) => journalSection(node, `${index + 1}`)),
  };
}

function createOfficialDocument(
  title: string,
  items: MarkdownItem[],
  metadata: MarkdownFrontMatter,
): OfficialDocument {
  const categoryValue = stringValue(metadata, 'documentCategory', 'category');
  // An unknown category used to become 通知 silently, so a notice with a
  // mistyped category rendered as a different kind of document and no check
  // could see it. An absent category still defaults; a wrong one is an error.
  if (categoryValue !== undefined && !OFFICIAL_CATEGORIES.includes(categoryValue as OfficialDocCategory)) {
    throw new Error(
      `Unsupported official document category: ${categoryValue}. Expected one of: ${OFFICIAL_CATEGORIES.join(', ')}`,
    );
  }
  const documentCategory = (categoryValue as OfficialDocCategory | undefined) ?? '通知';
  const body = items.map(item => item.kind === 'heading'
    ? ({ type: `heading${Math.min(item.level, 4)}`, text: item.text } as ContentBlock)
    : item.block);

  return {
    type: 'official',
    meta: {
      title,
      issuingAuthority: stringValue(metadata, 'issuingAuthority', 'authority', 'organization') ?? '',
      documentNumber: stringValue(metadata, 'documentNumber', 'document_number') ?? '',
      documentCategory,
      primaryRecipients: arrayValue(metadata, 'primaryRecipients', 'recipients') ?? [],
      ccRecipients: arrayValue(metadata, 'ccRecipients', 'cc'),
      signatureAuthority: stringValue(metadata, 'signatureAuthority', 'signature_authority'),
      date: stringValue(metadata, 'date') ?? '',
      urgency: urgencyValue(metadata.urgency),
    },
    body,
  };
}

function buildSectionTree(items: MarkdownItem[]): SectionNode[] {
  const roots: SectionNode[] = [];
  const stack: SectionNode[] = [];
  let preamble: ContentBlock[] = [];

  for (const item of items) {
    if (item.kind === 'block') {
      if (stack.length > 0) stack[stack.length - 1].content.push(item.block);
      else preamble.push(item.block);
      continue;
    }

    const node: SectionNode = { level: item.level, title: item.text, content: [], children: [] };
    while (stack.length > 0 && stack[stack.length - 1].level >= item.level) stack.pop();
    if (stack.length > 0) stack[stack.length - 1].children.push(node);
    else roots.push(node);
    stack.push(node);
  }

  if (preamble.length > 0) {
    roots.unshift({ level: 1, title: '正文', content: preamble, children: [] });
  }
  if (roots.length === 0) {
    roots.push({ level: 1, title: '正文', content: [], children: [] });
  }
  return roots;
}

function thesisSection(node: SectionNode, id: string): ThesisSection {
  return {
    id: `section-${id}-${slugify(node.title)}`,
    type: thesisSectionType(node.title),
    title: node.title,
    content: node.content,
    subsections: node.children.map((child, index) => thesisSection(child, `${id}-${index + 1}`)),
  };
}

function journalSection(node: SectionNode, id: string): JournalSection {
  return {
    id: `section-${id}-${slugify(node.title)}`,
    type: journalSectionType(node.title),
    title: node.title,
    content: node.content,
    subsections: node.children.map((child, index) => journalSection(child, `${id}-${index + 1}`)),
  };
}

function thesisSectionType(title: string): ThesisSection['type'] {
  const normalized = title.toLowerCase();
  if (/abstract|摘要/.test(normalized)) return 'abstract';
  if (/acknowledg|致谢/.test(normalized)) return 'acknowledgement';
  if (/references|bibliography|参考文献/.test(normalized)) return 'references';
  if (/appendix|附录/.test(normalized)) return 'appendix';
  if (/contents|目录/.test(normalized)) return 'toc';
  return 'chapter';
}

function journalSectionType(title: string): JournalSection['type'] {
  const normalized = title.toLowerCase();
  if (/method|方法/.test(normalized)) return 'methods';
  if (/material|材料/.test(normalized)) return 'materials';
  if (/result|结果/.test(normalized)) return 'results';
  if (/discussion|讨论/.test(normalized)) return 'discussion';
  if (/conclusion|结论/.test(normalized)) return 'conclusion';
  if (/background|背景/.test(normalized)) return 'background';
  if (/related|相关工作|文献综述/.test(normalized)) return 'related_work';
  if (/experiment|实验/.test(normalized)) return 'experiment';
  if (/analysis|分析/.test(normalized)) return 'analysis';
  if (/appendix|附录/.test(normalized)) return 'appendix';
  return 'introduction';
}

function resolveDocumentType(option: DocumentType | undefined, value: FrontMatterValue | undefined): DocumentType {
  if (option) return option;
  const candidate = typeof value === 'string' ? value : 'thesis';
  if (!DOCUMENT_TYPES.includes(candidate as DocumentType)) {
    throw new Error(`Unsupported Markdown document type: ${candidate}`);
  }
  return candidate as DocumentType;
}

function stringValue(metadata: MarkdownFrontMatter, ...keys: string[]): string | undefined {
  for (const key of keys) {
    const value = metadata[key];
    if (typeof value === 'string' && value.trim()) return value.trim();
    if (typeof value === 'number' || typeof value === 'boolean') return String(value);
  }
  return undefined;
}

function arrayValue(metadata: MarkdownFrontMatter, ...keys: string[]): string[] | undefined {
  for (const key of keys) {
    const value = metadata[key];
    if (Array.isArray(value)) return value.map(item => item.trim()).filter(Boolean);
    if (typeof value === 'string' && value.trim()) return splitList(value);
  }
  return undefined;
}

/**
 * Split a flow sequence on commas/semicolons, ignoring separators inside
 * quotes. Splitting first and stripping quotes afterwards tore
 * `["a, b", c]` into `'"a'`, `'b"'`, `'c'`.
 */
function splitList(value: string): string[] {
  const items: string[] = [];
  let current = '';
  let quote: string | null = null;

  for (const character of value) {
    if (quote) {
      if (character === quote) quote = null;
      current += character;
    } else if (character === '"' || character === "'") {
      quote = character;
      current += character;
    } else if (character === ',' || character === ';' || character === '；') {
      items.push(current);
      current = '';
    } else {
      current += character;
    }
  }
  items.push(current);

  return items.map(stripQuotes).map(item => item.trim()).filter(Boolean);
}

/**
 * Spellings an author is likely to write for the three degree levels.
 *
 * Anything else is refused rather than dropped: `degree: PhD` used to become
 * `undefined`, which JSON.stringify then removed, leaving no way to tell "the
 * author wrote a doctorate" from "no degree was given".
 */
const DEGREE_ALIASES: Record<string, ThesisDocument['meta']['degree']> = {
  bachelor: 'bachelor', bachelors: 'bachelor', bsc: 'bachelor', 'b.sc': 'bachelor',
  undergraduate: 'bachelor', 本科: 'bachelor', 学士: 'bachelor',
  master: 'master', masters: 'master', msc: 'master', 'm.sc': 'master', ma: 'master',
  graduate: 'master', 硕士: 'master', 研究生: 'master',
  doctor: 'doctor', doctorate: 'doctor', doctoral: 'doctor', phd: 'doctor',
  'ph.d': 'doctor', 博士: 'doctor',
};

function degreeValue(value: FrontMatterValue | undefined): ThesisDocument['meta']['degree'] {
  if (value === undefined) return undefined;
  const candidate = typeof value === 'string' ? value.trim().toLowerCase() : String(value);
  const degree = DEGREE_ALIASES[candidate];
  if (!degree) {
    throw new Error(
      `Unsupported degree: ${candidate}. Expected bachelor, master or doctor (or a common spelling such as PhD).`,
    );
  }
  return degree;
}

function articleTypeValue(value: FrontMatterValue | undefined): JournalArticle['meta']['articleType'] {
  const allowed = ['research', 'review', 'short_communication', 'case_study', 'letter'] as const;
  return typeof value === 'string' && allowed.includes(value as typeof allowed[number])
    ? value as typeof allowed[number]
    : undefined;
}

function urgencyValue(value: FrontMatterValue | undefined): OfficialDocument['meta']['urgency'] {
  return value === 'normal' || value === 'urgent' || value === 'most_urgent' ? value : undefined;
}

function cleanInline(text: string): string {
  return text
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/!\[([^\]]*)\]\([^)]+\)/g, '$1')
    .replace(/\[([^\]]+)\]\([^)]+\)/g, '$1')
    .replace(/`([^`]+)`/g, '$1')
    // Emphasis has to actually be emphasis. The previous `(\*|_)(.*?)\1`
    // treated any two delimiters as a pair, so `my_file_name` became
    // `myfilename` and `2 * 3 * 4` became `2  3  4`. CommonMark does not treat
    // intraword underscores as emphasis, and a delimiter followed by a space
    // opens nothing — both rules are encoded here.
    .replace(/\*\*(\S(?:[\s\S]*?\S)?)\*\*/g, '$1')
    .replace(/__(\S(?:[\s\S]*?\S)?)__/g, '$1')
    .replace(/\*(\S(?:[\s\S]*?\S)?)\*/g, '$1')
    .replace(/(?<!\w)_(\S(?:[\s\S]*?\S)?)_(?!\w)/g, '$1')
    .replace(/~~(\S(?:[\s\S]*?\S)?)~~/g, '$1')
    // Keep autolink targets, and only drop things that are really tags:
    // `<[^>]+>` also ate `a < b > c` and every `<https://…>` link.
    .replace(/<((?:https?|mailto):[^>\s]+)>/gi, '$1')
    .replace(/<\/?[A-Za-z][^>]*>/g, '')
    .replace(/\\([\\`*{}\[\]()#+.!_>-])/g, '$1')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .trim();
}

function joinSoftLines(lines: string[]): string {
  return lines.reduce((result, line) => {
    const current = line.trim();
    if (!result) return current;
    return result + (endsWithCjk(result) && startsWithCjk(current) ? '' : ' ') + current;
  }, '');
}

function startsWithCjk(value: string): boolean {
  return /^[\u3000-\u9fff]/u.test(value);
}

function endsWithCjk(value: string): boolean {
  return /[\u3000-\u9fff]$/u.test(value);
}

function stripQuotes(value: string): string {
  return value.replace(/^(?:"([\s\S]*)"|'([\s\S]*)')$/, '$1$2');
}

function stripAngleBrackets(value: string): string {
  return value.startsWith('<') && value.endsWith('>') ? value.slice(1, -1) : value;
}

function slugify(value: string): string {
  const slug = value.toLowerCase().replace(/[^\p{L}\p{N}]+/gu, '-').replace(/^-|-$/g, '');
  return slug || 'untitled';
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
