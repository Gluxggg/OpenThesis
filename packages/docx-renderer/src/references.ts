/**
 * GB/T 7714-2015 参考文献著录 formatting.
 *
 * The author supplies the parts; this module supplies the punctuation, the
 * 文献类型标志 (`[M]`, `[J]`, …), the 等 / et al. truncation and the
 * `年, 卷(期): 页码` layout. Getting those right by hand for every entry is
 * exactly the kind of work a generator should do instead.
 *
 * Rule structure, from 《信息与文献 参考文献著录规则》(GB/T 7714-2015):
 *
 *   专著    主要责任者. 题名: 其他题名信息[M]. 版本项. 出版地: 出版者, 出版年: 页码.
 *   期刊    主要责任者. 题名[J]. 刊名, 年, 卷(期): 页码.
 *   会议    主要责任者. 题名[C]//论文集主要责任者. 论文集题名. 出版地: 出版者, 年: 页码.
 *   学位    主要责任者. 题名[D]. 保存地: 保存者, 年.
 *   标准    主要责任者. 题名[S]. 出版地: 出版者, 年.
 *   专利    专利申请者. 专利题名: 专利号[P]. 公告日期.
 *   报告    主要责任者. 题名[R]. 出版地: 出版者, 年.
 *   报纸    主要责任者. 题名[N]. 报纸名, 出版日期(版次).
 *   电子    主要责任者. 题名[EB/OL]. (更新日期)[引用日期]. 访问路径.
 */
import type { Reference, ReferenceKind } from '@openthesis/document-schema';

/** The 文献类型标志 each kind is filed under. */
const TYPE_MARKER: Record<ReferenceKind, string> = {
  book: 'M',
  journal: 'J',
  chapter: 'M',   // 析出文献 from a 专著: 题名[M]//专著题名
  conference: 'C',
  thesis: 'D',
  standard: 'S',
  patent: 'P',
  report: 'R',
  newspaper: 'N',
  electronic: 'EB/OL',
  database: 'DB/OL',
  software: 'CP/DK',
  archive: 'A',
  other: 'Z',
};

/** How many authors are printed before the list is truncated. */
const MAX_AUTHORS = 3;

/**
 * A trailing element, normalised so joining never produces `., ` or `..`.
 */
function joinParts(parts: (string | undefined)[], separator = '. '): string {
  const kept = parts
    .map(part => (part ?? '').trim())
    .filter(part => part.length > 0)
    // An element that already ends in punctuation must not gain a second mark.
    .map(part => part.replace(/[.。]+$/, ''));
  return kept.join(separator);
}

/**
 * Turn an English name round to GB/T 7714's "姓在前, 名在后" order.
 *
 * The standard writes an English author as `SMITH J` — family name, a space, then
 * the initials of the given names without full stops. The common spellings are
 * handled: `John Smith`, `John A. Smith`, `J. Smith`, `Smith, John`. A name already
 * in `SMITH J` form is left alone, and anything that does not look like a name is
 * returned unchanged rather than mangled.
 */
export function invertEnglishName(name: string): string {
  const trimmed = name.trim();
  if (!trimmed) return trimmed;

  // `Smith, John` — the comma already separates family name from given names.
  const commaSplit = trimmed.split(',').map(part => part.trim()).filter(Boolean);
  if (commaSplit.length === 2) {
    const [family, given] = commaSplit;
    const inverted = `${family.toUpperCase()} ${initials(given)}`.trim();
    return inverted;
  }

  const parts = trimmed.split(/\s+/).filter(Boolean);
  if (parts.length < 2) {
    // A single token is a family name with no given name: upper-case it.
    return trimmed.toUpperCase();
  }

  // `SMITH J` — already inverted: a shouting family name then initials.
  if (parts[0] === parts[0].toUpperCase() && parts.slice(1).every(part => /^[A-Z]\.?$/.test(part))) {
    return trimmed.replace(/\./g, '');
  }

  const family = parts[parts.length - 1];
  const given = parts.slice(0, -1).join(' ');
  return `${family.toUpperCase()} ${initials(given)}`.trim();
}

/** `John A.` becomes `J A` — initials, no full stops, space separated. */
function initials(given: string): string {
  return given
    .split(/\s+/)
    .filter(Boolean)
    .map(word => word.replace(/\./g, ''))
    .map(word => (word.length === 1 ? word : word[0]))
    .map(letter => letter.toUpperCase())
    .join(' ');
}

/**
 * 主要责任者, truncated per the standard.
 *
 * Three or fewer are all printed; more are cut to three and closed with 等 for a
 * Chinese entry or et al. for an English one. An English name is inverted to
 * `SMITH J` unless the author already wrote it that way.
 */
export function formatAuthors(
  authors: string[] | undefined,
  language?: 'zh' | 'en',
  invertNames = true,
): string {
  const raw = (authors ?? []).map(author => author.trim()).filter(Boolean);
  if (raw.length === 0) return '';
  const list = language === 'en' && invertNames ? raw.map(invertEnglishName) : raw;
  if (list.length <= MAX_AUTHORS) return list.join(', ');
  const suffix = language === 'en' ? ', et al' : ', 等';
  return `${list.slice(0, MAX_AUTHORS).join(', ')}${suffix}`;
}

/** `年, 卷(期): 页码` — whichever of those the entry actually has. */
function journalTail(ref: Reference): string {
  const year = ref.year === undefined ? '' : String(ref.year).trim();
  let volumeIssue = (ref.volume ?? '').trim();
  if ((ref.issue ?? '').trim()) {
    volumeIssue = volumeIssue ? `${volumeIssue}(${ref.issue!.trim()})` : `(${ref.issue!.trim()})`;
  }
  const head = [year, volumeIssue].filter(Boolean).join(', ');
  const pages = (ref.pages ?? '').trim();
  return pages ? `${head}: ${pages}` : head;
}

/** `出版地: 出版者, 年` — a missing place leaves the publisher unqualified. */
function imprint(ref: Reference): string {
  const place = (ref.place ?? '').trim();
  const publisher = (ref.publisher ?? '').trim();
  const year = ref.year === undefined ? '' : String(ref.year).trim();
  const where = place && publisher ? `${place}: ${publisher}` : place || publisher;
  return [where, year].filter(Boolean).join(', ');
}

/**
 * An entry of 参考文献, as GB/T 7714-2015 writes it.
 *
 * A `text` the author wrote wins: it is their entry, and re-formatting text they
 * already punctuated would be presumptuous.
 */
export function formatReference(ref: Reference): string {
  if ((ref.text ?? '').trim()) return ref.text!.trim();

  const kind: ReferenceKind = ref.type ?? 'journal';
  const marker = TYPE_MARKER[kind] ?? 'Z';
  const authors = formatAuthors(ref.authors, ref.language);
  const title = (ref.title ?? '').trim();
  const container = (ref.container ?? '').trim();

  // The 题名 carries the type marker: 题名[J], 题名[M]//专著题名.
  let titleBlock = title;
  if (kind === 'chapter' || kind === 'conference') {
    // 析出文献: the source document follows the marker.
    const source = joinParts([ref.publisher && !container ? undefined : container]);
    titleBlock = `${title}[${marker}]${source ? `//${source}` : '//'}`;
  } else {
    titleBlock = `${title}[${marker}]`;
  }

  let tail = '';
  switch (kind) {
    case 'journal':
      tail = [container, journalTail(ref)].filter(Boolean).join(', ');
      break;
    case 'newspaper':
      tail = [container, (ref.year === undefined ? '' : String(ref.year)), (ref.issue ?? '').trim()]
        .filter(Boolean)
        .join(', ');
      break;
    case 'chapter':
    case 'conference':
      tail = [imprint(ref), (ref.pages ?? '').trim()].filter(Boolean).join(': ');
      break;
    case 'thesis':
      tail = [imprint(ref)].filter(Boolean).join(' ');
      break;
    case 'patent':
      tail = [
        (ref.issue ?? '').trim(),
        (ref.year === undefined ? '' : String(ref.year)),
      ].filter(Boolean).join(', ');
      break;
    case 'electronic':
    case 'database':
    case 'software': {
      // 更新日期 and 引用日期 form one bracket group; the 访问路径 follows after a
      // full stop, which is how the standard prints an online entry.
      const dates = [
        ref.year === undefined ? '' : `(${ref.year})`,
        ref.accessDate ? `[${ref.accessDate}]` : '',
      ].join('');
      tail = [dates, (ref.url ?? '').trim()].filter(Boolean).join('. ');
      break;
    }
    default:
      tail = [imprint(ref), (ref.pages ?? '').trim()].filter(Boolean).join(': ');
      break;
  }

  // DOI closes the entry when there is no access path of its own.
  const identifier = (ref.url ?? '').trim() && kind !== 'electronic' && kind !== 'database' && kind !== 'software'
    ? (ref.url ?? '').trim()
    : (ref.doi ?? '').trim();

  const body = joinParts([authors, titleBlock, tail, identifier]);
  return body ? `${body}.` : '';
}
