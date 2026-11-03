// ============================================================
// @openthesis/document-schema — content validation
// ============================================================
// The renderer used to be the only thing that inspected content, so a
// malformed document failed deep inside the OOXML builder with a message that
// named neither the file nor the offending block. This validator reports every
// problem it finds, each with a path, before any rendering starts.
// ============================================================

import type { OpenThesisDocument } from './index.js';

export interface ValidationIssue {
  /** Dotted path to the offending value, e.g. `sections[0].content[3].headers`. */
  path: string;
  message: string;
}

const DOCUMENT_TYPES = ['thesis', 'journal', 'official'] as const;

const THESIS_SECTION_TYPES = [
  'chapter', 'abstract', 'acknowledgement', 'toc',
  'list_of_figures', 'list_of_tables', 'references', 'appendix',
] as const;

const THESIS_DEGREES = ['bachelor', 'master', 'doctor'] as const;

const JOURNAL_SECTION_TYPES = [
  'introduction', 'methods', 'results', 'discussion', 'conclusion',
  'materials', 'background', 'related_work', 'experiment', 'analysis',
  'appendix', 'supplementary',
] as const;

const OFFICIAL_CATEGORIES = [
  '通知', '通告', '报告', '请示', '批复', '函', '纪要', '决定', '意见',
  '通报', '议案', '公告', '命令', '决议', '公报',
] as const;

const DOCUMENT_CATEGORIES = ['journal', 'book', 'conference', 'thesis', 'web', 'standard', 'other'] as const;

/** Block types that need only the listed non-empty string fields. */
const REQUIRED_STRING_FIELDS: Record<string, readonly string[]> = {
  heading1: ['text'],
  heading2: ['text'],
  heading3: ['text'],
  heading4: ['text'],
  paragraph: ['text'],
  paragraph_no_indent: ['text'],
  centered_text: ['text'],
  equation: ['latex'],
  equation_numbered: ['latex'],
  figure: ['path', 'caption'],
  list_item: ['text'],
  code_block: ['text'],
  blockquote: ['text'],
  red_header: ['text'],
  document_number: ['text'],
  recipient_line: ['text'],
  signature_block: ['authority', 'date'],
};

/** Block types handled by dedicated logic below. */
const SPECIAL_BLOCK_TYPES = ['table', 'spacer', 'page_break', 'horizontal_rule', 'attachment_note'] as const;

const BLOCK_TYPES = [
  ...Object.keys(REQUIRED_STRING_FIELDS),
  ...SPECIAL_BLOCK_TYPES,
];

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isNonEmptyString(value: unknown): boolean {
  return typeof value === 'string' && value.trim().length > 0;
}

function isStringArray(value: unknown): value is string[] {
  return Array.isArray(value) && value.every(item => typeof item === 'string');
}

function describe(value: unknown): string {
  if (value === undefined) return 'missing';
  if (value === null) return 'null';
  if (typeof value === 'string') {
    if (value.length === 0) return 'an empty string';
    const shown = value.length > 40 ? `${value.slice(0, 40)}…` : value;
    return JSON.stringify(shown);
  }
  if (Array.isArray(value)) return `an array of ${value.length}`;
  return `a ${typeof value}`;
}

function quoteList(values: readonly string[]): string {
  return values.map(value => `"${value}"`).join(', ');
}

/**
 * Validate a parsed content document.
 *
 * @returns every issue found, in document order. An empty array means valid.
 */
export function validateDocument(input: unknown): ValidationIssue[] {
  const issues: ValidationIssue[] = [];
  const report = (path: string, message: string) => { issues.push({ path, message }); };

  if (!isObject(input)) {
    report('', `expected a document object, received ${describe(input)}`);
    return issues;
  }

  const type = input.type;
  if (typeof type !== 'string' || !DOCUMENT_TYPES.includes(type as typeof DOCUMENT_TYPES[number])) {
    report('type', `must be one of ${quoteList(DOCUMENT_TYPES)}, received ${describe(type)}`);
    return issues;
  }

  const meta = input.meta;
  if (!isObject(meta)) {
    report('meta', `expected an object, received ${describe(meta)}`);
  } else if (!isNonEmptyString(meta.title)) {
    report('meta.title', `must be a non-empty string, received ${describe(meta.title)}`);
  }

  if (type === 'thesis') validateThesis(input, report);
  else if (type === 'journal') validateJournal(input, report);
  else validateOfficial(input, report);

  return issues;
}

/** Throw a single readable error listing every issue. */
export function assertValidDocument(input: unknown): asserts input is OpenThesisDocument {
  const issues = validateDocument(input);
  if (issues.length === 0) return;

  const detail = issues
    .map(issue => `  ${issue.path || '<document>'}: ${issue.message}`)
    .join('\n');
  throw new Error(
    `Invalid document content — ${issues.length} problem${issues.length === 1 ? '' : 's'}:\n${detail}`,
  );
}

// ── Per-document-type rules ───────────────────────────────

type Report = (path: string, message: string) => void;

function validateThesis(input: Record<string, unknown>, report: Report): void {
  const meta = input.meta;
  if (isObject(meta) && meta.degree !== undefined
      && !THESIS_DEGREES.includes(meta.degree as typeof THESIS_DEGREES[number])) {
    report('meta.degree', `must be one of ${quoteList(THESIS_DEGREES)}, received ${describe(meta.degree)}`);
  }

  requireArray(input, 'cover', report, (block, path) => validateBlock(block, path, report));
  validateSections(input.sections, 'sections', THESIS_SECTION_TYPES, report);

  const backMatter = input.backMatter;
  if (backMatter !== undefined) {
    if (!isObject(backMatter)) {
      report('backMatter', `expected an object, received ${describe(backMatter)}`);
    } else {
      validateReferences(backMatter.references, 'backMatter.references', report);
      validateSections(backMatter.appendices, 'backMatter.appendices', THESIS_SECTION_TYPES, report);
    }
  }
}

function validateJournal(input: Record<string, unknown>, report: Report): void {
  const meta = input.meta;
  if (isObject(meta)) {
    const authors = meta.authors;
    if (!Array.isArray(authors)) {
      report('meta.authors', `must be an array, received ${describe(authors)}`);
    } else if (authors.length === 0) {
      report('meta.authors', 'must list at least one author');
    } else {
      authors.forEach((author, index) => {
        const path = `meta.authors[${index}]`;
        if (!isObject(author)) {
          report(path, `expected an object, received ${describe(author)}`);
          return;
        }
        if (!isNonEmptyString(author.name)) {
          report(`${path}.name`, `must be a non-empty string, received ${describe(author.name)}`);
        }
        const affiliations = author.affiliations;
        if (!Array.isArray(affiliations)) {
          report(`${path}.affiliations`, `must be an array, received ${describe(affiliations)}`);
        } else {
          affiliations.forEach((affiliation, affIndex) => {
            const affPath = `${path}.affiliations[${affIndex}]`;
            if (!isObject(affiliation)) {
              report(affPath, `expected an object, received ${describe(affiliation)}`);
            } else if (!isNonEmptyString(affiliation.institution)) {
              report(`${affPath}.institution`, `must be a non-empty string, received ${describe(affiliation.institution)}`);
            }
          });
        }
      });
    }
    validateKeywords(meta.keywords, 'meta.keywords', report);
  }

  validateSections(input.sections, 'sections', JOURNAL_SECTION_TYPES, report);

  const backMatter = input.backMatter;
  if (backMatter !== undefined) {
    if (!isObject(backMatter)) {
      report('backMatter', `expected an object, received ${describe(backMatter)}`);
    } else {
      validateReferences(backMatter.references, 'backMatter.references', report);
    }
  }
}

function validateOfficial(input: Record<string, unknown>, report: Report): void {
  const meta = input.meta;
  if (isObject(meta)) {
    // GB/T 9704-2012 requires all of these to be present.
    for (const field of ['issuingAuthority', 'documentNumber', 'date'] as const) {
      if (!isNonEmptyString(meta[field])) {
        report(`meta.${field}`, `must be a non-empty string, received ${describe(meta[field])}`);
      }
    }

    const category = meta.documentCategory;
    if (typeof category !== 'string' || !OFFICIAL_CATEGORIES.includes(category as typeof OFFICIAL_CATEGORIES[number])) {
      report('meta.documentCategory', `must be one of ${quoteList(OFFICIAL_CATEGORIES)}, received ${describe(category)}`);
    }

    const recipients = meta.primaryRecipients;
    if (!isStringArray(recipients)) {
      report('meta.primaryRecipients', `must be an array of strings, received ${describe(recipients)}`);
    } else if (recipients.length === 0) {
      report('meta.primaryRecipients', 'must name at least one primary recipient');
    }

    if (meta.ccRecipients !== undefined && !isStringArray(meta.ccRecipients)) {
      report('meta.ccRecipients', `must be an array of strings, received ${describe(meta.ccRecipients)}`);
    }
  }

  requireArray(input, 'body', report, (block, path) => validateBlock(block, path, report));

  const attachments = input.attachments;
  if (attachments !== undefined) {
    if (!Array.isArray(attachments)) {
      report('attachments', `must be an array, received ${describe(attachments)}`);
    } else {
      attachments.forEach((attachment, index) => {
        const path = `attachments[${index}]`;
        if (!isObject(attachment)) {
          report(path, `expected an object, received ${describe(attachment)}`);
          return;
        }
        if (!Number.isInteger(attachment.order)) {
          report(`${path}.order`, `must be an integer, received ${describe(attachment.order)}`);
        }
        if (!isNonEmptyString(attachment.title)) {
          report(`${path}.title`, `must be a non-empty string, received ${describe(attachment.title)}`);
        }
        if (attachment.content !== undefined) {
          requireArray(attachment, 'content', report, (block, blockPath) =>
            validateBlock(block, `${path}.${blockPath}`, report));
        }
      });
    }
  }
}

// ── Shared pieces ─────────────────────────────────────────

function requireArray(
  owner: Record<string, unknown>,
  field: string,
  report: Report,
  each: (item: unknown, path: string) => void,
): void {
  const value = owner[field];
  if (!Array.isArray(value)) {
    report(field, `must be an array, received ${describe(value)}`);
    return;
  }
  value.forEach((item, index) => each(item, `${field}[${index}]`));
}

function validateSections(
  value: unknown,
  path: string,
  allowedTypes: readonly string[],
  report: Report,
): void {
  if (value === undefined) {
    if (path === 'sections') report(path, 'is required');
    return;
  }
  if (!Array.isArray(value)) {
    report(path, `must be an array, received ${describe(value)}`);
    return;
  }

  value.forEach((section, index) => {
    const sectionPath = `${path}[${index}]`;
    if (!isObject(section)) {
      report(sectionPath, `expected an object, received ${describe(section)}`);
      return;
    }
    if (!isNonEmptyString(section.id)) {
      report(`${sectionPath}.id`, `must be a non-empty string, received ${describe(section.id)}`);
    }
    if (!isNonEmptyString(section.title)) {
      report(`${sectionPath}.title`, `must be a non-empty string, received ${describe(section.title)}`);
    }
    const sectionType = section.type;
    if (typeof sectionType !== 'string' || !allowedTypes.includes(sectionType)) {
      report(`${sectionPath}.type`, `must be one of ${quoteList(allowedTypes)}, received ${describe(sectionType)}`);
    }
    requireArray(section, 'content', report, (block, blockPath) =>
      validateBlock(block, `${sectionPath}.${blockPath}`, report));
    validateSections(section.subsections, `${sectionPath}.subsections`, allowedTypes, report);
  });
}

function validateReferences(value: unknown, path: string, report: Report): void {
  if (value === undefined) return;
  if (!Array.isArray(value)) {
    report(path, `must be an array, received ${describe(value)}`);
    return;
  }
  value.forEach((reference, index) => {
    const referencePath = `${path}[${index}]`;
    if (!isObject(reference)) {
      report(referencePath, `expected an object, received ${describe(reference)}`);
      return;
    }
    for (const field of ['id', 'text'] as const) {
      if (!isNonEmptyString(reference[field])) {
        report(`${referencePath}.${field}`, `must be a non-empty string, received ${describe(reference[field])}`);
      }
    }
    const referenceType = reference.type;
    if (referenceType !== undefined && !DOCUMENT_CATEGORIES.includes(referenceType as typeof DOCUMENT_CATEGORIES[number])) {
      report(`${referencePath}.type`, `must be one of ${quoteList(DOCUMENT_CATEGORIES)}, received ${describe(referenceType)}`);
    }
  });
}

function validateKeywords(value: unknown, path: string, report: Report): void {
  if (value === undefined) return;
  if (!isStringArray(value)) {
    report(path, `must be an array of strings, received ${describe(value)}`);
  }
}

function validateBlock(value: unknown, path: string, report: Report): void {
  if (!isObject(value)) {
    report(path, `expected a block object, received ${describe(value)}`);
    return;
  }

  const blockType = value.type;
  if (typeof blockType !== 'string' || !BLOCK_TYPES.includes(blockType)) {
    report(`${path}.type`, `must be one of ${quoteList(BLOCK_TYPES)}, received ${describe(blockType)}`);
    return;
  }

  for (const field of REQUIRED_STRING_FIELDS[blockType] ?? []) {
    if (!isNonEmptyString(value[field])) {
      report(`${path}.${field}`, `${blockType} requires a non-empty string, received ${describe(value[field])}`);
    }
  }

  switch (blockType) {
    case 'table':
      validateTable(value, path, report);
      break;
    case 'spacer':
      if (!Number.isInteger(value.lines) || (value.lines as number) < 0) {
        report(`${path}.lines`, `must be a non-negative integer, received ${describe(value.lines)}`);
      }
      break;
    case 'recipient_line':
      if (value.recipientType !== 'primary' && value.recipientType !== 'cc') {
        report(`${path}.recipientType`, `must be "primary" or "cc", received ${describe(value.recipientType)}`);
      }
      break;
    case 'attachment_note':
      if (!isStringArray(value.attachments)) {
        report(`${path}.attachments`, `must be an array of strings, received ${describe(value.attachments)}`);
      }
      break;
    default:
      break;
  }
}

function validateTable(value: Record<string, unknown>, path: string, report: Report): void {
  const headers = value.headers;
  if (!isStringArray(headers)) {
    report(`${path}.headers`, `must be an array of strings, received ${describe(headers)}`);
    return;
  }
  if (headers.length === 0) {
    report(`${path}.headers`, 'must define at least one column');
    return;
  }

  const data = value.data;
  if (!Array.isArray(data)) {
    report(`${path}.data`, `must be an array, received ${describe(data)}`);
  } else {
    data.forEach((row, rowIndex) => {
      const rowPath = `${path}.data[${rowIndex}]`;
      if (!Array.isArray(row)) {
        report(rowPath, `expected an array of cells, received ${describe(row)}`);
        return;
      }
      // The renderer pads and truncates silently, which loses data without a word.
      if (row.length !== headers.length) {
        report(rowPath, `has ${row.length} cell${row.length === 1 ? '' : 's'} but the table declares ${headers.length} column${headers.length === 1 ? '' : 's'}`);
      }
      row.forEach((cell, cellIndex) => {
        if (typeof cell !== 'string') {
          report(`${rowPath}[${cellIndex}]`, `must be a string, received ${describe(cell)}`);
        }
      });
    });
  }

  const columnWidths = value.columnWidths;
  if (columnWidths !== undefined) {
    if (!Array.isArray(columnWidths)) {
      report(`${path}.columnWidths`, `must be an array of numbers, received ${describe(columnWidths)}`);
    } else if (columnWidths.length !== headers.length) {
      report(`${path}.columnWidths`, `has ${columnWidths.length} entries but the table declares ${headers.length} column${headers.length === 1 ? '' : 's'}`);
    } else {
      columnWidths.forEach((width, index) => {
        if (typeof width !== 'number' || !Number.isFinite(width) || width <= 0) {
          report(`${path}.columnWidths[${index}]`, `must be a positive number, received ${describe(width)}`);
        }
      });
    }
  }
}

/**
 * Validate the legacy `{cover_blocks, body_blocks}` format.
 *
 * It is still accepted by the CLI and still shipped as `examples/`, so it needs
 * the same guard rails as the structured schema.
 */
export function validateLegacyDocument(input: unknown): ValidationIssue[] {
  const issues: ValidationIssue[] = [];
  const report = (path: string, message: string) => { issues.push({ path, message }); };

  if (!isObject(input)) {
    report('', `expected a document object, received ${describe(input)}`);
    return issues;
  }
  if (!isNonEmptyString(input.title)) {
    report('title', `must be a non-empty string, received ${describe(input.title)}`);
  }
  if (input.report_header !== undefined && typeof input.report_header !== 'string') {
    report('report_header', `must be a string, received ${describe(input.report_header)}`);
  }

  for (const field of ['cover_blocks', 'body_blocks'] as const) {
    requireArray(input, field, report, (block, path) => validateBlock(block, path, report));
  }

  return issues;
}

/** Format issues for a CLI error message. */
export function formatIssues(issues: ValidationIssue[]): string {
  return issues.map(issue => `  ${issue.path || '<document>'}: ${issue.message}`).join('\n');
}
