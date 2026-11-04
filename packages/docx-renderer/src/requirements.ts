/**
 * Checking a journal submission against what its template requires.
 *
 * The template is the spec: it says `Title: [Max. 20 words]`, `Keywords: [At
 * least 3]`, which sections it expects and which fixed table it opens with.
 * Layout alone does not satisfy those, and a submission that misses one is
 * desk-rejected however good it looks.
 *
 * The checks report; they never edit the content. A caller decides what to do
 * with an over-long title, and the message says what the limit is.
 */
import type { JournalArticle, JournalRequirements } from '@openthesis/document-schema';
import { countWords } from './references.js';

/** One unmet requirement. */
export interface RequirementIssue {
  /** Where the problem is, e.g. `meta.title`. */
  path: string;
  message: string;
}

/** How the section titles in the content read, for matching against the template. */
function normalizeSectionTitle(title: string): string {
  return title.toLowerCase().replace(/[^a-z0-9\u4e00-\u9fff]+/g, ' ').trim();
}

/**
 * Check a journal article against its template's requirements.
 *
 * A requirement the template does not state is not checked: the template is the
 * only source of truth here, and inventing a limit would be worse than missing
 * one.
 */
export function validateJournalRequirements(
  article: JournalArticle,
  requirements: JournalRequirements | undefined,
): RequirementIssue[] {
  const issues: RequirementIssue[] = [];
  if (!requirements) return issues;

  if (requirements.maxTitleWords !== undefined) {
    const words = countWords(article.meta.title ?? '');
    if (words > requirements.maxTitleWords) {
      issues.push({
        path: 'meta.title',
        message: `the journal accepts at most ${requirements.maxTitleWords} words; this title has ${words}`,
      });
    }
  }

  if (requirements.maxAbstractWords !== undefined) {
    const words = countWords(article.meta.abstract ?? '');
    if (words > requirements.maxAbstractWords) {
      issues.push({
        path: 'meta.abstract',
        message: `the journal accepts at most ${requirements.maxAbstractWords} words; this abstract has ${words}`,
      });
    }
  }

  if (requirements.minKeywords !== undefined) {
    const count = (article.meta.keywords ?? []).length;
    if (count < requirements.minKeywords) {
      issues.push({
        path: 'meta.keywords',
        message: `the journal asks for at least ${requirements.minKeywords} keywords; this article lists ${count}`,
      });
    }
  }

  if (requirements.minReferences !== undefined) {
    const count = (article.backMatter?.references ?? []).length;
    if (count < requirements.minReferences) {
      issues.push({
        path: 'backMatter.references',
        message: `the journal asks for at least ${requirements.minReferences} references; this article lists ${count}`,
      });
    }
  }

  if (requirements.requiredSections && requirements.requiredSections.length > 0) {
    const present = new Set(
      (article.sections ?? []).map(section => normalizeSectionTitle(section.title ?? '')),
    );
    for (const required of requirements.requiredSections) {
      const wanted = normalizeSectionTitle(required);
      // A title that contains the required name satisfies it: a journal that
      // asks for `Build Instructions` accepts `Build Instructions` with a
      // sentence after it.
      const found = [...present].some(title => title === wanted || title.includes(wanted) || wanted.includes(title));
      if (!found) {
        issues.push({
          path: 'sections',
          message: `the journal expects a section titled "${required}"; none of the sections matches`,
        });
      }
    }
  }

  if (requirements.specificationLabels && requirements.specificationLabels.length > 0) {
    const supplied = new Set((article.specifications ?? []).map(row => normalizeSectionTitle(row.label)));
    for (const label of requirements.specificationLabels) {
      if (!supplied.has(normalizeSectionTitle(label))) {
        issues.push({
          path: 'specifications',
          message: `the journal's specifications table has a "${label}" row; it is missing`,
        });
      }
    }
  }

  return issues;
}
