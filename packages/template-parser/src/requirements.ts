/**
 * What a journal template states it requires.
 *
 * A journal template is usually a filled-in instruction sheet rather than a style
 * table: `Title: [Max. 20 words …]`, `Keywords: [At least 3 …]`, followed by the
 * sections the journal expects and, often, a fixed opening table. None of that is
 * layout, so a renderer that only copies fonts and margins can produce a document
 * that looks right and gets desk-rejected.
 *
 * Everything here is read out of the template's own text. Where the text does not
 * state a limit, no limit is invented.
 */
import type { JournalRequirements } from '@openthesis/document-schema';

/** The text of every paragraph in the template, in document order. */
export type ParagraphTexts = string[];

/** `Max. 20 words` / `maximum of 200 words` / `不超过 200 字`. */
function wordLimit(text: string, near: string): number | undefined {
  const lower = text.toLowerCase();
  const index = lower.indexOf(near);
  if (index < 0) return undefined;
  const window = text.slice(index, index + near.length + 180);
  const match = /(?:max(?:imum)?(?:\s+of)?|no\s+more\s+than|at\s+most|不超过|最多)\s*[.:]?\s*(\d+)\s*(?:words?|字)?/i.exec(window);
  return match ? Number(match[1]) : undefined;
}

/** `At least 3 keywords` / `至少 3 个关键词`. */
function minimumNear(text: string, near: string): number | undefined {
  const lower = text.toLowerCase();
  const index = lower.indexOf(near);
  if (index < 0) return undefined;
  const window = text.slice(index, index + 240);
  const match = /(?:at\s+least|minimum(?:\s+of)?|no\s+fewer\s+than|至少|不少于)\s*(\d+)/i.exec(window);
  return match ? Number(match[1]) : undefined;
}

/**
 * The sections a template expects, in the order it prints them.
 *
 * The instruction sheet's section headings are short lines that name a task —
 * `Design files`, `Build Instructions`, `Validation and Characterization` — and
 * that are followed by prose rather than by more labels. Table rows are short
 * lines too, which is why a bare length test is not enough: a line is a heading
 * when the next non-empty paragraph is a sentence.
 */
function requiredSections(texts: ParagraphTexts): string[] {
  const sections: string[] = [];
  for (let i = 0; i < texts.length; i += 1) {
    const line = texts[i].trim();
    if (!line || line.length > 64) continue;
    if (/[.;,:]/.test(line)) continue;
    // A cell that glued two instructions together reads as a question with no
    // punctuation between them; a section heading never does.
    if (/\?/.test(line)) continue;
    if (line.includes('[')) continue;
    const next = texts.slice(i + 1).find(text => text.trim().length > 0)?.trim() ?? '';
    if (next.length < 40 || !/[.]/.test(next)) continue;
    if (!/(instructions?|validation|characteri[sz]ation|design\s+files?|build|operation|safety|reusability|limitations|conclusions?|introduction|methods?|results?|discussion|bill\s+of\s+materials|hardware\s+in\s+context)/i.test(line)) {
      continue;
    }
    const label = line.replace(/\s+/g, ' ').trim();
    if (!sections.some(existing => existing.toLowerCase() === label.toLowerCase())) {
      sections.push(label);
    }
  }
  return sections;
}

/**
 * The labels of the fixed opening table, when the template has one.
 *
 * HardwareX asks for a `Specifications table` whose left column lists labels
 * (Hardware name, Subject area, Hardware type, …) and whose right column the
 * author fills in. The labels are read by name, because the option lists under
 * them are also short lines and guessing which is which goes wrong.
 */
function specificationLabels(texts: ParagraphTexts): string[] {
  const start = texts.findIndex(text => /specifications?\s+table/i.test(text));
  if (start < 0) return [];

  const known = [
    /hardware\s+name/i,
    /subject\s+area/i,
    /hardware\s+type/i,
    /open\s+source\s+license/i,
    /cost\s+of\s+hardware/i,
    /source\s+file\s+repository/i,
  ];

  const labels: string[] = [];
  for (let i = start + 1; i < Math.min(texts.length, start + 80); i += 1) {
    const line = texts[i].trim();
    if (!line || line.length > 64) continue;
    if (/[.:;]/.test(line) || line.includes('[')) continue;
    if (/^hardware\s+in\s+context$/i.test(line)) break;
    if (known.some(pattern => pattern.test(line))) labels.push(line.replace(/\s+/g, ' ').trim());
  }
  return labels;
}

/** Read a journal template's stated requirements, or nothing if it states none. */
export function readJournalRequirements(
  texts: ParagraphTexts,
  documentType?: string,
): JournalRequirements | undefined {
  if (documentType !== 'journal') return undefined;

  const titleText = texts.find(text => /^\s*title\s*:/i.test(text)) ?? '';
  const abstractText = texts.find(text => /^\s*abstract\s*:/i.test(text)) ?? '';
  const keywordText = texts.find(text => /^\s*keywords?\s*:/i.test(text)) ?? '';
  const referenceText = texts.find(text => /references?\s*:/i.test(text)) ?? '';

  const maxTitleWords = wordLimit(titleText, 'title');
  const maxAbstractWords = wordLimit(abstractText, 'abstract');
  const minKeywords = minimumNear(keywordText, 'keyword');
  const minReferences = minimumNear(referenceText, 'reference');
  const sections = requiredSections(texts);
  const specs = specificationLabels(texts);

  const requirements: JournalRequirements = {
    ...(maxTitleWords ? { maxTitleWords } : {}),
    ...(maxAbstractWords ? { maxAbstractWords } : {}),
    ...(minKeywords ? { minKeywords } : {}),
    ...(minReferences ? { minReferences } : {}),
    ...(sections.length > 0 ? { requiredSections: sections } : {}),
    ...(specs.length > 0 ? { specificationLabels: specs } : {}),
  };

  // A template that states nothing is not a requirements sheet; do not invent one.
  return Object.keys(requirements).length > 0 ? requirements : undefined;
}
