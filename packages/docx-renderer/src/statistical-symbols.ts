// ============================================================
// @openthesis/docx-renderer — Statistical Symbol Formatting
// ============================================================
// Academic manuscripts follow a strict convention for symbols:
//
//   - Statistics and population parameters — t, F, p, P, M, SD,
//     SE, Mdn, r, z, d, n, N, df, χ², and Greek parameters such
//     as μ (\mu) and σ (\sigma) — are set in Times New Roman
//     italic, never bold.
//   - Vector and matrix symbols are set in Times New Roman
//     bold italic.
//
// Plain paragraph/table text carries no inline markup, so the
// renderer has to *recognise* these symbols in running text.
// Recognition is deliberately conservative: a bare letter is
// only treated as a symbol when the surrounding context marks
// it as one (a comparison or assignment follows, it heads a
// "P 值 / t 检验" label, it is flanked by CJK text, or it is the
// whole cell). Letters inside ordinary words are never split,
// so normal English prose is left untouched.
// ============================================================

export type TextSegmentKind = 'plain' | 'statistic' | 'vector';

export interface TextSegment {
  text: string;
  kind: TextSegmentKind;
}

// ── Symbol tables ─────────────────────────────────────────

const GREEK_BY_LATEX: Record<string, string> = {
  alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε',
  varepsilon: 'ε', zeta: 'ζ', eta: 'η', theta: 'θ', vartheta: 'ϑ',
  iota: 'ι', kappa: 'κ', lambda: 'λ', mu: 'μ', nu: 'ν', xi: 'ξ',
  pi: 'π', rho: 'ρ', sigma: 'σ', tau: 'τ', upsilon: 'υ', phi: 'φ',
  varphi: 'φ', chi: 'χ', psi: 'ψ', omega: 'ω',
  Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ', Xi: 'Ξ',
  Pi: 'Π', Sigma: 'Σ', Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω',
};

const GREEK_CHARS = new Set(Object.values(GREEK_BY_LATEX));

const SUPERSCRIPT_DIGITS: Record<string, string> = {
  '0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴',
  '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹',
};

/** Symbols recognised when a comparison / assignment follows. */
const COMPARISON_SYMBOLS = new Set([
  't', 'F', 'p', 'P', 'r', 'z', 'd', 'M', 'N', 'n', 'U', 'W',
]);

/** Symbols recognised when followed by 值 (e.g. the header "P 值"). */
const VALUE_LABEL_SYMBOLS = new Set(['p', 'P', 't', 'F', 'r', 'M']);

/** Symbols recognised when followed by 检验 / 分布 / test. */
const TEST_LABEL_SYMBOLS = new Set(['t', 'F', 'z', 'U']);

/**
 * Symbols recognised as standalone tokens flanked by CJK text.
 * `M` is excluded here (it needs `=` / 值 / a whole cell) so that
 * strings such as "M码" are not mistaken for a mean.
 */
const CJK_FLANKED_SYMBOLS = new Set(['t', 'F', 'p', 'P', 'r', 'n', 'N']);

/** Multi-letter statistics, recognised as standalone tokens. */
const MULTI_LETTER_SYMBOLS = ['Mdn', 'SD', 'SE', 'df'];

/** Words after which a single letter denotes a vector / matrix. */
const VECTOR_KEYWORDS = ['向量', '矩阵', '矢量'];

// ── Character helpers ─────────────────────────────────────

function isAsciiAlnum(ch: string | undefined): boolean {
  return ch !== undefined && /[A-Za-z0-9]/.test(ch);
}

function isCjk(ch: string | undefined): boolean {
  return ch !== undefined
    && /[\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]/.test(ch);
}

function isCjkPunctuation(ch: string | undefined): boolean {
  return ch !== undefined && '，。、；：？！（）【】《》“”‘’—…·'.includes(ch);
}

function isCjkContext(ch: string | undefined): boolean {
  return isCjk(ch) || isCjkPunctuation(ch);
}

/** True when the text before `index` is CJK (allowing one space). */
function cjkBefore(text: string, index: number): boolean {
  if (isCjkContext(text[index - 1])) return true;
  return text[index - 1] === ' ' && isCjkContext(text[index - 2]);
}

/** True when the text at/after `index` is CJK (allowing one space). */
function cjkAfter(text: string, index: number): boolean {
  if (index >= text.length) return false;
  if (isCjkContext(text[index])) return true;
  return text[index] === ' ' && isCjkContext(text[index + 1]);
}

// ── Matchers (each anchored at `index`) ───────────────────

interface Match {
  /** Characters consumed from the source text. */
  consumed: number;
  /** Text to display (LaTeX commands are converted to Unicode). */
  display: string;
  kind: TextSegmentKind;
}

function matchExplicitVector(text: string, index: number): Match | null {
  const braced = /^\\(?:mathbf|boldsymbol|bm|vec)\{([^{}]+)\}/.exec(text.slice(index));
  if (braced) {
    return { consumed: braced[0].length, display: braced[1], kind: 'vector' };
  }
  const bare = /^\\(?:mathbf|boldsymbol|bm)\s*([A-Za-z])/.exec(text.slice(index));
  if (bare) {
    return { consumed: bare[0].length, display: bare[1], kind: 'vector' };
  }
  return null;
}

function matchLatexGreek(text: string, index: number): Match | null {
  const command = /^\\([A-Za-z]+)/.exec(text.slice(index));
  if (!command) return null;
  const greek = GREEK_BY_LATEX[command[1]];
  if (!greek) return null;
  let consumed = command[0].length;
  let display = greek;
  const rest = text.slice(index + consumed);
  const superscript = /^\^\{(\d)\}|^\^(\d)|^²/.exec(rest);
  if (superscript) {
    const digit = superscript[1] ?? superscript[2] ?? '2';
    display += SUPERSCRIPT_DIGITS[digit] ?? digit;
    consumed += superscript[0].length;
  }
  return { consumed, display, kind: 'statistic' };
}

function matchGreekChar(text: string, index: number): Match | null {
  const ch = text[index];
  if (!GREEK_CHARS.has(ch)) return null;
  if (isAsciiAlnum(text[index - 1]) || GREEK_CHARS.has(text[index - 1] ?? '')) return null;
  const rest = text.slice(index + 1);
  const superscript = /^\^\{(\d)\}|^\^(\d)|^²/.exec(rest);
  if (superscript) {
    const digit = superscript[1] ?? superscript[2] ?? '2';
    return {
      consumed: 1 + superscript[0].length,
      display: ch + (SUPERSCRIPT_DIGITS[digit] ?? digit),
      kind: 'statistic',
    };
  }
  // A lone Greek letter: the next character must not continue a word.
  if (isAsciiAlnum(text[index + 1]) || GREEK_CHARS.has(text[index + 1] ?? '')) return null;
  return { consumed: 1, display: ch, kind: 'statistic' };
}

function matchVectorKeyword(text: string, index: number): Match | null {
  const ch = text[index];
  if (!/[A-Za-z]/.test(ch) && !GREEK_CHARS.has(ch)) return null;
  if (isAsciiAlnum(text[index + 1])) return null;
  const before = text.slice(Math.max(0, index - 4), index);
  if (!VECTOR_KEYWORDS.some(keyword => new RegExp(`${keyword}\\s{0,2}$`).test(before))) {
    return null;
  }
  return { consumed: 1, display: ch, kind: 'vector' };
}

function matchMultiLetter(text: string, index: number): Match | null {
  if (isAsciiAlnum(text[index - 1])) return null;
  for (const symbol of MULTI_LETTER_SYMBOLS) {
    if (!text.startsWith(symbol, index)) continue;
    const next = text[index + symbol.length];
    if (isAsciiAlnum(next)) continue;
    // Followed directly by a CJK noun it is more likely a compound
    // word ("SD卡") than a statistic ("SD为…", "SD的值", "SD，…").
    if (isCjk(next) && !'为是的值'.includes(next)) continue;
    return { consumed: symbol.length, display: symbol, kind: 'statistic' };
  }
  return null;
}

function matchSingleLetter(text: string, index: number): Match | null {
  const ch = text[index];
  if (!COMPARISON_SYMBOLS.has(ch) && !CJK_FLANKED_SYMBOLS.has(ch)) return null;
  if (isAsciiAlnum(text[index - 1]) || text[index - 1] === '_') return null;

  const rest = text.slice(index + 1);
  const statistic: Match = { consumed: 1, display: ch, kind: 'statistic' };

  // Strong context: t(12) = …, F(1, 23) = …, p < 0.05, M = 3.20 …
  if (COMPARISON_SYMBOLS.has(ch)
    && /^(\(\s*[\d\s,.]+\s*\))?\s*(=|<|>|≤|≥|≈|~)/.test(rest)) {
    return statistic;
  }
  // Label context: "P 值", "P值", "t 值" (table headers, captions).
  if (VALUE_LABEL_SYMBOLS.has(ch) && /^\s*值/.test(rest)) return statistic;
  // Test / distribution context: "t 检验", "F分布", "t test".
  if (TEST_LABEL_SYMBOLS.has(ch) && (/^\s*(检验|分布|试验)/.test(rest) || /^\s+test\b/.test(rest))) {
    return statistic;
  }
  // "P value(s)" in English prose.
  if ((ch === 'p' || ch === 'P') && /^\s+values?\b/.test(rest)) return statistic;
  // The symbol is the whole (trimmed) text — e.g. a one-letter table cell.
  if (text.trim() === ch) return statistic;
  // CJK-flanked standalone token: "结果显示P为…", "其t值…", "样本量n为30".
  if (CJK_FLANKED_SYMBOLS.has(ch)
    && (index === 0 || cjkBefore(text, index))
    && (index + 1 >= text.length || cjkAfter(text, index + 1))) {
    return statistic;
  }
  return null;
}

function matchAt(text: string, index: number): Match | null {
  if (text[index] === '\\') {
    return matchExplicitVector(text, index) ?? matchLatexGreek(text, index);
  }
  return matchVectorKeyword(text, index)
    ?? matchGreekChar(text, index)
    ?? matchMultiLetter(text, index)
    ?? matchSingleLetter(text, index);
}

// ── Public API ────────────────────────────────────────────

/**
 * Split plain text into plain / statistic / vector segments.
 * Adjacent segments of the same kind are merged. Concatenating
 * the segment texts reproduces the input, except that LaTeX
 * forms (`\mu`, `\chi^2`, `\mathbf{A}`) are converted to their
 * Unicode display forms (`μ`, `χ²`, `A`).
 */
export function splitStatisticalSymbols(text: string): TextSegment[] {
  const segments: TextSegment[] = [];
  const push = (value: string, kind: TextSegmentKind) => {
    if (!value) return;
    const last = segments[segments.length - 1];
    if (last && last.kind === kind) last.text += value;
    else segments.push({ text: value, kind });
  };

  let index = 0;
  while (index < text.length) {
    const match = matchAt(text, index);
    if (match) {
      push(match.display, match.kind);
      index += match.consumed;
    } else {
      push(text[index], 'plain');
      index += 1;
    }
  }
  return segments;
}
