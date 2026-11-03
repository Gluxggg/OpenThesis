// ============================================================
// @openthesis/equation-engine — LaTeX → Office Math Engine
// ============================================================
// V3: LaTeX math AST for native DOCX OMML rendering.
// Keeps Unicode fallback and optional Pandoc extraction for compatibility.
// ============================================================

import { execFileSync } from 'child_process';
import { writeFileSync, unlinkSync, existsSync, mkdirSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';
import { randomUUID } from 'crypto';
import AdmZip from 'adm-zip';

// ── Public API ────────────────────────────────────────────

export type LatexMathNode =
  | { type: 'run'; text: string }
  | { type: 'fraction'; numerator: LatexMathNode[]; denominator: LatexMathNode[]; bar?: boolean }
  | { type: 'radical'; children: LatexMathNode[]; degree?: LatexMathNode[] }
  | { type: 'script'; base: LatexMathNode[]; subScript?: LatexMathNode[]; superScript?: LatexMathNode[] }
  | { type: 'sum' | 'integral'; subScript?: LatexMathNode[]; superScript?: LatexMathNode[] }
  | { type: 'nary'; operator: string; subScript?: LatexMathNode[]; superScript?: LatexMathNode[] }
  | { type: 'accent'; accent: string; children: LatexMathNode[] }
  | { type: 'bar'; position: 'top' | 'bottom'; children: LatexMathNode[] }
  | { type: 'delimiter'; opening: string; closing: string; children: LatexMathNode[] }
  | { type: 'function'; name: string; children: LatexMathNode[] }
  | { type: 'limit'; name: string; subScript?: LatexMathNode[]; superScript?: LatexMathNode[] }
  | { type: 'overUnder'; base: LatexMathNode[]; over?: LatexMathNode[]; under?: LatexMathNode[] }
  | { type: 'matrix'; rows: LatexMathNode[][][]; opening: string; closing: string; environment: MatrixEnvironment };

export type MatrixEnvironment =
  | 'matrix' | 'pmatrix' | 'bmatrix' | 'Bmatrix'
  | 'vmatrix' | 'Vmatrix' | 'cases' | 'aligned';

/**
 * Parse a practical LaTeX math subset into a renderer-neutral tree.
 * The DOCX renderer maps this tree to native Office Math (OMML) elements.
 */
export function latexToMathAst(latex: string): LatexMathNode[] {
  return new LatexMathParser(latex).parse();
}

export interface OmmlConversion {
  /** OMML XML when pandoc produced it, otherwise Unicode plain text. */
  value: string;
  /** Which path produced `value`. */
  source: 'pandoc' | 'unicode';
  /** True when the Unicode fallback could not fully represent the input. */
  lossy: boolean;
  /** Why the pandoc path was unavailable, when it was. */
  reason?: string;
}

/**
 * Convert LaTeX to OMML, reporting which path was taken.
 *
 * Prefer this over `latexToOMML` when the caller needs to know whether the
 * result is native Office Math or a lossy approximation — the fallback used to
 * be entirely silent.
 */
export function convertLatexToOmml(latex: string): OmmlConversion {
  const pandoc = generateOMMLviaPandoc(latex);
  if (pandoc.omml) {
    return { value: pandoc.omml, source: 'pandoc', lossy: false };
  }

  const plain = latexToPlainText(latex);
  return {
    value: plain,
    source: 'unicode',
    lossy: isLossyConversion(latex, plain),
    reason: pandoc.error,
  };
}

/**
 * Render a LaTeX equation string to OMML XML (native Word equation).
 * Uses pandoc when available and falls back to Unicode plain text.
 *
 * The fallback cannot express everything; it warns when it is taken. Use
 * `convertLatexToOmml` if the caller needs to inspect that programmatically.
 */
export function latexToOMML(latex: string): string {
  const conversion = convertLatexToOmml(latex);
  if (conversion.source === 'unicode') {
    console.warn(
      `No native Office Math for "${latex}"`
      + `${conversion.reason ? ` (${conversion.reason})` : ''}; `
      + `falling back to ${conversion.lossy ? 'lossy ' : ''}Unicode text.`,
    );
  }
  return conversion.value;
}

/**
 * The Unicode path cannot express environments, scalable delimiters, or any
 * command it does not recognise — a surviving backslash is the tell.
 */
function isLossyConversion(latex: string, plain: string): boolean {
  if (/\\[A-Za-z]+/.test(plain)) return true;
  return /\\(?:begin|end|left|right)\b/.test(latex);
}

/**
 * Command → Unicode, looked up by *exact* command name.
 *
 * The previous implementation ran an ordered chain of `.replace(/\\prod/g, …)`
 * calls, so any command that is a prefix of another was corrupted:
 * `\propto` became `Πto`, `\cdots` became `·s`, `\simeq` became `∼eq`.
 * Matching `\\([A-Za-z]+)` and looking the whole name up removes that class of
 * bug entirely, and unknown commands survive instead of being silently mangled.
 */
const PLAIN_TEXT_SYMBOLS: Record<string, string> = {
  Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ', Xi: 'Ξ', Pi: 'Π',
  Sigma: 'Σ', Upsilon: 'Υ', Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω',
  alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', varepsilon: 'ε',
  zeta: 'ζ', eta: 'η', theta: 'θ', vartheta: 'ϑ', iota: 'ι', kappa: 'κ',
  lambda: 'λ', mu: 'μ', nu: 'ν', xi: 'ξ', pi: 'π', rho: 'ρ',
  sigma: 'σ', tau: 'τ', upsilon: 'υ', phi: 'φ', varphi: 'φ', chi: 'χ',
  psi: 'ψ', omega: 'ω',
  infty: '∞', partial: '∂', nabla: '∇',
  propto: '∝', prod: 'Π', sum: 'Σ', int: '∫',
  times: '×', cdot: '·', pm: '±', mp: '∓',
  le: '≤', ge: '≥', leq: '≤', geq: '≥', ne: '≠', neq: '≠', approx: '≈',
  in: '∈', notin: '∉', ni: '∋',
  subset: '⊂', supset: '⊃', subseteq: '⊆', supseteq: '⊇',
  equiv: '≡', sim: '∼', parallel: '∥', perp: '⊥',
  sqrt: '√',
};

/** Read a balanced `{...}` group starting at `start` (which must be `{`). */
function readBalancedGroup(text: string, start: number): { content: string; end: number } | null {
  if (text[start] !== '{') return null;
  let depth = 0;
  for (let index = start; index < text.length; index += 1) {
    const character = text[index];
    if (character === '\\') { index += 1; continue; }
    if (character === '{') depth += 1;
    else if (character === '}') {
      depth -= 1;
      if (depth === 0) return { content: text.slice(start + 1, index), end: index + 1 };
    }
  }
  return null;
}

/**
 * Replace `\name{…}` (one or more balanced groups) with `transform(contents)`.
 *
 * A regex such as `/\\frac\{([^}]+)\}\{([^}]+)\}/` cannot match nested braces,
 * so `\frac{a_{1}}{b}` was left as literal LaTeX and the subsequent brace
 * cleanup turned it into `\fraca_1/b`.
 */
function replaceCommandGroups(
  text: string,
  names: string[],
  groupCount: number,
  transform: (contents: string[]) => string,
): string {
  // Longest name first, or `text` would match before `textrm`.
  const alternation = [...names].sort((a, b) => b.length - a.length).join('|');
  const pattern = new RegExp(`\\\\(?:${alternation})`, 'g');

  let result = '';
  let cursor = 0;
  let match: RegExpExecArray | null;

  while ((match = pattern.exec(text)) !== null) {
    let position = match.index + match[0].length;
    const contents: string[] = [];
    let complete = true;

    for (let group = 0; group < groupCount; group += 1) {
      while (text[position] === ' ') position += 1;
      const balanced = readBalancedGroup(text, position);
      if (!balanced) { complete = false; break; }
      contents.push(balanced.content);
      position = balanced.end;
    }
    if (!complete) continue;

    result += text.slice(cursor, match.index) + transform(contents);
    cursor = position;
    pattern.lastIndex = position;
  }

  return result + text.slice(cursor);
}

/**
 * Convert `^{…}` / `_{…}` (and the bare `^x` / `_x` forms) to Unicode.
 * The bare form was previously left as a literal caret/underscore.
 */
function replaceScripts(text: string): string {
  let result = '';
  let index = 0;

  while (index < text.length) {
    const character = text[index];
    if (character === '^' || character === '_') {
      const map = character === '^' ? SUPERSCRIPTS : SUBSCRIPTS;
      if (text[index + 1] === '{') {
        const balanced = readBalancedGroup(text, index + 1);
        if (balanced) {
          result += charMap(balanced.content, map);
          index = balanced.end;
          continue;
        }
      }
      const next = text[index + 1];
      if (next !== undefined && /[0-9A-Za-z+\-=()]/.test(next)) {
        result += map[next] ?? next;
        index += 2;
        continue;
      }
    }
    result += character;
    index += 1;
  }

  return result;
}

/**
 * Render a LaTeX equation to plain Unicode text (fallback when pandoc and the
 * native OMML path are unavailable).
 */
export function latexToPlainText(latex: string): string {
  // `\text{…}` keeps its spaces, and the whitespace that follows such a group
  // sits in math mode where it is insignificant. Pulling the groups out lets
  // the restore step below drop exactly that one following space, instead of
  // turning `\text{if } x>0` into `if  x>0`.
  const literals: string[] = [];
  let result = replaceCommandGroups(latex, ['text', 'textrm'], 1, ([content]) => {
    literals.push(content);
    return `\u0000${literals.length - 1}\u0000`;
  });

  // Structural commands, with brace-aware matching.
  result = replaceCommandGroups(result, ['frac', 'dfrac', 'tfrac'], 2, ([numerator, denominator]) => `(${numerator})/(${denominator})`);
  result = replaceCommandGroups(result, ['sqrt'], 1, ([content]) => `√(${content})`);
  result = replaceCommandGroups(result, ['mathrm', 'mathbf', 'mathit'], 1, ([content]) => content);
  result = replaceCommandGroups(result, ['bar'], 1, ([content]) => `${content}̄`);
  result = replaceCommandGroups(result, ['hat'], 1, ([content]) => `${content}̂`);
  result = replaceCommandGroups(result, ['dot'], 1, ([content]) => `${content}̇`);
  result = replaceCommandGroups(result, ['ddot'], 1, ([content]) => `${content}̈`);
  result = replaceCommandGroups(result, ['tilde'], 1, ([content]) => `${content}̃`);
  result = replaceCommandGroups(result, ['vec'], 1, ([content]) => `${content}⃗`);

  result = replaceScripts(result);

  // Exact-name symbol lookup — no prefix collisions.
  result = result.replace(/\\([A-Za-z]+)/g, (match: string, name: string) => PLAIN_TEXT_SYMBOLS[name] ?? match);

  result = result
    .replace(/\\\|/g, '‖')
    .replace(/\\,/g, ' ')
    .replace(/\\;/g, '  ')
    .replace(/\\quad/g, '    ')
    .replace(/\\qquad/g, '        ')
    .replace(/[{}]/g, '');

  // Restore the verbatim text groups, dropping the insignificant space that
  // follows one when the literal already ends with its own space.
  result = result.replace(/\u0000(\d+)\u0000([ \t]*)/g, (_match, index: string, trailing: string) => {
    const literal = literals[Number(index)] ?? '';
    return /\s$/.test(literal) ? literal : literal + trailing;
  });

  return result.trim();
}

/**
 * Check if a string looks like LaTeX math.
 */
export function isLatexMath(text: string): boolean {
  return /\\[a-zA-Z]+|\^\{|_\{|\\frac|\\sum|\\int|\\sqrt/.test(text);
}

/**
 * Render a LaTeX equation for DOCX insertion.
 * V2: Returns OMML XML string when pandoc available, Unicode otherwise.
 */
export function renderEquation(latex: string): string {
  return latexToOMML(latex);
}

// ── Native math AST ──────────────────────────────────────

class LatexMathParser {
  private index = 0;

  constructor(private readonly source: string) {}

  parse(stopCharacter?: string): LatexMathNode[] {
    const nodes: LatexMathNode[] = [];
    while (this.index < this.source.length) {
      if (stopCharacter && this.source[this.index] === stopCharacter) break;
      const atom = this.parseAtom();
      if (atom.length === 0) continue;

      let subScript: LatexMathNode[] | undefined;
      let superScript: LatexMathNode[] | undefined;
      this.skipWhitespace();
      while (this.source[this.index] === '_' || this.source[this.index] === '^') {
        const marker = this.source[this.index];
        this.index += 1;
        const value = this.parseScriptValue();
        // `x^{a}^{b}` is malformed LaTeX (double superscript). Keep the first
        // rather than silently discarding it.
        if (marker === '_') subScript ??= value;
        else superScript ??= value;
        this.skipWhitespace();
      }

      if (subScript || superScript) {
        const only = atom.length === 1 ? atom[0] : undefined;
        if (only?.type === 'sum' || only?.type === 'integral' || only?.type === 'nary' || only?.type === 'limit') {
          appendMathNode(nodes, { ...only, subScript, superScript });
        } else {
          appendMathNode(nodes, { type: 'script', base: atom, subScript, superScript });
        }
      } else {
        for (const node of atom) appendMathNode(nodes, node);
      }
    }
    return nodes;
  }

  private parseAtom(): LatexMathNode[] {
    const character = this.source[this.index];
    if (character === undefined) return [];
    if (/\s/.test(character)) {
      this.skipWhitespace();
      return [{ type: 'run', text: ' ' }];
    }
    if (character === '{') {
      this.index += 1;
      const children = this.parse('}');
      if (this.source[this.index] === '}') this.index += 1;
      return children;
    }
    if (character === '(' || character === '[') {
      this.index += 1;
      const closing = character === '(' ? ')' : ']';
      const children = this.parse(closing);
      if (this.source[this.index] === closing) this.index += 1;
      return [{ type: 'delimiter', opening: character, closing, children }];
    }
    if (character !== '\\') {
      this.index += 1;
      return [{ type: 'run', text: character }];
    }

    this.index += 1;
    const command = this.readCommand();
    if (!command) return [];
    if (command === 'frac' || command === 'dfrac' || command === 'tfrac') {
      return [{
        type: 'fraction',
        numerator: this.parseRequiredGroup(),
        denominator: this.parseRequiredGroup(),
      }];
    }
    if (command === 'sqrt') {
      const degree = this.parseOptionalGroup('[', ']');
      return [{ type: 'radical', children: this.parseRequiredGroup(), degree }];
    }
    if (command === 'sum') return [{ type: 'sum' }];
    if (command === 'int') return [{ type: 'integral' }];
    if (command === 'iint' || command === 'iiint') {
      return [{ type: 'nary', operator: command === 'iint' ? '∬' : '∭' }];
    }
    if (command === 'prod' || command === 'coprod' || command === 'bigcap' || command === 'bigcup') {
      const operators: Record<string, string> = { prod: '∏', coprod: '∐', bigcap: '⋂', bigcup: '⋃' };
      return [{ type: 'nary', operator: operators[command] }];
    }
    if (LIMIT_COMMANDS.has(command)) return [{ type: 'limit', name: command }];
    if (FUNCTION_COMMANDS.has(command)) {
      return [{ type: 'function', name: command, children: this.parseFunctionArgument() }];
    }
    if (command === 'operatorname') {
      // `\operatorname*{argmax}` is the same command with a star; without this
      // the star was read as the group and the name came back empty.
      if (this.source[this.index] === '*') this.index += 1;
      // Spacing macros inside the name (`\operatorname{arg\,max}`) are layout,
      // not part of the operator's text.
      const name = this.readRequiredGroupText().replace(/\\[,;:!\s]+/g, '');
      return [{ type: 'function', name, children: this.parseFunctionArgument() }];
    }
    if (command === 'binom') {
      const fraction: LatexMathNode = {
        type: 'fraction', numerator: this.parseRequiredGroup(), denominator: this.parseRequiredGroup(), bar: false,
      };
      return [{ type: 'delimiter', opening: '(', closing: ')', children: [fraction] }];
    }
    if (command === 'overset' || command === 'underset') {
      const annotation = this.parseRequiredGroup();
      const base = this.parseRequiredGroup();
      return [{
        type: 'overUnder', base,
        over: command === 'overset' ? annotation : undefined,
        under: command === 'underset' ? annotation : undefined,
      }];
    }
    if (command === 'xrightarrow' || command === 'xleftarrow') {
      const under = this.parseOptionalGroup('[', ']');
      const over = this.parseRequiredGroup();
      const arrow = command === 'xrightarrow' ? '→' : '←';
      return [{
        type: 'overUnder',
        base: [{ type: 'run', text: arrow }],
        over,
        under,
      }];
    }
    if (command in ACCENT_COMMANDS) {
      return [{ type: 'accent', accent: ACCENT_COMMANDS[command], children: this.parseRequiredGroup() }];
    }
    if (command === 'overline' || command === 'underline') {
      return [{ type: 'bar', position: command === 'overline' ? 'top' : 'bottom', children: this.parseRequiredGroup() }];
    }
    if (command === 'left') return [this.parseLeftRight()];
    if (command === 'right') {
      this.readDelimiterToken();
      return [];
    }
    if (command === 'begin') {
      const environment = this.readRequiredGroupText();
      return [this.parseEnvironment(environment)];
    }
    if (command === 'text' || command === 'textrm') {
      // Verbatim text: spaces are content here. Running it through the math
      // parser dropped them, so `\text{if } x>0` rendered as "ifx>0".
      return [{ type: 'run', text: this.readRequiredGroupText(false) }];
    }
    if (command === 'mathrm' || command === 'mathbf' || command === 'mathit' || command === 'boldsymbol' || command === 'bm') {
      return this.parseRequiredGroup();
    }
    if (command === ',' || command === ':' || command === ';' || command === 'quad') {
      return [{ type: 'run', text: command === 'quad' ? '    ' : ' ' }];
    }
    if (command === 'qquad') return [{ type: 'run', text: '        ' }];
    if (command === '!' || command === ' ') return [];

    const symbol = MATH_SYMBOLS[command];
    if (symbol !== undefined) return [{ type: 'run', text: symbol }];

    // A single non-letter command is an escaped literal (`\%`, `\{`, `\\`).
    if (!/^[A-Za-z]+$/.test(command)) return [{ type: 'run', text: command }];

    // An unknown command keeps its backslash, so it renders as visibly
    // unsupported instead of silently becoming a bare identifier.
    return [{ type: 'run', text: `\\${command}` }];
  }

  private parseRequiredGroup(): LatexMathNode[] {
    this.skipWhitespace();
    if (this.source[this.index] !== '{') return this.parseAtom();
    this.index += 1;
    const children = this.parse('}');
    if (this.source[this.index] === '}') this.index += 1;
    return children;
  }

  private parseOptionalGroup(open: string, close: string): LatexMathNode[] | undefined {
    this.skipWhitespace();
    if (this.source[this.index] !== open) return undefined;
    this.index += 1;
    const children = this.parse(close);
    if (this.source[this.index] === close) this.index += 1;
    return children;
  }

  private parseScriptValue(): LatexMathNode[] {
    this.skipWhitespace();
    return this.parseRequiredGroup();
  }

  private parseFunctionArgument(): LatexMathNode[] {
    this.skipWhitespace();
    if (this.source[this.index] === '^' || this.source[this.index] === '_' || this.index >= this.source.length) return [];
    const base = this.parseAtom();
    let subScript: LatexMathNode[] | undefined;
    let superScript: LatexMathNode[] | undefined;
    while (this.source[this.index] === '_' || this.source[this.index] === '^') {
      const marker = this.source[this.index];
      this.index += 1;
      const value = this.parseScriptValue();
      if (marker === '_') subScript = value;
      else superScript = value;
    }
    return subScript || superScript
      ? [{ type: 'script', base, subScript, superScript }]
      : base;
  }

  /**
   * Read a `{...}` group as raw text.
   * `trim` is off for `\text{…}`, where leading/trailing spaces are content.
   */
  private readRequiredGroupText(trim = true): string {
    this.skipWhitespace();
    if (this.source[this.index] !== '{') return '';
    this.index += 1;
    const start = this.index;
    let depth = 1;
    while (this.index < this.source.length && depth > 0) {
      if (this.source[this.index] === '{') depth += 1;
      else if (this.source[this.index] === '}') depth -= 1;
      this.index += 1;
    }
    const text = this.source.slice(start, depth === 0 ? this.index - 1 : this.index);
    return trim ? text.trim() : text;
  }

  private parseLeftRight(): LatexMathNode {
    const opening = this.readDelimiterToken();
    const start = this.index;
    // `(?![A-Za-z])` keeps `\leftarrow` / `\rightarrow` from being mistaken for
    // `\left` / `\right`: the bare prefix match consumed the real closer and
    // made the delimiter swallow the rest of the equation.
    const marker = /\\(left|right)(?![A-Za-z])\s*/g;
    marker.lastIndex = start;
    let depth = 0;
    let bodyEnd = this.source.length;
    let closing = '';
    let match: RegExpExecArray | null;
    while ((match = marker.exec(this.source))) {
      if (match[1] === 'left') {
        depth += 1;
        continue;
      }
      if (depth > 0) {
        depth -= 1;
        continue;
      }
      bodyEnd = match.index;
      this.index = marker.lastIndex;
      closing = this.readDelimiterToken();
      break;
    }
    if (bodyEnd === this.source.length) this.index = this.source.length;
    const children = new LatexMathParser(this.source.slice(start, bodyEnd)).parse();
    return { type: 'delimiter', opening, closing, children };
  }

  private readDelimiterToken(): string {
    this.skipWhitespace();
    if (this.source[this.index] !== '\\') {
      const token = this.source[this.index] ?? '';
      this.index += token ? 1 : 0;
      return token === '.' ? '' : token;
    }
    this.index += 1;
    const token = this.readCommand();
    return DELIMITER_COMMANDS[token] ?? (token === '.' ? '' : token);
  }

  private parseEnvironment(name: string): LatexMathNode {
    const body = this.readEnvironmentBody(name);
    if (!MATRIX_ENVIRONMENTS.has(name as MatrixEnvironment)) {
      return { type: 'delimiter', opening: '', closing: '', children: new LatexMathParser(body).parse() };
    }
    const environment = name as MatrixEnvironment;
    const delimiters = MATRIX_DELIMITERS[environment];
    const rows = splitEnvironmentRows(body).map(row => row.map(cell => new LatexMathParser(cell).parse()));
    return { type: 'matrix', rows, opening: delimiters[0], closing: delimiters[1], environment };
  }

  private readEnvironmentBody(name: string): string {
    const start = this.index;
    const marker = /\\(begin|end)\s*\{([^{}]+)\}/g;
    marker.lastIndex = start;
    let depth = 1;
    let match: RegExpExecArray | null;
    while ((match = marker.exec(this.source))) {
      if (match[1] === 'begin') depth += 1;
      else depth -= 1;
      if (depth === 0) {
        const body = this.source.slice(start, match.index);
        // Only consume the `\end` when it closes the environment we opened.
        // A mismatched name used to keep scanning to EOF and swallow the whole
        // remainder of the equation into this node.
        this.index = match[2].trim() === name ? marker.lastIndex : match.index;
        return body;
      }
    }
    this.index = this.source.length;
    return this.source.slice(start);
  }

  private readCommand(): string {
    const start = this.index;
    while (/[A-Za-z]/.test(this.source[this.index] ?? '')) this.index += 1;
    if (this.index > start) return this.source.slice(start, this.index);
    const command = this.source[this.index] ?? '';
    this.index += command ? 1 : 0;
    return command;
  }

  private skipWhitespace(): void {
    while (/\s/.test(this.source[this.index] ?? '')) this.index += 1;
  }
}

function appendMathNode(nodes: LatexMathNode[], node: LatexMathNode): void {
  const previous = nodes[nodes.length - 1];
  if (previous?.type === 'run' && node.type === 'run') previous.text += node.text;
  else nodes.push(node);
}

const MATH_SYMBOLS: Record<string, string> = {
  Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ', Xi: 'Ξ', Pi: 'Π',
  Sigma: 'Σ', Upsilon: 'Υ', Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω',
  alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', varepsilon: 'ε',
  zeta: 'ζ', eta: 'η', theta: 'θ', vartheta: 'ϑ', iota: 'ι', kappa: 'κ',
  lambda: 'λ', mu: 'μ', nu: 'ν', xi: 'ξ', pi: 'π', rho: 'ρ', sigma: 'σ',
  tau: 'τ', upsilon: 'υ', phi: 'φ', varphi: 'φ', chi: 'χ', psi: 'ψ', omega: 'ω',
  infinity: '∞', infty: '∞', partial: '∂', nabla: '∇', prod: '∏',
  propto: '∝', times: '×', cdot: '·', pm: '±', mp: '∓', div: '÷',
  ast: '∗', star: '⋆', circ: '∘', bullet: '•',
  le: '≤', ge: '≥', leq: '≤', geq: '≥', ne: '≠', neq: '≠', approx: '≈',
  in: '∈', notin: '∉', ni: '∋',
  subset: '⊂', supset: '⊃', subseteq: '⊆', supseteq: '⊇',
  cap: '∩', cup: '∪', forall: '∀', exists: '∃', neg: '¬',
  ll: '≪', gg: '≫',
  equiv: '≡', sim: '∼', parallel: '∥', perp: '⊥',
  to: '→', rightarrow: '→', leftarrow: '←', Rightarrow: '⇒', Leftarrow: '⇐',
  leftrightarrow: '↔', Leftrightarrow: '⇔',
  ldots: '…', cdots: '⋯', ell: 'ℓ', hbar: 'ℏ', Re: 'ℜ', Im: 'ℑ',
};

const FUNCTION_COMMANDS = new Set([
  'sin', 'cos', 'tan', 'cot', 'sec', 'csc',
  'sinh', 'cosh', 'tanh', 'log', 'ln', 'exp',
  'det', 'gcd', 'ker', 'arg', 'deg', 'dim', 'hom',
  'sgn', 'diag', 'rank', 'tr',
]);

const LIMIT_COMMANDS = new Set(['lim', 'min', 'max', 'inf', 'sup']);

const ACCENT_COMMANDS: Record<string, string> = {
  hat: '̂', widehat: '̂', bar: '̄', vec: '⃗',
  dot: '̇', ddot: '̈', tilde: '̃', widetilde: '̃',
};

const DELIMITER_COMMANDS: Record<string, string> = {
  lbrace: '{', rbrace: '}', langle: '⟨', rangle: '⟩',
  vert: '|', Vert: '‖', mid: '|', '{': '{', '}': '}', '|': '|',
};

const MATRIX_ENVIRONMENTS = new Set<MatrixEnvironment>([
  'matrix', 'pmatrix', 'bmatrix', 'Bmatrix', 'vmatrix', 'Vmatrix', 'cases', 'aligned',
]);

const MATRIX_DELIMITERS: Record<MatrixEnvironment, readonly [string, string]> = {
  matrix: ['', ''], pmatrix: ['(', ')'], bmatrix: ['[', ']'], Bmatrix: ['{', '}'],
  vmatrix: ['|', '|'], Vmatrix: ['‖', '‖'], cases: ['{', ''], aligned: ['', ''],
};

/**
 * True when the character at `index` is escaped by a preceding backslash.
 *
 * Only an *odd* run of backslashes escapes: `\&` is a literal ampersand, while
 * the `&` in `\\&` follows a row break and is a real cell separator. Testing
 * just the previous character read the second case as an escape and merged two
 * cells into one.
 */
function isEscaped(source: string, index: number): boolean {
  let backslashes = 0;
  for (let scan = index - 1; scan >= 0 && source[scan] === '\\'; scan -= 1) backslashes += 1;
  return backslashes % 2 === 1;
}

function splitEnvironmentRows(source: string): string[][] {
  const rows: string[][] = [];
  let cells: string[] = [];
  let current = '';
  let braceDepth = 0;
  let environmentDepth = 0;

  const pushCell = () => {
    cells.push(current.trim());
    current = '';
  };
  const pushRow = () => {
    pushCell();
    if (cells.some(cell => cell.length > 0)) rows.push(cells);
    cells = [];
  };

  for (let index = 0; index < source.length; index += 1) {
    if (source.startsWith('\\begin{', index)) environmentDepth += 1;
    if (source.startsWith('\\end{', index) && environmentDepth > 0) environmentDepth -= 1;
    const character = source[index];
    if (character === '{') braceDepth += 1;
    else if (character === '}' && braceDepth > 0) braceDepth -= 1;

    if (braceDepth === 0 && environmentDepth === 0 && character === '&' && !isEscaped(source, index)) {
      pushCell();
      continue;
    }
    if (braceDepth === 0 && environmentDepth === 0 && character === '\\' && source[index + 1] === '\\') {
      pushRow();
      index += 1;
      if (source[index + 1] === '[') {
        const spacingEnd = source.indexOf(']', index + 2);
        if (spacingEnd >= 0) index = spacingEnd;
      }
      continue;
    }
    current += character;
  }
  pushRow();
  return rows.length > 0 ? rows : [[source.trim()]];
}

// ── OMML generation via pandoc ────────────────────────────

/** Result of the pandoc attempt: either OMML, or why it could not be produced. */
interface PandocResult {
  omml?: string;
  error?: string;
}

function generateOMMLviaPandoc(latex: string): PandocResult {
  let mdPath: string | undefined;
  let docxPath: string | undefined;
  try {
    // Create a minimal Markdown file with the equation
    const tmpDir = join(tmpdir(), 'openthesis-omml');
    if (!existsSync(tmpDir)) mkdirSync(tmpDir, { recursive: true });

    const id = randomUUID().slice(0, 8);
    mdPath = join(tmpDir, `eq-${id}.md`);
    docxPath = join(tmpDir, `eq-${id}.docx`);

    // Wrap in display math for pandoc
    const safeLatex = latex.replace(/`/g, '\\`');
    writeFileSync(mdPath, `$$${safeLatex}$$`, 'utf-8');

    // Run pandoc
    execFileSync('pandoc', [mdPath, '-o', docxPath], {
      encoding: 'utf-8',
      stdio: 'pipe',
      timeout: 5000,
    });

    // Extract OMML from the generated DOCX
    const omml = extractOMML(docxPath);
    return omml ? { omml } : { error: 'pandoc produced no <m:oMath> element' };
  } catch (error: unknown) {
    // Report why: "pandoc is not installed" and "pandoc failed on this input"
    // need different responses from the caller.
    const message = error instanceof Error ? error.message : String(error);
    return { error: message.split('\n')[0] };
  } finally {
    for (const path of [mdPath, docxPath]) {
      if (path && existsSync(path)) {
        try { unlinkSync(path); } catch {}
      }
    }
  }
}

function extractOMML(docxPath: string): string | null {
  try {
    // DOCX is a ZIP file
    const zip = new AdmZip(docxPath);
    const docXml = zip.readAsText('word/document.xml');

    // Extract the first oMath element
    const match = docXml.match(/<m:oMath[\s\S]*?<\/m:oMath>/);
    return match ? match[0] : null;
  } catch {
    return null;
  }
}

// ── Unicode helpers ───────────────────────────────────────

const SUPERSCRIPTS: Record<string, string> = {
  '0':'⁰','1':'¹','2':'²','3':'³','4':'⁴','5':'⁵','6':'⁶','7':'⁷','8':'⁸','9':'⁹',
  '+':'⁺','-':'⁻','=':'⁼','(':'⁽',')':'⁾',
  'a':'ᵃ','b':'ᵇ','c':'ᶜ','d':'ᵈ','e':'ᵉ','f':'ᶠ','g':'ᵍ','h':'ʰ',
  'i':'ⁱ','j':'ʲ','k':'ᵏ','l':'ˡ','m':'ᵐ','n':'ⁿ','o':'ᵒ','p':'ᵖ',
  'r':'ʳ','s':'ˢ','t':'ᵗ','u':'ᵘ','v':'ᵛ','w':'ʷ','x':'ˣ','y':'ʸ','z':'ᶻ',
  'A':'ᴬ','B':'ᴮ','D':'ᴰ','E':'ᴱ','G':'ᴳ','H':'ᴴ','I':'ᴵ','J':'ᴶ',
  'K':'ᴷ','L':'ᴸ','M':'ᴹ','N':'ᴺ','O':'ᴼ','P':'ᴾ','R':'ᴿ','T':'ᵀ',
  'U':'ᵁ','V':'ⱽ','W':'ᵂ',
  'α':'ᵅ','β':'ᵝ','γ':'ᵞ','δ':'ᵟ','ε':'ᵋ','θ':'ᶿ',
  'φ':'ᵠ','χ':'ᵡ',
};

const SUBSCRIPTS: Record<string, string> = {
  '0':'₀','1':'₁','2':'₂','3':'₃','4':'₄','5':'₅','6':'₆','7':'₇','8':'₈','9':'₉',
  '+':'₊','-':'₋','=':'₌','(':'₍',')':'₎',
  'a':'ₐ','e':'ₑ','h':'ₕ','i':'ᵢ','j':'ⱼ','k':'ₖ','l':'ₗ',
  'm':'ₘ','n':'ₙ','o':'ₒ','p':'ₚ','r':'ᵣ','s':'ₛ','t':'ₜ',
  'u':'ᵤ','v':'ᵥ','x':'ₓ',
  // Greek subscripts (limited Unicode support)
  'β':'ᵦ','γ':'ᵧ','ρ':'ᵨ','φ':'ᵩ','χ':'ᵪ',
};

function charMap(text: string, map: Record<string, string>): string {
  return [...text].map(c => map[c] || c).join('');
}
