import test from 'node:test';
import assert from 'node:assert/strict';
import { isLatexMath, latexToMathAst, latexToPlainText, convertLatexToOmml } from '../packages/equation-engine/dist/index.js';

test('converts common LaTeX symbols and scripts to Unicode', () => {
  assert.equal(latexToPlainText('E = mc^{2}'), 'E = mc²');
  assert.equal(latexToPlainText('\\alpha + \\beta \\leq \\infty'), 'α + β ≤ ∞');
  assert.equal(latexToPlainText('x_{max} = \\frac{a}{b}'), 'xₘₐₓ = (a)/(b)');
});

test('detects LaTeX while ignoring ordinary prose', () => {
  assert.equal(isLatexMath('\\frac{a}{b}'), true);
  assert.equal(isLatexMath('x^{2}'), true);
  assert.equal(isLatexMath('ordinary text'), false);
});

test('builds a native-math AST for fractions, radicals, scripts, sums, and integrals', () => {
  const nodes = latexToMathAst('\\frac{a_1}{\\sqrt{x^2}} + \\sum_{i=1}^{n} x_i + \\int_0^1 f(x)');
  assert.equal(nodes[0].type, 'fraction');
  assert.equal(nodes[0].numerator[0].type, 'script');
  assert.equal(nodes[0].denominator[0].type, 'radical');
  assert.ok(nodes.some(node => node.type === 'sum' && node.subScript && node.superScript));
  assert.ok(nodes.some(node => node.type === 'integral' && node.subScript && node.superScript));
});

test('parses advanced macros into native-math nodes', () => {
  const nodes = latexToMathAst(
    '\\hat{x}+\\overline{AB}+\\lim_{x\\to0}\\frac{\\sin(x)}{x}+\\binom{n}{k}+\\overset{*}{=}',
  );
  assert.ok(nodes.some(node => node.type === 'accent' && node.accent === '̂'));
  assert.ok(nodes.some(node => node.type === 'bar' && node.position === 'top'));
  assert.ok(nodes.some(node => node.type === 'limit' && node.name === 'lim' && node.subScript));
  const fraction = nodes.find(node => node.type === 'fraction');
  assert.equal(fraction?.numerator[0].type, 'function');
  const binomial = nodes.find(node => node.type === 'delimiter' && node.opening === '(');
  assert.equal(binomial?.children[0].type, 'fraction');
  assert.equal(binomial?.children[0].bar, false);
  assert.ok(nodes.some(node => node.type === 'overUnder' && node.over));

  const nary = latexToMathAst('\\prod_{i=1}^{n}x_i')[0];
  assert.equal(nary.type, 'nary');
  assert.equal(nary.operator, '∏');
  assert.ok(nary.subScript && nary.superScript);
});

test('parses matrix, cases, and aligned environments with cells and rows', () => {
  const matrix = latexToMathAst('\\begin{pmatrix}a & \\frac{b}{c} \\\\ d & e\\end{pmatrix}')[0];
  assert.equal(matrix.type, 'matrix');
  assert.equal(matrix.environment, 'pmatrix');
  assert.deepEqual([matrix.opening, matrix.closing], ['(', ')']);
  assert.equal(matrix.rows.length, 2);
  assert.equal(matrix.rows[0].length, 2);
  assert.equal(matrix.rows[0][1][0].type, 'fraction');

  const cases = latexToMathAst('\\begin{cases}x^2 & x\\geq0 \\\\ -x & x<0\\end{cases}')[0];
  assert.equal(cases.type, 'matrix');
  assert.equal(cases.environment, 'cases');
  assert.deepEqual([cases.opening, cases.closing], ['{', '']);

  const aligned = latexToMathAst('\\begin{aligned}a&=b+c \\\\ d&=e\\end{aligned}')[0];
  assert.equal(aligned.type, 'matrix');
  assert.equal(aligned.rows.length, 2);
});

test('keeps a cell separator that follows a row break', () => {
  // `\\&` is a row break followed by a real separator. Testing only the
  // previous character read the `&` as escaped and merged two cells into one.
  const matrix = latexToMathAst('\\begin{matrix}1&2\\\\&3\\end{matrix}')[0];
  assert.equal(matrix.rows.length, 2);
  assert.equal(matrix.rows[1].length, 2);
  assert.deepEqual(matrix.rows[1][0], []);
  assert.equal(matrix.rows[1][1][0].text, '3');

  // A genuinely escaped ampersand still belongs to its own cell.
  const escaped = latexToMathAst('\\begin{matrix}a\\&b&c\\end{matrix}')[0];
  assert.equal(escaped.rows[0].length, 2);
  assert.equal(escaped.rows[0][0].map(node => node.text ?? '').join(''), 'a&b');
});

test('preserves scalable left/right delimiters in the AST', () => {
  const node = latexToMathAst('\\left\\langle \\frac{x}{y} \\right\\rangle')[0];
  assert.equal(node.type, 'delimiter');
  assert.deepEqual([node.opening, node.closing], ['⟨', '⟩']);
  assert.ok(node.children.some(child => child.type === 'fraction'));
});

test('keeps scripts on function arguments distinct from scripts on function names', () => {
  const argumentScript = latexToMathAst('\\sin x^2')[0];
  assert.equal(argumentScript.type, 'function');
  assert.equal(argumentScript.children[0].type, 'script');

  const functionScript = latexToMathAst('\\sin^2 x')[0];
  assert.equal(functionScript.type, 'script');
  assert.equal(functionScript.base[0].type, 'function');
  assert.deepEqual(functionScript.superScript, [{ type: 'run', text: '2' }]);
});

test('does not confuse a command with a longer command that starts the same way', () => {
  // An ordered replace chain turned `\propto` into `Πto` and `\cdots` into `·s`.
  assert.equal(latexToPlainText('\\propto'), '∝');
  assert.equal(latexToPlainText('\\prod'), 'Π');
  // Unknown commands survive intact rather than being silently mangled.
  assert.equal(latexToPlainText('\\cdots'), '\\cdots');
  assert.equal(latexToPlainText('\\simeq'), '\\simeq');
});

test('converts bare superscripts and subscripts to Unicode', () => {
  assert.equal(latexToPlainText('x^2'), 'x²');
  assert.equal(latexToPlainText('x_1'), 'x₁');
  assert.equal(latexToPlainText('x^{2}'), 'x²');
});

test('handles nested braces in structural commands', () => {
  // `[^}]+` cannot match nested braces, so this used to leak literal LaTeX.
  assert.equal(latexToPlainText('\\frac{a_{1}}{b}'), '(a₁)/(b)');
  assert.equal(latexToPlainText('\\sqrt{x_{1}}'), '√(x₁)');
});

test('keeps the spaces inside \\text and drops the insignificant one after it', () => {
  const nodes = latexToMathAst('\\text{if } x>0');
  // Adjacent runs merge, so the equation lands in a single run — the point is
  // that the space survives instead of collapsing to "ifx>0".
  assert.equal(nodes.map(node => node.text ?? '').join(''), 'if x>0');
  assert.equal(latexToPlainText('\\text{if } x>0'), 'if x>0');
  // A text group that does not end in a space still keeps the one after it.
  assert.equal(latexToPlainText('\\text{if} x>0'), 'if x>0');
});

test('does not mistake \\leftarrow for \\left', () => {
  const node = latexToMathAst('\\left( a \\leftarrow b \\right)')[0];
  assert.equal(node.type, 'delimiter');
  assert.equal(node.closing, ')');
  assert.match(JSON.stringify(node.children), /←/);
});

test('parses \\operatorname with an optional star and spacing macros', () => {
  // The star used to be read as the group, leaving the name empty.
  // `_{x}` then attaches to the operator, so the function sits inside a script.
  const starred = latexToMathAst('\\operatorname*{argmax}_{x}')[0];
  assert.equal(starred.type, 'script');
  assert.equal(starred.base[0].type, 'function');
  assert.equal(starred.base[0].name, 'argmax');

  const spaced = latexToMathAst('\\operatorname{arg\\,max}')[0];
  assert.equal(spaced.name, 'argmax');
});

test('keeps the backslash on an unknown command', () => {
  // Silently dropping it turned `\foo` into an identifier that looked intended.
  assert.deepEqual(latexToMathAst('\\foo')[0], { type: 'run', text: '\\foo' });
  // Escaped literals are still emitted bare.
  assert.deepEqual(latexToMathAst('\\%')[0], { type: 'run', text: '%' });
});

test('does not swallow the rest of the equation when an environment name mismatches', () => {
  const nodes = latexToMathAst('\\begin{aligned}a&=b\\end{matrix}+c');
  const text = JSON.stringify(nodes);
  // The old scanner ran to EOF on a mismatched name, so `+c` vanished.
  assert.match(text, /\+/);
  assert.match(text, /c/);
});

test('keeps the first superscript of a malformed double superscript', () => {
  const node = latexToMathAst('x^{a}^{b}')[0];
  assert.equal(node.type, 'script');
  assert.deepEqual(node.superScript, [{ type: 'run', text: 'a' }]);
});

test('reports which conversion path was used, and why the fallback was taken', () => {
  // pandoc may or may not be installed, so assert the contract rather than the
  // path: the conversion must always say where it came from.
  const plain = convertLatexToOmml('\\alpha + \\beta');
  assert.ok(['pandoc', 'unicode'].includes(plain.source));

  if (plain.source === 'pandoc') {
    assert.match(plain.value, /<m:oMath/);
    assert.equal(plain.lossy, false);
  } else {
    assert.equal(plain.value, 'α + β');
    assert.equal(plain.lossy, false);
    assert.ok(plain.reason, 'the fallback must say why pandoc was unavailable');
  }

  // Constructs the Unicode path cannot express are flagged as lossy.
  const matrix = convertLatexToOmml('\\begin{pmatrix}a&b\\\\c&d\\end{pmatrix}');
  if (matrix.source === 'unicode') {
    assert.equal(matrix.lossy, true);
    assert.ok(matrix.reason);
  }
});
