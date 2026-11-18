// ============================================================
// Statistical symbol formatting — regression lock
// ============================================================
// The convention under test (see statistical-symbols.ts):
// statistics and population parameters are set in Times New
// Roman italic, never bold; vectors/matrices bold italic.
// Every case here is one a reviewer has flagged or plausibly
// would: if a future change breaks one, this file fails first.
// ============================================================

import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { createUSTBTemplate, renderDocument, splitStatisticalSymbols } from '../packages/docx-renderer/dist/index.js';

const kinds = text => splitStatisticalSymbols(text).map(segment => [segment.text, segment.kind]);

test('recognises the core statistics in their usual contexts', () => {
  assert.deepEqual(kinds('P < 0.001'), [['P', 'statistic'], [' < 0.001', 'plain']]);
  assert.deepEqual(kinds('P 值'), [['P', 'statistic'], [' 值', 'plain']]);
  assert.deepEqual(kinds('P值'), [['P', 'statistic'], ['值', 'plain']]);
  assert.deepEqual(kinds('t(12) = 2.31'), [['t', 'statistic'], ['(12) = 2.31', 'plain']]);
  assert.deepEqual(kinds('F(1, 23) = 4.10'), [['F', 'statistic'], ['(1, 23) = 4.10', 'plain']]);
  assert.deepEqual(kinds('M = 3.20'), [['M', 'statistic'], [' = 3.20', 'plain']]);
  assert.deepEqual(kinds('r = -0.45'), [['r', 'statistic'], [' = -0.45', 'plain']]);
  assert.deepEqual(kinds('z = 1.96'), [['z', 'statistic'], [' = 1.96', 'plain']]);
  assert.deepEqual(kinds('d = 0.50'), [['d', 'statistic'], [' = 0.50', 'plain']]);
  assert.deepEqual(kinds('g = 0.30'), [['g', 'statistic'], [' = 0.30', 'plain']]);
  assert.deepEqual(kinds('H = 6.20'), [['H', 'statistic'], [' = 6.20', 'plain']]);
  assert.deepEqual(kinds('n = 30'), [['n', 'statistic'], [' = 30', 'plain']]);
  assert.deepEqual(kinds('N = 120'), [['N', 'statistic'], [' = 120', 'plain']]);
  assert.deepEqual(kinds('U = 45'), [['U', 'statistic'], [' = 45', 'plain']]);
  assert.deepEqual(kinds('W = 78'), [['W', 'statistic'], [' = 78', 'plain']]);
  assert.deepEqual(kinds('t检验'), [['t', 'statistic'], ['检验', 'plain']]);
  assert.deepEqual(kinds('F分布'), [['F', 'statistic'], ['分布', 'plain']]);
  assert.deepEqual(kinds('P value'), [['P', 'statistic'], [' value', 'plain']]);
});

test('recognises multi-letter statistics and Greek parameters', () => {
  assert.deepEqual(kinds('SD = 1.10'), [['SD', 'statistic'], [' = 1.10', 'plain']]);
  assert.deepEqual(kinds('SE为0.21'), [['SE', 'statistic'], ['为0.21', 'plain']]);
  assert.deepEqual(kinds('Mdn = 4'), [['Mdn', 'statistic'], [' = 4', 'plain']]);
  assert.deepEqual(kinds('df = 12'), [['df', 'statistic'], [' = 12', 'plain']]);
  assert.deepEqual(kinds('95% CI'), [['95% ', 'plain'], ['CI', 'statistic']]);
  assert.deepEqual(kinds('OR = 1.50'), [['OR', 'statistic'], [' = 1.50', 'plain']]);
  assert.deepEqual(kinds('μ'), [['μ', 'statistic']]);
  assert.deepEqual(kinds('σ = 0.5'), [['σ', 'statistic'], [' = 0.5', 'plain']]);
  assert.deepEqual(kinds('α系数'), [['α', 'statistic'], ['系数', 'plain']]);
  assert.deepEqual(kinds('\\mu与\\sigma'), [['μ', 'statistic'], ['与', 'plain'], ['σ', 'statistic']]);
});

test('normalises squared statistics typed without a superscript', () => {
  assert.deepEqual(kinds('χ²检验'), [['χ²', 'statistic'], ['检验', 'plain']]);
  assert.deepEqual(kinds('χ2检验'), [['χ²', 'statistic'], ['检验', 'plain']]);
  assert.deepEqual(kinds('\\chi^2 = 5.99'), [['χ²', 'statistic'], [' = 5.99', 'plain']]);
  assert.deepEqual(kinds('R² = 0.85'), [['R²', 'statistic'], [' = 0.85', 'plain']]);
  assert.deepEqual(kinds('R^2 = 0.85'), [['R²', 'statistic'], [' = 0.85', 'plain']]);
  assert.deepEqual(kinds('R2为0.85'), [['R²', 'statistic'], ['为0.85', 'plain']]);
  assert.deepEqual(kinds('R = 0.90'), [['R', 'statistic'], [' = 0.90', 'plain']]);
});

test('accepts the fullwidth forms Chinese manuscripts mix in', () => {
  assert.deepEqual(kinds('P＜0.001'), [['P', 'statistic'], ['＜0.001', 'plain']]);
  assert.deepEqual(kinds('t（12）＝2.31'), [['t', 'statistic'], ['（12）＝2.31', 'plain']]);
  assert.deepEqual(kinds('结果显示P为0.997'), [['结果显示', 'plain'], ['P', 'statistic'], ['为0.997', 'plain']]);
  assert.deepEqual(kinds('样本量n为30'), [['样本量', 'plain'], ['n', 'statistic'], ['为30', 'plain']]);
});

test('recognises vectors and matrices as bold italic material', () => {
  assert.deepEqual(kinds('矩阵A'), [['矩阵', 'plain'], ['A', 'vector']]);
  assert.deepEqual(kinds('向量 x'), [['向量 ', 'plain'], ['x', 'vector']]);
  assert.deepEqual(kinds('\\mathbf{B}'), [['B', 'vector']]);
  assert.deepEqual(kinds('\\vec{x}'), [['x', 'vector']]);
});

test('leaves ordinary prose and lookalike compounds untouched', () => {
  for (const text of [
    'The plan was practical and the approach worked well.',
    'M码与SD卡',
    'CI/CD 流水线',
    'R2D2 是机器人',
    'A组与B组',
    'Plan B was approved.',
    'e.g. the second one',
    'Vitamin D levels',
  ]) {
    assert.deepEqual(kinds(text), [[text, 'plain']], `expected no symbol split in: ${text}`);
  }
});

test('renders fullwidth and squared forms correctly in a document', async () => {
  const template = createUSTBTemplate();
  const buffer = await renderDocument({
    outputPath: '', template,
    document: {
      type: 'thesis', meta: { title: 'Hardening' }, cover: [],
      sections: [{
        id: 's', type: 'chapter', title: '结果',
        content: [
          { type: 'paragraph', text: '差异显著（P＜0.001），模型拟合良好，R2为0.85，χ2检验通过，95% CI为[0.12, 0.34]。' },
        ],
      }],
    },
  });
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file('word/document.xml').async('string');
  const runs = [...xml.matchAll(/<w:r>([\s\S]*?)<\/w:r>/g)].map(match => ({
    text: [...match[1].matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(part => part[1]).join(''),
    italic: /<w:i\/>|<w:i\s/.test(match[1]),
    boldOff: /<w:b w:val="false"\/>/.test(match[1]),
  }));
  for (const symbol of ['P', 'R²', 'χ²', 'CI']) {
    const run = runs.find(candidate => candidate.text === symbol);
    assert.ok(run, `expected a dedicated run for ${symbol}`);
    assert.equal(run.italic, true, `${symbol} should be italic`);
    assert.equal(run.boldOff, true, `${symbol} should not be bold`);
  }
  assert.doesNotMatch(xml, />R2</);
  assert.doesNotMatch(xml, />χ2</);
});
