import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createUSTBTemplate, renderDocument } from '../packages/docx-renderer/dist/index.js';
import { parseMarkdownBlocks } from '../packages/markdown-parser/dist/index.js';

const template = createUSTBTemplate();

// A 2x2 JPEG, the same bytes the other renderer tests embed.
const JPEG = Buffer.from(
  '/9j/4AAQSkZJRgABAQAAAAAAAAD/2wBDAAMCAgICAgMCAgIDAwMDBAYEBAQEBAgGBgUGCQgKCgkICQkKDA8MCgsOCwkJDRENDg8QEBEQCgwSExIQEw8QEBD/wAALCAACAAIBAREA/8QAFAABAAAAAAAAAAAAAAAAAAAACf/EABQQAQAAAAAAAAAAAAAAAAAAAAD/2gAIAQEAAD8AVN//2Q==',
  'base64',
);

async function documentXml(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  return await zip.file('word/document.xml').async('string');
}

/**
 * The body's top-level paragraphs and tables, in document order.
 *
 * The order of a caption and the block it names is the thing under test, so it
 * is read back out of the produced XML instead of being assumed from the order
 * the renderer happens to build its array in.
 */
function bodyBlocks(xml) {
  const body = xml.slice(xml.indexOf('<w:body>'), xml.indexOf('</w:body>'));
  const blocks = [];
  const element = /<w:(p|tbl)[ >][\s\S]*?<\/w:\1>/g;
  let match;
  while ((match = element.exec(body)) !== null) {
    blocks.push({
      kind: match[1],
      xml: match[0],
      text: [...match[0].matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map(t => t[1]).join(''),
    });
  }
  return blocks;
}

/** The `<w:spacing>` of the one block whose text contains `needle`. */
function spacingOf(blocks, needle) {
  const block = blocks.find(b => b.kind === 'p' && b.text.includes(needle));
  assert.ok(block, `expected a caption paragraph containing ${needle}`);
  const spacing = /<w:spacing\b[^>]*\/>/.exec(block.xml);
  assert.ok(spacing, `expected ${needle} to declare w:spacing`);
  return { xml: spacing[0], attr: name => new RegExp(`w:${name}="(-?\\d+)"`).exec(spacing[0])?.[1] };
}

function paragraphProperties(blocks, needle) {
  const block = blocks.find(b => b.kind === 'p' && b.text.includes(needle));
  assert.ok(block, `expected a caption paragraph containing ${needle}`);
  const properties = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(block.xml);
  assert.ok(properties, `expected ${needle} to declare w:pPr`);
  return properties[0];
}

async function renderWithCaptions() {
  const contentDir = mkdtempSync(join(tmpdir(), 'openthesis-caption-'));
  writeFileSync(join(contentDir, 'figure.png'), JPEG);
  const buffer = await renderDocument({
    outputPath: '', template, contentDir,
    document: {
      type: 'thesis', meta: { title: 'Captions' }, cover: [],
      sections: [{
        id: 'captions', type: 'chapter', title: 'Captions',
        content: [
          { type: 'table', number: '3-1', caption: '登记结果', headers: ['A', 'B'], data: [['1', '2']] },
          { type: 'figure', path: 'figure.png', number: '2-1', caption: '隧道纵向弯矩时程对比' },
        ],
      }],
    },
  });
  return bodyBlocks(await documentXml(buffer));
}

test('a table caption carries the table role spacing and precedes its table', async () => {
  // 《北京科技大学研究生学位论文书写指南》: 表题 above the table, 黑体 五号 (21
  // half-points) centred, 段前1行 / 段后0.1行. The template says the same in twips
  // through its `ub` style, and 段前1行 = beforeLines 100.
  const tableStyle = template.styles[template.roleWinners.table_caption];
  assert.equal(tableStyle.paragraph.spaceBefore, 150, '段前1行 is 150 twips in the template');
  assert.equal(tableStyle.paragraph.spaceAfter, 50, '段后0.1行 is 50 twips in the template');

  const blocks = await renderWithCaptions();
  const spacing = spacingOf(blocks, '3-1 登记结果');
  assert.equal(spacing.attr('before'), '150', `表题 段前, got ${spacing.xml}`);
  assert.equal(spacing.attr('after'), '50', `表题 段后, got ${spacing.xml}`);
  assert.equal(spacing.attr('line'), '312', 'single spacing within the template grid');

  const properties = paragraphProperties(blocks, '3-1 登记结果');
  assert.match(properties, /<w:jc w:val="center"\/>/, '表题 is centred');
  const run = /<w:r>[\s\S]*?<\/w:r>/.exec(blocks.find(b => b.text.includes('3-1 登记结果')).xml)[0];
  assert.match(run, /w:eastAsia="黑体"/, '表题 uses 黑体 for Chinese');
  assert.match(run, /w:ascii="Times New Roman"/, '表题 uses Times New Roman for Latin and digits');
  assert.match(run, /<w:sz w:val="21"\/>/, '表题 is 五号 (10.5pt = 21 half-points)');
  assert.match(run, /<w:b\/>|<w:b w:val="(?:true|1)"\/>/, '表题 is bold');

  const captionIndex = blocks.findIndex(b => b.kind === 'p' && b.text.includes('3-1 登记结果'));
  const tableIndex = blocks.findIndex(b => b.kind === 'tbl');
  assert.notEqual(tableIndex, -1, 'the table must be in the body');
  assert.ok(captionIndex < tableIndex, `表题 must precede its table (#${captionIndex} vs #${tableIndex})`);
  assert.equal(captionIndex + 1, tableIndex, 'nothing sits between the caption and its table');
});

test('a figure caption carries the figure role spacing and follows its figure', async () => {
  // 图题 below the figure, 段前0.1行 / 段后0.1行 — the template's `u4` style.
  const figureStyle = template.styles[template.roleWinners.figure_caption];
  assert.equal(figureStyle.paragraph.spaceBefore, 10, '段前0.1行 is 10 twips in the template');
  assert.equal(figureStyle.paragraph.spaceAfter, 10, '段后0.1行 is 10 twips in the template');

  const blocks = await renderWithCaptions();
  const spacing = spacingOf(blocks, '2-1 隧道纵向弯矩时程对比');
  assert.equal(spacing.attr('before'), '10', `图题 段前, got ${spacing.xml}`);
  assert.equal(spacing.attr('after'), '10', `图题 段后, got ${spacing.xml}`);
  assert.equal(spacing.attr('line'), '312', 'single spacing within the template grid');
  assert.doesNotMatch(spacing.xml, /w:(?:before|after)="120"/, 'the old hardcoded 120 must be gone');

  const properties = paragraphProperties(blocks, '2-1 隧道纵向弯矩时程对比');
  assert.match(properties, /<w:jc w:val="center"\/>/, '图题 is centred');

  const figureIndex = blocks.findIndex(b => b.kind === 'p' && /<w:drawing>/.test(b.xml));
  const captionIndex = blocks.findIndex(b => b.kind === 'p' && b.text.includes('2-1 隧道纵向弯矩时程对比'));
  assert.notEqual(figureIndex, -1, 'the figure must be in the body');
  assert.ok(captionIndex > figureIndex, `图题 must follow its figure (#${captionIndex} vs #${figureIndex})`);
  assert.equal(captionIndex, figureIndex + 1, 'nothing sits between the figure and its caption');
});

test('a caption separates its number from its title with exactly one space', async () => {
  const blocks = await renderWithCaptions();
  for (const needle of ['3-1 登记结果', '2-1 隧道纵向弯矩时程对比']) {
    const block = blocks.find(b => b.kind === 'p' && b.text.includes(needle));
    assert.ok(block, `expected a caption paragraph containing ${needle}`);
    assert.equal(block.text, needle, `caption text must use one space, got ${JSON.stringify(block.text)}`);
  }
  // A number that already carries its label must not be prefixed a second time.
  const numbered = await renderDocument({
    outputPath: '', template,
    document: {
      type: 'thesis', meta: { title: 'Labels' }, cover: [],
      sections: [{
        id: 'labels', type: 'chapter', title: 'Labels',
        content: [
          { type: 'table', number: '表3-1', caption: '  结果', headers: ['A'], data: [['1']] },
          { type: 'table', number: '表3-2', caption: '表3-2  结果', headers: ['A'], data: [['1']] },
        ],
      }],
    },
  });
  const texts = bodyBlocks(await documentXml(numbered))
    .filter(b => b.kind === 'p' && b.text.includes('结果'))
    .map(b => b.text);
  assert.deepEqual(texts, ['表3-1 结果', '表3-2 结果']);
});

test('a 表题 line above a table is the table caption, not body text', async () => {
  const blocks = parseMarkdownBlocks([
    '表 2  40 m资格模型结果',
    '',
    '| 方法 | 结果 |',
    '| --- | --- |',
    '| PCG | 通过 |',
  ].join('\n'));
  assert.deepEqual(blocks, [{
    type: 'table',
    caption: '表 2 40 m资格模型结果',
    headers: ['方法', '结果'],
    data: [['PCG', '通过']],
  }]);
  // 表3-1 and 表 3-1 are both numbering an author chose; only the separator
  // between the number and the title is the guide's business.
  assert.equal(
    parseMarkdownBlocks('表3-1  结果\n\n| A |\n| --- |\n| 1 |')[0].caption,
    '表3-1 结果',
  );
});

test('a 图题 line below a figure is the figure caption, not body text', async () => {
  const blocks = parseMarkdownBlocks('![图片](figure.png)\n\n图 1-1  隧道纵向弯矩');
  assert.equal(blocks.length, 1, `expected one block, got ${JSON.stringify(blocks)}`);
  assert.equal(blocks[0].type, 'figure');
  assert.equal(blocks[0].caption, '图 1-1 隧道纵向弯矩');
});

test('a sentence that merely starts with 图 is left as body text', async () => {
  const blocks = parseMarkdownBlocks('图1表明两种方法的结果一致。');
  assert.deepEqual(blocks, [{ type: 'paragraph', text: '图1表明两种方法的结果一致。' }]);
});
