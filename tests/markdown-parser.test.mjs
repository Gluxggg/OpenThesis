import test from 'node:test';
import assert from 'node:assert/strict';
import { parseMarkdown, parseMarkdownBlocks } from '../packages/markdown-parser/dist/index.js';

test('parses thesis front matter and nested Markdown sections', () => {
  const document = parseMarkdown(`---
type: thesis
title: "振动控制研究"
author: 张三
degree: master
keywords: [振动, 控制]
---
# This heading is not duplicated

## 第一章 绪论

研究背景正文。

### 研究方法

- 建立模型
- 完成实验

$$E = mc^2$$
`);

  assert.equal(document.type, 'thesis');
  assert.equal(document.meta.title, '振动控制研究');
  assert.deepEqual(document.meta.keywords, ['振动', '控制']);
  assert.equal(document.sections.length, 1);
  assert.equal(document.sections[0].title, '第一章 绪论');
  assert.equal(document.sections[0].subsections[0].title, '研究方法');
  assert.deepEqual(
    document.sections[0].subsections[0].content.map(block => block.type),
    ['list_item', 'list_item', 'equation'],
  );
});

test('parses GFM-style tables, figures, quotes, code, and page breaks', () => {
  const blocks = parseMarkdownBlocks(`
| 指标 | 数值 |
| --- | ---: |
| 精度 | 0.98 |

![实验结果](figures/result.png "Figure 1")

> 可复现性优先。

\`\`\`ts
const value = 42;
\`\`\`

<!-- pagebreak -->
`);

  assert.deepEqual(blocks.map(block => block.type), [
    'table', 'figure', 'blockquote', 'code_block', 'page_break',
  ]);
  assert.deepEqual(blocks[0].headers, ['指标', '数值']);
  assert.equal(blocks[1].path, 'figures/result.png');
  assert.equal(blocks[3].language, 'ts');
});

test('creates journal and official document metadata', () => {
  const journal = parseMarkdown(`---
type: journal
title: A Reproducible Study
authors: [Alice, Bob]
affiliation: Open Research Lab
keywords: [reproducibility, documents]
---
## Methods
Method details.
`);
  assert.equal(journal.type, 'journal');
  assert.equal(journal.meta.authors.length, 2);
  assert.equal(journal.meta.authors[0].affiliations[0].institution, 'Open Research Lab');
  assert.equal(journal.sections[0].type, 'methods');

  const official = parseMarkdown(`---
type: official
title: 关于开展测试工作的通知
authority: XX研究院
documentNumber: X研发〔2026〕1号
category: 通知
recipients: [各部门, 各中心]
date: 2026年8月30日
---
## 一、工作要求
请认真组织实施。
`);
  assert.equal(official.type, 'official');
  assert.equal(official.meta.issuingAuthority, 'XX研究院');
  assert.deepEqual(official.meta.primaryRecipients, ['各部门', '各中心']);
  assert.equal(official.body[0].type, 'heading2');
});

test('rejects unsupported front-matter document types', () => {
  assert.throws(
    () => parseMarkdown('---\ntype: slides\n---\n# Demo'),
    /Unsupported Markdown document type/,
  );
});


test('preserves quoted front-matter scalars as strings', () => {
  const document = parseMarkdown(`---
title: "00123"
studentId: '00042'
author: "false"
---
# Demo`);
  assert.equal(document.meta.title, '00123');
  assert.equal(document.meta.studentId, '00042');
  assert.equal(document.meta.author, 'false');
});

test('requires a closing code fence at least as long as its opener', () => {
  for (const marker of ['`', '~']) {
    const blocks = parseMarkdownBlocks([
      marker.repeat(4) + 'markdown',
      marker.repeat(3),
      '# literal heading',
      marker.repeat(3),
      marker.repeat(5),
      '',
      'After the fence.',
    ].join('\n'));
    assert.deepEqual(blocks, [
      { type: 'code_block', text: [marker.repeat(3), '# literal heading', marker.repeat(3)].join('\n'), language: 'markdown' },
      { type: 'paragraph', text: 'After the fence.' },
    ]);
  }
});

test('leaves ordinary punctuation alone while still stripping real emphasis', () => {
  // `(\*|_)(.*?)\1` treated any two delimiters as a pair, so these were mangled.
  assert.deepEqual(parseMarkdownBlocks('my_file_name and other_thing'), [
    { type: 'paragraph', text: 'my_file_name and other_thing' },
  ]);
  assert.deepEqual(parseMarkdownBlocks('compute 2 * 3 * 4 now'), [
    { type: 'paragraph', text: 'compute 2 * 3 * 4 now' },
  ]);
  // `<[^>]+>` also ate ordinary comparisons and every autolink.
  assert.deepEqual(parseMarkdownBlocks('if a < b > c then'), [
    { type: 'paragraph', text: 'if a < b > c then' },
  ]);
  assert.deepEqual(parseMarkdownBlocks('see <https://example.com/x> here'), [
    { type: 'paragraph', text: 'see https://example.com/x here' },
  ]);
  assert.deepEqual(parseMarkdownBlocks('this is *italic* and **bold**'), [
    { type: 'paragraph', text: 'this is italic and bold' },
  ]);
  assert.deepEqual(parseMarkdownBlocks('an _emphasised_ word'), [
    { type: 'paragraph', text: 'an emphasised word' },
  ]);
});

test('keeps a trailing hash that is not a closing sequence', () => {
  // CommonMark needs a space before the closing `#` run, so `# C#` keeps it.
  assert.equal(parseMarkdown('# C#\n\nbody').meta.title, 'C#');
  assert.equal(parseMarkdown('# Title ##\n\nbody').meta.title, 'Title');
});

test('reads a front-matter block sequence', () => {
  const document = parseMarkdown(`---
title: T
keywords:
  - alpha
  - beta
---
# Demo`, { documentType: 'journal' });
  // These items used to be skipped entirely and the key silently became "".
  assert.deepEqual(document.meta.keywords, ['alpha', 'beta']);
});

test('keeps a comma inside a quoted list element', () => {
  const document = parseMarkdown(`---
title: T
keywords: ["a, b", c]
---
# Demo`, { documentType: 'journal' });
  assert.deepEqual(document.meta.keywords, ['a, b', 'c']);
});

test('detects a single-column table', () => {
  assert.deepEqual(parseMarkdownBlocks('| A |\n| --- |\n| 1 |'), [
    { type: 'table', caption: '', headers: ['A'], data: [['1']] },
  ]);
});

test('accepts an extended fence info string', () => {
  const blocks = parseMarkdownBlocks('```js title=x\nconst a = 1;\n```\n\nafter paragraph');
  // The old `[\w.+-]*` info-string pattern failed to match, so the opener
  // became a paragraph and the closing fence swallowed the rest of the file.
  assert.deepEqual(blocks, [
    { type: 'code_block', text: 'const a = 1;', language: 'js title=x' },
    { type: 'paragraph', text: 'after paragraph' },
  ]);
});

test('derives list nesting from the marker column', () => {
  // `floor(indent / 2)` put 4-space nesting on level 2 and 8-space on level 4.
  const twoSpace = parseMarkdownBlocks('- top one\n- top two\n  - nested\n    - deeper\n- top three');
  assert.deepEqual(twoSpace.map(block => block.level), [0, 0, 1, 2, 0]);

  const fourSpace = parseMarkdownBlocks('- a\n    - b\n        - c');
  assert.deepEqual(fourSpace.map(block => block.level), [0, 1, 2]);

  // Dedenting returns to the enclosing level rather than starting a new one.
  const dedent = parseMarkdownBlocks('- a\n    - b\n- c');
  assert.deepEqual(dedent.map(block => block.level), [0, 1, 0]);
});

test('keeps the text that follows a closed display equation', () => {
  // `$$a=b$$ trailing` was read as an *opening* `$$`, so the scan below ate
  // every following line up to the next one ending in `$$` — headings included.
  const blocks = parseMarkdownBlocks('$$a=b$$ trailing\n\nnext paragraph\n\n## Heading');
  assert.deepEqual(blocks.map(block => block.type), ['equation', 'paragraph', 'paragraph', 'heading2']);
  assert.equal(blocks[0].latex, 'a=b');
  assert.equal(blocks[1].text, 'trailing');
  assert.match(blocks[2].text, /next paragraph/);
});

test('a leading `---` only strips a block that really is front matter', () => {
  // Everything between the two rules used to vanish, with `frontMatter: {}` and
  // no warning, because any later `---` ended a front-matter block.
  const blocks = parseMarkdownBlocks('---\nplain first line\nplain second line\n---\n# Heading');
  assert.deepEqual(blocks.map(block => block.type), ['horizontal_rule', 'paragraph', 'horizontal_rule', 'heading1']);
  assert.match(blocks[1].text, /plain first line/);
  assert.match(blocks[1].text, /plain second line/);

  // Real front matter is still stripped.
  assert.deepEqual(parseMarkdownBlocks('---\ntitle: T\n---\n# Heading'), [{ type: 'heading1', text: 'Heading' }]);
});

test('refuses metadata it cannot honour instead of coercing it', () => {
  // An unknown category became 通知 and `degree: PhD` became `undefined`, so a
  // mistyped value was indistinguishable from a deliberate one — and the
  // validator downstream never saw the invalid value at all.
  assert.throws(
    () => parseMarkdown('---\ntype: official\ncategory: BudgetRequest\n---\n# Notice'),
    /Unsupported official document category: BudgetRequest/,
  );
  assert.equal(parseMarkdown('---\ndegree: PhD\n---\n# C').meta.degree, 'doctor');
  assert.equal(parseMarkdown('---\ndegree: 硕士\n---\n# C').meta.degree, 'master');
  assert.throws(() => parseMarkdown('---\ndegree: postdoc\n---\n# C'), /Unsupported degree: postdoc/);
});
