# @openthesis/docx-renderer

> **Template-Driven DOCX Renderer** — convert structured JSON content into submission-ready `.docx` files.

## What it does

```
  template.json  +  content.json
        │               │
        └───────┬───────┘
                ▼
         docx-renderer
         (dolanmiu/docx)
                │
                ▼
          output.docx
```

## Key Features

- **Template-driven** — all formatting (fonts, sizes, margins, spacing) from parsed template
- **Block dispatcher pattern** — extensible content block → DOCX element routing
- **Complete table formatting** — gridlines, header shading, column widths
- **Headers/footers/page numbers** — configurable
- **Three document types** — thesis (学位论文), journal (期刊), official document (公文)
- **Backward compatible** — supports legacy `{cover_blocks, body_blocks}` JSON format

## Install

```bash
npm install @openthesis/docx-renderer
```

## Usage

```ts
import { renderDocument, createUSTBTemplate } from '@openthesis/docx-renderer';
import type { ThesisDocument } from '@openthesis/document-schema';

const template = createUSTBTemplate(); // or load from parsed template.json

const doc: ThesisDocument = {
  type: 'thesis',
  meta: { title: 'My Thesis' },
  cover: [
    { type: 'centered_text', text: 'XX大学', font_size_pt: 26, bold: true },
    { type: 'page_break' },
  ],
  sections: [
    {
      id: 'intro',
      type: 'chapter',
      title: 'Introduction',
      content: [
        { type: 'paragraph', text: 'This is the introduction...' },
      ],
    },
  ],
};

const buffer = await renderDocument({
  outputPath: 'output.docx',
  template,
  document: doc,
});
```

## Supported Content Blocks

| Block Type | Description |
|-----------|-------------|
| `heading1`–`heading4` | Section headings |
| `paragraph` / `paragraph_no_indent` | Body text |
| `centered_text` | Cover page titles |
| `equation` / `equation_numbered` | LaTeX equations |
| `figure` | Image with caption |
| `table` | Data table with caption |
| `list_item` | Ordered, unordered, or custom-marker list item |
| `code_block` | Monospaced multiline code block |
| `blockquote` | Indented quotation with a left rule |
| `red_header` | 公文红色发文机关标志 |
| `document_number` | 发文字号 |
| `recipient_line` | 主送/抄送机关 |
| `signature_block` | 发文机关署名+日期 |
| `attachment_note` | 附件说明 |
| `spacer` / `page_break` / `horizontal_rule` | Layout |

Equations are emitted as editable native Office Math for the syntax documented in [`@openthesis/equation-engine`](../equation-engine/README.md). Unsupported commands are preserved as readable text where possible.

## Related

- [OpenThesis](https://github.com/Gluxggg/OpenThesis) — Full document template engine
- [@openthesis/template-parser](https://www.npmjs.com/package/@openthesis/template-parser) — Parse DOCX → JSON
- [@openthesis/equation-engine](https://www.npmjs.com/package/@openthesis/equation-engine) — LaTeX → OMML

## License

MIT
