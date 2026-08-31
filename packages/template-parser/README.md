# @openthesis/template-parser

> **DOCX + PDF Template Parser** — parse DOCX styles or infer text-PDF layouts into the same structured JSON style DSL.

## What it does

```
  模板.docx / 格式说明.pdf
       │
       ▼
  template-parser
  (Open XML / PDF.js)
       │
       ▼
  template.json
  {
    "meta": { "organization": "XX大学", "documentType": "thesis" },
    "page": { "width": 11906, "height": 16838, "margins": {...} },
    "styles": {
      "heading1": { "font": { "eastAsia": "黑体", "size": 30 }, ... },
      "paragraph": { "font": { "eastAsia": "宋体", "size": 24 }, ... }
    },
    "styleRoles": { "heading1": "heading1", ... }
  }
```

## Key Features

- **Style inheritance resolution** — recursively resolves DOCX `basedOn` chains
- **Text-PDF layout inference** — clusters positioned text into fonts, sizes, spacing, alignment, margins, columns, and semantic roles
- **Page geometry extraction** — margins, page size, columns from section properties
- **Semantic role detection** — heuristically maps style names to block types
- **Chinese + Western font separation** — `eastAsia` vs `name` font attributes
- **Document-type agnostic** — works for thesis, journal, and official doc templates

## Install

```bash
npm install @openthesis/template-parser
```

## Usage

```ts
import { parseTemplate } from '@openthesis/template-parser';
import { parsePdfTemplate } from '@openthesis/template-parser';
import { readFileSync } from 'fs';

const buffer = readFileSync('template.docx');
const template = await parseTemplate(buffer, {
  organization: 'XX大学',
  documentType: 'thesis',
});

console.log(template.styles);
console.log(template.page);

const pdfTemplate = await parsePdfTemplate(readFileSync('format-guide.pdf'), {
  organization: 'XX期刊',
  documentType: 'journal',
});
```

## How it works

A `.docx` file is a ZIP archive. The parser:

1. **Unzips** the `.docx` with JSZip
2. **Parses** `word/styles.xml` to extract paragraph and character styles
3. **Resolves inheritance** — if style "Heading 1" is `basedOn="Normal"`, merges all inherited properties
4. **Extracts page settings** from `word/document.xml` section properties
5. **Detects semantic roles** — regex matching on style names + outline level fallback
6. **Outputs** a clean JSON DSL

For a text-based PDF, PDF.js extracts page geometry and positioned text. The parser groups text into lines, clusters typography into reusable styles, estimates margins/columns/spacing, and maps styles to semantic roles. Scanned PDFs require OCR, and PDF-derived warnings should be reviewed before production use.

## Related

- [OpenThesis](https://github.com/Gluxggg/OpenThesis) — Full document template engine
- [@openthesis/docx-renderer](https://www.npmjs.com/package/@openthesis/docx-renderer) — Render JSON → DOCX
- [@openthesis/document-schema](https://www.npmjs.com/package/@openthesis/document-schema) — TypeScript type definitions

## License

MIT
