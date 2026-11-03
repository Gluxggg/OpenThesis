# @openthesis/template-parser

> **Universal DOCX Template Parser** — parse any `.docx` template into structured JSON style DSL.

## What it does

```
  学校模板.docx
       │
       ▼
  template-parser
  (JSZip + fast-xml-parser)
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
import { readFileSync } from 'fs';

const buffer = readFileSync('template.docx');
const template = await parseTemplate(buffer, {
  organization: 'XX大学',
  documentType: 'thesis',
});

console.log(template.styles);
console.log(template.page);
```

## How it works

A `.docx` file is a ZIP archive. The parser:

1. **Unzips** the `.docx` with JSZip
2. **Parses** `word/styles.xml` to extract paragraph and character styles
3. **Resolves inheritance** — if style "Heading 1" is `basedOn="Normal"`, merges all inherited properties
4. **Extracts page settings** from `word/document.xml` section properties
5. **Detects semantic roles** — regex matching on style names + outline level fallback
6. **Outputs** a clean JSON DSL

## Related

- [OpenThesis](https://github.com/Gluxggg/OpenThesis) — Full document template engine
- [@openthesis/docx-renderer](https://www.npmjs.com/package/@openthesis/docx-renderer) — Render JSON → DOCX
- [@openthesis/document-schema](https://www.npmjs.com/package/@openthesis/document-schema) — TypeScript type definitions

## License

MIT
