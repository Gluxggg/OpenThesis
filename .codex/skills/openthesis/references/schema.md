# Structured content quick reference

Every content file is one of three tagged documents:

- Thesis: `{ "type": "thesis", "meta": {...}, "cover": [...], "sections": [...] }`
- Journal: `{ "type": "journal", "meta": {...}, "sections": [...] }`
- Official: `{ "type": "official", "meta": {...}, "body": [...] }`

Use the TypeScript definitions in `packages/document-schema/src/index.ts` as the authoritative contract. `validateDocument()` in the same package checks a content file against it, and `thesis build` runs that check before rendering — every problem is reported at once with a path such as `sections[0].content[2].headers`.

## Common blocks

```json
{ "type": "paragraph", "text": "Body text" }
{ "type": "heading2", "text": "Nested heading" }
{ "type": "equation", "latex": "E = mc^2" }
{ "type": "figure", "path": "figures/result.png", "caption": "Experimental result" }
{ "type": "table", "caption": "Results", "headers": ["A", "B"], "data": [["1", "2"]] }
{ "type": "list_item", "text": "First item", "ordered": true, "level": 0 }
{ "type": "code_block", "language": "ts", "text": "const x = 1;" }
{ "type": "blockquote", "text": "Quoted text" }
{ "type": "horizontal_rule" }
{ "type": "page_break" }
```

Every table row must have exactly one cell per header; `build` rejects a mismatch rather than padding or truncating it. If `columnWidths` is present, it must contain one positive number per header.

Figure paths may be absolute or relative to the content file. Supported detection includes PNG, JPEG, GIF, and BMP.

## Document-specific notes

- Thesis sections support recursive `subsections` and section types such as `chapter`, `abstract`, `references`, and `appendix`.
- Thesis `backMatter` may carry `references`, `appendices`, `authorBiography`, `declaration`, and `datasetInfo`. Only `references` is rendered today; the rest is preserved in the JSON.
- Journal metadata requires an `authors` array; each author has an `affiliations` array, which may be empty.
- Official metadata requires `issuingAuthority`, `documentNumber`, `documentCategory`, `primaryRecipients`, and `date`; `documentCategory` must be one of the GB/T 9704 categories (`通知`, `通告`, `报告`, …). Use block types such as `recipient_line`, `signature_block`, and `attachment_note` when direct control is needed.

Legacy `{cover_blocks, body_blocks}` JSON remains readable but should not be generated for new work. `validateLegacyDocument()` checks it with the same rules.
