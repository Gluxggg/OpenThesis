# OpenThesis — AI-powered Document Template Engine

## Project Overview
Parse any .docx template (thesis, journal article, or official document), write content in structured JSON, export a submission-ready DOCX. The same engine handles all three document types — the template parser and renderer are document-type-agnostic.

## Architecture

```
packages/
  document-schema/     → Domain models: Thesis, JournalArticle, OfficialDocument
  markdown-parser/     → Markdown + front matter → typed document schema
  template-parser/     → .docx → JSON style DSL (JSZip + fast-xml-parser)
  docx-renderer/       → JSON content + template → .docx (dolanmiu/docx)
  equation-engine/     → LaTeX math AST + Unicode/Pandoc compatibility fallbacks
  cli/                 → `thesis parse|import|build|init` commands

.codex/skills/openthesis/ → Agent-ready skill, wrapper, and focused references

examples/
  sample-thesis.json   → Sample thesis content
```

## Three Document Types

| Type | CLI flag | Standards |
|------|----------|-----------|
| Thesis (学位论文) | `--type thesis` | University-specific formatting |
| Journal Article (期刊论文) | `--type journal` | Elsevier, Springer, IEEE, etc. |
| Official Document (公文) | `--type official` | GB/T 9704-2012 (《党政机关公文格式》) |

## Key Design Decisions

### 1. Template-Driven, Not Hardcoded
All formatting parameters come from the parsed `DocumentTemplate` object. The `createUSTBTemplate()` function provides a fallback.

### 2. Document-Type-Agnostic Core
The template-parser and docx-renderer don't care what kind of document they're processing. Only the document-schema types and CLI commands are document-type-specific.

### 3. Style Inheritance Resolution
DOCX styles have a `basedOn` chain. The template-parser recursively resolves inheritance so each style's JSON is fully self-contained.

### 4. Semantic Role Detection (Heuristic V1)
Style roles are detected by regex matching on style names plus outline level fallback. In V4, this will be replaced by ML-based layout understanding.

### 5. Official Document Support (V1)
Complete block types for Chinese official documents: `red_header`, `document_number`, `recipient_line`, `signature_block`, `attachment_note`.

## Commands

```bash
pnpm install
pnpm build

# Parse templates
thesis parse <template.docx> --type thesis|journal|official --org <name>

# Create samples
thesis init --type thesis    # → thesis-content.json
thesis init --type journal   # → journal-content.json
thesis init --type official  # → official-content.json

# Build documents
thesis build <content.json> -t <template.json> -o output.docx

# Import or directly build Markdown
thesis import <manuscript.md> --type thesis -o content.json
thesis build <manuscript.md> --type thesis -t <template.json> -o output.docx
```

## Current Limitations

| Area | Limitation | Plan |
|------|-----------|------|
| Style role detection | Regex-based, fails on auto-numbered styles | V4: ML/layout-based |
| Equation rendering | Native OMML covers core math plus common functions, limits, accents, delimiters and matrix environments | V3: custom macros, array specs and uncommon AMS environments |
| Template parsing | Works best with real Word templates | Current: works for most templates |
| Markdown inline styling | Paragraph schema stores plain text | Future: rich inline ranges |
| Markdown dialect | Predictable manuscript subset | Expand only with compatibility tests |
| Multi-section templates | The cover, front matter and body each get their own section; a template that changes its setup mid-chapter (landscape, a two-column passage) is still flattened into the part it falls in | Map a part to a page setup the content can name |
| Document grid | The template snaps text to a 312-twip `w:docGrid`; the writer library emits its own `linePitch="360"` | Write `w:docGrid` in the post-pass |
| Unsupported LaTeX | The Unicode fallback is lossy; `convertLatexToOmml` reports it, but the AST path has no diagnostic channel | Surface unsupported syntax on the rendered document |

See `AUDIT.md` for the full audit, including the findings that are still open.

## Diagnostics

Both entry points now report problems instead of failing silently:

- `thesis parse` prints a `warnings` list — a template that carries no
  formatting at all, a style inheriting from an undefined style, or a document
  with more sections than can be rendered. The same array is stored on the
  template JSON.
- `thesis build` runs `validateDocument` / `validateLegacyDocument` first and
  reports every problem at once with a path such as
  `sections[0].content[2].headers`, instead of failing inside the OOXML builder.

## The built-in USTB template

`assets/ustb-thesis-template.json` was regenerated from
`《北京科技大学硕士学位论文模板》.docx` with the fixed parser: all 72 styles now
declare formatting, and `roleWinners` records which style drives each role. The
numbers match `《北京科技大学研究生学位论文书写指南》` — 一级标题 黑体 小三 加粗
居中, 二级/三级标题 黑体 四号 加粗, 正文 宋体 小四 with a 2-character first-line
indent.

Two parts of that file come from different places, and a regeneration must keep
them apart:

- **the style table is curated.** The template file does not declare everything
  the guide asks for — `w:pageBreakBefore` on chapter headings, the 1 cm / 1.25 cm
  hanging indents, and the three cover tiers were written into the asset from the
  guide. Regenerating the whole file from a `.docx` therefore *loses* them.
- **the section facts are parsed.** `page`, `pageSections`, `evenAndOddHeaders`
  and `warnings` are read out of the template's `document.xml`, `settings.xml`
  and `header*/footer*` parts. Refresh those four keys after a parser change
  rather than hand-editing them:

```bash
node packages/cli/dist/index.js parse "<template>.docx" --type thesis --org "北京科技大学"
```

## Sections

A thesis is rendered as up to three sections — cover, front matter, body — with
the split derived from the content (`cover`, then the `abstract` / `toc` /
`list_of_figures` / `list_of_tables` sections, then everything else plus the back
matter). A template that declares more than one `pageSections` entry drives each
part's geometry, page-number format and running heads; a template that declares
none still produces the single section it always did.

## Next Steps
1. Add explicit unsupported-syntax diagnostics and custom macro expansion to the equation engine
2. Get real university and journal templates to expand parser compatibility fixtures (`.gitignore` now allows `tests/fixtures/**/*.docx`)
3. Add ML-assisted layout-role understanding behind deterministic fallbacks
4. Add agent content-generation workflows on top of the typed schema
5. Build a community-contributed template marketplace
