# Markdown import contract

OpenThesis accepts UTF-8 `.md` and `.markdown` files. It intentionally supports a predictable manuscript subset rather than every Markdown extension.

## Front matter

Place YAML-style metadata at the top. Scalars, inline arrays, and block sequences are read:

```markdown
---
type: thesis
title: A Study of Open Documents
author: Jane Doe
degree: doctor
keywords: [documents, reproducibility]
acknowledgements:
  - The library staff
  - My supervisor
---
```

Shared fields: `type`, `title`, `date`, `keywords`, `abstract`.

Thesis fields: `titleEn`, `author`, `supervisor`, `department`, `major`, `degree`, `abstractEn`, `studentId`.

Journal fields: `authors`, `affiliation` or `affiliations`, `journalName`, `articleType`, `doi`. Inline author arrays create authors with the listed shared affiliations.

Official-document fields: `issuingAuthority` (aliases `authority`, `organization`), `documentNumber`, `documentCategory` (alias `category`), `primaryRecipients` (alias `recipients`), `ccRecipients`, `signatureAuthority`, `date`, `urgency`.

The `type` value is `thesis`, `journal`, or `official`. The CLI `--type` option overrides front matter when explicitly supplied.

Not supported — these are silently ignored, so do not rely on them: nested maps (`meta:\n  x: 1`), multi-line scalar continuations, and multi-document front matter.

## Supported body syntax

| Markdown | OpenThesis block |
| --- | --- |
| `#` through `######` | Nested sections or heading blocks |
| Paragraphs | `paragraph` |
| `- item` / `1. item` | `list_item` with nesting and order |
| Fenced code | `code_block`; the whole info string after the fence is kept as `language` |
| `> quote` | `blockquote` |
| `$$...$$` | `equation` |
| `![caption](path)` | `figure` |
| Pipe tables | `table` (single-column tables included) |
| `---` | `horizontal_rule` |
| `<!-- pagebreak -->` | `page_break` |

The first level-one heading becomes the document title and is not duplicated in the body; front-matter `title` takes precedence over its text. Any level-one heading is consumed this way, not only one at the very top. Remaining headings become nested thesis or journal sections; official-document headings remain body heading blocks.

An image caption falls back in this order: the alt text, then the title, then the file name. Reference-style images are not supported and appear as literal text.

Not supported: setext headings (`Title\n=====`), lazy blockquote continuation, and list items whose text continues on the next line — each becomes a separate paragraph.

Inline emphasis, links, and code are converted to plain text because the current content schema stores paragraph text without inline Markdown ranges.

## Media paths

During direct Markdown builds, relative image paths resolve from the Markdown file's directory. After importing to JSON, paths resolve from the JSON file's directory.
