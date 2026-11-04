# Submitting to a journal

A journal submission is a template-driven build, the same shape as a thesis, with
one extra step at the start: the journal's own template is the specification, and
what it says has to be read before anything is written.

```bash
# 1. Parse the journal's template
node .codex/skills/openthesis/scripts/openthesis.mjs parse journal-template.docx \
  --type journal --org "Journal Name"

# 2. Write the content (see "What the content looks like", below)
# 3. Build
node .codex/skills/openthesis/scripts/openthesis.mjs build manuscript.json \
  -t journal-template.template.json -o manuscript.docx
```

`parse` writes `journal-template.template.json` beside the source. Read the output
before going further: it lists which style drives each role, and every warning.

## Two kinds of journal template

Both are common, and they are read differently.

| Template | What it carries | How it is read |
|---|---|---|
| **A styled template** | named Word styles — Elsevier's `ElsArticleTitle`, `ElsAuthor`, `ElsAbstractText`, `ElsKeyword`, `ElsFigureCaption`, `ElsReferences`, `ElsHeading1..5`; ACM's `Title`, `Authors`, `Affiliation`, `Abstract`, `KeyWords`, `FigureCaption`, `TableCaption`, `Bibliography`; Pandoc's `reference.docx` names | **by style name**. The role detector knows the Elsevier, ACM, Pandoc and Chinese-thesis families |
| **A fill-in instruction sheet** | almost no named styles, but body text that states the rules: `Title: [Max. 20 words]`, `Keywords: [At least 3]`, the sections it expects, a fixed opening table | **by its text**. `parse` stores the limits on `template.requirements`, and `build` reports every unmet one |

A template with neither named styles nor instruction text yields the built-in
journal fallback; the output is readable but not the journal's own typography.

## What the content looks like

```json
{
  "type": "journal",
  "meta": {
    "title": "…",
    "authors": [
      { "name": "San Zhang",
        "affiliations": [{ "institution": "…", "department": "…", "city": "…", "country": "…" }],
        "isCorresponding": true, "email": "…" }
    ],
    "language": "en",
    "abstract": "…",
    "keywords": ["…", "…", "…"]
  },
  "specifications": [
    { "label": "Hardware name", "value": "OpenTurbidity" }
  ],
  "sections": [
    { "id": "intro", "type": "introduction", "title": "Introduction",
      "content": [
        { "type": "paragraph", "text": "… [[引用#r1]] …" },
        { "type": "table", "caption": "…", "headers": ["…"], "data": [["…"]] },
        { "type": "figure", "path": "figure.png", "caption": "…" },
        { "type": "equation_numbered", "latex": "P = \\frac{a}{b}" }
      ] }
  ],
  "backMatter": { "references": [ { "id": "r1", "…": "…" } ] }
}
```

- `meta.language` decides the labels: `Abstract.` / `Keywords:` against `摘要：` / `关键词：`.
- `specifications` is only for a journal that opens with a fixed label/value table
  (HardwareX does). It renders without a header row, because the journal's table has none.
- Leave the number out of a section title and the template's own numbering applies.
  Write `1. Introduction` and the template's numbering is suppressed, so the heading
  is not numbered twice.
- `[[引用#r1]]` becomes a superscript citation; `[[表1-1]]`, `[[图1-1]]`, `[[式1-1]]`
  become cross-reference fields. A reference to something that does not exist warns
  and stays as text rather than printing a number that points nowhere.

## What the generator does on its own

- Page geometry, margins, columns and column gutter, from the template.
- Fonts, sizes, bold, alignment, spacing, indents, from the template's style for each role.
- Section headings, at the template's own level of numbering.
- Figure and table captions, with their own styles, numbered straight through.
- Equations as native OMML, numbered `(1)`, `(2)`.
- References set out to GB/T 7714-2015, with English author names turned round to
  `SMITH J`, and a hanging indent.
- The journal's stated limits, checked and reported before anything is built.

## What still needs a person

| Item | Why |
|---|---|
| **A table spanning both columns** | A wide table on a two-column page needs to interrupt the column flow. The renderer keeps a table inside one column. |
| **Journal-specific blocks** | `ElsArticlehistory` (submission and acceptance dates), `ElsGraphAbs` (graphical abstract), `ElsFootnote`, `ElsChemEquation` are recognised as roles but the content model has no field to fill them. |
| **Per-journal flourishes** | Superscript affiliation markers, the corresponding-author star, an article-history line. These are small and journal-specific. |
| **Anything the template does not declare** | The fallback is readable, not the journal's typography. Ask for the publisher's Word template. |

## Checking the result

Render the output and look at it. A `.docx` that opens is not a `.docx` that is right:

```bash
# LibreOffice Kit, or any DOCX viewer
node <libreoffice-kit>/cli.js render --input manuscript.docx --output-dir render/
```

Look for: the title in the template's title style, the abstract and keywords
labelled in the article's language, captions numbered without a chapter prefix,
a table that fits its column, and references with the GB/T 7714 markers.
