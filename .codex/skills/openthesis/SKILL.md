---
name: openthesis
description: "Convert academic manuscripts and Chinese official documents between Markdown or structured OpenThesis JSON and submission-ready DOCX, and parse DOCX templates into reusable style JSON. Use for thesis, journal, 公文, Markdown-to-DOCX, Word template parsing, document rendering, format validation, or troubleshooting OpenThesis CLI workflows."
---

# OpenThesis

Use the repository's typed pipeline instead of constructing Word files ad hoc:

`DOCX template → template JSON` and `Markdown/JSON content + template JSON → DOCX`.

## Choose the workflow

- For Markdown input, read [references/markdown.md](references/markdown.md), then use direct build unless the user also needs editable JSON.
- For structured JSON, read [references/schema.md](references/schema.md) only when creating or repairing content blocks.
- For a custom `.docx` template, parse it first and then render with the resulting `.template.json`.
- For setup, exact commands, validation, and failure recovery, read [references/workflows.md](references/workflows.md).

## Execute

1. Locate the OpenThesis repository root. Prefer the current checkout; otherwise honor `OPENTHESIS_ROOT` when it points to a checkout.
2. Check that Node.js is at least 22 and that `packages/cli/dist/index.js` exists. If it does not, install with the frozen lockfile and build once.
3. Run `node .codex/skills/openthesis/scripts/openthesis.mjs <command> ...` from any directory inside the checkout. The subcommand is bare (`build`, `parse`, `import`, `init`); the `thesis build …` spelling in the README is the installed `thesis` bin and is also accepted.
4. Keep source files unchanged. Always use a distinct output path and resolve image paths relative to the Markdown or JSON content file. `build` refuses a path that resolves to the input or the template.
5. Use a user-provided parsed template whenever available. If none is supplied, state that the built-in USTB fallback is being used and do not claim institution-specific compliance. Relay any `Template: …` warnings `parse` or `build` prints.
6. After modifying the engine or skill, run `pnpm check` and the Skill Creator validator before handing off. The validator's path differs per platform; do not hardcode it.

## Guardrails

- Do not invent university, publisher, or GB/T formatting rules that are absent from the parsed template.
- Do not silently discard unsupported Markdown; preserve it as plain paragraph text when possible and report material limitations. A parsed template's `warnings` array lists what the engine could not represent — pass it on.
- Treat generated DOCX files as outputs, not source-of-truth content.
- For equations, describe the supported core subset as native OMML and disclose that advanced macros or environments may fall back or render literally.
