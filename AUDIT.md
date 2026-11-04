# OpenThesis — Engineering Audit

Date: 2026-09-28
Scope: `packages/*`, `tests/*`, CLI, agent skill, repository metadata
Method: full source read of all six packages, plus executable probes rendered against the built `dist/` — every defect below was reproduced, not inferred.

Baseline at audit time: `pnpm check` green, 29 tests passing, `pnpm audit --prod` **failing**.

---

## 1. Headline findings

Four problems made the engine produce wrong output on real templates. All four are
now fixed and covered by regression tests.

### 1.1 The built-in template produced headings identical to body text

`assets/ustb-thesis-template.json` resolved **all 72 styles to one identical
signature** — `Times New Roman / 宋体 / 12pt / justified / no indent`. Levels 1–3
of a nested heading rendered exactly like the surrounding body text, and level 4
rendered as 16pt 仿宋 (the *official-document* body font).

Two independent causes:

1. **`template-parser` read only style-level `w:rPr`** and ignored `w:pPr/w:rPr`,
   which is where Chinese Word/WPS templates put 黑体/三号 for headings.
2. **`rawToParagraphStyle` invented values** (`|| 'Times New Roman'`, `|| '宋体'`,
   `|| 24`, `?? 0`, `mapAlignment(undefined) → 'justified'`), so "the template is
   silent" became indistinguishable from "the template explicitly asks for 12pt".

The second cause is the important one: it is why the renderer could never apply a
sensible default. Verified against a realistic fixture, the parser itself was
correct — the failure was in what it emitted and what the renderer then trusted.

### 1.2 Tables overflowed the page

`renderTable` hardcoded `totalWidth = 8504` (an A4 constant) while the table
declared `<w:tblW w:type="pct" w:w="100%">` — contradictory OOXML.

Reproduced with an 8391-twip page and 1701-twip side margins (printable 4989):
cells still summed to **8504 twips, a 70% overflow** past the right margin.

### 1.3 Tables ignored the template entirely

Cell and caption runs hardcoded `Times New Roman` / 黑体 / 宋体 at 21 half-points.
A template requesting Georgia / 楷体 / 36 half-points was silently overruled; the
`_template` parameter was accepted and never read.

### 1.4 The repository's own CI was failing

`adm-zip@0.6.0` (a production dependency of `@openthesis/equation-engine`) carries
one **high** and one **moderate** advisory, both patched in `0.6.1`.

`pnpm audit --prod --audit-level high` is the last step of `.github/workflows/ci.yml`.
The last recorded CI run predates the advisory, so `main` is green on GitHub but
would fail on any push today.

---

## 2. Findings by severity

Status: **Fixed** = changed in this branch with a regression test.

### P0 — broken output or data loss

| # | Finding | Evidence | Status |
|---|---|---|---|
| P0-1 | `adm-zip@0.6.0` high + moderate advisories | `pnpm audit --prod` exit 1 | Fixed |
| P0-2 | `import a.md -o a.MD` overwrote the Markdown **source** with JSON — the guard was a case-sensitive string compare | reproduced | Fixed |
| P0-3 | `cleanInline` corrupted ordinary prose: `my_file_name` → `myfilename`, `2 * 3 * 4` → `2  3  4`, `if a < b > c then` → `if a  c then`; `<[^>]+>` also deleted every `<https://…>` autolink | reproduced | Fixed |
| P0-4 | `\text{if } x>0` rendered as `ifx>0` — whitespace was consumed before the script check and never emitted | reproduced | Fixed |

### P1 — wrong on real inputs

| # | Finding | Status |
|---|---|---|
| P1-1 | `heading4` absent from the fallback table → 4th-level headings resolved to the body paragraph style | Fixed |
| P1-2 | Table width hardcoded to 8504 twips, contradicting its own 100% declaration | Fixed |
| P1-3 | Table typography hardcoded; template styles ignored | Fixed |
| P1-4 | Thesis and 公文 shared one fallback table, so a thesis without a template body rendered as 仿宋 三号 | Fixed |
| P1-5 | Parser ignored `w:pPr/w:rPr` | Fixed |
| P1-6 | `w:firstLineChars` / `leftChars` / `rightChars` are hundredths of a *character* but were stored as twips (`200` instead of `480`) | Fixed |
| P1-7 | Unanchored `/^(表\|Table\|题注)/` matched Word's built-in `TableGrid` and `Table of Contents`, turning them into level-3 headings | Fixed |
| P1-8 | Role detection ignored `w:type`, so character (`uChar`, `Hyperlink`), table and numbering styles were all assigned `paragraph`; `resolveStyle` then picked whichever came first in the XML | Fixed |
| P1-9 | No pattern ever produced the `equation` role, so a template's equation style (`MTConvertedEquation`) could never be used | **Overstated — corrected.** The `equation` pattern was added, but `MTConvertedEquation` is a `w:type="character"` style and `detectStyleRoles` skips every non-paragraph style, so that specific style still receives no role (confirmed against the shipped asset: the key is absent from both `styleRoles` and `roleWinners`). What the fix does deliver is a role for *paragraph*-level equation styles (`公式`, `Equation`). Assigning `equation` to a character style would be meaningless anyway — the renderer styles a paragraph, and a character style carries no paragraph properties |
| P1-10 | `mathComponents` had no `default` case → an unhandled AST node put `undefined` into the OMML children array | Fixed |
| P1-11 | Equation parse failures were swallowed by a bare `catch` with no diagnostic | Fixed |
| P1-12 | `--flag=value` was silently ignored: `import x.md --type=official` exited 0 and produced `"type":"thesis"`; `init --type=journal` wrote `thesis-content.json` | Fixed |
| P1-13 | Unknown flags were silently ignored: `build x.json --out zzz.docx` wrote `x.docx` | Fixed |
| P1-14 | UTF-8-BOM JSON input failed with `Unexpected token '\uFEFF'` | Fixed |
| P1-15 | Front-matter block sequences (`keywords:` + `- item`) were skipped and the key became `""` | Fixed |
| P1-16 | Quoted list elements were split before quote handling: `["a, b", c]` → `['"a','b"','c']` | Fixed |
| P1-17 | `# C#` lost its `#`; the closing-hash pattern did not require a preceding space | Fixed |
| P1-18 | Single-column tables were never detected (separator regex required a second pipe group) | Fixed |
| P1-19 | A fence with an extended info string (` ```js title=x `) failed to match, so the opener became a paragraph and the closing fence swallowed the rest of the file | Fixed |
| P1-20 | `latexToPlainText` used an ordered replace chain, so any command that prefixes another was corrupted: `\propto` → `Πto`, `\cdots` → `·s`, `\simeq` → `∼eq` | Fixed |
| P1-21 | Bare `x^2` / `x_1` were left as literal caret/underscore | Fixed |
| P1-22 | `[^}]+` cannot match nested braces, so `\frac{a_{1}}{b}` leaked as `\fraca_1/b` | Fixed |
| P1-23 | `/\\(left\|right)\s*/` also matched `\leftarrow`, so the real `\right` was eaten and the delimiter swallowed the rest of the equation | Fixed |

### P2 — correctness and maintainability

Status: **Fixed** = changed on this branch with a regression test. Every finding
in the table below is now closed; the notes record what each fix does. §6 lists
what remains outside this table.

| # | Finding | Status |
|---|---|---|
| P2-1 | `latexToOMML` falls back to Unicode whenever pandoc is missing **or fails for any reason**, silently. | Fixed — `convertLatexToOmml` returns `{ source, lossy, reason }`, and `latexToOMML` warns. |
| P2-2 | `\operatorname*{argmax}_{x}`: the `*` is not consumed, yielding an empty function name plus a stray run and group. | Fixed |
| P2-3 | `\begin{aligned}…\end{matrix}` (mismatched name) silently consumes the rest of the input as the environment body. | Fixed — the scanner stops at depth 0 and leaves the index at the mismatched `\end`. |
| P2-4 | Unknown commands lose their backslash in the AST (`\foo` → run `foo`). | Fixed — `\foo` now stays `\foo`; escaped literals (`\%`) are still bare. |
| P2-5 | `x^{a}^{b}` silently overwrites the first superscript. | Fixed — the first wins. |
| P2-6 | Escaped `&` after a row break is mis-detected (`\\&`), merging two cells into one. | Fixed — a character is escaped by an *odd* run of backslashes, so `\\&` is a row break plus a separator and `\&` stays literal. |
| P2-7 | Only the body-level `w:sectPr` is read; paragraph-level section breaks are ignored. | Fixed — all sections are collected into `pageSections`; only the last is rendered, and that is now warned about. |
| P2-8 | `removeNSPrefix: false` hardcodes the `w:` prefix; a styles.xml using another prefix yields a zero-style template with no error. | Fixed — the prefix bound to the WordprocessingML namespace is read from the document's own declaration and normalised to `w:` before parsing, so an `x:styles` template parses like any other. A template that still declares nothing keeps its warning. |
| P2-9 | `basedOn` cycles are accepted and produce order-dependent merges. | Fixed — a loop is cut at its earliest-declared member, so the merge no longer depends on the entry point, and the loop is reported as a warning. A `basedOn` pointing at a missing style still warns. |
| P2-10 | `lineSpacing` stores the raw `w:line` and drops `lineRule`, so `auto` (240ths) and `exact` (twips) are conflated. | Fixed — `ParagraphStyle.lineSpacingRule` carries the rule from `w:lineRule` through to the rendered `w:spacing`. |
| P2-11 | Markdown list nesting is `floor(indent / 2)`, so standard 2- and 4-space nesting is mis-levelled. | Fixed — level comes from the marker's column. |
| P2-12 | `InlineRange` / `RichParagraph` exist in the schema and are referenced nowhere. | Fixed — both interfaces were removed. Nothing produced or consumed them, no block type can hold inline formatting, and the Markdown parser has no emphasis handling, so the types advertised a capability that did not exist. |
| P2-13 | Lists render as plain text runs with a manual `•` / `1.` prefix, not Word numbering. | Fixed — real OOXML numbering, each list restarting its own counter. |
| P2-14 | No document validator exists anywhere; the only runtime checks live in the renderer. | Fixed — `validateDocument` / `validateLegacyDocument`, run by `build`. |
| P2-15 | `$$a=b$$ trailing` mis-parses and then consumes following lines until one ends with `$$`. | Fixed — only an opener with no second `$$` starts a display block; the text after a closing `$$` becomes a paragraph. |
| P2-16 | A document starting with `---` and no front matter silently drops everything up to the next `---`. | Fixed — a `---` block is only stripped when its interior is front-matter-shaped, so a thematic break before a heading stays a thematic break. |
| P2-17 | An invalid official `category` silently becomes `通知`; `degree: PhD` silently becomes `undefined`. | Fixed — an unrecognised category is refused instead of coerced; degree spellings are normalised (`PhD` → `doctor`) and anything else is refused; `validateDocument` also checks `meta.degree`. |
| P2-18 | `resolveStyle` scans `styleRoles` linearly for every block — O(styles × blocks). | Fixed — `Object.entries(...).find(...)` materialised all 72 pairs before matching (9–11 µs per block, 18–24% of render time); the lookup now stops at the first match. |
| P2-19 | CLI: no `--version`, no stdin/stdout, no `mkdir -p` for the output directory, no overwrite protection for an existing output. | Fixed — `--version`/`-v`, `mkdir -p` for the output directory, an existing output refused unless `--force`/`-y`, and `-` for stdin (`build`, `import`) or stdout (`-o -`, with progress moved to stderr so the binary survives). |
| P2-20 | Agent-skill documentation mismatches. | Fixed — block sequences, fence info strings, caption precedence, H1 handling, `backMatter`, and the platform-specific validator path are all corrected. |

### Repository hygiene

| # | Finding | Status |
|---|---|---|
| H-1 | 24 references to the pre-rename owner `1771902720-lgtm` across 13 files (CI badge, clone URLs, star-history chart, every `package.json` `repository`/`homepage`/`bugs`). GitHub redirects them, so they mostly work, but npm metadata should be canonical. | Fixed |
| H-2 | `drm-fix-report.json` — a 549-line personal research report (北京科技大学 动力松弛法) committed to the public repository root. | Removed |
| H-3 | `thesis-content.json` — a `thesis init` output artifact committed at the root. | Removed |
| H-4 | `.gitignore`'s blanket `*.docx` also blocked template fixtures, directly obstructing the roadmap item "get real university and journal templates to expand parser compatibility fixtures". | Fixed (`examples/` and `tests/fixtures/` are now exempt) |

---

## 3. What changed

| Package | Change |
|---|---|
| `document-schema` | `FontSettings.name/eastAsia/size` and `ParagraphFormatting.alignment` are optional — a parsed template declares only what it specifies. Added `StyleRole` (`BlockType` + `table_header` + `figure_caption`). |
| `template-parser` | Merge `w:pPr/w:rPr` under the style-level `w:rPr`. Convert `*Chars` indents to twips. Stop inventing fallback values. Skip non-paragraph styles when assigning block roles. Anchor the caption pattern; add an `equation` pattern. Order role assignment so an explicit or `w:default` style wins the renderer's first-match lookup. |
| `docx-renderer` | Split `THESIS_STYLES` from `OFFICIAL_STYLES` (GB/T 9704-2012) and add `heading4`, `table`, `table_header`, `figure_caption`. `resolveStyle` merges the template **per property** over the role default. Table width derives from `page.width − margins`; table and caption typography resolve from the template. Image width fits the printable page. Exhaustive `mathComponents` switch. Equation fallback now warns. Lists use real OOXML numbering with per-list restart. |
| `equation-engine` | Exact-command symbol lookup (no prefix collisions). Brace-aware group matching for `\frac`, `\sqrt`, accents and `\text`. Bare-script conversion. `\text{}` keeps its spaces; the insignificant space after it is dropped. `\leftarrow` no longer matches `\left`. `\operatorname*`, unknown-command preservation, environment-name mismatch, double superscript, and a conversion diagnostic (`convertLatexToOmml`). |
| `cli` | `--flag=value` support, unknown-flag rejection, case-insensitive input/output equality (plus output ≠ template), BOM stripping, `init` shares the same option reader, the documented `thesis <command>` prefix, content validation before rendering, and template-warning relay. |
| `markdown-parser` | Emphasis requires real delimiters (no intraword `_`, no space-flanked `*`); only tag-shaped `<…>` is stripped and autolinks keep their target. Heading closing-hash requires a space. Front-matter block sequences. Quote-aware list splitting. Single-column tables. Extended fence info strings. List level derived from the marker column. |
| `assets/ustb-thesis-template.json` | Regenerated from `《北京科技大学硕士学位论文模板》.docx` with the fixed parser: 72 of 72 styles carry formatting (was 0), and `roleWinners` records which style drives each role. Verified attribute by attribute against `《北京科技大学研究生学位论文书写指南》` — see §5. |

Test count: **29 → 119**, all passing. `pnpm audit --prod --audit-level high` now exits 0.

---

## 4. Reproducing the original defects

The probes used during the audit are reproduced by the regression tests added in
`tests/`. To see the original behaviour, check out the parent commit and run:

```bash
pnpm install && pnpm build
pnpm audit --prod --audit-level high          # exits 1
node --test tests/*.test.mjs                  # 29 tests, all pass
```

Then render a four-level heading, or a table on a narrow page, and inspect
`word/document.xml` — the probes are described in §1.1–§1.3.

---

## 5. The built-in USTB template — regenerated

`assets/ustb-thesis-template.json` was generated by the **old** parser and carried
no per-style formatting: all 72 entries resolved to one identical signature. The
renderer compensated with role defaults, so output was plausible, but the file
itself said nothing.

The source `.docx` (`《北京科技大学硕士学位论文模板》.docx`) and the university's
`《北京科技大学研究生学位论文书写指南》.docx` were supplied, and the asset was
regenerated with the fixed parser:

```bash
node packages/cli/dist/index.js parse "《北京科技大学硕士学位论文模板》.docx" --type thesis --org "北京科技大学"
```

72 of 72 styles now declare formatting (was 0), and the rendered output matches
the guide, checked attribute by attribute:

| Element | Guide | Rendered |
|---|---|---|
| 一级标题 | 黑体 小三 (15pt), 加粗, 居中, 段前/段后 17pt, 1.3 倍行距 | 黑体 15pt, bold, center, 340/340 twips, line 312 auto |
| 二级标题 | 黑体 四号 (14pt), 加粗 | 黑体 14pt, bold |
| 三级标题 | 黑体 四号 (14pt), 加粗 | 黑体 14pt, bold |
| 正文 | 宋体 小四 (12pt), 首行缩进 2 字符 | 宋体 12pt, firstLine 480 twips |

Two defects had to be fixed before that was true, and both were found by
regenerating against a real template rather than by reading the code:

1. **Role assignment ignored what the template uses.** Word's stock `heading 2` /
   `heading 3` carry a heading role by name whether or not anyone applies them;
   the USTB document sets every heading in its own `u2级标题` / `u3级标题`. Ranking
   by key order cannot fix this either — a plain object enumerates integer-like
   keys first, and Word names those stock styles `1`, `2`, `3`. The parser now
   records an explicit `roleWinners` map: a name-pattern match beats an
   outline-level guess, and between equals the style the document applies more
   often wins.
2. **`w:firstLineChars` lost to the stale twips beside it.** Word writes both
   `w:firstLineChars="200"` (2 characters) and `w:firstLine="200"`, and the
   converter only used the character count when the twips attribute was absent —
   so a 2-character indent came out as 200 twips instead of 480.

---

## 6. Suggested next steps

Items 3–5 of the original list are done on this branch, and the built-in USTB
template has been regenerated (§5). What remains:

1. **Add real university and journal `.docx` files under `tests/fixtures/`** — now
   unblocked by the `.gitignore` change — and assert parsed geometry and roles.
2. **Emit real section breaks** for templates that declare more than one section
   geometry — currently all sections are read but only the last is rendered, and
   the template now warns about it.
3. **Inline formatting** (P2-12): the two unused types were removed. Modelling
   rich runs means new schema, `markdown-parser` emphasis handling, and
   `renderParagraph` splitting `text` at range boundaries.
4. **Unsupported LaTeX diagnostics**: surface unsupported syntax on the rendered
   document instead of only through `convertLatexToOmml`.

---

## 7. Submission readiness against the template

Measured against two authorities: `《北京科技大学研究生学位论文书写指南》.docx`, and a
**139-page accepted thesis** produced with the template, inspected page by page
(PyMuPDF: fonts, sizes, per-element ink boxes, odd/even text columns).

### Verified in the rendered OOXML

| Property | Evidence |
|---|---|
| Page size / margins | A4, 上3cm 下2cm 左右3cm — the template's own `w:pgMar` |
| 装订线 1cm | `w:gutter="567"`; the thesis adds it to the **inside** edge only |
| 对称页边距 | `w:mirrorMargins`; thesis text columns alternate 4.00/3.00 cm and 3.00/4.00 cm on odd/even pages |
| 一级标题 | 黑体 15pt bold centred, 340/340, 1.3 倍行距 |
| 二/三级标题 | 黑体 14pt bold, 260/260; hanging 1cm / 1.25cm (thesis measures level 3 at 708 twips) |
| 正文 | 宋体 12pt justified, `w:firstLine="480"`, `w:line="312"` |
| 特殊标题 (摘要/目录/序/附录/致谢/参考文献) | `w:line="579"` (2.41 倍), 340/330 — asserted in `tests/docx-renderer.test.mjs` |
| 图题 / 表题 | 黑体 10.5pt centred; the table caption has its own role and spacing |
| 参考文献条目 | 宋体 12pt, hanging 1cm, 10/10/312 |
| 页码与篇眉 | three sections: the cover has no head and no number, the front matter is `upperRoman` from I, the body restarts at 1 in decimal. The heads come from the template's own `header*.xml` parts — odd pages 北京科技大学硕士学位论文, even pages the *document's* own title, since the text a template carries in that slot is its author's sample thesis — with the 篇眉 0.5 pt rule, and `<w:evenAndOddHeaders/>` reaches `settings.xml` |
| 每节的页面设置 | each part carries its own `w:pgMar` (the template's cover footer distance is 851, the body's 850), `w:pgSz`, gutter and columns |
| 封面三级样式 | `cover_title` / `cover_line` / `cover_meta` resolve from the template: 小二 18pt bold centred, 四号 14pt bold justified, 五号 10.5pt |
| 目录 | a `type: "toc"` section emits the 目录 heading plus a real `TOC \h \o "1-3"` field |
| 装订线 1cm + 对称页边距 | `w:gutter="567"` and `w:mirrorMargins`; the accepted thesis alternates 4.00/3.00 cm columns. `w:mirrorMargins` is read from `settings.xml`, where it is actually declared |
| 三线表 (Three-line tables) | 学术论文表格默认输出标准三线表：顶底线 1.5pt，表头下线 0.75pt，无内外竖线，无表头底纹 |
| 公式编号靠右对齐 | 公式行居中，编号 `(X-Y)` 通过双制表位靠齐页面右边距（8504 twips） |
| 后置部分完整渲染 | `backMatter` 完整渲染 `references`、`declaration` (独创性说明)、`authorBiography` (作者简历及成果) 与 `appendices` |
| Markdown 封面生成 | Markdown front-matter 元数据自动生成带有 `cover_title`、`cover_line`、`cover_meta` 的北科大标准封面 |

**One step stays manual.** The directory page numbers come from the `TOC` field,
which only Word can evaluate: after `thesis build`, open the document and update
the field (Ctrl+A then F9, or right-click → 更新域). A generator has no layout
engine, so it cannot know on which page a chapter lands.

### Not yet delivered

1. **Cover page layout beyond the three tiers.** `cover_title` / `cover_line` /
   `cover_meta` are in place and verified, but the guide's 表 4 also fixes the
   *positions* of those boxes (间距 5 cm, 书脊), which the content model still
   expresses as spacers.
2. **TOC entry styles.** The field is emitted, but `TOC1/2/3` are still mapped to
   `paragraph`, so the entries Word generates keep Word's own TOC formatting
   rather than the template's.
3. **Sections that do not line up with a part boundary.** The renderer emits the
   cover, the front matter and the body; a template that changes its setup
   mid-chapter (a landscape page, a two-column passage) is still flattened into
   whichever part it falls in.
4. **Automatic chapter numbering.** The template's `numbering.xml` produces `1`,
   `1.1`, `2.3.3`; the renderer emits real list numbering for lists but not for
   headings, so a heading's number is whatever the content JSON says.
5. **Not expressible at all** (needs Word, documented rather than attempted):
   rasterised auto-numbers, TOC dot leaders,
   and English caption/title lines as separate styles.
6. **`w:docGrid`.** The template snaps text to a 312-twip line grid
   (`type="linesAndChars"` in the body); the writer library emits its own
   `linePitch="360"`, so a generated page will not break lines exactly where the
   template's does.

### Found by reading the template's own parts

The template package holds 47 parts; the tool now opens six of them
(`styles.xml`, `document.xml`, `settings.xml`, `document.xml.rels` and the
`header*/footer*` parts the document references), and still ignores the rest.
What reading them turned up, and what is still open:

1. **Chapter numbers exist only in `numbering.xml`.** They are bound to styles
   through `w:lvl/w:pStyle` plus the style's `w:numPr` (`%1`, `%1.%2`, `%1.%2.%3`
   for `1`, `u2`, `u3`; `附录 %1` from numId 4). A regex for `numPr|numId` over
   the parser's output is false, so nothing survives. Watch numId 0 on `u4` and
   the figure captions: that means "remove numbering", and a generator that
   treats every numId as "apply numbering" gets it backwards.
2. **`w:beforeLines` / `w:afterLines` are unparsed**, so `ua` reports 10 twips
   where the template means 0.1 line; and `firstLineChars` is converted with the
   style's *own* size rather than the inherited one, so `af8`/`aff` record 480
   where 420 is right.
3. **Equations change representation.** The template's twelve formulas are all
   OLE objects (8× MathType `Equation.DSMT4`, 4× AutoCAD) with WMF previews — it
   has **zero** `m:oMath` — while the renderer emits native OMML, and centres it
   where the template right-aligns at tab stop 8820. Its `MTConvertedEquation`
   style is `Cambria Math` italic and is never applied.
4. **`fixChineseFonts` is over-broad.** It strips every `w:*Theme` attribute and
   rewrites Latin `w:eastAsia` values to 仿宋. Harmless for this template,
   hostile to one that says what it means.

---

## 6. Follow-up round — defects this audit missed

A second review of the code after §1–§5 landed turned up ten more defects. Each
one below was reproduced against the built packages before it was changed, and
each has a named regression test.

| # | Finding | Evidence | Status |
|---|---|---|---|
| F-1 | A table without `caption` passed validation and then threw inside the renderer: `captionText` called `.trim()` on `undefined`. `TableBlock.caption` was declared required but `validateTable` only checked headers and data. | `thesis build` on a table block with no `caption` exited 1 with `Cannot read properties of undefined (reading 'trim')` | Fixed — the validator reports `…\.caption: must be a string`, and `captionText` accepts an absent caption. An **empty** caption stays valid: it means "no caption" and the renderer emits no caption paragraph, which `examples/sample-thesis.md` relies on |
| F-2 | `countStyleUsage` regex-scanned the **raw** `document.xml` for `<w:pStyle`, while the same part was parsed after `normalizeWordPrefix`. A template bound to another prefix counted zero uses for every style, so `pickRoleWinners` fell back to declaration order. | Two styles with `outlineLvl=0`, document applying `a2`: `w:` prefix → `a2` (right), `wp:` prefix → `a1` (an unused style) | Fixed — the part is normalised once and the count read from the parsed tree. The exported `countStyleUsage(string)` keeps its signature and now normalises internally |
| F-3 | `w:beforeLines` / `w:afterLines` were ignored: spacing read only the twips `w:before` / `w:after`, so a style declaring 段前/段后 in lines parsed to nothing. | `<w:spacing w:beforeLines="50" w:afterLines="30"/>` → `paragraph: {}` | Fixed — converted with the effective run size and preferred over the twips beside them, the same precedence `w:firstLineChars` already had |
| F-4 | `mergeStyles` built `rPr: { ...base.rPr, ...override.rPr }` unconditionally, so every inherited style ended with `rPr === {}`. Word writes `w:docDefaults` into every file, so the "no style declares any formatting" diagnostic could never fire. | A template with an empty `w:docDefaults` parsed to `font: {}` and `warnings: undefined`; deleting the element made the warning appear | Fixed — `mergeStyles` reuses `mergeRunProps`, which keeps `undefined` when both sides are absent |
| F-5 | Every render loaded and regenerated the whole package **twice**, and the second pass passed no `compression` option — so the file that reached disk was stored, not deflated. | `renderDocument` 19.5 ms, of which one load+DEFLATE repack is 6.7 ms; output 188.9 KB for 136.8 KB of content | Fixed — one `postProcess`: load once, patch every part, generate once with DEFLATE, and return the original buffer when nothing changed. Now 13.7 ms and 136.8 KB |
| F-6 | The journal renderer passed `docType: undefined` to `processBlocks`, so `resolveStyle` never consulted the template; a journal rendered from a journal template came out in the thesis fallback styles. 公文 is different — `resolveStyle` opts it out of the template on purpose, because GB/T 9704-2012 fixes its typography. | Journal heading rendered `黑体 30 half-points` only after the fix; before it followed `THESIS_STYLES` | Fixed for journal. 公文's hardcoded typography is intentional and documented at the call sites |
| F-7 | `flattenSections` and `flattenJournalSections` were character-for-character identical, and `HEADING_TYPES` was declared twice. | — | Fixed — one generic `flattenSections` |
| F-8 | `KNOWN_FLAGS` was one global set: every command accepted every flag, and a flag was honoured only if that command happened to call `readOption` for it. Same silent-no-op class the audit claimed to have closed for `--type=`. | `thesis init --org "Acme" -t missing.json` exited 0, wrote `thesis-content.json`, and used neither flag | Fixed — per-command flag tables, with `--force`/`-y` shared because all four commands enforce overwrite protection. `--help`/`--version` moved ahead of dispatch so `build --help` is no longer read as a content path |
| F-9 | `resolveStyle` ran at least once per block and three times per table, re-deriving an identical style: a `for…in` over 48 role entries for every role without a recorded winner (the shipped template records 12), then two `Object.entries`/`fromEntries` merges. | — | Fixed — cached in a `WeakMap` keyed by template identity, dropped at the start of every render so an in-place template edit is still read |
| F-10 | Dead plumbing: `mapAlignment(...) ?? AlignmentType.CENTER` could never fire (`mapAlignment` always returns a value); `renderSpacer` took a template it never used; `RenderOptions.outputPath` was declared required while every caller treated it as optional and the implementation guarded it. | — | Fixed — the unreachable fallback removed, the unused parameter dropped, `outputPath` declared optional to match how it is actually used |

`pnpm check`: 137 tests, all passing (120 before this round, 17 added).

---

## 7. Second follow-up round

The rest of the review backlog, plus the PDF feature from the long-stale PR #2.

| # | Finding | Evidence | Status |
|---|---|---|---|
| G-1 | `keepNext` / `keepLines` / `pageBreakBefore` never parsed out of a real document. Word writes them as *empty* elements, which fast-xml-parser turns into `""` — falsy — so `if (pPr['w:keepNext'])` never fired. | A template declaring `<w:keepNext/><w:keepLines/>` parsed to a style with neither flag | Fixed — one `flagOn` helper distinguishes "absent" from "present but empty", and the schema grew the `keepNext` / `keepLines` fields it never had, so the renderer can emit them |
| G-2 | Table cells passed `line` without `lineRule`, so an `exact` value was read as a multiple of the line height; they also hardcoded 段前/段后 40/20 over the template's own. | `spacing: { before: 40, after: 40, line: headerStyle.lineSpacing }` | Fixed — `lineStyle()` and the template's `spaceBefore`/`spaceAfter` with the old values only as fallbacks |
| G-3 | A caller-supplied figure `width` was never clamped to the printable page (only the no-size branch clamped); a zero-height header made `size.width / size.height` `Infinity`; and the placeholder said 未找到图片 even when the file was found but could not be embedded. | `width: 4000` reached `w:extent` unclamped | Fixed — every branch clamps while preserving the ratio, a zero dimension is rejected, and "无法嵌入图片" is now distinct from "未找到图片" |
| G-4 | `createUSTBTemplate` swallowed read/parse errors and silently returned the hand-written minimal template; it also re-read and re-parsed a 1,300-line JSON on every call. | — | Fixed — warns with the path and the error, then falls back; the parsed template is cached per process |
| G-5 | `isSamePath` lowercases unconditionally. On a case-sensitive filesystem `a.MD` and `a.md` are two files, so the guard refuses a legitimate output. | `tests/cli.test.mjs` pins the stricter behaviour | **Kept as-is, deliberately.** Folding only on Windows/macOS was tried and reverted: the same command line may later run where `a.md` and `a.MD` are one file, and an over-eager refusal cannot destroy a source file. The rationale is now recorded at the function |
| G-6 | `mapAlignment` returned `JUSTIFIED` for an undeclared alignment, so a paragraph the template left alone rendered justified instead of taking Word's default. | — | Fixed — returns `undefined`, and the library omits `w:jc` |
| G-7 | Header/footer parts were read one at a time in a serial loop, and each part was traversed three times. A template with 11 sections declares 10–20 such parts, and those round trips were most of `thesis parse`'s wall time. | — | Fixed — the relationships are read concurrently with `Promise.all` |
| G-8 | PR #2 (`codex/v3-advanced-omml`) had been open since 2026-08-31 with two commits of real work — text-based PDF template inference — based on pre-audit code, so it could no longer merge. | `git merge-base` = `1ada918e`; 9 conflicted files | Integrated in #8: replayed onto `main`, conflicts resolved by hand so nothing from #4/#7 was reverted, 3 PDF tests carried over. Adds `parsePdfTemplate` (`pdfjs-dist`), `thesis parse <template.docx\\|pdf>`, and a README scope note |

`pnpm check`: 145 tests, all passing.

---

## 8. Template-fidelity round

The three holes left in the "template-driven" promise, plus the journal front
matter that was still hardcoded.

| # | Finding | Evidence | Status |
|---|---|---|---|
| F-1 | Character units were converted with the style's **own** font size, before the `basedOn` chain was resolved. A style that inherits 10.5pt converted with the 12pt fallback. | A style `basedOn` a 10.5pt parent, `w:firstLineChars="200"` parsed to 480 twips; 420 is right. A style declaring 10.5pt itself got 420 | Fixed — the raw counts (`firstLineChars`, `leftChars`, `rightChars`, `beforeLines`, `afterLines`) are carried through the merge and converted in `rawToParagraphStyle`, where the inherited size is known. `mergeParagraphProps` also drops an inherited unit form when a style declares the twips form explicitly, so `override` still wins |
| F-2 | The template's `TOC1` / `TOC2` / `TOC3` styles were parsed and mapped to the `toc_entry` role, but the renderer never requested that role and the generated package carried none of those style ids — so the entries Word generates kept Word's own formatting. | `roleWinners.toc_entry = "TOC1"`; no `TOC1` in the emitted `styles.xml` | Fixed — `tocEntryStyles()` translates the template's three directory styles into the package's paragraph styles |
| F-3 | `w:docGrid` was neither parsed nor written. The writer library always emits `linePitch="360"`; the university template snaps the body to 312. | no `docGrid` anywhere in the parser or renderer | Fixed — `PageSettings.docGrid`, read from each `w:sectPr`, written back into the generated `document.xml` by the post-pass. The *shipped* asset still has no `docGrid`, because it was generated before this change and its source `.docx` is not in the repository; re-parsing it will fill the field in |
| F-4 | The journal front matter (title, authors, affiliations, abstract, keywords) was built from literal font, size and spacing values — seven `font: { ascii: … }` literals and zero `resolveStyle` calls — while the body already went through the template. | — | Fixed — five roles added to `StyleRole` (`journal_title`, `journal_author`, `journal_affiliation`, `journal_abstract`, `journal_keywords`) with fallbacks in the renderer, and every front-matter paragraph now resolves through `resolveStyle` |

`pnpm check`: 150 tests, all passing.

---

## 9. Code-quality round

The structural items left after §6–§8. No behaviour change to any document, but
the parser gets faster and its public surface stops advertising things it cannot
support.

| # | Finding | Evidence | Status |
|---|---|---|---|
| H-1 | `resolveStyleInheritance` walked the whole `basedOn` chain from scratch for every inheriting style — O(styles × depth) property-bag spreads, and the shipped template hangs almost every style off `Normal`. | — | Fixed — each style resolves once and its descendants reuse the result (`Map` cache, plus a path-local `resolving` set so a loop is still cut rather than mistaken for a hit) |
| H-2 | Five internal helpers were re-exported, but their signatures name `RawStyle` / `RawParagraphProps`, which are not exported — no external caller could write the parameter type, and nothing in the repository imported them. `countStyleUsage` also carried a public `string` signature that nothing called. | `dist/index.d.ts` declared `RawStyle` without exporting it | Fixed — the package exports `parseTemplate` and `parsePdfTemplate` only; `countStyleUsage`'s string entry point is gone (its tree-based form stays internal) |
| H-3 | The Chinese-font guard stripped **every** `w:*Theme` attribute and rewrote **every** Latin `w:eastAsia` to 仿宋, in styles.xml *and* document.xml. Harmless for a template that says nothing, hostile to one that says exactly what it wants. | — | Fixed — scoped to `w:docDefaults`, where the library bug actually is |
| H-4 | **The guard is inert with docx 9.7.1.** The emitted part is `<w:docDefaults><w:rPrDefault/><w:pPrDefault/></w:docDefaults>` — no `rFonts` at all, because the renderer never asks the library for a default document run. | inspected the emitted `styles.xml` | Recorded rather than removed — the symptom it addresses was real and a library bump could reintroduce it, and staying narrow costs nothing. The comment now says so instead of implying an active fix |
| H-5 | A table row with the wrong cell count was padded or truncated with no word to the caller. `validateDocument` reports it, but a caller using the renderer directly does not necessarily run the validator. | `styledRuns(row[ci] ?? '', bodyStyle)` | Fixed — the renderer warns which row, how many cells it has, and whether cells were added empty or dropped |

`pnpm check`: 154 tests, all passing.

---

## 10. Layout-fidelity round

The three items §9 left as "needs a content-model decision". Two of them now have
one, and the third got an explicit rule instead of a redesign.

| # | Finding | Decision | Status |
|---|---|---|---|
| L-1 | A cover gap could only be expressed as a line count. The guide fixes 5 cm between the cover tiers, and "five blank lines" is only as tall as the line height in force. | A spacer is a **height**: `lines`, `twips` or `cm`, exactly one. An exact height renders as one empty paragraph with `lineRule="exact"`, so 5 cm is 2834 twips regardless of typography. | Done — `SpacerBlock` widened, validated, rendered |
| L-2 | A heading's number could only come from the content file. The template binds its heading styles to `w:num` instances through `w:numPr`, and nothing read `numbering.xml`. | **A `number` in the content wins; the template's numbering applies otherwise.** The author who wrote an exact label meant it, and silently replacing it with an automatic one would be the surprise. | Done — `word/numbering.xml` is parsed into `DocumentTemplate.numbering`, each style carries its `w:numPr` binding, and a heading with a binding and no written number gets the template's reference and level. `w:numId="0"` ("no numbering") is honoured. Not read: `w:lvlOverride` |
| L-3 | A chapter that turns one passage landscape, or sets it in two columns, was flattened into whichever part it fell in. | A `section_break` block starts a new section and carries that section's page setup. Layout goes per-part whenever any part asks for its own, even if the template declares only one page setup. | Done — `SectionBreakBlock` with `orientation` / `columns` / `columnGutter`; the body is split at those markers, each chunk becomes its own section, and `PageSettings.orientation` now also reads a template's own `w:orient` |

Two details worth recording:

- The writer library swaps the page dimensions itself when told a page is
  landscape. Swapping them first and telling it too cancels out — the emitted
  section came out `w:w="11906" w:h="16838" w:orient="landscape"`, which is a
  portrait page with a landscape label. Orientation is now passed through and
  the swap is left to the library.
- `section_break` is honoured inside a thesis body. The cover, the front
  matter and the back matter are not split — they are already sections of their
  own — and a break there is reported by a warning rather than dropped in
  silence.

`pnpm check`: 161 tests, all passing.

---

## 11. Compatibility against real publisher templates

Up to this point every parser test built its own `.docx`. That proves the parser
handles what the tests imagine, not what a publisher ships. Four unmodified
templates were downloaded and kept as fixtures:

| Fixture | Source | Notes |
|---|---|---|
| `tests/fixtures/templates/acm-journal.docx` | [ACM interim Word template](https://www.acm.org/binaries/content/assets/publications/word_style/interim-template-style/interim-layout.docx) | 468 styles, 4 header/footer parts, `numbering.xml` |
| `tests/fixtures/templates/cn-cjstp.docx` | [中国科技期刊研究 论文模板](https://www.cjstp.cn/attached/file/20200402/20200402172721_69.docx) | 10 styles, **no heading style at all** |
| `tests/fixtures/templates/cn-dlmu-thesis.docx` | [《大连海事大学学报》模板](https://cbimg.cnki.net/Editor/2025/0311/dlhx/3cc38dd4-7f31-48b7-8bf2-3c361bbbcea4.docx) | 21 styles, body style is 硕士正文样式 |
| `tests/fixtures/templates/elsevier-hardwarex.docx` | [HardwareX (Elsevier) template](https://zenodo.org/record/4792320) | 15 styles, body style is `Standard` |

A 4.3 MB free-format journal template from Zenodo was also parsed during the
sweep but not committed, to keep the repository small.

### What the sweep found

| # | Finding | Evidence | Status |
|---|---|---|---|
| R-1 | The body-text pattern only knew the stock names (`Normal`, `Body`, `正文`). Real templates name it by hand. | `cn-dlmu-thesis.docx` calls its body style **硕士正文样式** and applies it to its body paragraphs; the parser settled for `Normal` (宋体 21 half-points) instead, missing the 12pt body the template actually uses. `elsevier-hardwarex.docx` calls it **`Standard`** and applies it 23 times. | Fixed — the pattern now covers 硕士/博士正文样式, 学位论文正文, `Standard`, `Body Text Indent` |
| R-2 | The heading patterns missed the names publishers use. | ACM and the Zenodo journal template name their headings `Head1` / `Heading1`; Chinese journals use 一级标题/二级标题. | Fixed — `一级标题`, `二级标题`, `三级标题`, `Head1..3`, `MainHeading` added |
| R-3 | **Some real templates declare no heading styles at all.** | `cn-cjstp.docx` has 10 styles: `Normal`, `footer`, `header`, `footnote text`, `HTML Preformatted`, `p0` — nothing for a heading. Journals routinely format headings directly. | **Not a defect.** The renderer falls back to the role defaults, which is the documented contract ("a template that carries no formatting still produces correct headings"). A test pins that such a template still parses and still finds its body style, rather than failing |
| R-4 | A template with hundreds of styles maps almost all of them to `paragraph`. | `acm-journal.docx`: 468 styles, ~300 of them land on `paragraph` because none of them match a name pattern or carry an outline level. | **Accepted.** `pickRoleWinners` prefers a style the document actually applies, so `Para` (used 6 times) loses to `Normal` (a stock style). Output is correct — headings get their own size, body gets the template font — but the role table is noisy. Recorded, not changed: narrowing it needs a "which style does the body *really* use" heuristic that would be guesswork |

### Verified behaviour

All four fixtures parse, and a document with a heading, a body paragraph, a
three-line table and an equation renders against each of them:

```
template                  heading2        body
acm-journal.docx          PMingLiU 26     PMingLiU 18
cn-dlmu-thesis.docx       黑体 28          宋体 24
elsevier-hardwarex.docx   黑体 28          Calibri 22
```

The table width comes out at the template's printable width (`10080` twips for
ACM's 216 × 279 mm page), not the old hardcoded 8504.

Eight local `.docx` files were also parsed as a cross-check — three OpenThesis
outputs and five 博导申请 forms — and every one parsed without throwing.

`pnpm check`: 166 tests, all passing.

---

## 12. Table and figure formatting, and caption cross-references

Two questions from the owner: does the engine read a template's **table and
figure formatting**, and do **表题/图题** follow a standard citation form?

### What was actually true before this round

| Area | Before | Now |
|---|---|---|
| Table borders | The template parser read **no** table XML at all. The renderer hardcoded one shape — 1.5 pt outer rules, 0.75 pt header rule, no verticals — so every template produced the same table. | The template's `w:tblBorders`, `w:jc`, `w:tblLayout`, `w:gridCol` widths, header `w:shd` fill and header run size are parsed into `TableFormat` and drive the output |
| Figure format | Nothing was read. Every figure was centred at a width clamped to the printable area. | `ImageFormat` records `w:jc`, `wp:inline` vs `wp:anchor`, and the first drawing's EMU extent |
| Caption format | 编号 + one space + 题名, centred, with the caption role's 段前/段后 — already correct | Unchanged in text, but the number now sits inside a **bookmark** |
| Cross-references | `如表3-1所示` was plain text. Insert a table, renumber, and the sentence is wrong. | `[[表3-1]]` / `[[图2-1]]` become a `REF` field pointing at that bookmark |

### What the four fixtures declare, and what comes out

```
template                 top-rule  grid   (source)
acm-journal.docx              12   no     no table in the document body
cn-cjstp.docx                  4   no     0.5 pt three-rule
cn-dlmu-thesis.docx           12   no     1.5 pt top/bottom
elsevier-hardwarex.docx        4   yes    full 0.5 pt grid
```

`acm-journal.docx` declares no table of its own, so it keeps the built-in shape —
which is the correct reading, not a miss.

### The cross-reference mechanism

A caption's number is wrapped in a `w:bookmarkStart` / `w:bookmarkEnd` pair; the
body's `[[表3-1]]` becomes a `w:fldSimple` with `w:instr=" REF _Ref表31 \h "`. Word
recomputes the field, so the sentence follows the caption.

Two details that matter:

- The field carries a **fallback result** in its run. A viewer that does not
  evaluate fields still shows `表3-1` rather than an empty gap.
- The bookmark name is derived from the label alone (`_Ref表31`), and the field
  derives it from the same function. The first attempt passed the kind separately
  and produced `_Ref表表31` on the caption against `_Ref31` in the field — a
  reference that silently points nowhere. The test asserts the two are **equal**,
  not merely present.

A manuscript that writes the number out longhand is left exactly as it was; only
the `[[…]]` form becomes a field. Both are pinned by tests.

### Two library details worth recording

`docx` 9.7.1 has no wrapper for `bookmarkStart` or `fldSimple`, so the renderer
constructs them through `XmlComponent`'s raw `root` convention: `['', { _attr },
text]`. Two traps:

- The attribute bag takes an object, but **a third slot of `''` emits
  `<w:t></w:t>`** — a text element with no text. An element that should be
  self-closing must pass no third slot at all.
- `XmlComponent` is a **value** import, not a type-only one, and its `root` is
  `protected` — a subclass has to `declare` it with the base class's exact type
  or TypeScript rejects the assignment.

`pnpm check`: 170 tests, all passing.

---

## 13. Table and figure series, part two

The audit for §12 asked a different question: of the fields that were *parsed*,
how many were actually *used*? Half of them were not.

| Field | Parsed in §12 | Used in §12 | Used now |
|---|---|---|---|
| `tableFormat.borders` | ✅ | ✅ | ✅ |
| `tableFormat.headerFill` / `headerSize` / `headerBold` | ✅ | ✅ | ✅ |
| `tableFormat.columnWidths` | ✅ | ❌ | ✅ — when the column count matches |
| `tableFormat.alignment` (`w:jc`) | ✅ | ❌ | ✅ |
| `tableFormat.layout` (`w:tblLayout`) | ✅ | ❌ | ✅ |
| `imageFormat.alignment` | ✅ | ❌ | ✅ |

Parsing a value and not using it is worse than not parsing it: the template JSON
claims the template was understood.

### What the four fixes changed

- **Column widths.** The template's grid applies only when the table has the
  same number of columns. Its 12-column example says nothing about a 2-column
  table, so an unmatched table still splits the printable width evenly.
- **Alignment and layout.** `w:jc` and `w:tblLayout` now reach `w:tblPr`, and the
  grid is emitted explicitly so `w:gridCol` stays in step with the cell widths.
- **Figure alignment.** A template that puts its figures left or right is honoured
  instead of every drawing being centred.

### Three more defects the audit turned up

| # | Finding | Evidence | Status |
|---|---|---|---|
| T-1 | A caption with no `number` rendered with no number at all, so it could never be the target of a reference. | a table with `caption: '无编号表'` and no `number` emitted no number | Fixed — captions are numbered per chapter in document order (表1-1, 图2-1), front matter outside a chapter numbers as 表1, and a chapter boundary resets the sequence |
| T-2 | A reference to a caption that does not exist emitted a `REF` field pointing at a bookmark that was never created. Word prints *"Error! Reference source not found."* and nothing said so. | `[[表9-9]]` with no such caption produced a field and no warning | Fixed — an unknown label warns and the reference is left as the literal text `[[表9-9]]` |
| T-3 | A missing image emitted only a placeholder line: no caption, no number, no bookmark, so `[[图2-1]]` could not reach it. | a `figure` with an unreadable path produced `[未找到图片: …]` and nothing else | Fixed — the placeholder is followed by the caption, number and bookmark, so the figure keeps its identity even when its pixels are gone |

Also added: `[[表#id]]` resolves through a block's `id` rather than its final
number. The author cannot know a table will become 表3-1, so naming it by id is
the form that survives automatic numbering.

`pnpm check`: 175 tests, all passing.
