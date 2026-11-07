# Create or edit a document

Produce a `.docx` that opens cleanly in Word with real heading styles, lists and tables — and change it later without destroying what is there.

## Create

Write a JSON spec next to the output, then run:

```bash
"<python>" "<cli>" docx create --spec "<spec.json>" --out "<out.docx>"
```

```json
{
  "title": "Regional revenue review",
  "subtitle": "FY26 Q2",
  "author": "Finance",
  "header": "Internal — do not circulate",
  "footer": "FY26 Q2 review",
  "page_size": "Letter",
  "orientation": "portrait",
  "blocks": [
    { "type": "heading", "level": 1, "text": "Summary" },
    { "type": "paragraph", "text": "Revenue grew 18% year over year." },
    { "type": "bullets", "items": ["Renewals at 92%", "42 new accounts"] },
    { "type": "table", "headers": ["Region", "Revenue"], "rows": [["EMEA", "4.1M"], ["APAC", "3.2M"]] },
    { "type": "page_break" }
  ]
}
```

The optional top-level `font` string sets the same family for Western, East Asian and complex-script text (for example, `"font": "Microsoft YaHei"`). The default is `Calibri`; choose an installed family with the glyph coverage your document needs. Body, headings, lists, tables, headers and footers share this family. Heading styles are prepared even when unused, so later `append_blocks.py` headings keep the same family. Headings and titles are black, and East Asian language metadata is `zh-CN`. Native equations keep their math font.

Block types and keys: `heading` (`level`, `text`), `paragraph` (`text`, `bold`, `italic`, `size`, `color`, `align`, `style`), `bullets` / `numbered` (`items`), `table` (`headers`, `rows`, optional `style`, default `Table Grid`), `image` (`path`, optional `width` in inches), `equation` (`latex`), `quote`, `page_break`, `spacer`. Page keys: `page_size` (`Letter` or `A4`), `orientation`, and `top_margin` / `bottom_margin` / `left_margin` / `right_margin` in inches.

The command prints counts for headings, paragraphs, list items, tables, images and equations. Check `equation_fallbacks`: a non-zero value means some LaTeX was emitted as styled text instead of OMML.

## Edit

The helper CLI never mutates a document: every edit reads the original, writes a copy, and refuses to overwrite its input. Without `--out` the copy lands next to the source as `<name>-edited.docx`.

Replace text — body, tables, and every header and footer:

```bash
"<python>" "<skill>/scripts/replace_text.py" --input "<original.docx>" --out "<original-edited.docx>" \
  --find "18%" --replace "21%" --find "Hardware flat" --replace "Hardware steady"
```

`--find` / `--replace` pair up in order; `--spec edits.json` takes `{"replace": [{"find": "...", "replace": "..."}]}` instead. The command prints a count per pair. Matches preserve existing run formatting, including unchanged text on either side of a replacement. Within a cross-run match, unchanged text keeps its original run; inserted or changed text inherits the formatting at the start of that changed span (an insertion at a run boundary uses the following run, or the preceding run at the paragraph end). `--require-all` exits 1 if a pair matched nothing — the check for a scripted edit.

Append blocks — the same block types `docx create` accepts:

```bash
"<python>" "<skill>/scripts/append_blocks.py" --input "<original.docx>" --out "<original-edited.docx>" --blocks "blocks.json"
```

```json
[{ "type": "heading", "level": 1, "text": "Appendix" }, { "type": "table", "headers": ["Region", "Revenue"], "rows": [["EMEA", "4.1M"]] }]
```

`--spec` takes a full document spec and appends its `blocks` array. The builder is the CLI's own, so anything that works in `docx create` works here, images (`path`, `width`) and equations (`latex`) included.

Set properties:

```bash
"<python>" "<skill>/scripts/core_properties.py" --input "<original.docx>" --out "<original-edited.docx" \
  --title "Regional revenue review (final)" --author "Finance" --keywords "revenue, FY26"
```

`--show` reports the current properties without writing; the same names work in `--spec core.json` under a `core` object. Each step writes a valid document, so a pipeline is append, then properties, and every stage is re-runnable and inspectable.

## When the CLI is not enough

| Need | Route |
|---|---|
| Edit inside a header, footer, footnote or comment | `docx_kit.py xml <file> <part>` to inspect, then the snippet below |
| Mark edits as tracked changes | `add_tracked_replacements.py` (`tasks/clean_tracked_changes.md`) |
| Header date, footer page numbers, first hyperlink | `docx_ooxml_patch.py` |
| Replace text across styled runs | `replace_text.py` above |

```bash
"<python>" - <<'PY'
from docx import Document
d = Document(r"<file.docx>")
# Structural edits use python-docx; text substitutions use replace_text.py above.
d.sections[0].header.paragraphs[0].add_run(" — Reviewed")
d.save(r"<file-edited.docx>")
PY
```

Plain word swaps can edit `w:t` through `docx_kit` directly; anything that changes paragraph structure goes through `python-docx` so the surrounding XML stays consistent.

## Verify

`docx read` to confirm text and counts, then render every page (`tasks/verify_render.md`). A spec that produces three pages instead of two means the table or margins need work, not that the renderer is wrong.

## Pitfalls

- Deleting and retyping a heading loses its outline level; edit the text and keep the style.
- `append_blocks.py` inserts before the final section properties, so a trailing landscape section or page setup keeps working.
- Replacing text in a field result changes only the cached value; the field recalculates on open in Word.
- Copying content between documents by hand drops styles and relationships — use `merge_docx_append.py`.
