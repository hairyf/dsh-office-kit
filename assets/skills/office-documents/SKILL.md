---
description: Create, read, edit, render, and verify Word documents (.docx) — build new files, revise existing ones, and check the rendered pages.
---

# Documents (.docx)

A `.docx` is a zip of XML parts: `word/document.xml` holds the body, with headers, footers, footnotes, comments, styles and numbering beside it. Two routes reach them: the helper CLI (`docx create`, `docx read`, `docx render`, no in-place mutation) and `scripts/*.py`, which edits everything structural — text, blocks, properties, comments, tracked changes, fields, protection, watermarks. Pick the cheapest route that yields the requested structure, then look at the pages. Never invent content; never deliver a file you have not looked at.

## Before any file work

Call `office_status`, then read the `## Office runtime` block at the end of this file — it holds the verified `python`, `cli`, `skill` and `shared` paths for this install. If there is no runtime, stop, give the user the `node "..."` remedy command, and call `office_status` again — never fake a document or claim success. Quote every path: spaces, non-ASCII names and `&` are normal on Windows.

## Delivery

Deliver with `present` plus a one-sentence summary. A bare path is not delivery: the transcript renders `.docx`/`.pptx`/`.xlsx`/`.pdf` links as "二进制文件，无法显示改动". Write new files into the workspace, verify, `present`. Never overwrite an original — write `<original>-edited.<ext>` (or the user's name), verify, `present`, and state both paths.

## Annotations

An `office-annotations` payload is `[{ page, text, requirement, region: { left, top, width, height } }]`: `region` normalised 0–1 against the rendered preview page, `requirement` free-form. Treat it as a change request, not editor state — find what the region covers, apply the smallest edit that satisfies `requirement`, re-render, confirm. Unmappable payload (empty region, no geometry): say what you could not locate; never guess. For `.docx`, `page` is a rendered PDF page, not a Word page: locate by `text` and its neighbouring paragraphs, use the region only for near-misses, and name the object you picked when a region lands on a table row, caption or header.

## Verify

```bash
"<python>" "<cli>" docx render --input "<file.docx>" --out-dir "<render-dir>"
```

Writes a PDF and one PNG per page and prints the paths. Open each PNG with `read_image`; check heading hierarchy, table width and cell wrapping, page breaks, header/footer, clipped or overflowing text, missing glyphs (dashes, CJK, symbols), equations, caption placement. Iterate, then deliver — "it parses" is not verification. Without a renderer (`office_status` reports `renderers.available: false`, with a `remedy`) ask the user to fix the install or fall back to `docx read` plus the audit scripts; say layout was not visually verified and never describe an unseen page. Re-render after every script edit: a package can be valid while its page is broken. Details: `tasks/verify_render.md`. Setup: `troubleshooting/libreoffice_headless.md`.

## Task index

| Job | Task doc | CLI or main tool |
|---|---|---|
| Create a document | `tasks/create_edit.md` | `docx create --spec`, `python-docx` snippet, `append_blocks.py` |
| Read and review a file | `tasks/read_review.md` | `docx read`, `docx_kit.py text`, `style_lint`, audits |
| Edit text, blocks or properties | `tasks/create_edit.md` | `replace_text.py`, `append_blocks.py`, `core_properties.py` |
| Comments | `tasks/comments_manage.md` | `comments_*.py` |
| Tracked changes | `tasks/clean_tracked_changes.md` | `accept_tracked_changes.py`, `add_tracked_replacements.py` |
| Fields and updates | `tasks/fields_update.md` | `fields_report.py`, `fields_materialize.py` |
| Table of contents | `tasks/toc_workflow.md` | `insert_toc.py`, `fields_*.py` |
| Captions and cross-references | `tasks/captions_crossrefs.md` | `captions_and_crossrefs.py`, `insert_ref_fields.py` |
| Navigation links | `tasks/navigation_internal_links.md` | `internal_nav.py` |
| Footnotes and endnotes | `tasks/footnotes_endnotes.md` | `insert_note.py`, `footnotes_report.py` |
| Headings and numbering | `tasks/headings_numbering.md` | `heading_audit.py` |
| Tables and geometry | `tasks/tables_spreadsheets.md` | `docx create` table block, `table_geometry.py`, `docx_table_to_csv.py`, `xlsx_to_docx_table.py` |
| Images and figures | `tasks/images_figures.md` | `images_audit.py` |
| Sections and page setup | `tasks/sections_layout.md` | `section_audit.py` |
| Styles | `tasks/style_lint_normalize.md` | `style_lint.py`, `style_normalize.py` |
| Templates | `tasks/templates_style_packs.md` | `apply_template_styles.py`, `references/template-elicitation.md` |
| Accessibility | `tasks/accessibility_a11y.md` | `a11y_audit.py` |
| Content controls | `tasks/forms_content_controls.md` | `content_controls.py` |
| Editing protection | `tasks/protection_restrict_editing.md` | `set_protection.py` |
| Watermarks | `tasks/watermarks_background.md` | `watermark_add.py`, `watermark_audit_remove.py` |
| Privacy and metadata | `tasks/privacy_scrub_metadata.md` | `privacy_scrub.py` |
| Redaction | `tasks/redaction_anonymization.md` | `redact_docx.py` |
| Merging documents | `tasks/multi_doc_merge.md` | `merge_docx_append.py` |
| Comparing versions | `tasks/compare_diff.md` | `docx render`, `render_and_diff.py` |
| Equations | `tasks/equations.md` | `docx create` `equation` block |
| Visual verification | `tasks/verify_render.md` | `docx render`, `render_docx.py` (skill root) |
| Fixtures and edge cases | `tasks/fixtures_edge_cases.md` | `make_fixtures.py` |

Where no verb exists, use the listed script or a short `python-docx`/`lxml` snippet — both run on the runtime `python` without a new dependency.

## Document defaults

- Letter 8.5×11 portrait (A4 when the user or template says so), 1in margins, body 11–12pt, single column.
- `docx create` accepts an optional top-level `"font": "Microsoft YaHei"` family string; omit it to retain Calibri. The selected family applies consistently to Western, East Asian and complex-script text through document defaults and styles used by generated blocks, with East Asian language `zh-CN`. Choose an installed font with the glyph coverage the content needs; the helper does not install or embed fonts. Native equations retain their math font.
- Black headings and body; colour only for deliberate emphasis, links and chart series.
- Tables: `#D9D9D9` borders, bold header row, `w:tblHeader` repeat, no vertical merges in wide tables.
- No callout boxes, drop shadows or coloured sidebars — structure comes from headings, lists and tables.
- Captions under figures and above tables, `Caption` style, staying with the object.
- Prefer native headings, lists, tables and OMML equations over hand-positioned runs: the next edit depends on that structure.

## Writing quality

Plain sentences, no filler. Checklist `references/writing-quality.md`; package parts `references/docx-structure.md`.

## Shared workflow

Cross-format habits and annotation payloads: `"<shared>/common-workflow.md"`, `"<shared>/annotation-recipes.md"`.
