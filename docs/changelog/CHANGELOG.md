# Changelog

## 2026-10-19

### Fixed

- Tool arguments typed as text (Studio Run Test) no longer fail with Pydantic type errors. List and object fields accept JSON text, and a single object counts as a one-item list. Empty fields use their default. Numbers sent as text are accepted for `heading_level`, `index`, and `indent`. Numbers in text fields are stored as text.
- `paragraphs` and `append_paragraphs` accept plain text, one paragraph per line. A paragraph item can be a plain string. `bullets` accepts one bullet per line. `headers` accepts comma-separated text.
- `tables` and other list fields that get plain text now report `Expected a JSON array`.
- `decode_base64` accepts the catalog alias `as`.
- `encode_base64` ignores an empty `content` or `json`, so filling one and leaving the other blank no longer fails with "not both". JSON object or array text in `json` or in `json_stringify.value` is parsed first.
- `json_parse` accepts a value that is already parsed.
- `create_docx` and `update_docx` no longer let Word merge adjacent tables into one. An empty paragraph now separates a table from a table right before it.
- `update_docx` replacements no longer duplicate hyperlink text or drop formatting outside the match. They also reach first-page and even-page headers and footers, tables in headers, and nested tables. They no longer add empty header and footer parts.
- `update_pptx` replacements keep formatting outside the match. `set_slide_titles` writes the title placeholder of decks made in PowerPoint instead of the first text box, and adds a title box when a slide has none.
- `update_xlsx` counts existing sheets toward the 20-sheet limit.
- `update_docx`, `update_xlsx`, `update_pptx`, and `decode_base64` accept URL-safe Base64, a `data:` URL, and missing padding in `content_base64`.
- `update_docx` replacements reach text boxes and shapes, in the body and in headers and footers. The VML fallback copy is edited too and counted once.
- `update_docx` replacements reach footnotes, endnotes, and comments. A notes part without a match keeps its bytes.
- `update_xlsx` keeps tables valid: a header write renames the table column (empty becomes `ColumnN`, a repeated name gets a number) and rewrites structured references to it (`Table1[Qty]`) in cell formulas on every sheet, defined names, and table formulas, and a totals row write replaces that column's total function with the written label.
- `update_xlsx` no longer drops shapes, images, charts, form controls, or VBA. It edits worksheet XML in place instead of rebuilding the workbook with openpyxl. A macro-enabled workbook is returned as `.xlsm`. Writing over a formula removes calcChain.xml and asks Excel to recalculate on open. A shared formula is split so its other cells keep working. A cell inside an array formula is rejected.

## 2026-10-15

### Added

- `create_docx`, `update_docx`, `create_xlsx`, `update_xlsx`, `create_pptx`, `update_pptx`. In-memory Office Open XML returned as `content_base64` for Drive `upload_file` or Office 365 `create_file`. Excel strings are text, not formulas. Worksheet XML stays in memory. See ADR-0001.

### Changed

### Fixed

### Removed
