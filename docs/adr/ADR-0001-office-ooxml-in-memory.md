# ADR-0001 — Office files are in-memory OOXML

## Context

Design and Work Item flows need a Word, Excel, or PowerPoint file before Drive or Office 365 upload. Asking a model to emit OOXML bytes is unreliable. Writing the file on the MCP host would couple the tool to a disk the user cannot see. ADR-0097 keeps this server free of cloud I/O.

## Decision

Add six tools: `create_docx`, `update_docx`, `create_xlsx`, `update_xlsx`, `create_pptx`, `update_pptx`. Each builds or edits an Open XML package in memory and returns `content_base64`, `file_name`, `mime_type`, and `byte_length`.

Excel strings are stored as text so a leading `=` is not a formula.

Updates take `content_base64` of the same package type. The server checks ZIP size and the expected OOXML part before parsing.

## Alternatives Considered

- Emit OOXML from the model. Rejected: truncation and corrupt packages.
- Write to a shared volume. Rejected: no durable caller-visible path, extra privilege.
- Add the builders to Drive and Office 365 MCPs. Rejected: every storage server would duplicate the same pure transform.
- Support legacy `.doc` / `.xls` / `.ppt`. Rejected: those are not ZIP packages and need another parser.

## Consequences

- Callers add an upload stage when the file should be stored.
- Image redeploy plus Tool Registry Setup is required before Deep can call the tools.
- Rollback is the previous image. Dependent workflows fail closed.

## Trade-offs

- Explicit stage instead of a hidden file write.
- New native libraries (`python-docx`, `openpyxl`, `python-pptx`, `lxml`) increase the image. They do not add network clients.

## Risks

- Edits are paragraph and cell text, not full layout fidelity for files produced by another editor. Word and PowerPoint replacement edits only the runs that hold the match; a match that spans runs takes the formatting of the run where it starts. Merged Word cells are edited once. Word replacement also reaches footnotes, endnotes, comments, and text boxes and shapes; the DrawingML copy and its VML fallback are both edited and counted once. `update_xlsx` does not load the workbook with openpyxl, which would drop shapes and VBA on save. It edits worksheet XML in place with lxml and copies every other part byte for byte, so shapes, images, charts, pivot tables, and macros are kept; a macro-enabled workbook is returned as `.xlsm`. A write to a table header renames the table column and rewrites structured references to it, as Excel does; without that Excel refuses to open the file. Checked in Excel 16 (build 20430), Word 16, and PowerPoint 16 on 2026-10-19 against files Excel and Word created. A sheet named `Sheet` is kept.
- The 8 MiB cap rejects large workbooks. Raising it needs a documented limit change.
