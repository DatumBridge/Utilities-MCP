# Security architecture

## Boundaries

- No credentials, tokens, or vault fields.
- No network calls from the Office tools.
- No path writes. Word and PowerPoint save to a memory buffer. `create_xlsx` writes worksheet XML with openpyxl's stream writer into a `BytesIO`, then into the ZIP buffer; the library's temp-file writer is not used. `update_xlsx` parses package XML with lxml (entities and network access off) and writes the new ZIP to a `BytesIO`.
- `file_name` is a base name. `/`, `\`, and `..` are rejected.

## Untrusted bytes

`content_base64` on update is data, not instructions. Before a library parses it, the server checks:

- standard Base64
- decoded size at most 8 MiB (8388608 bytes)
- ZIP local header
- at most 512 entries on an uploaded package (80 slides with notes is about 360 parts)
- uncompressed size at most 32 MiB
- the OOXML part for that tool (`word/document.xml`, `xl/workbook.xml`, or `ppt/presentation.xml`)
- entry names are not absolute and do not contain `..`

A wrong package type fails with `INVALID_OFFICE_FILE`.

## Spreadsheets

String cells are stored as text (`data_type` `s`, number format `@`). A value that starts with `=` is not an Excel formula when the file is opened.

## Logs

Failures log the error code only. Document text and file bytes are not logged.

## Residual

The Office libraries parse XML after those caps. The caps bound zip-bomb expansion declared in the ZIP directory. They do not add a separate XML sandbox.
