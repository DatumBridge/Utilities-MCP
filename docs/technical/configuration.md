# Configuration

No environment variables. Runtime dependencies are `requirements_mcp.txt` (`python-docx`, `openpyxl`, `python-pptx`, `lxml`).

Limits are constants in `app/services/office_documents.py`:

| Limit | Value |
|-------|-------|
| Decoded package | 8388608 bytes |
| Uncompressed ZIP | 33554432 bytes |
| ZIP entries on an uploaded package | 512 |
| Paragraphs / slides | 400 / 80 |
| Sheets / rows / columns | 20 / 5000 / 50 |
| Document text | 2000000 characters |

The 512-entry cap applies when reading `content_base64`. Eighty slides with speaker notes is about 360 parts, under that cap. Output size is also limited by the slide, sheet, and 8 MiB caps.

Changing a limit is a behavior change. Update `registry_docs` when a documented message includes the number.
