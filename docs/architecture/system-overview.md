# System overview — Utilities MCP

`utilities-mcp` is a side-effect-free MCP tool server. It reshapes text and builds Office Open XML packages in memory. It does not call a cloud API and does not write the caller's disk.

## What changed (2026-10-15)

| Area | Change | Why | Impacted components | Risks |
|------|--------|-----|---------------------|-------|
| Office tools | `create_docx`, `update_docx`, `create_xlsx`, `update_xlsx`, `create_pptx`, `update_pptx` return `content_base64` | Design and Work Item steps need a Word, Excel, or PowerPoint file before Drive `upload_file` or Office 365 `create_file` | Tool Registry after Setup; upload stages that consume `content_base64` | Large packages (cap 8 MiB; Office 365 simple upload is 4 MiB). Redeploy `utilities-mcp`, then Setup. |

Codec tools (`encode_base64`, `decode_base64`, `json_stringify`, `json_parse`) are unchanged.

## Dependencies

- FastMCP
- `python-docx`, `openpyxl`, `python-pptx` (in-process only)

## Risks

- An 8 MiB cap rejects larger decks and workbooks.
- Editing a file from another program rewrites the package. Word replacement flattens the matching paragraph into its first run. Excel macros are not kept. New slides use a blank layout plus Title and Body text boxes.
- Spreadsheet strings are text, so `=1+1` is not calculated.
