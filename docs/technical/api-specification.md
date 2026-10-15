# API specification

Transport: MCP Streamable HTTP on `POST /mcp`. Health: `GET /health` → `{"status":"ok","service":"utilities-mcp"}`.

No authentication on the tool payloads. The platform gateway may still front the server.

Codec tools (`encode_base64`, `decode_base64`, `json_stringify`, `json_parse`) are specified in [MCP_TOOL_SPECIFICATION.md](MCP_TOOL_SPECIFICATION.md).

## Office tools

All six tools return:

| Field | Type | Notes |
|-------|------|-------|
| success | boolean | |
| content_base64 | string | Present on success. Standard Base64. |
| file_name | string | Present on success. |
| mime_type | string | OOXML MIME for the format. |
| byte_length | integer | Decoded size. |
| replacements | integer or null | Set on Word and PowerPoint updates that passed `replacements`. |
| error | object | On failure: `error_code`, `error_message`, `retryable` (false), `original_provider_error` (null). |

| Tool | Required input | Edit order |
|------|----------------|------------|
| `create_docx` | title, paragraphs, or tables | title, paragraphs, tables |
| `update_docx` | `content_base64` plus one edit | replacements, then append |
| `create_xlsx` | `sheets` | one sheet per item |
| `update_xlsx` | `content_base64` plus one edit | add sheets, set cells, append rows |
| `create_pptx` | `slides` | one slide per item |
| `update_pptx` | `content_base64` plus one edit | replacements, titles, append slides |

Argument samples and error messages: `registry_docs/<tool>.md`.

## Error codes

`VALIDATION_ERROR`, `INVALID_BASE64` (update tools), `INVALID_OFFICE_FILE`.

Missing required arguments are JSON-RPC `-32602` before the handler runs.
