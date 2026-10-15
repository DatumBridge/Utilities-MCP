# Utilities MCP — Tool specification (v1)

Server id: `utilities` / registry host recommendation: `utilities-mcp-main`.  
Auth: none (pure functions).

## encode_base64

| Field | Type | Required | Notes |
|-------|------|----------|-------|
| content | string | one of | Plaintext UTF-8 (e.g. articlesText) |
| json | any | one of | Object/array/scalar; serialized with `ensure_ascii=False`, indent=2 |
| encoding | string | no | Default `utf-8` only |

**Output:** `{ success, content_base64, byte_length }` or `{ success:false, error }`.

## decode_base64

| Field | Type | Required | Notes |
|-------|------|----------|-------|
| content_base64 | string | yes | Standard Base64 |
| as_format | string | no | `text` (default) or `json` |
| encoding | string | no | `utf-8` only |

**Output:** `{ success, text?, json?, byte_length }` — fail closed on invalid Base64 / UTF-8 / JSON.

## json_stringify

| Field | Type | Required | Notes |
|-------|------|----------|-------|
| value | any | yes | |
| indent | int \| null | no | Default 2; null = compact |

## json_parse

| Field | Type | Required | Notes |
|-------|------|----------|-------|
| text | string | yes | |

## create_docx / update_docx / create_xlsx / update_xlsx / create_pptx / update_pptx

In-memory Office Open XML. Success: `{ success, content_base64, file_name, mime_type, byte_length, replacements }`.

| Tool | Required | Notes |
|------|----------|-------|
| create_docx | title, paragraphs, or tables | `.docx` |
| update_docx | content_base64 plus one edit | replacements, then append |
| create_xlsx | sheets | cell strings stored as text, not formulas |
| update_xlsx | content_base64 plus one edit | add sheets, set cells, append rows |
| create_pptx | slides | 16:9, Title and Body shapes |
| update_pptx | content_base64 plus one edit | replacements, titles, append slides |

Argument samples: `registry_docs/<tool>.md`. Limits: [configuration.md](configuration.md).

## Error codes

Codec: `VALIDATION_ERROR`, `UNSUPPORTED_ENCODING`, `INVALID_BASE64`, `UTF8_DECODE_ERROR`, `JSON_ENCODE_ERROR`, `JSON_PARSE_ERROR`.

Office: `VALIDATION_ERROR`, `INVALID_BASE64` (update only), `INVALID_OFFICE_FILE`. `retryable` is false. `original_provider_error` is null.
