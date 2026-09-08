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

## Error codes

`VALIDATION_ERROR`, `UNSUPPORTED_ENCODING`, `INVALID_BASE64`, `UTF8_DECODE_ERROR`, `JSON_ENCODE_ERROR`, `JSON_PARSE_ERROR`.
