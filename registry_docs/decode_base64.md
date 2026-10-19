# decode_base64

Decode Base64 to UTF-8 text or to a JSON value. Standard Base64, URL-safe Base64 (Gmail), a `data:...;base64,` URL, and missing padding are accepted. Invalid Base64, invalid UTF-8, and invalid JSON fail closed.

## Parameters

| Name | Required | Meaning |
|---|---|---|
| `content_base64` | yes | Base64 payload. Standard or URL-safe; a data URL or missing padding also works. |
| `as_format` | no | `text` or `json`. Catalog alias: `as`. Default `text`. |
| `encoding` | no | Only `utf-8`. |

## Cases

### Decode as text

Input:

```json
{"content_base64": "aGVsbG8=", "as_format": "text"}
```

Output:

```json
{"success": true, "text": "hello", "byte_length": 5}
```

### Decode as JSON

Input:

```json
{"content_base64": "eyJ0aXRsZSI6ICJOb3RlcyJ9", "as_format": "json"}
```

Output:

```json
{"success": true, "json": {"title": "Notes"}, "byte_length": 20}
```

### Invalid Base64

Input:

```json
{"content_base64": "!!!!"}
```

Output:

```json
{"success": false, "error": {"error_code": "INVALID_BASE64", "error_message": "Invalid Base64: ...", "retryable": false}}
```

### JSON mode with plaintext payload

Input:

```json
{"content_base64": "aGVsbG8=", "as_format": "json"}
```

Output:

```json
{"success": false, "error": {"error_code": "JSON_PARSE_ERROR", "error_message": "Decoded text is not JSON: ...", "retryable": false}}
```
