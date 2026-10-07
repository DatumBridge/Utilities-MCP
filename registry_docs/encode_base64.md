# encode_base64

Encode UTF-8 text or a JSON value to standard Base64. Use this before a tool that asks for `content_base64`. Pass `content` or `json`, not both.

## Parameters

| Name | Required | Meaning |
|---|---|---|
| `content` | no | UTF-8 plaintext. Use this or `json`. |
| `json` | no | A JSON value. It is serialized with Unicode kept, then encoded. |
| `encoding` | no | Only `utf-8`. |

## Cases

### Encode plaintext

Input:

```json
{"content": "hello"}
```

Output:

```json
{"success": true, "content_base64": "aGVsbG8=", "byte_length": 5}
```

### Encode a JSON value

Input:

```json
{"json": {"title": "Notes", "count": 2}}
```

Output:

```json
{"success": true, "content_base64": "ewogICJ0aXRsZSI6ICJOb3RlcyIsCiAgImNvdW50IjogMgp9", "byte_length": 36}
```

### Both content and json

Input:

```json
{"content": "hello", "json": {"a": 1}}
```

Output:

```json
{"success": false, "error": {"error_code": "VALIDATION_ERROR", "error_message": "Provide either content or json, not both", "retryable": false}}
```

### Neither content nor json

Input:

```json
{}
```

Output:

```json
{"success": false, "error": {"error_code": "VALIDATION_ERROR", "error_message": "Either content (string) or json (object/array/scalar) is required", "retryable": false}}
```
