# json_parse

Parse JSON text into a value. Invalid JSON fails closed.

## Parameters

| Name | Required | Meaning |
|---|---|---|
| `text` | yes | JSON text. |

## Cases

### Parse an object

Input:

```json
{"text": "{\"title\": \"Notes\", \"count\": 2}"}
```

Output:

```json
{"success": true, "value": {"title": "Notes", "count": 2}}
```

### Parse an array

Input:

```json
{"text": "[1, 2, 3]"}
```

Output:

```json
{"success": true, "value": [1, 2, 3]}
```

### Invalid JSON

Input:

```json
{"text": "{not json}"}
```

Output:

```json
{"success": false, "error": {"error_code": "JSON_PARSE_ERROR", "error_message": "Invalid JSON: ...", "retryable": false}}
```
