# json_stringify

Serialize a value to JSON text. Unicode is kept.

## Parameters

| Name | Required | Meaning |
|---|---|---|
| `value` | yes | Object, array, or scalar. |
| `indent` | no | Spaces to indent. Null or omit for compact JSON. Default 2. |

## Cases

### Pretty JSON

Input:

```json
{"value": {"title": "Notes", "count": 2}, "indent": 2}
```

Output:

```json
{"success": true, "text": "{\n  \"title\": \"Notes\",\n  \"count\": 2\n}"}
```

### Compact JSON

Input:

```json
{"value": ["a", "b"], "indent": null}
```

Output:

```json
{"success": true, "text": "[\"a\",\"b\"]"}
```

### Missing value

`value` is required. The server rejects the call before it runs.

Input:

```json
{}
```
