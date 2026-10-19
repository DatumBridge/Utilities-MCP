# Utilities MCP (ADR-0097)

Side-effect-free MCP server for Weaver Design / Work Item **reshape**:
Base64 encode/decode and JSON stringify/parse. No cloud credentials.

## Tools

| Tool | Description |
|------|-------------|
| `encode_base64` | UTF-8 `content` **or** `json` → `content_base64` + `byte_length` |
| `decode_base64` | `content_base64` → `text` or parsed `json` (`as_format`: `text` \| `json`) |
| `json_stringify` | Value → JSON text (`ensure_ascii=False`) |
| `json_parse` | JSON text → value (fail closed) |
| `create_docx` / `update_docx` | Word `.docx` in memory → `content_base64` |
| `create_xlsx` / `update_xlsx` | Excel `.xlsx` in memory → `content_base64` (strings are text, not formulas) |
| `create_pptx` / `update_pptx` | PowerPoint `.pptx` in memory → `content_base64` |

## Typical Design chain (Drive upload)

```text
stage-create-articles (llm)
  outputs: articlesText | marketingArticles
    → encode_base64
         content ← articlesText   (or json ← marketingArticles)
         outputs: content_base64
    → upload_file (Google Drive MCP)
         content_base64 ← encode_base64.output.content_base64
```

Do **not** ask an LLM to invent `contentBase64` or Office file bytes.

## Office file, then upload

```text
create_docx | create_xlsx | create_pptx
  outputs: content_base64, file_name, mime_type
    → Google Drive upload_file, or Office 365 create_file (confirm=true)
         content_base64 ← content_base64
         file_name ← file_name
```

Office 365 `create_file` is a simple upload capped at 4 MiB. This server allows 8 MiB, so a larger package must use Drive.

`update_*` takes `content_base64` from `create_*` or from a download. It does not write disk. See `docs/adr/ADR-0001-office-ooxml-in-memory.md`.

## Run locally

```bash
cd mcp/utilities-mcp
python3.11 -m venv .venv && source .venv/bin/activate   # Python ≥3.10 for FastMCP
pip install -r requirements_mcp.txt -r requirements_dev.txt
uvicorn app.mcp_server:http_app --host 0.0.0.0 --port 8000
# health: curl -s localhost:8000/health
pytest -q
```

Codec unit tests need only `requirements_dev.txt` (no FastMCP). Office tests need `requirements_mcp.txt` (`python-docx`, `openpyxl`, `python-pptx`, `lxml`).
stdio: `python -m app.mcp_server`

## Docker

```bash
docker build -t utilities-mcp .
docker run --rm -p 8000:8000 utilities-mcp
```

## Registry

Publish into Tool Registry / Integrations as **Utilities MCP** / host `utilities-mcp-main` (no vault inject — pure tools). See ADR-0097.

## Out of scope

Disk or network I/O, secrets, HTML→PDF, zip archives, legacy `.doc` / `.xls` / `.ppt`, and domain “marketing → Drive” recipes. Office Open XML create/update is in memory only (ADR-0001).
