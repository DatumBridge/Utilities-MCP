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

Do **not** ask an LLM to invent `contentBase64`.

## Run locally

```bash
cd mcp/utilities-mcp
python3.11 -m venv .venv && source .venv/bin/activate   # Python ≥3.10 for FastMCP
pip install -r requirements_mcp.txt -r requirements_dev.txt
uvicorn app.mcp_server:http_app --host 0.0.0.0 --port 8000
# health: curl -s localhost:8000/health
pytest -q
```

Codec unit tests need only `requirements_dev.txt` (no FastMCP).
stdio: `python -m app.mcp_server`

## Docker

```bash
docker build -t utilities-mcp .
docker run --rm -p 8000:8000 utilities-mcp
```

## Registry

Publish into Tool Registry / Integrations as **Utilities MCP** / host `utilities-mcp-main` (no vault inject — pure tools). See ADR-0097.

## Out of scope (v1)

File I/O, network, secrets, HTML→PDF, zip, domain “marketing → Drive” recipes.
