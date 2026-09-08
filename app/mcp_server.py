"""
Utilities MCP Server (ADR-0097)

Side-effect-free catalog tools for Design / Work Item reshape:
encode_base64, decode_base64, json_stringify, json_parse.

Usage:
    python -m app.mcp_server
    uvicorn app.mcp_server:http_app --host 0.0.0.0 --port 8000
"""

from __future__ import annotations

import logging

from fastmcp import FastMCP
from starlette.applications import Starlette
from starlette.responses import JSONResponse
from starlette.routing import Mount, Route

from app.tools import register_all

logger = logging.getLogger(__name__)

mcp = FastMCP(
    name="utilities",
    instructions="""
    Weaver Utilities MCP provides mechanical reshape tools (no cloud I/O):
    - encode_base64: UTF-8 string or JSON → standard Base64 (for Drive/Office upload_file.content_base64)
    - decode_base64: Base64 → text or JSON
    - json_stringify / json_parse: deterministic JSON text ↔ value

    Prefer these over LLM-invented Base64. Typical chain:
    articlesText → encode_base64 → upload_file.content_base64
    """,
)

register_all(mcp)

_base_app = mcp.http_app()


async def health(request):
    return JSONResponse({"status": "ok", "service": "utilities-mcp"})


http_app = Starlette(
    routes=[
        Route("/health", health),
        Mount("/", _base_app),
    ],
    lifespan=getattr(_base_app, "lifespan", None),
)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    logger.info("Starting Utilities MCP Server (stdio mode)")
    mcp.run()
