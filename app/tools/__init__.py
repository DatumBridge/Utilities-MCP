"""MCP tool registration for utilities-mcp."""

from app.tools import codec_tools


def register_all(mcp) -> None:
    codec_tools.register(mcp)
