"""Registry docs are attached to tools/list from registry_docs/<tool>.md."""

from app.capability_bind import bind_declared_capabilities


class _Tool:
    def __init__(self) -> None:
        self.name = "encode_base64"
        self.description = "Capabilities: utilities.encode_base64"
        self.parameters = {}


class _Manager:
    def __init__(self, tool: _Tool) -> None:
        self._tools = {"encode_base64": tool}


class _MCP:
    def __init__(self, tool: _Tool) -> None:
        self._tool_manager = _Manager(tool)


def test_bind_loads_registry_docs() -> None:
    tool = _Tool()
    bind_declared_capabilities(_MCP(tool))
    docs = tool.parameters["x-datumbridge-docs"]
    assert "Encode plaintext" in docs
    assert "aGVsbG8=" in docs
    assert tool.parameters["x-datumbridge-capabilities"] == ["utilities.encode_base64"]
