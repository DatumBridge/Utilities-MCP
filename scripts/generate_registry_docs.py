"""Keep registry_docs aligned with @mcp.tool names.

Guides are authored in registry_docs/. This script is the server-owned
generator so mcp/scripts/generate_registry_docs.py does not replace them
with generic stubs.
"""

from __future__ import annotations

import ast
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOCS = ROOT / "registry_docs"
TOOLS = ROOT / "app" / "tools"


def tool_names() -> list[str]:
    names: list[str] = []
    for path in sorted(TOOLS.glob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                continue
            for dec in node.decorator_list:
                target = dec.func if isinstance(dec, ast.Call) else dec
                if isinstance(target, ast.Attribute) and target.attr == "tool":
                    names.append(node.name)
                    break
    return names


def main() -> None:
    names = tool_names()
    if not names:
        raise SystemExit("no MCP tools found")
    missing = [name for name in names if not (DOCS / f"{name}.md").is_file()]
    if missing:
        raise SystemExit("missing registry docs: " + ", ".join(missing))
    print(f"utilities-mcp: {len(names)} guides")


if __name__ == "__main__":
    try:
        main()
    except SystemExit as exc:
        if exc.code not in (0, None):
            print(exc, file=sys.stderr)
        raise
