"""Office tool guides match the handler contract."""

from __future__ import annotations

import ast
from pathlib import Path

from app.capability_bind import registry_docs
from app.core.exceptions import DOCUMENTED_ERROR_CODES

ROOT = Path(__file__).resolve().parents[1]
DOCS = ROOT / "registry_docs"
OFFICE = ROOT / "app" / "tools" / "office_tools.py"

CREATE = {"create_docx", "create_xlsx", "create_pptx"}
UPDATE = {"update_docx", "update_xlsx", "update_pptx"}


def office_tool_names() -> list[str]:
    tree = ast.parse(OFFICE.read_text(encoding="utf-8"))
    names = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.FunctionDef):
            continue
        for dec in node.decorator_list:
            target = dec.func if isinstance(dec, ast.Call) else dec
            if isinstance(target, ast.Attribute) and target.attr == "tool":
                names.append(node.name)
                break
    return names


def test_office_guides_match_tools_and_error_codes():
    names = office_tool_names()
    assert names == [
        "create_docx",
        "update_docx",
        "create_xlsx",
        "update_xlsx",
        "create_pptx",
        "update_pptx",
    ]
    allowed = {code for code, _retryable in DOCUMENTED_ERROR_CODES}
    for name in names:
        path = DOCS / f"{name}.md"
        assert path.is_file()
        text = path.read_text(encoding="utf-8")
        assert registry_docs(name) == text.strip()
        assert f"# {name}" in text
        assert "| Name | Required | Type | Sample | Related tool | Meaning |" in text
        assert "## Error codes" in text
        assert "## Cases" in text
        assert "original_provider_error" in text
        assert "untrusted" in text.lower()
        assert "Do not invent a token" in text
        assert "-32602" in text
        assert "```json" in text
        assert "invalid_argument" not in text
        assert len(text) <= 24000
        assert "credentials_json" not in text.split("## Cases", 1)[1]
        for code in allowed:
            if name in CREATE and code == "INVALID_BASE64":
                continue
            assert f"`{code}`" in text
        if name in UPDATE:
            assert "Invalid Base64:" in text
