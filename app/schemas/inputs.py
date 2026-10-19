"""Before-validators for tool arguments typed by hand.

Studio Run Test sends form fields as text: lists and objects arrive as JSON
text, numbers as digits, and empty fields as "". These validators turn that
text into the declared type so the published JSON schema stays unchanged.
"""

from __future__ import annotations

import json
from typing import Annotated, Any

from pydantic import BeforeValidator
from pydantic_core import PydanticUseDefault


def _blank(value: Any) -> bool:
    return isinstance(value, str) and value.strip() == ""


def _load(text: str) -> Any:
    try:
        return json.loads(text)
    except ValueError:
        return None


def blank_default(value: Any) -> Any:
    """Empty text means the field was left empty, so the default applies."""
    if _blank(value):
        raise PydanticUseDefault()
    return value


def json_list(value: Any) -> Any:
    """A list, one object, or JSON text of either."""
    if _blank(value):
        raise PydanticUseDefault()
    if isinstance(value, str):
        parsed = _load(value.strip())
        if not isinstance(parsed, (list, dict)):
            raise ValueError("Expected a JSON array, for example [{...}]")
        value = parsed
    if isinstance(value, dict):
        return [value]
    return value


def paragraph_list(value: Any) -> Any:
    """Like json_list, but plain text becomes one paragraph per non-empty line."""
    if isinstance(value, str) and not value.strip().startswith(("[", "{")):
        lines = [line.strip() for line in value.splitlines()]
        return [{"text": line} for line in lines if line] or None
    return json_list(value)


def _text_list(separator: str | None):
    def convert(value: Any) -> Any:
        if _blank(value):
            raise PydanticUseDefault()
        if not isinstance(value, str):
            return value
        text = value.strip()
        if text.startswith("["):
            parsed = _load(text)
            if isinstance(parsed, list):
                return parsed
        parts = text.split(separator) if separator else text.splitlines()
        return [part.strip() for part in parts if part.strip()]

    return convert


def scalar_text(value: Any) -> Any:
    """Numbers become text, so a cell such as 2 is stored as "2"."""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return str(value)
    return value


def int_text(value: Any) -> Any:
    """Digits sent as text, such as "2", become an integer."""
    if _blank(value):
        raise PydanticUseDefault()
    if isinstance(value, str) and value.strip().lstrip("-").isdigit():
        return int(value.strip())
    return value


def json_value(value: Any) -> Any:
    """JSON object or array text is parsed. Other text stays a string."""
    if _blank(value):
        raise PydanticUseDefault()
    if isinstance(value, str) and value.strip().startswith(("[", "{")):
        parsed = _load(value.strip())
        if isinstance(parsed, (list, dict)):
            return parsed
    return value


def json_text(value: Any) -> Any:
    """An object or array that is already parsed is turned back into JSON text."""
    if isinstance(value, (list, dict)):
        return json.dumps(value, ensure_ascii=False)
    return value


def object_text(value: Any) -> Any:
    """A list item sent as JSON object text is parsed."""
    if isinstance(value, str) and value.strip().startswith("{"):
        parsed = _load(value.strip())
        if isinstance(parsed, dict):
            return parsed
    return value


BLANK_DEFAULT = BeforeValidator(blank_default)
JSON_LIST = BeforeValidator(json_list)
PARAGRAPH_LIST = BeforeValidator(paragraph_list)
LINE_LIST = BeforeValidator(_text_list(None))
COMMA_LIST = BeforeValidator(_text_list(","))
INT_TEXT = BeforeValidator(int_text)
JSON_VALUE = BeforeValidator(json_value)
JSON_TEXT = BeforeValidator(json_text)

Text = Annotated[str, BeforeValidator(scalar_text)]
