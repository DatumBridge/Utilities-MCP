"""Pure codec helpers — no I/O (ADR-0097)."""

from __future__ import annotations

import base64
import json
from typing import Any


class CodecError(Exception):
    def __init__(self, code: str, message: str):
        self.code = code
        self.message = message
        super().__init__(message)


def _normalize_encoding(encoding: str | None) -> str:
    enc = (encoding or "utf-8").strip().lower()
    if enc in ("utf8", "utf-8"):
        return "utf-8"
    raise CodecError("UNSUPPORTED_ENCODING", f"Unsupported encoding: {encoding!r} (only utf-8)")


def plaintext_from_inputs(
    content: str | None,
    json_value: Any | None,
    *,
    indent: int | None = 2,
) -> str:
    """Build UTF-8 text from exactly one of content or json_value."""
    content_provided = content is not None
    json_provided = json_value is not None

    if content_provided and json_provided:
        raise CodecError(
            "VALIDATION_ERROR",
            "Provide either content or json, not both",
        )
    if not content_provided and not json_provided:
        raise CodecError(
            "VALIDATION_ERROR",
            "Either content (string) or json (object/array/scalar) is required",
        )
    if content_provided:
        return content if content is not None else ""
    try:
        if indent is None:
            return json.dumps(json_value, ensure_ascii=False, separators=(",", ":"))
        return json.dumps(json_value, ensure_ascii=False, indent=indent)
    except (TypeError, ValueError) as exc:
        raise CodecError("JSON_ENCODE_ERROR", f"Cannot serialize json: {exc}") from exc


def encode_base64(
    content: str | None = None,
    json_value: Any | None = None,
    encoding: str | None = "utf-8",
) -> tuple[str, int]:
    enc = _normalize_encoding(encoding)
    text = plaintext_from_inputs(content, json_value)
    raw = text.encode(enc)
    return base64.b64encode(raw).decode("ascii"), len(raw)


def decode_base64(
    content_base64: str,
    *,
    as_format: str = "text",
    encoding: str | None = "utf-8",
) -> tuple[str | None, Any | None, int]:
    enc = _normalize_encoding(encoding)
    fmt = (as_format or "text").strip().lower()
    if fmt not in ("text", "json"):
        raise CodecError("VALIDATION_ERROR", "as must be 'text' or 'json'")
    s = "".join((content_base64 or "").split())
    if not s:
        raise CodecError("VALIDATION_ERROR", "content_base64 is required")
    try:
        raw = base64.b64decode(s, validate=True)
    except Exception as exc:
        raise CodecError("INVALID_BASE64", f"Invalid Base64: {exc}") from exc
    try:
        text = raw.decode(enc)
    except UnicodeDecodeError as exc:
        raise CodecError("UTF8_DECODE_ERROR", f"Decoded bytes are not valid {enc}: {exc}") from exc
    if fmt == "text":
        return text, None, len(raw)
    try:
        return None, json.loads(text), len(raw)
    except json.JSONDecodeError as exc:
        raise CodecError("JSON_PARSE_ERROR", f"Decoded text is not JSON: {exc}") from exc


def json_stringify(value: Any, indent: int | None = 2) -> str:
    try:
        if indent is None:
            return json.dumps(value, ensure_ascii=False, separators=(",", ":"))
        return json.dumps(value, ensure_ascii=False, indent=indent)
    except (TypeError, ValueError) as exc:
        raise CodecError("JSON_ENCODE_ERROR", f"Cannot serialize value: {exc}") from exc


def json_parse(text: str) -> Any:
    if text is None or str(text).strip() == "":
        raise CodecError("VALIDATION_ERROR", "text is required")
    try:
        return json.loads(text)
    except json.JSONDecodeError as exc:
        raise CodecError("JSON_PARSE_ERROR", f"Invalid JSON: {exc}") from exc
