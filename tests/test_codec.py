"""Unit tests for utilities codec (no FastMCP required)."""

from __future__ import annotations

import base64
import json

import pytest

from app.services import (
    CodecError,
    decode_base64,
    encode_base64,
    json_parse,
    json_stringify,
)


def test_encode_content_preserves_unicode():
    text = "Thiết kế nội thất\n\nArticle body"
    b64, n = encode_base64(content=text)
    assert n == len(text.encode("utf-8"))
    assert base64.b64decode(b64).decode("utf-8") == text


def test_encode_json_object_not_title_only():
    payload = {
        "title": "Short",
        "articlesText": "Full article with tiếng Việt",
        "marketingArticles": [{"title": "A", "body": "long body"}],
    }
    b64, _ = encode_base64(json_value=payload)
    decoded = base64.b64decode(b64).decode("utf-8")
    assert "Full article" in decoded
    assert "marketingArticles" in decoded
    assert json.loads(decoded)["articlesText"].startswith("Full")


def test_encode_rejects_both_inputs():
    with pytest.raises(CodecError) as ei:
        encode_base64(content="a", json_value={"x": 1})
    assert ei.value.code == "VALIDATION_ERROR"


def test_encode_requires_one_input():
    with pytest.raises(CodecError) as ei:
        encode_base64()
    assert ei.value.code == "VALIDATION_ERROR"


def test_decode_roundtrip():
    b64, _ = encode_base64(content="hello")
    text, parsed, n = decode_base64(b64, as_format="text")
    assert text == "hello"
    assert parsed is None
    assert n == 5


def test_decode_as_json():
    b64, _ = encode_base64(json_value={"a": 1})
    text, parsed, _ = decode_base64(b64, as_format="json")
    assert text is None
    assert parsed == {"a": 1}


def test_decode_invalid_base64():
    with pytest.raises(CodecError) as ei:
        decode_base64("not!!valid")
    assert ei.value.code == "INVALID_BASE64"


def test_json_stringify_parse_roundtrip():
    value = {"title": "Ngắn", "body": "dài"}
    text = json_stringify(value, indent=2)
    assert "Ngắn" in text
    assert json_parse(text) == value


def test_json_parse_fail_closed():
    with pytest.raises(CodecError) as ei:
        json_parse("{")
    assert ei.value.code == "JSON_PARSE_ERROR"
