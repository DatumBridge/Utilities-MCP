"""Base64 spellings that callers pass for the same bytes."""

from __future__ import annotations

_URL_SAFE = str.maketrans("-_", "+/")


def standard_base64(text: str) -> str:
    """Accept a data URL, the URL-safe alphabet (Gmail), or missing padding.

    Whitespace must already be removed. The result still needs strict decoding,
    so text that is not Base64 is rejected there.
    """
    if text.startswith("data:") and ";base64," in text:
        text = text.split(";base64,", 1)[1]
    text = text.translate(_URL_SAFE)
    return text + "=" * (-len(text) % 4)
