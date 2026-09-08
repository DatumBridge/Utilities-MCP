"""Register encode/decode and JSON reshape tools (ADR-0097)."""

from __future__ import annotations

import logging
from typing import Any, Optional

from pydantic import Field

from app.schemas import (
    DecodeBase64Response,
    EncodeBase64Response,
    ErrorBody,
    JsonParseResponse,
    JsonStringifyResponse,
)
from app.services import CodecError, decode_base64 as decode_b64
from app.services import encode_base64 as encode_b64
from app.services import json_parse as parse_json
from app.services import json_stringify as stringify_json

logger = logging.getLogger(__name__)


def _err(exc: CodecError) -> ErrorBody:
    return ErrorBody(error_code=exc.code, error_message=exc.message, retryable=False)


def register(mcp) -> None:
    @mcp.tool()
    def encode_base64(
        content: Optional[str] = Field(
            default=None,
            description="UTF-8 plaintext to encode (e.g. articlesText). Provide content OR json, not both.",
        ),
        json: Optional[Any] = Field(
            default=None,
            description="JSON value (object/array/scalar) to serialize with ensure_ascii=False then encode.",
        ),
        encoding: str = Field(
            default="utf-8",
            description="Text encoding before Base64 (only utf-8 supported).",
        ),
    ) -> EncodeBase64Response:
        """
        Encode UTF-8 string or JSON to standard Base64 for upload_file.content_base64.
        Prefer this tool over inventing Base64 in an LLM stage.
        """
        try:
            b64, n = encode_b64(content=content, json_value=json, encoding=encoding)
            return EncodeBase64Response(success=True, content_base64=b64, byte_length=n)
        except CodecError as e:
            logger.warning("encode_base64 failed: %s", e.message)
            return EncodeBase64Response(success=False, error=_err(e))

    @mcp.tool()
    def decode_base64(
        content_base64: str = Field(..., description="Standard Base64 payload"),
        as_format: str = Field(
            default="text",
            description="Decode as 'text' (UTF-8 string) or 'json' (parse after UTF-8 decode). Catalog alias: as.",
        ),
        encoding: str = Field(default="utf-8", description="Only utf-8 supported"),
    ) -> DecodeBase64Response:
        """Decode Base64 to UTF-8 text or JSON. Fail closed on invalid Base64 / UTF-8 / JSON."""
        try:
            text, parsed, n = decode_b64(
                content_base64, as_format=as_format, encoding=encoding
            )
            return DecodeBase64Response(
                success=True, text=text, json=parsed, byte_length=n
            )
        except CodecError as e:
            logger.warning("decode_base64 failed: %s", e.message)
            return DecodeBase64Response(success=False, error=_err(e))

    @mcp.tool()
    def json_stringify(
        value: Any = Field(..., description="Value to serialize (object/array/scalar)"),
        indent: Optional[int] = Field(
            default=2,
            description="Indent spaces; null for compact JSON",
        ),
    ) -> JsonStringifyResponse:
        """Serialize a value to JSON text (ensure_ascii=False)."""
        try:
            return JsonStringifyResponse(
                success=True, text=stringify_json(value, indent=indent)
            )
        except CodecError as e:
            return JsonStringifyResponse(success=False, error=_err(e))

    @mcp.tool()
    def json_parse(
        text: str = Field(..., description="JSON text to parse"),
    ) -> JsonParseResponse:
        """Parse JSON text into a value. Fail closed on invalid JSON."""
        try:
            return JsonParseResponse(success=True, value=parse_json(text))
        except CodecError as e:
            return JsonParseResponse(success=False, error=_err(e))
