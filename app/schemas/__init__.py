"""Response schemas for utilities MCP tools."""

from __future__ import annotations

from typing import Any, Optional

from pydantic import BaseModel, Field


class ErrorBody(BaseModel):
    error_code: str
    error_message: str
    retryable: bool = False


class EncodeBase64Response(BaseModel):
    success: bool
    content_base64: Optional[str] = None
    byte_length: Optional[int] = None
    error: Optional[ErrorBody] = None


class DecodeBase64Response(BaseModel):
    success: bool
    text: Optional[str] = None
    json: Optional[Any] = Field(default=None, description="Parsed JSON when as=json")
    byte_length: Optional[int] = None
    error: Optional[ErrorBody] = None


class JsonStringifyResponse(BaseModel):
    success: bool
    text: Optional[str] = None
    error: Optional[ErrorBody] = None


class JsonParseResponse(BaseModel):
    success: bool
    value: Optional[Any] = None
    error: Optional[ErrorBody] = None
