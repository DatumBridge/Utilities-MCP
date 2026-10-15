"""Request and response models for in-memory Office document tools."""

from __future__ import annotations

from typing import Any, Literal, Optional

from pydantic import BaseModel, Field


class OfficeErrorBody(BaseModel):
    error_code: str
    error_message: str
    retryable: bool = False
    original_provider_error: Optional[str] = None


class OfficeFileResponse(BaseModel):
    success: bool
    content_base64: Optional[str] = None
    file_name: Optional[str] = None
    mime_type: Optional[str] = None
    byte_length: Optional[int] = None
    replacements: Optional[int] = None
    error: Optional[OfficeErrorBody] = None


class DocxParagraph(BaseModel):
    text: str = Field(
        ...,
        description="Paragraph text. Unicode is kept.",
        json_schema_extra={"x-datumbridge-encoding": "plain"},
    )
    heading_level: Literal[0, 1, 2, 3] = Field(
        default=0,
        description="0 is body text. 1, 2, and 3 are headings.",
    )


class DocxTable(BaseModel):
    headers: list[str] = Field(
        default_factory=list,
        description="Header cells. Omit or pass [] when the table has no header row.",
    )
    rows: list[list[str]] = Field(
        default_factory=list,
        description="Body rows. Each row is an array of strings, no longer than the header.",
    )


class TextReplace(BaseModel):
    find: str = Field(..., description="Non-empty text to find. Matched literally, not as a pattern.")
    replace: str = Field(..., description="Replacement text. May be empty.")


class XlsxSheet(BaseModel):
    name: str = Field(..., description="Sheet name, 1–31 characters.")
    headers: Optional[list[str]] = Field(
        default=None,
        description="Optional first row. Each header is stored as text.",
    )
    rows: Optional[list[list[Any]]] = Field(
        default=None,
        description="Data rows. Each cell is a string, number, boolean, or null.",
    )


class XlsxSetCell(BaseModel):
    sheet: str = Field(..., description="Existing sheet name, or a sheet added in this call.")
    cell: str = Field(..., description="A1 reference such as B2.")
    value: Any = Field(
        default=None,
        description="String, number, boolean, or null. Null clears the cell. Strings are stored as text.",
    )


class XlsxAppendRows(BaseModel):
    sheet: str = Field(..., description="Sheet that receives the rows.")
    rows: list[list[Any]] = Field(..., description="Rows appended after the last used row.")


class PptxSlide(BaseModel):
    title: str = Field(
        default="",
        description="Slide title. Required when bullets is empty.",
        json_schema_extra={"x-datumbridge-encoding": "plain"},
    )
    bullets: list[str] = Field(
        default_factory=list,
        description="Body bullets. Required when title is empty.",
        json_schema_extra={"x-datumbridge-encoding": "plain"},
    )
    notes: Optional[str] = Field(
        default=None,
        description="Speaker notes.",
        json_schema_extra={"x-datumbridge-encoding": "plain"},
    )


class PptxSetTitle(BaseModel):
    index: int = Field(..., ge=0, description="Zero-based slide index.")
    title: str = Field(
        ...,
        description="New title text.",
        json_schema_extra={"x-datumbridge-encoding": "plain"},
    )
