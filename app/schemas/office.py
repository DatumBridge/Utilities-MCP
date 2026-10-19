"""Request and response models for in-memory Office document tools."""

from __future__ import annotations

from typing import Annotated, Any, Literal, Optional

from pydantic import BaseModel, Field, model_validator

from app.schemas.inputs import COMMA_LIST, INT_TEXT, JSON_LIST, LINE_LIST, Text, object_text


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


class _Item(BaseModel):
    """A list item may arrive as JSON object text."""

    @model_validator(mode="before")
    @classmethod
    def _parse_text(cls, value: Any) -> Any:
        return object_text(value)


class DocxParagraph(_Item):
    @model_validator(mode="before")
    @classmethod
    def _plain_text(cls, value: Any) -> Any:
        value = object_text(value)
        if isinstance(value, (str, int, float)) and not isinstance(value, bool):
            return {"text": str(value)}
        return value

    text: Text = Field(
        ...,
        description="Paragraph text. Unicode is kept.",
        json_schema_extra={"x-datumbridge-encoding": "plain"},
    )
    heading_level: Annotated[Literal[0, 1, 2, 3], INT_TEXT] = Field(
        default=0,
        description="0 is body text. 1, 2, and 3 are headings.",
    )


class DocxTable(_Item):
    headers: Annotated[list[Text], COMMA_LIST] = Field(
        default_factory=list,
        description="Header cells. Omit or pass [] when the table has no header row.",
    )
    rows: Annotated[list[list[Text]], JSON_LIST] = Field(
        default_factory=list,
        description="Body rows. Each row is an array of strings, no longer than the header.",
    )


class TextReplace(_Item):
    find: Text = Field(..., description="Non-empty text to find. Matched literally, not as a pattern.")
    replace: Text = Field(..., description="Replacement text. May be empty.")


class XlsxSheet(_Item):
    name: Text = Field(..., description="Sheet name, 1–31 characters.")
    headers: Annotated[Optional[list[Text]], COMMA_LIST] = Field(
        default=None,
        description="Optional first row. Each header is stored as text.",
    )
    rows: Annotated[Optional[list[list[Any]]], JSON_LIST] = Field(
        default=None,
        description="Data rows. Each cell is a string, number, boolean, or null.",
    )


class XlsxSetCell(_Item):
    sheet: Text = Field(..., description="Existing sheet name, or a sheet added in this call.")
    cell: str = Field(..., description="A1 reference such as B2.")
    value: Any = Field(
        default=None,
        description="String, number, boolean, or null. Null clears the cell. Strings are stored as text.",
    )


class XlsxAppendRows(_Item):
    sheet: Text = Field(..., description="Sheet that receives the rows.")
    rows: Annotated[list[list[Any]], JSON_LIST] = Field(..., description="Rows appended after the last used row.")


class PptxSlide(_Item):
    title: Text = Field(
        default="",
        description="Slide title. Required when bullets is empty.",
        json_schema_extra={"x-datumbridge-encoding": "plain"},
    )
    bullets: Annotated[list[Text], LINE_LIST] = Field(
        default_factory=list,
        description="Body bullets. Required when title is empty.",
        json_schema_extra={"x-datumbridge-encoding": "plain"},
    )
    notes: Optional[Text] = Field(
        default=None,
        description="Speaker notes.",
        json_schema_extra={"x-datumbridge-encoding": "plain"},
    )


class PptxSetTitle(_Item):
    index: Annotated[int, INT_TEXT] = Field(..., ge=0, description="Zero-based slide index.")
    title: Text = Field(
        ...,
        description="New title text.",
        json_schema_extra={"x-datumbridge-encoding": "plain"},
    )
