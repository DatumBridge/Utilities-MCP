"""Register in-memory Word, Excel, and PowerPoint tools."""

from __future__ import annotations

import logging
from typing import Annotated, Optional

from pydantic import Field

from app.core.exceptions import OfficeDocumentError
from app.schemas.inputs import JSON_LIST, PARAGRAPH_LIST, Text
from app.schemas.office import (
    DocxParagraph,
    DocxTable,
    OfficeErrorBody,
    OfficeFileResponse,
    PptxSetTitle,
    PptxSlide,
    TextReplace,
    XlsxAppendRows,
    XlsxSetCell,
    XlsxSheet,
)
from app.services.office_documents import (
    create_docx as build_docx,
    create_pptx as build_pptx,
    create_xlsx as build_xlsx,
    update_docx as edit_docx,
    update_pptx as edit_pptx,
    update_xlsx as edit_xlsx,
)

logger = logging.getLogger(__name__)

_BASE64 = {"x-datumbridge-encoding": "base64"}
_PLAIN = {"x-datumbridge-encoding": "plain"}
_PLAIN_TEXT = {"x-datumbridge-encoding": "plain", "x-datumbridge-item-field": "text"}
_PLAIN_TITLE = {"x-datumbridge-encoding": "plain", "x-datumbridge-item-field": "title"}


def _fail(exc: OfficeDocumentError) -> OfficeFileResponse:
    logger.warning("office document failed: %s", exc.code)
    return OfficeFileResponse(
        success=False,
        error=OfficeErrorBody(
            error_code=exc.code,
            error_message=exc.message,
            retryable=False,
            original_provider_error=None,
        ),
    )


def _unexpected(kind: str) -> OfficeFileResponse:
    logger.warning("office document failed: %s", kind)
    return OfficeFileResponse(
        success=False,
        error=OfficeErrorBody(
            error_code="INVALID_OFFICE_FILE",
            error_message=f"Could not build {kind}",
            retryable=False,
            original_provider_error=None,
        ),
    )


def _ok(result) -> OfficeFileResponse:
    return OfficeFileResponse(
        success=True,
        content_base64=result.content_base64,
        file_name=result.file_name,
        mime_type=result.mime_type,
        byte_length=result.byte_length,
        replacements=result.replacements,
    )


def _dump(items):
    if not items:
        return None
    dumped = []
    for item in items:
        dumped.append(item.model_dump() if hasattr(item, "model_dump") else item)
    return dumped


def register(mcp) -> None:
    @mcp.tool()
    def create_docx(
        file_name: Optional[str] = Field(
            default=None,
            description="Base name ending with .docx. Default document.docx. Not a path.",
        ),
        title: Optional[Text] = Field(
            default=None,
            description="Document title, written as Heading 1.",
            json_schema_extra={"x-datumbridge-encoding": "plain"},
        ),
        paragraphs: Annotated[Optional[list[DocxParagraph]], PARAGRAPH_LIST] = Field(
            default=None,
            description="Body paragraphs or headings, in order, after the title.",
            json_schema_extra=_PLAIN_TEXT,
        ),
        tables: Annotated[Optional[list[DocxTable]], JSON_LIST] = Field(
            default=None,
            description="Tables appended after the paragraphs.",
            json_schema_extra=_PLAIN,
        ),
    ) -> OfficeFileResponse:
        """Create a .docx in memory and return standard Base64. Store it with Drive upload_file or Office 365 create_file (confirm=true, 4 MiB). No disk write.

        Capabilities: utilities.create_docx
Outputs: success, content_base64, file_name, mime_type, byte_length
        """
        try:
            return _ok(
                build_docx(
                    file_name=file_name,
                    title=title,
                    paragraphs=_dump(paragraphs),
                    tables=_dump(tables),
                )
            )
        except OfficeDocumentError as exc:
            return _fail(exc)
        except Exception:
            return _unexpected("Word document")

    @mcp.tool()
    def update_docx(
        content_base64: str = Field(
            ...,
            description="Existing .docx as standard Base64 from create_docx or a download.",
            json_schema_extra=_BASE64,
        ),
        file_name: Optional[str] = Field(
            default=None,
            description="Base name ending with .docx for the returned package. Default document.docx.",
        ),
        append_paragraphs: Annotated[Optional[list[DocxParagraph]], PARAGRAPH_LIST] = Field(
            default=None,
            description="Paragraphs appended after existing content.",
            json_schema_extra=_PLAIN_TEXT,
        ),
        append_tables: Annotated[Optional[list[DocxTable]], JSON_LIST] = Field(
            default=None,
            description="Tables appended after existing content.",
            json_schema_extra=_PLAIN,
        ),
        replacements: Annotated[Optional[list[TextReplace]], JSON_LIST] = Field(
            default=None,
            description="Literal find/replace on existing paragraphs, table cells, headers, and footers. Applied before appends.",
            json_schema_extra=_PLAIN,
        ),
    ) -> OfficeFileResponse:
        """Update a .docx in memory. Replacements run first, then paragraphs and tables are appended. Returns new Base64.

        Capabilities: utilities.update_docx
Outputs: success, content_base64, file_name, mime_type, byte_length, replacements
        """
        try:
            return _ok(
                edit_docx(
                    content_base64=content_base64,
                    file_name=file_name,
                    append_paragraphs=_dump(append_paragraphs),
                    append_tables=_dump(append_tables),
                    replacements=_dump(replacements),
                )
            )
        except OfficeDocumentError as exc:
            return _fail(exc)
        except Exception:
            return _unexpected("Word document")

    @mcp.tool()
    def create_xlsx(
        sheets: Annotated[list[XlsxSheet], JSON_LIST] = Field(
            ...,
            description="At least one worksheet. Strings are stored as text, including values that start with =.",
            json_schema_extra=_PLAIN,
        ),
        file_name: Optional[str] = Field(
            default=None,
            description="Base name ending with .xlsx. Default workbook.xlsx. Not a path.",
        ),
    ) -> OfficeFileResponse:
        """Create an .xlsx workbook in memory and return standard Base64. Store it with Drive upload_file or Office 365 create_file (confirm=true, 4 MiB). No disk write.

        Capabilities: utilities.create_xlsx
Outputs: success, content_base64, file_name, mime_type, byte_length
        """
        try:
            return _ok(build_xlsx(file_name=file_name, sheets=_dump(sheets)))
        except OfficeDocumentError as exc:
            return _fail(exc)
        except Exception:
            return _unexpected("Excel workbook")

    @mcp.tool()
    def update_xlsx(
        content_base64: str = Field(
            ...,
            description="Existing .xlsx as standard Base64 from create_xlsx or a download.",
            json_schema_extra=_BASE64,
        ),
        file_name: Optional[str] = Field(
            default=None,
            description="Base name ending with .xlsx (.xlsm for a macro-enabled workbook) for the returned package. Default workbook.xlsx or workbook.xlsm.",
        ),
        add_sheets: Annotated[Optional[list[XlsxSheet]], JSON_LIST] = Field(
            default=None,
            description="Sheets to add first, so later set_cells and append_rows can target them.",
            json_schema_extra=_PLAIN,
        ),
        set_cells: Annotated[Optional[list[XlsxSetCell]], JSON_LIST] = Field(
            default=None,
            description="Cells to write. Strings are stored as text, not formulas.",
            json_schema_extra=_PLAIN,
        ),
        append_rows: Annotated[Optional[list[XlsxAppendRows]], JSON_LIST] = Field(
            default=None,
            description="Rows appended after the last used row of the named sheet.",
            json_schema_extra=_PLAIN,
        ),
    ) -> OfficeFileResponse:
        """Update an .xlsx workbook in memory. Order: add sheets, set cells, append rows. Returns new Base64.

        Capabilities: utilities.update_xlsx
Outputs: success, content_base64, file_name, mime_type, byte_length
        """
        try:
            return _ok(
                edit_xlsx(
                    content_base64=content_base64,
                    file_name=file_name,
                    set_cells=_dump(set_cells),
                    append_rows=_dump(append_rows),
                    add_sheets=_dump(add_sheets),
                )
            )
        except OfficeDocumentError as exc:
            return _fail(exc)
        except Exception:
            return _unexpected("Excel workbook")

    @mcp.tool()
    def create_pptx(
        slides: Annotated[list[PptxSlide], JSON_LIST] = Field(
            ...,
            description="At least one slide. Each slide needs a title or at least one bullet.",
            json_schema_extra=_PLAIN,
        ),
        file_name: Optional[str] = Field(
            default=None,
            description="Base name ending with .pptx. Default presentation.pptx. Not a path.",
        ),
    ) -> OfficeFileResponse:
        """Create a .pptx presentation in memory and return standard Base64. Store it with Drive upload_file or Office 365 create_file (confirm=true, 4 MiB). No disk write.

        Capabilities: utilities.create_pptx
Outputs: success, content_base64, file_name, mime_type, byte_length
        """
        try:
            return _ok(build_pptx(file_name=file_name, slides=_dump(slides)))
        except OfficeDocumentError as exc:
            return _fail(exc)
        except Exception:
            return _unexpected("PowerPoint presentation")

    @mcp.tool()
    def update_pptx(
        content_base64: str = Field(
            ...,
            description="Existing .pptx as standard Base64 from create_pptx or a download.",
            json_schema_extra=_BASE64,
        ),
        file_name: Optional[str] = Field(
            default=None,
            description="Base name ending with .pptx for the returned package. Default presentation.pptx.",
        ),
        replacements: Annotated[Optional[list[TextReplace]], JSON_LIST] = Field(
            default=None,
            description="Literal find/replace on slide text, tables, and speaker notes. Applied first.",
            json_schema_extra=_PLAIN,
        ),
        set_slide_titles: Annotated[Optional[list[PptxSetTitle]], JSON_LIST] = Field(
            default=None,
            description="Replace a slide title by zero-based index. Applied after replacements.",
            json_schema_extra=_PLAIN_TITLE,
        ),
        append_slides: Annotated[Optional[list[PptxSlide]], JSON_LIST] = Field(
            default=None,
            description="Slides appended at the end. Applied last.",
            json_schema_extra=_PLAIN,
        ),
    ) -> OfficeFileResponse:
        """Update a .pptx presentation in memory. Order: replacements, titles, then appended slides. Returns new Base64.

        Capabilities: utilities.update_pptx
Outputs: success, content_base64, file_name, mime_type, byte_length, replacements
        """
        try:
            return _ok(
                edit_pptx(
                    content_base64=content_base64,
                    file_name=file_name,
                    append_slides=_dump(append_slides),
                    replacements=_dump(replacements),
                    set_slide_titles=_dump(set_slide_titles),
                )
            )
        except OfficeDocumentError as exc:
            return _fail(exc)
        except Exception:
            return _unexpected("PowerPoint presentation")
