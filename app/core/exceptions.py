"""Error codes returned by Office document tools.

Codec tools keep their own messages in ``app.services``. Office tools emit only
the codes in ``DOCUMENTED_ERROR_CODES``. There is no remote provider, so
``original_provider_error`` is always null.
"""

from __future__ import annotations

DOCUMENTED_ERROR_CODES: tuple[tuple[str, bool], ...] = (
    ("VALIDATION_ERROR", False),
    ("INVALID_BASE64", False),
    ("INVALID_OFFICE_FILE", False),
)

_CODES = {code for code, _retryable in DOCUMENTED_ERROR_CODES}


class OfficeDocumentError(Exception):
    def __init__(self, code: str, message: str):
        if code not in _CODES:
            raise ValueError(f"undocumented office error code: {code}")
        self.code = code
        self.message = message
        super().__init__(message)
