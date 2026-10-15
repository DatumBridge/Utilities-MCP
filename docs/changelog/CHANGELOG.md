# Changelog

## 2026-10-15

### Added

- `create_docx`, `update_docx`, `create_xlsx`, `update_xlsx`, `create_pptx`, `update_pptx`. In-memory Office Open XML returned as `content_base64` for Drive `upload_file` or Office 365 `create_file`. Excel strings are text, not formulas. Worksheet XML stays in memory. See ADR-0001.

### Changed

### Fixed

### Removed
