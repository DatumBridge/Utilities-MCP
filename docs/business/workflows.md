# Workflows

## New Office file, then upload

```text
create_docx | create_xlsx | create_pptx
  outputs: content_base64, file_name, mime_type
    → Google Drive upload_file
         content_base64 ← content_base64
         file_name ← file_name
    or Office 365 create_file (confirm=true, max 4 MiB)
         content_base64 ← content_base64
         file_name ← file_name
```

`encode_base64` remains the path for prose or JSON that is not an Office package.

## Edit, then upload

```text
content_base64 from create_* or from a prior download
    → update_docx | update_xlsx | update_pptx
         outputs: content_base64, file_name
    → Google Drive upload_file, or Office 365 create_file (confirm=true, max 4 MiB)
```

This server allows 8 MiB. A package larger than 4 MiB cannot use the Office 365 simple upload.

Update order:

- Word: replacements, then appended paragraphs, then appended tables.
- Excel: add sheets, then set cells, then append rows.
- PowerPoint: replacements, then titles, then appended slides.
