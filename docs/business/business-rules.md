# Business rules

1. Utilities MCP stays a pure tool server. Office tools return bytes. They do not upload, email, or store the file.
2. Create and update Word (`.docx`), Excel (`.xlsx`), and PowerPoint (`.pptx`) only. Legacy `.doc`, `.xls`, and `.ppt` are out of scope.
3. Empty create or an update with no edit fails with `VALIDATION_ERROR`.
4. A replacement that matches no text, and no other edit, fails with `replacement matched no text`.
5. Excel strings are literal text. This server does not evaluate formulas.
6. The returned `content_base64` is the input to a storage tool when the user wants the file saved: Google Drive `upload_file`, or Office 365 `create_file` with `confirm=true`. Office 365 simple upload is 4 MiB. This server allows 8 MiB.
7. File names are base names with the matching extension. Paths are rejected.
8. Packages larger than 8 MiB are rejected.
9. Error codes are only `VALIDATION_ERROR`, `INVALID_BASE64`, and `INVALID_OFFICE_FILE`. All are `retryable: false`. `original_provider_error` is null because there is no provider.
10. Deep Agent guides in `registry_docs/<tool>.md` are the argument contract. Do not invent tokens or Base64 samples.
