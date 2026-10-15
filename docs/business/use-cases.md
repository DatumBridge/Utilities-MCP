# Use cases

1. A Design stage writes a status note as a Word file and uploads it to Drive.
2. A Work Item builds a budget workbook (headers and rows) and uploads the `.xlsx`.
3. A follow-up stage appends a slide to a deck whose bytes came from `create_pptx` or from a download, then uploads the new `.pptx`.
4. The user asks to replace a phrase in an existing document. `update_docx` returns the edited package; a storage tool saves it.
