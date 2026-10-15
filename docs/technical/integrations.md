# Integrations

Utilities MCP does not integrate with a vendor API.

Office bytes are consumed by:

- Google Drive MCP `upload_file.content_base64`
- Office 365 MCP `create_file` (`content_base64`, `confirm=true`). Simple upload is 4 MiB. This server allows 8 MiB, so a larger package must use Drive. There is no `upload_file` on the Office 365 server.

Those servers own OAuth. This server never receives their tokens.

`encode_base64` is still the reshape step for UTF-8 prose and JSON. Do not Base64-encode an Office package with `encode_base64`; `create_*` and `update_*` already return Base64.
