# Deployment

Image: `utilities-mcp`, HTTP port 8000, health `GET /health`.

Kubernetes publish name: `utilities-mcp-main` (see the platform pack `60-mcp-tools.yaml`). No vault inject.

After a tool change, redeploy this image and run Tool Registry Setup or publish so `x-datumbridge-capabilities` and Docs refresh. Then redeploy the agent if Deep should read the new guides.

Rollback: deploy the previous image. Workflows that call the Office tools fail closed at a missing tool.
