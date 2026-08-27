# hermes-mcp

An [MCP](https://modelcontextprotocol.io) server for [Hermes](https://github.com/hashicorp-forge/hermes) — the HashiCorp/IBM document management system. Connects AI assistants (like [Bob](https://bob.ibm.com)) to your Hermes instance so you can search, browse, and retrieve documents directly in chat.

## Prerequisites

- [Node.js](https://nodejs.org) v18 or later
- Access to a running Hermes instance (SharePoint backend)
- A valid Hermes session cookie (see [Authentication](#authentication))

## Installation

```bash
git clone https://github.com/HarutuneDavisIBM/hermes-mcp.git
cd hermes-mcp
npm install
npm run build
```

The compiled server will be at `build/index.js`.

## Authentication

Hermes (SharePoint deployment) authenticates via an AWS Application Load Balancer session cookie named `AWSELBAuthSessionCookie-0`. This cookie is set automatically when you log in through your browser.

### How to get your session cookie

1. Open your Hermes instance (e.g. `https://hermes-sharepoint.hashicorp.services`) and log in
2. Open **DevTools** (`Cmd+Option+I` on Mac, `F12` on Windows)
3. Go to the **Network** tab and click any `/api/v2/` request
4. In **Request Headers**, find the `cookie` field
5. Copy the value of `AWSELBAuthSessionCookie-0=...` (just that one cookie, including the name)

> **Note:** This session cookie expires periodically (typically after a few hours). You will need to refresh it by repeating the steps above.

## Configuration

Add the server to your MCP client's config file. For Bob, this is `~/.bob/settings/mcp.json`:

```json
{
  "mcpServers": {
    "hermes": {
      "command": "node",
      "args": ["/absolute/path/to/hermes-mcp/build/index.js"],
      "env": {
        "HERMES_BASE_URL": "https://your-hermes-instance.example.com",
        "HERMES_COOKIE": "AWSELBAuthSessionCookie-0=<your-session-cookie-value>"
      }
    }
  }
}
```

Replace:
- `/absolute/path/to/hermes-mcp` with the actual path where you cloned this repo
- `https://your-hermes-instance.example.com` with your Hermes URL
- `<your-session-cookie-value>` with the cookie value from DevTools

### Verify the connection

After saving the config, ask your AI assistant to run `hermes_me`. A successful response looks like:

```json
{
  "id": "...",
  "email": "you@example.com",
  "name": "Your Name"
}
```

If you get an authentication error, your session cookie has likely expired — repeat the [steps above](#how-to-get-your-session-cookie) to get a fresh one.

## Available Tools

| Tool | Description |
|---|---|
| `hermes_search` | Full-text search across all published documents. Supports optional `doc_type` and `product` filters. Each result includes an `objectID` field — use that with `hermes_get_document`. |
| `hermes_get_document` | Fetch full metadata for a document by its **SharePoint `objectID`** (e.g. `01XOO7K4...`), found in `hermes_search` results. **Not** the human-readable doc number like `RFC-123`. |
| `hermes_list_drafts` | List draft documents. Optionally filter by owner email. |
| `hermes_list_document_types` | List all document types configured in the instance (RFC, PRD, FRD, ADR, etc.) with their schemas and custom fields. |
| `hermes_list_products` | List all products and areas configured in the instance. Useful to know valid filter values for `hermes_search`. |
| `hermes_list_projects` | List projects that group related documents together. |
| `hermes_me` | Get the currently authenticated user's profile. Use this to verify your session cookie is working. |

### Example prompts

Once connected, you can ask your AI assistant things like:

- *"Search Hermes for RFCs about Vault authentication"*
- *"Find all PRDs related to the HCP platform"*
- *"Search for the document HVS-022 and get its full metadata"*
- *"List all draft documents owned by me"*
- *"What products are in Hermes?"*
- *"Show me projects related to Terraform"*

## Environment Variables

| Variable | Required | Description |
|---|---|---|
| `HERMES_BASE_URL` | No | Base URL of your Hermes instance. Defaults to `https://hermes-sharepoint.hashicorp.services` |
| `HERMES_COOKIE` | Yes* | Full cookie string including name, e.g. `AWSELBAuthSessionCookie-0=...` |
| `HERMES_TOKEN` | Yes* | ALB OIDC JWT token via `x-amzn-oidc-data` header (alternative to cookie) |

\* At least one of `HERMES_COOKIE` or `HERMES_TOKEN` is required.

## Development

```bash
# Install dependencies
npm install

# Build (compiles TypeScript → build/index.js)
npm run build

# Rebuild after making changes to src/index.ts
npm run build
```

The source is a single file: [`src/index.ts`](src/index.ts).

## Refreshing Your Session Cookie

Session cookies expire. When tools start returning authentication errors, grab a fresh cookie:

1. Log in to your Hermes instance in the browser
2. Open DevTools → Network → any `/api/v2/` request → Request Headers → `cookie`
3. Copy the `AWSELBAuthSessionCookie-0=...` value
4. Update `HERMES_COOKIE` in your `mcp.json`

Your MCP client will pick up the new value on the next request (no restart needed for most clients).

## Contributing

Pull requests are welcome. For major changes, please open an issue first.

## License

MIT
