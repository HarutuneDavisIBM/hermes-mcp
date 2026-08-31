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
3. Go to the **Application** tab → **Cookies** → select the site
4. Find the cookie named `AWSELBAuthSessionCookie-0`
5. Copy the **Name=Value** pair: `AWSELBAuthSessionCookie-0=<value>`

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
- `AWSELBAuthSessionCookie-0=<your-session-cookie-value>` with the full Name=Value pair from DevTools

### Verify the connection

After saving the config, restart the MCP server and ask your AI assistant to run `hermes_me`. A successful response looks like:

```json
{ "operation": "login" }
```

If you get an authentication error (`Unexpected token '<'` or similar), your session cookie has expired — repeat the [steps above](#how-to-get-your-session-cookie) to get a fresh one.

## Refreshing Your Session Cookie

Session cookies expire after a few hours. When tools start returning authentication errors:

1. Log in to your Hermes instance in the browser
2. Open DevTools → **Application** tab → **Cookies** → select the site
3. Find `AWSELBAuthSessionCookie-0` and copy the full `Name=Value` string
4. Update `HERMES_COOKIE` in your `mcp.json` using a safe JSON edit (not string replace — the value contains special characters):

```bash
node -e "
const fs = require('fs');
const p = require('os').homedir() + '/.bob/settings/mcp.json';
const c = JSON.parse(fs.readFileSync(p, 'utf8'));
c.mcpServers.hermes.env.HERMES_COOKIE = 'AWSELBAuthSessionCookie-0=<new-value>';
fs.writeFileSync(p, JSON.stringify(c, null, 2));
"
```

5. Kill the old process so it restarts with the new cookie:

```bash
pkill -f "hermes-mcp/build/index.js"
```

## Available Tools

### Read operations

| Tool | Description |
|---|---|
| `hermes_search` | Full-text search across all published documents. Supports optional `doc_type` and `product` filters. Each result includes an `objectID` field — use that with `hermes_get_document`. |
| `hermes_get_document` | Fetch full metadata for a document by its **SharePoint `objectID`** (e.g. `01XOO7K4...`), found in `hermes_search` results. **Not** the human-readable doc number like `RFC-123`. |
| `hermes_list_drafts` | List draft documents. Optionally filter by owner email. |
| `hermes_list_document_types` | List all document types configured in the instance (RFC, PRD, FRD, ADR, etc.) with their schemas and custom fields. |
| `hermes_list_products` | List all products and areas configured in the instance. Useful to know valid filter values for `hermes_search`. |
| `hermes_list_projects` | List projects that group related documents together. |
| `hermes_me` | Get the currently authenticated user's profile. Use this to verify your session cookie is working. |

### Write operations

| Tool | Description |
|---|---|
| `hermes_create_draft` | Create a new document draft (WIP status). Requires `title` and `doc_type`. Returns the new document's SharePoint objectID. |
| `hermes_update_draft` | Update a draft's metadata — title, summary, product, contributors, approvers, approver groups, owners, and custom fields. Only provided fields are changed. |
| `hermes_update_document` | Update a **published** document's metadata or status. Supported statuses: `In-Review`, `Approved`, `Obsolete`. Also updates title, summary, owners, contributors, approvers, approver groups, and custom fields. |
| `hermes_request_review` | Publish a draft and move it to `In-Review` status, notifying assigned approvers. The draft must have at least one approver set before calling this. |
| `hermes_approve_document` | Approve a document as the currently authenticated user. The document must be `In-Review` or `Approved` and the current user must be listed as an approver. |

### Typical workflow

1. **Create** — `hermes_create_draft` → get a document ID
2. **Edit** — `hermes_update_draft` (add title, summary, product, approvers, custom fields)
3. **Publish** — `hermes_request_review` (moves to `In-Review`, emails approvers)
4. **Approve** — `hermes_approve_document` (called by each approver)
5. **Mark obsolete** — `hermes_update_document` with `status: "Obsolete"`

### Example prompts

Once connected, you can ask your AI assistant things like:

- *"Search Hermes for RFCs about Vault authentication"*
- *"Find all PRDs related to the HCP platform"*
- *"Search for the document HVS-022 and get its full metadata"*
- *"List all draft documents owned by me"*
- *"What products are in Hermes?"*
- *"Show me projects related to Terraform"*
- *"Create a new RFC draft titled 'Improving Agent Authentication' for the Vault product"*
- *"Add jane@example.com as an approver on my draft RFC-??? and request review"*
- *"Mark document 01XOO7K4... as Obsolete"*
- *"Approve the document 01XOO7K4..."*

## Environment Variables

| Variable | Required | Description |
|---|---|---|
| `HERMES_BASE_URL` | No | Base URL of your Hermes instance. Defaults to `https://hermes-sharepoint.hashicorp.services` |
| `HERMES_COOKIE` | **Yes** | Full cookie Name=Value string: `AWSELBAuthSessionCookie-0=<value>` |

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

## Contributing

Pull requests are welcome. For major changes, please open an issue first.

## License

MIT
