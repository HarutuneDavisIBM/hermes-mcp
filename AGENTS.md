# Agent Instructions

This file provides instructions for AI agents (Bob, Copilot, etc.) working with this MCP server.

## Authentication — the only method that works

This Hermes deployment authenticates **exclusively via the `AWSELBAuthSessionCookie-0` session cookie**.

- `HERMES_TOKEN` / `x-amzn-oidc-data` is **not used** by this deployment. Do not suggest it.
- The only required env var is `HERMES_COOKIE`.

The correct value format is the **full cookie name=value string**:

```
AWSELBAuthSessionCookie-0=<value>
```

Include the `AWSELBAuthSessionCookie-0=` prefix — this is passed verbatim as the HTTP `Cookie` header.

## When authentication fails

If any tool returns an HTML error page (`Unexpected token '<'...` or `Search failed: Unexpected token`) instead of JSON, the session cookie has expired. This is the **only** failure mode.

**Do not** attempt to use `HERMES_TOKEN`. **Do not** look for a `x-amzn-oidc-data` header. Just ask the user for a new cookie.

When asking for a new cookie, give these exact instructions:

> 1. Go to `https://hermes-sharepoint.hashicorp.services` in your browser and make sure you're logged in
> 2. Open DevTools (`Cmd+Option+I`) → **Application** tab → **Cookies** → select the site
> 3. Find the cookie named `AWSELBAuthSessionCookie-0`
> 4. Copy the **Name=Value** pair (e.g. `AWSELBAuthSessionCookie-0=abc123...`)
> 5. Paste it here

## Updating the cookie

1. Use `execute_command` with a `node` one-liner to update `~/.bob/settings/mcp.json` — this is safer than string search/replace because the cookie value is long and contains special characters that confuse diff tools:

```bash
node -e "
const fs = require('fs');
const path = '/Users/<username>/.bob/settings/mcp.json';
const config = JSON.parse(fs.readFileSync(path, 'utf8'));
config.mcpServers.hermes.env.HERMES_COOKIE = '<new-cookie-value>';
fs.writeFileSync(path, JSON.stringify(config, null, 2));
console.log('Done');
"
```

2. Kill the old server process so it restarts with the new env:

```bash
pkill -f "hermes-mcp/build/index.js"
```

3. Wait for the server to reconnect, then verify with `hermes_me`.

## Verifying the connection

Call `hermes_me`. A working session returns:
```json
{ "operation": "login" }
```

An expired/invalid cookie returns an HTML page, which surfaces as a JSON parse error.

## Cookie expiry

Cookies typically expire after a few hours. This is normal — just ask the user for a new one using the steps above.
