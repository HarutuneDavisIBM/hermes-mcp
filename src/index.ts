#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";

const HERMES_BASE_URL =
  process.env.HERMES_BASE_URL ||
  "https://hermes-sharepoint.hashicorp.services";

// Hermes (SharePoint deployment) is fronted by an AWS ALB with OIDC.
// The ALB forwards the authenticated user's JWT in the "x-amzn-oidc-data"
// header. We pass that token along with every API request so Hermes can
// identify us. Copy the value from DevTools → Network → any request →
// Request Headers → x-amzn-oidc-data.
const HERMES_TOKEN = process.env.HERMES_TOKEN;

// Fallback: some deployments also accept a session cookie.
const HERMES_COOKIE = process.env.HERMES_COOKIE;

if (!HERMES_TOKEN && !HERMES_COOKIE) {
  console.error(
    "ERROR: Either HERMES_TOKEN or HERMES_COOKIE environment variable is required.\n\n" +
    "To get HERMES_TOKEN (preferred):\n" +
    "  1. Log in to " + HERMES_BASE_URL + " in your browser\n" +
    "  2. Open DevTools → Network tab, click any /api/ request\n" +
    "  3. Copy the value of the 'x-amzn-oidc-data' request header\n" +
    "  4. Set HERMES_TOKEN=<value> in your mcp.json env block\n\n" +
    "To get HERMES_COOKIE (fallback):\n" +
    "  1. Log in to " + HERMES_BASE_URL + " in your browser\n" +
    "  2. Open DevTools → Application → Cookies\n" +
    "  3. Copy all cookie values as a single string\n" +
    "  4. Set HERMES_COOKIE=<value> in your mcp.json env block"
  );
  process.exit(1);
}

async function hermesRequest(
  path: string,
  options: RequestInit = {}
): Promise<Response> {
  const url = `${HERMES_BASE_URL}${path}`;
  const headers: Record<string, string> = {
    "Content-Type": "application/json",
    ...(options.headers as Record<string, string>),
  };

  // Prefer ALB OIDC JWT header; fall back to cookie.
  if (HERMES_TOKEN) {
    headers["x-amzn-oidc-data"] = HERMES_TOKEN;
  } else if (HERMES_COOKIE) {
    headers["Cookie"] = HERMES_COOKIE;
  }

  return fetch(url, { ...options, headers });
}

async function hermesJson<T>(path: string, options: RequestInit = {}): Promise<T> {
  const res = await hermesRequest(path, options);
  if (!res.ok) {
    const text = await res.text().catch(() => res.statusText);
    throw new Error(`Hermes API error ${res.status}: ${text}`);
  }
  return res.json() as Promise<T>;
}

const server = new McpServer({
  name: "hermes",
  version: "1.0.0",
});

// ── Tool: search documents via Algolia proxy ──────────────────────────────────
server.registerTool(
  "hermes_search",
  {
    description:
      "Search documents in Hermes using full-text search. Returns matching published documents with titles, owners, product areas, and document types.",
    inputSchema: z.object({
      query: z.string().describe("Search query text"),
      doc_type: z
        .string()
        .optional()
        .describe("Filter by document type, e.g. 'RFC', 'PRD', 'FRD'"),
      product: z
        .string()
        .optional()
        .describe("Filter by product name"),
      page: z
        .number()
        .int()
        .min(0)
        .default(0)
        .describe("Page number (0-based, default 0)"),
      hits_per_page: z
        .number()
        .int()
        .min(1)
        .max(50)
        .default(10)
        .describe("Number of results per page (default 10, max 50)"),
    }),
  },
  async ({ query, doc_type, product, page, hits_per_page }) => {
    try {
      const facetFilters: string[][] = [];
      if (doc_type) facetFilters.push([`docType:${doc_type}`]);
      if (product) facetFilters.push([`product:${product}`]);

      const body = {
        query,
        page,
        hitsPerPage: hits_per_page,
        ...(facetFilters.length > 0 && { facetFilters }),
      };

      const data = await hermesJson<{ hits: unknown[]; nbHits: number; nbPages: number }>(
        "/1/indexes/docs/query",
        { method: "POST", body: JSON.stringify(body) }
      );

      return {
        content: [
          {
            type: "text" as const,
            text: JSON.stringify(
              {
                total_hits: data.nbHits,
                page,
                total_pages: data.nbPages,
                hits: data.hits,
              },
              null,
              2
            ),
          },
        ],
      };
    } catch (err) {
      return {
        content: [{ type: "text" as const, text: `Search failed: ${err instanceof Error ? err.message : String(err)}` }],
        isError: true,
      };
    }
  }
);

// ── Tool: get document details ────────────────────────────────────────────────
server.registerTool(
  "hermes_get_document",
  {
    description:
      "Get full metadata for a specific Hermes document by its ID (e.g. 'RFC-123'). Returns title, status, approvers, summary, custom fields, and links.",
    inputSchema: z.object({
      document_id: z.string().describe("Document ID, e.g. 'RFC-123' or 'PRD-45'"),
    }),
  },
  async ({ document_id }) => {
    try {
      const data = await hermesJson<unknown>(`/api/v2/documents/${document_id}`);
      return {
        content: [{ type: "text" as const, text: JSON.stringify(data, null, 2) }],
      };
    } catch (err) {
      return {
        content: [{ type: "text" as const, text: `Failed to get document: ${err instanceof Error ? err.message : String(err)}` }],
        isError: true,
      };
    }
  }
);

// ── Tool: list drafts ─────────────────────────────────────────────────────────
server.registerTool(
  "hermes_list_drafts",
  {
    description:
      "List draft documents in Hermes. Returns documents that are in draft/review state, not yet published.",
    inputSchema: z.object({
      owner: z
        .string()
        .optional()
        .describe("Filter drafts by owner email address"),
    }),
  },
  async ({ owner }) => {
    try {
      const params = owner ? `?owner=${encodeURIComponent(owner)}` : "";
      const data = await hermesJson<unknown>(`/api/v2/drafts${params}`);
      return {
        content: [{ type: "text" as const, text: JSON.stringify(data, null, 2) }],
      };
    } catch (err) {
      return {
        content: [{ type: "text" as const, text: `Failed to list drafts: ${err instanceof Error ? err.message : String(err)}` }],
        isError: true,
      };
    }
  }
);

// ── Tool: list document types ─────────────────────────────────────────────────
server.registerTool(
  "hermes_list_document_types",
  {
    description:
      "List all available document types configured in this Hermes instance (e.g. RFC, PRD, FRD) with their descriptions and custom fields.",
    inputSchema: z.object({}),
  },
  async () => {
    try {
      const data = await hermesJson<unknown>("/api/v2/document-types");
      return {
        content: [{ type: "text" as const, text: JSON.stringify(data, null, 2) }],
      };
    } catch (err) {
      return {
        content: [{ type: "text" as const, text: `Failed to list document types: ${err instanceof Error ? err.message : String(err)}` }],
        isError: true,
      };
    }
  }
);

// ── Tool: list products ───────────────────────────────────────────────────────
server.registerTool(
  "hermes_list_products",
  {
    description:
      "List all products/areas configured in this Hermes instance. Useful for knowing valid product filter values when searching.",
    inputSchema: z.object({}),
  },
  async () => {
    try {
      const data = await hermesJson<unknown>("/api/v2/products");
      return {
        content: [{ type: "text" as const, text: JSON.stringify(data, null, 2) }],
      };
    } catch (err) {
      return {
        content: [{ type: "text" as const, text: `Failed to list products: ${err instanceof Error ? err.message : String(err)}` }],
        isError: true,
      };
    }
  }
);

// ── Tool: list projects ───────────────────────────────────────────────────────
server.registerTool(
  "hermes_list_projects",
  {
    description:
      "List projects in Hermes. Projects group related documents together.",
    inputSchema: z.object({
      page: z
        .number()
        .int()
        .min(1)
        .default(1)
        .describe("Page number (1-based, default 1)"),
    }),
  },
  async ({ page }) => {
    try {
      const data = await hermesJson<unknown>(`/api/v2/projects?page=${page}`);
      return {
        content: [{ type: "text" as const, text: JSON.stringify(data, null, 2) }],
      };
    } catch (err) {
      return {
        content: [{ type: "text" as const, text: `Failed to list projects: ${err instanceof Error ? err.message : String(err)}` }],
        isError: true,
      };
    }
  }
);

// ── Tool: get current user ────────────────────────────────────────────────────
server.registerTool(
  "hermes_me",
  {
    description:
      "Get the currently authenticated user's profile in Hermes. Useful to confirm the session cookie is working.",
    inputSchema: z.object({}),
  },
  async () => {
    try {
      const data = await hermesJson<unknown>("/api/v2/me");
      return {
        content: [{ type: "text" as const, text: JSON.stringify(data, null, 2) }],
      };
    } catch (err) {
      return {
        content: [{ type: "text" as const, text: `Failed to get current user: ${err instanceof Error ? err.message : String(err)}` }],
        isError: true,
      };
    }
  }
);

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Hermes MCP server running on stdio");
}

main().catch((err) => {
  console.error("Fatal error:", err);
  process.exit(1);
});
