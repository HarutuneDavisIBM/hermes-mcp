#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";

const HERMES_BASE_URL =
  process.env.HERMES_BASE_URL ||
  "https://hermes-sharepoint.hashicorp.services";

// Hermes (SharePoint deployment) authenticates via an AWS ALB session cookie.
// Log in at your Hermes URL, then open DevTools → Application → Cookies and
// copy the AWSELBAuthSessionCookie-0 value (including the name prefix).
// Set HERMES_COOKIE="AWSELBAuthSessionCookie-0=<value>" in your mcp.json env.
const HERMES_COOKIE = process.env.HERMES_COOKIE;

if (!HERMES_COOKIE) {
  console.error(
    "ERROR: HERMES_COOKIE environment variable is required.\n\n" +
    "To get your session cookie:\n" +
    "  1. Log in to " + HERMES_BASE_URL + " in your browser\n" +
    "  2. Open DevTools → Application tab → Cookies → select the site\n" +
    "  3. Find the cookie named 'AWSELBAuthSessionCookie-0'\n" +
    "  4. Copy the Name=Value pair: AWSELBAuthSessionCookie-0=<value>\n" +
    "  5. Set HERMES_COOKIE=<that full string> in your mcp.json env block\n\n" +
    "Note: This cookie expires after a few hours. Repeat these steps to refresh it."
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
    "Cookie": HERMES_COOKIE!,
    ...(options.headers as Record<string, string>),
  };

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
      "Search documents in Hermes using full-text search. Returns matching published documents with titles, owners, product areas, and document types. " +
      "Each result contains an 'objectID' field (a SharePoint GUID like '01XOO7K4...') — pass that value to hermes_get_document to fetch full metadata. " +
      "Do NOT pass the human-readable 'docNumber' (e.g. 'RFC-123') to hermes_get_document; it will return 404.",
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
      "Get full metadata for a specific Hermes document by its SharePoint objectID. Returns title, status, approvers, summary, custom fields, and links. " +
      "IMPORTANT: The document_id must be the 'objectID' field from hermes_search results (a SharePoint GUID like '01XOO7K4...'), " +
      "NOT the human-readable docNumber like 'RFC-123' or 'HVS-022' — those will return a 404. " +
      "Always call hermes_search first to obtain the objectID, then pass it here.",
    inputSchema: z.object({
      document_id: z.string().describe(
        "The SharePoint objectID of the document — a GUID string like '01XOO7K4NVWSKOYH3XMVB3TF54U4DNOEC5'. " +
        "Found in the 'objectID' field of hermes_search results. Do NOT use docNumber (e.g. 'RFC-123')."
      ),
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
      "List draft documents in Hermes. Returns documents that are in WIP (work-in-progress) status. " +
      "Optionally filter by owner email address.",
    inputSchema: z.object({
      owner: z
        .string()
        .optional()
        .describe("Filter drafts by owner email address"),
    }),
  },
  async ({ owner }) => {
    try {
      // The /api/v2/drafts REST endpoint is broken on this deployment (always 500).
      // Fall back to querying the Algolia 'drafts' index directly, then filter
      // by owner client-side (the owners field is not facetable in Algolia).
      const PAGE_SIZE = 200; // Algolia max per request
      let page = 0;
      let totalPages = 1;
      const hits: unknown[] = [];

      do {
        const body = { query: "", hitsPerPage: PAGE_SIZE, page };
        const data = await hermesJson<{ hits: unknown[]; nbHits: number; nbPages: number }>(
          "/1/indexes/drafts/query",
          { method: "POST", body: JSON.stringify(body) }
        );
        totalPages = data.nbPages;
        hits.push(...data.hits);
        page++;
        // Stop early once we have enough results (avoid scanning all 3000+ drafts
        // when no owner filter is set — cap at first 5 pages = 1000 docs).
        if (!owner && page >= 5) break;
      } while (page < totalPages);

      // Filter by owner if requested (case-insensitive).
      const ownerLower = owner?.toLowerCase();
      const filtered = ownerLower
        ? hits.filter((h) => {
            const owners = (h as Record<string, unknown>).owners;
            if (Array.isArray(owners)) {
              return owners.some(
                (o) => typeof o === "string" && o.toLowerCase() === ownerLower
              );
            }
            return false;
          })
        : hits;

      // Strip _highlightResult noise from each hit.
      const clean = filtered.map((h) => {
        const { _highlightResult, ...rest } = h as Record<string, unknown>;
        void _highlightResult;
        return rest;
      });

      return {
        content: [
          {
            type: "text" as const,
            text: JSON.stringify(
              { total: clean.length, drafts: clean },
              null,
              2
            ),
          },
        ],
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

// ── Tool: create draft ────────────────────────────────────────────────────────
server.registerTool(
  "hermes_create_draft",
  {
    description:
      "Create a new document draft in Hermes. Returns the new document's ID which can be used with other tools. " +
      "The draft will have WIP (work-in-progress) status until published via hermes_request_review.",
    inputSchema: z.object({
      title: z.string().describe("Document title (required)"),
      doc_type: z
        .string()
        .describe(
          "Document type abbreviation, e.g. 'RFC', 'PRD', 'FRD', 'ADR', 'Memo', 'PRFAQ'. " +
          "Use hermes_list_document_types to see all available types."
        ),
      product: z
        .string()
        .optional()
        .describe("Product or area name. Use hermes_list_products for valid values."),
      product_abbreviation: z
        .string()
        .optional()
        .describe("Short product abbreviation used as the doc number prefix, e.g. 'HVS', 'TF'. Falls back to 'TODO' if omitted."),
      summary: z
        .string()
        .optional()
        .describe("Short summary or abstract of the document"),
      contributors: z
        .array(z.string())
        .optional()
        .describe("List of contributor email addresses"),
    }),
  },
  async ({ title, doc_type, product, product_abbreviation, summary, contributors }) => {
    try {
      const body: Record<string, unknown> = { title, docType: doc_type };
      if (product) body.product = product;
      if (product_abbreviation) body.productAbbreviation = product_abbreviation;
      if (summary) body.summary = summary;
      if (contributors && contributors.length > 0) body.contributors = contributors;

      const data = await hermesJson<{ id: string }>("/api/v2/drafts", {
        method: "POST",
        body: JSON.stringify(body),
      });
      return {
        content: [{ type: "text" as const, text: JSON.stringify(data, null, 2) }],
      };
    } catch (err) {
      return {
        content: [{ type: "text" as const, text: `Failed to create draft: ${err instanceof Error ? err.message : String(err)}` }],
        isError: true,
      };
    }
  }
);

// ── Tool: update draft ────────────────────────────────────────────────────────
server.registerTool(
  "hermes_update_draft",
  {
    description:
      "Update a document draft in Hermes (WIP status only). Supports updating title, summary, product, " +
      "contributors, approvers, approver groups, and custom fields. All fields are optional — only provided " +
      "fields are changed. The document_id must be the SharePoint objectID from hermes_list_drafts or hermes_create_draft.",
    inputSchema: z.object({
      document_id: z.string().describe(
        "SharePoint objectID of the draft to update (from hermes_list_drafts or hermes_create_draft)"
      ),
      title: z.string().optional().describe("New title for the document"),
      summary: z.string().optional().describe("New summary or abstract"),
      product: z.string().optional().describe("New product or area name"),
      contributors: z
        .array(z.string())
        .optional()
        .describe("Updated list of contributor email addresses (replaces existing)"),
      approvers: z
        .array(z.string())
        .optional()
        .describe("Updated list of approver email addresses (replaces existing)"),
      approver_groups: z
        .array(z.string())
        .optional()
        .describe("Updated list of approver group names (replaces existing)"),
      owners: z
        .array(z.string())
        .optional()
        .describe("Updated list of owner email addresses (replaces existing)"),
      custom_fields: z
        .array(
          z.object({
            name: z.string().describe(
              "Camel-case internal key for the custom field as returned by hermes_get_document " +
              "under 'customEditableFields' (e.g. 'currentVersion', 'prd', 'stakeholders'). " +
              "Do NOT use the display name here."
            ),
            value: z.unknown().describe("Custom field value"),
            type: z.string().optional().describe(
              "Custom field type — MUST be uppercase as returned by the API: 'STRING' or 'PEOPLE'. " +
              "Lowercase values ('string', 'people') are rejected with a 400 error."
            ),
            display_name: z.string().optional().describe(
              "Human-readable label for the field as shown in the UI " +
              "(e.g. 'Current Version', 'PRD', 'Stakeholders'). " +
              "Matches the 'displayName' in hermes_get_document customEditableFields."
            ),
          })
        )
        .optional()
        .describe(
          "Custom fields specific to the document type. " +
          "IMPORTANT: 'name' must be the camelCase key (e.g. 'currentVersion'), " +
          "'type' must be uppercase (e.g. 'STRING', 'PEOPLE'), and " +
          "'display_name' must match the displayName from the API exactly. " +
          "Call hermes_get_document on an existing doc or hermes_list_document_types to discover valid keys and types."
        ),
    }),
  },
  async ({
    document_id,
    title,
    summary,
    product,
    contributors,
    approvers,
    approver_groups,
    owners,
    custom_fields,
  }) => {
    try {
      const body: Record<string, unknown> = {};
      if (title !== undefined) body.title = title;
      if (summary !== undefined) body.summary = summary;
      if (product !== undefined) body.product = product;
      if (contributors !== undefined) body.contributors = contributors;
      if (approvers !== undefined) body.approvers = approvers;
      if (approver_groups !== undefined) body.approverGroups = approver_groups;
      if (owners !== undefined) body.owners = owners;
      if (custom_fields !== undefined) {
        body.customFields = custom_fields.map((cf) => ({
          name: cf.name,
          value: cf.value,
          ...(cf.type && { type: cf.type }),
          ...(cf.display_name && { displayName: cf.display_name }),
        }));
      }

      const res = await hermesRequest(`/api/v2/drafts/${document_id}`, {
        method: "PATCH",
        body: JSON.stringify(body),
      });
      if (!res.ok) {
        const text = await res.text().catch(() => res.statusText);
        throw new Error(`Hermes API error ${res.status}: ${text}`);
      }
      const data = await res.json().catch(() => ({}));
      return {
        content: [
          {
            type: "text" as const,
            text: JSON.stringify({ success: true, document_id, ...data }, null, 2),
          },
        ],
      };
    } catch (err) {
      return {
        content: [{ type: "text" as const, text: `Failed to update draft: ${err instanceof Error ? err.message : String(err)}` }],
        isError: true,
      };
    }
  }
);

// ── Tool: update published document ──────────────────────────────────────────
server.registerTool(
  "hermes_update_document",
  {
    description:
      "Update a published document's metadata in Hermes. Supports changing status (In-Review, Approved, Obsolete), " +
      "title, summary, owners, contributors, approvers, approver groups, and custom fields. " +
      "For drafts (WIP), use hermes_update_draft instead. " +
      "Valid statuses: 'In-Review', 'Approved', 'Obsolete'.",
    inputSchema: z.object({
      document_id: z.string().describe(
        "SharePoint objectID of the published document to update (from hermes_search or hermes_get_document)"
      ),
      status: z
        .enum(["In-Review", "Approved", "Obsolete"])
        .optional()
        .describe(
          "New document status. " +
          "'In-Review' — document is under active review by approvers. " +
          "'Approved' — document has been approved. " +
          "'Obsolete' — document is no longer current."
        ),
      title: z.string().optional().describe("New document title"),
      summary: z.string().optional().describe("New summary or abstract"),
      owners: z
        .array(z.string())
        .optional()
        .describe("Updated list of owner email addresses (replaces existing)"),
      contributors: z
        .array(z.string())
        .optional()
        .describe("Updated list of contributor email addresses (replaces existing)"),
      approvers: z
        .array(z.string())
        .optional()
        .describe("Updated list of approver email addresses (replaces existing)"),
      approver_groups: z
        .array(z.string())
        .optional()
        .describe("Updated list of approver group names (replaces existing)"),
      custom_fields: z
        .array(
          z.object({
            name: z.string().describe(
              "Camel-case internal key for the custom field as returned by hermes_get_document " +
              "under 'customEditableFields' (e.g. 'currentVersion', 'prd', 'stakeholders'). " +
              "Do NOT use the display name here."
            ),
            value: z.unknown().describe("Custom field value"),
            type: z.string().optional().describe(
              "Custom field type — MUST be uppercase as returned by the API: 'STRING' or 'PEOPLE'. " +
              "Lowercase values ('string', 'people') are rejected with a 400 error."
            ),
            display_name: z.string().optional().describe(
              "Human-readable label for the field as shown in the UI " +
              "(e.g. 'Current Version', 'PRD', 'Stakeholders'). " +
              "Matches the 'displayName' in hermes_get_document customEditableFields."
            ),
          })
        )
        .optional()
        .describe(
          "Custom fields specific to the document type. " +
          "IMPORTANT: 'name' must be the camelCase key (e.g. 'currentVersion'), " +
          "'type' must be uppercase (e.g. 'STRING', 'PEOPLE'), and " +
          "'display_name' must match the displayName from the API exactly. " +
          "Call hermes_get_document on an existing doc or hermes_list_document_types to discover valid keys and types."
        ),
    }),
  },
  async ({
    document_id,
    status,
    title,
    summary,
    owners,
    contributors,
    approvers,
    approver_groups,
    custom_fields,
  }) => {
    try {
      const body: Record<string, unknown> = {};
      if (status !== undefined) body.status = status;
      if (title !== undefined) body.title = title;
      if (summary !== undefined) body.summary = summary;
      if (owners !== undefined) body.owners = owners;
      if (contributors !== undefined) body.contributors = contributors;
      if (approvers !== undefined) body.approvers = approvers;
      if (approver_groups !== undefined) body.approverGroups = approver_groups;
      if (custom_fields !== undefined) {
        body.customFields = custom_fields.map((cf) => ({
          name: cf.name,
          value: cf.value,
          ...(cf.type && { type: cf.type }),
          ...(cf.display_name && { displayName: cf.display_name }),
        }));
      }

      const res = await hermesRequest(`/api/v2/documents/${document_id}`, {
        method: "PATCH",
        body: JSON.stringify(body),
      });
      if (!res.ok) {
        const text = await res.text().catch(() => res.statusText);
        throw new Error(`Hermes API error ${res.status}: ${text}`);
      }
      const data = await res.json().catch(() => ({}));
      return {
        content: [
          {
            type: "text" as const,
            text: JSON.stringify({ success: true, document_id, ...data }, null, 2),
          },
        ],
      };
    } catch (err) {
      return {
        content: [{ type: "text" as const, text: `Failed to update document: ${err instanceof Error ? err.message : String(err)}` }],
        isError: true,
      };
    }
  }
);

// ── Tool: request review (publish draft) ─────────────────────────────────────
server.registerTool(
  "hermes_request_review",
  {
    description:
      "Publish a draft document and request review from approvers. This moves the document from WIP " +
      "status to 'In-Review' and notifies the assigned approvers. " +
      "The draft must already have approvers set (use hermes_update_draft to add them first). " +
      "The document_id must be the SharePoint objectID of a draft.",
    inputSchema: z.object({
      document_id: z.string().describe(
        "SharePoint objectID of the draft to publish for review (from hermes_list_drafts or hermes_create_draft)"
      ),
    }),
  },
  async ({ document_id }) => {
    try {
      const res = await hermesRequest(`/api/v2/reviews/${document_id}`, {
        method: "POST",
        body: JSON.stringify({}),
      });
      if (!res.ok) {
        const text = await res.text().catch(() => res.statusText);
        throw new Error(`Hermes API error ${res.status}: ${text}`);
      }
      const data = await res.json().catch(() => ({}));
      return {
        content: [
          {
            type: "text" as const,
            text: JSON.stringify(
              { success: true, document_id, message: "Document published and review requested", ...data },
              null,
              2
            ),
          },
        ],
      };
    } catch (err) {
      return {
        content: [{ type: "text" as const, text: `Failed to request review: ${err instanceof Error ? err.message : String(err)}` }],
        isError: true,
      };
    }
  }
);

// ── Tool: approve document ────────────────────────────────────────────────────
server.registerTool(
  "hermes_approve_document",
  {
    description:
      "Approve a document in Hermes as the currently authenticated user. The document must be in " +
      "'In-Review' or 'Approved' status. The current user must be listed as an approver on the document. " +
      "Use hermes_get_document to check the current status and approver list before calling this.",
    inputSchema: z.object({
      document_id: z.string().describe(
        "SharePoint objectID of the document to approve (from hermes_search or hermes_get_document)"
      ),
    }),
  },
  async ({ document_id }) => {
    try {
      const res = await hermesRequest(`/api/v2/documents/${document_id}/approvals`, {
        method: "POST",
        body: JSON.stringify({}),
      });
      if (!res.ok) {
        const text = await res.text().catch(() => res.statusText);
        throw new Error(`Hermes API error ${res.status}: ${text}`);
      }
      return {
        content: [
          {
            type: "text" as const,
            text: JSON.stringify(
              { success: true, document_id, message: "Document approved successfully" },
              null,
              2
            ),
          },
        ],
      };
    } catch (err) {
      return {
        content: [{ type: "text" as const, text: `Failed to approve document: ${err instanceof Error ? err.message : String(err)}` }],
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
