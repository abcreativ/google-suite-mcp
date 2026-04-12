/**
 * Spreadsheet-level tools: create, list, get info, delete, copy, rename.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getSheetsClient, getDriveClient, extractFileId } from "../client/google-client.js";
import { formatSuccess, formatError } from "../utils/response.js";
import { withErrorHandling } from "../utils/error-handler.js";

export function registerSpreadsheetTools(server: McpServer): void {
  // ─── sheets_create ──────────────────────────────────────────────────────────

  server.tool(
    "sheets_create",
    "Create a new spreadsheet with optional tab names. Returns ID and URL.",
    {
      title: z.string().optional().describe("Spreadsheet title (default: 'Untitled spreadsheet')"),
      sheet_names: z
        .array(z.string())
        .optional()
        .describe("Names of sheets/tabs to create (default: ['Sheet1'])"),
    },
    withErrorHandling(async ({ title, sheet_names }) => {
      const sheets = await getSheetsClient();

      const sheetRequests =
        sheet_names && sheet_names.length > 0
          ? sheet_names.map((name, i) => ({
              properties: { title: name, index: i },
            }))
          : [{ properties: { title: "Sheet1", index: 0 } }];

      const res = await sheets.spreadsheets.create({
        requestBody: {
          properties: { title: title ?? "Untitled spreadsheet" },
          sheets: sheetRequests,
        },
      });

      const { spreadsheetId, spreadsheetUrl } = res.data;
      return formatSuccess(
        `Created: ${spreadsheetId}\nURL: ${spreadsheetUrl}`
      );
    })
  );

  // ─── sheets_list ────────────────────────────────────────────────────────────

  server.tool(
    "sheets_list",
    "List spreadsheets, optionally filtered by Drive query. Returns IDs, titles, URLs.",
    {
      query: z
        .string()
        .optional()
        .describe("Search query (appended to mimeType filter). E.g. 'name contains \"budget\"'"),
      page_token: z.string().optional().describe("Pagination token from a previous response"),
      page_size: z.number().int().min(1).max(1000).optional().describe("Results per page (default 50)"),
    },
    withErrorHandling(async ({ query, page_token, page_size }) => {
      const drive = await getDriveClient();

      const mimeFilter = "mimeType='application/vnd.google-apps.spreadsheet'";
      const q = query ? `${mimeFilter} and ${query}` : mimeFilter;

      const res = await drive.files.list({
        q,
        pageToken: page_token,
        pageSize: page_size ?? 50,
        fields: "nextPageToken, files(id, name, webViewLink, modifiedTime)",
        orderBy: "modifiedTime desc",
      });

      const files = res.data.files ?? [];
      if (files.length === 0) {
        return formatSuccess("No spreadsheets found.");
      }

      const lines = files.map(
        (f) => `${f.id}  ${f.name}  modified: ${f.modifiedTime ?? "?"}  ${f.webViewLink}`
      );

      const next = res.data.nextPageToken
        ? `\nnextPageToken: ${res.data.nextPageToken}`
        : "";

      return formatSuccess(lines.join("\n") + next);
    })
  );

  // ─── sheets_get_info ────────────────────────────────────────────────────────

  server.tool(
    "sheets_get_info",
    "Get spreadsheet metadata: title, locale, timezone, sheet list with dimensions.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
    },
    withErrorHandling(async ({ spreadsheet_id }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);

      const res = await sheets.spreadsheets.get({
        spreadsheetId: id,
        fields: "spreadsheetId,spreadsheetUrl,properties,sheets.properties",
      });

      const { spreadsheetId, spreadsheetUrl, properties, sheets: sheetList } = res.data;
      const props = properties ?? {};

      const sheetLines = (sheetList ?? []).map((s) => {
        const p = s.properties ?? {};
        return `  [${p.sheetId}] ${p.title}  (${p.gridProperties?.rowCount ?? "?"}r × ${p.gridProperties?.columnCount ?? "?"}c)`;
      });

      return formatSuccess(
        [
          `ID: ${spreadsheetId}`,
          `URL: ${spreadsheetUrl}`,
          `Title: ${props.title}`,
          `Locale: ${props.locale}`,
          `Timezone: ${props.timeZone}`,
          `Sheets (${sheetLines.length}):`,
          ...sheetLines,
        ].join("\n")
      );
    })
  );

  // ─── sheets_delete ──────────────────────────────────────────────────────────

  server.tool(
    "sheets_delete",
    "DESTRUCTIVE: Trash a spreadsheet (recoverable from Drive trash). Confirm with user first.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
    },
    withErrorHandling(async ({ spreadsheet_id }) => {
      const drive = await getDriveClient();
      const id = extractFileId(spreadsheet_id);

      await drive.files.update({
        fileId: id,
        requestBody: { trashed: true },
      });

      return formatSuccess(`Moved to trash: ${id}`);
    })
  );

  // ─── sheets_copy ────────────────────────────────────────────────────────────

  server.tool(
    "sheets_copy",
    "Copy a spreadsheet. Returns new ID and URL.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL to copy"),
      title: z.string().optional().describe("Title for the copy (default: 'Copy of <original>')"),
    },
    withErrorHandling(async ({ spreadsheet_id, title }) => {
      const drive = await getDriveClient();
      const id = extractFileId(spreadsheet_id);

      const res = await drive.files.copy({
        fileId: id,
        requestBody: title ? { name: title } : {},
        fields: "id, name, webViewLink",
      });

      return formatSuccess(
        `Copied: ${res.data.id}\nTitle: ${res.data.name}\nURL: ${res.data.webViewLink}`
      );
    })
  );

  // ─── sheets_rename ──────────────────────────────────────────────────────────

  server.tool(
    "sheets_rename",
    "Rename a spreadsheet.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      title: z.string().describe("New title"),
    },
    withErrorHandling(async ({ spreadsheet_id, title }) => {
      const drive = await getDriveClient();
      const id = extractFileId(spreadsheet_id);

      await drive.files.update({
        fileId: id,
        requestBody: { name: title },
      });

      return formatSuccess(`Renamed ${id} → "${title}"`);
    })
  );
}
