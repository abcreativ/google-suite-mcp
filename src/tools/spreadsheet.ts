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
    "Creates a new Google Sheets spreadsheet using spreadsheets.create; optionally sets the title and pre-creates named tabs. Use when the user asks to create a new spreadsheet from scratch. Use when you need a new spreadsheet file as a destination before writing data with sheets_write_range or sheets_build_sheet. Do not use when: adding a tab to an existing spreadsheet - use sheets_add_sheet instead; copying an existing spreadsheet - use sheets_copy instead; listing existing spreadsheets - use sheets_list instead; getting info about a spreadsheet - use sheets_get_info instead; deleting a spreadsheet - use sheets_delete instead; renaming a spreadsheet - use sheets_rename instead. Returns: 'Created: {spreadsheetId}\\nURL: {spreadsheetUrl}'. Parameters: - title: spreadsheet name (optional; defaults to 'Untitled spreadsheet') - sheet_names: array of tab names to create, e.g. ['Summary', 'Data']; defaults to ['Sheet1'].",
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
    "Lists Google Sheets spreadsheets in Drive using drive.files.list filtered to the spreadsheet MIME type; supports Drive query strings for filtering by name or other metadata. Use when the user asks which spreadsheets exist in their Drive, or to find a spreadsheet ID by name. Use when paginating through a large collection of spreadsheets using page_token. Do not use when: listing the tabs within a known spreadsheet - use sheets_list_sheets instead; getting full metadata for one spreadsheet - use sheets_get_info instead; creating a spreadsheet - use sheets_create instead; searching Drive for any file type - use drive_search instead. Returns: one line per file formatted as '{id}  {name}  modified: {timestamp}  {url}', or 'No spreadsheets found.' Appends 'nextPageToken: ...' when more pages exist. Parameters: - query: Drive query string appended to the MIME filter, e.g. 'name contains \"budget\"' - page_size: results per page (default 50, max 1000).",
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
    "Retrieves top-level metadata for a spreadsheet using spreadsheets.get, returning its ID, URL, title, locale, timezone, and a list of all tabs with their sheetIds and grid dimensions. Use when the user asks for details about a spreadsheet, or when you need to discover tab names and sheetIds before operating on them. Use when you need to confirm a spreadsheet exists and get its URL in one call. Do not use when: listing tabs only - use sheets_list_sheets for a more focused response; listing all spreadsheets in Drive - use sheets_list instead; creating a spreadsheet - use sheets_create instead; deleting a spreadsheet - use sheets_delete instead; renaming a spreadsheet - use sheets_rename instead; copying a spreadsheet - use sheets_copy instead. Returns: multi-line string with ID, URL, Title, Locale, Timezone, and a Sheets list where each tab shows '[{sheetId}] {title} ({rows}r × {cols}c)'.",
    {
      spreadsheet_id: z.string().describe("sheet ID from the URL (the token between /d/ and /edit) or the full URL"),
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
    "Moves a spreadsheet to the Google Drive Bin (trash) using drive.files.update with trashed=true; the file is not permanently deleted and can be restored from Drive Bin within 30 days. Use when the user asks to delete or trash a spreadsheet they no longer need. Use when cleaning up a test or temporary spreadsheet after a workflow completes. Do not use when: deleting a tab within a spreadsheet - use sheets_delete_sheet instead; permanently removing a file (requires emptying Drive Bin manually after trashing); creating a spreadsheet - use sheets_create instead; listing spreadsheets - use sheets_list instead; copying a spreadsheet - use sheets_copy instead; renaming a spreadsheet - use sheets_rename instead. Returns: 'Moved to trash: {spreadsheetId}'. Parameters: - spreadsheet_id: the spreadsheet ID or full URL, e.g. '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms' or the full https://docs.google.com URL.",
    {
      spreadsheet_id: z.string().describe("sheet ID from the URL (the token between /d/ and /edit) or the full URL"),
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
    "Creates a full copy of a spreadsheet using drive.files.copy, duplicating all tabs, data, formulas, and formatting into a new file; returns the new spreadsheet's ID and URL. Use when the user asks to make a backup of a spreadsheet before editing. Use when creating a new version of a template file for a new reporting period. Do not use when: duplicating a single tab within the same spreadsheet - use sheets_duplicate_sheet instead; creating a blank spreadsheet - use sheets_create instead; renaming a spreadsheet - use sheets_rename instead; listing spreadsheets - use sheets_list instead; deleting a spreadsheet - use sheets_delete instead; moving a spreadsheet to a folder - use drive_move instead. Returns: 'Copied: {newSpreadsheetId}\\nTitle: {name}\\nURL: {url}'. Parameters: - spreadsheet_id: source spreadsheet ID or URL - title: title for the copy, e.g. 'Budget 2025 Backup' (optional; defaults to 'Copy of {original title}').",
    {
      spreadsheet_id: z.string().describe("sheet ID from the URL (the token between /d/ and /edit) or the full URL of the spreadsheet to copy"),
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
    "Renames a spreadsheet file using drive.files.update with a new name property; the spreadsheet ID and all content remain unchanged. Use when the user asks to change the title of a spreadsheet. Use when correcting a misspelled or auto-generated spreadsheet title after creation. Do not use when: renaming a tab within a spreadsheet - use sheets_rename_sheet instead; copying a spreadsheet with a new name - use sheets_copy instead; creating a spreadsheet - use sheets_create instead; listing spreadsheets - use sheets_list instead; deleting a spreadsheet - use sheets_delete instead; getting spreadsheet metadata - use sheets_get_info instead. Returns: 'Renamed {spreadsheetId} → \"{title}\"'. Parameters: - spreadsheet_id: the spreadsheet ID or full URL - title: the new spreadsheet name, e.g. 'Q1 Sales Report 2025'.",
    {
      spreadsheet_id: z.string().describe("sheet ID from the URL (the token between /d/ and /edit) or the full URL"),
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
