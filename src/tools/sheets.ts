/**
 * Sheet/tab management tools: add, delete, rename, duplicate, reorder, list.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getSheetsClient, extractFileId } from "../client/google-client.js";
import { formatSuccess, formatError } from "../utils/response.js";
import { withErrorHandling } from "../utils/error-handler.js";
import { resolveSheetId } from "../utils/sheet-resolver.js";
import { parseColor } from "../utils/color.js";

// ─── Tool registration ────────────────────────────────────────────────────────

export function registerSheetTools(server: McpServer): void {
  // ─── sheets_add_sheet ──────────────────────────────────────────────────────

  server.tool(
    "sheets_add_sheet",
    "Add a new tab. Returns sheet ID.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      title: z.string().optional().describe("Tab title"),
      row_count: z.number().int().optional().describe("Initial row count (default: 1000)"),
      column_count: z.number().int().optional().describe("Initial column count (default: 26)"),
      tab_color: z.string().optional().describe("Tab color as hex (#RRGGBB)"),
    },
    withErrorHandling(async ({ spreadsheet_id, title, row_count, column_count, tab_color }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);

      const sheetProps: Record<string, unknown> = {};
      if (title) sheetProps.title = title;
      if (tab_color) sheetProps.tabColorStyle = { rgbColor: parseColor(tab_color) };

      const gridProps: Record<string, number> = {};
      if (row_count) gridProps.rowCount = row_count;
      if (column_count) gridProps.columnCount = column_count;
      if (Object.keys(gridProps).length > 0) sheetProps.gridProperties = gridProps;

      const res = await sheets.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: {
          requests: [{ addSheet: { properties: sheetProps } }],
        },
      });

      const newSheet = res.data.replies?.[0]?.addSheet?.properties;
      return formatSuccess(
        `Added sheet "${newSheet?.title}" (sheetId: ${newSheet?.sheetId})`
      );
    })
  );

  // ─── sheets_delete_sheet ───────────────────────────────────────────────────

  server.tool(
    "sheets_delete_sheet",
    "DESTRUCTIVE: Delete a tab by name or sheet ID. Cannot be undone.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      sheet: z.string().describe("Sheet name or numeric sheet ID"),
    },
    withErrorHandling(async ({ spreadsheet_id, sheet }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);
      const sheetId = await resolveSheetId(id, sheet);

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: {
          requests: [{ deleteSheet: { sheetId } }],
        },
      });

      return formatSuccess(`Deleted sheet "${sheet}"`);
    })
  );

  // ─── sheets_rename_sheet ───────────────────────────────────────────────────

  server.tool(
    "sheets_rename_sheet",
    "Rename a tab.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      sheet: z.string().describe("Current sheet name or numeric sheet ID"),
      new_title: z.string().describe("New tab title"),
    },
    withErrorHandling(async ({ spreadsheet_id, sheet, new_title }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);
      const sheetId = await resolveSheetId(id, sheet);

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: {
          requests: [
            {
              updateSheetProperties: {
                properties: { sheetId, title: new_title },
                fields: "title",
              },
            },
          ],
        },
      });

      return formatSuccess(`Renamed "${sheet}" → "${new_title}"`);
    })
  );

  // ─── sheets_duplicate_sheet ────────────────────────────────────────────────

  server.tool(
    "sheets_duplicate_sheet",
    "Duplicate a tab. Returns new sheet ID.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      sheet: z.string().describe("Sheet name or numeric sheet ID to duplicate"),
      new_title: z.string().optional().describe("Title for the duplicate"),
      insert_at_index: z.number().int().optional().describe("Position to insert the duplicate (0-based)"),
    },
    withErrorHandling(async ({ spreadsheet_id, sheet, new_title, insert_at_index }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);
      const sheetId = await resolveSheetId(id, sheet);

      const dupRequest: Record<string, unknown> = {
        sourceSheetId: sheetId,
      };
      if (insert_at_index !== undefined) dupRequest.insertSheetIndex = insert_at_index;
      if (new_title) dupRequest.newSheetName = new_title;

      const res = await sheets.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: {
          requests: [{ duplicateSheet: dupRequest }],
        },
      });

      const newSheet = res.data.replies?.[0]?.duplicateSheet?.properties;
      return formatSuccess(
        `Duplicated to "${newSheet?.title}" (sheetId: ${newSheet?.sheetId})`
      );
    })
  );

  // ─── sheets_reorder_sheets ─────────────────────────────────────────────────

  server.tool(
    "sheets_reorder_sheets",
    "Reorder tabs (list all sheet names/IDs in desired order).",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      sheet_order: z
        .array(z.string())
        .describe("Sheet names or IDs in the desired order (all sheets must be listed)"),
    },
    withErrorHandling(async ({ spreadsheet_id, sheet_order }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);

      const requests = await Promise.all(
        sheet_order.map(async (sheet, index) => {
          const sheetId = await resolveSheetId(id, sheet);
          return {
            updateSheetProperties: {
              properties: { sheetId, index },
              fields: "index",
            },
          };
        })
      );

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: { requests },
      });

      return formatSuccess(`Reordered ${sheet_order.length} sheets.`);
    })
  );

  // ─── sheets_list_sheets ────────────────────────────────────────────────────

  server.tool(
    "sheets_list_sheets",
    "List all tabs with IDs, names, and dimensions.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
    },
    withErrorHandling(async ({ spreadsheet_id }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);

      const res = await sheets.spreadsheets.get({
        spreadsheetId: id,
        fields: "sheets.properties",
      });

      const sheetList = res.data.sheets ?? [];
      if (sheetList.length === 0) {
        return formatSuccess("No sheets found.");
      }

      const lines = sheetList.map((s) => {
        const p = s.properties ?? {};
        const color = p.tabColorStyle?.rgbColor;
        const colorStr = color
          ? ` color:(${Math.round((color.red ?? 0) * 255)},${Math.round((color.green ?? 0) * 255)},${Math.round((color.blue ?? 0) * 255)})`
          : "";
        return `[${p.sheetId}] index:${p.index} "${p.title}"  ${p.gridProperties?.rowCount ?? "?"}r × ${p.gridProperties?.columnCount ?? "?"}c${colorStr}`;
      });

      return formatSuccess(lines.join("\n"));
    })
  );
}
