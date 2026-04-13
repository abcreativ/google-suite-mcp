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
    "Adds a new tab to an existing spreadsheet using spreadsheets.batchUpdate with addSheet; the new sheet is empty and returns its numeric sheetId. Use when the user asks to add a new tab to an existing spreadsheet, such as adding a 'Summary' or 'Config' sheet. Use when you need to create a destination sheet before writing data to it with sheets_write_range or sheets_build_sheet. Do not use when: creating a new spreadsheet file - use sheets_create instead; duplicating an existing tab with its data - use sheets_duplicate_sheet instead; renaming a tab - use sheets_rename_sheet instead; listing existing tabs - use sheets_list_sheets instead; deleting a tab - use sheets_delete_sheet instead; reordering tabs - use sheets_reorder_sheets instead. Returns: 'Added sheet \"{title}\" (sheetId: {sheetId})'. The returned sheetId is required by tools that take a numeric sheet ID. Parameters: - title: tab name (optional; Google assigns a default name if omitted) - tab_color: hex color string, e.g. '#FF0000'.",
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
    "Permanently deletes a tab from a spreadsheet using spreadsheets.batchUpdate with deleteSheet; all data on the sheet is lost and the action cannot be undone. Use when the user asks to remove a tab that is no longer needed, such as a staging sheet after its data has been merged. Use when cleaning up temporary sheets created during a multi-step workflow. Do not use when: adding a tab - use sheets_add_sheet instead; renaming a tab - use sheets_rename_sheet instead; listing tabs - use sheets_list_sheets instead; duplicating a tab - use sheets_duplicate_sheet instead; reordering tabs - use sheets_reorder_sheets instead; deleting the entire spreadsheet file - use sheets_delete instead. Returns: 'Deleted sheet \"{sheet}\"'. Parameters: - sheet: tab name or numeric sheet ID to delete, e.g. 'Staging' or '12345'.",
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
    "Renames a tab in a spreadsheet using spreadsheets.batchUpdate with updateSheetProperties targeting the title field; the numeric sheetId is unchanged. Use when the user asks to change a tab name, such as renaming 'Sheet1' to 'Sales Data'. Use when a sheet name in a formula reference needs to be updated to match a new naming convention. Do not use when: adding a new tab - use sheets_add_sheet instead; deleting a tab - use sheets_delete_sheet instead; duplicating a tab - use sheets_duplicate_sheet instead; renaming the entire spreadsheet file - use sheets_rename instead; reordering tabs - use sheets_reorder_sheets instead; listing tab names - use sheets_list_sheets instead. Returns: 'Renamed \"{sheet}\" → \"{new_title}\"'. Parameters: - sheet: current tab name or numeric sheet ID, e.g. 'Sheet1' - new_title: the replacement tab name, e.g. 'Sales Data'.",
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
    "Copies an existing tab - including its cell values, formulas, formatting, conditional rules, and data validation - to a new tab in the same spreadsheet using spreadsheets.batchUpdate with duplicateSheet; returns the new tab's sheetId. Use when the user asks to clone a sheet template or make a backup copy before editing. Use when creating a new period's data sheet by duplicating a prior period's template. Do not use when: adding a blank tab - use sheets_add_sheet instead; renaming a tab - use sheets_rename_sheet instead; deleting a tab - use sheets_delete_sheet instead; reordering tabs - use sheets_reorder_sheets instead; listing tabs - use sheets_list_sheets instead; copying a spreadsheet file - use sheets_copy instead. Returns: 'Duplicated to \"{new_title}\" (sheetId: {sheetId})'. Parameters: - sheet: source tab name or numeric sheet ID, e.g. 'Template' - new_title: title for the duplicate, e.g. 'April 2025' (optional; Google assigns a default if omitted) - insert_at_index: 0-based position to place the new tab, e.g. 0 = first tab.",
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
    "Reorders all tabs in a spreadsheet to a specified sequence using spreadsheets.batchUpdate with updateSheetProperties on each sheet's index; every tab in the spreadsheet must be listed or the call will error. Use when the user asks to rearrange tabs, such as moving a Summary tab to the front. Use when organizing a workbook with many sheets into a logical reading order. Do not use when: adding a tab - use sheets_add_sheet instead; deleting a tab - use sheets_delete_sheet instead; renaming a tab - use sheets_rename_sheet instead; duplicating a tab - use sheets_duplicate_sheet instead; listing current tab order - use sheets_list_sheets instead. Returns: 'Reordered {N} sheets.'. Parameters: - sheet_order: array of all tab names or numeric IDs in the desired final order, e.g. ['Summary', 'Data', 'Charts']; must include every tab - omitting any tab causes an error.",
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
    "Reads all tab metadata from a spreadsheet using spreadsheets.get with fields scoped to sheets.properties, returning each tab's sheetId, index, title, grid dimensions, and tab color. Use when you need to enumerate available tabs before operating on one by name or ID. Use when looking up a sheetId required by protection, validation, or formatting tools. Do not use when: adding a tab - use sheets_add_sheet instead; deleting a tab - use sheets_delete_sheet instead; renaming a tab - use sheets_rename_sheet instead; duplicating a tab - use sheets_duplicate_sheet instead; reordering tabs - use sheets_reorder_sheets instead; reading top-level spreadsheet metadata - use sheets_get_info instead. Returns: one line per tab formatted as '[{sheetId}] index:{N} \"{title}\"  {rows}r × {cols}c', or 'No sheets found.' Parameters: - spreadsheetId: the ID from the sheet URL (between /d/ and /edit).",
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
