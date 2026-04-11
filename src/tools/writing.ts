/**
 * Writing tools: write ranges, batch write, append rows, insert/delete rows/columns.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getSheetsClient, extractFileId } from "../client/google-client.js";
import { formatSuccess } from "../utils/response.js";
import { withErrorHandling } from "../utils/error-handler.js";
import { resolveSheetId } from "../utils/sheet-resolver.js";
import type { CellValue } from "../types/sheets.js";

// ─── Shared schema ────────────────────────────────────────────────────────────

const ValueInputOption = z
  .enum(["RAW", "USER_ENTERED"])
  .optional()
  .describe("How input is interpreted (default: USER_ENTERED)");

// ─── Tool registration ────────────────────────────────────────────────────────

export function registerWritingTools(server: McpServer): void {
  // ─── sheets_write_range ────────────────────────────────────────────────────

  server.tool(
    "sheets_write_range",
    "Write values to a range (overwrites existing content). Formulas: prefix with = (interpreted via USER_ENTERED). Bulk: sheets_write_multiple_ranges.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      range: z.string().describe("A1 notation range, e.g. 'Sheet1!A1'"),
      values: z
        .array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()])))
        .describe("2D array of values [rows][cols]"),
      value_input_option: ValueInputOption,
    },
    withErrorHandling(async ({ spreadsheet_id, range, values, value_input_option }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);

      const res = await sheets.spreadsheets.values.update({
        spreadsheetId: id,
        range,
        valueInputOption: value_input_option ?? "USER_ENTERED",
        requestBody: { values: values as CellValue[][] },
      });

      const updated = res.data.updatedCells ?? 0;
      const updatedRange = res.data.updatedRange ?? range;
      return formatSuccess(`Updated ${updated} cell(s) in ${updatedRange}`);
    })
  );

  // ─── sheets_write_multiple_ranges ─────────────────────────────────────────

  server.tool(
    "sheets_write_multiple_ranges",
    "Write values to multiple ranges in one call.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      data: z
        .array(
          z.object({
            range: z.string().describe("A1 notation range"),
            values: z
              .array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()])))
              .describe("2D array of values"),
          })
        )
        .describe("Array of { range, values } objects"),
      value_input_option: ValueInputOption,
    },
    withErrorHandling(async ({ spreadsheet_id, data, value_input_option }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);

      const res = await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: id,
        requestBody: {
          valueInputOption: value_input_option ?? "USER_ENTERED",
          data: data.map((d) => ({
            range: d.range,
            values: d.values as CellValue[][],
          })),
        },
      });

      const totalCells = res.data.totalUpdatedCells ?? 0;
      const totalRanges = res.data.responses?.length ?? data.length;
      return formatSuccess(`Updated ${totalCells} cell(s) across ${totalRanges} range(s)`);
    })
  );

  // ─── sheets_append_rows ────────────────────────────────────────────────────

  server.tool(
    "sheets_append_rows",
    "Append rows after the last occupied row. Non-destructive.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      range: z.string().describe("A1 notation range to append after (e.g. 'Sheet1!A1')"),
      values: z
        .array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()])))
        .describe("Rows to append"),
      value_input_option: ValueInputOption,
    },
    withErrorHandling(async ({ spreadsheet_id, range, values, value_input_option }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);

      const res = await sheets.spreadsheets.values.append({
        spreadsheetId: id,
        range,
        valueInputOption: value_input_option ?? "USER_ENTERED",
        insertDataOption: "INSERT_ROWS",
        requestBody: { values: values as CellValue[][] },
      });

      const updates = res.data.updates;
      return formatSuccess(
        `Appended ${updates?.updatedRows ?? values.length} row(s) to ${updates?.updatedRange ?? range}`
      );
    })
  );

  // ─── sheets_insert_rows ────────────────────────────────────────────────────

  server.tool(
    "sheets_insert_rows",
    "Insert empty rows at a position (0-based index).",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      sheet: z.string().describe("Sheet name or numeric sheet ID"),
      row_index: z.number().int().describe("0-based row index to insert before"),
      count: z.number().int().min(1).describe("Number of rows to insert"),
      inherit_from_before: z
        .boolean()
        .optional()
        .describe("Inherit formatting from the row above (default: false)"),
    },
    withErrorHandling(async ({ spreadsheet_id, sheet, row_index, count, inherit_from_before }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);
      const sheetId = await resolveSheetId(id, sheet);

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: {
          requests: [
            {
              insertDimension: {
                range: {
                  sheetId,
                  dimension: "ROWS",
                  startIndex: row_index,
                  endIndex: row_index + count,
                },
                inheritFromBefore: inherit_from_before ?? false,
              },
            },
          ],
        },
      });

      return formatSuccess(`Inserted ${count} row(s) at row ${row_index + 1} in "${sheet}"`);
    })
  );

  // ─── sheets_insert_columns ─────────────────────────────────────────────────

  server.tool(
    "sheets_insert_columns",
    "Insert empty columns at a position (0-based index).",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      sheet: z.string().describe("Sheet name or numeric sheet ID"),
      column_index: z.number().int().describe("0-based column index to insert before"),
      count: z.number().int().min(1).describe("Number of columns to insert"),
      inherit_from_before: z
        .boolean()
        .optional()
        .describe("Inherit formatting from the column to the left (default: false)"),
    },
    withErrorHandling(async ({ spreadsheet_id, sheet, column_index, count, inherit_from_before }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);
      const sheetId = await resolveSheetId(id, sheet);

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: {
          requests: [
            {
              insertDimension: {
                range: {
                  sheetId,
                  dimension: "COLUMNS",
                  startIndex: column_index,
                  endIndex: column_index + count,
                },
                inheritFromBefore: inherit_from_before ?? false,
              },
            },
          ],
        },
      });

      return formatSuccess(`Inserted ${count} column(s) at column ${column_index + 1} in "${sheet}"`);
    })
  );

  // ─── sheets_delete_rows_columns ────────────────────────────────────────────

  server.tool(
    "sheets_delete_rows_columns",
    "DESTRUCTIVE: Delete rows or columns by index range (0-based, end exclusive). Cannot be undone.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      sheet: z.string().describe("Sheet name or numeric sheet ID"),
      dimension: z.enum(["ROWS", "COLUMNS"]).describe("Whether to delete rows or columns"),
      start_index: z.number().int().describe("0-based start index (inclusive)"),
      end_index: z.number().int().describe("0-based end index (exclusive)"),
    },
    withErrorHandling(async ({ spreadsheet_id, sheet, dimension, start_index, end_index }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);
      const sheetId = await resolveSheetId(id, sheet);

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: {
          requests: [
            {
              deleteDimension: {
                range: {
                  sheetId,
                  dimension,
                  startIndex: start_index,
                  endIndex: end_index,
                },
              },
            },
          ],
        },
      });

      const count = end_index - start_index;
      return formatSuccess(
        `Deleted ${count} ${dimension.toLowerCase().slice(0, -1)}(s) from ${start_index + 1} to ${end_index} in "${sheet}"`
      );
    })
  );
}

