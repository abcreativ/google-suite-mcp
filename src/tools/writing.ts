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
    "Calls spreadsheets.values.update to overwrite a single contiguous range with a 2D value array. Use when replacing existing cell content at a known address (e.g. 'Sheet1!A1:C5'), or populating an empty area in one shot. Do not use when: targeting multiple disconnected ranges (use sheets_write_multiple_ranges); adding rows after existing data (use sheets_append_rows); inserting empty structural rows (use sheets_insert_rows); writing a single formula (use sheets_write_formula); dispersing formulas across cells (use sheets_write_formulas); building a new formatted table (use sheets_write_table). Returns: 'Updated {N} cell(s) in {range}'. Values starting with = are evaluated as formulas (USER_ENTERED mode). The values array is row-major: outer array indexes rows, inner array indexes columns.",
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
    "Calls spreadsheets.values.batchUpdate to write values to several disconnected ranges in one atomic API call. Use when populating non-contiguous areas in a single round trip, such as filling a header block and a data block at different locations; or when updating totals rows in multiple sections simultaneously. Do not use when: writing to a single range (use sheets_write_range); adding rows after existing data (use sheets_append_rows); inserting empty rows (use sheets_insert_rows); writing a single formula (use sheets_write_formula); dispersing formulas across cells (use sheets_write_formulas); building a new formatted table (use sheets_write_table). Returns: 'Updated {N} cell(s) across {M} range(s)'. Each entry in the data array requires a range in A1 notation (e.g. 'Sheet1!A1:C3') and a 2D values array (row-major: outer array = rows, inner array = cells). Defaults to USER_ENTERED so formulas starting with = are evaluated; pass RAW to store literal strings. All ranges must belong to the same spreadsheet.",
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
    "Calls spreadsheets.values.append with INSERT_ROWS to add rows after the last occupied row in the detected table range. Use when adding new records to an existing table without touching current data, or when the exact insertion row is unknown. Do not use when: writing to a fixed address (use sheets_write_range); targeting multiple disconnected ranges (use sheets_write_multiple_ranges); inserting empty structural rows to push content down (use sheets_insert_rows); writing a single formula (use sheets_write_formula); dispersing formulas across cells (use sheets_write_formulas); building a new formatted table (use sheets_write_table). Returns: 'Appended {N} row(s) to {range}'. The range parameter identifies the table to extend (e.g. 'Sheet1!A1' or 'Sheet1!A:A'); the API detects the actual end of data automatically. Existing rows are never overwritten.",
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
    "Calls spreadsheets.batchUpdate with an insertDimension request to push existing rows down and create empty rows at a given position. Use when making room for future data at a specific row without writing any values, or when inserting structural blank rows into an existing layout. Do not use when: adding data rows at the end of a table (use sheets_append_rows); writing values to existing rows (use sheets_write_range); targeting multiple disconnected ranges (use sheets_write_multiple_ranges); writing a single formula (use sheets_write_formula); dispersing formulas across cells (use sheets_write_formulas); building a new formatted table (use sheets_write_table); inserting columns instead (use sheets_insert_columns). Returns: 'Inserted {N} row(s) at row {row_index+1} in \"{sheet}\"'. row_index is 0-based; count sets how many empty rows to insert (minimum 1). Set inherit_from_before to copy formatting from the row above.",
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

