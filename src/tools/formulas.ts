/**
 * Formula tools: write formulas, read formulas, write array formulas.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getSheetsClient, extractFileId } from "../client/google-client.js";
import { formatSuccess } from "../utils/response.js";
import { withErrorHandling } from "../utils/error-handler.js";
import type { CellValue } from "../types/sheets.js";

export function registerFormulaTools(server: McpServer): void {
  // ─── sheets_write_formula ──────────────────────────────────────────────────

  server.tool(
    "sheets_write_formula",
    "Calls spreadsheets.values.update on a single cell with USER_ENTERED to write one formula. Use when placing a calculated expression in a specific cell, such as a summary at the bottom of a column; or when setting a standalone lookup or validation formula in an isolated cell. Do not use when: writing multiple formulas to scattered cells (use sheets_write_formulas); writing an ARRAYFORMULA that spills across a range (use sheets_write_array_formula); writing mixed data and formulas in a contiguous block (use sheets_write_range); targeting multiple disconnected ranges (use sheets_write_multiple_ranges); adding rows after existing data (use sheets_append_rows); inserting empty rows (use sheets_insert_rows); building a new formatted table (use sheets_write_table). Returns: 'Formula written to {range}'. The formula must start with =; the handler does not prepend it automatically.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      cell: z.string().describe("Target cell in A1 notation, e.g. 'Sheet1!B2' or 'B2'"),
      formula: z.string().describe("Formula string, e.g. '=SUM(A1:A10)'"),
    },
    withErrorHandling(async ({ spreadsheet_id, cell, formula }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);

      const res = await sheets.spreadsheets.values.update({
        spreadsheetId: id,
        range: cell,
        valueInputOption: "USER_ENTERED",
        requestBody: { values: [[formula]] },
      });

      return formatSuccess(
        `Formula written to ${res.data.updatedRange ?? cell}`
      );
    })
  );

  // ─── sheets_write_formulas (bulk) ───────────────────────────────────────────

  server.tool(
    "sheets_write_formulas",
    "Calls spreadsheets.values.batchUpdate with USER_ENTERED to write multiple formulas to scattered cells in one API call. Use when distributing calculated expressions across non-adjacent cells in a single round trip, such as setting column totals across a summary row; or when applying the same formula pattern to several named cells simultaneously. Do not use when: writing a single formula (use sheets_write_formula); writing an ARRAYFORMULA that spills results (use sheets_write_array_formula); writing mixed data and formulas in a contiguous block (use sheets_write_range); targeting multiple disconnected ranges (use sheets_write_multiple_ranges); adding rows after existing data (use sheets_append_rows); inserting empty rows (use sheets_insert_rows); building a new formatted table (use sheets_write_table). Returns: 'Wrote {N} formula(s)'. If a formula string does not start with =, the handler prepends it automatically. Each entry requires a cell address in A1 notation (e.g. 'Sheet1!C2') and a formula string.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      formulas: z
        .array(
          z.object({
            cell: z.string().describe("Target cell, e.g. 'Sheet1!C2'"),
            formula: z.string().describe("Formula starting with ="),
          })
        )
        .min(1)
        .describe("Array of {cell, formula} pairs"),
    },
    withErrorHandling(async ({ spreadsheet_id, formulas }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);

      // Ensure all formulas start with = to avoid writing literal text
      const data = formulas.map((f) => {
        const formula = f.formula.startsWith("=") ? f.formula : `=${f.formula}`;
        return { range: f.cell, values: [[formula]] };
      });

      const res = await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: id,
        requestBody: {
          valueInputOption: "USER_ENTERED",
          data,
        },
      });

      const totalCells = res.data.totalUpdatedCells ?? 0;
      return formatSuccess(`Wrote ${totalCells} formula(s)`);
    })
  );

  // ─── sheets_get_formulas ───────────────────────────────────────────────────

  server.tool(
    "sheets_get_formulas",
    "Reads raw formula strings from a cell range using spreadsheets.values.get with valueRenderOption=FORMULA; non-formula cells return their display value instead. Use when the user asks to inspect or audit the formulas in a sheet, or when you need to verify what expression is in a calculated cell before overwriting it. Use when debugging a formula result by reading its source text rather than its computed output. Do not use when: reading computed cell values - use sheets_read_range instead; reading a single cell's value and metadata - use sheets_get_cell_info instead; reading multiple non-contiguous ranges - use sheets_read_multiple_ranges instead; searching for a value - use sheets_search_values instead; writing a formula - use sheets_write_formula or sheets_write_formulas instead. Returns: '{range}:\\n{tab-separated formula strings per row}', or '{range}: (empty)' if the range has no content. Parameters: - range: A1 notation including sheet name, e.g. 'Sheet1!A1:D10'.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      range: z.string().describe("A1 notation range, e.g. 'Sheet1!A1:D10'"),
    },
    withErrorHandling(async ({ spreadsheet_id, range }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);

      const res = await sheets.spreadsheets.values.get({
        spreadsheetId: id,
        range,
        valueRenderOption: "FORMULA",
      });

      const values = (res.data.values ?? []) as CellValue[][];
      if (values.length === 0) {
        return formatSuccess(`${range}: (empty)`);
      }

      const rows = values.map((row) =>
        row.map((cell) => (cell === null || cell === undefined ? "" : String(cell))).join("\t")
      );

      return formatSuccess(`${res.data.range ?? range}:\n${rows.join("\n")}`);
    })
  );

  // ─── sheets_write_array_formula ────────────────────────────────────────────

  server.tool(
    "sheets_write_array_formula",
    "Writes an ARRAYFORMULA to a single cell using spreadsheets.values.update with USER_ENTERED; automatically wraps the expression in ARRAYFORMULA() if not already wrapped, and prepends = if missing. Use when the user asks to write a formula that spills results across a range from a single cell, such as =ARRAYFORMULA(A1:A10*B1:B10). Use when setting up a dynamic column that auto-fills as data is added below. Do not use when: writing a single non-array formula to one cell - use sheets_write_formula instead; writing multiple formulas to scattered cells - use sheets_write_formulas instead; writing an array of values (not a formula) - use sheets_write_range instead; writing mixed data and formulas in a block - use sheets_write_range with USER_ENTERED mode; building a full formatted table - use sheets_write_table instead. Returns: 'Array formula written to {updatedRange}: {finalFormula}' where finalFormula is the ARRAYFORMULA-wrapped expression actually written. Parameters: - cell: target cell in A1 notation, e.g. 'Sheet1!A1'; the formula spills from this anchor - formula: the expression to wrap, e.g. '=A1:A10*B1:B10' or '=ARRAYFORMULA(A1:A10*B1:B10)'.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      cell: z.string().describe("Target cell in A1 notation, e.g. 'Sheet1!A1' or 'A1'"),
      formula: z.string().describe("Formula to wrap in ARRAYFORMULA, e.g. '=A1:A10*B1:B10' or '=ARRAYFORMULA(A1:A10*B1:B10)'"),
    },
    withErrorHandling(async ({ spreadsheet_id, cell, formula }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);

      // Ensure formula starts with =
      const baseFormula = formula.startsWith("=") ? formula.slice(1) : formula;

      // Wrap in ARRAYFORMULA if not already
      const isWrapped = /^ARRAYFORMULA\s*\(/i.test(baseFormula.trim());
      const finalFormula = isWrapped
        ? `=${baseFormula}`
        : `=ARRAYFORMULA(${baseFormula})`;

      const res = await sheets.spreadsheets.values.update({
        spreadsheetId: id,
        range: cell,
        valueInputOption: "USER_ENTERED",
        requestBody: { values: [[finalFormula]] },
      });

      return formatSuccess(
        `Array formula written to ${res.data.updatedRange ?? cell}: ${finalFormula}`
      );
    })
  );
}
