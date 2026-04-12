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
    "Write a formula to a single cell (include leading '='). For many formulas at once, use sheets_write_formulas (plural).",
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
    "Write multiple formulas to scattered cells in one call. Bulk version of sheets_write_formula.",
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
    "Get formula text for cells in a range (non-formula cells show value).",
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
    "Write an ARRAYFORMULA to a single cell (auto-wraps if needed). For bulk formula writes, include formulas in sheets_write_range values with = prefix.",
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
