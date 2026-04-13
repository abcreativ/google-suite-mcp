/**
 * Reading tools: read ranges, multiple ranges, entire sheets, cell info, search.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getSheetsClient, extractFileId } from "../client/google-client.js";
import { formatSuccess, formatError } from "../utils/response.js";
import { withErrorHandling } from "../utils/error-handler.js";
import { columnToLetter, quoteSheetName } from "../utils/range.js";
import type { CellValue } from "../types/sheets.js";

// ─── Helpers ──────────────────────────────────────────────────────────────────

const ValueRenderOption = z
  .enum(["FORMATTED_VALUE", "UNFORMATTED_VALUE", "FORMULA"])
  .optional()
  .describe("How values are rendered (default: FORMATTED_VALUE)");

/** Serialises a 2D array to a compact text table. */
function formatGrid(
  values: CellValue[][] | null | undefined,
  range: string
): string {
  if (!values || values.length === 0) {
    return `${range}: (empty)`;
  }
  const rows = values.map((row) =>
    row.map((cell) => (cell === null || cell === undefined ? "" : String(cell))).join("\t")
  );
  return `${range}:\n${rows.join("\n")}`;
}

// ─── Tool registration ────────────────────────────────────────────────────────

export function registerReadingTools(server: McpServer): void {
  // ─── sheets_read_range ─────────────────────────────────────────────────────

  server.tool(
    "sheets_read_range",
    "Reads cell values from a single contiguous range using spreadsheets.values.get; by default returns formatted display values (FORMATTED_VALUE). Use when the user asks to see the contents of a specific region of a sheet. Use when you need to verify existing data before overwriting it with sheets_write_range. Do not use when: reading multiple non-contiguous ranges in one call - use sheets_read_multiple_ranges instead; reading an entire sheet tab - use sheets_read_sheet instead; reading a single cell's metadata, formula, note, or validation rule - use sheets_get_cell_info instead; searching for a specific value - use sheets_search_values instead; reading formula strings instead of computed values - use sheets_get_formulas instead. Returns: '{range}:\\n{tab-separated values per row}', or '{range}: (empty)' if the range has no content. Parameters: - range: A1 notation including sheet name, e.g. 'Sheet1!A1:C10' - value_render_option: FORMATTED_VALUE (default), UNFORMATTED_VALUE, or FORMULA.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      range: z.string().describe("A1 notation range, e.g. 'Sheet1!A1:C10' or 'A1:C10'"),
      value_render_option: ValueRenderOption,
    },
    withErrorHandling(async ({ spreadsheet_id, range, value_render_option }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);

      const res = await sheets.spreadsheets.values.get({
        spreadsheetId: id,
        range,
        valueRenderOption: value_render_option ?? "FORMATTED_VALUE",
      });

      return formatSuccess(formatGrid(res.data.values as CellValue[][], res.data.range ?? range));
    })
  );

  // ─── sheets_read_multiple_ranges ───────────────────────────────────────────

  server.tool(
    "sheets_read_multiple_ranges",
    "Reads values from several non-contiguous ranges in a single API call using spreadsheets.values.batchGet; returns each range's data as a separate section. Use when you need to read a header row and a data block at the same time without two round trips. Use when comparing values from different parts of a spreadsheet in one operation. Do not use when: reading a single contiguous range - use sheets_read_range instead; reading an entire sheet - use sheets_read_sheet instead; reading a single cell's details - use sheets_get_cell_info instead; searching for a value - use sheets_search_values instead; reading formula strings - use sheets_get_formulas instead. Returns: each range formatted as '{range}:\\n{tab-separated rows}', sections separated by a blank line. Parameters: - ranges: array of A1 notation strings, e.g. ['Sheet1!A1:C1', 'Sheet1!A10:C20'] - value_render_option: FORMATTED_VALUE (default), UNFORMATTED_VALUE, or FORMULA.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      ranges: z.array(z.string()).describe("Array of A1 notation ranges"),
      value_render_option: ValueRenderOption,
    },
    withErrorHandling(async ({ spreadsheet_id, ranges, value_render_option }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);

      const res = await sheets.spreadsheets.values.batchGet({
        spreadsheetId: id,
        ranges,
        valueRenderOption: value_render_option ?? "FORMATTED_VALUE",
      });

      const sections = (res.data.valueRanges ?? []).map((vr) =>
        formatGrid(vr.values as CellValue[][], vr.range ?? "(unknown range)")
      );

      return formatSuccess(sections.join("\n\n"));
    })
  );

  // ─── sheets_read_sheet ─────────────────────────────────────────────────────

  server.tool(
    "sheets_read_sheet",
    "Reads all cell values from a sheet tab using spreadsheets.values.get with the full sheet range; optional max_rows and max_columns limits prevent loading oversized sheets. Use when the user asks to see everything in a tab, or when the data range is unknown. Use when reading a small-to-medium sheet in full for analysis or inspection. Do not use when: the target range is already known - use sheets_read_range for efficiency; reading multiple specific regions - use sheets_read_multiple_ranges instead; reading a single cell's details - use sheets_get_cell_info instead; searching for a specific value - use sheets_search_values instead; reading formula strings - use sheets_get_formulas instead. Returns: '{range}:\\n{tab-separated values per row}', or '{range}: (empty)'. Parameters: - sheet_name: tab name (optional; defaults to the first tab if omitted) - max_rows: truncate output after this many rows - max_columns: truncate output after this many columns.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      sheet_name: z.string().optional().describe("Sheet/tab name (default: first sheet)"),
      max_rows: z.number().int().optional().describe("Maximum rows to return"),
      max_columns: z.number().int().optional().describe("Maximum columns to return"),
      value_render_option: ValueRenderOption,
    },
    withErrorHandling(async ({ spreadsheet_id, sheet_name, max_rows, max_columns, value_render_option }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);

      // Resolve the target sheet name
      let resolvedName: string;
      if (sheet_name) {
        resolvedName = sheet_name;
      } else {
        const meta = await sheets.spreadsheets.get({
          spreadsheetId: id,
          fields: "sheets.properties.title",
        });
        resolvedName = meta.data.sheets?.[0]?.properties?.title ?? "Sheet1";
      }

      // Build range - apply limits if set
      let range: string;
      const quoted = quoteSheetName(resolvedName);
      if (max_rows || max_columns) {
        const colStr = max_columns ? columnToLetter(max_columns - 1) : "ZZZ";
        const rowStr = max_rows ? String(max_rows) : "";
        range = `${quoted}!A1:${colStr}${rowStr}`;
      } else {
        range = quoted;
      }

      const res = await sheets.spreadsheets.values.get({
        spreadsheetId: id,
        range,
        valueRenderOption: value_render_option ?? "FORMATTED_VALUE",
      });

      return formatSuccess(formatGrid(res.data.values as CellValue[][], res.data.range ?? range));
    })
  );

  // ─── sheets_get_cell_info ──────────────────────────────────────────────────

  server.tool(
    "sheets_get_cell_info",
    "Retrieves detailed metadata for a single cell using spreadsheets.get with includeGridData, returning its effective value, formula string, formatted display value, number format, note, and data validation rule. Use when the user asks what formula or format a specific cell contains. Use when verifying a cell's data validation rule before modifying it with sheets_set_validation. Do not use when: reading a range of cell values - use sheets_read_range instead; reading formula strings across a range - use sheets_get_formulas instead; reading multiple non-contiguous ranges - use sheets_read_multiple_ranges instead; reading an entire sheet - use sheets_read_sheet instead; searching for a value - use sheets_search_values instead. Returns: multi-line string with 'Cell: {addr}', then lines for Value, Formula, Formatted, Number format, Note, and Validation only when each is present; or '{cell}: (empty)'. Parameters: - cell: single cell address in A1 notation, e.g. 'Sheet1!B3'.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      cell: z.string().describe("Cell address in A1 notation, e.g. 'Sheet1!B3' or 'B3'"),
    },
    withErrorHandling(async ({ spreadsheet_id, cell }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);

      const res = await sheets.spreadsheets.get({
        spreadsheetId: id,
        ranges: [cell],
        includeGridData: true,
        fields: "sheets.data.rowData.values(formattedValue,userEnteredValue,effectiveValue,note,userEnteredFormat.numberFormat,dataValidation)",
      });

      const row = res.data.sheets?.[0]?.data?.[0]?.rowData?.[0];
      const cellData = row?.values?.[0];

      if (!cellData) {
        return formatSuccess(`${cell}: (empty)`);
      }

      const lines: string[] = [`Cell: ${cell}`];

      const ev = cellData.effectiveValue;
      if (ev) {
        const val = ev.stringValue ?? ev.numberValue ?? ev.boolValue ?? ev.formulaValue ?? "(empty)";
        lines.push(`Value: ${val}`);
      }

      const uev = cellData.userEnteredValue;
      if (uev?.formulaValue) {
        lines.push(`Formula: ${uev.formulaValue}`);
      }

      if (cellData.formattedValue !== undefined) {
        lines.push(`Formatted: ${cellData.formattedValue}`);
      }

      const fmt = cellData.userEnteredFormat?.numberFormat;
      if (fmt) {
        lines.push(`Number format: ${fmt.type} "${fmt.pattern ?? ""}"`);
      }

      if (cellData.note) {
        lines.push(`Note: ${cellData.note}`);
      }

      const dv = cellData.dataValidation;
      if (dv?.condition) {
        lines.push(`Validation: ${dv.condition.type}`);
      }

      return formatSuccess(lines.join("\n"));
    })
  );

  // ─── sheets_search_values ──────────────────────────────────────────────────

  server.tool(
    "sheets_search_values",
    "Searches cell values across one or all tabs in a spreadsheet for a text pattern using a client-side scan of FORMATTED_VALUE data; supports regex and caps results at max_results. Use when the user asks to find which cells contain a specific value, name, or pattern. Use when you need the cell addresses of all matches before operating on them. Do not use when: replacing matched text - use sheets_find_replace or sheets_find_replace_many instead; reading a known cell range - use sheets_read_range instead; reading a single cell's details - use sheets_get_cell_info instead; reading formula strings - use sheets_get_formulas instead. Returns: '{N} match(es):\\n{SheetName!ColRow}: {value}' for each match, or 'No cells match \"{pattern}\".' A truncation notice is appended if max_results is reached. Parameters: - pattern: text to match (case-insensitive substring by default) - use_regex: true to treat pattern as a regular expression - sheet_name: limit search to one tab (optional; defaults to all tabs).",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      pattern: z.string().describe("Text to search for (case-insensitive substring match)"),
      sheet_name: z
        .string()
        .optional()
        .describe("Limit to this sheet (default: all)"),
      use_regex: z.boolean().optional().describe("Treat pattern as regex"),
      max_results: z.number().int().min(1).optional().describe("Max matches to return (default: 1000)"),
    },
    withErrorHandling(async ({ spreadsheet_id, pattern, sheet_name, use_regex, max_results }) => {
      const limit = max_results ?? 1000;
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);

      // Get all sheet names
      const meta = await sheets.spreadsheets.get({
        spreadsheetId: id,
        fields: "sheets.properties(title,sheetId)",
      });

      const allSheets = (meta.data.sheets ?? []).map((s) => s.properties?.title ?? "");
      const targetSheets = sheet_name
        ? allSheets.filter((n) => n.toLowerCase() === sheet_name.toLowerCase())
        : allSheets;

      if (targetSheets.length === 0) {
        return formatSuccess("No matching sheets found.");
      }

      let regex: RegExp;
      try {
        if (use_regex && pattern.length > 200) {
          return formatError("Regex pattern too long (max 200 chars). Use a simpler pattern or set use_regex=false.");
        }
        regex = use_regex
          ? new RegExp(pattern, "i")
          : new RegExp(escapeRegex(pattern), "i");
        // Quick ReDoS safety check - test against a moderate string
        if (use_regex) {
          const start = Date.now();
          regex.test("a".repeat(50));
          if (Date.now() - start > 100) {
            return formatError("Regex pattern is too expensive (potential ReDoS). Simplify the pattern.");
          }
        }
      } catch (e) {
        return formatError(`Invalid regex pattern: ${(e as Error).message}`);
      }

      const matches: string[] = [];
      let truncated = false;

      outer:
      for (const sheetTitle of targetSheets) {
        const res = await sheets.spreadsheets.values.get({
          spreadsheetId: id,
          range: quoteSheetName(sheetTitle),
          valueRenderOption: "FORMATTED_VALUE",
        });

        const values = (res.data.values ?? []) as string[][];
        for (let rowIdx = 0; rowIdx < values.length; rowIdx++) {
          const row = values[rowIdx];
          for (let colIdx = 0; colIdx < row.length; colIdx++) {
            if (regex.test(String(row[colIdx] ?? ""))) {
              const cellAddr = `${quoteSheetName(sheetTitle)}!${columnToLetter(colIdx)}${rowIdx + 1}`;
              matches.push(`${cellAddr}: ${row[colIdx]}`);
              if (matches.length >= limit) { truncated = true; break outer; }
            }
          }
        }
      }

      if (matches.length === 0) {
        return formatSuccess(`No cells match "${pattern}".`);
      }

      const suffix = truncated ? `\n(capped at ${limit} - set max_results for more)` : "";
      return formatSuccess(`${matches.length} match(es):\n${matches.join("\n")}${suffix}`);
    })
  );
}

// ─── Internal helpers ─────────────────────────────────────────────────────────

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
