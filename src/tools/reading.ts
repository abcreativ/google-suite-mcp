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
    "Read cell values from a range. Returns tab-separated rows.",
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
    "Read multiple ranges in one call. Returns each range with tab-separated values.",
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
    "Read entire sheet (or first N rows/cols). Returns tab-separated data.",
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
    "Inspect a single cell: value, formula, format, note, validation rule.",
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
    "Search for text across sheets. Returns matching cell addresses and values.",
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
