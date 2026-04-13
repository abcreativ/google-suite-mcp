/**
 * High-level table builder tool.
 *
 * sheets_write_table - write a complete styled data table in a single call:
 * data + formula-column expansion + header styling + banded rows + borders +
 * freeze + column sizing + number formats.
 *
 * Follows the same accumulate-then-batch pattern as sheets_build_dashboard.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getSheetsClient, extractFileId } from "../client/google-client.js";
import { formatSuccess } from "../utils/response.js";
import { withErrorHandling } from "../utils/error-handler.js";
import { parseColor } from "../utils/color.js";
import { columnToLetter, quoteSheetName, parseA1Notation } from "../utils/range.js";
import {
  resolveOrCreateSheet,
  makeGridRange,
  repeatCellRequest,
  DEFAULT_COLORS,
  solidBorder,
} from "../utils/sheet-builder.js";
import type { sheets_v4 } from "googleapis";
import type { CellValue } from "../types/sheets.js";

// ─── Tool registration ────────────────────────────────────────────────────────

export function registerTableTools(server: McpServer): void {
  server.tool(
    "sheets_write_table",
    "Builds a new formatted table by writing data, expanding row-by-row formula templates, and applying header styling, banded rows, borders, freeze, and column sizing in two to three batched API calls. Use when creating a complete formatted table from a headers array and row data in one operation; or when you need automatic column sizing and banding applied in the same call as the data write. Do not use when: editing an existing table's data or formatting (use sheets_write_range or sheets_format_cells); targeting multiple disconnected ranges (use sheets_write_multiple_ranges); adding rows to an existing table (use sheets_append_rows); inserting empty rows (use sheets_insert_rows); writing a single formula (use sheets_write_formula); dispersing formulas across cells (use sheets_write_formulas); building a multi-section dashboard with KPIs and charts (use sheets_build_dashboard). Returns: 'Table written: {M} cols x {N} rows in \"{sheet}\"'. headers: array of strings (column labels); rows: 2D row-major array (outer = rows, inner = cells) matching the values structure used in sheets_write_range. The sheet is created if it does not exist.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      sheet: z.string().optional().describe("Sheet name (created if not exists, default: first sheet)"),
      start_cell: z.string().optional().describe("Top-left cell, e.g. 'A1' (default: A1)"),

      // Data
      headers: z.array(z.string()).min(1).describe("Column header labels"),
      rows: z
        .array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()])))
        .describe("2D array of row data (formulas work - prefix with =)"),

      // Formula columns - server-side expansion
      column_formulas: z
        .array(
          z.object({
            column: z.number().int().describe("0-based column index"),
            formula_template: z
              .string()
              .describe("Formula with {row} placeholder for row number, e.g. '=C{row}*D{row}'"),
          })
        )
        .optional()
        .describe("Auto-expanded formulas: {row} is replaced with each data row number"),

      // Styling
      header_style: z
        .object({
          bold: z.boolean().optional(),
          background_color: z.string().optional(),
          text_color: z.string().optional(),
          font_size: z.number().optional(),
        })
        .optional()
        .describe("Header row style (defaults: bold, dark bg, white text)"),

      banded_rows: z.boolean().optional().describe("Alternate row shading (default: true)"),
      banded_color: z.string().optional().describe("Banded row color (default: light grey)"),
      border: z.boolean().optional().describe("Box border + inner grid (default: true)"),
      freeze_header: z.boolean().optional().describe("Freeze the header row (default: true)"),
      auto_resize: z.boolean().optional().describe("Auto-fit column widths (default: true)"),

      // Per-column number formats
      column_formats: z
        .array(
          z.object({
            column: z.number().int().describe("0-based column index"),
            format: z.string().describe("Number format pattern, e.g. '$#,##0.00', '0.0%'"),
          })
        )
        .optional()
        .describe("Number format per column"),

      // Column widths
      column_widths: z
        .array(
          z.object({
            column: z.number().int().describe("0-based column index"),
            width: z.number().int().describe("Width in pixels"),
          })
        )
        .optional(),
    },
    withErrorHandling(
      async ({
        spreadsheet_id,
        sheet,
        start_cell,
        headers,
        rows,
        column_formulas,
        header_style,
        banded_rows,
        banded_color,
        border,
        freeze_header,
        auto_resize,
        column_formats,
        column_widths,
      }) => {
        const sheetsClient = await getSheetsClient();
        const id = extractFileId(spreadsheet_id);

        // ── Resolve sheet ───────────────────────────────────────────────
        let sheetName: string;
        let sheetId: number;

        if (sheet) {
          sheetName = sheet;
        } else {
          const meta = await sheetsClient.spreadsheets.get({
            spreadsheetId: id,
            fields: "sheets.properties.title",
          });
          sheetName = meta.data.sheets?.[0]?.properties?.title ?? "Sheet1";
        }
        sheetId = await resolveOrCreateSheet(sheetsClient, id, sheetName);

        // ── Parse start cell ────────────────────────────────────────────
        const parsed = parseA1Notation(start_cell ?? "A1");
        const startRow = parsed.startRow;
        const startCol = parsed.startCol;
        const numCols = headers.length;
        const numDataRows = rows.length;
        const totalRows = 1 + numDataRows; // header + data
        const quotedSheet = quoteSheetName(sheetName);

        // ── Build values: expand formula templates ──────────────────────
        const allValues: CellValue[][] = [headers];

        for (let i = 0; i < numDataRows; i++) {
          const row = [...rows[i]];
          // Expand formula templates for this row
          if (column_formulas) {
            const actualRow = startRow + 1 + i + 1; // 1-based row number (header + offset + data index)
            for (const cf of column_formulas) {
              if (cf.column >= 0 && cf.column < numCols) {
                row[cf.column] = cf.formula_template.replace(/\{row\}/g, String(actualRow));
              }
            }
          }
          allValues.push(row);
        }

        // ── Write values ────────────────────────────────────────────────
        const endCol = startCol + numCols - 1;
        const endRow = startRow + totalRows - 1;
        const rangeA1 = `${quotedSheet}!${columnToLetter(startCol)}${startRow + 1}:${columnToLetter(endCol)}${endRow + 1}`;

        await sheetsClient.spreadsheets.values.update({
          spreadsheetId: id,
          range: rangeA1,
          valueInputOption: "USER_ENTERED",
          requestBody: { values: allValues as CellValue[][] },
        });

        // ── Accumulate formatting requests ──────────────────────────────
        const requests: sheets_v4.Schema$Request[] = [];

        // Header styling
        const hStyle = header_style ?? {};
        const headerBg = hStyle.background_color
          ? parseColor(hStyle.background_color)
          : DEFAULT_COLORS.TABLE_HEADER_BG;
        const headerFg = hStyle.text_color
          ? parseColor(hStyle.text_color)
          : DEFAULT_COLORS.TABLE_HEADER_FG;

        requests.push(
          repeatCellRequest(
            sheetId,
            startRow, startCol, startRow + 1, startCol + numCols,
            {
              backgroundColorStyle: { rgbColor: headerBg },
              textFormat: {
                bold: hStyle.bold ?? true,
                fontSize: hStyle.font_size ?? 10,
                foregroundColorStyle: { rgbColor: headerFg },
              },
              horizontalAlignment: "CENTER",
            },
            "userEnteredFormat.backgroundColorStyle,userEnteredFormat.textFormat,userEnteredFormat.horizontalAlignment"
          )
        );

        // Banded rows via native addBanding (default: true)
        if (banded_rows !== false && numDataRows > 0) {
          const bandColor = banded_color
            ? parseColor(banded_color)
            : DEFAULT_COLORS.ALT_ROW_BG;

          requests.push({
            addBanding: {
              bandedRange: {
                range: makeGridRange(
                  sheetId,
                  startRow, startCol,
                  startRow + totalRows, startCol + numCols
                ),
                rowProperties: {
                  headerColorStyle: { rgbColor: headerBg },
                  firstBandColorStyle: { rgbColor: { red: 1, green: 1, blue: 1 } },
                  secondBandColorStyle: { rgbColor: bandColor },
                },
              },
            },
          });
        }

        // Borders (default: true)
        if (border !== false) {
          const borderObj = solidBorder();
          const innerBorder: sheets_v4.Schema$Border = {
            style: "SOLID",
            colorStyle: { rgbColor: parseColor("#c5cae9") },
          };

          requests.push({
            updateBorders: {
              range: makeGridRange(
                sheetId,
                startRow, startCol,
                startRow + totalRows, startCol + numCols
              ),
              top: borderObj,
              bottom: borderObj,
              left: borderObj,
              right: borderObj,
              innerHorizontal: innerBorder,
              innerVertical: innerBorder,
            },
          });
        }

        // Freeze header (default: true)
        if (freeze_header !== false) {
          requests.push({
            updateSheetProperties: {
              properties: {
                sheetId,
                gridProperties: { frozenRowCount: startRow + 1 },
              },
              fields: "gridProperties.frozenRowCount",
            },
          });
        }

        // Column number formats
        if (column_formats) {
          for (const cf of column_formats) {
            const col = startCol + cf.column;
            requests.push(
              repeatCellRequest(
                sheetId,
                startRow + 1, col,       // data rows only (skip header)
                startRow + totalRows, col + 1,
                { numberFormat: { type: "NUMBER", pattern: cf.format } },
                "userEnteredFormat.numberFormat"
              )
            );
          }
        }

        // Auto-resize first (default: true), then explicit widths override
        if (auto_resize !== false) {
          requests.push({
            autoResizeDimensions: {
              dimensions: {
                sheetId,
                dimension: "COLUMNS",
                startIndex: startCol,
                endIndex: startCol + numCols,
              },
            },
          });
        }

        // Explicit column widths (after auto-resize so they take precedence)
        if (column_widths) {
          for (const cw of column_widths) {
            const col = startCol + cw.column;
            requests.push({
              updateDimensionProperties: {
                range: {
                  sheetId,
                  dimension: "COLUMNS",
                  startIndex: col,
                  endIndex: col + 1,
                },
                properties: { pixelSize: cw.width },
                fields: "pixelSize",
              },
            });
          }
        }

        // ── Apply all formatting in one batch ───────────────────────────
        if (requests.length > 0) {
          await sheetsClient.spreadsheets.batchUpdate({
            spreadsheetId: id,
            requestBody: { requests },
          });
        }

        return formatSuccess(
          `Table written: ${numCols} cols × ${numDataRows} rows in "${sheetName}"`
        );
      }
    )
  );
}
