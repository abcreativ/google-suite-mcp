/**
 * Dashboard builder tool.
 *
 * F4.5 - sheets_build_dashboard
 *
 * Orchestrates a complete professional dashboard in a single batchUpdate:
 * - Dashboard title / header row
 * - KPI section with green/red delta coloring
 * - Data tables with bold headers and alternating row colors
 * - Charts
 * - Frozen header row
 * - Auto-resized columns
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getSheetsClient, extractFileId } from "../client/google-client.js";
import { formatSuccess } from "../utils/response.js";
import { withErrorHandling } from "../utils/error-handler.js";
import { parseColor } from "../utils/color.js";
import { toGridRange, columnToLetter, quoteSheetName } from "../utils/range.js";
import { resolveSheetIdFromRange } from "../utils/sheet-resolver.js";
import {
  resolveOrCreateSheet,
  makeGridRange,
  repeatCellRequest,
  DEFAULT_COLORS,
  solidBorder,
} from "../utils/sheet-builder.js";
import type { sheets_v4 } from "googleapis";

// ─── Color aliases (from shared palette) ─────────────────────────────────────

const HEADER_BG = DEFAULT_COLORS.HEADER_BG;
const HEADER_FG = DEFAULT_COLORS.HEADER_FG;
const KPI_BG = DEFAULT_COLORS.KPI_BG;
const KPI_LABEL_FG = DEFAULT_COLORS.KPI_LABEL_FG;
const ALT_ROW_BG = DEFAULT_COLORS.ALT_ROW_BG;
const TABLE_HEADER_BG = DEFAULT_COLORS.TABLE_HEADER_BG;
const TABLE_HEADER_FG = DEFAULT_COLORS.TABLE_HEADER_FG;
const POSITIVE_FG = DEFAULT_COLORS.POSITIVE_FG;
const NEGATIVE_FG = DEFAULT_COLORS.NEGATIVE_FG;

// ─── Schema ───────────────────────────────────────────────────────────────────

const KpiSchema = z.object({
  label: z.string().describe("KPI label, e.g. 'Revenue'"),
  value: z.union([z.string(), z.number()]).describe("KPI value, e.g. '$1.2M' or 1200000"),
  delta: z.string().optional().describe("Change indicator, e.g. '+12%' or '-5%'"),
});

const DataTableSchema = z.object({
  title: z.string().describe("Table title"),
  headers: z.array(z.string()).describe("Column headers"),
  rows: z
    .array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()])))
    .describe("Table data rows"),
});

const ChartSpecSchema = z.object({
  type: z
    .enum(["BAR", "LINE", "PIE", "SCATTER", "AREA", "COLUMN"])
    .describe("Chart type"),
  data_range: z.string().describe("A1 notation data range"),
  title: z.string().optional().describe("Chart title"),
});

// ─── Tool registration ────────────────────────────────────────────────────────

export function registerDashboardTools(server: McpServer): void {
  server.tool(
    "sheets_build_dashboard",
    "Builds a formatted dashboard sheet with KPI metric cards, data tables, embedded charts, frozen header, and banded rows in batched API calls; creates or reuses the target tab. Use when the user asks to build an executive dashboard or summary view from structured data in one operation. Use when a sheet needs KPI cards, a data table, and charts composed together in a fixed layout. Do not use when: building a sheet with arbitrary section types and custom layouts - use sheets_build_sheet instead; creating a single formatted data table - use sheets_write_table instead; writing raw data without formatting - use sheets_write_range instead; adding a chart to an existing sheet - use sheets_create_chart instead. Returns: 'Dashboard built in sheet \"{sheet}\":\\n  Title: {t}\\n  KPIs: {N}\\n  Data tables: {N}\\n  Charts: {N}\\n  Formatting requests: {N}'. Parameters: - sheet_name: tab to build in (created if it does not exist) - kpis: array of {label, value} objects for metric cards - data_tables: array of table specs each with headers and rows - charts: array of chart specs referencing data ranges.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      sheet_name: z
        .string()
        .optional()
        .describe("Sheet name to write the dashboard to (created if not exists, default: 'Dashboard')"),
      title: z.string().optional().describe("Dashboard title displayed at the top"),
      kpis: z.array(KpiSchema).optional().describe("KPI cards to display at the top of the dashboard"),
      data_tables: z.array(DataTableSchema).optional().describe("Data tables to include"),
      charts: z.array(ChartSpecSchema).optional().describe("Charts to create below the data"),
    },
    withErrorHandling(
      async ({ spreadsheet_id, sheet_name, title, kpis, data_tables, charts }) => {
        const sheets = await getSheetsClient();
        const id = extractFileId(spreadsheet_id);
        const targetSheet = sheet_name ?? "Dashboard";
        const sheetId = await resolveOrCreateSheet(sheets, id, targetSheet);

        // ── Layout planning ──────────────────────────────────────────────────
        // Row 0:   Dashboard title (spanning columns A-F or as many as needed)
        // Row 1:   blank separator
        // Row 2+:  KPI section (one row per KPI: label | value | delta)
        // Row N+1: blank separator
        // Row N+2+: Data tables (title row + header row + data rows + blank)
        // Charts anchor below data

        const COL_COUNT = 6; // default width
        const requests: sheets_v4.Schema$Request[] = [];
        const valueData: { range: string; values: (string | number | boolean | null)[][] }[] = [];

        let currentRow = 0;

        // Helper to produce a sheet range string (0-indexed to A1)
        const rowToA1 = (row: number) => row + 1;
        const quotedSheet = quoteSheetName(targetSheet);
        const cellRef = (row: number, col: number) =>
          `${quotedSheet}!${columnToLetter(col)}${rowToA1(row)}`;
        const rangeRef = (startRow: number, startCol: number, endRow: number, endCol: number) =>
          `${quotedSheet}!${columnToLetter(startCol)}${rowToA1(startRow)}:${columnToLetter(endCol)}${rowToA1(endRow)}`;

        // ── Dashboard title ──────────────────────────────────────────────────
        if (title) {
          valueData.push({
            range: cellRef(currentRow, 0),
            values: [[title]],
          });

          // Title formatting - large bold white text on navy background
          requests.push(
            repeatCellRequest(
              sheetId,
              currentRow, 0, currentRow + 1, COL_COUNT,
              {
                backgroundColorStyle: { rgbColor: HEADER_BG },
                textFormat: {
                  bold: true,
                  fontSize: 18,
                  foregroundColorStyle: { rgbColor: HEADER_FG },
                  fontFamily: "Arial",
                },
                verticalAlignment: "MIDDLE",
                horizontalAlignment: "LEFT",
              },
              "userEnteredFormat.backgroundColorStyle,userEnteredFormat.textFormat,userEnteredFormat.verticalAlignment,userEnteredFormat.horizontalAlignment"
            )
          );

          // Merge title row across all columns
          requests.push({
            mergeCells: {
              range: makeGridRange(sheetId, currentRow, 0, currentRow + 1, COL_COUNT),
              mergeType: "MERGE_ALL",
            },
          });

          // Set title row height
          requests.push({
            updateDimensionProperties: {
              range: { sheetId, dimension: "ROWS", startIndex: currentRow, endIndex: currentRow + 1 },
              properties: { pixelSize: 48 },
              fields: "pixelSize",
            },
          });

          currentRow += 2; // title + blank separator
        }

        // ── KPI Section ──────────────────────────────────────────────────────
        if (kpis && kpis.length > 0) {
          // KPI section header
          valueData.push({
            range: cellRef(currentRow, 0),
            values: [["Key Performance Indicators"]],
          });

          requests.push(
            repeatCellRequest(
              sheetId,
              currentRow, 0, currentRow + 1, COL_COUNT,
              {
                backgroundColorStyle: { rgbColor: TABLE_HEADER_BG },
                textFormat: {
                  bold: true,
                  fontSize: 11,
                  foregroundColorStyle: { rgbColor: TABLE_HEADER_FG },
                  fontFamily: "Arial",
                },
              },
              "userEnteredFormat.backgroundColorStyle,userEnteredFormat.textFormat"
            )
          );

          currentRow++;

          // KPI column headers
          valueData.push({
            range: rangeRef(currentRow, 0, currentRow, 2),
            values: [["Metric", "Value", "Change"]],
          });

          requests.push(
            repeatCellRequest(
              sheetId,
              currentRow, 0, currentRow + 1, 3,
              {
                backgroundColorStyle: { rgbColor: KPI_BG },
                textFormat: {
                  bold: true,
                  fontSize: 10,
                  foregroundColorStyle: { rgbColor: KPI_LABEL_FG },
                },
                horizontalAlignment: "CENTER",
              },
              "userEnteredFormat.backgroundColorStyle,userEnteredFormat.textFormat,userEnteredFormat.horizontalAlignment"
            )
          );

          currentRow++;

          // KPI data rows
          const kpiStartRow = currentRow;
          for (let i = 0; i < kpis.length; i++) {
            const kpi = kpis[i];
            const row = currentRow + i;
            const bgColor = i % 2 === 0 ? null : ALT_ROW_BG;

            // Write label and value
            valueData.push({
              range: rangeRef(row, 0, row, 2),
              values: [[kpi.label, kpi.value, kpi.delta ?? ""]],
            });

            // Alternating row background
            if (bgColor) {
              requests.push(
                repeatCellRequest(
                  sheetId, row, 0, row + 1, 3,
                  { backgroundColorStyle: { rgbColor: bgColor } },
                  "userEnteredFormat.backgroundColorStyle"
                )
              );
            }

            // Value formatting - bold
            requests.push(
              repeatCellRequest(
                sheetId, row, 1, row + 1, 2,
                { textFormat: { bold: true, fontSize: 11 } },
                "userEnteredFormat.textFormat"
              )
            );

            // Delta conditional formatting - apply green for positive, red for negative
            // We use two conditional format rules directly
            if (kpi.delta) {
              const isPositive =
                kpi.delta.startsWith("+") ||
                (!kpi.delta.startsWith("-") && parseFloat(kpi.delta) > 0);

              const deltaColor = isPositive ? POSITIVE_FG : NEGATIVE_FG;
              requests.push(
                repeatCellRequest(
                  sheetId, row, 2, row + 1, 3,
                  {
                    textFormat: {
                      bold: true,
                      foregroundColorStyle: { rgbColor: deltaColor },
                    },
                  },
                  "userEnteredFormat.textFormat"
                )
              );
            }
          }

          currentRow += kpis.length;

          // Border around KPI block
          requests.push({
            updateBorders: {
              range: makeGridRange(sheetId, kpiStartRow - 1, 0, currentRow, 3),
              top: solidBorder(),
              bottom: solidBorder(),
              left: solidBorder(),
              right: solidBorder(),
              innerHorizontal: { style: "SOLID", colorStyle: { rgbColor: parseColor("#c5cae9") } },
            },
          });

          currentRow += 2; // blank separator after KPIs
        }

        // ── Data Tables ──────────────────────────────────────────────────────
        const tableChartAnchors: { row: number; dataRange: string; title?: string; type: string }[] = [];

        if (data_tables && data_tables.length > 0) {
          for (const table of data_tables) {
            const numCols = Math.max(table.headers.length, COL_COUNT);

            // Table title row
            valueData.push({
              range: cellRef(currentRow, 0),
              values: [[table.title]],
            });

            requests.push(
              repeatCellRequest(
                sheetId,
                currentRow, 0, currentRow + 1, numCols,
                {
                  backgroundColorStyle: { rgbColor: TABLE_HEADER_BG },
                  textFormat: {
                    bold: true,
                    fontSize: 11,
                    foregroundColorStyle: { rgbColor: TABLE_HEADER_FG },
                  },
                },
                "userEnteredFormat.backgroundColorStyle,userEnteredFormat.textFormat"
              )
            );

            requests.push({
              mergeCells: {
                range: makeGridRange(sheetId, currentRow, 0, currentRow + 1, numCols),
                mergeType: "MERGE_ALL",
              },
            });

            currentRow++;

            // Header row
            valueData.push({
              range: rangeRef(currentRow, 0, currentRow, table.headers.length - 1),
              values: [table.headers],
            });

            requests.push(
              repeatCellRequest(
                sheetId,
                currentRow, 0, currentRow + 1, table.headers.length,
                {
                  backgroundColorStyle: { rgbColor: KPI_BG },
                  textFormat: {
                    bold: true,
                    fontSize: 10,
                    foregroundColorStyle: { rgbColor: KPI_LABEL_FG },
                  },
                  horizontalAlignment: "CENTER",
                },
                "userEnteredFormat.backgroundColorStyle,userEnteredFormat.textFormat,userEnteredFormat.horizontalAlignment"
              )
            );

            currentRow++;

            // Data rows
            const dataStartRow = currentRow;
            const dataRangeForChart = rangeRef(
              currentRow - 1, // include header
              0,
              currentRow + table.rows.length - 1,
              table.headers.length - 1
            );

            if (table.rows.length > 0) {
              valueData.push({
                range: rangeRef(currentRow, 0, currentRow + table.rows.length - 1, table.headers.length - 1),
                values: table.rows,
              });

              // Alternating row colors
              for (let i = 0; i < table.rows.length; i++) {
                if (i % 2 === 1) {
                  requests.push(
                    repeatCellRequest(
                      sheetId,
                      dataStartRow + i, 0, dataStartRow + i + 1, table.headers.length,
                      { backgroundColorStyle: { rgbColor: ALT_ROW_BG } },
                      "userEnteredFormat.backgroundColorStyle"
                    )
                  );
                }
              }

              // Bottom border on data block
              requests.push({
                updateBorders: {
                  range: makeGridRange(
                    sheetId,
                    dataStartRow - 1, // header row
                    0,
                    dataStartRow + table.rows.length,
                    table.headers.length
                  ),
                  top: solidBorder(),
                  bottom: solidBorder(),
                  left: solidBorder(),
                  right: solidBorder(),
                  innerHorizontal: { style: "SOLID", colorStyle: { rgbColor: parseColor("#c5cae9") } },
                },
              });

              currentRow += table.rows.length;
            }

            // Store chart anchor info
            tableChartAnchors.push({
              row: currentRow + 1,
              dataRange: dataRangeForChart,
              title: table.title,
              type: "COLUMN",
            });

            currentRow += 2; // blank separator
          }
        }

        // ── Freeze header row (always freeze row 1) ─────────────────────────
        requests.push({
          updateSheetProperties: {
            properties: {
              sheetId,
              gridProperties: { frozenRowCount: 1 },
            },
            fields: "gridProperties.frozenRowCount",
          },
        });

        // ── Auto-resize columns ──────────────────────────────────────────────
        // Use the widest table's column count, not just COL_COUNT
        const maxCols = Math.max(
          COL_COUNT,
          ...(data_tables ?? []).map((t) => t.headers.length)
        );
        requests.push({
          autoResizeDimensions: {
            dimensions: {
              sheetId,
              dimension: "COLUMNS",
              startIndex: 0,
              endIndex: maxCols,
            },
          },
        });

        // ── Write values ─────────────────────────────────────────────────────
        if (valueData.length > 0) {
          await sheets.spreadsheets.values.batchUpdate({
            spreadsheetId: id,
            requestBody: {
              valueInputOption: "USER_ENTERED",
              data: valueData,
            },
          });
        }

        // ── Apply all formatting in one batchUpdate ───────────────────────────
        if (requests.length > 0) {
          await sheets.spreadsheets.batchUpdate({
            spreadsheetId: id,
            requestBody: { requests },
          });
        }

        // ── Create charts ─────────────────────────────────────────────────────
        const chartRequests: sheets_v4.Schema$Request[] = [];
        const inputCharts = charts ?? [];

        // Use provided charts first, then fall back to auto-generated from tables
        const allChartSpecs: { type: string; dataRange: string; title?: string; anchorRow: number }[] = [];

        if (inputCharts.length > 0) {
          inputCharts.forEach((c, i) => {
            allChartSpecs.push({
              type: c.type,
              dataRange: c.data_range,
              title: c.title,
              anchorRow: currentRow + i * 22, // space charts vertically
            });
          });
        } else if (tableChartAnchors.length > 0) {
          // Auto-generate one chart per data table
          tableChartAnchors.forEach((anchor, i) => {
            allChartSpecs.push({
              type: anchor.type,
              dataRange: anchor.dataRange,
              title: anchor.title,
              anchorRow: anchor.row + 1,
            });
          });
        }

        for (let i = 0; i < allChartSpecs.length; i++) {
          const spec = allChartSpecs[i];
          const chartSheetId = await resolveSheetIdFromRange(id, spec.dataRange);
          const fullRange = toGridRange(spec.dataRange, chartSheetId);

          // Split into domain (first column) and series (remaining columns)
          const domainStartCol = fullRange.startColumnIndex ?? 0;
          const domainRange: sheets_v4.Schema$GridRange = {
            ...fullRange,
            startColumnIndex: domainStartCol,
            endColumnIndex: domainStartCol + 1,
          };

          const colOffset = (i % 2) * 1;
          const rowOffset = Math.floor(i / 2) * 20;

          const position: sheets_v4.Schema$EmbeddedObjectPosition = {
            overlayPosition: {
              anchorCell: {
                sheetId,
                rowIndex: spec.anchorRow + rowOffset,
                columnIndex: colOffset * 4,
              },
              widthPixels: 480,
              heightPixels: 300,
            },
          };

          let chartSpecObj: sheets_v4.Schema$ChartSpec;

          if (spec.type === "PIE") {
            // PIE charts require pieChart spec, not basicChart
            const seriesRange: sheets_v4.Schema$GridRange = {
              ...fullRange,
              startColumnIndex: domainStartCol + 1,
              endColumnIndex: domainStartCol + 2, // PIE uses single series column
            };
            chartSpecObj = {
              title: spec.title ?? "",
              pieChart: {
                legendPosition: "BOTTOM_LEGEND",
                domain: { sourceRange: { sources: [domainRange] } },
                series: { sourceRange: { sources: [seriesRange] } },
                threeDimensional: false,
              },
            };
          } else {
            // Build per-column series entries
            const seriesStartCol = domainStartCol + 1;
            const seriesEndCol = fullRange.endColumnIndex ?? (seriesStartCol + 1);
            const seriesEntries: sheets_v4.Schema$BasicChartSeries[] = [];
            for (let col = seriesStartCol; col < seriesEndCol; col++) {
              const colRange: sheets_v4.Schema$GridRange = {
                ...fullRange,
                startColumnIndex: col,
                endColumnIndex: col + 1,
              };
              seriesEntries.push({
                series: { sourceRange: { sources: [colRange] } },
              });
            }
            if (seriesEntries.length === 0) {
              // Fallback: at least one series column
              seriesEntries.push({
                series: { sourceRange: { sources: [{ ...fullRange, startColumnIndex: seriesStartCol }] } },
              });
            }

            chartSpecObj = {
              title: spec.title ?? "",
              basicChart: {
                chartType: spec.type,
                legendPosition: "BOTTOM_LEGEND",
                domains: [{ domain: { sourceRange: { sources: [domainRange] } } }],
                series: seriesEntries,
                headerCount: 1,
              },
            };
          }

          chartRequests.push({
            addChart: {
              chart: { spec: chartSpecObj, position },
            },
          });
        }

        if (chartRequests.length > 0) {
          await sheets.spreadsheets.batchUpdate({
            spreadsheetId: id,
            requestBody: { requests: chartRequests },
          });
        }

        const summary: string[] = [
          `Dashboard built in sheet "${targetSheet}":`,
          `  Title: ${title ?? "(none)"}`,
          `  KPIs: ${kpis?.length ?? 0}`,
          `  Data tables: ${data_tables?.length ?? 0}`,
          `  Charts: ${chartRequests.length}`,
          `  Formatting requests: ${requests.length}`,
        ];

        return formatSuccess(summary.join("\n"));
      }
    )
  );
}

