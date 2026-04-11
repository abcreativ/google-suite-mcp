/**
 * Full-sheet declarative builder tool.
 *
 * sheets_build_sheet - build an entire sheet from a declarative spec:
 * title + ordered sections (tables, KPIs, text) + conditional formats + charts.
 * Everything fires in 2-3 API calls (values batch + format batch + chart batch).
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getSheetsClient, extractFileId } from "../client/google-client.js";
import { formatSuccess } from "../utils/response.js";
import { withErrorHandling } from "../utils/error-handler.js";
import { parseColor } from "../utils/color.js";
import { columnToLetter, quoteSheetName } from "../utils/range.js";
import { toGridRange } from "../utils/range.js";
import { resolveSheetIdFromRange } from "../utils/sheet-resolver.js";
import {
  resolveOrCreateSheet,
  makeGridRange,
  repeatCellRequest,
  DEFAULT_COLORS,
  solidBorder,
} from "../utils/sheet-builder.js";
import type { sheets_v4 } from "googleapis";
import type { CellValue } from "../types/sheets.js";

// ─── Section schemas ─────────────────────────────────────────────────────────

const TableSection = z.object({
  type: z.literal("table"),
  title: z.string().optional(),
  headers: z.array(z.string()).min(1),
  rows: z.array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()]))),
  column_formulas: z
    .array(z.object({
      column: z.number().int(),
      formula_template: z.string(),
    }))
    .optional(),
  column_formats: z
    .array(z.object({
      column: z.number().int(),
      format: z.string(),
    }))
    .optional(),
  header_style: z
    .object({
      bold: z.boolean().optional(),
      background_color: z.string().optional(),
      text_color: z.string().optional(),
    })
    .optional(),
  banded_rows: z.boolean().optional(),
  border: z.boolean().optional(),
});

const KpiSection = z.object({
  type: z.literal("kpis"),
  title: z.string().optional(),
  items: z.array(z.object({
    label: z.string(),
    value: z.union([z.string(), z.number()]),
    delta: z.string().optional(),
  })),
});

const TextSection = z.object({
  type: z.literal("text"),
  content: z.string(),
  style: z
    .object({
      bold: z.boolean().optional(),
      font_size: z.number().optional(),
      background_color: z.string().optional(),
      text_color: z.string().optional(),
    })
    .optional(),
});

const SectionSchema = z.union([TableSection, KpiSection, TextSection]);

const ConditionalFormatRule = z.object({
  range: z.string().describe("A1 range"),
  rule_type: z.enum(["condition", "color_scale", "formula"]),
  condition_type: z.string().optional(),
  condition_values: z.array(z.string()).optional(),
  format_background_color: z.string().optional(),
  format_text_color: z.string().optional(),
  format_bold: z.boolean().optional(),
  min_color: z.string().optional(),
  mid_color: z.string().optional(),
  max_color: z.string().optional(),
  formula: z.string().optional(),
});

const ChartSpec = z.object({
  type: z.enum(["BAR", "LINE", "PIE", "SCATTER", "AREA", "COLUMN"]),
  data_range: z.string().describe("A1 range for chart data"),
  title: z.string().optional(),
});

// ─── Tool registration ────────────────────────────────────────────────────────

export function registerSheetBuilderTools(server: McpServer): void {
  server.tool(
    "sheets_build_sheet",
    "Build an entire sheet from a declarative spec: title, sections (tables, KPIs, text), conditional formats, charts - all in one call. For building from scratch.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      sheet_name: z.string().optional().describe("Sheet name (created if not exists)"),

      title: z.string().optional().describe("Sheet title (merged row at top)"),
      title_style: z
        .object({
          background_color: z.string().optional(),
          text_color: z.string().optional(),
          font_size: z.number().optional(),
        })
        .optional(),

      sections: z
        .array(SectionSchema)
        .describe("Ordered sections to lay out top-to-bottom"),

      conditional_formats: z.array(ConditionalFormatRule).optional(),
      charts: z.array(ChartSpec).optional(),
      freeze_rows: z.number().optional().describe("Rows to freeze (default: 1)"),
      auto_resize: z.boolean().optional().describe("Auto-fit columns (default: true)"),
    },
    withErrorHandling(
      async ({
        spreadsheet_id,
        sheet_name,
        title,
        title_style,
        sections,
        conditional_formats,
        charts,
        freeze_rows,
        auto_resize,
      }) => {
        const sheetsClient = await getSheetsClient();
        const id = extractFileId(spreadsheet_id);
        const targetSheet = sheet_name ?? "Sheet1";
        const sheetId = await resolveOrCreateSheet(sheetsClient, id, targetSheet);
        const quotedSheet = quoteSheetName(targetSheet);

        const requests: sheets_v4.Schema$Request[] = [];
        const valueData: { range: string; values: CellValue[][] }[] = [];

        let currentRow = 0;
        // Pre-scan sections to find widest table for title/merge width
        let maxCols = Math.max(
          6,
          ...sections.map((s) => (s.type === "table" ? s.headers.length : 0))
        );

        // Helpers
        const cellRef = (row: number, col: number) =>
          `${quotedSheet}!${columnToLetter(col)}${row + 1}`;
        const rangeRef = (sRow: number, sCol: number, eRow: number, eCol: number) =>
          `${quotedSheet}!${columnToLetter(sCol)}${sRow + 1}:${columnToLetter(eCol)}${eRow + 1}`;

        // ── Title ─────────────────────────────────────────────────────────
        if (title) {
          valueData.push({ range: cellRef(currentRow, 0), values: [[title]] });

          const tStyle = title_style ?? {};
          const titleBg = tStyle.background_color
            ? parseColor(tStyle.background_color)
            : DEFAULT_COLORS.HEADER_BG;
          const titleFg = tStyle.text_color
            ? parseColor(tStyle.text_color)
            : DEFAULT_COLORS.HEADER_FG;

          requests.push(
            repeatCellRequest(
              sheetId,
              currentRow, 0, currentRow + 1, maxCols,
              {
                backgroundColorStyle: { rgbColor: titleBg },
                textFormat: {
                  bold: true,
                  fontSize: tStyle.font_size ?? 18,
                  foregroundColorStyle: { rgbColor: titleFg },
                  fontFamily: "Arial",
                },
                verticalAlignment: "MIDDLE",
                horizontalAlignment: "LEFT",
              },
              "userEnteredFormat.backgroundColorStyle,userEnteredFormat.textFormat,userEnteredFormat.verticalAlignment,userEnteredFormat.horizontalAlignment"
            )
          );

          requests.push({
            mergeCells: {
              range: makeGridRange(sheetId, currentRow, 0, currentRow + 1, maxCols),
              mergeType: "MERGE_ALL",
            },
          });

          requests.push({
            updateDimensionProperties: {
              range: { sheetId, dimension: "ROWS", startIndex: currentRow, endIndex: currentRow + 1 },
              properties: { pixelSize: 48 },
              fields: "pixelSize",
            },
          });

          currentRow += 2; // title + spacer
        }

        // ── Sections ──────────────────────────────────────────────────────
        for (const section of sections) {
          if (section.type === "text") {
            // Text section
            valueData.push({ range: cellRef(currentRow, 0), values: [[section.content]] });

            if (section.style) {
              const s = section.style;
              const cellFormat: sheets_v4.Schema$CellFormat = {};
              const fields: string[] = [];
              const textFormat: sheets_v4.Schema$TextFormat = {};

              if (s.bold) { textFormat.bold = true; fields.push("userEnteredFormat.textFormat.bold"); }
              if (s.font_size) { textFormat.fontSize = s.font_size; fields.push("userEnteredFormat.textFormat.fontSize"); }
              if (s.text_color) {
                textFormat.foregroundColorStyle = { rgbColor: parseColor(s.text_color) };
                fields.push("userEnteredFormat.textFormat.foregroundColorStyle");
              }
              if (Object.keys(textFormat).length > 0) cellFormat.textFormat = textFormat;
              if (s.background_color) {
                cellFormat.backgroundColorStyle = { rgbColor: parseColor(s.background_color) };
                fields.push("userEnteredFormat.backgroundColorStyle");
              }

              if (fields.length > 0) {
                requests.push(
                  repeatCellRequest(sheetId, currentRow, 0, currentRow + 1, maxCols, cellFormat, fields.join(","))
                );
              }
            }

            currentRow += 2; // text + spacer
          } else if (section.type === "kpis") {
            // KPI section
            if (section.title) {
              valueData.push({ range: cellRef(currentRow, 0), values: [[section.title]] });
              requests.push(
                repeatCellRequest(
                  sheetId, currentRow, 0, currentRow + 1, maxCols,
                  {
                    backgroundColorStyle: { rgbColor: DEFAULT_COLORS.TABLE_HEADER_BG },
                    textFormat: {
                      bold: true, fontSize: 11,
                      foregroundColorStyle: { rgbColor: DEFAULT_COLORS.TABLE_HEADER_FG },
                    },
                  },
                  "userEnteredFormat.backgroundColorStyle,userEnteredFormat.textFormat"
                )
              );
              currentRow++;
            }

            // KPI column headers
            valueData.push({ range: rangeRef(currentRow, 0, currentRow, 2), values: [["Metric", "Value", "Change"]] });
            requests.push(
              repeatCellRequest(
                sheetId, currentRow, 0, currentRow + 1, 3,
                {
                  backgroundColorStyle: { rgbColor: DEFAULT_COLORS.KPI_BG },
                  textFormat: { bold: true, fontSize: 10, foregroundColorStyle: { rgbColor: DEFAULT_COLORS.KPI_LABEL_FG } },
                  horizontalAlignment: "CENTER",
                },
                "userEnteredFormat.backgroundColorStyle,userEnteredFormat.textFormat,userEnteredFormat.horizontalAlignment"
              )
            );
            currentRow++;

            for (let i = 0; i < section.items.length; i++) {
              const kpi = section.items[i];
              const row = currentRow + i;

              valueData.push({
                range: rangeRef(row, 0, row, 2),
                values: [[kpi.label, kpi.value, kpi.delta ?? ""]],
              });

              if (i % 2 === 1) {
                requests.push(
                  repeatCellRequest(
                    sheetId, row, 0, row + 1, 3,
                    { backgroundColorStyle: { rgbColor: DEFAULT_COLORS.ALT_ROW_BG } },
                    "userEnteredFormat.backgroundColorStyle"
                  )
                );
              }

              requests.push(
                repeatCellRequest(
                  sheetId, row, 1, row + 1, 2,
                  { textFormat: { bold: true, fontSize: 11 } },
                  "userEnteredFormat.textFormat"
                )
              );

              if (kpi.delta) {
                const isPositive = kpi.delta.startsWith("+") || (!kpi.delta.startsWith("-") && parseFloat(kpi.delta) > 0);
                requests.push(
                  repeatCellRequest(
                    sheetId, row, 2, row + 1, 3,
                    { textFormat: { bold: true, foregroundColorStyle: { rgbColor: isPositive ? DEFAULT_COLORS.POSITIVE_FG : DEFAULT_COLORS.NEGATIVE_FG } } },
                    "userEnteredFormat.textFormat"
                  )
                );
              }
            }

            currentRow += section.items.length + 2; // data + spacer

          } else if (section.type === "table") {
            const numCols = section.headers.length;
            maxCols = Math.max(maxCols, numCols);

            // Table title
            if (section.title) {
              valueData.push({ range: cellRef(currentRow, 0), values: [[section.title]] });
              requests.push(
                repeatCellRequest(
                  sheetId, currentRow, 0, currentRow + 1, numCols,
                  {
                    backgroundColorStyle: { rgbColor: DEFAULT_COLORS.TABLE_HEADER_BG },
                    textFormat: {
                      bold: true, fontSize: 11,
                      foregroundColorStyle: { rgbColor: DEFAULT_COLORS.TABLE_HEADER_FG },
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
            }

            // Header row
            valueData.push({
              range: rangeRef(currentRow, 0, currentRow, numCols - 1),
              values: [section.headers],
            });

            const hStyle = section.header_style ?? {};
            const headerBg = hStyle.background_color
              ? parseColor(hStyle.background_color)
              : DEFAULT_COLORS.KPI_BG;
            const headerFg = hStyle.text_color
              ? parseColor(hStyle.text_color)
              : DEFAULT_COLORS.KPI_LABEL_FG;

            requests.push(
              repeatCellRequest(
                sheetId, currentRow, 0, currentRow + 1, numCols,
                {
                  backgroundColorStyle: { rgbColor: headerBg },
                  textFormat: {
                    bold: hStyle.bold ?? true,
                    fontSize: 10,
                    foregroundColorStyle: { rgbColor: headerFg },
                  },
                  horizontalAlignment: "CENTER",
                },
                "userEnteredFormat.backgroundColorStyle,userEnteredFormat.textFormat,userEnteredFormat.horizontalAlignment"
              )
            );

            currentRow++;
            const dataStartRow = currentRow;

            // Data rows (with formula expansion)
            if (section.rows.length > 0) {
              const expandedRows: CellValue[][] = [];
              for (let i = 0; i < section.rows.length; i++) {
                const row = [...section.rows[i]];
                if (section.column_formulas) {
                  const actualRow = dataStartRow + i + 1; // 1-based
                  for (const cf of section.column_formulas) {
                    if (cf.column >= 0 && cf.column < numCols) {
                      row[cf.column] = cf.formula_template.replace(/\{row\}/g, String(actualRow));
                    }
                  }
                }
                expandedRows.push(row);
              }

              valueData.push({
                range: rangeRef(dataStartRow, 0, dataStartRow + section.rows.length - 1, numCols - 1),
                values: expandedRows,
              });

              // Banded rows via native addBanding (default: true)
              if (section.banded_rows !== false) {
                requests.push({
                  addBanding: {
                    bandedRange: {
                      range: makeGridRange(
                        sheetId,
                        dataStartRow - 1, 0,  // include header row
                        dataStartRow + section.rows.length, numCols
                      ),
                      rowProperties: {
                        headerColorStyle: { rgbColor: headerBg },
                        firstBandColorStyle: { rgbColor: { red: 1, green: 1, blue: 1 } },
                        secondBandColorStyle: { rgbColor: DEFAULT_COLORS.ALT_ROW_BG },
                      },
                    },
                  },
                });
              }

              // Column number formats
              if (section.column_formats) {
                for (const cf of section.column_formats) {
                  requests.push(
                    repeatCellRequest(
                      sheetId,
                      dataStartRow, cf.column,
                      dataStartRow + section.rows.length, cf.column + 1,
                      { numberFormat: { type: "NUMBER", pattern: cf.format } },
                      "userEnteredFormat.numberFormat"
                    )
                  );
                }
              }

              // Borders (default: true)
              if (section.border !== false) {
                const borderObj = solidBorder();
                const innerBorder: sheets_v4.Schema$Border = {
                  style: "SOLID",
                  colorStyle: { rgbColor: parseColor("#c5cae9") },
                };
                requests.push({
                  updateBorders: {
                    range: makeGridRange(
                      sheetId,
                      dataStartRow - 1, 0,
                      dataStartRow + section.rows.length, numCols
                    ),
                    top: borderObj,
                    bottom: borderObj,
                    left: borderObj,
                    right: borderObj,
                    innerHorizontal: innerBorder,
                  },
                });
              }

              currentRow += section.rows.length;
            }

            currentRow += 2; // spacer
          }
        }

        // ── Freeze ────────────────────────────────────────────────────────
        const frozenRows = freeze_rows ?? 1;
        if (frozenRows > 0) {
          requests.push({
            updateSheetProperties: {
              properties: {
                sheetId,
                gridProperties: { frozenRowCount: frozenRows },
              },
              fields: "gridProperties.frozenRowCount",
            },
          });
        }

        // ── Auto-resize ───────────────────────────────────────────────────
        if (auto_resize !== false) {
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
        }

        // ── Write all values ──────────────────────────────────────────────
        if (valueData.length > 0) {
          await sheetsClient.spreadsheets.values.batchUpdate({
            spreadsheetId: id,
            requestBody: {
              valueInputOption: "USER_ENTERED",
              data: valueData,
            },
          });
        }

        // ── Apply all formatting ──────────────────────────────────────────
        if (requests.length > 0) {
          await sheetsClient.spreadsheets.batchUpdate({
            spreadsheetId: id,
            requestBody: { requests },
          });
        }

        // ── Conditional formats ───────────────────────────────────────────
        if (conditional_formats && conditional_formats.length > 0) {
          const cfRequests: sheets_v4.Schema$Request[] = [];

          for (const cf of conditional_formats) {
            // Use target sheet's sheetId for unqualified ranges (no sheet name prefix)
            const parsed = cf.range.includes("!") ? null : true;
            const cfSheetId = parsed ? sheetId : await resolveSheetIdFromRange(id, cf.range);
            const gridRange = toGridRange(cf.range, cfSheetId);

            if (cf.rule_type === "color_scale") {
              const gradientRule: sheets_v4.Schema$GradientRule = {
                minpoint: {
                  colorStyle: { rgbColor: parseColor(cf.min_color ?? "#ffffff") },
                  type: "MIN" as const,
                },
                maxpoint: {
                  colorStyle: { rgbColor: parseColor(cf.max_color ?? "#ff0000") },
                  type: "MAX" as const,
                },
              };
              if (cf.mid_color) {
                gradientRule.midpoint = {
                  colorStyle: { rgbColor: parseColor(cf.mid_color) },
                  type: "PERCENTILE" as const,
                  value: "50",
                };
              }
              cfRequests.push({
                addConditionalFormatRule: {
                  rule: { ranges: [gridRange], gradientRule },
                  index: 0,
                },
              });
            } else {
              const cellFormat: sheets_v4.Schema$CellFormat = {};
              const textFormat: sheets_v4.Schema$TextFormat = {};

              if (cf.format_background_color) {
                cellFormat.backgroundColorStyle = { rgbColor: parseColor(cf.format_background_color) };
              }
              if (cf.format_text_color) {
                textFormat.foregroundColorStyle = { rgbColor: parseColor(cf.format_text_color) };
              }
              if (cf.format_bold !== undefined) textFormat.bold = cf.format_bold;
              if (Object.keys(textFormat).length > 0) cellFormat.textFormat = textFormat;

              let condition: sheets_v4.Schema$BooleanCondition;
              if (cf.rule_type === "formula" && cf.formula) {
                condition = { type: "CUSTOM_FORMULA", values: [{ userEnteredValue: cf.formula }] };
              } else if (cf.condition_type) {
                condition = {
                  type: cf.condition_type,
                  values: (cf.condition_values ?? []).map((v) => ({ userEnteredValue: v })),
                };
              } else {
                throw new Error(
                  "Conditional format rule requires condition_type (for rule_type='condition') or formula (for rule_type='formula')."
                );
              }

              cfRequests.push({
                addConditionalFormatRule: {
                  rule: {
                    ranges: [gridRange],
                    booleanRule: { condition, format: cellFormat },
                  },
                  index: 0,
                },
              });
            }
          }

          if (cfRequests.length > 0) {
            await sheetsClient.spreadsheets.batchUpdate({
              spreadsheetId: id,
              requestBody: { requests: cfRequests },
            });
          }
        }

        // ── Charts ────────────────────────────────────────────────────────
        if (charts && charts.length > 0) {
          const chartRequests: sheets_v4.Schema$Request[] = [];

          for (let i = 0; i < charts.length; i++) {
            const chart = charts[i];
            // Use target sheet's sheetId for unqualified ranges
            const chartSheetId = chart.data_range.includes("!")
              ? await resolveSheetIdFromRange(id, chart.data_range)
              : sheetId;
            const fullRange = toGridRange(chart.data_range, chartSheetId);

            const domainStartCol = fullRange.startColumnIndex ?? 0;
            const domainRange: sheets_v4.Schema$GridRange = {
              ...fullRange,
              startColumnIndex: domainStartCol,
              endColumnIndex: domainStartCol + 1,
            };

            const position: sheets_v4.Schema$EmbeddedObjectPosition = {
              overlayPosition: {
                anchorCell: {
                  sheetId,
                  rowIndex: currentRow + Math.floor(i / 2) * 20,
                  columnIndex: (i % 2) * 4,
                },
                widthPixels: 480,
                heightPixels: 300,
              },
            };

            let chartSpec: sheets_v4.Schema$ChartSpec;

            if (chart.type === "PIE") {
              const seriesRange: sheets_v4.Schema$GridRange = {
                ...fullRange,
                startColumnIndex: domainStartCol + 1,
                endColumnIndex: domainStartCol + 2,
              };
              chartSpec = {
                title: chart.title ?? "",
                pieChart: {
                  legendPosition: "BOTTOM_LEGEND",
                  domain: { sourceRange: { sources: [domainRange] } },
                  series: { sourceRange: { sources: [seriesRange] } },
                  threeDimensional: false,
                },
              };
            } else {
              const seriesStartCol = domainStartCol + 1;
              const seriesEndCol = fullRange.endColumnIndex ?? (seriesStartCol + 1);
              const seriesEntries: sheets_v4.Schema$BasicChartSeries[] = [];
              for (let col = seriesStartCol; col < seriesEndCol; col++) {
                seriesEntries.push({
                  series: {
                    sourceRange: {
                      sources: [{ ...fullRange, startColumnIndex: col, endColumnIndex: col + 1 }],
                    },
                  },
                });
              }
              if (seriesEntries.length === 0) {
                seriesEntries.push({
                  series: { sourceRange: { sources: [{ ...fullRange, startColumnIndex: seriesStartCol }] } },
                });
              }
              chartSpec = {
                title: chart.title ?? "",
                basicChart: {
                  chartType: chart.type,
                  legendPosition: "BOTTOM_LEGEND",
                  domains: [{ domain: { sourceRange: { sources: [domainRange] } } }],
                  series: seriesEntries,
                  headerCount: 1,
                },
              };
            }

            chartRequests.push({
              addChart: { chart: { spec: chartSpec, position } },
            });
          }

          if (chartRequests.length > 0) {
            await sheetsClient.spreadsheets.batchUpdate({
              spreadsheetId: id,
              requestBody: { requests: chartRequests },
            });
          }
        }

        // ── Summary ───────────────────────────────────────────────────────
        const sectionCounts = {
          tables: sections.filter((s) => s.type === "table").length,
          kpis: sections.filter((s) => s.type === "kpis").length,
          text: sections.filter((s) => s.type === "text").length,
        };

        return formatSuccess(
          `Sheet built in "${targetSheet}": ${sections.length} section(s) (${sectionCounts.tables} table, ${sectionCounts.kpis} KPI, ${sectionCounts.text} text), ${conditional_formats?.length ?? 0} conditional format(s), ${charts?.length ?? 0} chart(s)`
        );
      }
    )
  );
}
