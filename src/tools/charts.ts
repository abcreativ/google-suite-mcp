/**
 * Chart tools: create and delete embedded charts.
 *
 * F4.4 - sheets_create_chart, sheets_delete_chart
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getSheetsClient, extractFileId } from "../client/google-client.js";
import { formatSuccess } from "../utils/response.js";
import { withErrorHandling } from "../utils/error-handler.js";
import { toGridRange, parseA1Notation, letterToColumn } from "../utils/range.js";
import { parseColor } from "../utils/color.js";
import { resolveSheetId, resolveSheetIdFromRange } from "../utils/sheet-resolver.js";
import type { sheets_v4 } from "googleapis";

// ─── Internal helpers ─────────────────────────────────────────────────────────

/**
 * Split a GridRange into a domain range (first column) and series range
 * (remaining columns). This is how Google Sheets charts expect data:
 * first column = X-axis labels, remaining columns = Y-axis data.
 */
function splitDomainSeries(
  gridRange: sheets_v4.Schema$GridRange
): { domain: sheets_v4.Schema$GridRange; series: sheets_v4.Schema$GridRange } {
  const startCol = gridRange.startColumnIndex ?? 0;
  const endCol = gridRange.endColumnIndex;

  const domain: sheets_v4.Schema$GridRange = {
    ...gridRange,
    startColumnIndex: startCol,
    endColumnIndex: startCol + 1,
  };

  const series: sheets_v4.Schema$GridRange = {
    ...gridRange,
    startColumnIndex: startCol + 1,
    endColumnIndex: endCol,
  };

  return { domain, series };
}

// ─── Chart type mapping ───────────────────────────────────────────────────────

type ChartTypeInput = "BAR" | "LINE" | "PIE" | "SCATTER" | "AREA" | "COLUMN" | "COMBO";

function mapChartType(type: ChartTypeInput): string {
  const map: Record<ChartTypeInput, string> = {
    BAR: "BAR",
    LINE: "LINE",
    PIE: "PIE",
    SCATTER: "SCATTER",
    AREA: "AREA",
    COLUMN: "COLUMN",
    COMBO: "COMBO",
  };
  return map[type];
}

// ─── Series schema ────────────────────────────────────────────────────────────

const SeriesSpec = z.object({
  data_range: z.string().describe("A1 notation range for this series data"),
  series_type: z
    .enum(["BAR", "LINE", "AREA", "COLUMN"])
    .optional()
    .describe("Series chart type (for COMBO charts)"),
  color: z.string().optional().describe("Series color (hex or named)"),
  label: z.string().optional().describe("Series label"),
});

// ─── Tool registration ────────────────────────────────────────────────────────

export function registerChartTools(server: McpServer): void {
  // ─── F4.4 sheets_create_chart ─────────────────────────────────────────────

  server.tool(
    "sheets_create_chart",
    "Creates an embedded chart in a Google Sheet using spreadsheets.batchUpdate with addChart; the first column of the data range is used as the X-axis domain and remaining columns as series. Use when the user asks to visualize data in a sheet as a bar, line, pie, scatter, area, column, or combo chart. Use when adding a chart to a dashboard sheet after building it with sheets_build_dashboard. Do not use when: removing a chart - use sheets_delete_chart with the returned chartId; building a full multi-chart dashboard - use sheets_build_dashboard instead. Returns: 'Created {chart_type} chart (chartId: {chartId}) in sheet \"{sheet}\"'. The returned chartId is required by sheets_delete_chart. Parameters: - chart_type: one of BAR, LINE, PIE, SCATTER, AREA, COLUMN, COMBO - data_range: A1 notation range where the first column is X-axis labels, e.g. 'Sheet1!A1:C20' - sheet: tab name or ID where the chart will be embedded.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      sheet: z.string().describe("Sheet name or ID for chart placement"),
      chart_type: z
        .enum(["BAR", "LINE", "PIE", "SCATTER", "AREA", "COLUMN", "COMBO"])
        .describe("Chart type"),
      data_range: z
        .string()
        .describe("Data range in A1 notation (first col = X-axis)"),
      title: z.string().optional().describe("Chart title"),
      x_axis_label: z.string().optional().describe("X-axis label"),
      y_axis_label: z.string().optional().describe("Y-axis label"),
      legend_position: z
        .enum(["BOTTOM_LEGEND", "LEFT_LEGEND", "RIGHT_LEGEND", "TOP_LEGEND", "NO_LEGEND"])
        .optional()
        .describe("Legend position (default: BOTTOM_LEGEND)"),
      anchor_row: z.number().int().min(0).optional().describe("0-based row index for chart anchor (default: 0)"),
      anchor_col: z.number().int().min(0).optional().describe("0-based column index for chart anchor (default: 0)"),
      offset_x: z.number().int().optional().describe("X offset in pixels from anchor cell"),
      offset_y: z.number().int().optional().describe("Y offset in pixels from anchor cell"),
      width: z.number().int().optional().describe("Chart width in pixels (default: 600)"),
      height: z.number().int().optional().describe("Chart height in pixels (default: 400)"),
      headers_in_first_row: z
        .boolean()
        .optional()
        .describe("First row is headers (default: true)"),
      series: z
        .array(SeriesSpec)
        .optional()
        .describe("Override series config (for COMBO charts)"),
      series_colors: z
        .array(z.string())
        .optional()
        .describe("Colors per series (hex or named)"),
      stacked: z
        .boolean()
        .optional()
        .describe("Stack series (BAR/COLUMN/AREA)"),
    },
    withErrorHandling(
      async ({
        spreadsheet_id,
        sheet,
        chart_type,
        data_range,
        title,
        x_axis_label,
        y_axis_label,
        legend_position,
        anchor_row,
        anchor_col,
        offset_x,
        offset_y,
        width,
        height,
        headers_in_first_row,
        series,
        series_colors,
        stacked,
      }) => {
        const sheets = await getSheetsClient();
        const id = extractFileId(spreadsheet_id);
        const anchorSheetId = await resolveSheetId(id, sheet);
        const dataSheetId = await resolveSheetIdFromRange(id, data_range);
        const dataGridRange = toGridRange(data_range, dataSheetId);

        // Position / overlay spec
        const overlayPosition: sheets_v4.Schema$OverlayPosition = {
          anchorCell: {
            sheetId: anchorSheetId,
            rowIndex: anchor_row ?? 0,
            columnIndex: anchor_col ?? 0,
          },
          offsetXPixels: offset_x ?? 0,
          offsetYPixels: offset_y ?? 0,
          widthPixels: width ?? 600,
          heightPixels: height ?? 400,
        };

        // Build spec based on chart type
        let chartSpec: sheets_v4.Schema$ChartSpec;

        const chartTitle = title ?? "";

        // Split the data range into domain (first column) and series (remaining)
        const { domain: domainRange, series: seriesGridRange } = splitDomainSeries(dataGridRange);

        if (chart_type === "PIE") {
          // PIE charts accept exactly one series column
          const pieSeriesRange: sheets_v4.Schema$GridRange = {
            ...seriesGridRange,
            endColumnIndex: (seriesGridRange.startColumnIndex ?? 1) + 1,
          };
          chartSpec = {
            title: chartTitle,
            pieChart: {
              legendPosition: legend_position ?? "BOTTOM_LEGEND",
              domain: {
                sourceRange: {
                  sources: [domainRange],
                },
              },
              series: {
                sourceRange: {
                  sources: [pieSeriesRange],
                },
              },
              threeDimensional: false,
            },
          };
        } else {
          // Build BasicChartSpec for all non-pie types
          const chartTypeStr = mapChartType(chart_type);

          const domains: sheets_v4.Schema$BasicChartDomain[] = [
            {
              domain: {
                sourceRange: {
                  sources: [domainRange],
                },
              },
            },
          ];

          // Build series list
          const chartSeries: sheets_v4.Schema$BasicChartSeries[] = [];

          if (series && series.length > 0) {
            // Explicit series configuration
            for (let i = 0; i < series.length; i++) {
              const s = series[i];
              const seriesSheetId = await resolveSheetIdFromRange(id, s.data_range);
              const seriesRange = toGridRange(s.data_range, seriesSheetId);
              const seriesEntry: sheets_v4.Schema$BasicChartSeries = {
                series: {
                  sourceRange: {
                    sources: [seriesRange],
                  },
                },
              };
              if (s.series_type) seriesEntry.type = s.series_type;
              if (s.color || (series_colors && series_colors[i])) {
                const colorStr = s.color ?? series_colors![i];
                seriesEntry.colorStyle = { rgbColor: parseColor(colorStr) };
              }
              chartSeries.push(seriesEntry);
            }
          } else {
            // Default: emit one series per remaining column for correct multi-series rendering
            const seriesStartCol = (seriesGridRange.startColumnIndex ?? 0);
            const seriesEndCol = seriesGridRange.endColumnIndex ?? undefined;
            if (seriesEndCol != null && seriesEndCol > seriesStartCol) {
              for (let col = seriesStartCol; col < seriesEndCol; col++) {
                const colRange: sheets_v4.Schema$GridRange = {
                  ...seriesGridRange,
                  startColumnIndex: col,
                  endColumnIndex: col + 1,
                };
                const seriesEntry: sheets_v4.Schema$BasicChartSeries = {
                  series: { sourceRange: { sources: [colRange] } },
                };
                const colorIdx = col - seriesStartCol;
                if (series_colors && series_colors[colorIdx]) {
                  seriesEntry.colorStyle = { rgbColor: parseColor(series_colors[colorIdx]) };
                }
                chartSeries.push(seriesEntry);
              }
            } else {
              // Single remaining column or open-ended
              const seriesEntry: sheets_v4.Schema$BasicChartSeries = {
                series: { sourceRange: { sources: [seriesGridRange] } },
              };
              if (series_colors && series_colors.length > 0) {
                seriesEntry.colorStyle = { rgbColor: parseColor(series_colors[0]) };
              }
              chartSeries.push(seriesEntry);
            }
          }

          // Build axes
          const axes: sheets_v4.Schema$BasicChartAxis[] = [];
          if (x_axis_label) {
            axes.push({ position: "BOTTOM_AXIS", title: x_axis_label });
          }
          if (y_axis_label) {
            axes.push({ position: "LEFT_AXIS", title: y_axis_label });
          }

          const basicChart: sheets_v4.Schema$BasicChartSpec = {
            chartType: chartTypeStr,
            legendPosition: legend_position ?? "BOTTOM_LEGEND",
            domains,
            series: chartSeries,
            headerCount: (headers_in_first_row ?? true) ? 1 : 0,
          };

          if (axes.length > 0) basicChart.axis = axes;

          if (stacked) {
            basicChart.stackedType = "STACKED";
          }

          chartSpec = {
            title: chartTitle,
            basicChart,
          };
        }

        const res = await sheets.spreadsheets.batchUpdate({
          spreadsheetId: id,
          requestBody: {
            requests: [
              {
                addChart: {
                  chart: {
                    spec: chartSpec,
                    position: {
                      overlayPosition,
                    },
                  },
                },
              },
            ],
          },
        });

        const addedChart = res.data.replies?.[0]?.addChart?.chart;
        const chartId = addedChart?.chartId;

        return formatSuccess(
          `Created ${chart_type} chart${chartId !== undefined ? ` (chartId: ${chartId})` : ""} in sheet "${sheet}"`
        );
      }
    )
  );

  // ─── F4.4 sheets_delete_chart ─────────────────────────────────────────────

  server.tool(
    "sheets_delete_chart",
    "Deletes an embedded chart from a spreadsheet using spreadsheets.batchUpdate with deleteEmbeddedObject; the chart is removed but the underlying data range is not affected. Use when the user asks to remove a chart that is no longer needed. Use when replacing a chart by deleting the old one and calling sheets_create_chart with updated parameters. Do not use when: creating a chart - use sheets_create_chart instead; building a full dashboard - use sheets_build_dashboard instead. Returns: 'Deleted chart {chart_id}'. Parameters: - chart_id: numeric chart ID returned by sheets_create_chart; obtain from the sheet if the ID is not known by inspecting the spreadsheet JSON.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      chart_id: z.number().int().describe("Chart ID to delete (returned by sheets_create_chart)"),
    },
    withErrorHandling(async ({ spreadsheet_id, chart_id }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: {
          requests: [
            {
              deleteEmbeddedObject: {
                objectId: chart_id,
              },
            },
          ],
        },
      });

      return formatSuccess(`Deleted chart ${chart_id}`);
    })
  );
}
