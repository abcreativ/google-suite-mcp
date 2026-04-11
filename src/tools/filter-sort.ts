/**
 * Filter and sort tools: set basic filters, sort ranges.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getSheetsClient, extractFileId } from "../client/google-client.js";
import { formatSuccess } from "../utils/response.js";
import { withErrorHandling } from "../utils/error-handler.js";
import { toGridRange } from "../utils/range.js";
import { resolveSheetIdFromRange } from "../utils/sheet-resolver.js";

export function registerFilterSortTools(server: McpServer): void {
  // ─── sheets_set_filter ────────────────────────────────────────────────────────

  server.tool(
    "sheets_set_filter",
    "Set auto-filter on a range with optional per-column criteria.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      range: z.string().describe("A1 notation range"),
      sheet_id: z.number().int().optional().describe("Numeric sheet ID (default: 0)"),
      criteria: z
        .array(
          z.object({
            column_index: z.number().int().describe("0-based column index within the filter range"),
            hidden_values: z
              .array(z.string())
              .optional()
              .describe("Values to hide (i.e. rows with these values in this column are hidden)"),
            condition_type: z
              .string()
              .optional()
              .describe("Condition type, e.g. TEXT_CONTAINS, NUMBER_GREATER, BLANK"),
            condition_values: z
              .array(z.string())
              .optional()
              .describe("Values for the condition"),
          })
        )
        .optional()
        .describe("Per-column filter criteria. Omit to set an empty filter (all rows visible)."),
    },
    withErrorHandling(async ({ spreadsheet_id, range, sheet_id, criteria }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);
      const resolvedSheetId = sheet_id ?? await resolveSheetIdFromRange(id, range);
      const gridRange = toGridRange(range, resolvedSheetId);

      // Build filterCriteria map keyed by column index
      const filterCriteria: Record<
        number,
        {
          hiddenValues?: string[];
          condition?: { type: string; values?: Array<{ userEnteredValue: string }> };
        }
      > = {};

      if (criteria && criteria.length > 0) {
        for (const c of criteria) {
          const entry: (typeof filterCriteria)[number] = {};
          if (c.hidden_values && c.hidden_values.length > 0) {
            entry.hiddenValues = c.hidden_values;
          }
          if (c.condition_type) {
            entry.condition = {
              type: c.condition_type,
              values: (c.condition_values ?? []).map((v) => ({ userEnteredValue: v })),
            };
          }
          filterCriteria[c.column_index] = entry;
        }
      }

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: {
          requests: [
            {
              setBasicFilter: {
                filter: {
                  range: gridRange,
                  ...(Object.keys(filterCriteria).length > 0
                    ? { criteria: filterCriteria }
                    : {}),
                },
              },
            },
          ],
        },
      });

      return formatSuccess(`Set basic filter on ${range}`);
    })
  );

  // ─── sheets_sort_range ────────────────────────────────────────────────────────

  server.tool(
    "sheets_sort_range",
    "Sort a range by one or more columns (0-based indices).",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      range: z.string().describe("A1 notation range (exclude headers)"),
      sheet_id: z.number().int().optional().describe("Numeric sheet ID (default: 0)"),
      sort_specs: z
        .array(
          z.object({
            column_index: z
              .number()
              .int()
              .describe("0-based column index within the range to sort by"),
            ascending: z.boolean().describe("Sort direction - true for A→Z / smallest first"),
          })
        )
        .min(1)
        .describe("Sort specifications - ordered by priority (first entry is the primary sort)"),
    },
    withErrorHandling(async ({ spreadsheet_id, range, sheet_id, sort_specs }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);
      const resolvedSheetId = sheet_id ?? await resolveSheetIdFromRange(id, range);
      const gridRange = toGridRange(range, resolvedSheetId);

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: {
          requests: [
            {
              sortRange: {
                range: gridRange,
                sortSpecs: sort_specs.map((s) => ({
                  dimensionIndex: s.column_index,
                  sortOrder: s.ascending ? "ASCENDING" : "DESCENDING",
                })),
              },
            },
          ],
        },
      });

      const specDesc = sort_specs
        .map((s) => `col ${s.column_index} ${s.ascending ? "ASC" : "DESC"}`)
        .join(", ");

      return formatSuccess(`Sorted ${range} by: ${specDesc}`);
    })
  );
}
