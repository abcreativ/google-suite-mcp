/**
 * Named range tools: create, list, delete named ranges.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getSheetsClient, extractFileId } from "../client/google-client.js";
import { formatSuccess, formatError } from "../utils/response.js";
import { withErrorHandling } from "../utils/error-handler.js";
import { toGridRange } from "../utils/range.js";
import { resolveSheetIdFromRange } from "../utils/sheet-resolver.js";

export function registerNamedRangeTools(server: McpServer): void {
  // ─── sheets_create_named_range ───────────────────────────────────────────────

  server.tool(
    "sheets_create_named_range",
    "Create a named range (usable in formulas).",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      name: z.string().describe("Name for the range (must be unique, no spaces)"),
      range: z.string().describe("A1 notation range, e.g. 'Sheet1!A1:B10'"),
      sheet_id: z.number().int().optional().describe("Numeric sheet ID (default: 0 for first sheet)"),
    },
    withErrorHandling(async ({ spreadsheet_id, name, range, sheet_id }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);

      const resolvedSheetId = sheet_id ?? await resolveSheetIdFromRange(id, range);
      const gridRange = toGridRange(range, resolvedSheetId);

      const res = await sheets.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: {
          requests: [
            {
              addNamedRange: {
                namedRange: {
                  name,
                  range: gridRange,
                },
              },
            },
          ],
        },
      });

      const namedRangeId =
        res.data.replies?.[0]?.addNamedRange?.namedRange?.namedRangeId ?? "unknown";

      return formatSuccess(`Created named range "${name}" (ID: ${namedRangeId}) → ${range}`);
    })
  );

  // ─── sheets_list_named_ranges ────────────────────────────────────────────────

  server.tool(
    "sheets_list_named_ranges",
    "List all named ranges defined in a spreadsheet.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
    },
    withErrorHandling(async ({ spreadsheet_id }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);

      const res = await sheets.spreadsheets.get({
        spreadsheetId: id,
        fields: "namedRanges",
      });

      const namedRanges = res.data.namedRanges ?? [];
      if (namedRanges.length === 0) {
        return formatSuccess("No named ranges found.");
      }

      const lines = namedRanges.map((nr) => {
        const r = nr.range ?? {};
        return `${nr.namedRangeId}  "${nr.name}"  sheetId=${r.sheetId} rows=${r.startRowIndex}-${r.endRowIndex} cols=${r.startColumnIndex}-${r.endColumnIndex}`;
      });

      return formatSuccess(`Named ranges (${namedRanges.length}):\n${lines.join("\n")}`);
    })
  );

  // ─── sheets_delete_named_range ───────────────────────────────────────────────

  server.tool(
    "sheets_delete_named_range",
    "Delete a named range by name or by its namedRangeId.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      name: z.string().optional().describe("Named range name (used to look up the ID if namedRangeId not given)"),
      named_range_id: z.string().optional().describe("Named range ID (takes precedence over name)"),
    },
    withErrorHandling(async ({ spreadsheet_id, name, named_range_id }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);

      let rangeId = named_range_id;

      if (!rangeId) {
        if (!name) {
          return formatError("Provide either name or named_range_id.");
        }

        const res = await sheets.spreadsheets.get({
          spreadsheetId: id,
          fields: "namedRanges",
        });

        const match = (res.data.namedRanges ?? []).find(
          (nr) => nr.name?.toLowerCase() === name.toLowerCase()
        );

        if (!match?.namedRangeId) {
          return formatError(`Named range "${name}" not found.`);
        }

        rangeId = match.namedRangeId;
      }

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: {
          requests: [
            {
              deleteNamedRange: {
                namedRangeId: rangeId,
              },
            },
          ],
        },
      });

      return formatSuccess(`Deleted named range (ID: ${rangeId})`);
    })
  );
}
