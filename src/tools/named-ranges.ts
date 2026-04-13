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
    "Creates a named range in a Google Sheet using spreadsheets.batchUpdate with addNamedRange, making a cell region addressable by name in formulas. Use when the user asks to label a cell region for formula reuse, such as defining 'TaxRate' for a fixed cell or 'SalesData' for a data block. Use when a formula in another cell needs to reference a region by a human-readable name rather than an A1 address. Do not use when: listing existing named ranges - use sheets_list_named_ranges instead; removing a named range - use sheets_delete_named_range instead; protecting a range from edits - use sheets_protect_range instead. Returns: 'Created named range \"{name}\" (ID: {namedRangeId}) → {range}'. Parameters: - name: unique identifier with no spaces, e.g. 'SalesData' or 'TaxRate' - range: A1 notation including sheet name, e.g. 'Sheet1!A1:B10' - sheet_id: numeric sheet ID (optional; resolved from range if omitted). Example: sheets_create_named_range(spreadsheetId, 'Revenue', 'Sheet1!C2:C50')",
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
    "Reads all named ranges from a spreadsheet using spreadsheets.get with fields=namedRanges, returning each range's ID, name, and grid coordinates. Use when the user asks what named ranges exist, or when you need a namedRangeId before calling sheets_delete_named_range. Use when inspecting a spreadsheet's formula-addressable regions before writing formulas that reference them by name. Do not use when: creating a new named range - use sheets_create_named_range instead; deleting a named range - use sheets_delete_named_range instead; reading cell values by name - use sheets_read_range with the name in A1 notation instead. Returns: 'Named ranges ({N}):\\n{namedRangeId}  \"{name}\"  sheetId=N rows=N-N cols=N-N' for each range, or 'No named ranges found.' Parameters: - spreadsheetId: the ID from the sheet URL (between /d/ and /edit).",
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
    "Deletes a named range from a spreadsheet using spreadsheets.batchUpdate with deleteNamedRange; accepts either the human-readable name or the namedRangeId directly. Use when the user asks to remove a named range that is no longer needed, or after renaming a region by deleting the old name and creating a new one. Use when a formula-addressable label must be cleaned up without affecting cell content. Do not use when: creating a named range - use sheets_create_named_range instead; listing named ranges to find the ID - use sheets_list_named_ranges first; protecting a range - use sheets_protect_range instead. Returns: 'Deleted named range (ID: {namedRangeId})'. Parameters: - name: human-readable range name (used to look up the ID if named_range_id is not given) - named_range_id: takes precedence over name if both are supplied; obtain from sheets_list_named_ranges.",
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
