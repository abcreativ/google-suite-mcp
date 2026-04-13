/**
 * Find and replace tools.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getSheetsClient, extractFileId } from "../client/google-client.js";
import { formatSuccess } from "../utils/response.js";
import { withErrorHandling } from "../utils/error-handler.js";

export function registerFindReplaceTools(server: McpServer): void {
  // ─── sheets_find_replace ──────────────────────────────────────────────────────

  server.tool(
    "sheets_find_replace",
    "Executes a single find-and-replace operation across one sheet or all sheets in a spreadsheet using spreadsheets.batchUpdate with findReplace; supports regex, case sensitivity, and formula-text search. Use when the user asks to update a value or label that appears in multiple cells, such as renaming a product code or correcting a misspelling throughout a sheet. Use when searching within formula text rather than displayed values by setting include_formulas=true. Do not use when: running multiple distinct find-and-replace patterns in one call - use sheets_find_replace_many instead; finding cells containing a value without replacing - use sheets_search_values instead; sorting or filtering rows - use sheets_sort_range or sheets_set_filter instead. Returns: 'Find & replace complete.\\nOccurrences replaced: {N}\\nRows affected: {N}'. Parameters: - find: literal text or regex pattern string, e.g. 'ProductCode_OLD' - replacement: text to substitute, e.g. 'ProductCode_NEW' - search_by_regex: true to treat find as a regex (default false) - all_sheets: true to search every tab; defaults to true when sheet_id is omitted.",
    {
      spreadsheet_id: z.string().describe("sheet ID from the URL (the token between /d/ and /edit) or the full URL"),
      find: z.string().describe("Text or regex pattern to search for"),
      replacement: z.string().describe("Replacement text"),
      match_case: z.boolean().optional().describe("Case-sensitive search (default: false)"),
      match_entire_cell: z
        .boolean()
        .optional()
        .describe("Match only if the entire cell content matches (default: false)"),
      search_by_regex: z
        .boolean()
        .optional()
        .describe("Treat find as a regular expression (default: false)"),
      include_formulas: z
        .boolean()
        .optional()
        .describe("Search within formula text rather than displayed values (default: false)"),
      all_sheets: z
        .boolean()
        .optional()
        .describe("Search across all sheets. When true, sheet_id is ignored (default: false)"),
      sheet_id: z
        .number()
        .int()
        .optional()
        .describe("Limit search to this sheet ID. Ignored when all_sheets is true."),
    },
    withErrorHandling(async ({
      spreadsheet_id,
      find,
      replacement,
      match_case,
      match_entire_cell,
      search_by_regex,
      include_formulas,
      all_sheets,
      sheet_id,
    }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);

      // Default to all sheets if no specific sheet_id is provided
      const effectiveAllSheets = all_sheets ?? (sheet_id === undefined);

      const findReplaceRequest: Record<string, unknown> = {
        find,
        replacement,
        matchCase: match_case ?? false,
        matchEntireCell: match_entire_cell ?? false,
        searchByRegex: search_by_regex ?? false,
        includeFormulas: include_formulas ?? false,
        allSheets: effectiveAllSheets,
      };

      if (!effectiveAllSheets && sheet_id !== undefined) {
        findReplaceRequest.sheetId = sheet_id;
      }

      const res = await sheets.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: {
          requests: [
            {
              findReplace: findReplaceRequest,
            },
          ],
        },
      });

      const reply = res.data.replies?.[0]?.findReplace;
      const occurrencesChanged = reply?.occurrencesChanged ?? 0;
      const rowsChanged = reply?.rowsChanged ?? 0;

      return formatSuccess(
        `Find & replace complete.\nOccurrences replaced: ${occurrencesChanged}\nRows affected: ${rowsChanged}`
      );
    })
  );

  // ─── sheets_find_replace_many (bulk) ────────────────────────────────────────

  server.tool(
    "sheets_find_replace_many",
    "Executes multiple find-and-replace operations in a single API call using spreadsheets.batchUpdate with multiple findReplace requests; each entry can target different patterns and different sheets. Use when the user asks to apply several text substitutions at once, such as updating a set of product codes or correcting multiple labels in one operation. Use when sequential single replacements would require multiple round trips to the API. Do not use when: executing a single find-and-replace - use sheets_find_replace instead; finding cells without replacing - use sheets_search_values instead; sorting or filtering - use sheets_sort_range or sheets_set_filter instead. Returns: 'Bulk find & replace: {N} operation(s), {N} occurrence(s) replaced across {N} row(s)'. Parameters: - replacements: array of objects each with find and replacement strings, plus optional sheet_id, all_sheets, match_case, match_entire_cell, search_by_regex, include_formulas; operations execute in array order.",
    {
      spreadsheet_id: z.string().describe("sheet ID from the URL (the token between /d/ and /edit) or the full URL"),
      replacements: z
        .array(
          z.object({
            find: z.string().describe("Text or regex to find"),
            replacement: z.string().describe("Replacement text"),
            sheet_id: z.number().int().optional().describe("Limit to this sheet ID"),
            all_sheets: z.boolean().optional().describe("Search all sheets (default: true)"),
            match_case: z.boolean().optional(),
            match_entire_cell: z.boolean().optional(),
            search_by_regex: z.boolean().optional(),
            include_formulas: z.boolean().optional(),
          })
        )
        .min(1)
        .describe("Array of find/replace pairs, executed in order"),
    },
    withErrorHandling(async ({ spreadsheet_id, replacements }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);

      const requests = replacements.map((r) => {
        const effectiveAllSheets = r.all_sheets ?? (r.sheet_id === undefined);
        const req: Record<string, unknown> = {
          find: r.find,
          replacement: r.replacement,
          matchCase: r.match_case ?? false,
          matchEntireCell: r.match_entire_cell ?? false,
          searchByRegex: r.search_by_regex ?? false,
          includeFormulas: r.include_formulas ?? false,
          allSheets: effectiveAllSheets,
        };
        if (!effectiveAllSheets && r.sheet_id !== undefined) {
          req.sheetId = r.sheet_id;
        }
        return { findReplace: req };
      });

      const res = await sheets.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: { requests },
      });

      const replies = res.data.replies ?? [];
      let totalOccurrences = 0;
      let totalRows = 0;
      for (const reply of replies) {
        const fr = reply?.findReplace;
        if (fr) {
          totalOccurrences += (fr as { occurrencesChanged?: number }).occurrencesChanged ?? 0;
          totalRows += (fr as { rowsChanged?: number }).rowsChanged ?? 0;
        }
      }

      return formatSuccess(
        `Bulk find & replace: ${replacements.length} operation(s), ${totalOccurrences} occurrence(s) replaced across ${totalRows} row(s)`
      );
    })
  );
}
