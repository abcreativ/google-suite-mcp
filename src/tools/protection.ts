/**
 * Range protection tools: protect, unprotect, list protected ranges.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getSheetsClient, extractFileId } from "../client/google-client.js";
import { formatSuccess } from "../utils/response.js";
import { withErrorHandling } from "../utils/error-handler.js";
import { toGridRange } from "../utils/range.js";
import { resolveSheetIdFromRange } from "../utils/sheet-resolver.js";

export function registerProtectionTools(server: McpServer): void {
  // ─── sheets_protect_range ─────────────────────────────────────────────────────

  server.tool(
    "sheets_protect_range",
    "Protects a cell range in a Google Sheet using spreadsheets.batchUpdate with addProtectedRange, restricting edits to specified users or displaying a warning to all editors. Use when the user asks to lock a header row or formula area so collaborators cannot accidentally overwrite it. Use when setting warning_only=true to flag a sensitive range without fully blocking edits. Do not use when: removing an existing protection - use sheets_unprotect_range with the returned protectedRangeId; listing current protections - use sheets_list_protected_ranges; creating a formula-addressable label - use sheets_create_named_range instead. Returns: 'Protected range created (ID: {protectedRangeId}) - \"{description}\"\\nRange: {range}' (description line omitted if no description supplied). The returned protectedRangeId is required by sheets_unprotect_range. Parameters: - range: A1 notation including sheet name, e.g. 'Sheet1!A1:Z1' - editor_emails: array of email strings allowed to edit; omit to restrict to owner only - warning_only: true shows a warning but does not block edits; overrides editor_emails.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      range: z.string().describe("A1 notation range"),
      sheet_id: z.number().int().optional().describe("Numeric sheet ID (default: 0)"),
      description: z.string().optional().describe("Human-readable description for this protection"),
      editor_emails: z
        .array(z.string())
        .optional()
        .describe("Emails of allowed editors (omit = owner only)"),
      warning_only: z
        .boolean()
        .optional()
        .describe(
          "If true, anyone can edit but will see a warning. Overrides editor_emails (default: false)."
        ),
    },
    withErrorHandling(async ({
      spreadsheet_id,
      range,
      sheet_id,
      description,
      editor_emails,
      warning_only,
    }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);
      const resolvedSheetId = sheet_id ?? await resolveSheetIdFromRange(id, range);
      const gridRange = toGridRange(range, resolvedSheetId);

      const protectedRange: Record<string, unknown> = {
        range: gridRange,
        warningOnly: warning_only ?? false,
      };

      if (description) {
        protectedRange.description = description;
      }

      if (!warning_only && editor_emails && editor_emails.length > 0) {
        protectedRange.editors = {
          users: editor_emails,
        };
      }

      const res = await sheets.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: {
          requests: [
            {
              addProtectedRange: {
                protectedRange,
              },
            },
          ],
        },
      });

      const newId =
        res.data.replies?.[0]?.addProtectedRange?.protectedRange?.protectedRangeId ?? "unknown";

      return formatSuccess(
        `Protected range created (ID: ${newId})${description ? ` - "${description}"` : ""}\nRange: ${range}`
      );
    })
  );

  // ─── sheets_unprotect_range ───────────────────────────────────────────────────

  server.tool(
    "sheets_unprotect_range",
    "Removes a range protection from a Google Sheet using spreadsheets.batchUpdate with deleteProtectedRange; requires the numeric protectedRangeId returned by sheets_protect_range or listed by sheets_list_protected_ranges. Use when the user asks to unlock a previously protected range so all collaborators can edit it again. Use when cleaning up stale protections after a workflow change. Do not use when: creating a new protection - use sheets_protect_range instead; finding the protectedRangeId first - use sheets_list_protected_ranges; the user wants a warning instead of full protection - use sheets_protect_range with warning_only=true. Returns: 'Removed protection (ID: {protected_range_id})'. Parameters: - protected_range_id: numeric integer ID returned by sheets_protect_range or sheets_list_protected_ranges; not the same as a sheet ID.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      protected_range_id: z.number().int().describe("The protectedRangeId to delete"),
    },
    withErrorHandling(async ({ spreadsheet_id, protected_range_id }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: {
          requests: [
            {
              deleteProtectedRange: {
                protectedRangeId: protected_range_id,
              },
            },
          ],
        },
      });

      return formatSuccess(`Removed protection (ID: ${protected_range_id})`);
    })
  );

  // ─── sheets_list_protected_ranges ─────────────────────────────────────────────

  server.tool(
    "sheets_list_protected_ranges",
    "Reads all protected ranges from a spreadsheet using spreadsheets.get with fields scoped to protectedRanges, returning each protection's ID, grid coordinates, allowed editors, and warning-only flag. Use when the user asks what ranges are locked, or when you need a protectedRangeId before calling sheets_unprotect_range. Use when auditing a spreadsheet's edit restrictions before modifying its structure. Do not use when: creating a protection - use sheets_protect_range instead; removing a protection - use sheets_unprotect_range instead; listing named ranges - use sheets_list_named_ranges instead. Returns: 'Protected ranges ({N}):\\n[{protectedRangeId}] sheetId=N rows=N-N cols=N-N | \"{description}\" | editors: {emails or \"owner only\"} | warningOnly: {bool}', or 'No protected ranges found.' Parameters: - spreadsheetId: the ID from the sheet URL (between /d/ and /edit).",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
    },
    withErrorHandling(async ({ spreadsheet_id }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);

      const res = await sheets.spreadsheets.get({
        spreadsheetId: id,
        fields: "sheets(properties.sheetId,protectedRanges)",
      });

      const sheetList = res.data.sheets ?? [];
      const allProtected: string[] = [];

      for (const sheet of sheetList) {
        const sheetId = sheet.properties?.sheetId;
        const ranges = sheet.protectedRanges ?? [];
        for (const pr of ranges) {
          const r = pr.range ?? {};
          const editors = pr.editors?.users?.join(", ") ?? "owner only";
          allProtected.push(
            `[${pr.protectedRangeId}] sheetId=${sheetId} rows=${r.startRowIndex}-${r.endRowIndex} cols=${r.startColumnIndex}-${r.endColumnIndex} | "${pr.description ?? ""}" | editors: ${editors} | warningOnly: ${pr.warningOnly ?? false}`
          );
        }
      }

      if (allProtected.length === 0) {
        return formatSuccess("No protected ranges found.");
      }

      return formatSuccess(`Protected ranges (${allProtected.length}):\n${allProtected.join("\n")}`);
    })
  );
}
