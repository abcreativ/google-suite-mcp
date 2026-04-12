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
    "Protect a range - restrict editing to specific users or show warning.",
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
    "Remove a range protection by its protectedRangeId.",
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
    "List all protected ranges in a spreadsheet.",
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
