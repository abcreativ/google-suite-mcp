/**
 * Raw batch update tool - power-user escape hatch for any Google Sheets API request.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getSheetsClient, extractFileId } from "../client/google-client.js";
import { formatSuccess } from "../utils/response.js";
import { withErrorHandling } from "../utils/error-handler.js";

export function registerBatchTools(server: McpServer): void {
  // ─── sheets_batch_update ──────────────────────────────────────────────────────

  server.tool(
    "sheets_batch_update",
    "Execute raw Sheets API batchUpdate requests. Power-user escape hatch for any unsupported operation.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
      requests: z
        .array(z.record(z.string(), z.unknown()))
        .min(1)
        .describe(
          "Array of Google Sheets API request objects, e.g. [{\"addSheet\": {\"properties\": {\"title\": \"NewSheet\"}}}]"
        ),
      include_spreadsheet_in_response: z
        .boolean()
        .optional()
        .describe("Include the updated spreadsheet resource in the response (default: false)"),
      response_ranges: z
        .array(z.string())
        .optional()
        .describe("Ranges to include in the response when include_spreadsheet_in_response is true"),
    },
    withErrorHandling(async ({
      spreadsheet_id,
      requests,
      include_spreadsheet_in_response,
      response_ranges,
    }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);

      const res = await sheets.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: {
          requests,
          includeSpreadsheetInResponse: include_spreadsheet_in_response ?? false,
          ...(response_ranges && response_ranges.length > 0
            ? { responseRanges: response_ranges }
            : {}),
        },
      });

      const replies = res.data.replies ?? [];

      // Extract useful info from each reply without dumping raw JSON
      const replyNotes: string[] = [];
      for (let i = 0; i < replies.length; i++) {
        const reply = replies[i] as Record<string, unknown> | null;
        if (!reply) continue;
        for (const [key, val] of Object.entries(reply)) {
          if (!val || typeof val !== "object") continue;
          // Dig into nested structures to find IDs and names
          const flat = JSON.parse(JSON.stringify(val));
          const ids: string[] = [];
          const extract = (o: Record<string, unknown>) => {
            for (const [k, v] of Object.entries(o)) {
              if ((k === "sheetId" || k === "chartId" || k === "spreadsheetId") && v != null) ids.push(`${k}=${v}`);
              if (k === "title" && typeof v === "string") ids.push(`title="${v}"`);
              if (v && typeof v === "object" && !Array.isArray(v)) extract(v as Record<string, unknown>);
            }
          };
          extract(flat);
          replyNotes.push(ids.length > 0 ? `${key}: ${ids.join(", ")}` : key);
        }
      }

      const summary = [
        `Batch update: ${requests.length} request(s), ${replies.length} reply(ies).`,
        ...(replyNotes.length > 0 ? replyNotes : []),
      ].join("\n");

      return formatSuccess(summary);
    })
  );
}
