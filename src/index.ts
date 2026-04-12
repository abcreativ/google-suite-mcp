#!/usr/bin/env node

/**
 * google-suite-mcp - MCP server entry point.
 *
 * Provides Google Sheets, Drive, and Docs tools for any MCP-compatible
 * AI client. Authentication is handled via OAuth2 with a persistent
 * refresh token stored in ~/.google-suite-mcp/tokens.json.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { getAuthClient, refreshTokens, getTokenInfo } from "./auth/oauth.js";
import { formatSuccess, formatError } from "./utils/response.js";
import { withErrorHandling } from "./utils/error-handler.js";

// ─── Phase 2: Sheets tool domain imports ─────────────────────────────────────
import { registerSpreadsheetTools } from "./tools/spreadsheet.js";
import { registerSheetTools }       from "./tools/sheets.js";
import { registerReadingTools }     from "./tools/reading.js";
import { registerWritingTools }     from "./tools/writing.js";
import { registerFormulaTools }     from "./tools/formulas.js";

// ─── Phase 3 & 6: Drive and Docs tool domain imports ─────────────────────────
import { registerDriveTools } from "./tools/drive.js";
import { registerDocsTools }  from "./tools/docs.js";

// ─── Phase 4: Formatting, Charts, Dashboard imports ──────────────────────────
import { registerFormattingTools } from "./tools/formatting.js";
import { registerChartTools }      from "./tools/charts.js";
import { registerDashboardTools }  from "./tools/dashboard.js";

// ─── Phase 5: Sheets power features ──────────────────────────────────────────
import { registerNamedRangeTools }  from "./tools/named-ranges.js";
import { registerValidationTools }  from "./tools/validation.js";
import { registerFilterSortTools }  from "./tools/filter-sort.js";
import { registerFindReplaceTools } from "./tools/find-replace.js";
import { registerProtectionTools }  from "./tools/protection.js";
import { registerBatchTools }       from "./tools/batch.js";

// ─── Phase 7: Apps Script ────────────────────────────────────────────────────
import { registerAppsScriptTools } from "./tools/apps-script.js";

// ─── Phase 8: Bulk / high-level tools ───────────────────────────────────────
import { registerTableTools }        from "./tools/table.js";
import { registerSheetBuilderTools } from "./tools/sheet-builder.js";

// ─── Server version ───────────────────────────────────────────────────────────

const VERSION = "1.0.0";

// ─── Server setup ─────────────────────────────────────────────────────────────

const server = new McpServer(
  {
    name: "google-suite-mcp",
    version: VERSION,
  },
  {
    capabilities: {
      logging: {},
    },
  }
);

// ─── Auth tools ───────────────────────────────────────────────────────────────

server.tool(
  "auth_status",
  "Check OAuth token state: expiry, refresh token presence, scopes. Use to diagnose auth errors.",
  {},
  withErrorHandling(async () => {
    const info = getTokenInfo();

    if (!info.hasTokens) {
      return formatError(
        `Not authenticated.\n\nNo tokens found at: ${info.tokensPath}\n\nRestart the server to trigger the OAuth2 browser flow.`
      );
    }

    const expiryStr = info.expiryDate
      ? new Date(info.expiryDate).toISOString()
      : "unknown";

    const isExpired = info.expiryDate ? Date.now() > info.expiryDate : false;
    const expiryLabel = isExpired
      ? `${expiryStr} (EXPIRED - use auth_refresh)`
      : expiryStr;

    const lines = [
      `Authentication Status`,
      `─────────────────────`,
      `Tokens file:      ${info.tokensPath}`,
      `Has refresh token: ${info.hasRefreshToken ? "yes" : "no (re-authentication required)"}`,
      `Access token expiry: ${expiryLabel}`,
      `Credentials file: ${info.credentialsPath}`,
      ``,
      `Scopes:`,
      `  • https://www.googleapis.com/auth/spreadsheets`,
      `  • https://www.googleapis.com/auth/drive`,
      `  • https://www.googleapis.com/auth/documents`,
      `  • https://www.googleapis.com/auth/script.projects`,
      `  • https://www.googleapis.com/auth/script.deployments`,
    ];

    return formatSuccess(lines.join("\n"));
  })
);

server.tool(
  "auth_refresh",
  "Force OAuth2 token refresh. Use when tools return 401 errors.",
  {},
  withErrorHandling(async () => {
    const result = await refreshTokens();

    const expiryStr = result.expiryDate
      ? new Date(result.expiryDate).toISOString()
      : "unknown";

    return formatSuccess(
      `Token refreshed successfully.\nNew access token expires at: ${expiryStr}`
    );
  })
);

// ─── Phase 2: Sheets tool domain registration ────────────────────────────────
registerSpreadsheetTools(server);
registerSheetTools(server);
registerReadingTools(server);
registerWritingTools(server);
registerFormulaTools(server);

// ─── Phase 3 & 6: Drive and Docs tool domain registration ────────────────────
registerDriveTools(server);
registerDocsTools(server);

// ─── Phase 4: Formatting, Charts, Dashboard registration ─────────────────────
registerFormattingTools(server);
registerChartTools(server);
registerDashboardTools(server);

// ─── Phase 5: Sheets power features registration ─────────────────────────────
registerNamedRangeTools(server);
registerValidationTools(server);
registerFilterSortTools(server);
registerFindReplaceTools(server);
registerProtectionTools(server);
registerBatchTools(server);

// ─── Phase 7: Apps Script registration ───────────────────────────────────────
registerAppsScriptTools(server);

// ─── Phase 8: Bulk / high-level tools ───────────────────────────────────────
registerTableTools(server);
registerSheetBuilderTools(server);

// ─── Main ─────────────────────────────────────────────────────────────────────

async function main(): Promise<void> {
  // Ensure authentication is ready before accepting tool calls
  try {
    await getAuthClient();
  } catch (err) {
    // Auth errors (e.g. missing credentials.json) are reported by getAuthClient
    // itself with a clear message before process.exit - this catch is a safety net.
    console.error("Failed to initialise authentication:", err);
    process.exit(1);
  }

  const transport = new StdioServerTransport();

  console.error(`google-suite-mcp v${VERSION} starting...`);
  console.error("─".repeat(50));
  console.error("Auth: auth_status, auth_refresh");
  console.error("─".repeat(50));
  console.error("Sheets: spreadsheet mgmt, sheet tabs, read, write, formulas");

  await server.connect(transport);
  console.error("google-suite-mcp running on stdio");
}

main().catch((err: unknown) => {
  console.error("Fatal error:", err);
  process.exit(1);
});
