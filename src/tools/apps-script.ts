/**
 * Apps Script tools - create, read, update, and run bound scripts.
 *
 * Enables the LLM to build self-sustaining spreadsheet applications:
 * custom functions, triggers, menus, automations, and complex logic
 * that humans can use without AI assistance.
 */

import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { getScriptClient, extractFileId } from "../client/google-client.js";
import { formatSuccess, formatError } from "../utils/response.js";
import { withErrorHandling } from "../utils/error-handler.js";

// ─── Helpers ──────────────────────────────────────────────────────────────────

/** Default manifest for a spreadsheet-bound script. */
function defaultManifest(timeZone = "America/New_York"): string {
  return JSON.stringify(
    {
      timeZone,
      dependencies: {},
      exceptionLogging: "STACKDRIVER",
      runtimeVersion: "V8",
    },
    null,
    2
  );
}

/** Format script files for display. */
function formatFiles(
  files: Array<{ name?: string | null; type?: string | null; source?: string | null }>
): string {
  return files
    .map((f) => {
      const header = `── ${f.name ?? "untitled"} (${f.type ?? "?"}) ──`;
      return `${header}\n${f.source ?? "(empty)"}`;
    })
    .join("\n\n");
}

// ─── Tool registration ────────────────────────────────────────────────────────

export function registerAppsScriptTools(server: McpServer): void {
  // ── script_create ───────────────────────────────────────────────────────────
  server.tool(
    "script_create",
    "Create an Apps Script project bound to a spreadsheet. Returns script ID. Use script_update to add code.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL to bind the script to"),
      title: z.string().optional().describe("Script project title (default: spreadsheet name + ' Script')"),
    },
    withErrorHandling(async (args) => {
      const script = await getScriptClient();
      const parentId = extractFileId(args.spreadsheet_id);

      const res = await script.projects.create({
        requestBody: {
          title: args.title ?? "Bound Script",
          parentId,
        },
      });

      const scriptId = res.data.scriptId;
      if (!scriptId) return formatError("Failed to create script project.");

      return formatSuccess(
        `Created Apps Script project\nscriptId: ${scriptId}\nBound to spreadsheet: ${parentId}\n\nNext: use script_update to add code.`
      );
    })
  );

  // ── script_get ──────────────────────────────────────────────────────────────
  server.tool(
    "script_get",
    "Read all files in an Apps Script project. Returns code, manifest, and metadata.",
    {
      script_id: z.string().describe("Apps Script project ID"),
    },
    withErrorHandling(async (args) => {
      const script = await getScriptClient();

      const res = await script.projects.getContent({
        scriptId: args.script_id,
      });

      const files = res.data.files ?? [];
      if (files.length === 0) return formatSuccess("(empty project - no files)");

      return formatSuccess(formatFiles(files));
    })
  );

  // ── script_update ───────────────────────────────────────────────────────────
  server.tool(
    "script_update",
    "Write or replace all files in a script project. Omitted files are deleted. Auto-generates manifest if not included (file name: 'appsscript', type: JSON).",
    {
      script_id: z.string().describe("Apps Script project ID"),
      files: z
        .array(
          z.object({
            name: z.string().describe("File name without extension (e.g. 'Code', 'Triggers', 'Utils')"),
            source: z.string().describe("Full source code for this file"),
            type: z
              .enum(["SERVER_JS", "HTML", "JSON"])
              .optional()
              .describe("File type (default: SERVER_JS). Use JSON only for appsscript.json"),
          })
        )
        .describe("Array of script files to write"),
      time_zone: z
        .string()
        .optional()
        .describe("Timezone for manifest if auto-generating (default: America/New_York)"),
    },
    withErrorHandling(async (args) => {
      const script = await getScriptClient();

      // Ensure manifest is included
      const hasManifest = args.files.some(
        (f) => f.name === "appsscript" && f.type === "JSON"
      );

      const files = args.files.map((f) => ({
        name: f.name,
        type: f.type ?? "SERVER_JS",
        source: f.source,
      }));

      if (!hasManifest) {
        files.push({
          name: "appsscript",
          type: "JSON",
          source: defaultManifest(args.time_zone),
        });
      }

      await script.projects.updateContent({
        scriptId: args.script_id,
        requestBody: { files },
      });

      const fileNames = files.map((f) => `${f.name} (${f.type})`).join(", ");
      return formatSuccess(
        `Updated ${files.length} file(s): ${fileNames}`
      );
    })
  );

  // ── script_get_bound ────────────────────────────────────────────────────────
  server.tool(
    "script_get_bound",
    "Find Apps Script projects bound to a spreadsheet via Drive metadata. Also checks common script naming patterns.",
    {
      spreadsheet_id: z.string().describe("Spreadsheet ID or URL"),
    },
    withErrorHandling(async (args) => {
      const { getDriveClient } = await import("../client/google-client.js");
      const drive = await getDriveClient();
      const parentId = extractFileId(args.spreadsheet_id);

      // Get the spreadsheet name for searching
      const fileMeta = await drive.files.get({
        fileId: parentId,
        fields: "name",
      });
      const ssName = fileMeta.data.name ?? "";

      // Strategy 1: Search Drive for scripts with the spreadsheet as parent
      const parentSearch = await drive.files.list({
        q: `mimeType='application/vnd.google-apps.script' and '${parentId}' in parents and trashed=false`,
        fields: "files(id, name)",
        pageSize: 5,
      });

      // Strategy 2: Search for scripts with matching names
      const nameEsc = ssName.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
      const nameSearch = await drive.files.list({
        q: `mimeType='application/vnd.google-apps.script' and name contains '${nameEsc}' and trashed=false`,
        fields: "files(id, name)",
        pageSize: 5,
      });

      // Merge and deduplicate results
      const seen = new Set<string>();
      const results: Array<{ id: string; name: string }> = [];
      for (const file of [...(parentSearch.data.files ?? []), ...(nameSearch.data.files ?? [])]) {
        if (file.id && !seen.has(file.id)) {
          seen.add(file.id);
          results.push({ id: file.id, name: file.name ?? "(untitled)" });
        }
      }

      if (results.length === 0) {
        return formatSuccess(
          "No Apps Script project found for this spreadsheet.\nUse script_create to create one.\n\nNote: bound scripts created in the Apps Script editor may not appear here - use the script ID from the editor URL."
        );
      }

      const lines = results.map((f) => `scriptId: ${f.id}\nname: ${f.name}`);
      return formatSuccess(`Found ${results.length} script(s):\n\n${lines.join("\n\n")}`);
    })
  );

  // ── script_run ──────────────────────────────────────────────────────────────
  server.tool(
    "script_run",
    "Execute a function via Apps Script API. Requires: (1) deployed version, (2) script's GCP project must match the OAuth credentials' GCP project. If 'not found', set the GCP project in Apps Script editor > Project Settings.",
    {
      script_id: z.string().describe("Apps Script project ID"),
      function_name: z.string().describe("Function name to execute"),
      parameters: z
        .array(z.unknown())
        .optional()
        .describe("Arguments to pass to the function"),
    },
    withErrorHandling(async (args) => {
      const script = await getScriptClient();

      const res = await script.scripts.run({
        scriptId: args.script_id,
        requestBody: {
          function: args.function_name,
          parameters: args.parameters ?? [],
        },
      });

      if (res.data.error) {
        const err = res.data.error;
        const details = err.details?.map(
          (d: Record<string, unknown>) =>
            `${d.errorType ?? "Error"}: ${d.errorMessage ?? "unknown"}`
        ) ?? [];
        return formatError(
          `Script error: ${details.join("\n") || JSON.stringify(err)}`
        );
      }

      const result = res.data.response?.result;
      if (result === undefined || result === null) {
        return formatSuccess("Function executed (no return value).");
      }

      return formatSuccess(
        typeof result === "string" ? result : JSON.stringify(result, null, 2)
      );
    })
  );

  // ── script_create_version ───────────────────────────────────────────────────
  server.tool(
    "script_create_version",
    "Create a versioned snapshot of the script. Required before deploying as API executable.",
    {
      script_id: z.string().describe("Apps Script project ID"),
      description: z.string().optional().describe("Version description"),
    },
    withErrorHandling(async (args) => {
      const script = await getScriptClient();

      const res = await script.projects.versions.create({
        scriptId: args.script_id,
        requestBody: {
          description: args.description ?? "",
        },
      });

      const version = res.data.versionNumber;
      return formatSuccess(
        `Created version ${version}${args.description ? `: ${args.description}` : ""}\n\nNext: use script_deploy to make it executable.`
      );
    })
  );

  // ── script_deploy ───────────────────────────────────────────────────────────
  server.tool(
    "script_deploy",
    "Deploy a script version as an API executable (required for script_run). Creates or updates the deployment.",
    {
      script_id: z.string().describe("Apps Script project ID"),
      version_number: z.number().int().describe("Version number to deploy (from script_create_version)"),
      description: z.string().optional().describe("Deployment description"),
    },
    withErrorHandling(async (args) => {
      const script = await getScriptClient();

      // Check for existing deployments
      const existing = await script.projects.deployments.list({
        scriptId: args.script_id,
      });

      const deployments = existing.data.deployments ?? [];
      // Only update API executable deployments - skip web app / add-on deployments
      // which have entryPoints with non-EXECUTION_API types
      const apiDeployment = deployments.find((d) => {
        if (!d.deploymentConfig?.versionNumber) return false;
        const entries = d.entryPoints ?? [];
        // If no entry points or has an EXECUTION_API entry, it's an API deployment
        return entries.length === 0 || entries.some((e) => e.entryPointType === "EXECUTION_API");
      });

      if (apiDeployment?.deploymentId) {
        // Update existing deployment
        await script.projects.deployments.update({
          scriptId: args.script_id,
          deploymentId: apiDeployment.deploymentId,
          requestBody: {
            deploymentConfig: {
              versionNumber: args.version_number,
              description: args.description ?? "",
              scriptId: args.script_id,
            },
          },
        });
        return formatSuccess(
          `Updated deployment ${apiDeployment.deploymentId} to version ${args.version_number}\n\nThe script is now executable via script_run.`
        );
      }

      // Create new deployment
      const res = await script.projects.deployments.create({
        scriptId: args.script_id,
        requestBody: {
          versionNumber: args.version_number,
          description: args.description ?? "",
        },
      });

      return formatSuccess(
        `Deployed version ${args.version_number}\ndeploymentId: ${res.data.deploymentId}\n\nThe script is now executable via script_run.`
      );
    })
  );
}
