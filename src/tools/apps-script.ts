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
    "Creates a new Apps Script project bound to a spreadsheet using script.projects.create; returns the scriptId needed for all subsequent script_ calls. Use when the user asks to add automation or custom functions to a spreadsheet for the first time. Use as the first step before calling script_update to write code and script_create_version to prepare for deployment. Do not use when: a script already exists for the spreadsheet - use script_get_bound to find it; reading existing script code - use script_get instead; writing code to an existing project - use script_update instead; running a function - use script_run instead; creating a version snapshot - use script_create_version instead; deploying a version - use script_deploy instead. Returns: 'Created Apps Script project\\nscriptId: {scriptId}\\nBound to spreadsheet: {parentId}\\n\\nNext: use script_update to add code.'. Parameters: - spreadsheet_id: spreadsheet ID or URL to bind the script to - title: display name for the script project, e.g. 'Inventory Automation' (optional).",
    {
      spreadsheet_id: z.string().describe("sheet ID from the URL (the token between /d/ and /edit) or the full URL; the script will be bound to this sheet"),
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
    "Reads all files in an Apps Script project using script.projects.getContent, returning each file's name, type (SERVER_JS, HTML, JSON), and full source code. Use when the user asks to see the current code in a script project. Use when inspecting existing automation before modifying it with script_update. Do not use when: creating a new script project - use script_create instead; writing or replacing code - use script_update instead; finding a script bound to a spreadsheet - use script_get_bound instead; running a function - use script_run instead; creating a version - use script_create_version instead; deploying - use script_deploy instead. Returns: one block per file formatted as '-- {name} ({type}) --\\n{source}'; blocks separated by blank lines. Parameters: - script_id: Apps Script project ID.",
    {
      script_id: z.string().describe("Apps Script project ID (57-char token visible in the Apps Script editor URL, or returned by script_create, or looked up from a Sheet/Doc via script_get_bound)"),
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
    "Writes or replaces all files in an Apps Script project using script.projects.updateContent; files not included in the call are deleted. Auto-generates the appsscript.json manifest if not provided. Use when the user asks to write, update, or replace automation code in a script project. Use when adding new files or functions after creating a project with script_create. Do not use when: creating a new project - use script_create instead; reading existing code before editing - use script_get first; finding a bound script - use script_get_bound instead; running a function - use script_run instead; creating a snapshot before deploying - use script_create_version instead; deploying the script - use script_deploy instead. Returns: 'Updated {N} file(s): {name (TYPE), name (TYPE), ...}'. Parameters: - script_id: Apps Script project ID - files: array of objects with name (no extension), source (full code), and optional type (SERVER_JS default, HTML, JSON) - time_zone: IANA timezone for auto-generated manifest (optional; default America/New_York).",
    {
      script_id: z.string().describe("Apps Script project ID (57-char token visible in the Apps Script editor URL, or returned by script_create, or looked up from a Sheet/Doc via script_get_bound)"),
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
    "Finds Apps Script projects bound to a spreadsheet by searching Drive for scripts with the spreadsheet as parent and by name pattern; deduplicates results from both strategies. Use when the user asks whether a spreadsheet already has a script, or when looking up the scriptId before calling script_get or script_update. Use before script_create to avoid creating a duplicate project. Do not use when: creating a new project - use script_create instead; reading a known script's code - use script_get instead; writing code - use script_update instead; running a function - use script_run instead; creating a version - use script_create_version instead; deploying - use script_deploy instead. Returns: 'Found {N} script(s):\\n\\nscriptId: {id}\\nname: {name}' per result; or a not-found message with instruction to use script_create. Parameters: - spreadsheet_id: spreadsheet ID or URL to search against.",
    {
      spreadsheet_id: z.string().describe("sheet ID from the URL (the token between /d/ and /edit) or the full URL"),
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
    "Executes a named function in a deployed Apps Script project via script.scripts.run; requires an active deployment (created with script_deploy) and the script's GCP project must match the OAuth credentials' GCP project. Use when the user asks to trigger a custom function or automation in a script. Use when testing a deployed function after updating code and creating a new version. Do not use when: the script has no deployment - call script_create_version then script_deploy first; creating a new project - use script_create instead; reading code - use script_get instead; writing code - use script_update instead; creating a version snapshot - use script_create_version instead; deploying a version - use script_deploy instead. Returns: function return value as a string or pretty-printed JSON; returns 'Function executed (no return value).' if the function returns nothing. Parameters: - script_id: Apps Script project ID - function_name: name of the function to call, e.g. 'sendWeeklyReport' - parameters: array of arguments to pass (optional).",
    {
      script_id: z.string().describe("Apps Script project ID (57-char token visible in the Apps Script editor URL, or returned by script_create, or looked up from a Sheet/Doc via script_get_bound)"),
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
    "Creates a versioned snapshot of an Apps Script project using script.projects.versions.create; the version number returned is required by script_deploy. Use when the user asks to save a snapshot before deploying, or after updating code with script_update and ready to publish. Use as the required step between script_update and script_deploy in the standard automation workflow. Do not use when: creating a new project - use script_create instead; writing code - use script_update instead; finding a bound script - use script_get_bound instead; reading code - use script_get instead; deploying a version - that is the next step, use script_deploy after this call; running a function - use script_run instead. Returns: 'Created version {N}: {description}\\n\\nNext: use script_deploy to make it executable.' Parameters: - script_id: Apps Script project ID - description: optional human-readable label, e.g. 'Added email trigger'.",
    {
      script_id: z.string().describe("Apps Script project ID (57-char token visible in the Apps Script editor URL, or returned by script_create, or looked up from a Sheet/Doc via script_get_bound)"),
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
    "Deploys a versioned Apps Script project as an API executable using script.projects.deployments.create or .update; after deployment the script can be invoked with script_run. Updates an existing API executable deployment if one exists, otherwise creates a new one. Use when the user asks to deploy or publish a script so it can be executed. Use as the final step after script_update and script_create_version in the standard workflow. Do not use when: creating a new project - use script_create instead; writing code - use script_update instead; creating the version number needed here - use script_create_version first; running a function after deploying - use script_run instead; finding a bound script - use script_get_bound instead. Returns: 'Updated deployment {deploymentId} to version {N}...' if updating, or 'Deployed version {N}\\ndeploymentId: {id}...' if creating new. Parameters: - script_id: Apps Script project ID - version_number: version number from script_create_version, e.g. 3 - description: optional deployment label.",
    {
      script_id: z.string().describe("Apps Script project ID (57-char token visible in the Apps Script editor URL, or returned by script_create, or looked up from a Sheet/Doc via script_get_bound)"),
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
