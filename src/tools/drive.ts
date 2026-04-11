/**
 * Drive tools - file search, operations, sharing, and permissions.
 */

import * as fs from "fs";
import * as path from "path";
import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { getDriveClient, extractFileId } from "../client/google-client.js";
import { formatSuccess, formatError } from "../utils/response.js";
import { withErrorHandling } from "../utils/error-handler.js";

// ─── Helpers ──────────────────────────────────────────────────────────────────

const MIME_SHORTCUTS: Record<string, string> = {
  sheet: "application/vnd.google-apps.spreadsheet",
  doc: "application/vnd.google-apps.document",
  folder: "application/vnd.google-apps.folder",
  slide: "application/vnd.google-apps.presentation",
  form: "application/vnd.google-apps.form",
};

const EXPORT_MIME: Record<string, Record<string, string>> = {
  "application/vnd.google-apps.spreadsheet": {
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    csv: "text/csv",
    pdf: "application/pdf",
  },
  "application/vnd.google-apps.document": {
    docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    pdf: "application/pdf",
    txt: "text/plain",
    html: "text/html",
  },
  "application/vnd.google-apps.presentation": {
    pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    pdf: "application/pdf",
  },
};

function guessMimeType(filePath: string): string {
  const ext = path.extname(filePath).slice(1).toLowerCase();
  const map: Record<string, string> = {
    pdf: "application/pdf",
    txt: "text/plain",
    html: "text/html",
    htm: "text/html",
    csv: "text/csv",
    json: "application/json",
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    gif: "image/gif",
    svg: "image/svg+xml",
    mp4: "video/mp4",
    docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    zip: "application/zip",
  };
  return map[ext] ?? "application/octet-stream";
}

// ─── Register all Drive tools ─────────────────────────────────────────────────

export function registerDriveTools(server: McpServer): void {
  // ── drive_search ────────────────────────────────────────────────────────────
  server.tool(
    "drive_search",
    "Search Drive by name, type, date, owner. Returns id, name, type, url. Paginated.",
    {
      query: z.string().optional().describe("Text to search in file names"),
      mimeType: z.string().optional().describe("Filter: sheet|doc|folder|slide|form or full MIME"),
      modifiedAfter: z.string().optional().describe("ISO date string, e.g. 2024-01-01"),
      owner: z.string().optional().describe("Owner email address"),
      pageToken: z.string().optional().describe("Token for next page"),
      pageSize: z.number().int().min(1).max(100).default(20),
    },
    withErrorHandling(async (args) => {
      const drive = await getDriveClient();

      // Sanitize all user inputs interpolated into the Drive query string
      const esc = (s: string) => s.replace(/\\/g, "\\\\").replace(/'/g, "\\'");

      const parts: string[] = [];
      if (args.query) parts.push(`name contains '${esc(args.query)}'`);
      if (args.mimeType) {
        const resolved = MIME_SHORTCUTS[args.mimeType] ?? args.mimeType;
        parts.push(`mimeType = '${esc(resolved)}'`);
      }
      if (args.modifiedAfter) {
        // Validate ISO date format to prevent query injection
        const dateStr = args.modifiedAfter.replace(/[^0-9T:\-Z.]/g, "");
        parts.push(`modifiedTime > '${dateStr}'`);
      }
      if (args.owner) parts.push(`'${esc(args.owner)}' in owners`);
      parts.push("trashed = false");

      const q = parts.join(" and ");

      const res = await drive.files.list({
        q,
        pageSize: args.pageSize,
        pageToken: args.pageToken,
        fields: "nextPageToken, files(id, name, mimeType, modifiedTime, webViewLink, owners)",
      });

      const files = res.data.files ?? [];
      if (files.length === 0) return formatSuccess("No files found.");

      const lines = files.map(
        (f) =>
          `${f.name ?? "(untitled)"}\n  id: ${f.id}\n  type: ${f.mimeType}\n  modified: ${f.modifiedTime ?? "unknown"}\n  url: ${f.webViewLink ?? "n/a"}`
      );

      const nextPage = res.data.nextPageToken
        ? `\nnextPageToken: ${res.data.nextPageToken}`
        : "";

      return formatSuccess(lines.join("\n\n") + nextPage);
    })
  );

  // ── drive_get_info ───────────────────────────────────────────────────────────
  server.tool(
    "drive_get_info",
    "Get file metadata: name, type, size, dates, owner, sharing, parent.",
    {
      fileId: z.string().describe("File ID or URL"),
    },
    withErrorHandling(async (args) => {
      const drive = await getDriveClient();
      const id = extractFileId(args.fileId);

      const res = await drive.files.get({
        fileId: id,
        fields:
          "id, name, mimeType, size, createdTime, modifiedTime, owners, shared, sharingUser, parents, webViewLink, capabilities",
      });

      const f = res.data;
      const owner = f.owners?.[0]?.emailAddress ?? "unknown";
      const parent = f.parents?.[0] ?? "root";

      const lines = [
        `Name:     ${f.name}`,
        `ID:       ${f.id}`,
        `Type:     ${f.mimeType}`,
        `Size:     ${f.size ? `${f.size} bytes` : "n/a (Google format)"}`,
        `Created:  ${f.createdTime}`,
        `Modified: ${f.modifiedTime}`,
        `Owner:    ${owner}`,
        `Shared:   ${f.shared ? "yes" : "no"}`,
        `Parent:   ${parent}`,
        `URL:      ${f.webViewLink ?? "n/a"}`,
      ];

      return formatSuccess(lines.join("\n"));
    })
  );

  // ── drive_create_folder ─────────────────────────────────────────────────────
  server.tool(
    "drive_create_folder",
    "Create a new folder in Drive. Returns folder id and URL.",
    {
      name: z.string().describe("Folder name"),
      parentId: z.string().optional().describe("Parent folder ID or URL (defaults to My Drive root)"),
    },
    withErrorHandling(async (args) => {
      const drive = await getDriveClient();

      const metadata: {
        name: string;
        mimeType: string;
        parents?: string[];
      } = {
        name: args.name,
        mimeType: "application/vnd.google-apps.folder",
      };

      if (args.parentId) {
        metadata.parents = [extractFileId(args.parentId)];
      }

      const res = await drive.files.create({
        requestBody: metadata,
        fields: "id, name, webViewLink",
      });

      return formatSuccess(
        `Folder created: ${res.data.name}\nid: ${res.data.id}\nurl: ${res.data.webViewLink}`
      );
    })
  );

  // ── drive_upload ─────────────────────────────────────────────────────────────
  server.tool(
    "drive_upload",
    "Upload a local file to Drive. Returns file ID and URL.",
    {
      localPath: z.string().describe("Absolute local file path"),
      name: z.string().optional().describe("Override file name in Drive"),
      parentId: z.string().optional().describe("Destination folder ID or URL"),
      mimeType: z.string().optional().describe("Override MIME type (auto-detected if omitted)"),
    },
    withErrorHandling(async (args) => {
      if (!path.isAbsolute(args.localPath)) {
        return formatError("localPath must be an absolute path.");
      }
      if (!fs.existsSync(args.localPath)) {
        return formatError(`File not found: ${args.localPath}`);
      }

      const drive = await getDriveClient();
      const fileName = args.name ?? path.basename(args.localPath);
      const detectedMime = args.mimeType ?? guessMimeType(args.localPath);

      const metadata: {
        name: string;
        parents?: string[];
      } = { name: fileName };

      if (args.parentId) {
        metadata.parents = [extractFileId(args.parentId)];
      }

      const res = await drive.files.create({
        requestBody: metadata,
        media: {
          mimeType: detectedMime,
          body: fs.createReadStream(args.localPath),
        },
        fields: "id, name, webViewLink",
      });

      return formatSuccess(
        `Uploaded: ${res.data.name}\nid: ${res.data.id}\nurl: ${res.data.webViewLink}`
      );
    })
  );

  // ── drive_download ──────────────────────────────────────────────────────────
  server.tool(
    "drive_download",
    "Download a file to local path. Google-native files need exportFormat (pdf/docx/xlsx/pptx/csv/txt/html).",
    {
      fileId: z.string().describe("File ID or URL"),
      localPath: z.string().describe("Absolute local destination path"),
      exportFormat: z
        .enum(["pdf", "docx", "xlsx", "pptx", "csv", "txt", "html"])
        .optional()
        .describe("Required for Google-native files"),
    },
    withErrorHandling(async (args) => {
      if (!path.isAbsolute(args.localPath)) {
        return formatError("localPath must be an absolute path.");
      }

      const drive = await getDriveClient();
      const id = extractFileId(args.fileId);

      // Get file metadata to check mimeType
      const meta = await drive.files.get({ fileId: id, fields: "mimeType, name" });
      const fileMime = meta.data.mimeType ?? "";

      const isGoogleNative = fileMime.startsWith("application/vnd.google-apps.");
      let response;

      if (isGoogleNative) {
        if (!args.exportFormat) {
          return formatError(
            `This is a Google-native file (${fileMime}). Specify exportFormat: pdf, docx, xlsx, pptx, csv, txt, or html.`
          );
        }
        const exportMimeMap = EXPORT_MIME[fileMime];
        if (!exportMimeMap) {
          return formatError(`No export options available for ${fileMime}.`);
        }
        const exportMime = exportMimeMap[args.exportFormat];
        if (!exportMime) {
          return formatError(
            `Format "${args.exportFormat}" is not supported for ${fileMime}. Supported: ${Object.keys(exportMimeMap).join(", ")}`
          );
        }
        response = await drive.files.export(
          { fileId: id, mimeType: exportMime },
          { responseType: "stream" }
        );
      } else {
        response = await drive.files.get(
          { fileId: id, alt: "media" },
          { responseType: "stream" }
        );
      }

      // Ensure destination directory exists
      const dir = path.dirname(args.localPath);
      if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
      }

      await new Promise<void>((resolve, reject) => {
        const source = response.data as NodeJS.ReadableStream;
        const dest = fs.createWriteStream(args.localPath);
        source.on("error", reject);
        source.pipe(dest);
        dest.on("finish", resolve);
        dest.on("error", reject);
      });

      return formatSuccess(`Downloaded to: ${args.localPath}`);
    })
  );

  // ── drive_move ───────────────────────────────────────────────────────────────
  server.tool(
    "drive_move",
    "Move a file to a different folder.",
    {
      fileId: z.string().describe("File ID or URL"),
      destinationId: z.string().describe("Destination folder ID or URL"),
    },
    withErrorHandling(async (args) => {
      const drive = await getDriveClient();
      const id = extractFileId(args.fileId);
      const destId = extractFileId(args.destinationId);

      // Get current parents
      const meta = await drive.files.get({ fileId: id, fields: "parents" });
      const currentParents = (meta.data.parents ?? []).join(",");

      await drive.files.update({
        fileId: id,
        addParents: destId,
        removeParents: currentParents,
        fields: "id, parents",
      });

      return formatSuccess(`Moved ${id} to folder ${destId}`);
    })
  );

  // ── drive_copy ───────────────────────────────────────────────────────────────
  server.tool(
    "drive_copy",
    "Copy a file. Optional new name and destination folder.",
    {
      fileId: z.string().describe("File ID or URL"),
      name: z.string().optional().describe("New file name"),
      destinationId: z.string().optional().describe("Destination folder ID or URL"),
    },
    withErrorHandling(async (args) => {
      const drive = await getDriveClient();
      const id = extractFileId(args.fileId);

      const requestBody: { name?: string; parents?: string[] } = {};
      if (args.name) requestBody.name = args.name;
      if (args.destinationId) requestBody.parents = [extractFileId(args.destinationId)];

      const res = await drive.files.copy({
        fileId: id,
        requestBody,
        fields: "id, name, webViewLink",
      });

      return formatSuccess(
        `Copied: ${res.data.name}\nid: ${res.data.id}\nurl: ${res.data.webViewLink}`
      );
    })
  );

  // ── drive_rename ─────────────────────────────────────────────────────────────
  server.tool(
    "drive_rename",
    "Rename a file or folder.",
    {
      fileId: z.string().describe("File ID or URL"),
      name: z.string().describe("New name"),
    },
    withErrorHandling(async (args) => {
      const drive = await getDriveClient();
      const id = extractFileId(args.fileId);

      await drive.files.update({
        fileId: id,
        requestBody: { name: args.name },
        fields: "id, name",
      });

      return formatSuccess(`Renamed to: ${args.name}`);
    })
  );

  // ── drive_trash ──────────────────────────────────────────────────────────────
  server.tool(
    "drive_trash",
    "DESTRUCTIVE: Move file to trash (recoverable). Confirm with user first.",
    {
      fileId: z.string().describe("File ID or URL"),
    },
    withErrorHandling(async (args) => {
      const drive = await getDriveClient();
      const id = extractFileId(args.fileId);

      await drive.files.update({
        fileId: id,
        requestBody: { trashed: true },
        fields: "id, trashed",
      });

      return formatSuccess(`Moved to trash: ${id}`);
    })
  );

  // ── drive_share ──────────────────────────────────────────────────────────────
  server.tool(
    "drive_share",
    "Share a file by email (reader/commenter/writer/owner). Owner transfer is IRREVERSIBLE. Confirm with user first.",
    {
      fileId: z.string().describe("File ID or URL"),
      email: z.string().email().describe("Recipient email"),
      role: z
        .enum(["reader", "commenter", "writer", "owner"])
        .default("reader"),
      notify: z
        .boolean()
        .default(true)
        .describe("Send notification email to recipient"),
    },
    withErrorHandling(async (args) => {
      const drive = await getDriveClient();
      const id = extractFileId(args.fileId);

      const res = await drive.permissions.create({
        fileId: id,
        sendNotificationEmail: args.notify,
        transferOwnership: args.role === "owner" ? true : undefined,
        requestBody: {
          type: "user",
          role: args.role,
          emailAddress: args.email,
        },
        fields: "id, role, emailAddress",
      });

      return formatSuccess(
        `Shared with ${args.email} as ${res.data.role}\npermissionId: ${res.data.id}`
      );
    })
  );

  // ── drive_update_permission ──────────────────────────────────────────────────
  server.tool(
    "drive_update_permission",
    "Change a user's permission role on a file.",
    {
      fileId: z.string().describe("File ID or URL"),
      permissionId: z.string().describe("Permission ID (from drive_list_permissions)"),
      role: z.enum(["reader", "commenter", "writer", "owner"]),
    },
    withErrorHandling(async (args) => {
      const drive = await getDriveClient();
      const id = extractFileId(args.fileId);

      const res = await drive.permissions.update({
        fileId: id,
        permissionId: args.permissionId,
        transferOwnership: args.role === "owner" ? true : undefined,
        requestBody: { role: args.role },
        fields: "id, role, emailAddress",
      });

      return formatSuccess(
        `Updated permission ${res.data.id} to ${res.data.role} for ${res.data.emailAddress}`
      );
    })
  );

  // ── drive_list_permissions ───────────────────────────────────────────────────
  server.tool(
    "drive_list_permissions",
    "List all users/roles with access to a file.",
    {
      fileId: z.string().describe("File ID or URL"),
    },
    withErrorHandling(async (args) => {
      const drive = await getDriveClient();
      const id = extractFileId(args.fileId);

      const res = await drive.permissions.list({
        fileId: id,
        fields: "permissions(id, type, role, emailAddress, displayName)",
      });

      const perms = res.data.permissions ?? [];
      if (perms.length === 0) return formatSuccess("No permissions found.");

      const lines = perms.map(
        (p) =>
          `${p.displayName ?? p.emailAddress ?? p.type}\n  permissionId: ${p.id}\n  role: ${p.role}\n  type: ${p.type}`
      );

      return formatSuccess(lines.join("\n\n"));
    })
  );
}
