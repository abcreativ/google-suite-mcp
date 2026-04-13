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
    "Searches Google Drive files using drive.files.list with a query built from name, MIME type, modification date, and owner filters; excludes trashed files automatically and supports pagination. Use when the user asks to find a file or folder by name or type in Drive. Use when discovering file IDs before operating on files with other drive_ tools. Do not use when: retrieving metadata for a known file ID - use drive_get_info instead; listing spreadsheets only - use sheets_list instead; getting info about a specific file - use drive_get_info instead. Returns: one block per file with name, id, type, modified timestamp, and url; blocks separated by blank lines. Appends 'nextPageToken: ...' when more pages exist. Parameters: - query: text to match in file names, e.g. 'budget 2025' - mimeType: sheet, doc, folder, slide, form, or a full MIME type string - modifiedAfter: ISO date string, e.g. '2024-01-01' - pageSize: results per page (default 20, max 100).",
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
    "Retrieves detailed metadata for a specific Drive file or folder using drive.files.get, returning name, MIME type, size, creation/modification dates, owner email, sharing status, parent folder ID, and web URL. Use when the user asks for details about a known file, such as confirming its owner or finding its parent folder. Use when verifying a file exists and getting its URL before sharing it or operating on it. Do not use when: searching for files by name or type - use drive_search instead; listing all spreadsheets - use sheets_list instead; listing file permissions - use drive_list_permissions instead. Returns: multi-line string with Name, ID, Type, Size, Created, Modified, Owner, Shared, Parent, and URL fields. Parameters: - fileId: Drive file ID or full URL.",
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
    "Creates a new folder in Google Drive using drive.files.create with the folder MIME type; optionally places it inside a parent folder. Use when the user asks to create a new folder for organizing files. Use when creating a destination folder before uploading files with drive_upload or moving files with drive_move. Do not use when: uploading a file - use drive_upload instead; moving an existing file - use drive_move instead; copying a file - use drive_copy instead; renaming a file or folder - use drive_rename instead; trashing a folder - use drive_trash instead. Returns: 'Folder created: {name}\\nid: {id}\\nurl: {url}'. Parameters: - name: folder display name - parentId: parent folder ID or URL (optional; defaults to My Drive root).",
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
    "Uploads a local file to Google Drive using drive.files.create with a multipart media upload; auto-detects the MIME type from the file extension unless overridden. Use when the user asks to upload a local file to Drive, such as a CSV, PDF, or image. Use when placing a file into a specific folder before sharing it with drive_share. Do not use when: downloading a file from Drive - use drive_download instead; creating a new folder - use drive_create_folder instead; copying an existing Drive file - use drive_copy instead; moving a file to a different folder - use drive_move instead. Returns: 'Uploaded: {name}\\nid: {id}\\nurl: {url}'. Parameters: - localPath: absolute path to the local file to upload - name: override the Drive file name (optional; defaults to the local filename) - parentId: destination folder ID or URL (optional) - mimeType: override MIME type (optional; auto-detected from extension if omitted).",
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
    "Downloads a file from Drive to a local path using drive.files.get with alt=media; Google-native files (Docs, Sheets, Slides) require an exportFormat because they cannot be downloaded in their native format. Use when the user asks to download a Drive file to disk. Use when exporting a Google Sheet as CSV or a Google Doc as PDF for offline use. Do not use when: uploading a file to Drive - use drive_upload instead; reading a Google Doc's text in-session - use docs_get_text instead; exporting a Google Doc in-session - use docs_export instead. Returns: 'Downloaded to: {localPath}'. Parameters: - fileId: Drive file ID or URL - localPath: absolute local destination path (directory is created if it does not exist) - exportFormat: required for Google-native files; one of pdf, docx, xlsx, pptx, csv, txt, html; omit for non-native files like PDFs and images.",
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
    "Moves a file to a different folder in Drive using drive.files.update by adding the new parent and removing all current parents; the file ID and content are unchanged. Use when the user asks to relocate a file from one folder to another. Use when organizing uploaded files into a project folder after creating it with drive_create_folder. Do not use when: copying a file without removing the original - use drive_copy instead; renaming a file - use drive_rename instead; trashing a file - use drive_trash instead; creating a folder - use drive_create_folder instead; uploading a new file - use drive_upload instead; downloading a file - use drive_download instead. Returns: 'Moved {fileId} to folder {destinationFolderId}'. Parameters: - fileId: Drive file ID or URL of the file to move - destinationId: Drive folder ID or URL of the destination folder.",
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
    "Creates a copy of a Drive file using drive.files.copy; the original is not moved or modified. Use when the user asks to duplicate a file, such as creating a backup before editing or cloning a template. Use when copying a file to a different folder by specifying destinationId. Do not use when: copying a full spreadsheet and getting a Sheets-specific result - prefer sheets_copy for spreadsheets; moving a file without keeping the original - use drive_move instead; renaming without copying - use drive_rename instead; creating a new folder - use drive_create_folder instead; uploading a local file - use drive_upload instead; trashing a file - use drive_trash instead. Returns: 'Copied: {name}\\nid: {newFileId}\\nurl: {url}'. Parameters: - fileId: source Drive file ID or URL - name: display name for the copy (optional; defaults to 'Copy of {original}') - destinationId: destination folder ID or URL (optional; defaults to same folder as original).",
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
    "Renames a file or folder in Drive using drive.files.update with a new name field; the file ID, content, and location are unchanged. Use when the user asks to rename an existing file or folder. Use when correcting a display name before sharing a file. Do not use when: copying a file and giving the copy a new name - use drive_copy instead; moving a file to a different folder - use drive_move instead; trashing a file - use drive_trash instead; creating a new folder - use drive_create_folder instead; uploading a file - use drive_upload instead; downloading a file - use drive_download instead. Returns: 'Renamed to: {name}'. Parameters: - fileId: Drive file ID or URL - name: new display name for the file or folder.",
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
    "Moves a file or folder to the Google Drive Bin using drive.files.update with trashed=true; the item is recoverable from the Bin for 30 days before permanent deletion. Use when the user asks to delete or trash a file and does not need it immediately. Use when removing a file you plan to restore within 30 days. Do not use when: permanently deleting without recovery - warn the user that Drive Bin holds items 30 days; renaming a file - use drive_rename instead; moving a file to a different folder - use drive_move instead; copying a file - use drive_copy instead; sharing a file - use drive_share instead; listing permissions - use drive_list_permissions instead. Returns: 'Moved to trash: {id}'. Parameters: - fileId: Drive file ID or URL of the file or folder to trash.",
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
    "Grants access to a Drive file for a specific user via drive.permissions.create; sets role to reader, commenter, writer, or owner. Owner transfer is IRREVERSIBLE and cannot be undone. Use when the user asks to share a file with a collaborator by email. Use when granting view-only access to a report or write access to a shared project file. Do not use when: changing the role of an existing collaborator - use drive_update_permission instead; listing who has access - use drive_list_permissions instead; trashing a file - use drive_trash instead; moving a file - use drive_move instead; renaming a file - use drive_rename instead. Returns: 'Shared with {email} as {role}\\npermissionId: {id}'. Parameters: - fileId: Drive file ID or URL - email: recipient email address - role: reader, commenter, writer, or owner (default reader; owner transfer is irreversible) - notify: send notification email to recipient (default true).",
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
    "Changes the role of an existing permission on a Drive file using drive.permissions.update; requires a permissionId obtained from drive_share or drive_list_permissions. Owner promotion is IRREVERSIBLE. Use when the user asks to upgrade a viewer to editor or downgrade a writer to reader on a file. Use when adjusting access for an existing collaborator without re-inviting them. Do not use when: granting access to a new user - use drive_share instead; listing current permissions to find the permissionId - use drive_list_permissions instead; trashing a file - use drive_trash instead; renaming a file - use drive_rename instead; moving a file - use drive_move instead. Returns: 'Updated permission {id} to {role} for {email}'. Parameters: - fileId: Drive file ID or URL - permissionId: permission ID from drive_list_permissions or drive_share - role: new role - reader, commenter, writer, or owner (owner promotion is irreversible).",
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
    "Retrieves all permissions on a Drive file using drive.permissions.list, returning each collaborator's display name, permissionId, role, and type. Use when the user asks who has access to a file or wants to audit sharing settings. Use when retrieving permissionIds before calling drive_update_permission to change a collaborator's role. Do not use when: granting new access - use drive_share instead; changing an existing permission - use drive_update_permission instead; getting general file metadata - use drive_get_info instead; searching for files - use drive_search instead. Returns: one block per permission with displayName (or email or type), permissionId, role, and type; blocks separated by blank lines. Parameters: - fileId: Drive file ID or URL.",
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
