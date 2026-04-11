/**
 * Google API client factory.
 *
 * Provides lazy singleton instances of the Sheets, Drive, and Docs API
 * clients, all sharing a single authenticated OAuth2Client.
 *
 * Usage:
 *   const sheets = await getSheetsClient();
 *   const drive  = await getDriveClient();
 *   const docs   = await getDocsClient();
 */

import { google } from "googleapis";
import type { sheets_v4, drive_v3, docs_v1, script_v1 } from "googleapis";
import { getAuthClient } from "../auth/oauth.js";

// ─── Singletons ───────────────────────────────────────────────────────────────

let sheetsInstance: sheets_v4.Sheets | null = null;
let driveInstance: drive_v3.Drive | null = null;
let docsInstance: docs_v1.Docs | null = null;
let scriptInstance: script_v1.Script | null = null;

/**
 * Returns a lazy singleton Sheets v4 API client.
 */
export async function getSheetsClient(): Promise<sheets_v4.Sheets> {
  if (!sheetsInstance) {
    const auth = await getAuthClient();
    sheetsInstance = google.sheets({ version: "v4", auth });
  }
  return sheetsInstance;
}

/**
 * Returns a lazy singleton Drive v3 API client.
 */
export async function getDriveClient(): Promise<drive_v3.Drive> {
  if (!driveInstance) {
    const auth = await getAuthClient();
    driveInstance = google.drive({ version: "v3", auth });
  }
  return driveInstance;
}

/**
 * Returns a lazy singleton Docs v1 API client.
 */
export async function getDocsClient(): Promise<docs_v1.Docs> {
  if (!docsInstance) {
    const auth = await getAuthClient();
    docsInstance = google.docs({ version: "v1", auth });
  }
  return docsInstance;
}

/**
 * Returns a lazy singleton Apps Script v1 API client.
 */
export async function getScriptClient(): Promise<script_v1.Script> {
  if (!scriptInstance) {
    const auth = await getAuthClient();
    scriptInstance = google.script({ version: "v1", auth });
  }
  return scriptInstance;
}

// ─── Utilities ────────────────────────────────────────────────────────────────

/**
 * Extracts a Google file ID from a URL or returns the input as-is if it looks
 * like a raw ID already.
 *
 * Handles:
 *   - Raw IDs:                  "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms"
 *   - Spreadsheet URLs:         "https://docs.google.com/spreadsheets/d/{ID}/edit"
 *   - Document URLs:            "https://docs.google.com/document/d/{ID}/edit"
 *   - Drive file URLs:          "https://drive.google.com/file/d/{ID}/view"
 *   - Drive open URLs:          "https://drive.google.com/open?id={ID}"
 *   - Drive sharing URLs:       "https://drive.google.com/uc?id={ID}"
 */
export function extractFileId(input: string): string {
  const trimmed = input.trim();

  // Drive open/sharing URLs with ?id= query parameter
  const idParam = new URLSearchParams(
    trimmed.includes("?") ? trimmed.split("?")[1] : ""
  ).get("id");
  if (idParam) return idParam;

  // URLs with /d/{ID}/ path segment (Sheets, Docs, Drive file URLs)
  const dMatch = trimmed.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (dMatch) return dMatch[1];

  // Drive folder URLs: /drive/folders/{ID}
  const folderMatch = trimmed.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (folderMatch) return folderMatch[1];

  // Looks like a raw ID (alphanumeric + underscores/hyphens, 20+ chars)
  if (/^[a-zA-Z0-9_-]{20,}$/.test(trimmed)) {
    return trimmed;
  }

  // Return as-is and let the API reject it with a meaningful error
  return trimmed;
}
