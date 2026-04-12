/**
 * Shared sheet resolution utilities.
 *
 * Resolves a sheet name or numeric sheetId to a numeric sheet ID,
 * used by tools that need to target specific sheets in batchUpdate requests.
 */

import { getSheetsClient } from "../client/google-client.js";
import { parseA1Notation } from "./range.js";

/**
 * Resolves a sheet name or numeric sheetId to a numeric sheet ID.
 * Throws if the sheet is not found.
 */
export async function resolveSheetId(
  spreadsheetId: string,
  nameOrId: string | number
): Promise<number> {
  const sheets = await getSheetsClient();
  const res = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: "sheets.properties(sheetId,title)",
  });

  const sheetList = res.data.sheets ?? [];

  if (typeof nameOrId === "number") {
    const found = sheetList.find((s) => s.properties?.sheetId === nameOrId);
    if (!found) throw new Error(`Sheet ID ${nameOrId} not found.`);
    return nameOrId;
  }

  // Try parsing as integer first
  const asInt = parseInt(nameOrId, 10);
  if (!isNaN(asInt) && String(asInt) === nameOrId) {
    const found = sheetList.find((s) => s.properties?.sheetId === asInt);
    if (found) return asInt;
  }

  // Match by title (case-insensitive)
  const byName = sheetList.find(
    (s) => s.properties?.title?.toLowerCase() === nameOrId.toLowerCase()
  );
  if (byName?.properties?.sheetId === undefined) {
    throw new Error(`Sheet "${nameOrId}" not found.`);
  }
  return byName.properties.sheetId as number;
}

/**
 * Resolve sheetId from a range string (uses the sheet name portion if present,
 * otherwise fetches the first sheet's actual ID).
 */
export async function resolveSheetIdFromRange(
  spreadsheetId: string,
  range: string
): Promise<number> {
  const parsed = parseA1Notation(range);
  if (parsed.sheetName) {
    return resolveSheetId(spreadsheetId, parsed.sheetName);
  }
  // Fetch the actual first sheet's ID (not always 0)
  const sheets = await getSheetsClient();
  const res = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: "sheets.properties.sheetId",
  });
  return res.data.sheets?.[0]?.properties?.sheetId ?? 0;
}

/**
 * Cached variant of resolveSheetIdFromRange for bulk tools.
 * Pass a Map that persists for the duration of one tool call to avoid
 * redundant API round-trips when many ranges target the same sheet.
 */
export async function resolveSheetIdCached(
  spreadsheetId: string,
  range: string,
  cache: Map<string, number>
): Promise<number> {
  const parsed = parseA1Notation(range);
  const key = parsed.sheetName ?? "__default__";
  const cached = cache.get(key);
  if (cached !== undefined) return cached;
  const resolved = await resolveSheetIdFromRange(spreadsheetId, range);
  cache.set(key, resolved);
  return resolved;
}
