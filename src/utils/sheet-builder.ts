/**
 * Shared helpers for high-level sheet-building tools.
 *
 * Extracted from dashboard.ts so that sheets_write_table, sheets_build_sheet,
 * and sheets_build_dashboard can all reuse the same primitives.
 */

import { parseColor } from "./color.js";
import type { sheets_v4 } from "googleapis";

// ─── Resolve or create a sheet by name ───────────────────────────────────────

export async function resolveOrCreateSheet(
  sheets: sheets_v4.Sheets,
  spreadsheetId: string,
  sheetName: string
): Promise<number> {
  const res = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: "sheets.properties(sheetId,title)",
  });

  const existing = (res.data.sheets ?? []).find(
    (s) => s.properties?.title?.toLowerCase() === sheetName.toLowerCase()
  );

  if (existing?.properties?.sheetId !== undefined) {
    return existing.properties.sheetId as number;
  }

  const addRes = await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [{ addSheet: { properties: { title: sheetName } } }],
    },
  });

  const newSheet = addRes.data.replies?.[0]?.addSheet?.properties;
  if (newSheet?.sheetId === undefined) {
    throw new Error(`Failed to create sheet "${sheetName}"`);
  }
  return newSheet.sheetId as number;
}

// ─── GridRange builder ───────────────────────────────────────────────────────

export function makeGridRange(
  sheetId: number,
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number
): sheets_v4.Schema$GridRange {
  return {
    sheetId,
    startRowIndex: startRow,
    startColumnIndex: startCol,
    endRowIndex: endRow,
    endColumnIndex: endCol,
  };
}

// ─── repeatCell request builder ──────────────────────────────────────────────

export function repeatCellRequest(
  sheetId: number,
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number,
  format: sheets_v4.Schema$CellFormat,
  fields: string
): sheets_v4.Schema$Request {
  return {
    repeatCell: {
      range: makeGridRange(sheetId, startRow, startCol, endRow, endCol),
      cell: { userEnteredFormat: format },
      fields,
    },
  };
}

// ─── Default color palette ───────────────────────────────────────────────────

export const DEFAULT_COLORS = {
  HEADER_BG: parseColor("#1a237e"),       // deep navy
  HEADER_FG: parseColor("#ffffff"),       // white
  KPI_BG: parseColor("#e8eaf6"),          // light indigo tint
  KPI_LABEL_FG: parseColor("#3949ab"),
  ALT_ROW_BG: parseColor("#f5f5f5"),     // light grey alternate rows
  TABLE_HEADER_BG: parseColor("#3949ab"),
  TABLE_HEADER_FG: parseColor("#ffffff"),
  POSITIVE_FG: parseColor("#2e7d32"),     // green
  NEGATIVE_FG: parseColor("#c62828"),     // red
  BORDER_COLOR: parseColor("#9fa8da"),
};

// ─── Border helper ───────────────────────────────────────────────────────────

export function solidBorder(
  color = DEFAULT_COLORS.BORDER_COLOR
): sheets_v4.Schema$Border {
  return {
    style: "SOLID",
    colorStyle: { rgbColor: color },
  };
}
