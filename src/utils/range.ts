/**
 * A1 notation utilities for Google Sheets range conversion.
 *
 * Handles both sheet-qualified ranges (e.g. "Sheet1!A1:B10") and bare ranges
 * (e.g. "A1:B10"), as well as single-cell references (e.g. "A1").
 */

/**
 * Quotes a sheet name for use in A1 notation if it contains special
 * characters (spaces, apostrophes, etc.). Plain names are left bare.
 */
export function quoteSheetName(name: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) return name;
  return `'${name.replace(/'/g, "''")}'`;
}

export interface ParsedA1Range {
  sheetName?: string;
  startRow: number;
  startCol: number;
  endRow?: number;
  endCol?: number;
}

/**
 * Converts a 0-indexed column number to its A1 letter representation.
 * Examples: 0 → "A", 1 → "B", 25 → "Z", 26 → "AA", 27 → "AB"
 */
export function columnToLetter(col: number): string {
  let result = "";
  let n = col + 1; // convert to 1-based
  while (n > 0) {
    const remainder = (n - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

/**
 * Converts a column letter (or multi-letter, e.g. "AA") to a 0-indexed column number.
 * Examples: "A" → 0, "B" → 1, "Z" → 25, "AA" → 26
 */
export function letterToColumn(letter: string): number {
  const upper = letter.toUpperCase();
  let result = 0;
  for (let i = 0; i < upper.length; i++) {
    result = result * 26 + (upper.charCodeAt(i) - 64);
  }
  return result - 1; // convert to 0-based
}

/**
 * Parses an A1 notation string into its component parts.
 *
 * Supports:
 *   - "A1"             → single cell
 *   - "A1:B10"         → range
 *   - "Sheet1!A1:B10"  → sheet-qualified range
 *   - "Sheet1!A1"      → sheet-qualified single cell
 *   - "A:B"            → entire columns
 *   - "1:10"           → entire rows
 */
export function parseA1Notation(range: string): ParsedA1Range {
  let sheetName: string | undefined;
  let rangeStr = range;

  // Extract sheet name if present (format: "SheetName!Range" or "'Sheet Name'!Range")
  // Use lastIndexOf to handle sheet names containing '!' (e.g. "'Sheet!1'!A1")
  const sheetSep = range.lastIndexOf("!");
  if (sheetSep !== -1) {
    sheetName = range.slice(0, sheetSep)
      .replace(/^'|'$/g, "")  // strip surrounding quotes
      .replace(/''/g, "'");    // unescape doubled apostrophes
    rangeStr = range.slice(sheetSep + 1);
  }

  const parts = rangeStr.split(":");
  const startRef = parts[0];
  const endRef = parts[1];

  const parseRef = (ref: string): { row?: number; col?: number } => {
    const match = ref.match(/^([A-Za-z]*)(\d*)$/);
    if (!match) return {};
    const colPart = match[1];
    const rowPart = match[2];
    return {
      col: colPart ? letterToColumn(colPart) : undefined,
      row: rowPart ? parseInt(rowPart, 10) - 1 : undefined, // convert to 0-based
    };
  };

  const start = parseRef(startRef);
  const result: ParsedA1Range = {
    sheetName,
    startRow: start.row ?? 0,
    startCol: start.col ?? 0,
  };

  if (endRef) {
    const end = parseRef(endRef);
    result.endRow = end.row;
    result.endCol = end.col;
  }

  return result;
}

/**
 * Converts an A1 notation string to a Google Sheets GridRange object.
 *
 * GridRange uses 0-based, exclusive end indices.
 * endRowIndex and endColumnIndex are omitted when the range extends to the
 * end of the sheet (i.e. when no end row/col is specified).
 *
 * @param range  - A1 notation range string
 * @param sheetId - The numeric sheet ID (defaults to 0 for the first sheet)
 */
export function toGridRange(
  range: string,
  sheetId: number = 0
): {
  sheetId: number;
  startRowIndex: number;
  startColumnIndex: number;
  endRowIndex?: number;
  endColumnIndex?: number;
} {
  const parsed = parseA1Notation(range);

  // Determine if this is a single-cell reference (no end row or column parsed).
  // For ranges like "A:B" or "1:10", at least one end dimension will be set.
  const isSingleCell = parsed.endRow === undefined && parsed.endCol === undefined;

  const result: {
    sheetId: number;
    startRowIndex: number;
    startColumnIndex: number;
    endRowIndex?: number;
    endColumnIndex?: number;
  } = {
    sheetId,
    startRowIndex: parsed.startRow,
    startColumnIndex: parsed.startCol,
  };

  if (parsed.endRow !== undefined) {
    result.endRowIndex = parsed.endRow + 1;
  } else if (isSingleCell) {
    // Single-cell reference (no colon): close the range to one cell.
    // Without this, batchUpdate operations apply to the entire remainder
    // of the sheet from the start cell.
    result.endRowIndex = parsed.startRow + 1;
  }
  // else: open-ended range (e.g. "A:B") - leave endRowIndex undefined

  if (parsed.endCol !== undefined) {
    result.endColumnIndex = parsed.endCol + 1;
  } else if (isSingleCell) {
    result.endColumnIndex = parsed.startCol + 1;
  }

  return result;
}
