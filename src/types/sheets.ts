/**
 * Shared TypeScript interfaces for Google Sheets operations.
 *
 * These types map to the structures used by the Sheets API v4 and are
 * referenced by multiple tool handlers throughout the project.
 */

/** A single cell value - can be a string, number, boolean, or null (empty). */
export type CellValue = string | number | boolean | null;

/** A row of cell values. */
export type RowValues = CellValue[];

/** A two-dimensional array of cell values representing a range of data. */
export type RangeValues = RowValues[];

/**
 * Represents a parsed cell style definition used when applying formatting.
 * All fields are optional - only provided fields will be written to the cell.
 */
export interface CellStyle {
  /** Background color as hex string or named color (e.g. "#FF0000", "red"). */
  backgroundColor?: string;
  /** Text color as hex string or named color. */
  textColor?: string;
  /** Font size in points. */
  fontSize?: number;
  /** Whether the text should be bold. */
  bold?: boolean;
  /** Whether the text should be italic. */
  italic?: boolean;
  /** Whether the text should be underlined. */
  underline?: boolean;
  /** Whether the text should be struck through. */
  strikethrough?: boolean;
  /**
   * Horizontal alignment of cell content.
   * Matches the Sheets API HorizontalAlign enum values.
   */
  horizontalAlignment?: "LEFT" | "CENTER" | "RIGHT";
  /**
   * Vertical alignment of cell content.
   * Matches the Sheets API VerticalAlign enum values.
   */
  verticalAlignment?: "TOP" | "MIDDLE" | "BOTTOM";
  /** Whether to wrap text within the cell. */
  wrapStrategy?: "OVERFLOW_CELL" | "LEGACY_WRAP" | "CLIP" | "WRAP";
  /** Number format pattern (e.g. "#,##0.00", "MM/dd/yyyy"). */
  numberFormat?: string;
  /**
   * Number format type.
   * Matches the Sheets API NumberFormatType enum values.
   */
  numberFormatType?:
    | "TEXT"
    | "NUMBER"
    | "PERCENT"
    | "CURRENCY"
    | "DATE"
    | "TIME"
    | "DATE_TIME"
    | "SCIENTIFIC";
}

/**
 * Represents a Google Sheets spreadsheet (minimal fields for display/lookup).
 */
export interface SpreadsheetInfo {
  spreadsheetId: string;
  title: string;
  sheets: SheetInfo[];
  spreadsheetUrl: string;
}

/** Represents a single sheet (tab) within a spreadsheet. */
export interface SheetInfo {
  sheetId: number;
  title: string;
  index: number;
  sheetType: "GRID" | "OBJECT" | "DATA_SOURCE";
  rowCount: number;
  columnCount: number;
}

/**
 * Parameters for a batch update operation, used internally when constructing
 * batchUpdate requests.
 */
export interface BatchUpdateRequest {
  requests: object[];
}
