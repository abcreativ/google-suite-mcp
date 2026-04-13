/**
 * Formatting tools: cell formatting, borders, merges, resize, freeze, conditional formatting.
 *
 * F4.1 - sheets_format_cells
 * F4.2 - sheets_set_borders, sheets_merge_cells, sheets_resize_columns, sheets_resize_rows, sheets_freeze_panes
 * F4.3 - sheets_add_conditional_format
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getSheetsClient, extractFileId } from "../client/google-client.js";
import { formatSuccess } from "../utils/response.js";
import { withErrorHandling } from "../utils/error-handler.js";
import { toGridRange } from "../utils/range.js";
import { parseColor } from "../utils/color.js";
import { resolveSheetId, resolveSheetIdFromRange, resolveSheetIdCached } from "../utils/sheet-resolver.js";
import type { sheets_v4 } from "googleapis";

// ─── Border style schema ──────────────────────────────────────────────────────

const BorderStyleEnum = z.enum([
  "SOLID",
  "DASHED",
  "DOTTED",
  "DOUBLE",
  "SOLID_MEDIUM",
  "SOLID_THICK",
  "NONE",
]);

const BorderSpec = z
  .object({
    style: BorderStyleEnum.optional().describe("Border line style (default: SOLID)"),
    color: z.string().optional().describe("Border color (hex or named)"),
    width: z.number().int().min(1).max(3).optional().describe("Border width in pixels (1-3)"),
  })
  .optional();

function buildBorder(
  spec: { style?: string; color?: string; width?: number } | undefined | null
): sheets_v4.Schema$Border | undefined {
  if (!spec) return undefined;
  const border: sheets_v4.Schema$Border = {
    style: spec.style ?? "SOLID",
  };
  if (spec.color) {
    border.colorStyle = { rgbColor: parseColor(spec.color) };
  }
  if (spec.width !== undefined) {
    border.width = spec.width;
  }
  return border;
}

// ─── Tool registration ────────────────────────────────────────────────────────

export function registerFormattingTools(server: McpServer): void {
  // ─── F4.1 sheets_format_cells ─────────────────────────────────────────────

  server.tool(
    "sheets_format_cells",
    "Applies visual formatting to a single contiguous cell range using spreadsheets.batchUpdate with repeatCell; only the properties you supply are changed - unspecified properties are preserved. Use when the user asks to bold a header row, change a column's number format, or set a background color on a range. Use when targeting a single contiguous range with one or more formatting properties at once. Do not use when: formatting multiple non-contiguous ranges in one call - use sheets_apply_formats instead; setting borders specifically - use sheets_set_borders (or include borders via sheets_apply_formats); merging cells - use sheets_merge_cells instead; resizing columns - use sheets_resize_columns instead; resizing rows - use sheets_resize_rows instead; freezing panes - use sheets_freeze_panes instead; adding conditional formatting - use sheets_add_conditional_format instead. Returns: 'Applied formatting to {range} ({N} field(s) updated)', or 'No formatting options provided - nothing to update.' Parameters: - range: A1 notation including sheet name, e.g. 'Sheet1!A1:D1' - number_format_pattern: format string, e.g. '$#,##0.00', '0.0%', 'yyyy-mm-dd' - horizontal_alignment: LEFT, CENTER, or RIGHT - wrap_strategy: OVERFLOW_CELL, WRAP, or CLIP.",
    {
      spreadsheet_id: z.string().describe("sheet ID from the URL (the token between /d/ and /edit) or the full URL"),
      range: z.string().describe("A1 notation range, e.g. 'Sheet1!A1:C10'"),
      bold: z.boolean().optional().describe("Bold text"),
      italic: z.boolean().optional().describe("Italic text"),
      strikethrough: z.boolean().optional().describe("Strikethrough text"),
      font_family: z.string().optional().describe("Font family (e.g. 'Arial', 'Roboto')"),
      font_size: z.number().int().min(1).optional().describe("Font size in points"),
      text_color: z.string().optional().describe("Text/foreground color (hex or named)"),
      background_color: z.string().optional().describe("Cell background color (hex or named)"),
      horizontal_alignment: z
        .enum(["LEFT", "CENTER", "RIGHT"])
        .optional()
        .describe("Horizontal text alignment"),
      vertical_alignment: z
        .enum(["TOP", "MIDDLE", "BOTTOM"])
        .optional()
        .describe("Vertical text alignment"),
      number_format_pattern: z
        .string()
        .optional()
        .describe(
          "Number format pattern, e.g. '$#,##0.00', '0.0%', 'yyyy-mm-dd', '#,##0'"
        ),
      number_format_type: z
        .enum(["TEXT", "NUMBER", "PERCENT", "CURRENCY", "DATE", "TIME", "DATE_TIME", "SCIENTIFIC"])
        .optional()
        .describe("Number format type (used together with number_format_pattern)"),
      wrap_strategy: z
        .enum(["OVERFLOW_CELL", "WRAP", "CLIP"])
        .optional()
        .describe("Text wrapping strategy"),
    },
    withErrorHandling(
      async ({
        spreadsheet_id,
        range,
        bold,
        italic,
        strikethrough,
        font_family,
        font_size,
        text_color,
        background_color,
        horizontal_alignment,
        vertical_alignment,
        number_format_pattern,
        number_format_type,
        wrap_strategy,
      }) => {
        const sheets = await getSheetsClient();
        const id = extractFileId(spreadsheet_id);
        const sheetId = await resolveSheetIdFromRange(id, range);
        const gridRange = toGridRange(range, sheetId);

        // Build the cell format and update field mask
        const cellFormat: sheets_v4.Schema$CellFormat = {};
        const updateFields: string[] = [];

        // Text format
        const textFormat: sheets_v4.Schema$TextFormat = {};
        if (bold !== undefined) { textFormat.bold = bold; updateFields.push("userEnteredFormat.textFormat.bold"); }
        if (italic !== undefined) { textFormat.italic = italic; updateFields.push("userEnteredFormat.textFormat.italic"); }
        if (strikethrough !== undefined) { textFormat.strikethrough = strikethrough; updateFields.push("userEnteredFormat.textFormat.strikethrough"); }
        if (font_family !== undefined) { textFormat.fontFamily = font_family; updateFields.push("userEnteredFormat.textFormat.fontFamily"); }
        if (font_size !== undefined) { textFormat.fontSize = font_size; updateFields.push("userEnteredFormat.textFormat.fontSize"); }
        if (text_color !== undefined) {
          textFormat.foregroundColorStyle = { rgbColor: parseColor(text_color) };
          updateFields.push("userEnteredFormat.textFormat.foregroundColorStyle");
        }
        if (Object.keys(textFormat).length > 0) {
          cellFormat.textFormat = textFormat;
        }

        if (background_color !== undefined) {
          cellFormat.backgroundColorStyle = { rgbColor: parseColor(background_color) };
          updateFields.push("userEnteredFormat.backgroundColorStyle");
        }

        if (horizontal_alignment !== undefined) {
          cellFormat.horizontalAlignment = horizontal_alignment;
          updateFields.push("userEnteredFormat.horizontalAlignment");
        }

        if (vertical_alignment !== undefined) {
          cellFormat.verticalAlignment = vertical_alignment;
          updateFields.push("userEnteredFormat.verticalAlignment");
        }

        if (number_format_pattern !== undefined || number_format_type !== undefined) {
          cellFormat.numberFormat = {
            type: number_format_type ?? "NUMBER",
            pattern: number_format_pattern ?? "",
          };
          updateFields.push("userEnteredFormat.numberFormat");
        }

        if (wrap_strategy !== undefined) {
          cellFormat.wrapStrategy = wrap_strategy;
          updateFields.push("userEnteredFormat.wrapStrategy");
        }

        if (updateFields.length === 0) {
          return formatSuccess("No formatting options provided - nothing to update.");
        }

        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: id,
          requestBody: {
            requests: [
              {
                repeatCell: {
                  range: gridRange,
                  cell: { userEnteredFormat: cellFormat },
                  fields: updateFields.join(","),
                },
              },
            ],
          },
        });

        return formatSuccess(
          `Applied formatting to ${range} (${updateFields.length} field(s) updated)`
        );
      }
    )
  );

  // ─── F4.2 sheets_set_borders ──────────────────────────────────────────────

  server.tool(
    "sheets_set_borders",
    "Sets border styles on the outer and inner edges of a single cell range using spreadsheets.batchUpdate with updateBorders; each edge (top, bottom, left, right, inner horizontal, inner vertical) can be styled independently. Use when the user asks to add a box border around a header row or draw grid lines within a table. Use when borders need precise per-edge control that sheets_format_cells or sheets_apply_formats do not expose. Do not use when: applying borders alongside cell formatting in a single call - use sheets_apply_formats instead (it combines borders and formatting per range); applying formatting without borders - use sheets_format_cells instead; merging cells - use sheets_merge_cells instead; resizing columns - use sheets_resize_columns instead; resizing rows - use sheets_resize_rows instead; adding conditional formatting - use sheets_add_conditional_format instead. Returns: 'Borders applied to {range}'. Parameters: - range: A1 notation including sheet name, e.g. 'Sheet1!A1:E10' - top/bottom/left/right/inner_horizontal/inner_vertical: each accepts style (SOLID, DASHED, DOTTED, DOUBLE, SOLID_MEDIUM, SOLID_THICK, NONE) and optional color (hex) and width_pixels.",
    {
      spreadsheet_id: z.string().describe("sheet ID from the URL (the token between /d/ and /edit) or the full URL"),
      range: z.string().describe("A1 notation range"),
      top: BorderSpec.describe("Top border"),
      bottom: BorderSpec.describe("Bottom border"),
      left: BorderSpec.describe("Left border"),
      right: BorderSpec.describe("Right border"),
      inner_horizontal: BorderSpec.describe("Inner horizontal borders"),
      inner_vertical: BorderSpec.describe("Inner vertical borders"),
    },
    withErrorHandling(
      async ({ spreadsheet_id, range, top, bottom, left, right, inner_horizontal, inner_vertical }) => {
        const sheets = await getSheetsClient();
        const id = extractFileId(spreadsheet_id);
        const sheetId = await resolveSheetIdFromRange(id, range);
        const gridRange = toGridRange(range, sheetId);

        const updateBordersRequest: sheets_v4.Schema$UpdateBordersRequest = {
          range: gridRange,
        };

        if (top) updateBordersRequest.top = buildBorder(top);
        if (bottom) updateBordersRequest.bottom = buildBorder(bottom);
        if (left) updateBordersRequest.left = buildBorder(left);
        if (right) updateBordersRequest.right = buildBorder(right);
        if (inner_horizontal) updateBordersRequest.innerHorizontal = buildBorder(inner_horizontal);
        if (inner_vertical) updateBordersRequest.innerVertical = buildBorder(inner_vertical);

        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: id,
          requestBody: {
            requests: [{ updateBorders: updateBordersRequest }],
          },
        });

        return formatSuccess(`Borders applied to ${range}`);
      }
    )
  );

  // ─── F4.2 sheets_merge_cells ──────────────────────────────────────────────

  server.tool(
    "sheets_merge_cells",
    "Merges or unmerges a cell range in a Google Sheet using spreadsheets.batchUpdate with mergeCells or unMergeCells; merged cells display their top-left cell's content and the remaining cells' content is discarded. Use when the user asks to merge a header cell across multiple columns, such as a title row spanning A1:E1. Use when unmerging a previously merged region to restore individual cell editing. Do not use when: formatting cells without merging - use sheets_format_cells instead; applying borders - use sheets_set_borders instead; resizing columns or rows - use sheets_resize_columns or sheets_resize_rows instead; freezing panes - use sheets_freeze_panes instead; applying conditional formatting - use sheets_add_conditional_format instead. Returns: 'Merged cells in {range}' or 'Unmerged cells in {range}'. Parameters: - range: A1 notation including sheet name, e.g. 'Sheet1!A1:E1' - merge_type: MERGE_ALL (all cells into one), MERGE_COLUMNS (merge within each column), or MERGE_ROWS (merge within each row) - unmerge: set to true to unmerge instead of merge.",
    {
      spreadsheet_id: z.string().describe("sheet ID from the URL (the token between /d/ and /edit) or the full URL"),
      range: z.string().describe("A1 notation range"),
      merge_type: z
        .enum(["MERGE_ALL", "MERGE_COLUMNS", "MERGE_ROWS"])
        .optional()
        .describe("How to merge cells (default: MERGE_ALL)"),
      unmerge: z
        .boolean()
        .optional()
        .describe("Unmerge cells in range instead of merging (default: false)"),
    },
    withErrorHandling(async ({ spreadsheet_id, range, merge_type, unmerge }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);
      const sheetId = await resolveSheetIdFromRange(id, range);
      const gridRange = toGridRange(range, sheetId);

      const request: sheets_v4.Schema$Request = unmerge
        ? { unmergeCells: { range: gridRange } }
        : {
            mergeCells: {
              range: gridRange,
              mergeType: merge_type ?? "MERGE_ALL",
            },
          };

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: { requests: [request] },
      });

      const action = unmerge ? "Unmerged" : "Merged";
      return formatSuccess(`${action} cells in ${range}`);
    })
  );

  // ─── F4.2 sheets_resize_columns ───────────────────────────────────────────

  server.tool(
    "sheets_resize_columns",
    "Sets the pixel width of a column range using spreadsheets.batchUpdate with updateDimensionProperties; omit pixel_size to trigger auto-resize based on cell content using autoResizeDimensions. Use when the user asks to set a specific column width or auto-fit columns to their content. Use when adjusting column widths after writing data to make a sheet readable. Do not use when: sheets_write_table or sheets_build_sheet already include auto-resize - prefer those for full table builds; resizing rows - use sheets_resize_rows instead; formatting cell content - use sheets_format_cells instead; merging cells - use sheets_merge_cells instead; freezing panes - use sheets_freeze_panes instead; applying borders - use sheets_set_borders instead. Returns: 'Resized {N} column(s) (cols {start}–{end}) to {px}px' or 'to auto'. Parameters: - start_column: 0-based inclusive start index, e.g. 0 = column A - end_column: 0-based inclusive end index, e.g. 3 = column D - pixel_size: width in pixels; omit to auto-fit.",
    {
      spreadsheet_id: z.string().describe("sheet ID from the URL (the token between /d/ and /edit) or the full URL"),
      sheet: z.string().describe("Sheet name or numeric sheet ID"),
      start_column: z.number().int().describe("0-based start column index"),
      end_column: z.number().int().describe("0-based end column index (inclusive)"),
      pixel_size: z
        .number()
        .int()
        .min(0)
        .optional()
        .describe("Column width in pixels. Omit to auto-resize."),
    },
    withErrorHandling(async ({ spreadsheet_id, sheet, start_column, end_column, pixel_size }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);
      const sheetId = await resolveSheetId(id, sheet);

      const dimensionRange: sheets_v4.Schema$DimensionRange = {
        sheetId,
        dimension: "COLUMNS",
        startIndex: start_column,
        endIndex: end_column + 1,
      };

      const request: sheets_v4.Schema$Request =
        pixel_size !== undefined
          ? {
              updateDimensionProperties: {
                range: dimensionRange,
                properties: { pixelSize: pixel_size },
                fields: "pixelSize",
              },
            }
          : {
              autoResizeDimensions: {
                dimensions: dimensionRange,
              },
            };

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: { requests: [request] },
      });

      const colCount = end_column - start_column + 1;
      const sizeLabel = pixel_size !== undefined ? `${pixel_size}px` : "auto";
      return formatSuccess(
        `Resized ${colCount} column(s) (cols ${start_column}–${end_column}) to ${sizeLabel}`
      );
    })
  );

  // ─── F4.2 sheets_resize_rows ──────────────────────────────────────────────

  server.tool(
    "sheets_resize_rows",
    "Sets the pixel height of a row range using spreadsheets.batchUpdate with updateDimensionProperties; omit pixel_size to auto-fit row height to cell content using autoResizeDimensions. Use when the user asks to set a specific row height or auto-fit row heights after wrapping text. Use when adjusting header row height for visual emphasis in a table layout. Do not use when: resizing columns - use sheets_resize_columns instead; merging cells - use sheets_merge_cells instead; formatting cell content - use sheets_format_cells instead; freezing panes - use sheets_freeze_panes instead; applying borders - use sheets_set_borders instead; applying conditional formatting - use sheets_add_conditional_format instead. Returns: 'Resized {N} row(s) (rows {start+1}–{end+1}) to {px}px' or 'to auto'. Parameters: - start_row: 0-based inclusive start index, e.g. 0 = row 1 - end_row: 0-based inclusive end index, e.g. 0 = row 1 (single row) - pixel_size: height in pixels; omit to auto-fit.",
    {
      spreadsheet_id: z.string().describe("sheet ID from the URL (the token between /d/ and /edit) or the full URL"),
      sheet: z.string().describe("Sheet name or numeric sheet ID"),
      start_row: z.number().int().describe("0-based start row index"),
      end_row: z.number().int().describe("0-based end row index (inclusive)"),
      pixel_size: z
        .number()
        .int()
        .min(0)
        .optional()
        .describe("Row height in pixels. Omit to auto-resize."),
    },
    withErrorHandling(async ({ spreadsheet_id, sheet, start_row, end_row, pixel_size }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);
      const sheetId = await resolveSheetId(id, sheet);

      const dimensionRange: sheets_v4.Schema$DimensionRange = {
        sheetId,
        dimension: "ROWS",
        startIndex: start_row,
        endIndex: end_row + 1,
      };

      const request: sheets_v4.Schema$Request =
        pixel_size !== undefined
          ? {
              updateDimensionProperties: {
                range: dimensionRange,
                properties: { pixelSize: pixel_size },
                fields: "pixelSize",
              },
            }
          : {
              autoResizeDimensions: {
                dimensions: dimensionRange,
              },
            };

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: { requests: [request] },
      });

      const rowCount = end_row - start_row + 1;
      const sizeLabel = pixel_size !== undefined ? `${pixel_size}px` : "auto";
      return formatSuccess(
        `Resized ${rowCount} row(s) (rows ${start_row + 1}–${end_row + 1}) to ${sizeLabel}`
      );
    })
  );

  // ─── F4.2 sheets_freeze_panes ─────────────────────────────────────────────

  server.tool(
    "sheets_freeze_panes",
    "Freezes a number of rows and/or columns in a sheet using spreadsheets.batchUpdate with updateSheetProperties targeting gridProperties; pass 0 to unfreeze. Use when the user asks to keep a header row or label column visible while scrolling through a large dataset. Use when unfreezing previously frozen rows by setting frozen_row_count=0. Do not use when: sheets_write_table or sheets_build_sheet already include freeze - prefer those for full table builds; formatting cells - use sheets_format_cells instead; merging cells - use sheets_merge_cells instead; resizing columns - use sheets_resize_columns instead; resizing rows - use sheets_resize_rows instead; applying borders - use sheets_set_borders instead; applying conditional formatting - use sheets_add_conditional_format instead. Returns: 'Froze {N} row(s) and/or {N} column(s) in sheet \"{sheet}\"', or 'Nothing to freeze - provide frozen_row_count and/or frozen_column_count.' Parameters: - sheet: tab name or numeric sheet ID - frozen_row_count: number of rows to freeze from the top (0 unfreezes) - frozen_column_count: number of columns to freeze from the left (0 unfreezes).",
    {
      spreadsheet_id: z.string().describe("sheet ID from the URL (the token between /d/ and /edit) or the full URL"),
      sheet: z.string().describe("Sheet name or numeric sheet ID"),
      frozen_row_count: z
        .number()
        .int()
        .min(0)
        .optional()
        .describe("Number of rows to freeze (0 = unfreeze rows)"),
      frozen_column_count: z
        .number()
        .int()
        .min(0)
        .optional()
        .describe("Number of columns to freeze (0 = unfreeze columns)"),
    },
    withErrorHandling(async ({ spreadsheet_id, sheet, frozen_row_count, frozen_column_count }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);
      const sheetId = await resolveSheetId(id, sheet);

      const properties: sheets_v4.Schema$GridProperties = {};
      const fields: string[] = [];

      if (frozen_row_count !== undefined) {
        properties.frozenRowCount = frozen_row_count;
        fields.push("gridProperties.frozenRowCount");
      }
      if (frozen_column_count !== undefined) {
        properties.frozenColumnCount = frozen_column_count;
        fields.push("gridProperties.frozenColumnCount");
      }

      if (fields.length === 0) {
        return formatSuccess("Nothing to freeze - provide frozen_row_count and/or frozen_column_count.");
      }

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: {
          requests: [
            {
              updateSheetProperties: {
                properties: {
                  sheetId,
                  gridProperties: properties,
                },
                fields: fields.join(","),
              },
            },
          ],
        },
      });

      const parts: string[] = [];
      if (frozen_row_count !== undefined) parts.push(`${frozen_row_count} row(s)`);
      if (frozen_column_count !== undefined) parts.push(`${frozen_column_count} column(s)`);
      return formatSuccess(`Froze ${parts.join(" and ")} in sheet "${sheet}"`);
    })
  );

  // ─── F4.3 sheets_add_conditional_format ──────────────────────────────────

  server.tool(
    "sheets_add_conditional_format",
    "Adds a single conditional formatting rule to a cell range using spreadsheets.batchUpdate with addConditionalFormatRule; supports gradient color scales, boolean conditions (cell value comparisons), and custom formula conditions. Use when the user asks to highlight cells above a threshold, apply a red-green gradient scale to a column, or color cells matching a custom formula. Use when adding one targeted rule to a range without batching. Do not use when: adding multiple rules in one call - use sheets_add_conditional_formats instead; applying static formatting - use sheets_format_cells instead; applying borders - use sheets_set_borders instead; merging cells - use sheets_merge_cells instead; setting data validation rules - use sheets_set_validation instead; clearing an existing conditional rule - use sheets_format_cells to overwrite. Returns: 'Added conditional formatting rule to {range}'. Parameters: - range: A1 notation including sheet name, e.g. 'Sheet1!A2:A100' - rule_type: 'gradient' for color scales, 'boolean' for value comparisons, or 'custom_formula' for formula-driven rules.",
    {
      spreadsheet_id: z.string().describe("sheet ID from the URL (the token between /d/ and /edit) or the full URL"),
      range: z.string().describe("A1 notation range"),
      rule_type: z
        .enum(["condition", "color_scale", "formula"])
        .describe("Type of conditional formatting rule"),

      // ── Condition-based options ──
      condition_type: z
        .enum([
          "NUMBER_GREATER",
          "NUMBER_GREATER_THAN_EQ",
          "NUMBER_LESS",
          "NUMBER_LESS_THAN_EQ",
          "NUMBER_BETWEEN",
          "NUMBER_NOT_BETWEEN",
          "NUMBER_EQ",
          "NUMBER_NOT_EQ",
          "TEXT_CONTAINS",
          "TEXT_NOT_CONTAINS",
          "TEXT_STARTS_WITH",
          "TEXT_ENDS_WITH",
          "TEXT_EQ",
          "BLANK",
          "NOT_BLANK",
          "CUSTOM_FORMULA",
        ])
        .optional()
        .describe("Condition type (required when rule_type='condition')"),
      condition_values: z
        .array(z.string())
        .optional()
        .describe("Condition values (1 or 2 depending on type)"),

      // ── Format to apply (condition and formula) ──
      format_background_color: z.string().optional().describe("Background color to apply"),
      format_text_color: z.string().optional().describe("Text color to apply"),
      format_bold: z.boolean().optional().describe("Bold text"),
      format_italic: z.boolean().optional().describe("Italic text"),
      format_strikethrough: z.boolean().optional().describe("Strikethrough text"),

      // ── Color scale options ──
      min_color: z.string().optional().describe("Color for min value in color scale"),
      mid_color: z
        .string()
        .optional()
        .describe("Color for mid value in color scale (omit for 2-color scale)"),
      max_color: z.string().optional().describe("Color for max value in color scale"),
      min_type: z
        .enum(["MIN", "NUMBER", "PERCENT", "PERCENTILE"])
        .optional()
        .describe("Min point type (default: MIN)"),
      mid_type: z
        .enum(["NUMBER", "PERCENT", "PERCENTILE"])
        .optional()
        .describe("Mid point type (default: PERCENTILE at 50)"),
      max_type: z
        .enum(["MAX", "NUMBER", "PERCENT", "PERCENTILE"])
        .optional()
        .describe("Max point type (default: MAX)"),
      min_value: z.string().optional().describe("Min point value (for NUMBER/PERCENT/PERCENTILE types)"),
      mid_value: z.string().optional().describe("Mid point value"),
      max_value: z.string().optional().describe("Max point value"),

      // ── Custom formula ──
      formula: z
        .string()
        .optional()
        .describe("Formula for rule_type='formula', e.g. '=A1>100'. Must start with '='."),
    },
    withErrorHandling(
      async ({
        spreadsheet_id,
        range,
        rule_type,
        condition_type,
        condition_values,
        format_background_color,
        format_text_color,
        format_bold,
        format_italic,
        format_strikethrough,
        min_color,
        mid_color,
        max_color,
        min_type,
        mid_type,
        max_type,
        min_value,
        mid_value,
        max_value,
        formula,
      }) => {
        const sheets = await getSheetsClient();
        const id = extractFileId(spreadsheet_id);
        const sheetId = await resolveSheetIdFromRange(id, range);
        const gridRange = toGridRange(range, sheetId);

        let request: sheets_v4.Schema$Request;

        if (rule_type === "color_scale") {
          // Build color scale rule
          const minPoint: sheets_v4.Schema$InterpolationPoint = {
            colorStyle: { rgbColor: parseColor(min_color ?? "#ffffff") },
            type: min_type ?? "MIN",
          };
          if (min_value !== undefined) minPoint.value = min_value;

          const maxPoint: sheets_v4.Schema$InterpolationPoint = {
            colorStyle: { rgbColor: parseColor(max_color ?? "#ff0000") },
            type: max_type ?? "MAX",
          };
          if (max_value !== undefined) maxPoint.value = max_value;

          // Build proper color scale with interpolation points
          const gradientRule: sheets_v4.Schema$GradientRule = {
            minpoint: minPoint,
            maxpoint: maxPoint,
          };

          if (mid_color) {
            const midPoint: sheets_v4.Schema$InterpolationPoint = {
              colorStyle: { rgbColor: parseColor(mid_color) },
              type: mid_type ?? "PERCENTILE",
              value: mid_value ?? "50",
            };
            gradientRule.midpoint = midPoint;
          }

          request = {
            addConditionalFormatRule: {
              rule: {
                ranges: [gridRange],
                gradientRule,
              },
              index: 0,
            },
          };
        } else {
          // Condition or formula rule
          const booleanRule: sheets_v4.Schema$BooleanRule = {
            condition: {},
            format: {},
          };

          // Determine condition
          if (rule_type === "formula" && formula) {
            booleanRule.condition = {
              type: "CUSTOM_FORMULA",
              values: [{ userEnteredValue: formula }],
            };
          } else if (condition_type) {
            const condValues = (condition_values ?? []).map((v) => ({
              userEnteredValue: v,
            }));
            booleanRule.condition = {
              type: condition_type,
              values: condValues,
            };
          } else {
            throw new Error(
              "For rule_type='condition', provide condition_type. For rule_type='formula', provide formula."
            );
          }

          // Build format
          const cellFormat: sheets_v4.Schema$CellFormat = {};
          const textFormat: sheets_v4.Schema$TextFormat = {};

          if (format_background_color) {
            cellFormat.backgroundColorStyle = { rgbColor: parseColor(format_background_color) };
          }
          if (format_text_color) {
            textFormat.foregroundColorStyle = { rgbColor: parseColor(format_text_color) };
          }
          if (format_bold !== undefined) textFormat.bold = format_bold;
          if (format_italic !== undefined) textFormat.italic = format_italic;
          if (format_strikethrough !== undefined) textFormat.strikethrough = format_strikethrough;

          if (Object.keys(textFormat).length > 0) {
            cellFormat.textFormat = textFormat;
          }

          booleanRule.format = cellFormat;
          request = {
            addConditionalFormatRule: {
              rule: {
                ranges: [gridRange],
                booleanRule,
              },
              index: 0,
            },
          };
        }

        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: id,
          requestBody: { requests: [request] },
        });

        return formatSuccess(`Added conditional formatting rule to ${range}`);
      }
    )
  );

  // ─── sheets_add_conditional_formats (bulk) ────────────────────────────────

  server.tool(
    "sheets_add_conditional_formats",
    "Adds multiple conditional formatting rules in a single API call using spreadsheets.batchUpdate with multiple addConditionalFormatRule requests; each rule can target a different range and use a different rule type. Use when the user asks to apply several conditional formatting rules at once, such as highlighting different status values with different colors across a column. Use when batching rules is needed to avoid multiple round trips to the API. Do not use when: adding a single conditional formatting rule - use sheets_add_conditional_format instead; applying static formatting to multiple ranges - use sheets_apply_formats instead; formatting a single range - use sheets_format_cells instead; applying borders - use sheets_set_borders instead; setting data validation - use sheets_set_validation instead. Returns: 'Added {N} conditional formatting rule(s)'. Parameters: - rules: array where each entry has a range (A1 notation) and a rule definition with rule_type (gradient, boolean, or custom_formula) and type-specific fields.",
    {
      spreadsheet_id: z.string().describe("sheet ID from the URL (the token between /d/ and /edit) or the full URL"),
      rules: z
        .array(
          z.object({
            range: z.string().describe("A1 notation range"),
            rule_type: z.enum(["condition", "color_scale", "formula"]),
            condition_type: z
              .enum([
                "NUMBER_GREATER", "NUMBER_GREATER_THAN_EQ", "NUMBER_LESS",
                "NUMBER_LESS_THAN_EQ", "NUMBER_BETWEEN", "NUMBER_NOT_BETWEEN",
                "NUMBER_EQ", "NUMBER_NOT_EQ", "TEXT_CONTAINS", "TEXT_NOT_CONTAINS",
                "TEXT_STARTS_WITH", "TEXT_ENDS_WITH", "TEXT_EQ",
                "BLANK", "NOT_BLANK", "CUSTOM_FORMULA",
              ])
              .optional(),
            condition_values: z.array(z.string()).optional(),
            format_background_color: z.string().optional(),
            format_text_color: z.string().optional(),
            format_bold: z.boolean().optional(),
            format_italic: z.boolean().optional(),
            format_strikethrough: z.boolean().optional(),
            min_color: z.string().optional(),
            mid_color: z.string().optional(),
            max_color: z.string().optional(),
            min_type: z.enum(["MIN", "NUMBER", "PERCENT", "PERCENTILE"]).optional(),
            mid_type: z.enum(["NUMBER", "PERCENT", "PERCENTILE"]).optional(),
            max_type: z.enum(["MAX", "NUMBER", "PERCENT", "PERCENTILE"]).optional(),
            min_value: z.string().optional(),
            mid_value: z.string().optional(),
            max_value: z.string().optional(),
            formula: z.string().optional(),
          })
        )
        .min(1)
        .describe("Array of conditional format rules"),
    },
    withErrorHandling(async ({ spreadsheet_id, rules }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);
      const requests: sheets_v4.Schema$Request[] = [];
      const sheetIdCache = new Map<string, number>();

      for (let ruleIdx = 0; ruleIdx < rules.length; ruleIdx++) {
        const rule = rules[ruleIdx];
        const sheetId = await resolveSheetIdCached(id, rule.range, sheetIdCache);
        const gridRange = toGridRange(rule.range, sheetId);

        if (rule.rule_type === "color_scale") {
          const minPoint: sheets_v4.Schema$InterpolationPoint = {
            colorStyle: { rgbColor: parseColor(rule.min_color ?? "#ffffff") },
            type: rule.min_type ?? "MIN",
          };
          if (rule.min_value !== undefined) minPoint.value = rule.min_value;

          const maxPoint: sheets_v4.Schema$InterpolationPoint = {
            colorStyle: { rgbColor: parseColor(rule.max_color ?? "#ff0000") },
            type: rule.max_type ?? "MAX",
          };
          if (rule.max_value !== undefined) maxPoint.value = rule.max_value;

          const gradientRule: sheets_v4.Schema$GradientRule = {
            minpoint: minPoint,
            maxpoint: maxPoint,
          };

          if (rule.mid_color) {
            gradientRule.midpoint = {
              colorStyle: { rgbColor: parseColor(rule.mid_color) },
              type: rule.mid_type ?? "PERCENTILE",
              value: rule.mid_value ?? "50",
            };
          }

          requests.push({
            addConditionalFormatRule: {
              rule: { ranges: [gridRange], gradientRule },
              index: ruleIdx,
            },
          });
        } else {
          const booleanRule: sheets_v4.Schema$BooleanRule = {
            condition: {},
            format: {},
          };

          if (rule.rule_type === "formula" && rule.formula) {
            booleanRule.condition = {
              type: "CUSTOM_FORMULA",
              values: [{ userEnteredValue: rule.formula }],
            };
          } else if (rule.condition_type) {
            booleanRule.condition = {
              type: rule.condition_type,
              values: (rule.condition_values ?? []).map((v) => ({ userEnteredValue: v })),
            };
          } else {
            throw new Error(
              "For rule_type='condition', provide condition_type. For rule_type='formula', provide formula."
            );
          }

          const cellFormat: sheets_v4.Schema$CellFormat = {};
          const textFormat: sheets_v4.Schema$TextFormat = {};

          if (rule.format_background_color) {
            cellFormat.backgroundColorStyle = { rgbColor: parseColor(rule.format_background_color) };
          }
          if (rule.format_text_color) {
            textFormat.foregroundColorStyle = { rgbColor: parseColor(rule.format_text_color) };
          }
          if (rule.format_bold !== undefined) textFormat.bold = rule.format_bold;
          if (rule.format_italic !== undefined) textFormat.italic = rule.format_italic;
          if (rule.format_strikethrough !== undefined) textFormat.strikethrough = rule.format_strikethrough;

          if (Object.keys(textFormat).length > 0) cellFormat.textFormat = textFormat;
          booleanRule.format = cellFormat;

          requests.push({
            addConditionalFormatRule: {
              rule: { ranges: [gridRange], booleanRule },
              index: ruleIdx,
            },
          });
        }
      }

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: { requests },
      });

      return formatSuccess(`Added ${rules.length} conditional formatting rule(s)`);
    })
  );

  // ─── sheets_apply_formats (bulk multi-range formatting) ───────────────────

  server.tool(
    "sheets_apply_formats",
    "Applies formatting - including font, color, alignment, number format, and borders - to multiple cell ranges in a single API call using spreadsheets.batchUpdate with multiple repeatCell and updateBorders requests; each range in the operations array gets its own independent formatting specification. Use when the user asks to format several non-contiguous ranges at once, such as styling a header row, a totals row, and a data body in one operation. Use when combining border and cell formatting for the same range to avoid two separate API calls. Do not use when: formatting a single range - use sheets_format_cells or sheets_set_borders instead; adding conditional formatting - use sheets_add_conditional_format or sheets_add_conditional_formats instead; merging cells - use sheets_merge_cells instead; resizing columns - use sheets_resize_columns instead; freezing panes - use sheets_freeze_panes instead. Returns: 'Applied formatting: {N} range(s), {M} operation(s)', or 'No formatting options provided - nothing to update.' Parameters: - operations: array where each entry has a range (A1 notation) and optional formatting fields (bold, italic, font_size, text_color, background_color, horizontal_alignment, number_format_pattern, borders).",
    {
      spreadsheet_id: z.string().describe("sheet ID from the URL (the token between /d/ and /edit) or the full URL"),
      operations: z
        .array(
          z.object({
            range: z.string().describe("A1 notation range"),
            bold: z.boolean().optional(),
            italic: z.boolean().optional(),
            strikethrough: z.boolean().optional(),
            font_family: z.string().optional(),
            font_size: z.number().int().min(1).optional(),
            text_color: z.string().optional(),
            background_color: z.string().optional(),
            horizontal_alignment: z.enum(["LEFT", "CENTER", "RIGHT"]).optional(),
            vertical_alignment: z.enum(["TOP", "MIDDLE", "BOTTOM"]).optional(),
            number_format_pattern: z.string().optional(),
            number_format_type: z
              .enum(["TEXT", "NUMBER", "PERCENT", "CURRENCY", "DATE", "TIME", "DATE_TIME", "SCIENTIFIC"])
              .optional(),
            wrap_strategy: z.enum(["OVERFLOW_CELL", "WRAP", "CLIP"]).optional(),
            borders: z
              .object({
                top: z.boolean().optional(),
                bottom: z.boolean().optional(),
                left: z.boolean().optional(),
                right: z.boolean().optional(),
                inner_horizontal: z.boolean().optional(),
                inner_vertical: z.boolean().optional(),
                style: BorderStyleEnum.optional(),
                color: z.string().optional(),
              })
              .optional(),
          })
        )
        .min(1)
        .describe("Array of format operations, each targeting a range"),
    },
    withErrorHandling(async ({ spreadsheet_id, operations }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);
      const requests: sheets_v4.Schema$Request[] = [];
      const sheetIdCache = new Map<string, number>();

      for (const op of operations) {
        const sheetId = await resolveSheetIdCached(id, op.range, sheetIdCache);
        const gridRange = toGridRange(op.range, sheetId);

        // Build cell format
        const cellFormat: sheets_v4.Schema$CellFormat = {};
        const updateFields: string[] = [];

        const textFormat: sheets_v4.Schema$TextFormat = {};
        if (op.bold !== undefined) { textFormat.bold = op.bold; updateFields.push("userEnteredFormat.textFormat.bold"); }
        if (op.italic !== undefined) { textFormat.italic = op.italic; updateFields.push("userEnteredFormat.textFormat.italic"); }
        if (op.strikethrough !== undefined) { textFormat.strikethrough = op.strikethrough; updateFields.push("userEnteredFormat.textFormat.strikethrough"); }
        if (op.font_family !== undefined) { textFormat.fontFamily = op.font_family; updateFields.push("userEnteredFormat.textFormat.fontFamily"); }
        if (op.font_size !== undefined) { textFormat.fontSize = op.font_size; updateFields.push("userEnteredFormat.textFormat.fontSize"); }
        if (op.text_color !== undefined) {
          textFormat.foregroundColorStyle = { rgbColor: parseColor(op.text_color) };
          updateFields.push("userEnteredFormat.textFormat.foregroundColorStyle");
        }
        if (Object.keys(textFormat).length > 0) cellFormat.textFormat = textFormat;

        if (op.background_color !== undefined) {
          cellFormat.backgroundColorStyle = { rgbColor: parseColor(op.background_color) };
          updateFields.push("userEnteredFormat.backgroundColorStyle");
        }
        if (op.horizontal_alignment !== undefined) {
          cellFormat.horizontalAlignment = op.horizontal_alignment;
          updateFields.push("userEnteredFormat.horizontalAlignment");
        }
        if (op.vertical_alignment !== undefined) {
          cellFormat.verticalAlignment = op.vertical_alignment;
          updateFields.push("userEnteredFormat.verticalAlignment");
        }
        if (op.number_format_pattern !== undefined || op.number_format_type !== undefined) {
          cellFormat.numberFormat = {
            type: op.number_format_type ?? "NUMBER",
            pattern: op.number_format_pattern ?? "",
          };
          updateFields.push("userEnteredFormat.numberFormat");
        }
        if (op.wrap_strategy !== undefined) {
          cellFormat.wrapStrategy = op.wrap_strategy;
          updateFields.push("userEnteredFormat.wrapStrategy");
        }

        if (updateFields.length > 0) {
          requests.push({
            repeatCell: {
              range: gridRange,
              cell: { userEnteredFormat: cellFormat },
              fields: updateFields.join(","),
            },
          });
        }

        // Borders
        if (op.borders) {
          const borderSpec = {
            style: op.borders.style ?? "SOLID",
            color: op.borders.color,
          };
          const border = buildBorder(borderSpec);
          const updateBorders: sheets_v4.Schema$UpdateBordersRequest = { range: gridRange };
          if (op.borders.top && border) updateBorders.top = border;
          if (op.borders.bottom && border) updateBorders.bottom = border;
          if (op.borders.left && border) updateBorders.left = border;
          if (op.borders.right && border) updateBorders.right = border;
          if (op.borders.inner_horizontal && border) updateBorders.innerHorizontal = border;
          if (op.borders.inner_vertical && border) updateBorders.innerVertical = border;
          requests.push({ updateBorders });
        }
      }

      if (requests.length === 0) {
        return formatSuccess("No formatting options provided - nothing to update.");
      }

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: { requests },
      });

      return formatSuccess(
        `Applied formatting: ${operations.length} range(s), ${requests.length} operation(s)`
      );
    })
  );
}
