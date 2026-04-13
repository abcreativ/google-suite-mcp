/**
 * Data validation tools: set and clear cell validation rules.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getSheetsClient, extractFileId } from "../client/google-client.js";
import { formatSuccess } from "../utils/response.js";
import { withErrorHandling } from "../utils/error-handler.js";
import { toGridRange } from "../utils/range.js";
import { resolveSheetIdFromRange } from "../utils/sheet-resolver.js";

// ─── Condition type enums ─────────────────────────────────────────────────────

const NumberCondition = z.enum([
  "BETWEEN",
  "NOT_BETWEEN",
  "GREATER_THAN",
  "LESS_THAN",
  "GREATER_THAN_OR_EQUAL",
  "LESS_THAN_OR_EQUAL",
  "EQ",
  "NOT_EQ",
]);

const DateCondition = z.enum([
  "BETWEEN",
  "NOT_BETWEEN",
  "BEFORE",
  "AFTER",
  "ON",
  "NOT_ON",
  "IS_VALID",
]);

const TextCondition = z.enum([
  "TEXT_CONTAINS",
  "TEXT_NOT_CONTAINS",
  "TEXT_STARTS_WITH",
  "TEXT_ENDS_WITH",
  "TEXT_EQ",
  "IS_VALID_EMAIL",
  "IS_VALID_URL",
]);

export function registerValidationTools(server: McpServer): void {
  // ─── sheets_set_validation ───────────────────────────────────────────────────

  server.tool(
    "sheets_set_validation",
    "Applies a data validation rule to a cell range using spreadsheets.batchUpdate with setDataValidation; supported types are dropdown, number, date, text, checkbox, and custom_formula. Use when the user asks to restrict a column to a predefined list of values, enforce a number range, or add a checkbox toggle to cells. Use when driving a dropdown from a range on another sheet by supplying dropdown_range instead of dropdown_values. Do not use when: removing an existing validation rule - use sheets_clear_validation instead; applying conditional formatting based on cell value - use sheets_add_conditional_format instead; reading current validation settings - use sheets_get_cell_info instead. Returns: 'Set {type} validation on {range}'. Parameters: - type: one of dropdown, number, date, text, checkbox, custom_formula - dropdown_values: array of allowed strings for type=dropdown, e.g. ['Yes','No','Maybe'] - dropdown_range: A1 range to pull dropdown options from, e.g. 'Sheet2!A1:A10'; takes precedence over dropdown_values - condition: condition string for number/date/text types, e.g. BETWEEN, GREATER_THAN, TEXT_CONTAINS - values: 1-2 condition threshold strings, e.g. ['10','100'] for BETWEEN - strict: true rejects invalid input; false shows warning only (default true).",
    {
      spreadsheet_id: z.string().describe("sheet ID from the URL (the token between /d/ and /edit) or the full URL"),
      range: z.string().describe("A1 notation range to apply validation to, e.g. 'Sheet1!A1:A100'"),
      sheet_id: z.number().int().optional().describe("Numeric sheet ID (default: 0)"),
      type: z
        .enum(["dropdown", "number", "date", "text", "checkbox", "custom_formula"])
        .describe("Validation type"),
      // Dropdown
      dropdown_values: z
        .array(z.string())
        .optional()
        .describe("For type=dropdown: list of allowed values"),
      dropdown_range: z
        .string()
        .optional()
        .describe("For type=dropdown: range reference to pull values from (e.g. 'Sheet2!A1:A10')"),
      // Number / Date / Text conditions
      condition: z
        .string()
        .optional()
        .describe("Condition type - e.g. BETWEEN, GREATER_THAN, TEXT_CONTAINS"),
      values: z
        .array(z.string())
        .optional()
        .describe("Condition values (1 or 2 depending on condition type)"),
      // Checkbox
      checked_value: z.string().optional().describe("For type=checkbox: value when checked (default: TRUE)"),
      unchecked_value: z.string().optional().describe("For type=checkbox: value when unchecked (default: FALSE)"),
      // Custom formula
      formula: z.string().optional().describe("For type=custom_formula: formula string starting with ="),
      // Behaviour
      strict: z
        .boolean()
        .optional()
        .describe("If true, reject invalid input. If false, show warning only (default: true)"),
      show_input_message: z.boolean().optional().describe("Show a help message on focus (default: false)"),
      input_message_title: z.string().optional(),
      input_message_body: z.string().optional(),
    },
    withErrorHandling(async ({
      spreadsheet_id,
      range,
      sheet_id,
      type,
      dropdown_values,
      dropdown_range,
      condition,
      values,
      checked_value,
      unchecked_value,
      formula,
      strict,
      show_input_message,
      input_message_title,
      input_message_body,
    }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);
      const resolvedSheetId = sheet_id ?? await resolveSheetIdFromRange(id, range);
      const gridRange = toGridRange(range, resolvedSheetId);

      // Build the BooleanCondition object
      let conditionObj: {
        type: string;
        values?: Array<{ userEnteredValue?: string; relativeDate?: string }>;
      };

      switch (type) {
        case "dropdown": {
          if (dropdown_range) {
            conditionObj = {
              type: "ONE_OF_RANGE",
              values: [{ userEnteredValue: `=${dropdown_range}` }],
            };
          } else {
            conditionObj = {
              type: "ONE_OF_LIST",
              values: (dropdown_values ?? []).map((v) => ({ userEnteredValue: v })),
            };
          }
          break;
        }
        case "number": {
          conditionObj = {
            type: condition ?? "NUMBER_GREATER",
            values: (values ?? []).map((v) => ({ userEnteredValue: v })),
          };
          break;
        }
        case "date": {
          conditionObj = {
            type: condition ?? "DATE_IS_VALID",
            values: (values ?? []).map((v) => ({ userEnteredValue: v })),
          };
          break;
        }
        case "text": {
          conditionObj = {
            type: condition ?? "TEXT_CONTAINS",
            values: (values ?? []).map((v) => ({ userEnteredValue: v })),
          };
          break;
        }
        case "checkbox": {
          if (checked_value || unchecked_value) {
            conditionObj = {
              type: "BOOLEAN",
              values: [
                { userEnteredValue: checked_value ?? "TRUE" },
                { userEnteredValue: unchecked_value ?? "FALSE" },
              ],
            };
          } else {
            conditionObj = { type: "BOOLEAN" };
          }
          break;
        }
        case "custom_formula": {
          conditionObj = {
            type: "CUSTOM_FORMULA",
            values: [{ userEnteredValue: formula ?? "" }],
          };
          break;
        }
      }

      const rule: Record<string, unknown> = {
        condition: conditionObj,
        strict: strict ?? true,
        showCustomUi: type === "dropdown",
      };

      if (show_input_message) {
        rule.inputMessage = input_message_title
          ? `${input_message_title}: ${input_message_body ?? ""}`
          : (input_message_body ?? "");
      }

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: {
          requests: [
            {
              setDataValidation: {
                range: gridRange,
                rule,
              },
            },
          ],
        },
      });

      return formatSuccess(`Set ${type} validation on ${range}`);
    })
  );

  // ─── sheets_clear_validation ─────────────────────────────────────────────────

  server.tool(
    "sheets_clear_validation",
    "Removes all data validation rules from a cell range using spreadsheets.batchUpdate with setDataValidation and an empty rule object; cell content is preserved, only the constraint is removed. Use when the user asks to unlock a restricted column so any value can be entered without warnings or rejections. Use when replacing one validation type with another by clearing first and then calling sheets_set_validation with new parameters. Do not use when: setting a new validation rule - use sheets_set_validation instead; applying conditional formatting - use sheets_add_conditional_format instead; clearing cell content - use sheets_write_range with empty values instead. Returns: 'Cleared validation on {range}'. Parameters: - range: A1 notation including sheet name, e.g. 'Sheet1!A1:A100'.",
    {
      spreadsheet_id: z.string().describe("sheet ID from the URL (the token between /d/ and /edit) or the full URL"),
      range: z.string().describe("A1 notation range, e.g. 'Sheet1!A1:A100'"),
      sheet_id: z.number().int().optional().describe("Numeric sheet ID (default: 0)"),
    },
    withErrorHandling(async ({ spreadsheet_id, range, sheet_id }) => {
      const sheets = await getSheetsClient();
      const id = extractFileId(spreadsheet_id);
      const resolvedSheetId = sheet_id ?? await resolveSheetIdFromRange(id, range);
      const gridRange = toGridRange(range, resolvedSheetId);

      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: {
          requests: [
            {
              setDataValidation: {
                range: gridRange,
                // omitting rule clears validation
              },
            },
          ],
        },
      });

      return formatSuccess(`Cleared validation on ${range}`);
    })
  );
}
