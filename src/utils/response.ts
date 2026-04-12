/**
 * Standard MCP response formatters.
 *
 * All tool handlers must return objects matching these shapes so that
 * MCP clients can process results uniformly.
 */

export interface McpTextContent {
  type: "text";
  text: string;
}

// The index signature ([x: string]: unknown) is required to satisfy the MCP SDK's
// CallToolResult type, which uses a Zod `$loose` schema (z.core.$loose).
export interface McpSuccessResponse {
  content: McpTextContent[];
  [x: string]: unknown;
}

export interface McpErrorResponse {
  content: McpTextContent[];
  isError: true;
  [x: string]: unknown;
}

/**
 * Wraps a plain text string in the MCP success response envelope.
 */
export function formatSuccess(text: string): McpSuccessResponse {
  return {
    content: [{ type: "text", text }],
  };
}

/**
 * Wraps a plain text message in the MCP error response envelope.
 * The `isError: true` flag signals to MCP clients that the call failed.
 */
export function formatError(message: string): McpErrorResponse {
  return {
    content: [{ type: "text", text: message }],
    isError: true,
  };
}
