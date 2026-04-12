/**
 * Error handling utilities for MCP tool handlers.
 *
 * Provides a `withErrorHandling` wrapper that catches Google API errors and
 * converts them into actionable MCP error responses.
 */

import { formatError } from "./response.js";
import type { McpSuccessResponse, McpErrorResponse } from "./response.js";

/** Shape that all tool handlers must return. */
export type ToolResult = McpSuccessResponse | McpErrorResponse;

/** A Google API error response body (partial - only the fields we inspect). */
interface GoogleApiError {
  code?: number;
  message?: string;
  status?: string;
  errors?: Array<{ message: string; reason?: string }>;
}

interface ErrorWithResponse {
  response?: {
    data?: {
      error?: GoogleApiError;
    };
    status?: number;
  };
  message?: string;
  code?: string | number;
}

/**
 * Maps a raw error into a human-readable, actionable message.
 */
function describeError(err: unknown): string {
  const error = err as ErrorWithResponse;

  // Google API errors surface inside `error.response.data.error`
  const apiError = error?.response?.data?.error;

  if (apiError) {
    const code = apiError.code ?? error?.response?.status;
    const msg = apiError.message ?? "Unknown API error";

    // Include structured error reasons when available (helps LLM self-correct)
    const reasons = apiError.errors?.map((e) => e.reason ?? e.message).filter(Boolean);
    const reasonSuffix = reasons && reasons.length > 0
      ? ` Reasons: ${reasons.join("; ")}.`
      : "";

    switch (code) {
      case 400:
        return `Bad request: ${msg}.${reasonSuffix} Check range, ID, and parameters.`;
      case 401:
        return `Auth failed: ${msg}. Run auth_refresh or re-authenticate.`;
      case 403:
        return `Permission denied: ${msg}.${reasonSuffix}`;
      case 404:
        // Add context for Apps Script "not found" errors
        if (msg.includes("entity was not found") || msg.includes("Requested entity")) {
          return `Not found: ${msg}. For script_run: ensure the script's GCP project matches your OAuth credentials' GCP project (Apps Script editor > Project Settings).`;
        }
        return `Not found: ${msg}. Check file ID or range.`;
      case 429:
        return `Rate limited: ${msg}. Retry after a brief wait.`;
      case 500:
      case 503:
        return `Google API error (${code}): ${msg}. Temporary - retry.`;
      default:
        return `Google API error (${code ?? "?"}): ${msg}${reasonSuffix}`;
    }
  }

  // Node.js / network-level errors
  if (error?.code === "ENOTFOUND" || error?.code === "ECONNREFUSED") {
    return `Network error: Could not reach the Google API. Check your internet connection.`;
  }

  if (error?.message) {
    return error.message;
  }

  return String(err);
}

/**
 * Wraps a tool handler function with error handling.
 *
 * Usage:
 * ```ts
 * server.tool("my_tool", "...", schema, withErrorHandling(async (args) => {
 *   // your implementation
 * }));
 * ```
 */
export function withErrorHandling<TArgs>(
  handler: (args: TArgs) => Promise<ToolResult>
): (args: TArgs) => Promise<ToolResult> {
  return async (args: TArgs): Promise<ToolResult> => {
    try {
      return await handler(args);
    } catch (err) {
      const message = describeError(err);
      console.error("[google-suite-mcp] Tool error:", message);
      return formatError(message);
    }
  };
}
