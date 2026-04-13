/**
 * Docs tools - document management, content writing, formatting, and rich content.
 */

import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { docs_v1 } from "googleapis";
import { getDocsClient, getDriveClient, extractFileId } from "../client/google-client.js";
import { formatSuccess, formatError } from "../utils/response.js";
import { withErrorHandling } from "../utils/error-handler.js";
import { parseColor } from "../utils/color.js";

// ─── Helpers ──────────────────────────────────────────────────────────────────

const HEADING_STYLE: Record<number, string> = {
  1: "HEADING_1",
  2: "HEADING_2",
  3: "HEADING_3",
  4: "HEADING_4",
  5: "HEADING_5",
  6: "HEADING_6",
};

const EXPORT_MIME: Record<string, string> = {
  pdf: "application/pdf",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  txt: "text/plain",
  html: "text/html",
  epub: "application/epub+zip",
};

/**
 * Extract plain text from a Google Docs document, preserving structural markers
 * for headings, paragraphs, and list items.
 */
function extractDocText(document: docs_v1.Schema$Document): string {
  const body = document.body?.content ?? [];
  const lines: string[] = [];

  for (const element of body) {
    // ── Tables ──
    if (element.table) {
      const table = element.table;
      const rows = table.tableRows ?? [];
      for (const row of rows) {
        const cells = (row.tableCells ?? []).map((cell) => {
          // Each cell contains paragraphs - extract their text
          return (cell.content ?? [])
            .filter((el) => el.paragraph)
            .map((el) =>
              (el.paragraph!.elements ?? [])
                .map((run) => run.textRun?.content ?? "")
                .join("")
                .replace(/\n$/, "")
            )
            .join(" ");
        });
        lines.push(cells.join("\t"));
      }
      continue;
    }

    // ── Paragraphs ──
    if (!element.paragraph) continue;

    const para = element.paragraph;
    const style = para.paragraphStyle?.namedStyleType ?? "NORMAL_TEXT";
    const listId = para.bullet?.listId;
    const nestingLevel = para.bullet?.nestingLevel ?? 0;

    // Collect text runs
    const text = (para.elements ?? [])
      .map((el) => el.textRun?.content ?? "")
      .join("")
      .replace(/\n$/, ""); // trim trailing newline added by Docs

    if (!text.trim()) continue;

    // Apply structural markers
    if (style.startsWith("HEADING_")) {
      const level = parseInt(style.replace("HEADING_", ""), 10);
      const prefix = "#".repeat(level);
      lines.push(`${prefix} ${text}`);
    } else if (listId !== undefined) {
      const indent = "  ".repeat(nestingLevel);
      lines.push(`${indent}- ${text}`);
    } else {
      lines.push(text);
    }
  }

  return lines.join("\n");
}

/**
 * Get the end index of a document (one before the final segment end).
 */
function getDocEndIndex(document: docs_v1.Schema$Document): number {
  const body = document.body?.content ?? [];
  if (body.length === 0) return 1;
  const last = body[body.length - 1];
  // Ensure we never return 0 - the Docs API requires index >= 1 for insertions.
  return Math.max(1, (last.endIndex ?? 1) - 1);
}

// ─── Register all Docs tools ──────────────────────────────────────────────────

export function registerDocsTools(server: McpServer): void {
  // ── docs_create ──────────────────────────────────────────────────────────────
  server.tool(
    "docs_create",
    "Creates a new Google Docs document using documents.create, optionally inserting initial plain text at index 1; returns the document ID and edit URL. Use when the user asks to create a new document or draft. Use when you need a document ID before appending content with docs_write_text. Do not use when: reading an existing document - use docs_get_text instead; replacing text in an existing document - use docs_replace_text instead; inserting text at a specific position - use docs_write_text instead; formatting text - use docs_format_text instead; exporting a document - use docs_export instead; inserting a table - use docs_insert_table instead; inserting an image - use docs_insert_image instead. Returns: 'Created: {title}\\nid: {docId}\\nurl: https://docs.google.com/document/d/{docId}/edit'. Parameters: - title: document name (default 'Untitled Document') - content: optional initial text inserted at the document start (index 1).",
    {
      title: z.string().default("Untitled Document"),
      content: z.string().optional().describe("Initial plain text content"),
    },
    withErrorHandling(async (args) => {
      const docs = await getDocsClient();

      const res = await docs.documents.create({
        requestBody: { title: args.title },
      });

      const docId = res.data.documentId;
      if (!docId) return formatError("Failed to create document: no ID returned.");

      // Insert initial content if provided
      if (args.content) {
        await docs.documents.batchUpdate({
          documentId: docId,
          requestBody: {
            requests: [
              {
                insertText: {
                  location: { index: 1 },
                  text: args.content,
                },
              },
            ],
          },
        });
      }

      return formatSuccess(
        `Created: ${args.title}\nid: ${docId}\nurl: https://docs.google.com/document/d/${docId}/edit`
      );
    })
  );

  // ── docs_get_text ────────────────────────────────────────────────────────────
  server.tool(
    "docs_get_text",
    "Reads all text content from a Google Doc using documents.get and extracts it as plain text; headings are prefixed with # marks, list items with dashes, and table cells are tab-separated. Use when the user asks to read or summarize a document. Use when you need to find a character index before calling docs_write_text or docs_format_text with a numeric position. Do not use when: creating a document - use docs_create instead; replacing specific text - use docs_replace_text instead; inserting text - use docs_write_text instead; formatting text - use docs_format_text instead; exporting in a specific format - use docs_export instead. Returns: the extracted document text as a string, or '(empty document)'; appends a truncation notice if max_chars is reached. Parameters: - docId: document ID or full URL - max_chars: character limit for the response (default 50000).",
    {
      docId: z.string().describe("Document ID or URL"),
      max_chars: z.number().int().min(1).optional().describe("Truncate output to N chars (default: 50000)"),
    },
    withErrorHandling(async (args) => {
      const docs = await getDocsClient();
      const id = extractFileId(args.docId);
      const limit = args.max_chars ?? 50_000;

      const res = await docs.documents.get({ documentId: id });
      const text = extractDocText(res.data);

      if (!text.trim()) return formatSuccess("(empty document)");
      if (text.length > limit) {
        return formatSuccess(text.slice(0, limit) + `\n\n(truncated at ${limit} chars - set max_chars for more)`);
      }
      return formatSuccess(text);
    })
  );

  // ── docs_export ──────────────────────────────────────────────────────────────
  server.tool(
    "docs_export",
    "Exports a Google Doc to a specified format using drive.files.export; saves the file to a local path when localPath is provided, or returns the content directly for txt/html or a size summary for binary formats. Use when the user asks to download or export a document as PDF, Word, plain text, HTML, or EPUB. Use when exporting to a local file path for further processing. Do not use when: reading the document text in-session - use docs_get_text instead; creating a document - use docs_create instead; replacing text - use docs_replace_text instead; inserting content - use docs_write_text or docs_insert_table or docs_insert_image instead. Returns: 'Exported to: {localPath}' when saved; for txt/html returns the raw content; for binary formats without localPath returns 'Export ready ({format}, {N} bytes). Provide localPath to save to disk.' Parameters: - format: one of pdf, docx, txt, html, epub - localPath: absolute local file path to save to (optional; must be absolute if provided).",
    {
      docId: z.string().describe("Document ID or URL"),
      format: z.enum(["pdf", "docx", "txt", "html", "epub"]),
      localPath: z.string().optional().describe("Absolute local path to save file"),
    },
    withErrorHandling(async (args) => {
      const drive = await getDriveClient();
      const id = extractFileId(args.docId);
      const mimeType = EXPORT_MIME[args.format];

      const response = await drive.files.export(
        { fileId: id, mimeType },
        { responseType: "arraybuffer" }
      );

      const buffer = Buffer.from(response.data as ArrayBuffer);

      if (args.localPath) {
        const fs = await import("fs");
        const path = await import("path");
        if (!path.default.isAbsolute(args.localPath)) {
          return formatError("localPath must be an absolute path.");
        }
        const dir = path.default.dirname(args.localPath);
        if (!fs.default.existsSync(dir)) {
          fs.default.mkdirSync(dir, { recursive: true });
        }
        fs.default.writeFileSync(args.localPath, buffer);
        return formatSuccess(`Exported to: ${args.localPath}`);
      }

      // Return as base64 for small text exports, or size info for binary
      if (args.format === "txt" || args.format === "html") {
        return formatSuccess(buffer.toString("utf-8"));
      }

      return formatSuccess(
        `Export ready (${args.format}, ${buffer.length} bytes). Provide localPath to save to disk.`
      );
    })
  );

  // ── docs_write_text ──────────────────────────────────────────────────────────
  server.tool(
    "docs_write_text",
    "Inserts text at the beginning, end, or a 1-based character index in a Google Doc using documents.batchUpdate with insertText; the Docs API uses 1-based indices where index 1 is the document start. Use when the user asks to append content to a document, prepend a header, or insert text at a known position. Use when building up document content section by section after creating it with docs_create. Do not use when: replacing existing text by pattern - use docs_replace_text instead; reading document content - use docs_get_text instead; formatting a range of text - use docs_format_text instead; inserting a table - use docs_insert_table instead; inserting an image - use docs_insert_image instead; creating a document - use docs_create instead. Returns: 'Inserted {N} characters at index {index}.' Parameters: - position: 'beginning' (index 1), 'end' (auto-detected), or a 1-based integer index; use docs_get_text to find the current document length before targeting a specific index.",
    {
      docId: z.string().describe("Document ID or URL"),
      text: z.string().describe("Text to insert"),
      position: z
        .union([z.enum(["beginning", "end"]), z.number().int().min(1)])
        .default("end")
        .describe("'beginning', 'end', or a numeric index"),
    },
    withErrorHandling(async (args) => {
      const docs = await getDocsClient();
      const id = extractFileId(args.docId);

      let index: number;

      if (args.position === "beginning") {
        index = 1;
      } else if (args.position === "end") {
        const docRes = await docs.documents.get({ documentId: id, fields: "body" });
        index = getDocEndIndex(docRes.data);
      } else {
        index = args.position as number;
      }

      await docs.documents.batchUpdate({
        documentId: id,
        requestBody: {
          requests: [
            {
              insertText: {
                location: { index },
                text: args.text,
              },
            },
          ],
        },
      });

      return formatSuccess(`Inserted ${args.text.length} characters at index ${index}.`);
    })
  );

  // ── docs_replace_text ────────────────────────────────────────────────────────
  server.tool(
    "docs_replace_text",
    "Replaces all occurrences of a text pattern throughout a Google Doc using documents.batchUpdate with replaceAllText; the replacement is applied document-wide. Use when the user asks to update a recurring label, name, or placeholder in a document, such as replacing '{{CLIENT_NAME}}' with an actual name. Use when standardizing terminology across an entire document in one operation. Do not use when: inserting new content at a position - use docs_write_text instead; formatting existing text - use docs_format_text instead; reading document content - use docs_get_text instead; creating a document - use docs_create instead; inserting a table - use docs_insert_table instead; inserting an image - use docs_insert_image instead. Returns: 'Replaced {N} occurrence(s) of \"{find}\".' Parameters: - find: literal text string to search for - replace: replacement text - matchCase: true for case-sensitive matching (default false).",
    {
      docId: z.string().describe("Document ID or URL"),
      find: z.string().describe("Text to search for"),
      replace: z.string().describe("Replacement text"),
      matchCase: z.boolean().default(false),
    },
    withErrorHandling(async (args) => {
      const docs = await getDocsClient();
      const id = extractFileId(args.docId);

      const res = await docs.documents.batchUpdate({
        documentId: id,
        requestBody: {
          requests: [
            {
              replaceAllText: {
                containsText: {
                  text: args.find,
                  matchCase: args.matchCase,
                },
                replaceText: args.replace,
              },
            },
          ],
        },
      });

      const count =
        res.data.replies?.[0]?.replaceAllText?.occurrencesChanged ?? 0;
      return formatSuccess(`Replaced ${count} occurrence(s) of "${args.find}".`);
    })
  );

  // ── docs_format_text ─────────────────────────────────────────────────────────
  server.tool(
    "docs_format_text",
    "Applies inline text formatting (bold, italic, underline, strikethrough, font size, color, hyperlink) or paragraph heading style to a character range in a Google Doc using documents.batchUpdate with updateTextStyle and updateParagraphStyle; the range is specified by 1-based startIndex and endIndex, or resolved automatically from searchText. Use when the user asks to bold a heading, set a font size, color a phrase, or apply a heading style to a paragraph. Use when formatting a known text phrase by passing it as searchText to avoid calculating indices manually. Do not use when: inserting new text - use docs_write_text instead; replacing text - use docs_replace_text instead; reading document content - use docs_get_text instead; inserting a table - use docs_insert_table instead; inserting an image - use docs_insert_image instead. Returns: 'Applied formatting to range [{startIndex}, {endIndex}).' Parameters: - startIndex/endIndex: 1-based character indices (endIndex is exclusive); use docs_get_text to read content and count character positions - searchText: alternative to index pair; finds the first occurrence automatically.",
    {
      docId: z.string().describe("Document ID or URL"),
      startIndex: z.number().int().min(1).optional().describe("Range start index (inclusive)"),
      endIndex: z.number().int().min(1).optional().describe("Range end index (exclusive)"),
      searchText: z.string().optional().describe("Find text to determine range (uses first occurrence)"),
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      underline: z.boolean().optional(),
      strikethrough: z.boolean().optional(),
      fontSize: z.number().positive().optional().describe("Font size in pt"),
      foregroundColor: z.string().optional().describe("Hex color, e.g. #FF0000"),
      headingLevel: z.number().int().min(1).max(6).optional().describe("Apply heading style (1-6)"),
      linkUrl: z.string().url().optional().describe("Apply hyperlink"),
    },
    withErrorHandling(async (args) => {
      const docs = await getDocsClient();
      const id = extractFileId(args.docId);

      let startIndex = args.startIndex;
      let endIndex = args.endIndex;

      // Resolve range via text search if indices not provided
      if (startIndex === undefined || endIndex === undefined) {
        if (!args.searchText) {
          return formatError(
            "Provide either startIndex+endIndex or searchText to identify the range."
          );
        }

        const docRes = await docs.documents.get({ documentId: id });
        const body = docRes.data.body?.content ?? [];
        let found = false;

        // Search across concatenated paragraph text (not individual textRuns)
        // to find text that spans formatting boundaries.
        for (const element of body) {
          if (!element.paragraph) continue;
          const elements = element.paragraph.elements ?? [];
          // Build the full paragraph text and track each character's document index
          let fullText = "";
          const indexMap: number[] = []; // fullText offset → document index
          for (const el of elements) {
            if (!el.textRun) continue;
            const content = el.textRun.content ?? "";
            const elStart = el.startIndex ?? 0;
            for (let i = 0; i < content.length; i++) {
              indexMap.push(elStart + i);
              fullText += content[i];
            }
          }
          const pos = fullText.indexOf(args.searchText);
          if (pos !== -1) {
            startIndex = indexMap[pos];
            endIndex = indexMap[pos + args.searchText.length - 1] + 1;
            found = true;
            break;
          }
        }

        if (!found) {
          return formatError(`Text "${args.searchText}" not found in document.`);
        }
      }

      const requests: docs_v1.Schema$Request[] = [];

      // Build text style request
      const textStyle: docs_v1.Schema$TextStyle = {};
      const textStyleFields: string[] = [];

      if (args.bold !== undefined) {
        textStyle.bold = args.bold;
        textStyleFields.push("bold");
      }
      if (args.italic !== undefined) {
        textStyle.italic = args.italic;
        textStyleFields.push("italic");
      }
      if (args.underline !== undefined) {
        textStyle.underline = args.underline;
        textStyleFields.push("underline");
      }
      if (args.strikethrough !== undefined) {
        textStyle.strikethrough = args.strikethrough;
        textStyleFields.push("strikethrough");
      }
      if (args.fontSize !== undefined) {
        textStyle.fontSize = { magnitude: args.fontSize, unit: "PT" };
        textStyleFields.push("fontSize");
      }
      if (args.foregroundColor !== undefined) {
        const parsed = parseColor(args.foregroundColor);
        textStyle.foregroundColor = { color: { rgbColor: { red: parsed.red, green: parsed.green, blue: parsed.blue } } };
        textStyleFields.push("foregroundColor");
      }
      if (args.linkUrl !== undefined) {
        textStyle.link = { url: args.linkUrl };
        textStyleFields.push("link");
      }

      if (textStyleFields.length > 0) {
        requests.push({
          updateTextStyle: {
            range: { startIndex, endIndex },
            textStyle,
            fields: textStyleFields.join(","),
          },
        });
      }

      // Heading style via paragraph style
      if (args.headingLevel !== undefined) {
        requests.push({
          updateParagraphStyle: {
            range: { startIndex, endIndex },
            paragraphStyle: {
              namedStyleType: HEADING_STYLE[args.headingLevel],
            },
            fields: "namedStyleType",
          },
        });
      }

      if (requests.length === 0) {
        return formatError("No formatting properties specified.");
      }

      await docs.documents.batchUpdate({
        documentId: id,
        requestBody: { requests },
      });

      return formatSuccess(
        `Applied formatting to range [${startIndex}, ${endIndex}).`
      );
    })
  );

  // ── docs_insert_table ────────────────────────────────────────────────────────
  server.tool(
    "docs_insert_table",
    "Inserts a table into a Google Doc using documents.batchUpdate with insertTable, then populates cells with data in a second pass if data is provided; the Docs API requires a 1-based insertion index at a paragraph boundary. Use when the user asks to add a data table to a document. Use when inserting a comparison table or a structured grid of values into a report. Do not use when: inserting text - use docs_write_text instead; inserting an image - use docs_insert_image instead; formatting existing text - use docs_format_text instead; replacing text - use docs_replace_text instead. Returns: 'Table inserted: {rows} rows × {columns} columns.' Parameters: - rows: number of table rows (max 50) - columns: number of table columns (max 20) - index: 1-based insertion index at a paragraph boundary; omit to insert at end (the handler creates a new paragraph automatically) - data: optional array of arrays of strings, outer = rows, inner = cells; e.g. [['Header A','Header B'],['Val 1','Val 2']].",
    {
      docId: z.string().describe("Document ID or URL"),
      rows: z.number().int().min(1).max(50),
      columns: z.number().int().min(1).max(20),
      index: z.number().int().min(1).optional().describe("Insertion index (defaults to end)"),
      data: z
        .array(z.array(z.string()))
        .optional()
        .describe("Row data: outer array = rows, inner = cells"),
    },
    withErrorHandling(async (args) => {
      const docs = await getDocsClient();
      const id = extractFileId(args.docId);

      let insertIndex: number;
      if (args.index !== undefined) {
        insertIndex = args.index;
      } else {
        // The Docs API requires insertTable at a paragraph start boundary.
        // We insert a newline to create a fresh empty paragraph, then use its
        // startIndex for the table.
        const docRes = await docs.documents.get({ documentId: id });
        const body = docRes.data.body?.content ?? [];
        // Find the last paragraph's content area for the newline insertion
        const lastParagraph = [...body].reverse().find((el) => el.paragraph);
        const insertAt = lastParagraph
          ? Math.max(1, (lastParagraph.endIndex ?? 2) - 1)
          : 1;

        await docs.documents.batchUpdate({
          documentId: id,
          requestBody: {
            requests: [{ insertText: { location: { index: insertAt }, text: "\n" } }],
          },
        });
        // Re-fetch to get the updated structure with the new paragraph
        const updatedDoc = await docs.documents.get({ documentId: id });
        const updatedBody = updatedDoc.data.body?.content ?? [];
        // The new empty paragraph is the last structural element
        const lastElement = updatedBody[updatedBody.length - 1];
        insertIndex = lastElement?.startIndex ?? insertAt + 1;
      }

      // Insert the table
      await docs.documents.batchUpdate({
        documentId: id,
        requestBody: {
          requests: [
            {
              insertTable: {
                rows: args.rows,
                columns: args.columns,
                location: { index: insertIndex },
              },
            },
          ],
        },
      });

      // Populate data if provided
      if (args.data && args.data.length > 0) {
        // Re-fetch doc to get actual table cell indices
        const docRes = await docs.documents.get({ documentId: id });
        const body = docRes.data.body?.content ?? [];

        // Find the table we just inserted - look for table elements after our insert point
        const tables = body.filter((el) => el.table !== undefined);
        if (tables.length === 0) {
          return formatSuccess(
            `Table inserted (${args.rows}×${args.columns}) but could not populate data.`
          );
        }

        // Use the last table (most recently inserted)
        const table = tables[tables.length - 1].table;
        const tableRows = table?.tableRows ?? [];
        const requests: docs_v1.Schema$Request[] = [];

        for (let r = 0; r < Math.min(args.data.length, tableRows.length); r++) {
          const rowData = args.data[r];
          const cells = tableRows[r]?.tableCells ?? [];
          for (let c = 0; c < Math.min(rowData.length, cells.length); c++) {
            const cellText = rowData[c];
            if (!cellText) continue;
            const cellContent = cells[c]?.content ?? [];
            if (cellContent.length === 0) continue;
            const cellIndex = cellContent[0]?.startIndex;
            if (cellIndex === undefined || cellIndex === null) continue;
            requests.push({
              insertText: {
                location: { index: cellIndex },
                text: cellText,
              },
            });
          }
        }

        if (requests.length > 0) {
          // Reverse-order to preserve indices
          requests.reverse();
          await docs.documents.batchUpdate({
            documentId: id,
            requestBody: { requests },
          });
        }
      }

      return formatSuccess(`Table inserted: ${args.rows} rows × ${args.columns} columns.`);
    })
  );

  // ── docs_insert_image ────────────────────────────────────────────────────────
  server.tool(
    "docs_insert_image",
    "Inserts an inline image into a Google Doc using documents.batchUpdate with insertInlineImage; the image source is either a public URL or a Drive file ID, and the insertion uses a 1-based character index. Use when the user asks to add an image or logo to a document. Use when embedding a chart screenshot or diagram from Drive into a report. Do not use when: inserting text - use docs_write_text instead; inserting a table - use docs_insert_table instead; formatting text - use docs_format_text instead; reading document content - use docs_get_text instead. Returns: 'Image inserted at index {index}.' Parameters: - imageUrl: public HTTPS URL of the image (mutually exclusive with driveFileId) - driveFileId: Drive file ID of an image already in Drive (mutually exclusive with imageUrl) - index: 1-based insertion index; omit to insert at the document end - width/height: optional dimensions in points (1 pt = 1/72 inch).",
    {
      docId: z.string().describe("Document ID or URL"),
      imageUrl: z.string().optional().describe("Public image URL"),
      driveFileId: z.string().optional().describe("Drive file ID of an image"),
      index: z.number().int().min(1).optional().describe("Insertion index (defaults to end)"),
      width: z.number().positive().optional().describe("Width in pt"),
      height: z.number().positive().optional().describe("Height in pt"),
    },
    withErrorHandling(async (args) => {
      if (!args.imageUrl && !args.driveFileId) {
        return formatError("Provide either imageUrl or driveFileId.");
      }

      const docs = await getDocsClient();
      const id = extractFileId(args.docId);

      let insertIndex: number;
      if (args.index !== undefined) {
        insertIndex = args.index;
      } else {
        const docRes = await docs.documents.get({ documentId: id, fields: "body" });
        insertIndex = getDocEndIndex(docRes.data);
      }

      const imageProperties: docs_v1.Schema$Size = {};
      if (args.width !== undefined) {
        imageProperties.width = { magnitude: args.width, unit: "PT" };
      }
      if (args.height !== undefined) {
        imageProperties.height = { magnitude: args.height, unit: "PT" };
      }

      const uri =
        args.imageUrl ??
        `https://drive.google.com/uc?id=${extractFileId(args.driveFileId!)}`;

      const imageRequest: docs_v1.Schema$Request = {
        insertInlineImage: {
          location: { index: insertIndex },
          uri,
          ...(Object.keys(imageProperties).length > 0
            ? { objectSize: imageProperties }
            : {}),
        },
      };

      await docs.documents.batchUpdate({
        documentId: id,
        requestBody: { requests: [imageRequest] },
      });

      return formatSuccess(`Image inserted at index ${insertIndex}.`);
    })
  );
}
