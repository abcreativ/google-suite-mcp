# Changelog

## [1.1.0] - 2026-04-13

### Improved

- Rewrote all 82 tool descriptions using a canonical 6-slot template optimized for Glama.ai TDQS scoring.
- Every description now opens with the underlying API method, includes two concrete "Use when" scenarios, lists all sibling tools with routing conditions in "Do not use when", documents the exact Returns string with template literals, and provides format examples for key parameters.
- Descriptions are capped at 180 words to preserve the Conciseness dimension score.
- Projected TDQS scores: Overall 4.75, Purpose Clarity 5.0, Usage Guidelines 5.0, Contextual Completeness 5.0, Behavioral Transparency 4.5, Parameter Semantics 4.5, Conciseness 4.5.

### Fixed

- Corrected 8 Returns block mismatches where descriptions used ASCII hyphens and "x" instead of the Unicode arrows, en-dashes, and multiplication signs that handlers actually emit (sheets_create_named_range, sheets_rename_sheet, sheets_rename, sheets_list_sheets, sheets_get_info, sheets_resize_columns, sheets_resize_rows, docs_insert_table).

## [1.0.0] - 2026-04-12

### Added

- Initial release: 82 tools covering Google Sheets (spreadsheet management, tab management, reading, writing, formulas, formatting, validation, charts, named ranges, protection, filters, Apps Script), Google Drive (file search, upload, download, move, copy, share, permissions), Google Docs (create, write, format, replace, tables, images, export), and Apps Script (create, read, update, deploy, run).
- OAuth2 authentication with automatic browser flow and token refresh.
- Compatible with Claude Desktop, Cursor, Windsurf, VS Code, Gemini CLI, and any MCP client.
