# Changelog

## [1.1.0] - 2026-04-13

This release is description-quality only. No API changes, no handler behavior changes, no breaking changes. Upgrading is safe without re-testing any tool calls.

### Improved

- Rewrote all 83 tool descriptions using a canonical 6-slot template optimized for Glama.ai Tool Description Quality Score (TDQS). Every description now opens with the underlying Google API method, lists two concrete "Use when" scenarios, enumerates sibling tools with routing conditions in "Do not use when", documents the exact Returns string as a template literal, and provides format examples for key parameters.
- Upgraded Zod `.describe()` strings on `spreadsheet_id` and `script_id` parameters across all tools to include format hints (e.g. "sheet ID from the URL, the token between /d/ and /edit") instead of bare type labels.
- Strengthened sibling cross-references for the write-range, chart, and Drive-search clusters so every tool in a cluster names every adjacent tool by exact name with a one-clause routing reason.
- Descriptions are capped at ~180 words to preserve the Conciseness TDQS dimension.

Baseline internal TDQS (measured before this release against Glama's six-dimension rubric) was 3.24/5 overall, with Usage Guidelines at 2.1/5. After this rewrite an internal re-grade against the same rubric clears all phase targets (Overall ≥ 4.6, Usage Guidelines = 5/5, every other dimension ≥ 4.5 except Conciseness floor ≥ 3.8). The authoritative score will be whatever Glama's own crawl reports.

### Fixed

- Corrected 9 Returns block mismatches where descriptions used ASCII characters ("x", hyphens, ASCII arrows) instead of the Unicode characters (×, en-dashes, →) that handlers actually emit. Affected tools: sheets_create_named_range, sheets_rename_sheet, sheets_rename, sheets_list_sheets, sheets_get_info, sheets_resize_columns, sheets_resize_rows, sheets_write_table, docs_insert_table.

## [1.0.0] - 2026-04-12

### Added

- Initial release: 83 tools covering Google Sheets (spreadsheet management, tab management, reading, writing, formulas, formatting, validation, charts, named ranges, protection, filters, find-replace, Apps Script integration), Google Drive (file search, metadata, upload, download, move, copy, rename, trash, share, permissions), Google Docs (create, write, format, replace, tables, images, export), and Apps Script (create, read, update, version, deploy, run).
- OAuth2 authentication with automatic browser flow and persistent refresh token.
- Compatible with Claude Desktop, Cursor, Windsurf, VS Code, Gemini CLI, and any MCP client.
