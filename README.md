<!--
GitHub repo description (paste into the "About" field):
Google Workspace MCP server for Claude Desktop, Cursor, Windsurf, VS Code, Gemini CLI, and any MCP client. 82 tools across Sheets, Docs, Drive, and Apps Script. Full read/write. 30/30 live-API tests. MIT.
-->

# google-suite-mcp

**You have an AI that can think. Now give it hands.**

An open-source Google Workspace MCP server that gives Claude Desktop, Cursor, Windsurf, Zed, VS Code (GitHub Copilot), Gemini CLI, and any other Model Context Protocol client full read/write control over Google Sheets, Docs, Drive, and Apps Script. 82 tools. One server. MIT licensed.

`google-suite-mcp` is the first **Workspace Operator**: the only MCP that treats Google Workspace as a runtime an AI can operate. It builds dashboards, deploys Apps Scripts, and formats documents in a single call, instead of exposing Workspace as a read-only surface to sample.

One 10-minute OAuth setup is the entire cost of entry. After that, a single natural-language prompt can build a multi-cell KPI dashboard (`sheets_build_dashboard`), turn a schema into a fully formatted sheet (`sheets_build_sheet`), or publish a Google Apps Script as a live web app (`script_deploy`). Those are three composed primitives no other Google MCP we've seen ships. Every tool is proven: **30 out of 30 end-to-end tests pass against live Google APIs, not mocks.** If a tool is listed here, it has been executed against Google's production endpoints and returned the expected result.

> MCP is a protocol, not a Claude feature. Any client that speaks the Model Context Protocol can use this server, regardless of which model is behind it.

> **Not sure where to start?** Paste this repo URL into Claude Code, Claude Desktop, Cursor, or ChatGPT and ask it to help you install. [`SETUP.md`](./SETUP.md) is written so your AI can walk you through every step, open the Google Cloud links for you, and verify everything works at the end.

---

## What you can do in 60 seconds of prompting

These are real single-prompt outcomes, not roadmap items. Paste any of them into your MCP client once the server is wired in.

- "Build me a KPI dashboard in the Q4 sheet with revenue, CAC, churn, and MRR, formatted, with conditional colors." → one call to `sheets_build_dashboard`.
- "Create a new sheet called Clients with these 12 columns, header styling, data validation, and frozen first row." → one call to `sheets_build_sheet`.
- "Find every instance of 'Q3 2025' across all tabs in this workbook and replace it with 'Q4 2025'." → one call to `sheets_find_replace_many`.
- "Create an Apps Script bound to this sheet that emails me a summary every Monday at 8am, and deploy it as a web app." → one call to `script_deploy`.
- "Insert a 5-column pricing table into this Google Doc with these rows." → one call to `docs_insert_table`.
- "Create a shared folder in Drive, move these three files into it, and set permissions to anyone-with-link can view."
- "Add a conditional format to highlight any row where margin is under 15 percent red."
- "Protect the formulas in column H so nobody else on the sheet can edit them."

No code. No manual steps. No context-switching out of your AI client.

---

## Why another Google Workspace MCP server?

Most Google MCPs on GitHub fall into one of two buckets: read-only connectors that can query a spreadsheet but not change it, or narrow Sheets-only adapters that ignore Docs, Drive, and Apps Script entirely. They are **connectors**: thin wrappers around the REST API, handed to an AI that then has to spend forty tool calls and a fortune in tokens to accomplish anything.

`google-suite-mcp` is not a connector. It is an **operator**. The primitives are *outcomes* (build this dashboard, deploy this script, format this report), not *endpoints* (read range, write cell, list file). Operators cover the full suite because real work crosses tools. Operators ship composed primitives because real work is never one cell edit.

They zig. We zag.

### Capability comparison

| Capability | Read-only MCPs | Sheets-only MCPs | google-suite-mcp |
|---|---|---|---|
| Read Google Sheets | Yes | Yes | Yes |
| Write to Google Sheets | No | Yes | Yes |
| Rich formatting and styles | No | Partial | Yes |
| Conditional formatting | No | Rare | Yes |
| Charts, named ranges, protected ranges | No | Rare | Yes |
| Data validation, filters, sorts | No | Rare | Yes |
| One-call dashboard builder | No | No | Yes (`sheets_build_dashboard`) |
| Schema-to-sheet builder | No | No | Yes (`sheets_build_sheet`) |
| Cross-sheet find and replace | No | No | Yes (`sheets_find_replace_many`) |
| Google Docs read and write | No | No | Yes |
| Docs table builder | No | No | Yes (`docs_insert_table`) |
| Google Drive file operations | No | No | Yes |
| Apps Script create, run, deploy | No | No | Yes (`script_deploy`) |
| Live-API test coverage | Unknown | Partial | 30 / 30 E2E tests |
| Token-efficient responses | No | No | Yes, audited |
| License | Mixed | Mixed | MIT |

---

## Requirements

- **Node.js 20 or later**
- **A Google account** with access to the Workspace files you want your AI to touch
- **A Google Cloud project** (free tier is fine)
- **OAuth 2.0 Desktop credentials**
- **An MCP-compatible client**: Claude Desktop, Cursor, Windsurf, Zed, VS Code (GitHub Copilot), Gemini CLI, Cline, Goose, any agent built on the OpenAI Agents SDK, or any other client that speaks MCP

Budget roughly ten minutes for first-time setup if you have never touched Google Cloud before. Done once, never again.

---

## How do I install google-suite-mcp?

```bash
git clone https://github.com/abcreativ/google-suite-mcp.git
cd google-suite-mcp
npm install
npm run build
```

That compiles the TypeScript server into `dist/`. Point your MCP client at `dist/index.js` once OAuth is configured (next section).

---

## How do I connect this to Google Workspace?

This is the only part of setup that needs real attention. Every link below opens the exact Google Cloud Console page where your next click lives, so you never have to hunt for anything.

**Full walkthrough:** [`SETUP.md`](./SETUP.md) has every step in order with troubleshooting for the top errors. The short version:

1. **[Create a Google Cloud project](https://console.cloud.google.com/projectcreate)** (10 seconds, free tier)
2. **Enable the four APIs** (click each link and press Enable):
   - [Sheets API](https://console.cloud.google.com/apis/library/sheets.googleapis.com)
   - [Docs API](https://console.cloud.google.com/apis/library/docs.googleapis.com)
   - [Drive API](https://console.cloud.google.com/apis/library/drive.googleapis.com)
   - [Apps Script API](https://console.cloud.google.com/apis/library/script.googleapis.com)
3. **[Configure the OAuth consent screen](https://console.cloud.google.com/apis/credentials/consent)**: choose External and add your own Google email as a test user
4. **[Create an OAuth 2.0 Desktop client](https://console.cloud.google.com/apis/credentials)**: Create Credentials > OAuth client ID > Desktop app. Copy the Client ID and Client Secret.
5. **Paste credentials into `.env`**:

   ```bash
   cp .env.example .env
   ```

   Then edit `.env` and paste the Client ID and Client Secret from Step 4.

The first time the server runs a tool it opens your browser, walks you through Google's consent flow, and caches the refresh token locally. You will not authenticate again unless you revoke access.

If you get stuck on any step, paste [`SETUP.md`](./SETUP.md) into your AI assistant and ask it to walk you through one step at a time.

---

## How do I connect my MCP client?

Every MCP-compatible client accepts the same two things: a `command` to run and a list of `args`. Use the absolute path to `dist/index.js` from the install step.

### Claude Desktop

Open your config file:

- **macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows:** `%APPDATA%\Claude\claude_desktop_config.json`

Add this entry:

```json
{
  "mcpServers": {
    "google-suite": {
      "command": "node",
      "args": ["/absolute/path/to/google-suite-mcp/dist/index.js"]
    }
  }
}
```

Restart Claude Desktop. The 82 tools appear in the tool picker.

### Cursor, Windsurf, Zed, VS Code

Each client has its own MCP config location but accepts the same `command` and `args` shape. Consult your client's MCP docs and paste the block above under its MCP server section. No other changes are required.

### Gemini CLI

Add the server to your Gemini CLI MCP config (`~/.gemini/config.json` or equivalent) using the same `command` and `args`.

### OpenAI Agents SDK

The OpenAI Agents SDK (Python and TypeScript) supports MCP servers natively. Pass `google-suite-mcp` as an MCP server when constructing your agent and the 82 tools become available to any OpenAI model you choose.

### Any other MCP client

If it speaks the Model Context Protocol, it works. Use the same `command` and `args` pattern wherever your client defines MCP servers.

---

## Usage examples

Once the server is wired in, you talk to your AI the way you always have. It just has hands now.

**Build a live KPI dashboard from a schema**

> "In the workbook titled 'Q4 Forecast', create a new tab called 'Dashboard' and build a KPI dashboard with four cells: Revenue, Gross Margin, CAC, and Churn. Pull the values from the 'Raw' tab, format headers bold, numbers as currency, and highlight anything below target in red."

One `sheets_build_dashboard` call. Done.

**Turn a schema into a sheet**

> "Create a new sheet called 'Client Tracker' with columns for Name, Email, Status (dropdown: Lead / Active / Churned), Last Contact (date), and Notes. Add conditional formatting so Churned rows turn red."

One `sheets_build_sheet` call. Done.

**Write a Doc with real structure**

> "Draft the kickoff brief for the Henderson project in Docs. Include a stakeholder table, a timeline table, and a risks section with bullets."

`docs_insert_table` gives you actual tables, not ASCII imitations.

**Deploy an Apps Script web app**

> "Create a new Apps Script project bound to this sheet, add a doGet that returns the Summary tab as JSON, version it, and deploy it as a web app I can curl."

`script_deploy` publishes it. You get a live URL.

**Rename a field across every sheet in a workbook**

> "In the Expenses workbook, replace every occurrence of 'customer_id' with 'account_id' across all sheets in one pass."

`sheets_find_replace_many` handles it in one call.

**Organize Drive**

> "Create a folder called '2026 Client Intake', move every file in my Drive with 'intake' in the name into it, and share the folder with view access to anyone with the link."

---

## Tool list summary (82 tools)

Grouped by Google surface:

- **Google Sheets (53 tools).** Create, list, read, write, append, format, conditional format, chart, named range, protected range, filter, sort, validation, borders, merge, freeze panes, resize, find/replace (single and `sheets_find_replace_many` across every tab), search, formulas, array formulas, batch update, duplicate, rename, reorder, delete, **`sheets_build_sheet`**, **`sheets_build_dashboard`**.
- **Google Docs (8 tools).** Create, write, format text, get text, replace text, insert image, **`docs_insert_table`**, export.
- **Google Drive (12 tools).** Upload, download, search, get info, move, copy, rename, trash, create folder, share, list and update permissions.
- **Google Apps Script (7 tools).** Create, update, get, get bound, run, create version, **`script_deploy`**.
- **Auth (2 tools).** Status, refresh.

Every tool ships with **token-efficient responses**. Every payload was audited end-to-end and tightened so your context window stays lean on long agentic runs.

For the authoritative list with live schemas, call `tools/list` from any MCP client once connected, or see `src/tools/` in the source tree.

---

## FAQ

### Does this work with Cursor, Gemini CLI, or other non-Claude MCP clients?

Yes. google-suite-mcp is a standard Model Context Protocol server. Any MCP-compatible client (Claude Desktop, Cursor, Windsurf, Zed, VS Code with GitHub Copilot, Gemini CLI, Cline, Goose, or any agent built on the OpenAI Agents SDK) can connect using the same `command` and `args` pattern shown above. Consult your client's MCP configuration docs for where to paste the block. MCP is a protocol, not a Claude feature.

### Can I use this with OpenAI models like GPT-4 or GPT-5?

Yes. The OpenAI Agents SDK supports MCP servers natively, so any agent you build on top of OpenAI models can mount `google-suite-mcp` and get all 82 tools. The server does not care which model is on the other end of the connection.

### Can I deploy Apps Script from AI?

Yes. google-suite-mcp is one of the only MCP servers that exposes `script_deploy`, letting your AI publish an Apps Script project as a web app, an API executable, or a scheduled trigger without leaving the chat. You can create a project, upload source files, version it, and deploy it end to end in a single conversation.

### Is it read-only?

No. It is full read and write across all four surfaces: Sheets, Docs, Drive, and Apps Script. Every tool that makes sense to write is writable, including formatting, validation, protected ranges, permissions, and Apps Script deployment.

### Can I self-host it?

Yes. Self-hosted is the default and only supported configuration. The server runs locally on your machine using your own OAuth 2.0 credentials. No data passes through any third-party service. You control the Google account it authenticates against and can revoke access at any time from your Google account security settings.

### Is it safe to give an AI write access to my Google Drive?

The security model is: your machine, your OAuth credentials, your account. Nothing is sent to a third party. That said, write access is write access. Always review what your AI is about to change in shared or production documents before you tell it to proceed. Scope the OAuth client to a dedicated Google account if you want a hard isolation boundary.

### Does it work with personal Gmail and paid Google Workspace accounts?

Both. The underlying Google APIs are identical across personal Gmail and paid Workspace tenants. Workspace administrators may need to allowlist the OAuth client in their admin console for organization-wide use.

### How does this compare to Zapier, Make, or n8n for Google Workspace?

Those are visual automation tools where you pre-build every workflow in advance. google-suite-mcp is different: the AI decides which tool to call at runtime based on what you ask for in natural language. There is no workflow to build. You describe the outcome, the model picks the tools, the server executes them.

### Is google-suite-mcp production ready?

The test suite runs 30 end-to-end tests against live Google APIs (not mocks), and all 30 pass. Treat it like any other open-source project: audit the code, scope the OAuth credentials tightly, and watch what it writes. If a tool breaks against live APIs, that is a bug and the test suite is designed to catch it before it ships.

---

## Contributing

Pull requests welcome. The test suite is the contract.

1. **Tests hit live APIs, not mocks.** The current suite runs 30 of 30 green against real Google services and that bar does not move. If you add a tool, add an end-to-end test that hits the live endpoint.
2. **Composed primitives beat endpoint mirrors.** If your PR adds a new tool, explain what *outcome* it delivers, not just what endpoint it wraps. If you find yourself writing a thin wrapper around a single REST endpoint, stop and ask whether the AI needs that call or a composed outcome two layers up.
3. **Token efficiency is a feature.** Every token the server returns is a token the user pays for downstream. Keep responses lean. See existing tools for the pattern.
4. **One zag at a time.** Open an issue first for anything that changes the category shape.

Run `npm run build` and the E2E suite locally (you will need your own Google Cloud credentials) before opening a pull request. Bug reports and feature requests go to GitHub Issues. Security issues: please open a private advisory rather than a public issue.

---

## License

MIT. Use it commercially, fork it, ship it inside your own product. No attribution required, though a GitHub star is appreciated.

See [`LICENSE`](./LICENSE) for full text.

---

## Topics

`google-workspace` `google-sheets` `google-docs` `google-drive` `google-apps-script` `mcp` `mcp-server` `model-context-protocol` `claude` `claude-desktop` `cursor` `openai` `gemini` `anthropic` `ai-agents` `ai-tools` `llm-tools` `typescript` `nodejs` `oauth2` `spreadsheet-automation` `workspace-automation`
