# Installing google-suite-mcp

This guide is written to be followed by a human OR by an AI assistant helping a human. If you want AI help, paste this file (or the link to it) into Claude Code, Claude Desktop, Cursor, or ChatGPT and ask the AI to walk you through every step. The steps below are imperative and copy-pasteable on purpose.

## What you need before you start

- Node.js 20 or later installed (`node --version` to check)
- A Google account you can sign into
- About 10 minutes
- A text editor of any kind

## Step 1: Clone and build

Run these commands in your terminal, one at a time:

```bash
git clone https://github.com/abcreativ/google-suite-mcp.git
cd google-suite-mcp
npm install
npm run build
```

After the last command finishes, there will be a `dist/` folder in the project directory with `dist/index.js` inside it. You will need the absolute path to `dist/index.js` in Step 7. To get it now, run:

```bash
echo "$(pwd)/dist/index.js"
```

Copy the output and keep it somewhere. You will paste it later.

## Step 2: Create a Google Cloud project

Open this link: https://console.cloud.google.com/projectcreate

Fill in:
- Project name: `google-suite-mcp` (or any name you like)
- Organization: leave at the default
- Location: leave at the default

Click **Create**. Wait about 10 seconds for the project to finish creating. Leave the browser tab open.

## Step 3: Enable the four Google APIs

Open each of these four links in turn. On every page, click the blue **Enable** button and wait for it to finish before moving on.

1. Sheets API: https://console.cloud.google.com/apis/library/sheets.googleapis.com
2. Docs API: https://console.cloud.google.com/apis/library/docs.googleapis.com
3. Drive API: https://console.cloud.google.com/apis/library/drive.googleapis.com
4. Apps Script API: https://console.cloud.google.com/apis/library/script.googleapis.com

Each one takes about 5 seconds. If you see a page that says "API Enabled" with a green check, you are done with that API.

## Step 4: Configure the OAuth consent screen

Open: https://console.cloud.google.com/apis/credentials/consent

On the first page:
- **User Type**: choose **External**
- Click **Create**

On the next page fill in:
- **App name**: `google-suite-mcp`
- **User support email**: your own Google address
- **Developer contact information**: your own Google address

Click **Save and Continue**. On the **Scopes** page, click **Save and Continue** again (no changes needed).

On the **Test users** page, click **Add Users** and enter your own Google email address. This is the account you will use with google-suite-mcp. Click **Save and Continue**.

Click **Back to Dashboard**.

## Step 5: Create OAuth 2.0 Desktop credentials

Open: https://console.cloud.google.com/apis/credentials

Click **Create Credentials**, then **OAuth client ID**.

Fill in:
- **Application type**: `Desktop app` (this is important, do not pick Web application)
- **Name**: `google-suite-mcp` (or any name)

Click **Create**. A dialog appears showing your **Client ID** and **Client Secret**. Copy both values. You will paste them in the next step.

## Step 6: Paste credentials into .env

In the project directory from Step 1, run:

```bash
cp .env.example .env
```

Open `.env` in your text editor. You will see two lines:

```
GOOGLE_CLIENT_ID=your-client-id.apps.googleusercontent.com
GOOGLE_CLIENT_SECRET=your-client-secret
```

Replace the placeholder values with the real Client ID and Client Secret from Step 5. Save the file and close it.

## Step 7: Add google-suite-mcp to your AI client

### Claude Desktop

Open your Claude Desktop config file:

- macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`
- Windows: `%APPDATA%\Claude\claude_desktop_config.json`

If the file does not exist yet, create it. Add this block (if the file already has content, merge the `mcpServers` key into the existing JSON):

```json
{
  "mcpServers": {
    "google-suite": {
      "command": "node",
      "args": ["PASTE_YOUR_ABSOLUTE_PATH_HERE"]
    }
  }
}
```

Replace `PASTE_YOUR_ABSOLUTE_PATH_HERE` with the absolute path you copied at the end of Step 1. Claude Desktop does not expand `~` inside the JSON, so the path must be fully qualified. The string you paste will end in `/dist/index.js`.

Save the file. Fully quit Claude Desktop (on macOS, use the menu bar Quit, not just close the window). Reopen Claude Desktop.

### Cursor, Windsurf, Zed, VS Code, Gemini CLI, and other MCP clients

Use the same `command` and `args` block above. Each client has its own config location, but the shape is identical. See your client's MCP documentation for where the config file lives.

## Step 8: Authorize with Google and test

In Claude Desktop (or your MCP client), start a new chat and ask:

> Use the auth_status tool from google-suite to check if it is connected.

The first time ANY google-suite-mcp tool runs, your default web browser will open a Google consent screen. Sign in with the same Google account you added as a test user in Step 4. You will see a warning that says "Google hasn't verified this app". This is expected because you are the app developer in testing mode. Click **Advanced**, then **Go to google-suite-mcp (unsafe)**, then **Continue**. Grant the requested permissions.

The browser will show a success page. The refresh token is now cached on your machine in `tokens.json` inside the project directory. You will not have to authenticate again unless you revoke access.

To confirm everything works, try this prompt:

> List the first 5 files I have in my Google Drive.

If the AI returns a list of real files from your Drive, setup is complete. You are done.

## Troubleshooting

**"Error 403: access_denied" on the Google consent screen**
You did not add your Google address as a test user in Step 4. Go back to https://console.cloud.google.com/apis/credentials/consent, scroll to Test Users, click Add Users, and add your email.

**"redirect_uri_mismatch" error**
You picked the wrong OAuth client type in Step 5. It must be **Desktop app**, not Web application. Go to https://console.cloud.google.com/apis/credentials, delete the existing OAuth client, and create a new one with the correct type.

**Claude Desktop says "MCP server failed to start"**
The path in your Claude Desktop config is wrong. Run `pwd` in the project directory, confirm `dist/index.js` exists by running `ls dist/index.js`, then paste the full absolute path into the config. On Windows, use forward slashes or doubled backslashes inside the JSON string.

**"No tools found" or the google-suite tools do not appear in Claude Desktop**
You did not fully quit Claude Desktop before reopening. On macOS, click **Claude** in the top menu bar and choose **Quit Claude**, not just close the window. On Windows, right-click the tray icon and choose Quit.

**Anything else**
Open a GitHub issue at https://github.com/abcreativ/google-suite-mcp/issues with the error message and the step you were on.
