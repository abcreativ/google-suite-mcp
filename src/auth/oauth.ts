/**
 * OAuth2 authentication for Google APIs.
 *
 * Credentials and tokens are stored in ~/.google-suite-mcp/:
 *   credentials.json  - OAuth2 client ID + secret (user provides this)
 *   tokens.json       - Access/refresh tokens (written after first auth)
 *
 * On first run the user is prompted to visit a URL in their browser, then
 * the auth code is captured via a localhost redirect server.
 *
 * On subsequent runs, the saved refresh token is used automatically.
 */

import fs from "node:fs";
import http from "node:http";
import os from "node:os";
import path from "node:path";
import { URL } from "node:url";
import { OAuth2Client } from "google-auth-library";
import open from "open";

// ─── Constants ───────────────────────────────────────────────────────────────

const CONFIG_DIR = path.join(os.homedir(), ".google-suite-mcp");
const CREDENTIALS_PATH = path.join(CONFIG_DIR, "credentials.json");
const TOKENS_PATH = path.join(CONFIG_DIR, "tokens.json");

const SCOPES = [
  "https://www.googleapis.com/auth/spreadsheets",
  "https://www.googleapis.com/auth/drive",
  "https://www.googleapis.com/auth/documents",
  "https://www.googleapis.com/auth/script.projects",
  "https://www.googleapis.com/auth/script.deployments",
];

// ─── Types ────────────────────────────────────────────────────────────────────

interface Credentials {
  installed?: OAuthClientConfig;
  web?: OAuthClientConfig;
}

interface OAuthClientConfig {
  client_id: string;
  client_secret: string;
  redirect_uris: string[];
}

interface TokenData {
  access_token?: string | null;
  refresh_token?: string | null;
  token_type?: string | null;
  expiry_date?: number | null;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

/**
 * Reads and validates credentials.json from the config directory.
 * Exits with a clear error message if the file is missing or malformed.
 */
function loadCredentials(): OAuthClientConfig {
  if (!fs.existsSync(CREDENTIALS_PATH)) {
    console.error(`
╔══════════════════════════════════════════════════════════════╗
║         google-suite-mcp - First-Time Setup Required         ║
╚══════════════════════════════════════════════════════════════╝

No credentials found at: ${CREDENTIALS_PATH}

To set up authentication:

  1. Go to https://console.cloud.google.com/
  2. Create a project (or select an existing one)
  3. Enable the following APIs:
       • Google Sheets API
       • Google Drive API
       • Google Docs API
  4. Navigate to "APIs & Services" → "Credentials"
  5. Click "Create Credentials" → "OAuth client ID"
  6. Choose "Desktop app" as the application type
  7. Download the JSON file and save it to:
       ${CREDENTIALS_PATH}
  8. Run this server again to complete authentication.
`);
    process.exit(1);
  }

  let raw: Credentials;
  try {
    raw = JSON.parse(fs.readFileSync(CREDENTIALS_PATH, "utf-8")) as Credentials;
  } catch {
    console.error(`Failed to parse credentials.json at ${CREDENTIALS_PATH}`);
    process.exit(1);
  }

  const config = raw.installed ?? raw.web;
  if (!config?.client_id || !config?.client_secret) {
    console.error(
      `credentials.json is missing client_id or client_secret. Please re-download it from Google Cloud Console.`
    );
    process.exit(1);
  }

  return config;
}

/**
 * Runs the interactive browser-based OAuth2 flow.
 * Opens the consent URL in the default browser and waits for the redirect.
 */
async function runAuthFlow(client: OAuth2Client): Promise<TokenData> {
  // Start the redirect server first so we know the port
  let resolvePort!: (port: number) => void;
  const portPromise = new Promise<number>((res) => {
    resolvePort = res;
  });

  // We need to start the server before generating the URL so we know the port.
  // Build a minimal HTTP server here, capture the port, then generate the URL.
  let authCodeResolve!: (code: string) => void;
  let authCodeReject!: (err: Error) => void;
  const authCodePromise = new Promise<string>((res, rej) => {
    authCodeResolve = res;
    authCodeReject = rej;
  });

  const server = http.createServer((req, res) => {
    if (!req.url) {
      res.end();
      return;
    }
    const base = `http://127.0.0.1`;
    const url = new URL(req.url, base);
    const code = url.searchParams.get("code");
    const error = url.searchParams.get("error");

    if (error) {
      res.writeHead(400, { "Content-Type": "text/html" });
      res.end(`<h1>Authentication failed</h1><p>${error}</p>`);
      server.close();
      authCodeReject(new Error(`OAuth2 authorization error: ${error}`));
      return;
    }

    if (code) {
      res.writeHead(200, { "Content-Type": "text/html" });
      res.end(
        `<h1>Authentication successful!</h1><p>You may close this tab and return to your terminal.</p>`
      );
      server.close(() => authCodeResolve(code));
      return;
    }

    res.writeHead(204);
    res.end();
  });

  await new Promise<void>((res, rej) => {
    server.listen(0, "127.0.0.1", () => {
      const address = server.address();
      if (!address || typeof address === "string") {
        rej(new Error("Failed to start redirect server"));
        return;
      }
      resolvePort(address.port);
      res();
    });
    server.on("error", rej);
  });

  const port = await portPromise;
  const redirectUri = `http://127.0.0.1:${port}`;

  const authUrl = client.generateAuthUrl({
    access_type: "offline",
    scope: SCOPES,
    prompt: "consent", // always request refresh_token
    redirect_uri: redirectUri,
  });

  console.error(`\nOpening browser for Google authentication...`);
  console.error(`If the browser does not open automatically, visit:\n\n  ${authUrl}\n`);

  try {
    await open(authUrl);
  } catch {
    // Silently ignore - user can visit the URL manually
  }

  console.error(`Waiting for authorization (listening on ${redirectUri})...`);
  const code = await authCodePromise;

  const { tokens } = await client.getToken({ code, redirect_uri: redirectUri });
  return tokens as TokenData;
}

/**
 * Persists token data to disk so subsequent runs skip the browser flow.
 */
function saveTokens(tokens: TokenData): void {
  fs.mkdirSync(CONFIG_DIR, { recursive: true });
  // Create with restricted permissions from the start (rw-------)
  fs.writeFileSync(TOKENS_PATH, JSON.stringify(tokens, null, 2), {
    encoding: "utf-8",
    mode: 0o600,
  });
}

// ─── Singleton ────────────────────────────────────────────────────────────────

let cachedClient: OAuth2Client | null = null;

/**
 * Returns an authenticated OAuth2Client, ready for use with googleapis.
 *
 * - If tokens.json exists: loads saved credentials and returns immediately.
 * - If tokens.json is missing: launches browser auth flow, saves tokens, returns.
 *
 * The returned client automatically refreshes the access token using the
 * stored refresh token when it expires.
 */
export async function getAuthClient(): Promise<OAuth2Client> {
  if (cachedClient) return cachedClient;

  const credentials = loadCredentials();

  const client = new OAuth2Client({
    clientId: credentials.client_id,
    clientSecret: credentials.client_secret,
    // Redirect URI will be set dynamically during the auth flow
    redirectUri: "http://127.0.0.1",
  });

  let tokens: TokenData;

  if (fs.existsSync(TOKENS_PATH)) {
    // Load previously saved tokens
    try {
      tokens = JSON.parse(fs.readFileSync(TOKENS_PATH, "utf-8")) as TokenData;
    } catch {
      console.error(
        `Failed to parse tokens.json. Re-authenticating...`
      );
      tokens = await runAuthFlow(client);
      saveTokens(tokens);
    }
    client.setCredentials(tokens);
  } else {
    // First run - launch the browser flow
    tokens = await runAuthFlow(client);
    client.setCredentials(tokens);
    saveTokens(tokens);
    console.error("Authentication successful. Tokens saved to", TOKENS_PATH);
  }

  // Single listener for automatic token refreshes (avoids duplicate registration)
  client.on("tokens", (newTokens) => {
    const merged = { ...tokens, ...newTokens };
    try {
      saveTokens(merged);
    } catch (err) {
      console.error("[google-suite-mcp] Failed to persist refreshed tokens:", err);
    }
    tokens = merged;
    client.setCredentials(merged);
  });

  cachedClient = client;
  return client;
}

/**
 * Forces a token refresh, updating tokens.json with fresh credentials.
 * Returns the new expiry date so callers can report it to the user.
 */
export async function refreshTokens(): Promise<{ expiryDate: number | null }> {
  const client = await getAuthClient();
  const response = await client.refreshAccessToken();
  const tokens = response.credentials as TokenData;
  saveTokens({ ...client.credentials, ...tokens });
  return { expiryDate: tokens.expiry_date ?? null };
}

/**
 * Returns the current token state for display purposes.
 */
export function getTokenInfo(): {
  hasTokens: boolean;
  hasRefreshToken: boolean;
  expiryDate: number | null;
  credentialsPath: string;
  tokensPath: string;
} {
  const hasTokens = fs.existsSync(TOKENS_PATH);

  if (!hasTokens || !cachedClient) {
    return {
      hasTokens,
      hasRefreshToken: false,
      expiryDate: null,
      credentialsPath: CREDENTIALS_PATH,
      tokensPath: TOKENS_PATH,
    };
  }

  const creds = cachedClient.credentials;
  return {
    hasTokens: true,
    hasRefreshToken: !!creds.refresh_token,
    expiryDate: (creds.expiry_date as number | undefined) ?? null,
    credentialsPath: CREDENTIALS_PATH,
    tokensPath: TOKENS_PATH,
  };
}
