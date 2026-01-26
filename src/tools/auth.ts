/**
 * Authentication tools for Outlook MCP Server
 */

import { z } from "zod";
import http from "http";
import { exec, spawn } from "child_process";
import path from "path";
import { fileURLToPath } from "url";
import { loadTokens, createTestTokens } from "../api.js";

// =============================================================================
// Configuration
// =============================================================================

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const AUTH_SERVER_URL = "http://localhost:1337";
const CLIENT_ID = process.env.AZURE_CLIENT_ID || "";
const USE_TEST_MODE = process.env.USE_TEST_MODE === "true";
const SERVER_VERSION = "2.0.0";

// =============================================================================
// Schemas
// =============================================================================

export const aboutSchema = z.object({});

export const authenticateSchema = z.object({
  force: z.boolean().optional().describe("Force re-authentication even if already authenticated"),
});

export const checkAuthStatusSchema = z.object({});

// =============================================================================
// Helper Functions
// =============================================================================

async function isAuthServerRunning(): Promise<boolean> {
  return new Promise((resolve) => {
    const req = http.get(`${AUTH_SERVER_URL}/`, () => {
      resolve(true);
    });
    req.on("error", () => {
      resolve(false);
    });
    req.setTimeout(2000, () => {
      req.destroy();
      resolve(false);
    });
  });
}

async function autoStartAuthServer(): Promise<{ success: boolean; message: string }> {
  return new Promise((resolve) => {
    try {
      const authServerPath = path.join(__dirname, "..", "auth-server.js");

      const child = spawn("node", [authServerPath], {
        detached: true,
        stdio: "ignore",
        env: { ...process.env },
      });

      child.unref();

      // Wait for server to start
      setTimeout(async () => {
        const running = await isAuthServerRunning();
        if (running) {
          resolve({ success: true, message: "Auth server started automatically" });
        } else {
          resolve({
            success: false,
            message: "Failed to start auth server automatically. Please run: pnpm auth-server",
          });
        }
      }, 2000);
    } catch (error) {
      resolve({
        success: false,
        message: `Failed to start auth server: ${error instanceof Error ? error.message : "Unknown error"}`,
      });
    }
  });
}

async function openBrowser(url: string): Promise<{ success: boolean; message: string }> {
  return new Promise((resolve) => {
    const platform = process.platform;
    let command: string;

    if (platform === "darwin") {
      command = `open "${url}"`;
    } else if (platform === "win32") {
      command = `start "" "${url}"`;
    } else {
      command = `xdg-open "${url}"`;
    }

    exec(command, (error) => {
      if (error) {
        resolve({ success: false, message: `Could not open browser automatically: ${error.message}` });
      } else {
        resolve({ success: true, message: "Browser opened" });
      }
    });
  });
}

function formatTokenExpiration(expiresAt: number | undefined): string {
  if (!expiresAt) return "Unknown";

  const now = Date.now();
  const expiresIn = expiresAt - now;

  if (expiresIn <= 0) {
    return "Expired";
  }

  const minutes = Math.floor(expiresIn / 60000);
  const hours = Math.floor(minutes / 60);

  if (hours > 0) {
    return `${hours} hour${hours > 1 ? "s" : ""} ${minutes % 60} minute${minutes % 60 !== 1 ? "s" : ""}`;
  }
  return `${minutes} minute${minutes !== 1 ? "s" : ""}`;
}

// =============================================================================
// Handlers
// =============================================================================

export async function handleAbout(): Promise<{ content: Array<{ type: "text"; text: string }> }> {
  return {
    content: [
      {
        type: "text",
        text: `Outlook Assistant MCP Server v${SERVER_VERSION}

Provides access to Microsoft Outlook email, calendar, and contacts through Microsoft Graph API.
Built with TypeScript and the Model Context Protocol SDK.`,
      },
    ],
  };
}

export async function handleAuthenticate(
  args: z.infer<typeof authenticateSchema>
): Promise<{ content: Array<{ type: "text"; text: string }> }> {
  const force = args?.force === true;

  // For test mode, create a test token
  if (USE_TEST_MODE) {
    createTestTokens();
    return {
      content: [
        {
          type: "text",
          text: "Successfully authenticated with Microsoft Graph API (test mode)",
        },
      ],
    };
  }

  // Check if auth server is running
  let serverRunning = await isAuthServerRunning();
  let serverMessage = "";

  // If not running, try to auto-start it
  if (!serverRunning) {
    const startResult = await autoStartAuthServer();
    serverRunning = startResult.success;
    serverMessage = startResult.message;
  }

  // If still not running, return error with helpful instructions
  if (!serverRunning) {
    return {
      content: [
        {
          type: "text",
          text: `Authentication server is not running.

To start the auth server manually, run:
  pnpm auth-server

If port 1337 is in use, run:
  npx kill-port 1337

Then try authenticating again.

${serverMessage ? `Details: ${serverMessage}` : ""}`,
        },
      ],
    };
  }

  // Generate auth URL
  const authUrl = `${AUTH_SERVER_URL}/auth?client_id=${CLIENT_ID}`;

  // Try to open browser automatically
  const browserResult = await openBrowser(authUrl);

  // Build response message
  let responseText = "";

  if (serverMessage) {
    responseText += `${serverMessage}\n\n`;
  }

  if (browserResult.success) {
    responseText += `Opening authentication page in your browser...

Steps:
1. Sign in with your Microsoft account in the browser window
2. Grant the requested permissions
3. You'll be redirected back automatically

If the browser didn't open, use this URL:
${authUrl}`;
  } else {
    responseText += `Please open this URL in your browser to authenticate:
${authUrl}

Steps:
1. Click or copy the URL above
2. Sign in with your Microsoft account
3. Grant the requested permissions
4. You'll be redirected back automatically

${browserResult.message}`;
  }

  return {
    content: [{ type: "text", text: responseText }],
  };
}

export async function handleCheckAuthStatus(): Promise<{ content: Array<{ type: "text"; text: string }> }> {
  const serverRunning = await isAuthServerRunning();
  const tokens = loadTokens();

  let statusText = "";

  // Server status
  statusText += `Auth Server: ${serverRunning ? "Running" : "Not running"}\n`;

  if (!tokens?.access_token) {
    statusText += `Authentication: Not authenticated\n`;
    statusText += `\nTo authenticate, use the "authenticate" tool.`;

    if (!serverRunning) {
      statusText += `\nNote: Auth server needs to be running for authentication.`;
    }

    return {
      content: [{ type: "text", text: statusText }],
    };
  }

  // Check if token is expired
  const now = Date.now();
  const isExpired = tokens.expires_at && tokens.expires_at < now;
  const expirationInfo = formatTokenExpiration(tokens.expires_at);

  statusText += `Authentication: ${isExpired ? "Expired" : "Authenticated"}\n`;
  statusText += `Token expires: ${isExpired ? "Expired" : `in ${expirationInfo}`}\n`;

  if (isExpired) {
    statusText += `\nYour token has expired. Use the "authenticate" tool to re-authenticate.`;
  } else {
    statusText += `\nReady to access Outlook data.`;
  }

  return {
    content: [{ type: "text", text: statusText }],
  };
}
