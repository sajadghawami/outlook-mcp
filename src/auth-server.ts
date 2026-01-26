#!/usr/bin/env node
/**
 * Outlook OAuth Authentication Server
 *
 * Handles the OAuth 2.0 callback flow for Microsoft Graph API authentication.
 */

import http from "http";
import https from "https";
import url from "url";
import querystring from "querystring";
import fs from "fs";
import path from "path";
import os from "os";

// =============================================================================
// Configuration
// =============================================================================

const PORT = 1337;
const REDIRECT_URI = `http://localhost:${PORT}/auth/callback`;

const AUTH_CONFIG = {
  clientId: process.env.AZURE_CLIENT_ID || "",
  clientSecret: process.env.AZURE_CLIENT_SECRET || "",
  redirectUri: REDIRECT_URI,
  scopes: [
    "offline_access",
    "User.Read",
    "Mail.Read",
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.Read",
    "Calendars.ReadWrite",
    "MailboxSettings.Read",
  ],
  tokenStorePath: path.join(
    process.env.HOME || process.env.USERPROFILE || os.homedir() || "/tmp",
    ".outlook-mcp-keys.json"
  ),
};

// =============================================================================
// Token Exchange
// =============================================================================

interface TokenResponse {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
  expires_at?: number;
  token_type?: string;
  scope?: string;
}

function exchangeCodeForTokens(code: string): Promise<TokenResponse> {
  return new Promise((resolve, reject) => {
    const postData = querystring.stringify({
      client_id: AUTH_CONFIG.clientId,
      client_secret: AUTH_CONFIG.clientSecret,
      code: code,
      redirect_uri: AUTH_CONFIG.redirectUri,
      grant_type: "authorization_code",
      scope: AUTH_CONFIG.scopes.join(" "),
    });

    const options = {
      hostname: "login.microsoftonline.com",
      path: "/common/oauth2/v2.0/token",
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Content-Length": Buffer.byteLength(postData),
      },
    };

    const req = https.request(options, (res) => {
      let data = "";

      res.on("data", (chunk) => {
        data += chunk;
      });

      res.on("end", () => {
        if (res.statusCode && res.statusCode >= 200 && res.statusCode < 300) {
          try {
            const tokenResponse: TokenResponse = JSON.parse(data);

            // Calculate expiration time
            const expiresAt = Date.now() + tokenResponse.expires_in * 1000;
            tokenResponse.expires_at = expiresAt;

            // Save tokens to file
            fs.writeFileSync(AUTH_CONFIG.tokenStorePath, JSON.stringify(tokenResponse, null, 2), "utf8");
            console.log(`Tokens saved to ${AUTH_CONFIG.tokenStorePath}`);

            resolve(tokenResponse);
          } catch (error) {
            reject(new Error(`Error parsing token response: ${error instanceof Error ? error.message : "Unknown error"}`));
          }
        } else {
          reject(new Error(`Token exchange failed with status ${res.statusCode}: ${data}`));
        }
      });
    });

    req.on("error", (error) => {
      reject(error);
    });

    req.write(postData);
    req.end();
  });
}

// =============================================================================
// HTML Templates
// =============================================================================

function errorPage(title: string, message: string, description?: string): string {
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>${title}</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      background: #fafafa;
      color: #333;
    }
    .container { text-align: center; padding: 2rem; }
    .icon { font-size: 3rem; color: #b33; margin-bottom: 1.5rem; }
    h1 { font-size: 1.5rem; font-weight: 500; margin-bottom: 0.75rem; }
    .message { color: #666; font-size: 0.95rem; margin-bottom: 0.5rem; }
    .hint { color: #999; font-size: 0.85rem; margin-top: 1.5rem; }
  </style>
</head>
<body>
  <div class="container">
    <div class="icon">✗</div>
    <h1>${title}</h1>
    <p class="message">${message}</p>
    ${description ? `<p class="message">${description}</p>` : ""}
    <p class="hint">Close and try again</p>
  </div>
</body>
</html>
`;
}

function successPage(): string {
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Connected</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      background: #fafafa;
      color: #333;
    }
    .container { text-align: center; padding: 2rem; }
    .icon { font-size: 3rem; color: #2a2; margin-bottom: 1.5rem; }
    h1 { font-size: 1.5rem; font-weight: 500; margin-bottom: 0.75rem; }
    .hint { color: #999; font-size: 0.85rem; }
  </style>
</head>
<body>
  <div class="container">
    <div class="icon">✓</div>
    <h1>Connected</h1>
    <p class="hint">You can close this window</p>
  </div>
</body>
</html>
`;
}

function homePage(): string {
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Outlook Auth Server</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      background: #fafafa;
      color: #333;
    }
    .container { text-align: center; padding: 2rem; }
    h1 { font-size: 1.5rem; font-weight: 500; margin-bottom: 0.75rem; }
    .status { color: #999; font-size: 0.85rem; }
  </style>
</head>
<body>
  <div class="container">
    <h1>Outlook Auth Server</h1>
    <p class="status">Waiting for authentication...</p>
  </div>
</body>
</html>
`;
}

// =============================================================================
// Server
// =============================================================================

const server = http.createServer((req, res) => {
  const parsedUrl = url.parse(req.url || "", true);
  const pathname = parsedUrl.pathname;

  console.log(`Request received: ${pathname}`);

  if (pathname === "/auth/callback") {
    const query = parsedUrl.query;

    if (query.error) {
      console.error(`Authentication error: ${query.error} - ${query.error_description}`);
      res.writeHead(400, { "Content-Type": "text/html" });
      res.end(errorPage("Authentication Error", String(query.error), String(query.error_description || "")));
      return;
    }

    if (query.code) {
      console.log("Authorization code received, exchanging for tokens...");

      exchangeCodeForTokens(String(query.code))
        .then(() => {
          console.log("Token exchange successful");
          res.writeHead(200, { "Content-Type": "text/html" });
          res.end(successPage());

          // Auto-shutdown after successful authentication
          console.log("Authentication complete. Shutting down auth server in 2 seconds...");
          setTimeout(() => {
            console.log("Auth server shutting down after successful authentication");
            server.close(() => {
              process.exit(0);
            });
          }, 2000);
        })
        .catch((error) => {
          console.error(`Token exchange error: ${error.message}`);
          res.writeHead(500, { "Content-Type": "text/html" });
          res.end(errorPage("Token Exchange Error", error.message));
        });
    } else {
      console.error("No authorization code provided");
      res.writeHead(400, { "Content-Type": "text/html" });
      res.end(errorPage("Missing Authorization Code", "No authorization code was provided in the callback."));
    }
  } else if (pathname === "/auth") {
    console.log("Auth request received, redirecting to Microsoft login...");

    if (!AUTH_CONFIG.clientId || !AUTH_CONFIG.clientSecret) {
      res.writeHead(500, { "Content-Type": "text/html" });
      res.end(
        errorPage(
          "Configuration Error",
          "Microsoft Graph API credentials are not set. Please set AZURE_CLIENT_ID and AZURE_CLIENT_SECRET environment variables."
        )
      );
      return;
    }

    const query = parsedUrl.query;
    const clientId = String(query.client_id || AUTH_CONFIG.clientId);

    const authParams = {
      client_id: clientId,
      response_type: "code",
      redirect_uri: AUTH_CONFIG.redirectUri,
      scope: AUTH_CONFIG.scopes.join(" "),
      response_mode: "query",
      state: Date.now().toString(),
    };

    const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?${querystring.stringify(authParams)}`;
    console.log(`Redirecting to: ${authUrl}`);

    res.writeHead(302, { Location: authUrl });
    res.end();
  } else if (pathname === "/") {
    res.writeHead(200, { "Content-Type": "text/html" });
    res.end(homePage());
  } else {
    res.writeHead(404, { "Content-Type": "text/plain" });
    res.end("Not Found");
  }
});

// =============================================================================
// Startup
// =============================================================================

server.listen(PORT, () => {
  console.log(`Authentication server running at http://localhost:${PORT}`);
  console.log(`Waiting for authentication callback at ${AUTH_CONFIG.redirectUri}`);
  console.log(`Token will be stored at: ${AUTH_CONFIG.tokenStorePath}`);

  if (!AUTH_CONFIG.clientId || !AUTH_CONFIG.clientSecret) {
    console.log("\nWARNING: Microsoft Graph API credentials are not set.");
    console.log("Please set the AZURE_CLIENT_ID and AZURE_CLIENT_SECRET environment variables.");
  }
});

// Handle termination
process.on("SIGINT", () => {
  console.log("Authentication server shutting down");
  process.exit(0);
});

process.on("SIGTERM", () => {
  console.log("Authentication server shutting down");
  process.exit(0);
});
