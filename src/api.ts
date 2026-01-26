/**
 * Microsoft Graph API client for Outlook MCP Server
 */

import https from "https";
import fs from "fs";
import path from "path";
import os from "os";
import type { TokenData, GraphAPIResponse } from "./types.js";

// =============================================================================
// Configuration
// =============================================================================

const GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0/";

const getTokenStorePath = (): string => {
  const homeDir = process.env.HOME || process.env.USERPROFILE || os.homedir() || "/tmp";
  return path.join(homeDir, ".outlook-mcp-keys.json");
};

// =============================================================================
// Token Management
// =============================================================================

let cachedTokens: TokenData | null = null;

export function loadTokens(): TokenData | null {
  try {
    const tokenPath = getTokenStorePath();

    if (!fs.existsSync(tokenPath)) {
      return null;
    }

    const tokenData = fs.readFileSync(tokenPath, "utf8");
    const tokens = JSON.parse(tokenData) as TokenData;

    if (!tokens.access_token) {
      return null;
    }

    // Check token expiration
    const now = Date.now();
    if (tokens.expires_at && now > tokens.expires_at) {
      return null;
    }

    cachedTokens = tokens;
    return tokens;
  } catch {
    return null;
  }
}

export function saveTokens(tokens: TokenData): boolean {
  try {
    const tokenPath = getTokenStorePath();
    fs.writeFileSync(tokenPath, JSON.stringify(tokens, null, 2));
    cachedTokens = tokens;
    return true;
  } catch {
    return false;
  }
}

export function getAccessToken(): string | null {
  if (cachedTokens?.access_token) {
    // Check if token is still valid
    if (cachedTokens.expires_at && Date.now() > cachedTokens.expires_at) {
      cachedTokens = null;
      return null;
    }
    return cachedTokens.access_token;
  }

  const tokens = loadTokens();
  return tokens?.access_token ?? null;
}

export function createTestTokens(): TokenData {
  const testTokens: TokenData = {
    access_token: "test_access_token_" + Date.now(),
    refresh_token: "test_refresh_token_" + Date.now(),
    expires_at: Date.now() + 3600 * 1000, // 1 hour
  };

  saveTokens(testTokens);
  return testTokens;
}

export function ensureAuthenticated(): string {
  const accessToken = getAccessToken();
  if (!accessToken) {
    throw new Error("Authentication required");
  }
  return accessToken;
}

// =============================================================================
// Graph API Client
// =============================================================================

export interface QueryParams {
  [key: string]: string | number | boolean | undefined;
  $filter?: string;
  $select?: string;
  $top?: number;
  $orderby?: string;
  $search?: string;
  $count?: string;
  startDateTime?: string;
  endDateTime?: string;
}

/**
 * Makes a request to the Microsoft Graph API
 */
export async function callGraphAPI<T = unknown>(
  accessToken: string,
  method: "GET" | "POST" | "PATCH" | "PUT" | "DELETE",
  apiPath: string,
  data: unknown = null,
  queryParams: QueryParams = {},
  extraHeaders: Record<string, string> = {}
): Promise<T> {
  // Build the URL
  let finalUrl: string;

  if (apiPath.startsWith("http://") || apiPath.startsWith("https://")) {
    // Path is already a full URL (from pagination nextLink)
    finalUrl = apiPath;
  } else {
    // Encode path segments properly
    const encodedPath = apiPath
      .split("/")
      .map((segment) => encodeURIComponent(segment))
      .join("/");

    // Build query string from parameters with special handling for OData filters
    let queryString = "";
    if (Object.keys(queryParams).length > 0) {
      // Handle $filter parameter specially to ensure proper URI encoding
      const filter = queryParams.$filter;
      const paramsWithoutFilter = { ...queryParams };
      if (filter !== undefined) {
        delete paramsWithoutFilter.$filter;
      }

      // Build query string with proper encoding for regular params
      const params = new URLSearchParams();
      for (const [key, value] of Object.entries(paramsWithoutFilter)) {
        if (value !== undefined) {
          params.append(key, String(value));
        }
      }

      queryString = params.toString();

      // Add filter parameter separately with proper encoding
      if (filter) {
        if (queryString) {
          queryString += `&$filter=${encodeURIComponent(filter)}`;
        } else {
          queryString = `$filter=${encodeURIComponent(filter)}`;
        }
      }

      if (queryString) {
        queryString = "?" + queryString;
      }
    }

    finalUrl = `${GRAPH_API_ENDPOINT}${encodedPath}${queryString}`;
  }

  return new Promise((resolve, reject) => {
    const options = {
      method: method,
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
        ...extraHeaders,
      },
    };

    const req = https.request(finalUrl, options, (res) => {
      let responseData = "";

      res.on("data", (chunk) => {
        responseData += chunk;
      });

      res.on("end", () => {
        if (res.statusCode && res.statusCode >= 200 && res.statusCode < 300) {
          try {
            responseData = responseData || "{}";
            const jsonResponse = JSON.parse(responseData);
            resolve(jsonResponse as T);
          } catch {
            reject(new Error(`Error parsing API response`));
          }
        } else if (res.statusCode === 401) {
          reject(new Error("UNAUTHORIZED"));
        } else {
          reject(new Error(`API call failed with status ${res.statusCode}: ${responseData}`));
        }
      });
    });

    req.on("error", (error) => {
      reject(new Error(`Network error during API call: ${error.message}`));
    });

    if (data && (method === "POST" || method === "PATCH" || method === "PUT")) {
      req.write(JSON.stringify(data));
    }

    req.end();
  });
}

/**
 * Calls Graph API with pagination support to retrieve all results up to maxCount
 */
export async function callGraphAPIPaginated<T>(
  accessToken: string,
  method: "GET",
  apiPath: string,
  queryParams: QueryParams = {},
  maxCount: number = 0,
  extraHeaders: Record<string, string> = {}
): Promise<GraphAPIResponse<T>> {
  const allItems: T[] = [];
  let nextLink: string | undefined;
  let currentUrl = apiPath;
  let currentParams = queryParams;

  try {
    do {
      const response = await callGraphAPI<GraphAPIResponse<T>>(
        accessToken,
        method,
        currentUrl,
        null,
        currentParams,
        extraHeaders
      );

      if (response.value && Array.isArray(response.value)) {
        allItems.push(...response.value);
      }

      // Check if we've reached the desired count
      if (maxCount > 0 && allItems.length >= maxCount) {
        break;
      }

      // Get next page URL
      nextLink = response["@odata.nextLink"];

      if (nextLink) {
        currentUrl = nextLink;
        currentParams = {}; // nextLink already contains all params
      }
    } while (nextLink);

    // Trim to exact count if needed
    const finalItems = maxCount > 0 ? allItems.slice(0, maxCount) : allItems;

    return {
      value: finalItems,
      "@odata.count": finalItems.length,
    };
  } catch (error) {
    throw error;
  }
}

// =============================================================================
// Convenience Methods
// =============================================================================

export async function graphGet<T = unknown>(
  apiPath: string,
  queryParams: QueryParams = {},
  extraHeaders: Record<string, string> = {}
): Promise<T> {
  const accessToken = ensureAuthenticated();
  return callGraphAPI<T>(accessToken, "GET", apiPath, null, queryParams, extraHeaders);
}

export async function graphPost<T = unknown>(apiPath: string, data: unknown): Promise<T> {
  const accessToken = ensureAuthenticated();
  return callGraphAPI<T>(accessToken, "POST", apiPath, data);
}

export async function graphPatch<T = unknown>(apiPath: string, data: unknown): Promise<T> {
  const accessToken = ensureAuthenticated();
  return callGraphAPI<T>(accessToken, "PATCH", apiPath, data);
}

export async function graphDelete(apiPath: string): Promise<void> {
  const accessToken = ensureAuthenticated();
  await callGraphAPI(accessToken, "DELETE", apiPath);
}

export async function graphGetPaginated<T>(
  apiPath: string,
  queryParams: QueryParams = {},
  maxCount: number = 0,
  extraHeaders: Record<string, string> = {}
): Promise<GraphAPIResponse<T>> {
  const accessToken = ensureAuthenticated();
  return callGraphAPIPaginated<T>(accessToken, "GET", apiPath, queryParams, maxCount, extraHeaders);
}
