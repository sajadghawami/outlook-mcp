/**
 * Email tools for Outlook MCP Server
 */

import { z } from "zod";
import {
  graphGet,
  graphPost,
  graphPatch,
  graphGetPaginated,
  ensureAuthenticated,
  callGraphAPI,
  callGraphAPIPaginated,
  type QueryParams,
} from "../api.js";
import type { Email, MailFolder, GraphAPIResponse, SendEmailPayload, SendEmailRecipient } from "../types.js";
import { WELL_KNOWN_FOLDERS } from "../types.js";

// =============================================================================
// Configuration
// =============================================================================

const EMAIL_SELECT_FIELDS =
  "id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,hasAttachments,importance,isRead";
const EMAIL_DETAIL_FIELDS =
  "id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,bodyPreview,body,hasAttachments,importance,isRead,internetMessageHeaders";
const MAX_RESULT_COUNT = 50;

// =============================================================================
// Schemas
// =============================================================================

export const listEmailsSchema = z.object({
  folder: z
    .string()
    .optional()
    .describe("Email folder to list (e.g., 'inbox', 'sent', 'drafts', default: 'inbox')"),
  count: z.number().optional().describe("Number of emails to retrieve (default: 10, max: 50)"),
});

export const searchEmailsSchema = z.object({
  query: z.string().optional().describe("Search query text to find in emails"),
  folder: z.string().optional().describe("Email folder to search in (default: 'inbox')"),
  from: z.string().optional().describe("Filter by sender email address or name"),
  to: z.string().optional().describe("Filter by recipient email address or name"),
  subject: z.string().optional().describe("Filter by email subject"),
  hasAttachments: z.boolean().optional().describe("Filter to only emails with attachments"),
  unreadOnly: z.boolean().optional().describe("Filter to only unread emails"),
  count: z.number().optional().describe("Number of results to return (default: 10, max: 50)"),
});

export const readEmailSchema = z.object({
  id: z.string().describe("ID of the email to read"),
});

export const sendEmailSchema = z.object({
  to: z.string().describe("Comma-separated list of recipient email addresses"),
  cc: z.string().optional().describe("Comma-separated list of CC recipient email addresses"),
  bcc: z.string().optional().describe("Comma-separated list of BCC recipient email addresses"),
  subject: z.string().describe("Email subject"),
  body: z.string().describe("Email body content (can be plain text or HTML)"),
  importance: z
    .enum(["normal", "high", "low"])
    .optional()
    .describe("Email importance (normal, high, low)"),
  saveToSentItems: z.boolean().optional().describe("Whether to save the email to sent items"),
});

export const markAsReadSchema = z.object({
  id: z.string().describe("ID of the email to mark as read/unread"),
  isRead: z
    .boolean()
    .optional()
    .describe("Whether to mark as read (true) or unread (false). Default: true"),
});

// =============================================================================
// Helper Functions
// =============================================================================

async function getAllFolders(accessToken: string): Promise<MailFolder[]> {
  try {
    const response = await callGraphAPI<GraphAPIResponse<MailFolder>>(
      accessToken,
      "GET",
      "me/mailFolders",
      null,
      {
        $top: 100,
        $select: "id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount",
      }
    );

    if (!response.value) {
      return [];
    }

    const foldersWithChildren = response.value.filter((f) => f.childFolderCount > 0);

    const childFolderPromises = foldersWithChildren.map(async (folder) => {
      try {
        const childResponse = await callGraphAPI<GraphAPIResponse<MailFolder>>(
          accessToken,
          "GET",
          `me/mailFolders/${folder.id}/childFolders`,
          null,
          {
            $select: "id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount",
          }
        );
        return childResponse.value || [];
      } catch {
        return [];
      }
    });

    const childFolders = await Promise.all(childFolderPromises);
    return [...response.value, ...childFolders.flat()];
  } catch {
    return [];
  }
}

async function getFolderIdByName(accessToken: string, folderName: string): Promise<string | null> {
  // Try direct match in top-level folders
  try {
    const response = await callGraphAPI<GraphAPIResponse<MailFolder>>(
      accessToken,
      "GET",
      "me/mailFolders",
      null,
      { $filter: `displayName eq '${folderName}'` }
    );

    if (response.value && response.value.length > 0) {
      return response.value[0].id;
    }

    // Try case-insensitive match
    const allFoldersResponse = await callGraphAPI<GraphAPIResponse<MailFolder>>(
      accessToken,
      "GET",
      "me/mailFolders",
      null,
      { $top: 100 }
    );

    if (allFoldersResponse.value) {
      const lowerFolderName = folderName.toLowerCase();
      const matchingFolder = allFoldersResponse.value.find(
        (folder) => folder.displayName.toLowerCase() === lowerFolderName
      );

      if (matchingFolder) {
        return matchingFolder.id;
      }
    }
  } catch {
    // Fall through to subfolder search
  }

  // Search recursively through ALL folders including subfolders
  try {
    const allFolders = await getAllFolders(accessToken);
    const lowerName = folderName.toLowerCase();
    const match = allFolders.find((f) => f.displayName.toLowerCase() === lowerName);

    if (match) {
      return match.id;
    }
  } catch {
    // No match found
  }

  return null;
}

async function resolveFolderPath(accessToken: string, folderName: string): Promise<string> {
  if (!folderName) {
    return WELL_KNOWN_FOLDERS["inbox"];
  }

  const lowerFolderName = folderName.toLowerCase();
  if (WELL_KNOWN_FOLDERS[lowerFolderName]) {
    return WELL_KNOWN_FOLDERS[lowerFolderName];
  }

  const folderId = await getFolderIdByName(accessToken, folderName);
  if (folderId) {
    return `me/mailFolders/${folderId}/messages`;
  }

  throw new Error(`Folder "${folderName}" not found. Use list-folders to see available folders.`);
}

function addBooleanFilters(
  params: QueryParams,
  filterTerms: { hasAttachments?: boolean; unreadOnly?: boolean }
): void {
  const filterConditions: string[] = [];

  if (filterTerms.hasAttachments === true) {
    filterConditions.push("hasAttachments eq true");
  }

  if (filterTerms.unreadOnly === true) {
    filterConditions.push("isRead eq false");
  }

  if (filterConditions.length > 0) {
    params.$filter = filterConditions.join(" and ");
  }
}

function addBooleanFiltersToFilter(
  params: QueryParams,
  filterTerms: { hasAttachments?: boolean; unreadOnly?: boolean }
): void {
  const additionalConditions: string[] = [];

  if (filterTerms.hasAttachments === true) {
    additionalConditions.push("hasAttachments eq true");
  }

  if (filterTerms.unreadOnly === true) {
    additionalConditions.push("isRead eq false");
  }

  if (additionalConditions.length > 0) {
    if (params.$filter) {
      params.$filter = `(${params.$filter}) and ${additionalConditions.join(" and ")}`;
    } else {
      params.$filter = additionalConditions.join(" and ");
    }
  }
}

interface SearchTerms {
  query?: string;
  from?: string;
  to?: string;
  subject?: string;
}

interface FilterTerms {
  hasAttachments?: boolean;
  unreadOnly?: boolean;
}

async function progressiveSearch(
  endpoint: string,
  accessToken: string,
  searchTerms: SearchTerms,
  filterTerms: FilterTerms,
  maxCount: number
): Promise<GraphAPIResponse<Email> & { _searchInfo?: { strategies: string[]; failed?: boolean } }> {
  const searchAttempts: string[] = [];

  // Strategy 1: $filter for sender address with exact match
  if (searchTerms.from && searchTerms.from.includes("@")) {
    const escapedFrom = searchTerms.from.replace(/'/g, "''").toLowerCase();

    try {
      searchAttempts.push("filter-from-exact");

      const filterParams: QueryParams = {
        $top: Math.min(50, maxCount),
        $select: EMAIL_SELECT_FIELDS,
        $filter: `sender/emailAddress/address eq '${escapedFrom}'`,
        $count: "true",
      };

      addBooleanFiltersToFilter(filterParams, filterTerms);

      const response = await callGraphAPIPaginated<Email>(
        accessToken,
        "GET",
        "me/messages",
        filterParams,
        maxCount,
        { ConsistencyLevel: "eventual" }
      );

      if (response.value && response.value.length > 0) {
        return { ...response, _searchInfo: { strategies: searchAttempts } };
      }
    } catch {
      // Try next strategy
    }
  }

  // Strategy 2: $filter for subject (using startswith)
  if (searchTerms.subject) {
    try {
      const escapedSubject = searchTerms.subject.replace(/'/g, "''");
      searchAttempts.push("filter-subject-startswith");

      const filterParams: QueryParams = {
        $top: Math.min(50, maxCount),
        $select: EMAIL_SELECT_FIELDS,
        $orderby: "receivedDateTime desc",
        $filter: `startswith(subject,'${escapedSubject}')`,
      };

      addBooleanFiltersToFilter(filterParams, filterTerms);

      const response = await callGraphAPIPaginated<Email>(accessToken, "GET", endpoint, filterParams, maxCount);

      if (response.value && response.value.length > 0) {
        return { ...response, _searchInfo: { strategies: searchAttempts } };
      }
    } catch {
      // Try next strategy
    }
  }

  // Strategy 3: Try KQL $search combined
  try {
    const kqlTerms: string[] = [];

    if (searchTerms.query) kqlTerms.push(searchTerms.query);
    if (searchTerms.subject) kqlTerms.push(`subject:"${searchTerms.subject}"`);
    if (searchTerms.from) kqlTerms.push(`from:"${searchTerms.from}"`);
    if (searchTerms.to) kqlTerms.push(`to:"${searchTerms.to}"`);

    if (kqlTerms.length > 0) {
      searchAttempts.push("kql-combined-search");

      const params: QueryParams = {
        $top: Math.min(50, maxCount),
        $select: EMAIL_SELECT_FIELDS,
        $orderby: "receivedDateTime desc",
        $search: kqlTerms.join(" "),
      };

      addBooleanFilters(params, filterTerms);

      const response = await callGraphAPIPaginated<Email>(accessToken, "GET", endpoint, params, maxCount);

      if (response.value && response.value.length > 0) {
        return { ...response, _searchInfo: { strategies: searchAttempts } };
      }
    }
  } catch {
    // Try next strategy
  }

  // Strategy 4: Try each KQL search term individually
  const searchPriority: (keyof SearchTerms)[] = ["subject", "from", "to", "query"];

  for (const term of searchPriority) {
    if (searchTerms[term]) {
      try {
        searchAttempts.push(`kql-single-term-${term}`);

        const simplifiedParams: QueryParams = {
          $top: Math.min(50, maxCount),
          $select: EMAIL_SELECT_FIELDS,
          $orderby: "receivedDateTime desc",
        };

        if (term === "query") {
          simplifiedParams.$search = `"${searchTerms[term]}"`;
        } else {
          simplifiedParams.$search = `${term}:"${searchTerms[term]}"`;
        }

        addBooleanFilters(simplifiedParams, filterTerms);

        const response = await callGraphAPIPaginated<Email>(
          accessToken,
          "GET",
          endpoint,
          simplifiedParams,
          maxCount
        );

        if (response.value && response.value.length > 0) {
          return { ...response, _searchInfo: { strategies: searchAttempts } };
        }
      } catch {
        // Try next strategy
      }
    }
  }

  // Strategy 5: Try with only boolean filters
  if (filterTerms.hasAttachments === true || filterTerms.unreadOnly === true) {
    try {
      searchAttempts.push("boolean-filters-only");

      const filterOnlyParams: QueryParams = {
        $top: Math.min(50, maxCount),
        $select: EMAIL_SELECT_FIELDS,
        $orderby: "receivedDateTime desc",
      };

      addBooleanFilters(filterOnlyParams, filterTerms);

      const response = await callGraphAPIPaginated<Email>(
        accessToken,
        "GET",
        endpoint,
        filterOnlyParams,
        maxCount
      );

      return { ...response, _searchInfo: { strategies: searchAttempts } };
    } catch {
      // Fall through to empty results
    }
  }

  // All strategies failed
  return {
    value: [],
    _searchInfo: {
      strategies: searchAttempts,
      failed: true,
    },
  };
}

function formatSearchResults(
  response: GraphAPIResponse<Email> & { _searchInfo?: { strategies: string[]; failed?: boolean } }
): { content: Array<{ type: "text"; text: string }> } {
  if (!response.value || response.value.length === 0) {
    const searchFailed = response._searchInfo?.failed;
    const message = searchFailed
      ? "No emails found matching your search criteria. The search may have failed - try a simpler query or use list-emails to browse the folder."
      : "No emails found matching your search criteria.";

    return {
      content: [{ type: "text", text: message }],
    };
  }

  const emailList = response.value
    .map((email, index) => {
      const sender = email.from?.emailAddress || { name: "Unknown", address: "unknown" };
      const date = new Date(email.receivedDateTime).toLocaleString();
      const readStatus = email.isRead ? "" : "[UNREAD] ";

      return `${index + 1}. ${readStatus}${date} - From: ${sender.name} (${sender.address})
Subject: ${email.subject}
ID: ${email.id}`;
    })
    .join("\n\n");

  let additionalInfo = "";
  if (response._searchInfo) {
    const lastStrategy = response._searchInfo.strategies[response._searchInfo.strategies.length - 1];
    additionalInfo = `\n(Search used ${lastStrategy} strategy)`;
  }

  return {
    content: [
      {
        type: "text",
        text: `Found ${response.value.length} emails matching your search criteria:${additionalInfo}\n\n${emailList}`,
      },
    ],
  };
}

// =============================================================================
// Handlers
// =============================================================================

export async function handleListEmails(
  args: z.infer<typeof listEmailsSchema>
): Promise<{ content: Array<{ type: "text"; text: string }> }> {
  try {
    const accessToken = ensureAuthenticated();
    const folder = args.folder || "inbox";
    const requestedCount = args.count || 10;

    const endpoint = await resolveFolderPath(accessToken, folder);

    const response = await callGraphAPIPaginated<Email>(
      accessToken,
      "GET",
      endpoint,
      {
        $top: Math.min(50, requestedCount),
        $orderby: "receivedDateTime desc",
        $select: EMAIL_SELECT_FIELDS,
      },
      requestedCount
    );

    if (!response.value || response.value.length === 0) {
      return {
        content: [{ type: "text", text: `No emails found in ${folder}.` }],
      };
    }

    const emailList = response.value
      .map((email, index) => {
        const sender = email.from?.emailAddress || { name: "Unknown", address: "unknown" };
        const date = new Date(email.receivedDateTime).toLocaleString();
        const readStatus = email.isRead ? "" : "[UNREAD] ";

        return `${index + 1}. ${readStatus}${date} - From: ${sender.name} (${sender.address})
Subject: ${email.subject}
ID: ${email.id}`;
      })
      .join("\n\n");

    return {
      content: [{ type: "text", text: `Found ${response.value.length} emails in ${folder}:\n\n${emailList}` }],
    };
  } catch (error) {
    if (error instanceof Error && error.message === "Authentication required") {
      return {
        content: [
          { type: "text", text: "Authentication required. Please use the 'authenticate' tool first." },
        ],
      };
    }
    if (error instanceof Error && error.message.includes("not found")) {
      return {
        content: [{ type: "text", text: error.message }],
      };
    }
    return {
      content: [
        { type: "text", text: `Error listing emails: ${error instanceof Error ? error.message : "Unknown error"}` },
      ],
    };
  }
}

export async function handleSearchEmails(
  args: z.infer<typeof searchEmailsSchema>
): Promise<{ content: Array<{ type: "text"; text: string }> }> {
  try {
    const accessToken = ensureAuthenticated();
    const folder = args.folder || "inbox";
    const requestedCount = args.count || 10;

    const endpoint = await resolveFolderPath(accessToken, folder);

    const response = await progressiveSearch(
      endpoint,
      accessToken,
      { query: args.query, from: args.from, to: args.to, subject: args.subject },
      { hasAttachments: args.hasAttachments, unreadOnly: args.unreadOnly },
      requestedCount
    );

    return formatSearchResults(response);
  } catch (error) {
    if (error instanceof Error && error.message === "Authentication required") {
      return {
        content: [
          { type: "text", text: "Authentication required. Please use the 'authenticate' tool first." },
        ],
      };
    }
    if (error instanceof Error && error.message.includes("not found")) {
      return {
        content: [{ type: "text", text: error.message }],
      };
    }
    return {
      content: [
        { type: "text", text: `Error searching emails: ${error instanceof Error ? error.message : "Unknown error"}` },
      ],
    };
  }
}

export async function handleReadEmail(args: z.infer<typeof readEmailSchema>): Promise<{ content: Array<{ type: "text"; text: string }> }> {
  try {
    if (!args.id) {
      return {
        content: [{ type: "text", text: "Email ID is required." }],
      };
    }

    const email = await graphGet<Email>(`me/messages/${encodeURIComponent(args.id)}`, {
      $select: EMAIL_DETAIL_FIELDS,
    });

    if (!email) {
      return {
        content: [{ type: "text", text: `Email with ID ${args.id} not found.` }],
      };
    }

    const sender = email.from
      ? `${email.from.emailAddress.name} (${email.from.emailAddress.address})`
      : "Unknown";
    const to = email.toRecipients
      ? email.toRecipients.map((r) => `${r.emailAddress.name} (${r.emailAddress.address})`).join(", ")
      : "None";
    const cc =
      email.ccRecipients && email.ccRecipients.length > 0
        ? email.ccRecipients.map((r) => `${r.emailAddress.name} (${r.emailAddress.address})`).join(", ")
        : "None";
    const bcc =
      email.bccRecipients && email.bccRecipients.length > 0
        ? email.bccRecipients.map((r) => `${r.emailAddress.name} (${r.emailAddress.address})`).join(", ")
        : "None";
    const date = new Date(email.receivedDateTime).toLocaleString();

    let body = "";
    if (email.body) {
      body = email.body.contentType === "html" ? email.body.content.replace(/<[^>]*>/g, "") : email.body.content;
    } else {
      body = email.bodyPreview || "No content";
    }

    const formattedEmail = `From: ${sender}
To: ${to}
${cc !== "None" ? `CC: ${cc}\n` : ""}${bcc !== "None" ? `BCC: ${bcc}\n` : ""}Subject: ${email.subject}
Date: ${date}
Importance: ${email.importance || "normal"}
Has Attachments: ${email.hasAttachments ? "Yes" : "No"}

${body}`;

    return {
      content: [{ type: "text", text: formattedEmail }],
    };
  } catch (error) {
    if (error instanceof Error && error.message === "Authentication required") {
      return {
        content: [
          { type: "text", text: "Authentication required. Please use the 'authenticate' tool first." },
        ],
      };
    }
    if (error instanceof Error && error.message.includes("doesn't belong to the targeted mailbox")) {
      return {
        content: [
          {
            type: "text",
            text: "The email ID seems invalid or doesn't belong to your mailbox. Please try with a different email ID.",
          },
        ],
      };
    }
    return {
      content: [
        { type: "text", text: `Error reading email: ${error instanceof Error ? error.message : "Unknown error"}` },
      ],
    };
  }
}

export async function handleSendEmail(args: z.infer<typeof sendEmailSchema>): Promise<{ content: Array<{ type: "text"; text: string }> }> {
  try {
    if (!args.to) {
      return {
        content: [{ type: "text", text: "Recipient (to) is required." }],
      };
    }

    if (!args.subject) {
      return {
        content: [{ type: "text", text: "Subject is required." }],
      };
    }

    if (!args.body) {
      return {
        content: [{ type: "text", text: "Body content is required." }],
      };
    }

    const toRecipients: SendEmailRecipient[] = args.to.split(",").map((email) => ({
      emailAddress: { address: email.trim() },
    }));

    const ccRecipients: SendEmailRecipient[] = args.cc
      ? args.cc.split(",").map((email) => ({
          emailAddress: { address: email.trim() },
        }))
      : [];

    const bccRecipients: SendEmailRecipient[] = args.bcc
      ? args.bcc.split(",").map((email) => ({
          emailAddress: { address: email.trim() },
        }))
      : [];

    const emailPayload: SendEmailPayload = {
      message: {
        subject: args.subject,
        body: {
          contentType: args.body.includes("<html") ? "html" : "text",
          content: args.body,
        },
        toRecipients,
        ccRecipients: ccRecipients.length > 0 ? ccRecipients : undefined,
        bccRecipients: bccRecipients.length > 0 ? bccRecipients : undefined,
        importance: args.importance || "normal",
      },
      saveToSentItems: args.saveToSentItems !== false,
    };

    await graphPost("me/sendMail", emailPayload);

    return {
      content: [
        {
          type: "text",
          text: `Email sent successfully!

Subject: ${args.subject}
Recipients: ${toRecipients.length}${ccRecipients.length > 0 ? ` + ${ccRecipients.length} CC` : ""}${bccRecipients.length > 0 ? ` + ${bccRecipients.length} BCC` : ""}
Message Length: ${args.body.length} characters`,
        },
      ],
    };
  } catch (error) {
    if (error instanceof Error && error.message === "Authentication required") {
      return {
        content: [
          { type: "text", text: "Authentication required. Please use the 'authenticate' tool first." },
        ],
      };
    }
    return {
      content: [
        { type: "text", text: `Error sending email: ${error instanceof Error ? error.message : "Unknown error"}` },
      ],
    };
  }
}

export async function handleMarkAsRead(
  args: z.infer<typeof markAsReadSchema>
): Promise<{ content: Array<{ type: "text"; text: string }> }> {
  try {
    if (!args.id) {
      return {
        content: [{ type: "text", text: "Email ID is required." }],
      };
    }

    const isRead = args.isRead !== undefined ? args.isRead : true;

    await graphPatch(`me/messages/${encodeURIComponent(args.id)}`, { isRead });

    const status = isRead ? "read" : "unread";

    return {
      content: [{ type: "text", text: `Email successfully marked as ${status}.` }],
    };
  } catch (error) {
    if (error instanceof Error && error.message === "Authentication required") {
      return {
        content: [
          { type: "text", text: "Authentication required. Please use the 'authenticate' tool first." },
        ],
      };
    }
    if (error instanceof Error && error.message.includes("doesn't belong to the targeted mailbox")) {
      return {
        content: [
          {
            type: "text",
            text: "The email ID seems invalid or doesn't belong to your mailbox. Please try with a different email ID.",
          },
        ],
      };
    }
    return {
      content: [
        {
          type: "text",
          text: `Error marking email: ${error instanceof Error ? error.message : "Unknown error"}`,
        },
      ],
    };
  }
}

// Export helper functions for use in other modules
export { getFolderIdByName, getAllFolders };
