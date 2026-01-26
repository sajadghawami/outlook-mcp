/**
 * Folder management tools for Outlook MCP Server
 */

import { z } from "zod";
import {
  graphGet,
  graphPost,
  ensureAuthenticated,
  callGraphAPI,
} from "../api.js";
import type { MailFolder, GraphAPIResponse } from "../types.js";
import { getFolderIdByName, getAllFolders } from "./email.js";

// =============================================================================
// Schemas
// =============================================================================

export const listFoldersSchema = z.object({
  includeItemCounts: z.boolean().optional().describe("Include counts of total and unread items"),
  includeChildren: z.boolean().optional().describe("Include child folders in hierarchy"),
});

export const createFolderSchema = z.object({
  name: z.string().describe("Name of the folder to create"),
  parentFolder: z.string().optional().describe("Optional parent folder name (default is root)"),
});

export const moveEmailsSchema = z.object({
  emailIds: z.string().describe("Comma-separated list of email IDs to move"),
  targetFolder: z.string().describe("Name of the folder to move emails to"),
  sourceFolder: z.string().optional().describe("Optional name of the source folder (default is inbox)"),
});

// =============================================================================
// Helper Functions
// =============================================================================

async function getAllFoldersHierarchy(
  accessToken: string,
  includeItemCounts: boolean
): Promise<MailFolder[]> {
  try {
    const selectFields = includeItemCounts
      ? "id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount"
      : "id,displayName,parentFolderId,childFolderCount";

    const response = await callGraphAPI<GraphAPIResponse<MailFolder>>(
      accessToken,
      "GET",
      "me/mailFolders",
      null,
      {
        $top: 100,
        $select: selectFields,
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
          { $select: selectFields }
        );

        const childFolders = childResponse.value || [];
        childFolders.forEach((child) => {
          child.parentFolder = folder.displayName;
        });

        return childFolders;
      } catch {
        return [];
      }
    });

    const childFolders = await Promise.all(childFolderPromises);
    const allChildFolders = childFolders.flat();

    const topLevelFolders = response.value.map((folder) => ({
      ...folder,
      isTopLevel: true,
    }));

    return [...topLevelFolders, ...allChildFolders];
  } catch {
    throw new Error("Failed to retrieve folders");
  }
}

function formatFolderList(folders: MailFolder[], includeItemCounts: boolean): string {
  if (!folders || folders.length === 0) {
    return "No folders found.";
  }

  const wellKnownFolderNames = ["Inbox", "Drafts", "Sent Items", "Deleted Items", "Junk Email", "Archive"];

  const sortedFolders = [...folders].sort((a, b) => {
    const aIsWellKnown = wellKnownFolderNames.includes(a.displayName);
    const bIsWellKnown = wellKnownFolderNames.includes(b.displayName);

    if (aIsWellKnown && !bIsWellKnown) return -1;
    if (!aIsWellKnown && bIsWellKnown) return 1;

    if (aIsWellKnown && bIsWellKnown) {
      return wellKnownFolderNames.indexOf(a.displayName) - wellKnownFolderNames.indexOf(b.displayName);
    }

    return a.displayName.localeCompare(b.displayName);
  });

  const folderLines = sortedFolders.map((folder) => {
    let folderInfo = folder.displayName;

    if (folder.parentFolder) {
      folderInfo += ` (in ${folder.parentFolder})`;
    }

    if (includeItemCounts) {
      const unreadCount = folder.unreadItemCount || 0;
      const totalCount = folder.totalItemCount || 0;
      folderInfo += ` - ${totalCount} items`;

      if (unreadCount > 0) {
        folderInfo += ` (${unreadCount} unread)`;
      }
    }

    return folderInfo;
  });

  return `Found ${folders.length} folders:\n\n${folderLines.join("\n")}`;
}

function formatFolderHierarchy(folders: MailFolder[], includeItemCounts: boolean): string {
  if (!folders || folders.length === 0) {
    return "No folders found.";
  }

  interface FolderNode extends MailFolder {
    children: string[];
  }

  const folderMap = new Map<string, FolderNode>();
  const rootFolders: string[] = [];

  folders.forEach((folder) => {
    folderMap.set(folder.id, {
      ...folder,
      children: [],
    });

    if (folder.isTopLevel) {
      rootFolders.push(folder.id);
    }
  });

  folders.forEach((folder) => {
    if (!folder.isTopLevel && folder.parentFolderId) {
      const parent = folderMap.get(folder.parentFolderId);
      if (parent) {
        parent.children.push(folder.id);
      } else {
        rootFolders.push(folder.id);
      }
    }
  });

  function formatSubtree(folderId: string, level: number = 0): string {
    const folder = folderMap.get(folderId);
    if (!folder) return "";

    const indent = "  ".repeat(level);
    let line = `${indent}${folder.displayName}`;

    if (includeItemCounts) {
      const unreadCount = folder.unreadItemCount || 0;
      const totalCount = folder.totalItemCount || 0;
      line += ` - ${totalCount} items`;

      if (unreadCount > 0) {
        line += ` (${unreadCount} unread)`;
      }
    }

    const childLines = folder.children
      .map((childId) => formatSubtree(childId, level + 1))
      .filter((line) => line.length > 0)
      .join("\n");

    return childLines.length > 0 ? `${line}\n${childLines}` : line;
  }

  const formattedHierarchy = rootFolders.map((folderId) => formatSubtree(folderId)).join("\n");

  return `Folder Hierarchy:\n\n${formattedHierarchy}`;
}

// =============================================================================
// Handlers
// =============================================================================

export async function handleListFolders(
  args: z.infer<typeof listFoldersSchema>
): Promise<{ content: Array<{ type: "text"; text: string }> }> {
  try {
    const accessToken = ensureAuthenticated();
    const includeItemCounts = args.includeItemCounts === true;
    const includeChildren = args.includeChildren === true;

    const folders = await getAllFoldersHierarchy(accessToken, includeItemCounts);

    if (includeChildren) {
      return {
        content: [{ type: "text", text: formatFolderHierarchy(folders, includeItemCounts) }],
      };
    } else {
      return {
        content: [{ type: "text", text: formatFolderList(folders, includeItemCounts) }],
      };
    }
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
        { type: "text", text: `Error listing folders: ${error instanceof Error ? error.message : "Unknown error"}` },
      ],
    };
  }
}

export async function handleCreateFolder(
  args: z.infer<typeof createFolderSchema>
): Promise<{ content: Array<{ type: "text"; text: string }> }> {
  try {
    if (!args.name) {
      return {
        content: [{ type: "text", text: "Folder name is required." }],
      };
    }

    const accessToken = ensureAuthenticated();

    // Check if folder already exists
    const existingFolder = await getFolderIdByName(accessToken, args.name);
    if (existingFolder) {
      return {
        content: [{ type: "text", text: `A folder named "${args.name}" already exists.` }],
      };
    }

    // If parent folder specified, find its ID
    let endpoint = "me/mailFolders";
    if (args.parentFolder) {
      const parentId = await getFolderIdByName(accessToken, args.parentFolder);
      if (!parentId) {
        return {
          content: [
            {
              type: "text",
              text: `Parent folder "${args.parentFolder}" not found. Please specify a valid parent folder or leave it blank to create at the root level.`,
            },
          ],
        };
      }
      endpoint = `me/mailFolders/${parentId}/childFolders`;
    }

    const response = await callGraphAPI<MailFolder>(accessToken, "POST", endpoint, {
      displayName: args.name,
    });

    if (response && response.id) {
      const locationInfo = args.parentFolder ? `inside "${args.parentFolder}"` : "at the root level";

      return {
        content: [{ type: "text", text: `Successfully created folder "${args.name}" ${locationInfo}.` }],
      };
    } else {
      return {
        content: [{ type: "text", text: "Failed to create folder. The server didn't return a folder ID." }],
      };
    }
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
        { type: "text", text: `Error creating folder: ${error instanceof Error ? error.message : "Unknown error"}` },
      ],
    };
  }
}

export async function handleMoveEmails(
  args: z.infer<typeof moveEmailsSchema>
): Promise<{ content: Array<{ type: "text"; text: string }> }> {
  try {
    if (!args.emailIds) {
      return {
        content: [
          {
            type: "text",
            text: "Email IDs are required. Please provide a comma-separated list of email IDs to move.",
          },
        ],
      };
    }

    if (!args.targetFolder) {
      return {
        content: [{ type: "text", text: "Target folder name is required." }],
      };
    }

    const accessToken = ensureAuthenticated();

    const ids = args.emailIds
      .split(",")
      .map((id) => id.trim())
      .filter((id) => id);

    if (ids.length === 0) {
      return {
        content: [{ type: "text", text: "No valid email IDs provided." }],
      };
    }

    // Get the target folder ID
    const targetFolderId = await getFolderIdByName(accessToken, args.targetFolder);
    if (!targetFolderId) {
      return {
        content: [
          {
            type: "text",
            text: `Target folder "${args.targetFolder}" not found. Please specify a valid folder name.`,
          },
        ],
      };
    }

    // Track successful and failed moves
    const results = {
      successful: [] as string[],
      failed: [] as { id: string; error: string }[],
    };

    // Process each email
    for (const emailId of ids) {
      try {
        await callGraphAPI(accessToken, "POST", `me/messages/${emailId}/move`, {
          destinationId: targetFolderId,
        });
        results.successful.push(emailId);
      } catch (error) {
        results.failed.push({
          id: emailId,
          error: error instanceof Error ? error.message : "Unknown error",
        });
      }
    }

    // Generate result message
    let message = "";

    if (results.successful.length > 0) {
      message += `Successfully moved ${results.successful.length} email(s) to "${args.targetFolder}".`;
    }

    if (results.failed.length > 0) {
      if (message) message += "\n\n";
      message += `Failed to move ${results.failed.length} email(s). Errors:`;

      const maxErrors = Math.min(results.failed.length, 3);
      for (let i = 0; i < maxErrors; i++) {
        const failure = results.failed[i];
        message += `\n- Email ${i + 1}: ${failure.error}`;
      }

      if (results.failed.length > maxErrors) {
        message += `\n...and ${results.failed.length - maxErrors} more.`;
      }
    }

    return {
      content: [{ type: "text", text: message }],
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
        { type: "text", text: `Error moving emails: ${error instanceof Error ? error.message : "Unknown error"}` },
      ],
    };
  }
}
