#!/usr/bin/env node
/**
 * Outlook MCP Server - Main Entry Point
 *
 * A Model Context Protocol server that provides access to
 * Microsoft Outlook through the Microsoft Graph API.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";

// Auth tools
import {
  aboutSchema,
  authenticateSchema,
  checkAuthStatusSchema,
  handleAbout,
  handleAuthenticate,
  handleCheckAuthStatus,
} from "./tools/auth.js";

// Calendar tools
import {
  listEventsSchema,
  declineEventSchema,
  createEventSchema,
  updateEventSchema,
  cancelEventSchema,
  deleteEventSchema,
  listCategoriesSchema,
  createCategorySchema,
  deleteCategorySchema,
  handleListEvents,
  handleDeclineEvent,
  handleCreateEvent,
  handleUpdateEvent,
  handleCancelEvent,
  handleDeleteEvent,
  handleListCategories,
  handleCreateCategory,
  handleDeleteCategory,
} from "./tools/calendar.js";

// Email tools
import {
  listEmailsSchema,
  searchEmailsSchema,
  readEmailSchema,
  sendEmailSchema,
  markAsReadSchema,
  handleListEmails,
  handleSearchEmails,
  handleReadEmail,
  handleSendEmail,
  handleMarkAsRead,
} from "./tools/email.js";

// Folder tools
import {
  listFoldersSchema,
  createFolderSchema,
  moveEmailsSchema,
  handleListFolders,
  handleCreateFolder,
  handleMoveEmails,
} from "./tools/folders.js";

// Rules tools
import {
  listRulesSchema,
  createRuleSchema,
  editRuleSequenceSchema,
  handleListRules,
  handleCreateRule,
  handleEditRuleSequence,
} from "./tools/rules.js";

// =============================================================================
// Configuration
// =============================================================================

const SERVER_NAME = "outlook-assistant";
const SERVER_VERSION = "2.0.0";

// =============================================================================
// Server Setup
// =============================================================================

const server = new McpServer({
  name: SERVER_NAME,
  version: SERVER_VERSION,
});

// =============================================================================
// Auth Tools Registration
// =============================================================================

server.tool("about", "Returns information about this Outlook Assistant server", aboutSchema.shape, async () => {
  return handleAbout();
});

server.tool(
  "authenticate",
  "Authenticate with Microsoft Graph API to access Outlook data",
  authenticateSchema.shape,
  async (args) => {
    return handleAuthenticate(args);
  }
);

server.tool(
  "check-auth-status",
  "Check the current authentication status with Microsoft Graph API",
  checkAuthStatusSchema.shape,
  async () => {
    return handleCheckAuthStatus();
  }
);

// =============================================================================
// Calendar Tools Registration
// =============================================================================

server.tool(
  "list-events",
  "Lists events from your calendar. By default shows upcoming events. Use startDate and endDate to query past events or a specific date range.",
  listEventsSchema.shape,
  async (args) => {
    return handleListEvents(args);
  }
);

server.tool("decline-event", "Declines a calendar event", declineEventSchema.shape, async (args) => {
  return handleDeclineEvent(args);
});

server.tool("create-event", "Creates a new calendar event", createEventSchema.shape, async (args) => {
  return handleCreateEvent(args);
});

server.tool("update-event", "Updates an existing calendar event", updateEventSchema.shape, async (args) => {
  return handleUpdateEvent(args);
});

server.tool("cancel-event", "Cancels a calendar event", cancelEventSchema.shape, async (args) => {
  return handleCancelEvent(args);
});

server.tool("delete-event", "Deletes a calendar event", deleteEventSchema.shape, async (args) => {
  return handleDeleteEvent(args);
});

server.tool(
  "list-categories",
  "Lists all calendar categories",
  listCategoriesSchema.shape,
  async () => {
    return handleListCategories();
  }
);

server.tool(
  "create-category",
  "Creates a new calendar category",
  createCategorySchema.shape,
  async (args) => {
    return handleCreateCategory(args);
  }
);

server.tool(
  "delete-category",
  "Deletes a calendar category",
  deleteCategorySchema.shape,
  async (args) => {
    return handleDeleteCategory(args);
  }
);

// =============================================================================
// Email Tools Registration
// =============================================================================

server.tool(
  "list-emails",
  "Lists recent emails from your inbox",
  listEmailsSchema.shape,
  async (args) => {
    return handleListEmails(args);
  }
);

server.tool(
  "search-emails",
  "Search for emails using various criteria",
  searchEmailsSchema.shape,
  async (args) => {
    return handleSearchEmails(args);
  }
);

server.tool(
  "read-email",
  "Reads the content of a specific email",
  readEmailSchema.shape,
  async (args) => {
    return handleReadEmail(args);
  }
);

server.tool(
  "send-email",
  "Composes and sends a new email",
  sendEmailSchema.shape,
  async (args) => {
    return handleSendEmail(args);
  }
);

server.tool(
  "mark-as-read",
  "Marks an email as read or unread",
  markAsReadSchema.shape,
  async (args) => {
    return handleMarkAsRead(args);
  }
);

// =============================================================================
// Folder Tools Registration
// =============================================================================

server.tool(
  "list-folders",
  "Lists mail folders in your Outlook account",
  listFoldersSchema.shape,
  async (args) => {
    return handleListFolders(args);
  }
);

server.tool(
  "create-folder",
  "Creates a new mail folder",
  createFolderSchema.shape,
  async (args) => {
    return handleCreateFolder(args);
  }
);

server.tool(
  "move-emails",
  "Moves emails from one folder to another",
  moveEmailsSchema.shape,
  async (args) => {
    return handleMoveEmails(args);
  }
);

// =============================================================================
// Rules Tools Registration
// =============================================================================

server.tool(
  "list-rules",
  "Lists inbox rules in your Outlook account",
  listRulesSchema.shape,
  async (args) => {
    return handleListRules(args);
  }
);

server.tool(
  "create-rule",
  "Creates a new inbox rule",
  createRuleSchema.shape,
  async (args) => {
    return handleCreateRule(args);
  }
);

server.tool(
  "edit-rule-sequence",
  "Changes the execution order of an existing inbox rule",
  editRuleSequenceSchema.shape,
  async (args) => {
    return handleEditRuleSequence(args);
  }
);

// =============================================================================
// Server Startup
// =============================================================================

async function main() {
  const transport = new StdioServerTransport();

  await server.connect(transport);

  // Handle graceful shutdown
  process.on("SIGINT", async () => {
    await server.close();
    process.exit(0);
  });

  process.on("SIGTERM", async () => {
    await server.close();
    process.exit(0);
  });
}

main().catch((error) => {
  console.error("Server error:", error);
  process.exit(1);
});
