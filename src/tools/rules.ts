/**
 * Email rules management tools for Outlook MCP Server
 */

import { z } from "zod";
import { graphGet, graphPost, graphPatch, ensureAuthenticated, callGraphAPI } from "../api.js";
import type { InboxRule, GraphAPIResponse, CreateRulePayload } from "../types.js";
import { getFolderIdByName } from "./email.js";

// =============================================================================
// Schemas
// =============================================================================

export const listRulesSchema = z.object({
  includeDetails: z.boolean().optional().describe("Include detailed rule conditions and actions"),
});

export const createRuleSchema = z.object({
  name: z.string().describe("Name of the rule to create"),
  fromAddresses: z
    .string()
    .optional()
    .describe("Comma-separated list of sender email addresses for the rule"),
  containsSubject: z.string().optional().describe("Subject text the email must contain"),
  hasAttachments: z.boolean().optional().describe("Whether the rule applies to emails with attachments"),
  moveToFolder: z.string().optional().describe("Name of the folder to move matching emails to"),
  markAsRead: z.boolean().optional().describe("Whether to mark matching emails as read"),
  isEnabled: z
    .boolean()
    .optional()
    .describe("Whether the rule should be enabled after creation (default: true)"),
  sequence: z
    .number()
    .optional()
    .describe("Order in which the rule is executed (lower numbers run first, default: 100)"),
});

export const editRuleSequenceSchema = z.object({
  ruleName: z.string().describe("Name of the rule to modify"),
  sequence: z.number().describe("New sequence value for the rule (lower numbers run first)"),
});

// =============================================================================
// Helper Functions
// =============================================================================

async function getInboxRules(accessToken: string): Promise<InboxRule[]> {
  try {
    const response = await callGraphAPI<GraphAPIResponse<InboxRule>>(
      accessToken,
      "GET",
      "me/mailFolders/inbox/messageRules",
      null
    );
    return response.value || [];
  } catch (error) {
    throw error;
  }
}

function formatRuleConditions(rule: InboxRule): string {
  const conditions: string[] = [];

  if (rule.conditions?.fromAddresses && rule.conditions.fromAddresses.length > 0) {
    const senders = rule.conditions.fromAddresses.map((addr) => addr.emailAddress.address).join(", ");
    conditions.push(`From: ${senders}`);
  }

  if (rule.conditions?.subjectContains && rule.conditions.subjectContains.length > 0) {
    conditions.push(`Subject contains: "${rule.conditions.subjectContains.join(", ")}"`);
  }

  if (rule.conditions?.bodyContains && rule.conditions.bodyContains.length > 0) {
    conditions.push(`Body contains: "${rule.conditions.bodyContains.join(", ")}"`);
  }

  if (rule.conditions?.hasAttachment === true) {
    conditions.push("Has attachment");
  }

  if (rule.conditions?.importance) {
    conditions.push(`Importance: ${rule.conditions.importance}`);
  }

  return conditions.join("; ");
}

function formatRuleActions(rule: InboxRule): string {
  const actions: string[] = [];

  if (rule.actions?.moveToFolder) {
    actions.push(`Move to folder: ${rule.actions.moveToFolder}`);
  }

  if (rule.actions?.copyToFolder) {
    actions.push(`Copy to folder: ${rule.actions.copyToFolder}`);
  }

  if (rule.actions?.markAsRead === true) {
    actions.push("Mark as read");
  }

  if (rule.actions?.markImportance) {
    actions.push(`Mark importance: ${rule.actions.markImportance}`);
  }

  if (rule.actions?.forwardTo && rule.actions.forwardTo.length > 0) {
    const recipients = rule.actions.forwardTo.map((r) => r.emailAddress.address).join(", ");
    actions.push(`Forward to: ${recipients}`);
  }

  if (rule.actions?.delete === true) {
    actions.push("Delete");
  }

  return actions.join("; ");
}

function formatRulesList(rules: InboxRule[], includeDetails: boolean): string {
  if (!rules || rules.length === 0) {
    return "No inbox rules found.\n\nTip: You can create rules using the 'create-rule' tool. Rules are processed in order of their sequence number (lower numbers are processed first).";
  }

  const sortedRules = [...rules].sort((a, b) => {
    return (a.sequence || 9999) - (b.sequence || 9999);
  });

  if (includeDetails) {
    const detailedRules = sortedRules.map((rule, index) => {
      let ruleText = `${index + 1}. ${rule.displayName}${rule.isEnabled ? "" : " (Disabled)"} - Sequence: ${rule.sequence || "N/A"}`;

      const conditions = formatRuleConditions(rule);
      if (conditions) {
        ruleText += `\n   Conditions: ${conditions}`;
      }

      const actions = formatRuleActions(rule);
      if (actions) {
        ruleText += `\n   Actions: ${actions}`;
      }

      return ruleText;
    });

    return `Found ${rules.length} inbox rules (sorted by execution order):\n\n${detailedRules.join("\n\n")}\n\nRules are processed in order of their sequence number. You can change rule order using the 'edit-rule-sequence' tool.`;
  } else {
    const simpleRules = sortedRules.map((rule, index) => {
      return `${index + 1}. ${rule.displayName}${rule.isEnabled ? "" : " (Disabled)"} - Sequence: ${rule.sequence || "N/A"}`;
    });

    return `Found ${rules.length} inbox rules (sorted by execution order):\n\n${simpleRules.join("\n")}\n\nTip: Use 'list-rules with includeDetails=true' to see more information about each rule.`;
  }
}

// =============================================================================
// Handlers
// =============================================================================

export async function handleListRules(
  args: z.infer<typeof listRulesSchema>
): Promise<{ content: Array<{ type: "text"; text: string }> }> {
  try {
    const accessToken = ensureAuthenticated();
    const includeDetails = args.includeDetails === true;

    const rules = await getInboxRules(accessToken);
    const formattedRules = formatRulesList(rules, includeDetails);

    return {
      content: [{ type: "text", text: formattedRules }],
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
        { type: "text", text: `Error listing rules: ${error instanceof Error ? error.message : "Unknown error"}` },
      ],
    };
  }
}

export async function handleCreateRule(
  args: z.infer<typeof createRuleSchema>
): Promise<{ content: Array<{ type: "text"; text: string }> }> {
  try {
    if (!args.name) {
      return {
        content: [{ type: "text", text: "Rule name is required." }],
      };
    }

    // Validate sequence if provided
    if (args.sequence !== undefined && (isNaN(args.sequence) || args.sequence < 1)) {
      return {
        content: [{ type: "text", text: "Sequence must be a positive number greater than zero." }],
      };
    }

    // Validate that at least one condition or action is specified
    const hasCondition = args.fromAddresses || args.containsSubject || args.hasAttachments === true;
    const hasAction = args.moveToFolder || args.markAsRead === true;

    if (!hasCondition) {
      return {
        content: [
          {
            type: "text",
            text: "At least one condition is required. Specify fromAddresses, containsSubject, or hasAttachments.",
          },
        ],
      };
    }

    if (!hasAction) {
      return {
        content: [{ type: "text", text: "At least one action is required. Specify moveToFolder or markAsRead." }],
      };
    }

    const accessToken = ensureAuthenticated();

    // Determine sequence
    let ruleSequence = args.sequence;
    if (!ruleSequence) {
      try {
        ruleSequence = 100;
        const existingRules = await getInboxRules(accessToken);
        if (existingRules && existingRules.length > 0) {
          const highestSequence = Math.max(...existingRules.map((r) => r.sequence || 0));
          ruleSequence = Math.max(highestSequence + 1, 100);
        }
      } catch {
        ruleSequence = 100;
      }
    }

    ruleSequence = Math.max(1, Math.floor(ruleSequence));

    // Build rule object
    const rule: CreateRulePayload = {
      displayName: args.name,
      isEnabled: args.isEnabled !== false,
      sequence: ruleSequence,
      conditions: {},
      actions: {},
    };

    // Add conditions
    if (args.fromAddresses) {
      const emailAddresses = args.fromAddresses
        .split(",")
        .map((email) => email.trim())
        .filter((email) => email)
        .map((email) => ({
          emailAddress: {
            name: "",
            address: email,
          },
        }));

      if (emailAddresses.length > 0) {
        rule.conditions.fromAddresses = emailAddresses;
      }
    }

    if (args.containsSubject) {
      rule.conditions.subjectContains = [args.containsSubject];
    }

    if (args.hasAttachments === true) {
      rule.conditions.hasAttachment = true;
    }

    // Add actions
    if (args.moveToFolder) {
      const folderId = await getFolderIdByName(accessToken, args.moveToFolder);
      if (!folderId) {
        return {
          content: [
            {
              type: "text",
              text: `Target folder "${args.moveToFolder}" not found. Please specify a valid folder name.`,
            },
          ],
        };
      }
      rule.actions.moveToFolder = folderId;
    }

    if (args.markAsRead === true) {
      rule.actions.markAsRead = true;
    }

    // Create the rule
    const response = await callGraphAPI<InboxRule>(
      accessToken,
      "POST",
      "me/mailFolders/inbox/messageRules",
      rule
    );

    if (response && response.id) {
      let responseText = `Successfully created rule "${args.name}" with sequence ${ruleSequence}.`;

      if (!args.sequence) {
        responseText +=
          "\n\nTip: You can specify a 'sequence' parameter when creating rules to control their execution order. Lower sequence numbers run first.";
      }

      return {
        content: [{ type: "text", text: responseText }],
      };
    } else {
      return {
        content: [{ type: "text", text: "Failed to create rule. The server didn't return a rule ID." }],
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
        { type: "text", text: `Error creating rule: ${error instanceof Error ? error.message : "Unknown error"}` },
      ],
    };
  }
}

export async function handleEditRuleSequence(
  args: z.infer<typeof editRuleSequenceSchema>
): Promise<{ content: Array<{ type: "text"; text: string }> }> {
  try {
    if (!args.ruleName) {
      return {
        content: [
          { type: "text", text: "Rule name is required. Please specify the exact name of an existing rule." },
        ],
      };
    }

    if (!args.sequence || isNaN(args.sequence) || args.sequence < 1) {
      return {
        content: [
          {
            type: "text",
            text: "A positive sequence number is required. Lower numbers run first (higher priority).",
          },
        ],
      };
    }

    const accessToken = ensureAuthenticated();

    const rules = await getInboxRules(accessToken);
    const rule = rules.find((r) => r.displayName === args.ruleName);

    if (!rule) {
      return {
        content: [{ type: "text", text: `Rule with name "${args.ruleName}" not found.` }],
      };
    }

    await callGraphAPI(accessToken, "PATCH", `me/mailFolders/inbox/messageRules/${rule.id}`, {
      sequence: args.sequence,
    });

    return {
      content: [
        {
          type: "text",
          text: `Successfully updated the sequence of rule "${args.ruleName}" to ${args.sequence}.`,
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
        {
          type: "text",
          text: `Error updating rule sequence: ${error instanceof Error ? error.message : "Unknown error"}`,
        },
      ],
    };
  }
}
