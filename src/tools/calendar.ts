/**
 * Calendar tools for Outlook MCP Server
 */

import { z } from "zod";
import { graphGet, graphPost, graphPatch, graphDelete } from "../api.js";
import type {
  CalendarEvent,
  Category,
  GraphAPIResponse,
  CreateEventPayload,
  UpdateEventPayload,
} from "../types.js";
import { PRESET_COLORS, COLOR_NAMES } from "../types.js";

// =============================================================================
// Configuration
// =============================================================================

const DEFAULT_TIMEZONE = "Central European Standard Time";
const CALENDAR_SELECT_FIELDS =
  "id,subject,bodyPreview,start,end,location,organizer,attendees,isAllDay,isCancelled,categories";
const MAX_RESULT_COUNT = 50;

// =============================================================================
// Schemas
// =============================================================================

export const listEventsSchema = z.object({
  count: z.number().optional().describe("Number of events to retrieve (default: 10, max: 50)"),
  startDate: z
    .string()
    .optional()
    .describe("Start date for the query in ISO 8601 format (e.g., '2025-01-01'). Defaults to now."),
  endDate: z
    .string()
    .optional()
    .describe(
      "End date for the query in ISO 8601 format (e.g., '2025-12-31'). If provided, returns events between startDate and endDate."
    ),
});

export const declineEventSchema = z.object({
  eventId: z.string().describe("The ID of the event to decline"),
  comment: z.string().optional().describe("Optional comment for declining the event"),
});

export const createEventSchema = z.object({
  subject: z.string().describe("The subject of the event"),
  start: z.string().describe("The start time of the event in ISO 8601 format"),
  end: z.string().describe("The end time of the event in ISO 8601 format"),
  attendees: z.array(z.string()).optional().describe("List of attendee email addresses"),
  body: z.string().optional().describe("Optional body content for the event"),
});

export const updateEventSchema = z.object({
  eventId: z.string().describe("The ID of the event to update"),
  subject: z.string().optional().describe("New subject for the event"),
  start: z.string().optional().describe("New start time in ISO 8601 format"),
  end: z.string().optional().describe("New end time in ISO 8601 format"),
  location: z.string().optional().describe("New location for the event"),
  body: z.string().optional().describe("New body content for the event"),
  attendees: z.array(z.string()).optional().describe("New list of attendee email addresses"),
  categories: z.array(z.string()).optional().describe("New list of category names"),
});

export const cancelEventSchema = z.object({
  eventId: z.string().describe("The ID of the event to cancel"),
  comment: z.string().optional().describe("Optional comment for cancelling the event"),
});

export const deleteEventSchema = z.object({
  eventId: z.string().describe("The ID of the event to delete"),
});

export const listCategoriesSchema = z.object({});

export const createCategorySchema = z.object({
  name: z.string().describe("The name of the category"),
  color: z
    .string()
    .optional()
    .describe(
      "The color preset (none, preset0-preset24). Colors: preset0=Red, preset1=Orange, preset2=Brown, preset3=Yellow, preset4=Green, preset5=Teal, preset6=Olive, preset7=Blue, preset8=Purple, preset9=Cranberry"
    ),
});

export const deleteCategorySchema = z.object({
  id: z.string().describe("The ID of the category to delete (use list-categories to find IDs)"),
});

// =============================================================================
// Handlers
// =============================================================================

export async function handleListEvents(
  args: z.infer<typeof listEventsSchema>
): Promise<{ content: Array<{ type: "text"; text: string }> }> {
  try {
    const count = Math.min(args.count || 10, MAX_RESULT_COUNT);
    const startDate = args.startDate ? new Date(args.startDate) : new Date();
    const endDate = args.endDate ? new Date(args.endDate) : null;

    const queryStartDate = startDate.toISOString();
    const queryEndDate = endDate
      ? endDate.toISOString()
      : new Date(startDate.getTime() + 30 * 24 * 60 * 60 * 1000).toISOString();

    // Determine sort order: descending for past events, ascending for future
    const isPastQuery = endDate && endDate < new Date();
    const orderDirection = isPastQuery ? "desc" : "asc";

    const response = await graphGet<GraphAPIResponse<CalendarEvent>>("me/calendarView", {
      startDateTime: queryStartDate,
      endDateTime: queryEndDate,
      $top: count,
      $orderby: `start/dateTime ${orderDirection}`,
      $select: CALENDAR_SELECT_FIELDS,
    });

    if (!response.value || response.value.length === 0) {
      return {
        content: [{ type: "text", text: "No calendar events found." }],
      };
    }

    const eventList = response.value
      .map((event, index) => {
        const start = new Date(event.start.dateTime).toLocaleString();
        const end = new Date(event.end.dateTime).toLocaleString();
        const location = event.location?.displayName || "No location";
        const categories = event.categories?.length ? event.categories.join(", ") : "None";

        return `${index + 1}. ${event.subject} - Location: ${location}
Start: ${start}
End: ${end}
Categories: ${categories}
Summary: ${event.bodyPreview || "No summary"}
ID: ${event.id}`;
      })
      .join("\n\n");

    return {
      content: [{ type: "text", text: `Found ${response.value.length} events:\n\n${eventList}` }],
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
        { type: "text", text: `Error listing events: ${error instanceof Error ? error.message : "Unknown error"}` },
      ],
    };
  }
}

export async function handleDeclineEvent(
  args: z.infer<typeof declineEventSchema>
): Promise<{ content: Array<{ type: "text"; text: string }> }> {
  try {
    if (!args.eventId) {
      return {
        content: [{ type: "text", text: "Event ID is required to decline an event." }],
      };
    }

    await graphPost(`me/events/${args.eventId}/decline`, {
      comment: args.comment || "Declined via API",
    });

    return {
      content: [{ type: "text", text: `Event with ID ${args.eventId} has been successfully declined.` }],
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
        { type: "text", text: `Error declining event: ${error instanceof Error ? error.message : "Unknown error"}` },
      ],
    };
  }
}

export async function handleCreateEvent(
  args: z.infer<typeof createEventSchema>
): Promise<{ content: Array<{ type: "text"; text: string }> }> {
  try {
    if (!args.subject || !args.start || !args.end) {
      return {
        content: [{ type: "text", text: "Subject, start, and end times are required to create an event." }],
      };
    }

    const payload: CreateEventPayload = {
      subject: args.subject,
      start: { dateTime: args.start, timeZone: DEFAULT_TIMEZONE },
      end: { dateTime: args.end, timeZone: DEFAULT_TIMEZONE },
      attendees: args.attendees?.map((email) => ({
        emailAddress: { address: email },
        type: "required",
      })),
      body: args.body ? { contentType: "html", content: args.body } : undefined,
    };

    await graphPost("me/events", payload);

    return {
      content: [{ type: "text", text: `Event '${args.subject}' has been successfully created.` }],
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
        { type: "text", text: `Error creating event: ${error instanceof Error ? error.message : "Unknown error"}` },
      ],
    };
  }
}

export async function handleUpdateEvent(
  args: z.infer<typeof updateEventSchema>
): Promise<{ content: Array<{ type: "text"; text: string }> }> {
  try {
    if (!args.eventId) {
      return {
        content: [{ type: "text", text: "Error: Event ID is required." }],
      };
    }

    const updateData: UpdateEventPayload = {};

    if (args.subject !== undefined) updateData.subject = args.subject;
    if (args.start !== undefined) updateData.start = { dateTime: args.start, timeZone: DEFAULT_TIMEZONE };
    if (args.end !== undefined) updateData.end = { dateTime: args.end, timeZone: DEFAULT_TIMEZONE };
    if (args.location !== undefined) updateData.location = { displayName: args.location };
    if (args.body !== undefined) updateData.body = { contentType: "text", content: args.body };
    if (args.attendees !== undefined) {
      updateData.attendees = args.attendees.map((email) => ({
        emailAddress: { address: email },
        type: "required",
      }));
    }
    if (args.categories !== undefined) updateData.categories = args.categories;

    if (Object.keys(updateData).length === 0) {
      return {
        content: [
          {
            type: "text",
            text: "Error: No fields to update. Provide at least one of: subject, start, end, location, body, attendees, categories.",
          },
        ],
      };
    }

    const response = await graphPatch<CalendarEvent>(`me/events/${args.eventId}`, updateData);

    const startDate = new Date(response.start.dateTime).toLocaleString();
    const endDate = new Date(response.end.dateTime).toLocaleString();
    const eventCategories = response.categories?.length ? response.categories.join(", ") : "None";

    return {
      content: [
        {
          type: "text",
          text: `Event updated successfully!

Subject: ${response.subject}
Start: ${startDate}
End: ${endDate}
Location: ${response.location?.displayName || "No location"}
Categories: ${eventCategories}
ID: ${response.id}`,
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
        { type: "text", text: `Error updating event: ${error instanceof Error ? error.message : "Unknown error"}` },
      ],
    };
  }
}

export async function handleCancelEvent(
  args: z.infer<typeof cancelEventSchema>
): Promise<{ content: Array<{ type: "text"; text: string }> }> {
  try {
    if (!args.eventId) {
      return {
        content: [{ type: "text", text: "Event ID is required to cancel an event." }],
      };
    }

    await graphPost(`me/events/${args.eventId}/cancel`, {
      comment: args.comment || "Cancelled via API",
    });

    return {
      content: [{ type: "text", text: `Event with ID ${args.eventId} has been successfully cancelled.` }],
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
        { type: "text", text: `Error cancelling event: ${error instanceof Error ? error.message : "Unknown error"}` },
      ],
    };
  }
}

export async function handleDeleteEvent(
  args: z.infer<typeof deleteEventSchema>
): Promise<{ content: Array<{ type: "text"; text: string }> }> {
  try {
    if (!args.eventId) {
      return {
        content: [{ type: "text", text: "Event ID is required to delete an event." }],
      };
    }

    await graphDelete(`me/events/${args.eventId}`);

    return {
      content: [{ type: "text", text: `Event with ID ${args.eventId} has been successfully deleted.` }],
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
        { type: "text", text: `Error deleting event: ${error instanceof Error ? error.message : "Unknown error"}` },
      ],
    };
  }
}

export async function handleListCategories(): Promise<{ content: Array<{ type: "text"; text: string }> }> {
  try {
    const response = await graphGet<GraphAPIResponse<Category>>("me/outlook/masterCategories");

    if (!response.value || response.value.length === 0) {
      return {
        content: [{ type: "text", text: "No categories found." }],
      };
    }

    const categoryList = response.value
      .map((cat, index) => {
        const colorName = COLOR_NAMES[cat.color] || cat.color;
        return `${index + 1}. ${cat.displayName} (${colorName})\n   ID: ${cat.id}`;
      })
      .join("\n\n");

    return {
      content: [{ type: "text", text: `Found ${response.value.length} categories:\n\n${categoryList}` }],
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
        { type: "text", text: `Error listing categories: ${error instanceof Error ? error.message : "Unknown error"}` },
      ],
    };
  }
}

export async function handleCreateCategory(
  args: z.infer<typeof createCategorySchema>
): Promise<{ content: Array<{ type: "text"; text: string }> }> {
  try {
    if (!args.name) {
      return {
        content: [{ type: "text", text: "Error: Category name is required." }],
      };
    }

    const categoryColor = args.color || "none";
    if (!PRESET_COLORS.includes(categoryColor as (typeof PRESET_COLORS)[number])) {
      return {
        content: [{ type: "text", text: `Error: Invalid color. Valid colors are: ${PRESET_COLORS.join(", ")}` }],
      };
    }

    const response = await graphPost<Category>("me/outlook/masterCategories", {
      displayName: args.name,
      color: categoryColor,
    });

    const colorName = COLOR_NAMES[response.color] || response.color;

    return {
      content: [
        {
          type: "text",
          text: `Category created successfully!\n\nName: ${response.displayName}\nColor: ${colorName}\nID: ${response.id}`,
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
        { type: "text", text: `Error creating category: ${error instanceof Error ? error.message : "Unknown error"}` },
      ],
    };
  }
}

export async function handleDeleteCategory(
  args: z.infer<typeof deleteCategorySchema>
): Promise<{ content: Array<{ type: "text"; text: string }> }> {
  try {
    if (!args.id) {
      return {
        content: [{ type: "text", text: "Error: Category ID is required. Use list-categories to find the ID." }],
      };
    }

    await graphDelete(`me/outlook/masterCategories/${args.id}`);

    return {
      content: [{ type: "text", text: "Category deleted successfully." }],
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
        { type: "text", text: `Error deleting category: ${error instanceof Error ? error.message : "Unknown error"}` },
      ],
    };
  }
}
