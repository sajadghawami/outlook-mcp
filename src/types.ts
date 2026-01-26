/**
 * TypeScript type definitions for Outlook MCP Server
 */

// =============================================================================
// Authentication Types
// =============================================================================

export interface TokenData {
  access_token: string;
  refresh_token?: string;
  expires_at: number;
  token_type?: string;
  scope?: string;
}

// =============================================================================
// Email Types
// =============================================================================

export interface EmailAddress {
  name: string;
  address: string;
}

export interface Recipient {
  emailAddress: EmailAddress;
}

export interface EmailBody {
  contentType: "text" | "html";
  content: string;
}

export interface Email {
  id: string;
  subject: string;
  from?: Recipient;
  toRecipients?: Recipient[];
  ccRecipients?: Recipient[];
  bccRecipients?: Recipient[];
  receivedDateTime: string;
  bodyPreview?: string;
  body?: EmailBody;
  hasAttachments: boolean;
  importance: "low" | "normal" | "high";
  isRead: boolean;
  internetMessageHeaders?: Array<{
    name: string;
    value: string;
  }>;
}

export interface SendEmailRecipient {
  emailAddress: {
    address: string;
    name?: string;
  };
}

export interface SendEmailMessage {
  subject: string;
  body: EmailBody;
  toRecipients: SendEmailRecipient[];
  ccRecipients?: SendEmailRecipient[];
  bccRecipients?: SendEmailRecipient[];
  importance?: "low" | "normal" | "high";
}

export interface SendEmailPayload {
  message: SendEmailMessage;
  saveToSentItems: boolean;
}

// =============================================================================
// Calendar Types
// =============================================================================

export interface DateTimeTimeZone {
  dateTime: string;
  timeZone: string;
}

export interface Location {
  displayName?: string;
}

export interface Attendee {
  emailAddress: EmailAddress;
  type: "required" | "optional" | "resource";
  status?: {
    response: "none" | "organizer" | "tentativelyAccepted" | "accepted" | "declined" | "notResponded";
    time?: string;
  };
}

export interface CalendarEvent {
  id: string;
  subject: string;
  bodyPreview?: string;
  body?: EmailBody;
  start: DateTimeTimeZone;
  end: DateTimeTimeZone;
  location?: Location;
  organizer?: {
    emailAddress: EmailAddress;
  };
  attendees?: Attendee[];
  isAllDay?: boolean;
  isCancelled?: boolean;
  categories?: string[];
}

export interface CreateEventPayload {
  subject: string;
  start: DateTimeTimeZone;
  end: DateTimeTimeZone;
  attendees?: Array<{
    emailAddress: { address: string };
    type: "required" | "optional";
  }>;
  body?: EmailBody;
  location?: Location;
  categories?: string[];
}

export interface UpdateEventPayload {
  subject?: string;
  start?: DateTimeTimeZone;
  end?: DateTimeTimeZone;
  location?: Location;
  body?: EmailBody;
  attendees?: Array<{
    emailAddress: { address: string };
    type: "required" | "optional";
  }>;
  categories?: string[];
}

// =============================================================================
// Category Types
// =============================================================================

export interface Category {
  id: string;
  displayName: string;
  color: string;
}

export const PRESET_COLORS = [
  "none",
  "preset0",
  "preset1",
  "preset2",
  "preset3",
  "preset4",
  "preset5",
  "preset6",
  "preset7",
  "preset8",
  "preset9",
  "preset10",
  "preset11",
  "preset12",
  "preset13",
  "preset14",
  "preset15",
  "preset16",
  "preset17",
  "preset18",
  "preset19",
  "preset20",
  "preset21",
  "preset22",
  "preset23",
  "preset24",
] as const;

export type PresetColor = (typeof PRESET_COLORS)[number];

export const COLOR_NAMES: Record<string, string> = {
  none: "No color",
  preset0: "Red",
  preset1: "Orange",
  preset2: "Brown",
  preset3: "Yellow",
  preset4: "Green",
  preset5: "Teal",
  preset6: "Olive",
  preset7: "Blue",
  preset8: "Purple",
  preset9: "Cranberry",
  preset10: "Steel",
  preset11: "DarkSteel",
  preset12: "Gray",
  preset13: "DarkGray",
  preset14: "Black",
  preset15: "DarkRed",
  preset16: "DarkOrange",
  preset17: "DarkBrown",
  preset18: "DarkYellow",
  preset19: "DarkGreen",
  preset20: "DarkTeal",
  preset21: "DarkOlive",
  preset22: "DarkBlue",
  preset23: "DarkPurple",
  preset24: "DarkCranberry",
};

// =============================================================================
// Folder Types
// =============================================================================

export interface MailFolder {
  id: string;
  displayName: string;
  parentFolderId?: string;
  childFolderCount: number;
  totalItemCount?: number;
  unreadItemCount?: number;
  isTopLevel?: boolean;
  parentFolder?: string;
}

export const WELL_KNOWN_FOLDERS: Record<string, string> = {
  inbox: "me/mailFolders/inbox/messages",
  drafts: "me/mailFolders/drafts/messages",
  sent: "me/mailFolders/sentItems/messages",
  deleted: "me/mailFolders/deletedItems/messages",
  junk: "me/mailFolders/junkemail/messages",
  archive: "me/mailFolders/archive/messages",
};

// =============================================================================
// Rules Types
// =============================================================================

export interface RuleConditions {
  fromAddresses?: Array<{ emailAddress: EmailAddress }>;
  subjectContains?: string[];
  bodyContains?: string[];
  hasAttachment?: boolean;
  importance?: "low" | "normal" | "high";
}

export interface RuleActions {
  moveToFolder?: string;
  copyToFolder?: string;
  markAsRead?: boolean;
  markImportance?: "low" | "normal" | "high";
  forwardTo?: Array<{ emailAddress: EmailAddress }>;
  delete?: boolean;
}

export interface InboxRule {
  id: string;
  displayName: string;
  sequence: number;
  isEnabled: boolean;
  conditions?: RuleConditions;
  actions?: RuleActions;
}

export interface CreateRulePayload {
  displayName: string;
  sequence: number;
  isEnabled: boolean;
  conditions: RuleConditions;
  actions: RuleActions;
}

// =============================================================================
// Graph API Types
// =============================================================================

export interface GraphAPIResponse<T> {
  value: T[];
  "@odata.count"?: number;
  "@odata.nextLink"?: string;
}

export interface GraphAPIError {
  error: {
    code: string;
    message: string;
    innerError?: {
      "request-id"?: string;
      date?: string;
    };
  };
}

// =============================================================================
// Configuration Types
// =============================================================================

export interface AuthConfig {
  clientId: string;
  clientSecret: string;
  redirectUri: string;
  scopes: string[];
  tokenStorePath: string;
  authServerUrl: string;
}

export interface ServerConfig {
  name: string;
  version: string;
  useTestMode: boolean;
  authConfig: AuthConfig;
  graphApiEndpoint: string;
  emailSelectFields: string;
  emailDetailFields: string;
  calendarSelectFields: string;
  defaultPageSize: number;
  maxResultCount: number;
  defaultTimezone: string;
}

// =============================================================================
// MCP Tool Types
// =============================================================================

// Note: Tool handlers should return objects directly compatible with MCP SDK
// The SDK expects: { content: Array<{type: "text", text: string}>, isError?: boolean }
