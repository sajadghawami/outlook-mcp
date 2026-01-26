# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Development Commands

- `pnpm install` - **ALWAYS run first** to install dependencies
- `pnpm build` - Compile TypeScript to JavaScript in build/
- `pnpm start` - Start the compiled MCP server
- `pnpm dev` - Watch mode for TypeScript compilation
- `pnpm auth-server` - Start the OAuth authentication server on port 1337 (**required for authentication**)
- `pnpm inspect` - Build and use MCP Inspector to test the server interactively
- `npx kill-port 1337` - Kill process using port 1337 if auth server won't start

## Architecture Overview

This is a TypeScript MCP (Model Context Protocol) server that provides Claude with access to Microsoft Outlook via the Microsoft Graph API.

### Project Structure
```
outlook-mcp/
├── src/
│   ├── index.ts           # Main MCP server entry point
│   ├── api.ts             # Microsoft Graph API client & token management
│   ├── types.ts           # TypeScript interfaces and type definitions
│   ├── auth-server.ts     # Standalone OAuth server for authentication
│   └── tools/
│       ├── auth.ts        # Authentication tools (about, authenticate, check-auth-status)
│       ├── calendar.ts    # Calendar tools (list-events, create-event, etc.)
│       ├── email.ts       # Email tools (list-emails, search, read, send, mark-as-read)
│       ├── folders.ts     # Folder tools (list-folders, create-folder, move-emails)
│       └── rules.ts       # Rules tools (list-rules, create-rule, edit-rule-sequence)
├── build/                 # Compiled JavaScript output (gitignored)
├── package.json           # Dependencies and scripts
└── tsconfig.json          # TypeScript configuration
```

### Key Components

- **MCP Server (index.ts)**: Registers all 23 tools using McpServer from @modelcontextprotocol/sdk
- **API Client (api.ts)**: Handles Microsoft Graph API calls, token management, and pagination
- **Types (types.ts)**: TypeScript interfaces for emails, calendar events, folders, rules, etc.
- **Auth Server (auth-server.ts)**: Standalone HTTP server for OAuth 2.0 callback handling
- **Tool Modules**: Each module exports Zod schemas for validation and handler functions

### Tool Registration Pattern

Tools are defined with Zod schemas and registered in index.ts:
```typescript
server.tool(
  "tool-name",
  "Tool description",
  zodSchema.shape,
  async (args) => handleTool(args)
);
```

## Authentication Flow

1. Azure app registration required with specific permissions (Mail.Read, Mail.Send, Calendars.ReadWrite, etc.)
2. Auth server auto-starts when using authenticate tool (or run manually: `pnpm auth-server`)
3. Use authenticate tool to open OAuth URL in browser
4. Complete browser authentication with Microsoft
5. Tokens automatically stored in `~/.outlook-mcp-keys.json`

## Configuration Requirements

### Environment Variables
Configure credentials in your `.mcp.json` file using:
- `AZURE_CLIENT_ID` - Your Azure application (client) ID
- `AZURE_CLIENT_SECRET` - Your Azure client secret VALUE (not the Secret ID)

### Common Setup Issues
1. **Missing dependencies**: Always run `pnpm install` first
2. **Build errors**: Run `pnpm build` before `pnpm start`
3. **Wrong secret**: Use Azure secret VALUE, not ID (AADSTS7000215 error)
4. **Auth server not running**: Start `pnpm auth-server` before authenticating
5. **Port conflicts**: Use `npx kill-port 1337` if port is in use

## Available Tools (23 total)

### Auth (3)
- `about` - Server information
- `authenticate` - Initiate OAuth flow
- `check-auth-status` - Check authentication state

### Calendar (9)
- `list-events` - Query calendar events
- `create-event` - Create new event
- `update-event` - Modify event
- `decline-event` - RSVP decline
- `cancel-event` - Cancel event
- `delete-event` - Remove event
- `list-categories` - Get categories
- `create-category` - Add category
- `delete-category` - Remove category

### Email (5)
- `list-emails` - List inbox emails
- `search-emails` - Search with filters
- `read-email` - Get email content
- `send-email` - Compose and send
- `mark-as-read` - Toggle read status

### Folders (3)
- `list-folders` - List mail folders
- `create-folder` - Create new folder
- `move-emails` - Move emails between folders

### Rules (3)
- `list-rules` - List inbox rules
- `create-rule` - Create new rule
- `edit-rule-sequence` - Change rule order

## Error Handling

- Authentication failures return "Authentication required" error
- Graph API errors include status codes and response details
- Token expiration triggers re-authentication flow
- All handlers wrap errors with descriptive messages
