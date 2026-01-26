# Outlook MCP Server

A TypeScript MCP server that connects Claude with Microsoft Outlook via the Microsoft Graph API. Features **23 tools** for managing emails, calendar events, folders, and inbox rules.

## Features

- **Email Management** - List, search, read, send emails, and manage read status
- **Calendar Management** - Create, update, decline, cancel, and delete calendar events
- **Category Management** - List, create, and delete calendar categories
- **Folder Management** - List and create mail folders, move emails between folders
- **Rules Management** - List, create, and modify inbox rules
- **OAuth 2.0 Authentication** - Secure Microsoft Graph API integration

## Prerequisites

- Node.js 18+
- pnpm, npm, or yarn
- Azure account for app registration

## Installation

```bash
git clone https://github.com/sajadghawami/outlook-mcp.git
cd outlook-mcp

# Install dependencies (choose one)
pnpm install   # or: npm install | yarn install

# Build
pnpm build     # or: npm run build | yarn build
```

## Configuration

### Azure App Setup

Before using this server, you'll need to register an application in Azure to obtain API credentials.

#### Register Your Application

1. Sign in to the [Azure Portal](https://portal.azure.com/) with your Microsoft account
2. Navigate to "App registrations" using the search bar
3. Select "New registration" to create a new application
4. Provide a display name for your app (e.g., "Outlook MCP")
5. Under supported account types, choose "Accounts in any organizational directory and personal Microsoft accounts"
6. Configure the redirect URI by selecting "Web" and entering `http://localhost:1337/auth/callback`
7. Click "Register" to complete the process
8. On the overview page, locate and save the **Application (client) ID** — you'll need this for configuration

#### Configure API Permissions

Your application needs permission to access Outlook data through Microsoft Graph:

1. From your app's settings page, select "API permissions"
2. Click "Add a permission", then choose "Microsoft Graph"
3. Select "Delegated permissions" and enable the following:
   - `offline_access` — Maintains access when you're not actively using the app
   - `User.Read` — Reads basic profile information
   - `Mail.Read` — Reads email messages
   - `Mail.ReadWrite` — Manages email messages and folders
   - `Mail.Send` — Sends email on your behalf
   - `Calendars.Read` — Reads calendar events
   - `Calendars.ReadWrite` — Creates and modifies calendar events
   - `MailboxSettings.Read` — Reads mailbox configuration
4. Click "Add permissions" to save

#### Generate a Client Secret

The server needs a secret key to authenticate with Microsoft:

1. In your app's settings, go to "Certificates & secrets"
2. Under "Client secrets", click "New client secret"
3. Add a description (e.g., "MCP Server") and choose an expiration period
4. Click "Add" to generate the secret
5. **Important**: Copy the secret's **Value** immediately — this is shown only once and cannot be retrieved later. Do not copy the "Secret ID".

### Claude Desktop

Add to your Claude Desktop config file:

- macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`
- Windows: `%APPDATA%\Claude\claude_desktop_config.json`

```json
{
  "mcpServers": {
    "outlook-assistant": {
      "command": "node",
      "args": ["/absolute/path/to/outlook-mcp/build/index.js"],
      "env": {
        "AZURE_CLIENT_ID": "your-client-id",
        "AZURE_CLIENT_SECRET": "your-client-secret-VALUE"
      }
    }
  }
}
```

### Claude Code

Add to your `.mcp.json`:

```json
{
  "mcpServers": {
    "outlook-assistant": {
      "command": "node",
      "args": ["/absolute/path/to/outlook-mcp/build/index.js"],
      "env": {
        "AZURE_CLIENT_ID": "your-client-id",
        "AZURE_CLIENT_SECRET": "your-client-secret-VALUE"
      }
    }
  }
}
```

## Available Tools (23 total)

### Authentication (3 tools)

| Tool | Description |
|------|-------------|
| `about` | Returns server information |
| `authenticate` | Initiates OAuth flow with Microsoft |
| `check-auth-status` | Checks current authentication state |

### Calendar (9 tools)

| Tool | Description |
|------|-------------|
| `list-events` | Query calendar events with date range |
| `create-event` | Create new event with attendees |
| `update-event` | Modify event details |
| `decline-event` | RSVP decline to an event |
| `cancel-event` | Cancel an event |
| `delete-event` | Remove an event |
| `list-categories` | Get calendar categories |
| `create-category` | Add category with color |
| `delete-category` | Remove a category |

### Email (5 tools)

| Tool | Description |
|------|-------------|
| `list-emails` | List emails from any folder |
| `search-emails` | Search with filters (from, subject, attachments, unread) |
| `read-email` | Get full email content |
| `send-email` | Compose and send email |
| `mark-as-read` | Toggle read/unread status |

### Folders (3 tools)

| Tool | Description |
|------|-------------|
| `list-folders` | List mail folders with hierarchy |
| `create-folder` | Create new mail folder |
| `move-emails` | Move emails between folders |

### Rules (3 tools)

| Tool | Description |
|------|-------------|
| `list-rules` | List inbox rules |
| `create-rule` | Create rule with conditions and actions |
| `edit-rule-sequence` | Change rule execution order |

## Development

### Build Commands

```bash
pnpm install      # Install dependencies
pnpm build        # Compile TypeScript
pnpm dev          # Watch mode
pnpm start        # Start MCP server
pnpm auth-server  # Start OAuth server on port 1337
pnpm inspect      # Test with MCP Inspector
```

### Project Structure

```
outlook-mcp/
├── src/
│   ├── index.ts           # Main MCP server entry point
│   ├── api.ts             # Microsoft Graph API client
│   ├── types.ts           # TypeScript interfaces
│   ├── auth-server.ts     # OAuth callback server
│   └── tools/
│       ├── auth.ts        # Authentication tools
│       ├── calendar.ts    # Calendar tools
│       ├── email.ts       # Email tools
│       ├── folders.ts     # Folder tools
│       └── rules.ts       # Rules tools
├── build/                 # Compiled output
├── package.json
└── tsconfig.json
```

## License

ISC
