# @mcp-z/mcp-outlook

Docs: https://mcp-z.github.io/mcp-outlook
Outlook MCP server for searching, reading, and sending Microsoft 365 mail.

## Common uses

- Search and read messages
- Send and reply to emails
- Manage categories and export messages to CSV

## Transports

MCP supports stdio and HTTP.

**Stdio**
```json
{
  "mcpServers": {
    "outlook": {
      "command": "npx",
      "args": ["-y", "@mcp-z/mcp-outlook"]
    }
  }
}
```

**HTTP**
```json
{
  "mcpServers": {
    "outlook": {
      "type": "http",
      "url": "http://localhost:9003/mcp",
      "start": {
        "command": "npx",
        "args": ["-y", "@mcp-z/mcp-outlook", "--port=9003"]
      }
    }
  }
}
```

`start` is an extension used by `npx @mcp-z/cli up` to launch HTTP servers for you.

## Create a Microsoft app

1. Go to [Azure Portal](https://portal.azure.com/).
2. Navigate to Azure Active Directory > App registrations.
3. Click New registration.
4. Choose a name and select a supported account type.
5. Copy the Application (client) ID and Directory (tenant) ID.

## OAuth modes

Configure via environment variables or the `env` block in `.mcp.json`. See `server.json` for the full list of options.

### Loopback OAuth (default)

Environment variables:

```bash
MS_CLIENT_ID=your-client-id
MS_TENANT_ID=common
MS_CLIENT_SECRET=your-client-secret
```

Example:
```json
{
  "mcpServers": {
    "outlook": {
      "command": "npx",
      "args": ["-y", "@mcp-z/mcp-outlook"],
      "env": {
        "MS_CLIENT_ID": "your-client-id",
        "MS_TENANT_ID": "common",
        "MS_CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

### Device code

Useful for headless or remote environments.

```json
{
  "mcpServers": {
    "outlook": {
      "command": "npx",
      "args": ["-y", "@mcp-z/mcp-outlook", "--auth=device-code"],
      "env": {
        "MS_CLIENT_ID": "your-client-id",
        "MS_TENANT_ID": "common"
      }
    }
  }
}
```

### DCR (self-hosted)

HTTP only. Requires a public base URL.

```json
{
  "mcpServers": {
    "outlook-dcr": {
      "command": "npx",
      "args": [
        "-y",
        "@mcp-z/mcp-outlook",
        "--auth=dcr",
        "--port=3456",
        "--base-url=https://oauth.example.com"
      ],
      "env": {
        "MS_CLIENT_ID": "your-client-id",
        "MS_TENANT_ID": "common",
        "MS_CLIENT_SECRET": "your-client-secret"
      }
    }
  }
}
```

## How to use

```bash
# List tools
mcp-z inspect --servers outlook --tools

# Call a tool
mcp-z call outlook message-search '{"query":"from:alice@example.com"}'
```

## Tools

1. categories-list
2. label-add
3. label-delete
4. labels-list
5. message-get
6. message-mark-read
7. message-move-to-trash
8. message-respond
9. message-search
10. message-send
11. messages-export-csv

## Resources

1. email

## Prompts

1. draft-email
2. query-syntax

## Configuration reference

See `server.json` for all supported environment variables, CLI arguments, and defaults.
