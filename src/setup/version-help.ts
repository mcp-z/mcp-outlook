import * as fs from 'fs';
import moduleRoot from 'module-root-sync';
import * as path from 'path';
import * as url from 'url';
import { parseArgs } from 'util';

// Kept dependency-free (fs/path/url/module-root-sync only) so `--version`/`--help`/`version`
// resolve nothing beyond Node startup: config.ts's parseConfig pulls in @mcp-z/oauth-microsoft
// and @mcp-z/server, which this path must not touch.
const pkg = JSON.parse(fs.readFileSync(path.join(moduleRoot(url.fileURLToPath(import.meta.url)), 'package.json'), 'utf-8'));

const HELP_TEXT = `
Usage: mcp-outlook [options]

MCP server for Outlook/Microsoft email management with OAuth authentication.

Options:
  --version              Show version number
  --help                 Show this help message
  --auth=<mode>          Authentication mode (default: loopback-oauth)
                         Modes: loopback-oauth, device-code, dcr
  --headless             Disable browser auto-open, return auth URL instead
  --redirect-uri=<uri>   OAuth redirect URI (default: ephemeral loopback)
  --tenant-id=<id>       Microsoft tenant ID (overrides MS_TENANT_ID env var)
  --dcr-mode=<mode>      DCR mode (self-hosted or external, default: self-hosted)
  --dcr-verify-url=<url> External verification endpoint (required for external mode)
  --dcr-store-uri=<uri>  DCR client storage URI (required for self-hosted mode)
  --port=<port>          Enable HTTP transport on specified port
  --stdio                Enable stdio transport (default if no port)
  --log-level=<level>    Logging level (default: info)
  --resource-store-uri=<uri>    Resource store URI for CSV file storage (default: file://~/.mcp-z/mcp-outlook/files)
  --base-url=<url>       Base URL for HTTP file serving (default: http://localhost for HTTP transports)

Commands:
  version                Show version number

Environment Variables:
  MS_CLIENT_ID           OAuth client ID (REQUIRED)
  MS_TENANT_ID           Microsoft tenant ID (REQUIRED)
  MS_CLIENT_SECRET       OAuth client secret (optional)
  AUTH_MODE              Default authentication mode (optional)
  HEADLESS               Disable browser auto-open (optional)
  DCR_MODE               DCR mode (optional, same format as --dcr-mode)
  DCR_VERIFY_URL         External verification URL (optional, same as --dcr-verify-url)
  DCR_STORE_URI          DCR storage URI (optional, same as --dcr-store-uri)
  TOKEN_STORE_URI        Token storage URI (optional)
  PORT                   Default HTTP port (optional)
  LOG_LEVEL              Default logging level (optional)
  RESOURCE_STORE_URI            Resource store URI (optional, file://)
  BASE_URL               Base URL for HTTP file serving (optional)

OAuth Scopes:
  openid profile offline_access https://graph.microsoft.com/User.Read https://graph.microsoft.com/Mail.ReadWrite https://graph.microsoft.com/Mail.Send https://graph.microsoft.com/MailboxSettings.ReadWrite

Examples:
  mcp-outlook                           # Use default settings
  mcp-outlook version                   # Print version number
  mcp-outlook --auth=device-code        # Use device code auth
  mcp-outlook --port=3000               # HTTP transport on port 3000
  mcp-outlook --tenant-id=xxx           # Set tenant ID
  mcp-outlook --resource-store-uri=file:///tmp/emails    # Custom resource store URI
  MS_CLIENT_ID=xxx mcp-outlook          # Set client ID via env var
`.trim();

/** Package metadata read from package.json, resolved relative to this file. */
export function readPkg(): { name: string; version: string; repository?: string | { url?: string } } {
  return pkg;
}

/**
 * Handle --version/--help flags and the `version` subcommand before config parsing.
 * These must work without requiring any configuration or heavy dependency.
 */
export function handleVersionHelp(args: string[]): { handled: boolean; output?: string } {
  const { values, positionals } = parseArgs({
    args,
    options: {
      version: { type: 'boolean' },
      help: { type: 'boolean' },
    },
    strict: false,
    allowPositionals: true,
  });

  if (values.version || positionals[0] === 'version') return { handled: true, output: pkg.version };
  if (values.help) return { handled: true, output: HELP_TEXT };
  return { handled: false };
}
