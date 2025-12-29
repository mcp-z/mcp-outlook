/**
 * Outlook MCP Server Constants
 *
 * These scopes are required for Microsoft Outlook functionality and are hardcoded
 * rather than externally configured since this server knows its own requirements.
 */

import { EMAIL_CHUNK_SIZE, EMAIL_MAX_BATCH_SIZE } from '@mcp-z/email';

// Microsoft OAuth scopes required for Outlook operations
export const MS_SCOPE = 'openid profile offline_access https://graph.microsoft.com/User.Read https://graph.microsoft.com/Mail.ReadWrite https://graph.microsoft.com/Mail.Send https://graph.microsoft.com/MailboxSettings.ReadWrite';

// Batch processing configuration (from shared email constants)
export const CHUNK_SIZE = EMAIL_CHUNK_SIZE;
export const MAX_BATCH_SIZE = EMAIL_MAX_BATCH_SIZE;

// Pagination configuration
export const MAX_PAGE_SIZE = 1000; // Maximum number of items per page
