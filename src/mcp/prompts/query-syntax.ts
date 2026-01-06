import type { PromptModule } from '@mcp-z/server';
import type { RequestHandlerExtra } from '@modelcontextprotocol/sdk/shared/protocol.js';
import type { ServerNotification, ServerRequest } from '@modelcontextprotocol/sdk/types.js';

export default function createPrompt() {
  const config = {
    description: 'Reference guide for Outlook query syntax',
  };

  const handler = async (_args: { [x: string]: unknown }, _extra: RequestHandlerExtra<ServerRequest, ServerNotification>) => {
    return {
      messages: [
        {
          role: 'user' as const,
          content: {
            type: 'text' as const,
            text: `# Outlook Query Syntax Reference

## Logical Operators
- \`$and\`: Array of conditions that ALL must match
- \`$or\`: Array of conditions where ANY must match
- \`$not\`: Condition that must NOT match

## Email Address Fields
- \`from\`, \`to\`, \`cc\`, \`bcc\`: String or field operators

## Content Fields
- \`subject\`: Search subject line
- \`body\`: Search message body
- \`text\`: Search all text content
- \`exactPhrase\`: Strict exact phrase matching (KQL); wrap the phrase in quotes to keep the words together.

## Search Semantics
- Including fields such as subject, text, body, exactPhrase, or kqlQuery runs a Microsoft Graph $search. The structured filters (from, to, cc, bcc, categories, label, hasAttachment, etc.) are then applied client-side after the matching messages arrive, so those rules run once the search hits are returned.

## Boolean Flags
- \`hasAttachment\`: true/false
- \`isRead\`: true/false

## Date Range
\`\`\`json
{ "date": { "$gte": "2024-01-01", "$lt": "2024-12-31" } }
\`\`\`

## Outlook-Specific
- \`categories\`: work, personal, family, travel, important, urgent
- \`label\`: User categories (case-sensitive, use outlook-categories-list to discover)
- \`importance\`: high, normal, low
- \`kqlQuery\`: Escape hatch for advanced KQL syntax

## Field Operators (for multi-value fields)
- \`$any\`: OR - matches if ANY value matches
- \`$all\`: AND - matches if ALL values match
- \`$none\`: NOT - matches if NONE match

## Example Queries
\`\`\`json
// Unread from specific sender
{ "from": "boss@company.com", "isRead": false }

// High importance with attachment
{ "importance": "high", "hasAttachment": true }

// Multiple senders
{ "from": { "$any": ["alice@example.com", "bob@example.com"] } }

// Complex: work OR important category, unread
{ "$and": [
  { "categories": { "$any": ["work", "important"] } },
  { "isRead": false }
]}
\`\`\``,
          },
        },
      ],
    };
  };

  return {
    name: 'query-syntax',
    config,
    handler,
  } satisfies PromptModule;
}
