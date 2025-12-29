import { EMAIL_COMMON_PATTERNS, EMAIL_FIELD_DESCRIPTIONS, EMAIL_FIELDS, EmailContentTypeSchema, EmailDetailSchema, ExcludeThreadHistorySchema, extractCurrentMessageFromHtml, extractCurrentMessageFromHtmlToText, normalizeDateToISO as toIsoUtc } from '@mcp-z/email';
import type { EnrichedExtra } from '@mcp-z/oauth-microsoft';
import { schemas } from '@mcp-z/oauth-microsoft';

const { AuthRequiredBranchSchema } = schemas;

import { createFieldsSchema, filterFields, parseFields, type ToolModule } from '@mcp-z/server';
import { Client } from '@microsoft/microsoft-graph-client';
import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { type CallToolResult, ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { z } from 'zod';

const inputSchema = z.object({
  id: z.coerce.string().trim().min(1).describe('Outlook message ID to retrieve'),
  fields: createFieldsSchema({
    availableFields: EMAIL_FIELDS,
    fieldDescriptions: EMAIL_FIELD_DESCRIPTIONS,
    commonPatterns: EMAIL_COMMON_PATTERNS,
    resourceName: 'email message',
  }),
  contentType: EmailContentTypeSchema,
  excludeThreadHistory: ExcludeThreadHistorySchema,
});

// Success branch schema - uses item: wrapper for consistency with standard vocabulary
const successBranchSchema = z.object({
  type: z.literal('success'),
  item: EmailDetailSchema,
});

// Output schema with auth_required support
const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Get an Outlook message by ID with flexible field selection',
  inputSchema: inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ id, fields, contentType, excludeThreadHistory }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;

  const requestedFields = parseFields(fields, EMAIL_FIELDS);

  logger.info('outlook.message.get called', { id, fields: fields || 'all' });

  if (!id) {
    logger.info('outlook.message.get missing id');
    throw new McpError(ErrorCode.InvalidParams, 'Missing id');
  }

  try {
    const graph = Client.initWithMiddleware({ authProvider: extra.authContext.auth });
    const message = (await graph.api(`/me/messages/${encodeURIComponent(id)}`).get()) as MicrosoftGraph.Message;

    // Process body based on contentType and excludeThreadHistory options
    let bodyContent = message.body?.content ?? '';
    const isHtml = message.body?.contentType?.toLowerCase() === 'html';

    if (isHtml && excludeThreadHistory) {
      // Remove thread history from HTML
      bodyContent = extractCurrentMessageFromHtml(bodyContent);
    }

    if (isHtml && contentType === 'text') {
      // Convert HTML to plain text (thread extraction happens inside if excludeThreadHistory was true)
      bodyContent = excludeThreadHistory ? extractCurrentMessageFromHtmlToText(bodyContent) : extractCurrentMessageFromHtmlToText(message.body?.content ?? '');
    }

    // Build full message data
    const payload = {
      id: message.id,
      subject: message.subject,
      from: message.from?.emailAddress?.address,
      date: toIsoUtc(message.receivedDateTime) || message.receivedDateTime,
      snippet: message.bodyPreview,
      body: bodyContent,
      bodyContentType: contentType,
    };

    // Filter based on requested fields
    const filteredPayload = filterFields(payload, requestedFields);

    logger.info('outlook.message.get success', { id: payload.id, subject: payload.subject, from: payload.from, date: payload.date, fields: fields || 'all' });

    const result: Output = {
      type: 'success' as const,
      item: filteredPayload as z.infer<typeof EmailDetailSchema>,
    };

    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(result),
        },
      ],
      structuredContent: { result },
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error('outlook.message.get error', { error: message });

    throw new McpError(ErrorCode.InternalError, `Error getting message: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'message-get',
    config,
    handler,
  } satisfies ToolModule;
}
