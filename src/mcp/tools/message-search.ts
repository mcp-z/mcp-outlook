import { EMAIL_COMMON_PATTERNS, EMAIL_FIELD_DESCRIPTIONS, EMAIL_FIELDS, EmailContentTypeSchema, type EmailDetail, EmailDetailSchema, ExcludeThreadHistorySchema, extractCurrentMessageFromHtml, extractCurrentMessageFromHtmlToText } from '@mcp-z/email';
import type { EnrichedExtra } from '@mcp-z/oauth-microsoft';
import { schemas } from '@mcp-z/oauth-microsoft';

const { AuthRequiredBranchSchema } = schemas;

import { createFieldsSchema, createPaginationSchema, createShapeSchema, filterFields, parseFields, type ToolModule, toColumnarFormat } from '@mcp-z/server';
import { Client } from '@microsoft/microsoft-graph-client';
import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { type CallToolResult, ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { z } from 'zod';
import { executeQuery as executeOutlookQuery } from '../../email/querying/execute-query.js';
import { OutlookQuerySchema } from '../../schemas/outlook-query-schema.js';

const inputSchema = z.object({
  query: OutlookQuerySchema.optional().describe('Structured query object for filtering messages. Use query-syntax prompt for reference.'),
  fields: createFieldsSchema({
    availableFields: EMAIL_FIELDS,
    fieldDescriptions: EMAIL_FIELD_DESCRIPTIONS,
    commonPatterns: EMAIL_COMMON_PATTERNS,
    resourceName: 'email message',
  }),
  ...createPaginationSchema({
    defaultPageSize: 50,
    maxPageSize: 500,
    provider: 'outlook',
  }).shape,
  shape: createShapeSchema(),
  contentType: EmailContentTypeSchema,
  excludeThreadHistory: ExcludeThreadHistorySchema,
});

// Success branch schemas for different shapes
const successObjectsBranchSchema = z.object({
  type: z.literal('success'),
  shape: z.literal('objects'),
  items: z.array(EmailDetailSchema).describe('Matching email messages'),
  nextPageToken: z.string().optional().describe('Token for fetching next page of results'),
});

const successArraysBranchSchema = z.object({
  type: z.literal('success'),
  shape: z.literal('arrays'),
  columns: z.array(z.string()).describe('Column names in canonical order'),
  rows: z.array(z.array(z.unknown())).describe('Row data matching column order'),
  nextPageToken: z.string().optional().describe('Token for fetching next page of results'),
});

// Output schema with auth_required support
// Using z.union instead of discriminatedUnion since we have two success branches with different shapes
const outputSchema = z.union([successObjectsBranchSchema, successArraysBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Search Outlook messages using structured query objects with flexible field selection.',
  inputSchema: inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ query, pageSize = 50, pageToken, fields, shape = 'arrays', contentType, excludeThreadHistory }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;

  const requestedFields = parseFields(fields, EMAIL_FIELDS);
  const includeBody = requestedFields === 'all' || requestedFields.has('body');

  logger.info('outlook.search called with', {
    query,
    pageSize,
    pageToken: pageToken ? '[provided]' : undefined,
    fields: fields || 'all',
    includeBody,
  });

  try {
    const graph = Client.initWithMiddleware({
      authProvider: extra.authContext.auth,
    });
    const started = Date.now();
    const exec = await executeOutlookQuery(
      graph,
      query,
      {
        logger,
        pageSize,
        ...(pageToken ? { pageToken } : {}),
        includeBody,
        limit: pageSize,
      },
      (m: unknown) => {
        const message = m as MicrosoftGraph.Message;
        const to = Array.isArray(message?.toRecipients)
          ? message.toRecipients
              .map((r: MicrosoftGraph.Recipient) => r?.emailAddress?.address)
              .filter(Boolean)
              .join(', ')
          : undefined;
        const cc = Array.isArray(message?.ccRecipients)
          ? message.ccRecipients
              .map((r: MicrosoftGraph.Recipient) => r?.emailAddress?.address)
              .filter(Boolean)
              .join(', ')
          : undefined;
        const bcc = Array.isArray(message?.bccRecipients)
          ? message.bccRecipients
              .map((r: MicrosoftGraph.Recipient) => r?.emailAddress?.address)
              .filter(Boolean)
              .join(', ')
          : undefined;
        const fromAddr = message?.from?.emailAddress?.address ?? null;
        const fromName = message?.from?.emailAddress?.name ?? null;
        // Process body based on contentType and excludeThreadHistory options
        let bodyContent: string | undefined;
        if (includeBody && message?.body?.content) {
          const isHtml = message.body.contentType?.toLowerCase() === 'html';
          bodyContent = message.body.content;

          if (isHtml && excludeThreadHistory) {
            bodyContent = extractCurrentMessageFromHtml(bodyContent);
          }

          if (isHtml && contentType === 'text') {
            bodyContent = excludeThreadHistory ? extractCurrentMessageFromHtmlToText(bodyContent) : extractCurrentMessageFromHtmlToText(message.body.content);
          }
        }

        const mapped: Partial<EmailDetail> = {
          id: message?.id ?? undefined,
          subject: message?.subject ?? undefined,
          from: fromAddr ?? undefined,
          fromName: fromName ?? undefined,
          to,
          cc,
          bcc,
          date: message?.receivedDateTime ?? undefined,
          snippet: message?.bodyPreview ?? undefined,
          ...(bodyContent && { body: bodyContent, bodyContentType: contentType }),
        };
        return mapped;
      }
    );
    const durationMs = Date.now() - started;

    const filteredItems = exec.items.map((item) => filterFields(item, requestedFields));

    logger.info('outlook.message.search returning', {
      query,
      pageSize,
      itemsLength: filteredItems.length,
      fields: fields || 'all',
    });
    logger.info('outlook.message.search metrics', {
      durationMs,
      metadata: exec.metadata,
    });

    // Type-safe metadata access
    const metadataObj = exec.metadata as { nextPageToken?: string } | undefined;

    // Build result based on shape
    const result: Output =
      shape === 'arrays'
        ? {
            type: 'success' as const,
            shape: 'arrays' as const,
            ...toColumnarFormat(filteredItems, requestedFields, EMAIL_FIELDS),
            ...(metadataObj?.nextPageToken && { nextPageToken: metadataObj.nextPageToken }),
          }
        : {
            type: 'success' as const,
            shape: 'objects' as const,
            items: filteredItems,
            ...(metadataObj?.nextPageToken && { nextPageToken: metadataObj.nextPageToken }),
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
    logger.error('outlook.message.search error', { error: message });

    throw new McpError(ErrorCode.InternalError, `Error searching messages: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'message-search',
    config,
    handler,
  } satisfies ToolModule;
}
