/** Outlook message CSV export tool - streams results to file without loading all data into context */

import { EmailContentTypeSchema, ExcludeThreadHistorySchema, extractCurrentMessageFromHtml, extractCurrentMessageFromHtmlToText } from '@mcp-z/email';
import type { EnrichedExtra } from '@mcp-z/oauth-microsoft';
import { schemas } from '@mcp-z/oauth-microsoft';

const { AuthRequiredBranchSchema } = schemas;

import { getFileUri, reserveFile, type ToolModule } from '@mcp-z/server';
import { Client } from '@microsoft/microsoft-graph-client';
import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { type CallToolResult, ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { stringify } from 'csv-stringify/sync';
import { createWriteStream } from 'fs';
import { unlink } from 'fs/promises';
import { z } from 'zod';
import { executeQuery as executeOutlookQuery } from '../../email/querying/execute-query.ts';
import { OutlookQuerySchema } from '../../schemas/outlook-query-schema.ts';
import type { StorageExtra } from '../../types.ts';

const DEFAULT_PAGE_SIZE = 50;
const DEFAULT_MAX_ITEMS = 10000;
const MAX_EXPORT_ITEMS = 50000;

const ExportResultSchema = z.object({
  uri: z.string().describe('File URI (file:// or http://)'),
  filename: z.string().describe('Stored filename'),
  rowCount: z.number().describe('Number of messages exported'),
  truncated: z.boolean().describe('Whether export was truncated at maxItems'),
});

const inputSchema = z.object({
  query: OutlookQuerySchema.optional().describe('Structured query object for filtering messages. Use query-syntax prompt for reference.'),
  maxItems: z.number().int().positive().max(MAX_EXPORT_ITEMS).default(DEFAULT_MAX_ITEMS).describe(`Maximum messages to export (default: ${DEFAULT_MAX_ITEMS}, max: ${MAX_EXPORT_ITEMS})`),
  filename: z.string().trim().min(1).default('outlook-messages.csv').describe('Output filename (default: outlook-messages.csv)'),
  contentType: EmailContentTypeSchema,
  excludeThreadHistory: ExcludeThreadHistorySchema,
});

// Success branch schema
const successBranchSchema = ExportResultSchema.extend({
  type: z.literal('success'),
});

// Output schema with auth_required support
const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Export Outlook messages to CSV with streaming pagination. Returns file URI. Use query-syntax prompt for query reference.',
  inputSchema: inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

/**
 * Handler for outlook-messages-export-csv tool
 *
 * CRITICAL: Streaming implementation per user requirements
 * - Generate UUID upfront
 * - Write CSV header immediately
 * - Append rows as batches arrive
 * - Delete partial file on error
 * - NO RETRIES (fail fast on error)
 */
async function handler({ query, maxItems, filename, contentType, excludeThreadHistory }: Input, extra: EnrichedExtra & StorageExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  const { storageContext } = extra;
  const { transport, resourceStoreUri, baseUrl } = storageContext;

  logger.info('outlook.messages.export-csv called', {
    query,
    maxItems,
    filename,
    accountId: extra.authContext.accountId,
  });

  // Reserve file location for streaming write (creates directory, generates ID, formats filename)
  const reservation = await reserveFile(filename, {
    resourceStoreUri,
  });
  const { storedName, fullPath } = reservation;

  logger.info('outlook.messages.export-csv starting streaming export', { path: fullPath, maxItems });

  try {
    const graph = Client.initWithMiddleware({ authProvider: extra.authContext.auth });

    // Create CSV headers (all email fields)
    const csvHeaders = ['id', 'threadId', 'from', 'to', 'cc', 'bcc', 'subject', 'date', 'snippet', 'body', 'provider', 'labels'];

    // Create write stream and write headers immediately
    const writeStream = createWriteStream(fullPath, { encoding: 'utf-8' });
    const headerLine = stringify([csvHeaders], { header: false, quoted: true, quote: '"', escape: '"' });
    writeStream.write(headerLine);

    // Internal pagination loop - append to CSV with each batch
    // NO RETRIES: If any error occurs, fail the whole operation and clean up
    let totalRows = 0;
    let nextPageToken: string | undefined;
    const started = Date.now();

    while (totalRows < maxItems) {
      const remainingItems = maxItems - totalRows;
      const pageSize = Math.min(remainingItems, DEFAULT_PAGE_SIZE);

      const exec: {
        items: Array<{
          id: string;
          threadId?: string;
          from: string;
          to: string;
          cc: string;
          bcc: string;
          subject: string;
          date: string;
          snippet: string;
          body: string;
          provider: string;
          labels: string;
        }>;
        metadata?: { nextPageToken?: string };
      } = await executeOutlookQuery(
        graph,
        query,
        {
          logger,
          pageSize,
          ...(nextPageToken !== undefined && { pageToken: nextPageToken }),
          includeBody: true, // Always include body for CSV export
          limit: pageSize,
        },
        (m: unknown) => {
          const message = m as MicrosoftGraph.Message;
          const to = Array.isArray(message?.toRecipients)
            ? message.toRecipients
                .map((r: MicrosoftGraph.Recipient) => r?.emailAddress?.address ?? (r as { address?: string })?.address)
                .filter(Boolean)
                .join(', ')
            : '';
          const cc = Array.isArray(message?.ccRecipients)
            ? message.ccRecipients
                .map((r: MicrosoftGraph.Recipient) => r?.emailAddress?.address ?? (r as { address?: string })?.address)
                .filter(Boolean)
                .join(', ')
            : '';
          const bcc = Array.isArray(message?.bccRecipients)
            ? message.bccRecipients
                .map((r: MicrosoftGraph.Recipient) => r?.emailAddress?.address ?? (r as { address?: string })?.address)
                .filter(Boolean)
                .join(', ')
            : '';
          const fromAddr = message?.from?.emailAddress?.address ?? (message?.from as { address?: string })?.address ?? '';
          const categories = Array.isArray(message?.categories) ? message.categories.join(';') : '';

          // Process body based on contentType and excludeThreadHistory options
          let body = message?.body?.content ?? '';
          const isHtml = message?.body?.contentType?.toLowerCase() === 'html';

          if (isHtml && excludeThreadHistory) {
            body = extractCurrentMessageFromHtml(body);
          }

          if (isHtml && contentType === 'text') {
            body = excludeThreadHistory ? extractCurrentMessageFromHtmlToText(body) : extractCurrentMessageFromHtmlToText(message?.body?.content ?? '');
          }

          return {
            id: String(message?.id ?? ''),
            threadId: message?.conversationId ? String(message.conversationId) : '',
            from: fromAddr,
            to,
            cc,
            bcc,
            subject: message?.subject ?? '',
            date: message?.receivedDateTime ?? '',
            snippet: message?.bodyPreview ?? '',
            body,
            provider: 'outlook' as const,
            labels: categories,
          };
        }
      );

      const csvRows = exec.items.map((item) => {
        return [item.id, item.threadId, item.from, item.to, item.cc, item.bcc, item.subject, item.date, item.snippet, item.body, item.provider, item.labels];
      });

      // Append rows to CSV file immediately
      if (csvRows.length > 0) {
        const rowsContent = stringify(csvRows, { header: false, quoted: true, quote: '"', escape: '"' });
        writeStream.write(rowsContent);
      }

      totalRows += exec.items.length;
      nextPageToken = exec.metadata?.nextPageToken;

      logger.info('outlook.messages.export-csv batch written', {
        batchSize: exec.items.length,
        totalRows,
        hasMore: Boolean(nextPageToken),
      });

      // Exit if no more results or reached maxItems
      if (!nextPageToken || exec.items.length === 0) {
        break;
      }
    }

    // Close write stream
    await new Promise<void>((resolve, reject) => {
      writeStream.end(() => resolve());
      writeStream.on('error', reject);
    });

    const durationMs = Date.now() - started;
    const truncated = totalRows >= maxItems && Boolean(nextPageToken);

    logger.info('outlook.messages.export-csv completed', {
      rowCount: totalRows,
      truncated,
      durationMs,
      filename: storedName,
    });

    // Generate URI based on transport type (stdio: file://, HTTP: http://)
    const uri = getFileUri(storedName, transport, {
      resourceStoreUri,
      ...(baseUrl && { baseUrl }),
      endpoint: '/files',
    });

    const result: Output = {
      type: 'success' as const,
      uri,
      filename: storedName,
      rowCount: totalRows,
      truncated,
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
    // CRITICAL: Clean up partial CSV file on error
    try {
      await unlink(fullPath);
      logger.debug('Cleaned up partial CSV file after error', { path: fullPath });
    } catch (_cleanupError) {
      logger.debug('Could not clean up CSV file (may not exist)', { path: fullPath });
    }

    const message = error instanceof Error ? error.message : String(error);
    logger.error('outlook.messages.export-csv error', { error: message });

    throw new McpError(ErrorCode.InternalError, `Error exporting messages to CSV: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'messages-export-csv',
    config,
    handler,
  } satisfies ToolModule;
}
