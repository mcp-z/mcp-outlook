import type { ExecutionResult } from '@mcp-z/email';
import type { Client } from '@microsoft/microsoft-graph-client';
import type { OutlookQuery as QueryNode } from '../../schemas/outlook-query-schema.js';
import type { Logger } from '../../types.js';
import { searchMessages } from './search-execution.js';

export interface ExecuteQueryOptions {
  readonly logger: Logger;
  readonly limit?: number;
  readonly includeBody?: boolean;
  readonly maxPages?: number | null;
  readonly pageSize?: number;
  readonly pageToken?: string;
  readonly timezone?: string;
  readonly abortSignal?: AbortSignal;
}

/**
 * Execute an Outlook query with direct, single-attempt execution.
 * No planning, no fallbacks, no retries.
 * Microsoft Graph errors are returned directly to the caller for actionable feedback.
 */
export async function executeQuery<T>(graph: Client, query: QueryNode | undefined, options: ExecuteQueryOptions, transform: (item: unknown) => T): Promise<ExecutionResult<T>> {
  const { logger, limit, includeBody, maxPages, pageSize, pageToken } = options;

  // Single execution - direct query to Microsoft Graph API
  logger.info('executeQuery: executing direct Outlook query');

  try {
    const result = await searchMessages(graph, {
      ...(query !== undefined && { query }),
      limit: limit ?? 50,
      includeBody: includeBody ?? false,
      maxPages: maxPages ?? null,
      pageSize: pageSize ?? 50,
      pageToken: pageToken ?? undefined,
      logger,
    });

    // Transform directly - if Microsoft Graph returns invalid data, fail loud
    const transformedResults = result.messages.map((item, index) => {
      try {
        return transform(item);
      } catch (transformError) {
        logger.error(`Transform function failed for item ${index}`, {
          itemIndex: index,
          error: transformError instanceof Error ? transformError.message : String(transformError),
        });
        throw new Error(`Transform failed at item ${index}: ${transformError}`);
      }
    });

    logger.info(`executeQuery: succeeded with ${transformedResults.length} results`);

    return {
      success: true,
      items: transformedResults,
      metadata: {
        nextPageToken: result.nextPageToken,
        totalFetched: result.messages.length,
      },
    };
  } catch (error) {
    // Re-throw errors directly - no fallback logic
    logger.error('executeQuery: failed', error as Record<string, unknown>);
    throw error;
  }
}
