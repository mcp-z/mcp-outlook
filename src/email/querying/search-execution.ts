import type { Client } from '@microsoft/microsoft-graph-client';
import type { Message } from '@microsoft/microsoft-graph-types';
import type { OutlookQuery as QueryNode } from '../../schemas/outlook-query-schema.js';
import type { Logger } from '../../types.js';
import { toGraphFilter } from './query-builder.js';

export interface OutlookSearchOptions {
  query?: QueryNode;
  limit?: number;
  includeBody?: boolean;
  maxPages?: number | null;
  pageSize?: number;
  pageToken?: string | undefined;
  logger: Logger;
}

export interface OutlookSearchResult {
  messages: Message[];
  nextPageToken?: string | undefined;
}

function buildGraphQuery(query: QueryNode | undefined): {
  search?: string | undefined;
  filter?: string | undefined;
} {
  if (!query) return {};
  const r = toGraphFilter(query) || { search: undefined, filter: undefined };
  return { search: r.search ?? undefined, filter: r.filter ?? undefined };
}

/**
 * Execute a direct Microsoft Graph API query.
 * Single attempt with no fallbacks or retries.
 * Errors are returned directly for actionable feedback.
 * Returns raw Graph API Message objects for transformation by caller.
 */
export async function searchMessages(graph: Client, opts: OutlookSearchOptions): Promise<OutlookSearchResult> {
  const { query, includeBody = false, pageSize = 50, pageToken, logger } = opts;

  const { search, filter } = buildGraphQuery(query);

  let request = graph.api('/me/messages');

  if (search) {
    request = request.search(search);
  }

  if (filter) {
    request = request.filter(filter);
  }

  request = request.top(Math.min(pageSize, 1000)); // Graph API max is 1000

  const selectFields = includeBody ? ['id', 'conversationId', 'receivedDateTime', 'subject', 'bodyPreview', 'body', 'from', 'toRecipients', 'ccRecipients', 'bccRecipients'] : ['id', 'conversationId', 'receivedDateTime', 'subject', 'bodyPreview', 'from', 'toRecipients', 'ccRecipients', 'bccRecipients'];
  request = request.select(selectFields.join(','));

  if (pageToken && pageToken.trim().length > 0) {
    request = request.skipToken(pageToken);
  }

  logger.info('Executing direct Outlook query', {
    hasSearch: !!search,
    hasFilter: !!filter,
    pageSize,
    hasPageToken: !!pageToken,
  });

  try {
    // Execute the request
    const response = await request.get();

    // Extract messages and nextLink
    const messages: Message[] = response.value || [];
    const nextLink: string | undefined = response['@odata.nextLink'];

    // Extract skipToken from nextLink if present
    let nextPageToken: string | undefined;
    if (nextLink) {
      try {
        const url = new URL(nextLink);
        nextPageToken = url.searchParams.get('$skiptoken') || undefined;
      } catch {
        logger.info('Failed to parse nextLink for skipToken', { nextLink });
      }
    }

    logger.info(`Query succeeded with ${messages.length} results`, {
      hasNextPage: !!nextPageToken,
    });

    return {
      messages,
      nextPageToken,
    };
  } catch (error) {
    // Log and re-throw error directly - no fallback logic
    logger.error('Outlook query failed', {
      error: error instanceof Error ? error.message : String(error),
      hasSearch: !!search,
      hasFilter: !!filter,
    });
    throw error;
  }
}
