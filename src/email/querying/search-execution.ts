import type { Client } from '@microsoft/microsoft-graph-client';
import type { Message } from '@microsoft/microsoft-graph-types';
import type { OutlookQuery as QueryNode } from '../../schemas/outlook-query-schema.ts';
import type { Logger } from '../../types.ts';
import { buildClientPredicate } from './client-filter.ts';
import { decodeNextPageToken, encodeNextPageToken, FIRST_PAGE_SKIP_TOKEN } from './pagination-token.ts';
import { toGraphFilter } from './query-builder.ts';

export interface OutlookSearchOptions {
  query?: QueryNode;
  limit?: number;
  includeBody?: boolean;
  maxPages?: number | null;
  maxItemsScanned?: number | null;
  pageSize?: number;
  pageToken?: string | undefined;
  logger: Logger;
}

export interface OutlookSearchResult {
  messages: Message[];
  nextPageToken?: string | undefined;
}

export async function searchMessages(graph: Client, opts: OutlookSearchOptions): Promise<OutlookSearchResult> {
  const { query, includeBody = false, pageSize = 50, pageToken, maxPages = null, maxItemsScanned = null, logger } = opts;

  const graphPlan = query
    ? toGraphFilter(query)
    : {
        search: null,
        filter: null,
        requireBodyClientFilter: false,
        hasFullText: false,
      };
  const pageState = decodeNextPageToken(pageToken);
  const useClientFiltering = Boolean(query && graphPlan.hasFullText);
  const predicate = useClientFiltering && query ? buildClientPredicate(query) : null;
  const requireBody = graphPlan.requireBodyClientFilter;
  const effectiveIncludeBody = includeBody || (requireBody && !includeBody);

  logger.info('Executing Outlook search', {
    hasSearch: !!graphPlan.search,
    hasFilter: !!graphPlan.filter,
    useClientFiltering,
    pageSize,
    includeBody: effectiveIncludeBody,
    pageToken: pageToken ? '[provided]' : undefined,
    forcedBody: requireBody && !includeBody,
  });

  if (!useClientFiltering) {
    const result = await fetchGraphPage(graph, {
      search: graphPlan.search ?? undefined,
      filter: graphPlan.filter ?? undefined,
      skipToken: pageState.skipToken,
      includeBody: effectiveIncludeBody,
      pageSize,
      logger,
    });
    return {
      messages: result.messages,
      nextPageToken: result.nextSkipToken,
    };
  }

  let currentSkipToken = pageState.skipToken;
  let offset = pageState.offset;
  const maxPagesCap = maxPages ?? null;
  const maxItemsCap = maxItemsScanned ?? null;
  const collected: Message[] = [];
  let lastNextSkipToken: string | undefined;
  let midPageToken: string | undefined;
  let capHit = false;
  let pagesFetched = 0;
  let scannedItems = 0;

  mainLoop: while (collected.length < pageSize) {
    if (maxPagesCap !== null && pagesFetched >= maxPagesCap) {
      capHit = true;
      break;
    }

    const pageTokenUsed = currentSkipToken;
    const graphPage = await fetchGraphPage(graph, {
      search: graphPlan.search ?? undefined,
      skipToken: pageTokenUsed,
      includeBody: effectiveIncludeBody,
      pageSize,
      logger,
    });

    pagesFetched += 1;
    lastNextSkipToken = graphPage.nextSkipToken;
    const pageItems = graphPage.messages;
    if (!pageItems.length) {
      break;
    }

    const startIndex = offset;
    const sliced = startIndex > 0 ? pageItems.slice(startIndex) : pageItems;
    offset = 0;

    for (let idx = 0; idx < sliced.length; idx += 1) {
      const message = sliced[idx];
      scannedItems += 1;
      if (maxItemsCap !== null && scannedItems > maxItemsCap) {
        capHit = true;
        break mainLoop;
      }
      if (predicate && !predicate(message)) {
        continue;
      }
      collected.push(message);
      if (collected.length === pageSize) {
        const nextOffset = startIndex + idx + 1;
        midPageToken = encodeNextPageToken({
          skipToken: pageTokenUsed,
          offset: nextOffset,
          mode: 'search',
        });
        break mainLoop;
      }
    }

    currentSkipToken = graphPage.nextSkipToken;
    if (!currentSkipToken) {
      break;
    }
  }

  if (capHit) {
    return {
      messages: collected,
    };
  }

  if (midPageToken) {
    return {
      messages: collected,
      nextPageToken: midPageToken,
    };
  }

  if (lastNextSkipToken) {
    return {
      messages: collected,
      nextPageToken: encodeNextPageToken({ skipToken: lastNextSkipToken, offset: 0, mode: 'search' }),
    };
  }

  return {
    messages: collected,
  };
}

interface GraphPageParams {
  search?: string;
  filter?: string;
  skipToken?: string;
  includeBody: boolean;
  pageSize: number;
  logger: Logger;
}

async function fetchGraphPage(graph: Client, params: GraphPageParams): Promise<{ messages: Message[]; nextSkipToken?: string }> {
  const { search, filter, skipToken, includeBody, pageSize, logger } = params;
  let request = graph.api('/me/messages');

  if (search) {
    request = request.search(search);
  }
  if (filter) {
    request = request.filter(filter);
  }

  const top = Math.min(pageSize, 1000);
  request = request.top(top);

  const selectFields = includeBody ? ['id', 'conversationId', 'receivedDateTime', 'subject', 'bodyPreview', 'body', 'from', 'toRecipients', 'ccRecipients', 'bccRecipients'] : ['id', 'conversationId', 'receivedDateTime', 'subject', 'bodyPreview', 'from', 'toRecipients', 'ccRecipients', 'bccRecipients'];
  request = request.select(selectFields.join(','));

  if (skipToken && skipToken !== FIRST_PAGE_SKIP_TOKEN) {
    request = request.skipToken(skipToken);
  }

  logger.info('Fetching Outlook Graph page', {
    search: !!search,
    filter: !!filter,
    skipToken: skipToken ? '[provided]' : undefined,
    pageSize: top,
  });

  const response = await request.get();
  const messages: Message[] = response.value || [];
  let nextSkipToken: string | undefined;
  if (response['@odata.nextLink']) {
    try {
      const nextLinkUrl = new URL(response['@odata.nextLink']);
      nextSkipToken = nextLinkUrl.searchParams.get('$skiptoken') || undefined;
    } catch {
      logger.info('Failed to parse nextLink skipToken', { nextLink: response['@odata.nextLink'] });
    }
  }

  logger.info(`Graph page returned ${messages.length} messages`, {
    hasNextPage: !!nextSkipToken,
  });

  return { messages, nextSkipToken };
}
