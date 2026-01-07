import type { Client } from '@microsoft/microsoft-graph-client';
import type { Message } from '@microsoft/microsoft-graph-types';
import type { OutlookQuery as QueryNode } from '../../schemas/outlook-query-schema.ts';
import type { Logger } from '../../types.ts';
import { buildClientPredicate } from './client-filter.ts';
import { decodeNextPageToken, encodeNextPageToken, FIRST_PAGE_SKIP_TOKEN } from './pagination-token.ts';
import { buildQueryFilter, toGraphFilter } from './query-builder.ts';

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

  const searchResult = await collectClientFilteredPages({
    graph,
    search: graphPlan.search ?? undefined,
    filter: undefined,
    startSkipToken: pageState.skipToken,
    startOffset: pageState.offset,
    predicate,
    includeBody: effectiveIncludeBody,
    pageSize,
    logger,
    maxPages: maxPages ?? null,
    maxItemsScanned: maxItemsScanned ?? null,
    mode: 'search',
  });

  if (query && 'text' in query && typeof query.text === 'object' && Array.isArray(query.text.$all)) {
    console.log(
      'SEARCH RESULT IDS',
      searchResult.collected.map((msg) => msg.id)
    );
  }

  if (searchResult.capHit) {
    return { messages: searchResult.collected };
  }
  if (searchResult.midPageToken) {
    return { messages: searchResult.collected, nextPageToken: searchResult.midPageToken };
  }
  if (searchResult.nextSkipToken) {
    return {
      messages: searchResult.collected,
      nextPageToken: encodeNextPageToken({ skipToken: searchResult.nextSkipToken, offset: 0, mode: 'search' }),
    };
  }
  if (searchResult.collected.length > 0) {
    return { messages: searchResult.collected };
  }

  if (predicate && query) {
    const fallbackFilter = buildQueryFilter(query).trim();
    if (fallbackFilter) {
      const fallbackResult = await collectClientFilteredPages({
        graph,
        filter: fallbackFilter,
        includeBody: effectiveIncludeBody,
        pageSize,
        predicate,
        logger,
        maxPages: maxPages ?? null,
        maxItemsScanned: maxItemsScanned ?? null,
        startSkipToken: undefined,
        startOffset: 0,
        mode: 'filter',
      });

      if (fallbackResult.capHit) {
        return { messages: fallbackResult.collected };
      }
      if (fallbackResult.midPageToken) {
        return { messages: fallbackResult.collected, nextPageToken: fallbackResult.midPageToken };
      }
      if (fallbackResult.nextSkipToken) {
        return {
          messages: fallbackResult.collected,
          nextPageToken: encodeNextPageToken({ skipToken: fallbackResult.nextSkipToken, offset: 0, mode: 'filter' }),
        };
      }
      if (fallbackResult.collected.length > 0) {
        return { messages: fallbackResult.collected };
      }
    }
  }

  return { messages: searchResult.collected };
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

interface CollectParams {
  graph: Client;
  search?: string;
  filter?: string;
  startSkipToken?: string;
  startOffset?: number;
  predicate: ((message: Message) => boolean) | null;
  includeBody: boolean;
  pageSize: number;
  logger: Logger;
  maxPages: number | null;
  maxItemsScanned: number | null;
  mode: 'search' | 'filter';
}

interface CollectResult {
  collected: Message[];
  midPageToken?: string;
  nextSkipToken?: string;
  capHit: boolean;
}

async function collectClientFilteredPages(params: CollectParams): Promise<CollectResult> {
  const { graph, search, filter, startSkipToken, startOffset = 0, predicate, includeBody, pageSize, logger, maxPages, maxItemsScanned, mode } = params;
  let currentSkipToken = startSkipToken;
  let offset = startOffset;
  const collected: Message[] = [];
  let midPageToken: string | undefined;
  let lastNextSkipToken: string | undefined;
  let capHit = false;
  let pagesFetched = 0;
  let scannedItems = 0;

  mainLoop: while (collected.length < pageSize) {
    if (maxPages !== null && pagesFetched >= maxPages) {
      capHit = true;
      break;
    }

    const pageTokenUsed = currentSkipToken;
    const graphPage = await fetchGraphPage(graph, {
      search,
      filter,
      skipToken: currentSkipToken,
      includeBody,
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
      if (maxItemsScanned !== null && scannedItems > maxItemsScanned) {
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
          mode,
        });
        break mainLoop;
      }
    }

    currentSkipToken = graphPage.nextSkipToken;
    if (!currentSkipToken) {
      break;
    }
  }

  return {
    collected,
    midPageToken,
    nextSkipToken: lastNextSkipToken,
    capHit,
  };
}
