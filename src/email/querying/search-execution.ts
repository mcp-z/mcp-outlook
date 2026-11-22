import type { Client, GraphRequest } from '@microsoft/microsoft-graph-client';
import type { Message } from '@microsoft/microsoft-graph-types';
import type { OutlookQuery as QueryNode } from '../../schemas/outlook-query-schema.ts';
import type { Logger } from '../../types.ts';
import { buildClientPredicate } from './client-filter.ts';
import { checkNextLink } from './next-link.ts';
import { decodeNextPageToken, encodeNextPageToken } from './pagination-token.ts';
import { midPageToken, resumeCollection } from './paging-decisions.ts';
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
      link: pageState.link,
      includeBody: effectiveIncludeBody,
      pageSize,
      logger,
    });
    return {
      messages: result.messages,
      ...(result.nextLink && { nextPageToken: encodeNextPageToken({ link: result.nextLink, offset: 0 }) }),
    };
  }

  const resumeFilter = pageToken !== undefined && resumeCollection(pageState.mode) === 'filter';
  const fallbackFilter = query ? buildQueryFilter(query).trim() : '';

  const searchResult = resumeFilter
    ? EMPTY_COLLECT
    : await collectClientFilteredPages({
        graph,
        search: graphPlan.search ?? undefined,
        filter: undefined,
        startLink: pageState.link,
        startOffset: pageState.offset,
        predicate,
        includeBody: effectiveIncludeBody,
        pageSize,
        logger,
        maxPages: maxPages ?? null,
        maxItemsScanned: maxItemsScanned ?? null,
        mode: 'search',
      });

  if (searchResult.capHit) {
    return { messages: searchResult.collected };
  }
  if (searchResult.midPageToken) {
    return { messages: searchResult.collected, nextPageToken: searchResult.midPageToken };
  }
  if (searchResult.nextLink) {
    return {
      messages: searchResult.collected,
      nextPageToken: encodeNextPageToken({ link: searchResult.nextLink, offset: 0, mode: 'search' }),
    };
  }
  if (searchResult.collected.length > 0) {
    return { messages: searchResult.collected };
  }

  if (predicate && query) {
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
        startLink: resumeFilter ? pageState.link : undefined,
        startOffset: resumeFilter ? pageState.offset : 0,
        mode: 'filter',
      });

      if (fallbackResult.capHit) {
        return { messages: fallbackResult.collected };
      }
      if (fallbackResult.midPageToken) {
        return { messages: fallbackResult.collected, nextPageToken: fallbackResult.midPageToken };
      }
      if (fallbackResult.nextLink) {
        return {
          messages: fallbackResult.collected,
          nextPageToken: encodeNextPageToken({ link: fallbackResult.nextLink, offset: 0, mode: 'filter' }),
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
  /** Path and query of a next-page link; when set it is requested as given and the other parameters are not used. */
  link?: string | undefined;
  includeBody: boolean;
  pageSize: number;
  logger: Logger;
}

async function fetchGraphPage(graph: Client, params: GraphPageParams): Promise<{ messages: Message[]; nextLink?: string }> {
  const { search, filter, link, includeBody, pageSize, logger } = params;
  let request: GraphRequest;

  if (link) {
    request = graph.api(link);
  } else {
    request = graph.api('/me/messages');
    if (search) {
      request = request.search(search);
    }
    if (filter) {
      request = request.filter(filter);
    }
    request = request.top(Math.min(pageSize, 1000));
    const selectFields = includeBody ? ['id', 'conversationId', 'receivedDateTime', 'subject', 'bodyPreview', 'body', 'from', 'toRecipients', 'ccRecipients', 'bccRecipients'] : ['id', 'conversationId', 'receivedDateTime', 'subject', 'bodyPreview', 'from', 'toRecipients', 'ccRecipients', 'bccRecipients'];
    request = request.select(selectFields.join(','));
  }

  logger.info('Fetching Outlook Graph page', {
    search: !!search,
    filter: !!filter,
    link: link ? '[provided]' : undefined,
    pageSize: Math.min(pageSize, 1000),
  });

  const response = await request.get();
  const messages: Message[] = response.value || [];
  const nextLink: string | undefined = response['@odata.nextLink'] ? checkNextLink(response['@odata.nextLink']) : undefined;

  logger.info(`Graph page returned ${messages.length} messages`, {
    hasNextPage: !!nextLink,
  });

  return { messages, ...(nextLink && { nextLink }) };
}

interface CollectParams {
  graph: Client;
  search?: string;
  filter?: string;
  startLink?: string | undefined;
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
  nextLink?: string;
  capHit: boolean;
}

const EMPTY_COLLECT: CollectResult = { collected: [], capHit: false };

async function collectClientFilteredPages(params: CollectParams): Promise<CollectResult> {
  const { graph, search, filter, startLink, startOffset = 0, predicate, includeBody, pageSize, logger, maxPages, maxItemsScanned, mode } = params;
  let currentLink = startLink;
  let offset = startOffset;
  const collected: Message[] = [];
  let midToken: string | undefined;
  let lastNextLink: string | undefined;
  let capHit = false;
  let pagesFetched = 0;
  let scannedItems = 0;

  mainLoop: while (collected.length < pageSize) {
    if (maxPages !== null && pagesFetched >= maxPages) {
      capHit = true;
      break;
    }

    const pageLinkUsed = currentLink;
    const graphPage = await fetchGraphPage(graph, {
      search,
      filter,
      link: currentLink,
      includeBody,
      pageSize,
      logger,
    });

    pagesFetched += 1;
    lastNextLink = graphPage.nextLink;
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
        midToken = midPageToken({ link: pageLinkUsed, pageStart: startIndex, index: idx, mode });
        break mainLoop;
      }
    }

    currentLink = graphPage.nextLink;
    if (!currentLink) {
      break;
    }
  }

  return {
    collected,
    midPageToken: midToken,
    nextLink: lastNextLink,
    capHit,
  };
}
