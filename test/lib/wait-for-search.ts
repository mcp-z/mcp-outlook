import type { Client } from '@microsoft/microsoft-graph-client';
import { setTimeout as delay } from 'timers/promises';
import { toGraphFilter } from '../../src/email/querying/query-builder.js';
import type { OutlookQuery } from '../../src/schemas/outlook-query-schema.js';

// Type for error objects that may have status/code properties
type ErrorWithStatus = {
  status?: number;
  statusCode?: number;
  code?: number | string;
};

// Type for message objects returned by Graph API
type GraphMessage = {
  id?: string;
  [key: string]: unknown;
};

interface WaitForSearchOptions {
  timeout?: number; // Timeout in ms (default: 10000)
  expectedId?: string; // Wait for specific message ID in results
  select?: string; // Fields to select (default: 'id,subject')
}

/**
 * Wait for Microsoft Graph message search results using OutlookQuery compilation with exponential backoff.
 * Uses the same query compilation path (OutlookQuery → toGraphFilter) that message-search tool uses, ensuring consistency.
 * Only retries on 404 and 5xx errors (transient failures).
 * Throws immediately on auth errors (401), rate limits (429), or client errors (4xx).
 *
 * @param client - Microsoft Graph client
 * @param query - OutlookQuery object (same format as message-search tool)
 * @param opts - Optional timeout, expectedId, and field selection
 * @returns Array of found messages
 * @throws Error on timeout or non-retryable errors
 */
export default async function waitForSearch(client: Client, query: OutlookQuery, opts?: WaitForSearchOptions): Promise<unknown[]> {
  const timeoutMs = opts?.timeout ?? 10000;
  const select = opts?.select ?? 'id,subject';
  const start = Date.now();
  let interval = 100;
  const maxInterval = 1000;

  // Compile OutlookQuery to Graph $search/$filter - same path as message-search tool
  const { search, filter } = toGraphFilter(query) || {};

  while (Date.now() - start < timeoutMs) {
    try {
      // Build request using compiled query - matches message-search tool behavior
      let request = client.api('/me/messages').select(select).top(50);

      if (search) {
        request = request.search(search);
      }
      if (filter) {
        request = request.filter(filter);
      }

      const response = await request.get();

      const messages = response?.value || [];

      if (opts?.expectedId) {
        const found = messages.some((m: GraphMessage) => m.id === opts.expectedId);
        if (found) {
          return messages;
        }
        // expectedId not found yet, continue polling
      } else if (messages.length > 0) {
        // No expectedId, just return first results
        return messages;
      }
    } catch (error: unknown) {
      // Only retry on transient server errors (5xx) and 404 (not indexed yet)
      const statusRaw = error && typeof error === 'object' && ('status' in error || 'statusCode' in error || 'code' in error) ? (error as ErrorWithStatus).status || (error as ErrorWithStatus).statusCode || (error as ErrorWithStatus).code : undefined;
      const status = typeof statusRaw === 'number' ? statusRaw : undefined;
      if (status !== 404 && !(status !== undefined && status >= 500 && status < 600)) {
        // Fail immediately on:
        // - 401 (auth errors)
        // - 429 (rate limits)
        // - 4xx client errors (except 404)
        throw error;
      }
      // 404 or 5xx - continue retry loop (message not indexed yet or transient error)
    }

    await delay(interval);
    // Exponential backoff with cap
    interval = Math.min(interval * 1.5, maxInterval);
  }

  throw new Error(`waitForSearch: timeout after ${timeoutMs}ms waiting for search results. Query: ${JSON.stringify(query)}${opts?.expectedId ? `, Expected ID: ${opts.expectedId}` : ''}`);
}
