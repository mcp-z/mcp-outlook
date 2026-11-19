import '../../../lib/env-loader.ts';
import { ProtocolError, ProtocolErrorCode } from '@mcp-z/server';
import type { Client } from '@microsoft/microsoft-graph-client';
import assert from 'assert';
import { searchMessages } from '../../../../src/email/querying/search-execution.ts';
import type { OutlookQuery } from '../../../../src/schemas/outlook-query-schema.ts';
import { InvalidQueryError } from '../../../../src/schemas/outlook-query-schema.ts';

describe('searchMessages', () => {
  const logger = { debug: () => {}, info: () => {}, warn: () => {}, error: () => {} };

  it('rejects a bad v2 pageToken as InvalidParams before any Graph request', async () => {
    // The token is decoded before the client is used, so an empty client proves no request was made.
    const graph = {} as Client;
    await assert.rejects(searchMessages(graph, { pageToken: 'v2:!!!', logger }), (error: unknown) => {
      assert.ok(error instanceof ProtocolError);
      assert.strictEqual(error.code, ProtocolErrorCode.InvalidParams);
      return true;
    });
  });

  it('rejects an invalid category as InvalidQueryError before any Graph request', async () => {
    // The filter is built before the client is used, so an empty client proves no request was made.
    const graph = {} as Client;
    const query = { categories: 'bogus' } as unknown as OutlookQuery;
    await assert.rejects(searchMessages(graph, { query, logger }), (error: unknown) => {
      assert.ok(error instanceof InvalidQueryError);
      assert.match(error.message, /Invalid Outlook category: "bogus"/);
      return true;
    });
  });
});
