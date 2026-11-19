import '../../../lib/env-loader.ts';
import type { EnrichedExtra } from '@mcp-z/oauth-microsoft';
import { ProtocolError, ProtocolErrorCode } from '@mcp-z/server';
import assert from 'assert';
import createTool from '../../../../src/mcp/tools/message-search.ts';

describe('message-search invalid category', () => {
  it('is InvalidParams, raised before the auth provider is used (no Graph request)', async () => {
    const logger = { debug: () => {}, info: () => {}, warn: () => {}, error: () => {} };
    let authUsed = false;
    const auth = {
      getAccessToken: async () => {
        authUsed = true;
        return 'unused';
      },
    };
    const extra = { logger, authContext: { auth, accountId: 'test' } } as unknown as EnrichedExtra;
    const handler = createTool().handler as unknown as (input: unknown, extra: EnrichedExtra) => Promise<unknown>;

    // Bypasses the input schema (as a direct caller can) so the query reaches the category check.
    await assert.rejects(handler({ query: { categories: 'bogus' } }, extra), (error: unknown) => {
      assert.ok(error instanceof ProtocolError);
      assert.strictEqual(error.code, ProtocolErrorCode.InvalidParams);
      assert.match(error.message, /Invalid/);
      return true;
    });
    assert.strictEqual(authUsed, false);
  });
});
