import type { Logger, MicrosoftAuthProvider } from '@mcp-z/oauth-microsoft';
import { Client } from '@microsoft/microsoft-graph-client';
import assert from 'assert';
import createTool, { type Input, type Output } from '../../../../src/mcp/tools/message-move-to-trash.ts';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.ts';
import createMiddlewareContext from '../../../lib/create-middleware-context.ts';
import { createTestDraftMessage, deleteTestMessage } from '../../../lib/message-helpers.ts';
import waitForMessage from '../../../lib/wait-for-message.ts';

// Type for objects with a parse method (like Zod schemas)
type SchemaLike = { parse: (data: unknown) => unknown };

// Type for objects with an id property
type ItemWithId = { id?: string; [key: string]: unknown };

describe('outlook-message-move-to-trash', () => {
  // Shared context and Graph client for all tests
  let auth: MicrosoftAuthProvider;
  let logger: Logger;
  let tool: ReturnType<typeof createTool>;
  let wrappedTool: ReturnType<Awaited<ReturnType<typeof createMiddlewareContext>>['middleware']['withToolAuth']>;
  let handler: TypedHandler<Input>;
  let sharedGraph: Client;

  before(async () => {
    const middlewareContext = await createMiddlewareContext();
    auth = middlewareContext.auth;
    logger = middlewareContext.logger;
    const middleware = middlewareContext.middleware;
    tool = createTool();
    wrappedTool = middleware.withToolAuth(tool);
    handler = wrappedTool.handler as TypedHandler<Input>;
    sharedGraph = await Client.initWithMiddleware({ authProvider: auth });
  });

  it('returns structured success or error for non-existent message', async () => {
    const res = await handler({ ids: ['non-existent-id'] }, createExtra());
    assert.ok(res && Array.isArray(res.content), 'move-to-trash did not return content array');

    // Ensure structuredContent exists and matches the declared output schema
    assert.ok(res.structuredContent && res.structuredContent, 'missing structuredContent');

    const schema = tool.config?.outputSchema;
    assert.ok(schema, 'tool.outputSchema missing from tool metadata');

    try {
      // Wrap structuredContent in { result } to match the outputSchema structure
      (schema as SchemaLike).parse(res.structuredContent);
    } catch (err) {
      const message = err instanceof Error && 'issues' in err ? JSON.stringify((err as { issues: unknown }).issues) : String(err);
      assert.fail(`structuredContent failed schema validation: ${message}`);
    }

    // The structured machine-readable payload is canonical; validate it instead of parsing a JSON mirror in content[0]
    const structured = res.structuredContent?.result as Output | undefined;
    assert.ok(structured, 'missing structuredContent');

    if (structured?.type === 'success') {
      assert.strictEqual(structured.totalRequested, 1, 'totalRequested should be 1');
      assert.strictEqual(structured.successCount, 0, 'successCount should be 0 for non-existent');
      assert.strictEqual(structured.failureCount, 1, 'failureCount should be 1 for non-existent');
    }
  });

  it('moves single message to trash', async () => {
    const graph = sharedGraph;

    // Track created resource ids locally to ensure per-test close
    const createdIds: string[] = [];

    try {
      // Get user's email address
      const profile = await graph.api('/me').get();
      const userEmail = profile.mail || profile.userPrincipalName;
      if (!userEmail) assert.fail('Unable to determine test email address');

      const draftId = await createTestDraftMessage(graph, { to: userEmail });
      createdIds.push(draftId);

      // Wait for message to be indexed
      await waitForMessage(graph, draftId);

      // Test single ID in array
      const trashResp = await handler({ ids: [draftId] }, createExtra());

      // Check structured response
      const structured = trashResp.structuredContent?.result as Output | undefined;
      assert.ok(structured, 'structuredContent missing');

      if (structured?.type === 'success') {
        assert.strictEqual(structured.totalRequested, 1, 'totalRequested should be 1');
        assert.strictEqual(structured.successCount, 1, 'successCount should be 1');
        assert.strictEqual(structured.failureCount, 0, 'failureCount should be 0');
        assert.strictEqual(structured.results.length, 1, 'results length should be 1');
        const firstResult = structured.results[0];
        if (firstResult) {
          assert.strictEqual(firstResult.id, draftId, 'result id should match');
          assert.strictEqual(firstResult.success, true, 'result success should be true');
        }
      } else {
        assert.fail(`expected success branch but received error: ${JSON.stringify(structured)}`);
      }
    } finally {
      // Per-test close: delete created messages
      if (createdIds.length > 0) {
        for (const id of createdIds) {
          await deleteTestMessage(graph, id, logger);
        }
      }
    }
  });

  it('moves multiple messages to trash in batch', async () => {
    const graph = sharedGraph;

    // Track created resource ids locally to ensure per-test close
    const createdIds: string[] = [];

    try {
      // Get user's email address
      const profile = await graph.api('/me').get();
      const userEmail = profile.mail || profile.userPrincipalName;
      if (!userEmail) assert.fail('Unable to determine test email address');

      // Create multiple test draft messages
      const draftId1 = await createTestDraftMessage(graph, { subject: `ci-batch-1-${Date.now()}`, to: userEmail });
      const draftId2 = await createTestDraftMessage(graph, { subject: `ci-batch-2-${Date.now()}`, to: userEmail });
      createdIds.push(draftId1, draftId2);

      // Wait for messages to be indexed
      await waitForMessage(graph, draftId1);
      await waitForMessage(graph, draftId2);

      // Test array of IDs
      const trashResp = await handler({ ids: [draftId1, draftId2] }, createExtra());

      // Check structured response
      const structured = trashResp.structuredContent?.result as Output | undefined;
      assert.ok(structured, 'structuredContent missing');

      if (structured?.type === 'success') {
        assert.strictEqual(structured.totalRequested, 2, 'totalRequested should be 2');
        assert.strictEqual(structured.successCount, 2, 'successCount should be 2');
        assert.strictEqual(structured.failureCount, 0, 'failureCount should be 0');
        assert.strictEqual(structured.results.length, 2, 'results length should be 2');

        // Check each result
        const result1 = structured.results.find((r: ItemWithId) => r.id === draftId1);
        const result2 = structured.results.find((r: ItemWithId) => r.id === draftId2);

        assert.ok(result1, 'result for draftId1 should exist');
        assert.strictEqual(result1.success, true, 'result1 success should be true');

        assert.ok(result2, 'result for draftId2 should exist');
        assert.strictEqual(result2.success, true, 'result2 success should be true');
      } else {
        assert.fail(`expected success branch but received error: ${JSON.stringify(structured)}`);
      }
    } finally {
      // Per-test close: delete created messages
      if (createdIds.length > 0) {
        for (const id of createdIds) {
          await deleteTestMessage(graph, id, logger);
        }
      }
    }
  });

  it('handles mixed success/failure batch operation', async () => {
    const graph = sharedGraph;

    // Track created resource ids locally to ensure per-test close
    const createdIds: string[] = [];

    try {
      // Get user's email address
      const profile = await graph.api('/me').get();
      const userEmail = profile.mail || profile.userPrincipalName;
      if (!userEmail) assert.fail('Unable to determine test email address');

      // Create one valid draft message
      const validId = await createTestDraftMessage(graph, { subject: `ci-mixed-${Date.now()}`, to: userEmail });
      createdIds.push(validId);

      // Wait for message to be indexed
      await waitForMessage(graph, validId);

      const nonExistentId = `non_existent_${Date.now()}`;

      // Test mixed valid and invalid IDs
      const trashResp = await handler({ ids: [validId, nonExistentId] }, createExtra());

      // Check structured response
      const structured = trashResp.structuredContent?.result as Output | undefined;
      assert.ok(structured, 'structuredContent missing');

      if (structured?.type === 'success') {
        assert.strictEqual(structured.totalRequested, 2, 'totalRequested should be 2');
        assert.strictEqual(structured.successCount, 1, 'successCount should be 1');
        assert.strictEqual(structured.failureCount, 1, 'failureCount should be 1');
        assert.strictEqual(structured.results.length, 2, 'results length should be 2');

        // Check results
        const validResult = structured.results.find((r: ItemWithId) => r.id === validId);
        const invalidResult = structured.results.find((r: ItemWithId) => r.id === nonExistentId);

        assert.ok(validResult, 'valid result should exist');
        assert.strictEqual(validResult.success, true, 'valid result should be successful');

        assert.ok(invalidResult, 'invalid result should exist');
        assert.strictEqual(invalidResult.success, false, 'invalid result should fail');
        assert.ok(invalidResult.error, 'invalid result should have error message');
      } else {
        assert.fail(`expected success branch but received error: ${JSON.stringify(structured)}`);
      }
    } finally {
      // Per-test close: delete created messages
      if (createdIds.length > 0) {
        for (const id of createdIds) {
          await deleteTestMessage(graph, id, logger);
        }
      }
    }
  });
});
