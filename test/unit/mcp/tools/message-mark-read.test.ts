import type { Logger, MicrosoftAuthProvider } from '@mcp-z/oauth-microsoft';
import { Client } from '@microsoft/microsoft-graph-client';
import assert from 'assert';
import createTool, { type Input, type Output } from '../../../../src/mcp/tools/message-mark-read.js';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.js';
import createMiddlewareContext from '../../../lib/create-middleware-context.js';
import { createTestDraftMessage, deleteTestMessage } from '../../../lib/message-helpers.js';
import waitForMessage from '../../../lib/wait-for-message.js';

describe('outlook-message-mark-read', () => {
  let auth: MicrosoftAuthProvider;
  let logger: Logger;
  let _wrappedTool: ReturnType<Awaited<ReturnType<typeof createMiddlewareContext>>['middleware']['withToolAuth']>;
  let handler: TypedHandler<Input>;
  let sharedGraph: Client;

  before(async () => {
    const middlewareContext = await createMiddlewareContext();
    auth = middlewareContext.auth;
    logger = middlewareContext.logger;
    const middleware = middlewareContext.middleware;
    const tool = createTool();
    const wrappedTool = middleware.withToolAuth(tool);
    handler = wrappedTool.handler;
    sharedGraph = await Client.initWithMiddleware({ authProvider: auth });
  });

  it('marks unread message as read and verifies state change', async () => {
    const createdIds: string[] = [];

    try {
      // 1. Get user email
      const profile = await sharedGraph.api('/me').get();
      const userEmail = profile.mail || profile.userPrincipalName;
      if (!userEmail) assert.fail('Unable to determine test email address');

      // 2. Create test draft message (drafts are unread by default, avoids rate limits)
      const draftId = await createTestDraftMessage(sharedGraph, {
        subject: `ci-mark-read-test-${Date.now()}`,
        to: userEmail,
      });
      createdIds.push(draftId);

      // 3. Wait for message to be ready
      await waitForMessage(sharedGraph, draftId);

      // 4. Mark message as read (drafts start as unread by default)
      const res = await handler({ id: draftId }, createExtra());

      // 5. Validate response structure
      assert.ok(res && Array.isArray(res.content), 'mark-read did not return content array');
      assert.ok(res.structuredContent && res.structuredContent, 'missing structuredContent');

      const branch: Output | undefined = res.structuredContent?.result as Output;

      if (branch?.type === 'success') {
        // 6. Verify message is now marked as read (fetch and validate state change)
        const afterMessage = await sharedGraph.api(`/me/messages/${draftId}`).select('isRead').get();
        assert.strictEqual(afterMessage.isRead, true, 'Message should be marked as read after operation');
      } else {
        assert.fail(`Expected success but got: ${JSON.stringify(branch)}`);
      }
    } finally {
      // 7. Cleanup (fail loud on errors)
      for (const id of createdIds) {
        await deleteTestMessage(sharedGraph, id, logger);
      }
    }
  });

  it('mark-read returns error for nonexistent message', async () => {
    // Errors are now thrown as McpError instead of returned as structuredContent
    try {
      await handler({ id: 'non-existent-id' }, createExtra());
      assert.fail('Expected McpError to be thrown for nonexistent message');
    } catch (err) {
      assert.ok(err instanceof Error, 'Error should be an Error instance');
      assert.ok(err.message.includes('Error marking message') || err.message.includes('malformed'), 'Error message should indicate marking failure');
    }
  });
});
