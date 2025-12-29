import type { Logger, MicrosoftAuthProvider } from '@mcp-z/oauth-microsoft';
import { Client } from '@microsoft/microsoft-graph-client';
import assert from 'assert';
import createTool, { type Input, type Output } from '../../../../src/mcp/tools/label-add.js';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.js';
import createMiddlewareContext from '../../../lib/create-middleware-context.js';
import { createTestDraftMessage, deleteTestMessage } from '../../../lib/message-helpers.js';
import waitForMessage from '../../../lib/wait-for-message.js';

describe('outlook-label-add', () => {
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

  it('adds label to real message and verifies it was added', async () => {
    const createdIds: string[] = [];

    try {
      // 1. Get user email
      const profile = await sharedGraph.api('/me').get();
      const userEmail = profile.mail || profile.userPrincipalName;
      if (!userEmail) assert.fail('Unable to determine test email address');

      // 2. Create test draft message (avoids hitting daily sending limits)
      const draftId = await createTestDraftMessage(sharedGraph, {
        subject: `ci-label-test-${Date.now()}`,
        to: userEmail,
      });
      createdIds.push(draftId);

      // 3. Wait for message to be ready
      await waitForMessage(sharedGraph, draftId);

      // 4. Add label to SPECIFIC message
      const labelName = `ci-label-${Date.now()}`;
      const res = await handler({ id: draftId, labels: [labelName] }, createExtra());

      // 5. Validate response structure
      assert.ok(res && Array.isArray(res.content), 'add_label did not return content array');
      assert.ok(res.structuredContent && res.structuredContent, 'missing structuredContent');

      const branch = res.structuredContent?.result as Output | undefined;

      if (branch?.type === 'success') {
        // 6. Verify label was ACTUALLY added (fetch message again)
        const message = await sharedGraph.api(`/me/messages/${draftId}`).get();
        const categories = message.categories || [];
        assert.ok(categories.includes(labelName), `Label ${labelName} should be in message categories. Found: ${categories.join(', ')}`);
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
});
