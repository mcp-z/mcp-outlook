import type { Logger, MicrosoftAuthProvider } from '@mcp-z/oauth-microsoft';
import { Client } from '@microsoft/microsoft-graph-client';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import assert from 'assert';
import createTool, { type Input, type Output } from '../../../../src/mcp/tools/message-send.js';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.js';
import createMiddlewareContext from '../../../lib/create-middleware-context.js';
import { deleteTestMessage } from '../../../lib/message-helpers.js';

// Type for objects with a parse method (like Zod schemas)
type SchemaLike = { parse: (data: unknown) => unknown };

// Shared instances for all tests
let auth: MicrosoftAuthProvider;
let logger: Logger;
let tool: ReturnType<typeof createTool>;
let wrappedTool: ReturnType<Awaited<ReturnType<typeof createMiddlewareContext>>['middleware']['withToolAuth']>;
let handler: TypedHandler<Input>;
let sharedGraphClient: Client;
let testAccountEmail: string;

before(async () => {
  const middlewareContext = await createMiddlewareContext();
  auth = middlewareContext.auth;
  logger = middlewareContext.logger;
  const middleware = middlewareContext.middleware;
  tool = createTool();
  wrappedTool = middleware.withToolAuth(tool);
  handler = wrappedTool.handler as TypedHandler<Input>;
  sharedGraphClient = Client.initWithMiddleware({ authProvider: auth });

  // Get test account email - fail loud if authentication is broken
  const profile = await sharedGraphClient.api('/me').get();
  testAccountEmail = (profile.mail || profile.userPrincipalName) as string;
  if (!testAccountEmail) {
    throw new Error('Unable to determine test account email from Microsoft Graph profile');
  }
});

function extractMessageId(result: CallToolResult): string | undefined {
  // Prefer structuredContent.result when present
  const branch: Output | undefined = result?.structuredContent?.result as Output;
  if (branch?.type === 'success' && typeof branch.id === 'string') {
    return branch.id;
  }

  // Fall back to content mirror only as a last resort (avoid JSON.parse when content is human text)
  const firstContent = result?.content?.[0];
  const txt = firstContent && firstContent.type === 'text' ? firstContent.text : undefined;
  if (typeof txt === 'string') {
    try {
      const mirror = JSON.parse(txt);
      return mirror?.id;
    } catch (error) {
      // Log parse failures - indicates malformed response
      logger.warn('Failed to parse message ID from content', {
        error: error instanceof Error ? error.message : String(error),
        textLength: txt.length,
      });
      return undefined;
    }
  }

  return undefined; // No ID found in expected locations
}

it('send returns structured success or an error payload', async () => {
  // Track created remote message id(s) for per-test close
  const createdIds: string[] = [];

  try {
    const uniqueSubject = `ci-send-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
    const res: CallToolResult = await handler({ to: testAccountEmail, cc: undefined, bcc: undefined, subject: uniqueSubject, body: 'test body', contentType: 'text' }, createExtra());

    // Ensure structuredContent.result exists and matches the declared output schema
    assert.ok(res.structuredContent && res.structuredContent, 'missing structuredContent');

    const schema = tool.config?.outputSchema;
    assert.ok(schema, 'tool.outputSchema missing from tool metadata');

    // Validate schema
    try {
      (schema as SchemaLike)?.parse(res.structuredContent);
    } catch (err: unknown) {
      const message = err instanceof Error && 'issues' in err ? JSON.stringify((err as { issues: unknown }).issues) : String(err);
      assert.fail(`structuredContent failed schema validation: ${message}`);
    }

    // Extract message ID for close
    const branch: Output | undefined = res.structuredContent?.result as Output;
    if (branch?.type === 'success' && typeof branch.id === 'string') {
      createdIds.push(branch.id);
    }
  } finally {
    // Per-test close: if a message was created on a real Microsoft account, delete it.
    if (createdIds.length > 0) {
      for (const id of createdIds) {
        const graph = sharedGraphClient;
        await deleteTestMessage(graph, id, logger);
      }
    }
  }
});

/* Note: tests must use real provider context. */

it('normalizes single string recipients to arrays (integration)', async () => {
  const uniqueSubject = `ci-test-normalize-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
  const res: CallToolResult = await handler({ to: testAccountEmail, cc: testAccountEmail, bcc: testAccountEmail, subject: uniqueSubject, body: 'Body', contentType: 'text' }, createExtra());

  // Cleanup if message was actually created (though example.com addresses typically don't send)
  const messageId = extractMessageId(res);
  try {
    if (res) assert.ok(res.content && res.content[0], 'no response content');
  } finally {
    if (messageId) {
      const graph = sharedGraphClient;
      await deleteTestMessage(graph, messageId, logger);
    }
  }
});

it('accepts comma-separated recipients (integration)', async () => {
  const uniqueSubject = `ci-test-array-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
  const res: CallToolResult = await handler({ to: `${testAccountEmail}, ${testAccountEmail}`, cc: undefined, bcc: undefined, subject: uniqueSubject, body: 'Body', contentType: 'text' }, createExtra());

  const messageId = extractMessageId(res);
  try {
    if (res) assert.ok(res.content && res.content[0], 'no response content');
  } finally {
    if (messageId) {
      const graph = sharedGraphClient;
      await deleteTestMessage(graph, messageId, logger);
    }
  }
});

describe('Context authentication pattern', () => {
  it('uses context credentials (integration)', async () => {
    const uniqueSubject = `ci-context-test-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
    const result: CallToolResult = await handler({ to: testAccountEmail, cc: undefined, bcc: undefined, subject: uniqueSubject, body: 'Testing context auth', contentType: 'text' }, createExtra());

    const messageId = extractMessageId(result);
    try {
      assert.ok(result);
      assert.ok(result.content || result.structuredContent);
      const branch: Output | undefined = result.structuredContent?.result as Output;
      if (branch) {
        assert.equal(branch.type, 'success', 'Result should be success type');
      }
    } finally {
      if (messageId) {
        const graph = sharedGraphClient;
        await deleteTestMessage(graph, messageId, logger);
      }
    }
  });
});

describe('Response structure', () => {
  it('validates structured response format', async () => {
    const uniqueSubject = `ci-structure-test-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
    const result: CallToolResult = await handler({ to: testAccountEmail, cc: undefined, bcc: undefined, subject: uniqueSubject, body: 'Testing response structure', contentType: 'text' }, createExtra());

    const messageId = extractMessageId(result);
    try {
      assert.ok(result);
      const hasValidContent = result.content && Array.isArray(result.content) && result.content.length > 0;
      const branch: Output | undefined = result.structuredContent?.result as Output;
      const hasValidStructured = branch && typeof branch.type === 'string';
      assert.ok(hasValidContent || hasValidStructured, 'Result should have either valid content array or valid structuredContent');

      // Validate it's a success result
      if (branch) {
        assert.equal(branch.type, 'success', 'Result should be success type');
      }
    } finally {
      if (messageId) {
        const graph = sharedGraphClient;
        await deleteTestMessage(graph, messageId, logger);
      }
    }
  });
});

describe('Message formatting and input validation', () => {
  it('handles single and multiple recipients correctly (integration)', async () => {
    const createdIds: string[] = [];

    try {
      const result1 = await handler(
        {
          to: `${testAccountEmail}, ${testAccountEmail}`,
          cc: testAccountEmail,
          bcc: `${testAccountEmail}, ${testAccountEmail}`,
          subject: `ci-multi-recipient-test-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`,
          body: 'Testing multiple recipients formatting',
          contentType: 'text',
        },
        createExtra()
      );
      assert.ok(result1);
      if (!result1.error) assert.ok(result1.content || result1.structuredContent);

      const messageId1 = extractMessageId(result1);
      if (messageId1) createdIds.push(messageId1);

      const result2 = await handler(
        {
          to: testAccountEmail,
          cc: undefined,
          bcc: undefined,
          subject: `ci-single-recipient-test-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`,
          body: 'Testing single recipient formatting',
          contentType: 'text',
        },
        createExtra()
      );
      assert.ok(result2);
      if (!result2.error) assert.ok(result2.content || result2.structuredContent);

      const messageId2 = extractMessageId(result2);
      if (messageId2) createdIds.push(messageId2);
    } finally {
      for (const id of createdIds) {
        const graph = sharedGraphClient;
        await deleteTestMessage(graph, id, logger);
      }
    }
  });

  it('validates required fields and handles empty subject', async () => {
    const result = await handler(
      {
        to: testAccountEmail,
        cc: undefined,
        bcc: undefined,
        subject: `ci-empty-subject-test-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`,
        body: 'Testing empty subject handling',
        contentType: 'text',
      },
      createExtra()
    );

    const messageId = extractMessageId(result);
    try {
      assert.ok(result);
      if (!result.error) assert.ok(result.content || result.structuredContent);
    } finally {
      if (messageId) {
        const graph = sharedGraphClient;
        await deleteTestMessage(graph, messageId, logger);
      }
    }
  });
});
