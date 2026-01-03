import assert from 'assert';
import createTool, { type Input } from '../../../../src/mcp/tools/message-respond.ts';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.ts';
import createMiddlewareContext from '../../../lib/create-middleware-context.ts';

describe('outlook-message-respond', () => {
  let _wrappedTool: ReturnType<Awaited<ReturnType<typeof createMiddlewareContext>>['middleware']['withToolAuth']>;
  let handler: TypedHandler<Input>;

  before(async () => {
    const middlewareContext = await createMiddlewareContext();
    const middleware = middlewareContext.middleware;
    const tool = createTool();
    const wrappedTool = middleware.withToolAuth(tool);
    handler = wrappedTool.handler as TypedHandler<Input>;
  });

  it('respond returns structured success or error', async () => {
    // Errors are now thrown as McpError instead of returned as structuredContent
    try {
      await handler({ id: 'non-existent-id', body: 'test reply', contentType: 'text' }, createExtra());
      assert.fail('Expected McpError to be thrown for nonexistent message');
    } catch (err) {
      assert.ok(err instanceof Error, 'Error should be an Error instance');
      assert.ok(err.message.includes('Error replying') || err.message.includes('malformed'), 'Error message should indicate reply failure');
    }
  });
});
