import type { Logger, MicrosoftAuthProvider } from '@mcp-z/oauth-microsoft';
import { Client } from '@microsoft/microsoft-graph-client';
import assert from 'assert';
import createTool, { type Input, type Output } from '../../../../src/mcp/tools/message-get.js';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.js';
import createMiddlewareContext from '../../../lib/create-middleware-context.js';
import { createTestDraftMessage, deleteTestMessage } from '../../../lib/message-helpers.js';
import waitForMessage from '../../../lib/wait-for-message.js';

describe('outlook-message-get tests', () => {
  // Shared instances for all tests
  let auth: MicrosoftAuthProvider;
  let logger: Logger;
  let _wrappedTool: ReturnType<Awaited<ReturnType<typeof createMiddlewareContext>>['middleware']['withToolAuth']>;
  let handler: TypedHandler<Input>;
  let sharedGraphClient: Client;
  let sharedMessageId: string;
  let sharedMessageSubject: string;

  before(async () => {
    const middlewareContext = await createMiddlewareContext();
    auth = middlewareContext.auth;
    logger = middlewareContext.logger;
    const middleware = middlewareContext.middleware;
    const tool = createTool();
    const wrappedTool = middleware.withToolAuth(tool);
    handler = wrappedTool.handler as TypedHandler<Input>;
    sharedGraphClient = await Client.initWithMiddleware({ authProvider: auth });

    // Create SHARED test draft message for all field tests (avoids rate limits)
    const profile = await sharedGraphClient.api('/me').get();
    const userEmail = profile.mail || profile.userPrincipalName;
    if (!userEmail) throw new Error('Unable to determine test email address');

    sharedMessageSubject = `ci-message-get-test-${Date.now()}`;
    sharedMessageId = await createTestDraftMessage(sharedGraphClient, {
      subject: sharedMessageSubject,
      body: 'Test message for message-get field tests',
      to: userEmail,
    });

    // Wait for message to be indexed
    await waitForMessage(sharedGraphClient, sharedMessageId);
  });

  after(async () => {
    // Cleanup shared test message
    if (sharedMessageId) {
      await deleteTestMessage(sharedGraphClient, sharedMessageId, logger);
    }
  });
  it('get returns structured payload or error for invalid id', async () => {
    // Errors are now thrown as McpError instead of returned as structuredContent
    try {
      await handler({ id: 'non-existent-id', fields: 'id,subject,from,body', contentType: 'text', excludeThreadHistory: false }, createExtra());
      assert.fail('Expected McpError to be thrown for invalid id');
    } catch (err) {
      // Verify it's an McpError with appropriate error message
      assert.ok(err instanceof Error, 'Error should be an Error instance');
      assert.ok(err.message.includes('Error fetching message') || err.message.includes('malformed'), 'Error message should indicate fetch failure or malformed id');
    }
  });

  describe('fields parameter tests', () => {
    it('default multiple fields returns full message item', async () => {
      const res = await handler({ id: sharedMessageId, fields: 'id,subject,from,body', contentType: 'text', excludeThreadHistory: false }, createExtra());

      // Errors are now thrown as McpError, not returned
      if (res) {
        const branch = res.structuredContent?.result as Output | undefined;

        if (branch?.type === 'success') {
          // Should have item field with full message object
          assert.ok(branch.item, 'should have item wrapper');

          // Verify full message data is present
          assert.equal(typeof branch.item.id, 'string', 'item should have id field');
          assert.ok(branch.item.id && branch.item.id.length > 0, 'id should have a value');

          // When fields are requested, they MUST exist in response (but can be empty/undefined for some message types)
          assert.ok('subject' in branch.item, 'subject field should exist when requested');
          assert.ok('from' in branch.item, 'from field should exist when requested');
          assert.ok('body' in branch.item, 'body field should exist when requested');

          // Validate field types and values for fields we expect to be populated
          assert.ok(typeof branch.item.subject === 'string', 'subject should be string type');
          assert.ok(branch.item.subject.length > 0, 'subject should have a value in our test message');

          // body should have value since we set it in createTestDraftMessage
          assert.ok(typeof branch.item.body === 'string', 'body should be string type');
          assert.ok(branch.item.body.length > 0, 'body should have a value in our test message');

          // from field exists when requested, but can be undefined for draft messages
          // Draft messages don't have a sender yet
          if (branch.item.from !== undefined) {
            assert.ok(typeof branch.item.from === 'string', 'from should be string type when present');
          }

          // Verify SPECIFIC content we created
          assert.ok(branch.item.subject.includes('ci-message-get-test'), 'subject should contain test identifier');
        }
      }
    });

    it('multiple fields explicitly returns full message item', async () => {
      const res = await handler({ id: sharedMessageId, fields: 'id,subject,from,body', contentType: 'text', excludeThreadHistory: false }, createExtra());

      // Errors are now thrown as McpError, not returned
      if (res) {
        const branch = res.structuredContent?.result as Output | undefined;

        if (branch?.type === 'success') {
          // Should have item field with full message object
          assert.ok(branch.item, 'should have item wrapper');

          // Verify full message data is present
          assert.equal(typeof branch.item.id, 'string', 'item should have id field');
          assert.ok(branch.item.id && branch.item.id.length > 0, 'id should have a value');

          // When fields are requested, they MUST exist in response (but can be empty/undefined for some message types)
          assert.ok('subject' in branch.item, 'subject field should exist when requested');
          assert.ok('from' in branch.item, 'from field should exist when requested');
          assert.ok('body' in branch.item, 'body field should exist when requested');

          // Validate field types and values for fields we expect to be populated
          assert.ok(typeof branch.item.subject === 'string', 'subject should be string type');
          assert.ok(branch.item.subject.length > 0, 'subject should have a value in our test message');

          // body should have value since we set it in createTestDraftMessage
          assert.ok(typeof branch.item.body === 'string', 'body should be string type');
          assert.ok(branch.item.body.length > 0, 'body should have a value in our test message');

          // from field exists when requested, but can be undefined for draft messages
          // Draft messages don't have a sender yet
          if (branch.item.from !== undefined) {
            assert.ok(typeof branch.item.from === 'string', 'from should be string type when present');
          }

          // Verify SPECIFIC content we created
          assert.ok(branch.item.subject.includes('ci-message-get-test'), 'subject should contain test identifier');
        }
      }
    });

    it('minimal fields returns messageId only', async () => {
      const res = await handler({ id: sharedMessageId, fields: 'id', contentType: 'text', excludeThreadHistory: false }, createExtra());

      // Errors are now thrown as McpError, not returned
      if (res) {
        const branch = res.structuredContent?.result as Output | undefined;

        if (branch?.type === 'success') {
          // Should have item field with only id property when fields='id'
          assert.ok(branch.item, 'should have item wrapper');
          assert.ok(branch.item.id, 'should have messageId field');
          assert.equal(typeof branch.item.id, 'string', 'item.id should be string');

          // Verify item.id matches the expected ID
          assert.equal(branch.item.id, sharedMessageId, 'item.id should match requested message ID');
        }
      }
    });

    it('minimal fields with nonexistent message', async () => {
      const nonexistentId = 'nonexistent-message-id-123';

      // Errors are now thrown as McpError instead of returned as structuredContent
      try {
        await handler({ id: nonexistentId, fields: 'id', contentType: 'text', excludeThreadHistory: false }, createExtra());
        assert.fail('Expected McpError to be thrown for nonexistent message');
      } catch (err) {
        assert.ok(err instanceof Error, 'Error should be an Error instance');
        assert.ok(err.message.includes('Error fetching message') || err.message.includes('malformed'), 'Error message should indicate fetch failure');
      }
    });

    it('fields parameter behavior with missing id parameter', async () => {
      // Errors are now thrown as McpError instead of returned as structuredContent
      // Test with multiple fields
      try {
        await handler({ fields: 'id,subject,from,body', contentType: 'text', excludeThreadHistory: false } as Input, createExtra());
        assert.fail('Expected McpError to be thrown when id is missing (multiple fields)');
      } catch (err) {
        assert.ok(err instanceof Error, 'Error should be an Error instance');
        assert.ok(err.message.includes('Missing id') || err.message.includes('required'), 'Error message should indicate missing id');
      }

      // Test with minimal fields
      try {
        await handler({ fields: 'id', contentType: 'text', excludeThreadHistory: false } as Input, createExtra());
        assert.fail('Expected McpError to be thrown when id is missing (minimal fields)');
      } catch (err) {
        assert.ok(err instanceof Error, 'Error should be an Error instance');
        assert.ok(err.message.includes('Missing id') || err.message.includes('required'), 'Error message should indicate missing id');
      }
    });
  });

  describe('negative test cases - error handling', () => {
    it('returns error for completely invalid message id', async () => {
      const invalidId = 'this-is-definitely-not-a-valid-message-id-format';

      // Errors are now thrown as McpError instead of returned as structuredContent
      try {
        await handler({ id: invalidId, fields: 'id,subject', contentType: 'text', excludeThreadHistory: false }, createExtra());
        assert.fail('Expected McpError to be thrown for invalid id format');
      } catch (err) {
        assert.ok(err instanceof Error, 'Error should be an Error instance');
        assert.ok(err.message.length > 0, 'Error message should be non-empty');
      }
    });

    it('returns error for empty string id', async () => {
      // Errors are now thrown as McpError instead of returned as structuredContent
      try {
        await handler({ id: '', fields: 'id', contentType: 'text', excludeThreadHistory: false }, createExtra());
        assert.fail('Expected McpError to be thrown for empty id');
      } catch (err) {
        assert.ok(err instanceof Error, 'Error should be an Error instance');
        assert.ok(err.message.includes('Missing id') || err.message.includes('required') || err.message.includes('empty'), 'Error message should indicate missing/empty id');
      }
    });

    it('validates field selection filters correctly', async () => {
      // Request only id and subject, verify body is NOT returned
      const res = await handler({ id: sharedMessageId, fields: 'id,subject', contentType: 'text', excludeThreadHistory: false }, createExtra());
      const branch = res.structuredContent?.result as Output | undefined;

      if (branch?.type === 'success') {
        // Should have item wrapper
        assert.ok(branch.item, 'should have item wrapper');

        // Requested fields must exist
        assert.ok('id' in branch.item, 'id should exist when requested');
        assert.ok('subject' in branch.item, 'subject should exist when requested');

        // Body should NOT exist since not requested
        assert.ok(!('body' in branch.item), 'body should not exist when not requested in fields');
        assert.ok(!('from' in branch.item), 'from should not exist when not requested in fields');
      }
    });

    it('handles malformed fields parameter gracefully', async () => {
      // Test with invalid field names
      const res = await handler({ id: sharedMessageId, fields: 'id,invalidFieldName,anotherInvalidField', contentType: 'text', excludeThreadHistory: false }, createExtra());
      const branch = res.structuredContent?.result as Output | undefined;

      // Should still succeed and return id field (valid fields should work)
      if (branch?.type === 'success') {
        // Should have item wrapper
        assert.ok(branch.item, 'should have item wrapper');
        assert.ok('id' in branch.item, 'should return valid field (id)');

        // Invalid fields should be ignored, not cause error
        assert.ok(!('invalidFieldName' in branch.item), 'invalid field should not appear in response');
      }
    });

    it('returns structured error for Graph API failures', async () => {
      // Use an id that looks valid but doesn't exist
      const nonExistentButValidFormatId = 'AAMkAGZmYWI4ZDk2LTM5YzQtNGE4Yi1iZjg1LWQ4NTg5YmI4NmI4YwBGAAAAAADq8M4vxkqXTaHq3qK0kZ4HBwCE_invalid_id_test';

      // Errors are now thrown as McpError instead of returned as structuredContent
      try {
        await handler({ id: nonExistentButValidFormatId, fields: 'id,subject,from', contentType: 'text', excludeThreadHistory: false }, createExtra());
        assert.fail('Expected McpError to be thrown for Graph API failure');
      } catch (err) {
        assert.ok(err instanceof Error, 'Error should be an Error instance');
        // Should have proper error message from Graph API
        assert.ok(err.message, 'error should have message field');
        assert.ok(typeof err.message === 'string', 'error message should be string');
      }
    });
  });
});
