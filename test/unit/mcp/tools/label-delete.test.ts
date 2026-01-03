import type { Logger, MicrosoftAuthProvider } from '@mcp-z/oauth-microsoft';
import { Client } from '@microsoft/microsoft-graph-client';
import assert from 'assert';
import createTool, { type Input, type Output } from '../../../../src/mcp/tools/label-delete.ts';
import { categoryExists, createTestCategory, deleteTestCategory, waitForCategoryDeleted } from '../../../lib/category-helpers.ts';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.ts';
import createMiddlewareContext from '../../../lib/create-middleware-context.ts';
import waitForCategory from '../../../lib/wait-for-category.ts';

// Type for objects with an id property
type ItemWithId = { id?: string; [key: string]: unknown };

describe('outlook-label-delete', () => {
  // Shared context and Graph client for all tests
  let auth: MicrosoftAuthProvider;
  let logger: Logger;
  let sharedGraph: Client;
  let _wrappedTool: ReturnType<Awaited<ReturnType<typeof createMiddlewareContext>>['middleware']['withToolAuth']>;
  let handler: TypedHandler<Input>;

  before(async () => {
    const middlewareContext = await createMiddlewareContext();
    auth = middlewareContext.auth;
    logger = middlewareContext.logger;
    const middleware = middlewareContext.middleware;
    sharedGraph = await Client.initWithMiddleware({ authProvider: auth });
    const tool = createTool();
    const wrappedTool = middleware.withToolAuth(tool);
    handler = wrappedTool.handler as TypedHandler<Input>;
  });

  it('deletes single category successfully', async () => {
    const graph = sharedGraph;

    // Track created resource ids locally to ensure per-test close
    const createdIds: string[] = [];

    try {
      // Create a test category
      const categoryId = await createTestCategory(graph);
      createdIds.push(categoryId);

      // Wait for category to be indexed
      await waitForCategory(graph, categoryId);

      // Verify category exists
      assert.ok(await categoryExists(graph, categoryId), 'Test category should exist before deletion');

      // Delete the category using single ID in array
      const response = await handler({ ids: [categoryId] }, createExtra());

      // Check structured response
      const structured = response?.structuredContent?.result as Output | undefined;
      assert.ok(structured, 'structuredContent missing');

      if (structured?.type === 'success') {
        assert.strictEqual(structured.totalRequested, 1, 'totalRequested should be 1');
        assert.strictEqual(structured.successCount, 1, 'successCount should be 1');
        assert.strictEqual(structured.failureCount, 0, 'failureCount should be 0');
        assert.strictEqual(structured.results.length, 1, 'results length should be 1');
        const firstResult = structured.results[0];
        if (firstResult) {
          assert.strictEqual(firstResult.id, categoryId, 'result id should match');
          assert.strictEqual(firstResult.success, true, 'result success should be true');
          assert.strictEqual(firstResult.error, undefined, 'result error should be undefined');
        }

        // Wait for category to be deleted (eventual consistency)
        await waitForCategoryDeleted(graph, categoryId, { timeout: 5000 });

        // Verify category no longer exists
        assert.ok(!(await categoryExists(graph, categoryId)), 'Category should not exist after deletion');

        // Remove from close list since deletion was successful
        const index = createdIds.indexOf(categoryId);
        if (index > -1) {
          createdIds.splice(index, 1);
        }
      } else {
        assert.fail(`expected success branch but received error: ${JSON.stringify(structured)}`);
      }
    } finally {
      // Cleanup any remaining categories (only if test failed before deletion)
      for (const id of createdIds) {
        await deleteTestCategory(graph, id, logger);
      }
    }
  });

  it('deletes multiple categories in batch', async () => {
    const graph = sharedGraph;

    // Track created resource ids locally to ensure per-test close
    const createdIds: string[] = [];

    try {
      // Create multiple test categories
      const categoryId1 = await createTestCategory(graph, { displayName: `ci-test-batch-1-${Date.now()}` });
      const categoryId2 = await createTestCategory(graph, { displayName: `ci-test-batch-2-${Date.now()}` });
      createdIds.push(categoryId1, categoryId2);

      // Wait for BOTH categories to be indexed
      await Promise.all([waitForCategory(graph, categoryId1, { timeout: 20000 }), waitForCategory(graph, categoryId2, { timeout: 20000 })]);

      // Verify categories exist before deletion
      const exists1 = await categoryExists(graph, categoryId1);
      const exists2 = await categoryExists(graph, categoryId2);
      assert.ok(exists1, 'Test category 1 should exist before deletion');
      assert.ok(exists2, 'Test category 2 should exist before deletion');

      // Delete the categories using array of IDs
      const response = await handler({ ids: [categoryId1, categoryId2] }, createExtra());

      // Check structured response
      const structured = response?.structuredContent?.result as Output | undefined;
      assert.ok(structured, 'structuredContent missing');

      if (structured.type === 'success') {
        // Tool successfully deleted - clear close list BEFORE assertions
        createdIds.length = 0;

        assert.strictEqual(structured.totalRequested, 2, 'totalRequested should be 2');
        assert.strictEqual(structured.successCount, 2, 'successCount should be 2');
        assert.strictEqual(structured.failureCount, 0, 'failureCount should be 0');
        assert.strictEqual(structured.results.length, 2, 'results length should be 2');

        // Check each result
        const result1 = structured.results.find((r: ItemWithId) => r.id === categoryId1);
        const result2 = structured.results.find((r: ItemWithId) => r.id === categoryId2);

        assert.ok(result1, 'result for categoryId1 should exist');
        assert.strictEqual(result1.success, true, 'result1 success should be true');
        assert.strictEqual(result1.error, undefined, 'result1 error should be undefined');

        assert.ok(result2, 'result for categoryId2 should exist');
        assert.strictEqual(result2.success, true, 'result2 success should be true');
        assert.strictEqual(result2.error, undefined, 'result2 error should be undefined');

        // Wait for categories to be deleted (eventual consistency)
        await waitForCategoryDeleted(graph, categoryId1, { timeout: 5000 });
        await waitForCategoryDeleted(graph, categoryId2, { timeout: 5000 });

        // Verify categories no longer exist
        assert.ok(!(await categoryExists(graph, categoryId1)), 'Category 1 should not exist after deletion');
        assert.ok(!(await categoryExists(graph, categoryId2)), 'Category 2 should not exist after deletion');
      } else {
        assert.fail(`expected success branch but received error: ${JSON.stringify(structured)}`);
      }
    } finally {
      // Cleanup any remaining categories (only if test failed before deletion)
      for (const id of createdIds) {
        await deleteTestCategory(graph, id, logger);
      }
    }
  });

  it('handles non-existent category deletion gracefully', async () => {
    const nonExistentId = `non-existent-category-${Date.now()}`;

    // Try to delete a non-existent category
    const response = await handler({ ids: [nonExistentId] }, createExtra());

    // Check structured response
    const structured = response?.structuredContent?.result as Output | undefined;
    assert.ok(structured, 'structuredContent missing');

    if (structured?.type === 'success') {
      assert.strictEqual(structured.totalRequested, 1, 'totalRequested should be 1');
      assert.strictEqual(structured.successCount, 0, 'successCount should be 0');
      assert.strictEqual(structured.failureCount, 1, 'failureCount should be 1');
      assert.strictEqual(structured.results.length, 1, 'results length should be 1');
      const firstResult = structured.results[0];
      if (firstResult) {
        assert.strictEqual(firstResult.id, nonExistentId, 'result id should match');
        assert.strictEqual(firstResult.success, false, 'result success should be false');
        assert.ok(firstResult.error, 'error should be present');
      }
    } else {
      assert.fail(`expected success branch but received error: ${JSON.stringify(structured)}`);
    }
  });

  it('handles mixed success/failure batch operation', async () => {
    const graph = sharedGraph;

    // Track created resource ids locally to ensure per-test close
    const createdIds: string[] = [];

    try {
      // Create one test category
      const validCategoryId = await createTestCategory(graph, { displayName: `ci-test-mixed-${Date.now()}` });
      createdIds.push(validCategoryId);

      // Wait for category to be indexed
      await waitForCategory(graph, validCategoryId);

      // Try to delete valid category + non-existent category
      const nonExistentId = `non-existent-category-${Date.now()}`;
      const response = await handler(
        {
          ids: [validCategoryId, nonExistentId],
        },
        createExtra()
      );

      // Check structured response
      const structured = response?.structuredContent?.result as Output | undefined;
      assert.ok(structured, 'structuredContent missing');

      if (structured.type === 'success') {
        assert.strictEqual(structured.totalRequested, 2, 'totalRequested should be 2');
        assert.strictEqual(structured.successCount, 1, 'successCount should be 1 (valid category)');
        assert.strictEqual(structured.failureCount, 1, 'failureCount should be 1 (non-existent)');
        assert.strictEqual(structured.results.length, 2, 'results length should be 2');

        // Check results
        const validResult = structured.results.find((r: ItemWithId) => r.id === validCategoryId);
        const nonExistentResult = structured.results.find((r: ItemWithId) => r.id === nonExistentId);

        assert.ok(validResult, 'valid result should exist');
        assert.strictEqual(validResult.success, true, 'valid result should be successful');

        assert.ok(nonExistentResult, 'non-existent result should exist');
        assert.strictEqual(nonExistentResult.success, false, 'non-existent result should fail');

        // Wait for valid category to be deleted (eventual consistency)
        await waitForCategoryDeleted(graph, validCategoryId, { timeout: 5000 });

        // Verify valid category no longer exists
        assert.ok(!(await categoryExists(graph, validCategoryId)), 'Valid category should not exist after deletion');

        // Remove successfully deleted category from close list
        const index = createdIds.indexOf(validCategoryId);
        if (index > -1) {
          createdIds.splice(index, 1);
        }
      } else {
        assert.fail(`expected success branch but received error: ${JSON.stringify(structured)}`);
      }
    } finally {
      // Cleanup any remaining categories (only if test failed before deletion)
      for (const id of createdIds) {
        await deleteTestCategory(graph, id, logger);
      }
    }
  });
});
