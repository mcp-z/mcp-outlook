import type { Logger, MicrosoftAuthProvider } from '@mcp-z/oauth-microsoft';
import { Client } from '@microsoft/microsoft-graph-client';
import assert from 'assert';
import createTool, { type Input, type Output } from '../../../../src/mcp/tools/categories-list.js';
import { createTestCategory, deleteTestCategory } from '../../../lib/category-helpers.js';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.js';
import createMiddlewareContext from '../../../lib/create-middleware-context.js';
import waitForCategory from '../../../lib/wait-for-category.js';

// Type for objects with an id property
type ItemWithId = { id?: string; [key: string]: unknown };

describe('outlook-categories-list tool', () => {
  let auth: MicrosoftAuthProvider;
  let logger: Logger;
  let tool: ReturnType<typeof createTool>;
  let wrappedTool: ReturnType<Awaited<ReturnType<typeof createMiddlewareContext>>['middleware']['withToolAuth']>;
  let handler: TypedHandler<Input>;
  let graph: Client;

  before(async () => {
    const middlewareContext = await createMiddlewareContext();
    auth = middlewareContext.auth;
    logger = middlewareContext.logger;
    const middleware = middlewareContext.middleware;
    tool = createTool();
    wrappedTool = middleware.withToolAuth(tool);
    handler = wrappedTool.handler as TypedHandler<Input>;
    graph = await Client.initWithMiddleware({ authProvider: auth });
  });

  it('creates specific test categories and validates they are returned', async () => {
    const createdIds: string[] = [];

    try {
      // 1. Create SPECIFIC test categories with unique names
      const category1Name = `ci-cat-1-${Date.now()}`;
      const category2Name = `ci-cat-2-${Date.now()}`;

      const cat1Id = await createTestCategory(graph, { displayName: category1Name, color: 'preset1' });
      createdIds.push(cat1Id);

      const cat2Id = await createTestCategory(graph, { displayName: category2Name, color: 'preset2' });
      createdIds.push(cat2Id);

      // 2. Wait for categories to be indexed
      await waitForCategory(graph, cat1Id);
      await waitForCategory(graph, cat2Id);

      // 3. Call categories-list tool
      const res = await handler({}, createExtra());

      // 4. Validate response structure
      assert.ok(res && Array.isArray(res.content), 'categories_list did not return content array');
      assert.ok(res.structuredContent && res.structuredContent, 'missing structuredContent');

      const branch = res.structuredContent?.result as Output | undefined;

      if (branch?.type === 'success') {
        // 5. Validate SPECIFIC categories are found in the list
        const items = branch.items;
        assert.ok(Array.isArray(items), 'success response should have items array');

        const foundCat1 = items.find((item: ItemWithId) => item.id === cat1Id);
        const foundCat2 = items.find((item: ItemWithId) => item.id === cat2Id);

        assert.ok(foundCat1, `Category 1 (${cat1Id}) should be found in results`);
        assert.ok(foundCat2, `Category 2 (${cat2Id}) should be found in results`);

        // Validate category 1 properties
        assert.strictEqual(foundCat1.displayName, category1Name, 'Category 1 displayName should match');
        assert.strictEqual(foundCat1.color, 'preset1', 'Category 1 color should match');

        // Validate category 2 properties
        assert.strictEqual(foundCat2.displayName, category2Name, 'Category 2 displayName should match');
        assert.strictEqual(foundCat2.color, 'preset2', 'Category 2 color should match');
      } else {
        assert.fail(`Expected success but got: ${JSON.stringify(branch)}`);
      }
    } finally {
      // 6. Cleanup (fail loud on errors)
      for (const id of createdIds) {
        await deleteTestCategory(graph, id, logger);
      }
    }
  });

  it('tool follows MCP naming convention', () => {
    // Tool module should export a name that matches expected pattern
    assert.strictEqual(tool.name, 'categories-list', 'tool name should follow {resource}-{action} pattern');
  });

  it('validates tool metadata and configuration', () => {
    // Validate tool name follows convention
    assert.strictEqual(tool.config.description?.includes('categories'), true, 'description should mention categories');

    // Validate input schema (should be empty object for this tool)
    assert.ok(tool.config.inputSchema, 'tool should have inputSchema');
    // For Zod schemas, check the shape instead of object keys
    const parsedEmpty = (tool.config.inputSchema as { safeParse: (data: unknown) => { success: boolean } }).safeParse({});
    assert.ok(parsedEmpty.success, 'inputSchema should accept empty object for categories-list');

    // Validate output schema structure
    assert.ok(tool.config.outputSchema, 'tool should have outputSchema');
    const schema = tool.config?.outputSchema;
    assert.ok(schema, 'outputSchema should be defined');
  });

  it('tool follows MCP naming convention', () => {
    // Tool module should export a name that matches expected pattern
    assert.strictEqual(tool.name, 'categories-list', 'tool name should follow {resource}-{action} pattern');
  });
});
