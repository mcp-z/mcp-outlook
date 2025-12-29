import type { Logger, MicrosoftAuthProvider } from '@mcp-z/oauth-microsoft';
import { Client } from '@microsoft/microsoft-graph-client';
import assert from 'assert';
import type { ZodTypeAny } from 'zod';
import categoriesFactory, { type Output as CategoriesOutput } from '../../../../src/mcp/tools/categories-list.js';
import createTool, { type Input, type Output } from '../../../../src/mcp/tools/labels-list.js';
import { createTestCategory, deleteTestCategory } from '../../../lib/category-helpers.js';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.js';
import createMiddlewareContext from '../../../lib/create-middleware-context.js';
import waitForCategory from '../../../lib/wait-for-category.js';

// Type for objects with a parse method (like Zod schemas)
type SchemaLike = { parse: (data: unknown) => unknown };

// Type for objects with an id property
type ItemWithId = { id?: string; [key: string]: unknown };

describe('outlook-labels-list tool', () => {
  let auth: MicrosoftAuthProvider;
  let logger: Logger;
  let tool: ReturnType<typeof createTool>;
  let wrappedTool: ReturnType<Awaited<ReturnType<typeof createMiddlewareContext>>['middleware']['withToolAuth']>;
  let handler: TypedHandler<Input>;
  let middleware: Awaited<ReturnType<typeof createMiddlewareContext>>['middleware'];
  let graph: Client;

  before(async () => {
    const middlewareContext = await createMiddlewareContext();
    auth = middlewareContext.auth;
    logger = middlewareContext.logger;
    middleware = middlewareContext.middleware;
    tool = createTool();
    wrappedTool = middleware.withToolAuth(tool);
    handler = wrappedTool.handler as TypedHandler<Input>;
    graph = await Client.initWithMiddleware({ authProvider: auth });
  });

  it('labels_list returns structured success or error (service-backed)', async () => {
    const res = await handler({}, createExtra());

    // Basic MCP response structure validation
    assert.ok(res && Array.isArray(res.content), 'labels_list did not return content array');
    assert.ok(res.structuredContent && res.structuredContent, 'missing structuredContent');

    // Validate against declared output schema
    const schema = tool.config?.outputSchema;
    assert.ok(schema, 'tool.outputSchema missing from tool metadata');

    try {
      (schema as SchemaLike).parse(res.structuredContent);
    } catch (err) {
      const message = err instanceof Error && 'issues' in err ? JSON.stringify((err as { issues: unknown }).issues) : String(err);
      assert.fail(`structuredContent failed schema validation: ${message}`);
    }

    const result = res.structuredContent?.result as Output | undefined;

    // Test different response types based on actual result
    if (result?.type === 'success') {
      // Validate success response structure
      assert.ok(Array.isArray(result.items), 'success response should have items array');

      // If labels/categories exist, validate their structure
      if (result.items.length > 0) {
        const label = result.items[0];
        if (label) {
          assert.ok(typeof label.id === 'string', 'label should have string id');
          assert.ok(typeof label.displayName === 'string', 'label should have string displayName');
          assert.ok(typeof label.color === 'string', 'label should have string color');

          // Validate color is a valid Outlook preset color
          const validColors = [
            'preset0',
            'preset1',
            'preset2',
            'preset3',
            'preset4',
            'preset5',
            'preset6',
            'preset7',
            'preset8',
            'preset9',
            'preset10',
            'preset11',
            'preset12',
            'preset13',
            'preset14',
            'preset15',
            'preset16',
            'preset17',
            'preset18',
            'preset19',
            'preset20',
            'preset21',
            'preset22',
            'preset23',
            'preset24',
          ];
          assert.ok(validColors.includes(label.color), `label color should be valid preset: ${label.color}`);

          // Verify case-sensitive handling - displayName should preserve exact case
          assert.strictEqual(typeof label.displayName, 'string');
          assert.ok(label.displayName.length > 0);
        }
      }
    } else if (result?.type === 'auth_required') {
      // Validate auth_required response
      assert.ok(typeof result.provider === 'string', 'auth_required should have provider');
      assert.ok(typeof result.message === 'string', 'auth_required should have message');
      // url is optional in auth_required
      if (result.url) {
        assert.ok(typeof result.url === 'string', 'auth_required url should be string');
      }
    }
  });

  it('validates tool metadata and configuration', () => {
    // Validate tool name follows convention
    assert.ok(tool.config.description?.includes('label'), 'description should mention labels');
    assert.ok(tool.config.description?.includes('categories'), 'description should explain that Outlook uses categories as labels');

    // Validate input schema (should be empty object for this tool)
    assert.ok(tool.config.inputSchema, 'tool should have inputSchema');
    // Cast to ZodTypeAny to access safeParse method - our tooling always uses Zod schemas
    const parsedEmpty = (tool.config.inputSchema as ZodTypeAny).safeParse({});
    assert.ok(parsedEmpty.success, 'inputSchema should accept empty object for labels-list');

    // Validate output schema structure (reuses OutlookCategorySchema)
    assert.ok(tool.config.outputSchema, 'tool should have outputSchema');
    const schema = tool.config?.outputSchema;
    assert.ok(schema, 'outputSchema should be defined');
  });

  it('tool follows MCP naming convention', () => {
    // Tool module should export a name that matches expected pattern
    assert.strictEqual(tool.name, 'labels-list', 'tool name should follow {resource}-{action} pattern');
  });

  it('tool returns same data as categories-list (since they are equivalent in Outlook)', async () => {
    // Create a unique test category to verify equivalence
    const testName = `ci-test-equiv-${Date.now()}`;
    const testId = await createTestCategory(graph, { displayName: testName });

    try {
      // Wait for category to be indexed
      await waitForCategory(graph, testId);

      // Get results from both tools
      const labelsResult = await handler({}, createExtra());
      const rawCategoriesTool = categoriesFactory();
      const categoriesTool = middleware.withToolAuth(rawCategoriesTool);
      const categoriesResult = await categoriesTool.handler({}, createExtra());

      // Extract structured content
      const labelsPayload = labelsResult.structuredContent?.result as Output | undefined;
      const categoriesPayload = categoriesResult.structuredContent?.result as CategoriesOutput | undefined;

      assert.ok(labelsPayload, 'labels result should have structuredContent');
      assert.ok(categoriesPayload, 'categories result should have structuredContent');

      // Both should succeed
      assert.strictEqual(labelsPayload?.type, 'success', 'labels should succeed');
      assert.strictEqual(categoriesPayload?.type, 'success', 'categories should succeed');

      // Find our test category in both results
      const labelItem = labelsPayload?.type === 'success' ? labelsPayload.items.find((i: ItemWithId) => i.id === testId) : undefined;
      const categoryItem = categoriesPayload?.type === 'success' ? categoriesPayload.items.find((i: ItemWithId) => i.id === testId) : undefined;

      // Both should have our test category
      assert.ok(labelItem, 'labels should return test category');
      assert.ok(categoryItem, 'categories should return test category');

      // The items should be identical
      assert.deepStrictEqual(labelItem, categoryItem, 'same category should have identical data in both endpoints');
    } finally {
      // Cleanup test category
      await deleteTestCategory(graph, testId, logger);
    }
  });
});
