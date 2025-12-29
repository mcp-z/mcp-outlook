/**
 * Outlook Server Spawn Integration Test
 *
 */

import { createServerRegistry, type ManagedClient, type ServerRegistry } from '@mcp-z/client';
import assert from 'assert';

describe('Outlook Server Spawn Integration', () => {
  let client: ManagedClient;
  let registry: ServerRegistry;

  before(async () => {
    registry = createServerRegistry(
      {
        outlook: {
          command: 'node',
          args: ['bin/server.js', '--headless'],
          env: {
            NODE_ENV: 'test',
            MS_CLIENT_ID: process.env.MS_CLIENT_ID || '',
            MS_CLIENT_SECRET: process.env.MS_CLIENT_SECRET || '',
            MS_TENANT_ID: process.env.MS_TENANT_ID || 'common',
            HEADLESS: 'true',
            LOG_LEVEL: 'error',
          },
        },
      },
      { cwd: process.cwd() }
    );

    client = await registry.connect('outlook');
  });

  after(async () => {
    if (client) await client.close();
    if (registry) await registry.close();
  });

  it('should connect to Outlook server', async () => {
    // Client is already connected via registry.connect() in before hook
    assert.ok(client, 'Should have connected Outlook client');
  });

  it('should list tools via MCP protocol', async () => {
    const result = await client.listTools();

    assert.ok(result.tools, 'Should return tools');
    assert.ok(result.tools.length > 0, 'Should have at least one tool');

    // Verify specific tools exist
    const includes = (name: string) => result.tools.some((t) => t.name.includes(name));
    assert.ok(includes('message-search'), 'Should have message-search tool');
    assert.ok(includes('message-get'), 'Should have message-get tool');
    assert.ok(includes('message-send'), 'Should have message-send tool');
    assert.ok(includes('labels-list'), 'Should have labels-list tool');
  });
});
