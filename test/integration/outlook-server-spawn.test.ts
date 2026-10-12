/**
 * Outlook Server Spawn Integration Test
 *
 */

import '../lib/env-loader.ts';
import { createServerRegistry, type ManagedClient, type ServerRegistry } from '@mcp-z/client';
import assert from 'assert';
import { randomUUID } from 'crypto';
import * as fs from 'fs';
import { safeRmSync } from 'fs-remove-compat';
import * as path from 'path';
import { requiredEnv } from 'portable-env';
import { fileURLToPath, pathToFileURL } from 'url';
import { throwFailures } from '../lib/throw-failures.ts';

describe('Outlook Server Spawn Integration', () => {
  let client: ManagedClient;
  let registry: ServerRegistry | undefined;
  let projectDir: string | undefined;
  let spawnAttempted = false;

  before(async () => {
    const serverRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
    const serverPath = path.join(serverRoot, 'bin/server.js');

    // The discovered project owns logs and stores instead of the user's home directory.
    projectDir = path.join(serverRoot, '.tmp', `outlook-spawn-test-${randomUUID()}`);
    fs.mkdirSync(projectDir, { recursive: true });
    fs.writeFileSync(path.join(projectDir, '.mcp.json'), '{ "mcpServers": {} }\n');
    const stateDir = path.join(projectDir, '.mcp-z');

    spawnAttempted = true;
    registry = createServerRegistry(
      {
        outlook: {
          command: process.execPath,
          args: [serverPath, '--headless', '--stdio'],
          env: {
            NODE_ENV: 'test',
            MS_CLIENT_ID: requiredEnv('MS_CLIENT_ID'),
            MS_CLIENT_SECRET: process.env.MS_CLIENT_SECRET || '',
            MS_TENANT_ID: process.env.MS_TENANT_ID || 'common',
            HEADLESS: 'true',
            LOG_LEVEL: 'error',
            TOKEN_STORE_URI: pathToFileURL(path.join(stateDir, 'tokens.json')).href,
            DCR_STORE_URI: pathToFileURL(path.join(stateDir, 'dcr.json')).href,
            RESOURCE_STORE_URI: pathToFileURL(path.join(stateDir, 'files')).href,
          },
        },
      },
      { cwd: projectDir }
    );

    client = await registry.connect('outlook');
  });

  after(async () => {
    const failures: unknown[] = [];
    // The registry owns both the client and the child; only its resolved close proves exit.
    let childClosed = !spawnAttempted;
    if (registry) {
      try {
        const result = await registry.close();
        childClosed = true;
        assert.deepStrictEqual({ timedOut: result.timedOut, killedCount: result.killedCount }, { timedOut: false, killedCount: 0 }, 'Server should shut down cooperatively');
      } catch (error) {
        failures.push(error);
      }
    }
    if (projectDir && childClosed) {
      try {
        safeRmSync(projectDir, { recursive: true, force: true });
      } catch (error) {
        failures.push(error);
      }
    } else if (projectDir) {
      failures.push(new Error(`Server closure unconfirmed; retained scratch ${projectDir}`));
    }
    throwFailures('Outlook spawn test teardown failed', failures);
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
