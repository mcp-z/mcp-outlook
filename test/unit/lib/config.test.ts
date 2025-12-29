import assert from 'assert';
import { parseConfig } from '../../../src/setup/config.js';

describe('parseConfig', () => {
  it('defaults to stdio transport with no args or env', () => {
    const config = parseConfig([], {
      MS_CLIENT_ID: 'test-client-id',
      MS_CLIENT_SECRET: 'test-client-secret',
      MS_TENANT_ID: 'test-tenant-id',
    });

    assert.strictEqual(config.transport.type, 'stdio');
  });

  it('defaults headless to true for tests', () => {
    const config = parseConfig([], {
      MS_CLIENT_ID: 'test-client-id',
      MS_CLIENT_SECRET: 'test-client-secret',
      MS_TENANT_ID: 'test-tenant-id',
      HEADLESS: 'true', // Explicit HEADLESS env var (no NODE_ENV magic)
    });

    assert.strictEqual(config.headless, true);
  });

  it('uses --headless CLI arg to override env var', () => {
    const config = parseConfig(['--headless'], {
      MS_CLIENT_ID: 'test-client-id',
      MS_CLIENT_SECRET: 'test-client-secret',
      MS_TENANT_ID: 'test-tenant-id',
      HEADLESS: 'false', // Env says false, but CLI arg overrides
    });

    // CLI arg --headless should override HEADLESS env var
    assert.strictEqual(config.headless, true);
  });

  it('parses config from env object parameter', () => {
    const testEnv = {
      MS_CLIENT_ID: 'test-client-id',
      MS_CLIENT_SECRET: 'test-client-secret',
      MS_TENANT_ID: 'test-tenant-id',
    };

    const config = parseConfig([], testEnv);

    assert.strictEqual(config.clientId, 'test-client-id');
    assert.strictEqual(config.tenantId, 'test-tenant-id');
  });

  it('uses empty array for args when args parameter is undefined', () => {
    const config = parseConfig([], {
      MS_CLIENT_ID: 'test-client-id',
      MS_CLIENT_SECRET: 'test-client-secret',
      MS_TENANT_ID: 'test-tenant-id',
    });

    // Should parse successfully without CLI args
    assert.strictEqual(config.transport.type, 'stdio');
  });

  it('parses HTTP port from env in test config', () => {
    const testPort = 5568; // Explicit test port
    const config = parseConfig([], {
      PORT: testPort.toString(),
      MS_CLIENT_ID: 'test-client-id',
      MS_CLIENT_SECRET: 'test-client-secret',
      MS_TENANT_ID: 'test-tenant-id',
    });

    assert.strictEqual(config.transport.type, 'http');
    assert.strictEqual(config.transport.port, testPort);
    // redirectUri is only set when explicitly provided via --redirect-uri
    assert.strictEqual(config.redirectUri, undefined);
  });

  it('parses HTTP port from CLI --port flag (overrides env)', () => {
    const envPort = 5568;
    const cliPort = 5569;
    const config = parseConfig([`--port=${cliPort}`], {
      PORT: envPort.toString(), // Env var should be overridden by CLI flag
      MS_CLIENT_ID: 'test-client-id',
      MS_CLIENT_SECRET: 'test-client-secret',
      MS_TENANT_ID: 'test-tenant-id',
    });

    assert.strictEqual(config.transport.type, 'http');
    assert.strictEqual(config.transport.port, cliPort); // CLI flag wins
    // redirectUri is only set when explicitly provided via --redirect-uri
    assert.strictEqual(config.redirectUri, undefined);
  });

  it('parses --redirect-uri when explicitly provided', () => {
    const config = parseConfig(['--redirect-uri=https://example.com/callback'], {
      MS_CLIENT_ID: 'test-client-id',
      MS_CLIENT_SECRET: 'test-client-secret',
      MS_TENANT_ID: 'test-tenant-id',
    });

    assert.strictEqual(config.redirectUri, 'https://example.com/callback');
  });

  it('parses --stdio explicitly', () => {
    const config = parseConfig(['--stdio'], {
      MS_CLIENT_ID: 'test-client-id',
      MS_CLIENT_SECRET: 'test-client-secret',
      MS_TENANT_ID: 'test-tenant-id',
    });

    assert.strictEqual(config.transport.type, 'stdio');
  });

  it('defaults to loopback-oauth auth mode', () => {
    const config = parseConfig([], {
      MS_CLIENT_ID: 'test-client-id',
      MS_CLIENT_SECRET: 'test-client-secret',
      MS_TENANT_ID: 'test-tenant-id',
    });

    assert.strictEqual(config.auth, 'loopback-oauth');
  });

  it('defaults to loopback-oauth auth mode', () => {
    const config = parseConfig([], {
      MS_CLIENT_ID: 'test-client-id',
      MS_CLIENT_SECRET: 'test-client-secret',
      MS_TENANT_ID: 'test-tenant-id',
    });

    assert.strictEqual(config.auth, 'loopback-oauth');
  });

  it('parses --auth=loopback-oauth', () => {
    const config = parseConfig(['--auth=loopback-oauth'], {
      MS_CLIENT_ID: 'test-client-id',
      MS_CLIENT_SECRET: 'test-client-secret',
      MS_TENANT_ID: 'test-tenant-id',
    });

    assert.strictEqual(config.auth, 'loopback-oauth');
  });

  it('defaults logLevel to info', () => {
    const config = parseConfig([], {
      MS_CLIENT_ID: 'test-client-id',
      MS_CLIENT_SECRET: 'test-client-secret',
      MS_TENANT_ID: 'test-tenant-id',
    });

    assert.strictEqual(config.logLevel, 'info');
  });

  it('parses LOG_LEVEL from env', () => {
    const config = parseConfig([], {
      MS_CLIENT_ID: 'test-client-id',
      MS_CLIENT_SECRET: 'test-client-secret',
      MS_TENANT_ID: 'test-tenant-id',
      LOG_LEVEL: 'debug',
    });

    assert.strictEqual(config.logLevel, 'debug');
  });

  it('parses --log-level from CLI', () => {
    const config = parseConfig(['--log-level=error'], {
      MS_CLIENT_ID: 'test-client-id',
      MS_CLIENT_SECRET: 'test-client-secret',
      MS_TENANT_ID: 'test-tenant-id',
    });

    assert.strictEqual(config.logLevel, 'error');
  });

  it('CLI --log-level overrides LOG_LEVEL env var', () => {
    const config = parseConfig(['--log-level=warn'], {
      MS_CLIENT_ID: 'test-client-id',
      MS_CLIENT_SECRET: 'test-client-secret',
      MS_TENANT_ID: 'test-tenant-id',
      LOG_LEVEL: 'debug',
    });

    assert.strictEqual(config.logLevel, 'warn');
  });
});
