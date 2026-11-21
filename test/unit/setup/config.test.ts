import '../../lib/env-loader.ts';
import { setup } from '@mcp-z/mcp-outlook';
import assert from 'assert';
import * as fs from 'fs';
import * as path from 'path';

describe('setup.parseConfig', () => {
  const baseEnv = {
    MS_CLIENT_ID: 'test-client-id',
    MS_TENANT_ID: 'common',
  };

  describe('Basic OAuth configuration', () => {
    it('parses config with all OAuth environment variables', () => {
      const env = {
        ...baseEnv,
        MS_CLIENT_SECRET: 'test-client-secret',
      };

      const config = setup.parseConfig([], env);

      assert.strictEqual(config.clientId, 'test-client-id');
      assert.strictEqual(config.clientSecret, 'test-client-secret');
      assert.strictEqual(config.tenantId, 'common');
      assert.strictEqual(config.auth, 'loopback-oauth');
    });

    it('parses config with optional client secret omitted', () => {
      const config = setup.parseConfig([], baseEnv);

      assert.strictEqual(config.clientId, 'test-client-id');
      assert.strictEqual(config.clientSecret, undefined);
      assert.strictEqual(config.tenantId, 'common');
    });
  });

  describe('Authentication modes', () => {
    it('parses --auth=loopback-oauth', () => {
      const config = setup.parseConfig(['--auth=loopback-oauth'], baseEnv);

      assert.strictEqual(config.auth, 'loopback-oauth');
      assert.strictEqual(config.dcrConfig, undefined);
    });

    it('parses --auth=device-code', () => {
      const config = setup.parseConfig(['--auth=device-code'], baseEnv);

      assert.strictEqual(config.auth, 'device-code');
      assert.strictEqual(config.dcrConfig, undefined);
    });
  });

  describe('DCR mode configuration', () => {
    describe('Self-hosted DCR mode', () => {
      it('parses DCR mode with self-hosted configuration', () => {
        const env = {
          ...baseEnv,
          DCR_MODE: 'self-hosted',
          DCR_STORE_URI: 'file://.dcr.json',
        };

        const config = setup.parseConfig(['--auth=dcr'], env);

        assert.strictEqual(config.auth, 'dcr');
        assert.ok(config.dcrConfig);
        assert.strictEqual(config.dcrConfig.mode, 'self-hosted');
        assert.strictEqual(config.dcrConfig.storeUri, 'file://.dcr.json');
        assert.strictEqual(config.dcrConfig.verifyUrl, undefined);
        assert.strictEqual(config.dcrConfig.clientId, 'test-client-id');
        assert.strictEqual(config.dcrConfig.tenantId, 'common');
      });

      it('parses DCR mode with CLI --dcr-store-uri', () => {
        const env = {
          ...baseEnv,
        };

        const config = setup.parseConfig(['--auth=dcr', '--dcr-store-uri=file://custom-path/store.json'], env);

        assert.strictEqual(config.auth, 'dcr');
        assert.ok(config.dcrConfig);
        assert.strictEqual(config.dcrConfig.mode, 'self-hosted');
        assert.strictEqual(config.dcrConfig.storeUri, 'file://custom-path/store.json');
      });
    });

    describe('External DCR mode is refused with --auth=dcr', () => {
      const notSupported = { name: 'Error', message: /^external DCR mode is not supported/ };
      const verifyUrlNotSupported = { name: 'Error', message: /DCR_VERIFY_URL is not supported/ };

      it('refuses DCR_MODE=external', () => {
        assert.throws(() => setup.parseConfig(['--auth=dcr'], { ...baseEnv, DCR_MODE: 'external', DCR_VERIFY_URL: 'https://auth.example.com/oauth/verify' }), notSupported);
      });

      it('refuses --dcr-mode=external', () => {
        assert.throws(() => setup.parseConfig(['--auth=dcr', '--dcr-mode=external'], { ...baseEnv, DCR_VERIFY_URL: 'https://auth.example.com/oauth/verify' }), notSupported);
      });

      it('refuses external mode without a verify URL with the same error', () => {
        assert.throws(() => setup.parseConfig(['--auth=dcr'], { ...baseEnv, DCR_MODE: 'external' }), notSupported);
      });

      it('refuses --dcr-mode=external over DCR_MODE=self-hosted', () => {
        assert.throws(() => setup.parseConfig(['--auth=dcr', '--dcr-mode=external'], { ...baseEnv, DCR_MODE: 'self-hosted', DCR_STORE_URI: 'file://.dcr.json' }), notSupported);
      });

      it('refuses --dcr-verify-url', () => {
        assert.throws(() => setup.parseConfig(['--auth=dcr', '--dcr-verify-url=https://new.example.com/verify'], baseEnv), verifyUrlNotSupported);
      });

      it('refuses DCR_VERIFY_URL', () => {
        assert.throws(() => setup.parseConfig(['--auth=dcr'], { ...baseEnv, DCR_VERIFY_URL: 'https://old.example.com/verify' }), verifyUrlNotSupported);
      });

      it('does not refuse stale external settings outside --auth=dcr', () => {
        const config = setup.parseConfig(['--auth=loopback-oauth'], { ...baseEnv, DCR_MODE: 'external', DCR_VERIFY_URL: 'https://auth.example.com/oauth/verify' });
        assert.strictEqual(config.auth, 'loopback-oauth');
        assert.strictEqual(config.dcrConfig, undefined);
      });
    });

    describe('DCR mode defaults', () => {
      it('defaults to self-hosted mode when DCR_MODE not specified', () => {
        const config = setup.parseConfig(['--auth=dcr'], { ...baseEnv, DCR_STORE_URI: 'file://.dcr.json' });

        assert.strictEqual(config.auth, 'dcr');
        assert.ok(config.dcrConfig);
        assert.strictEqual(config.dcrConfig.mode, 'self-hosted');
        assert.strictEqual(config.dcrConfig.storeUri, 'file://.dcr.json');
      });

      it('accepts --dcr-mode=self-hosted given explicitly', () => {
        const config = setup.parseConfig(['--auth=dcr', '--dcr-mode=self-hosted'], baseEnv);
        assert.strictEqual(config.dcrConfig?.mode, 'self-hosted');
      });
    });

    describe('DCR CLI overrides', () => {
      it('CLI --dcr-store-uri overrides DCR_STORE_URI env var', () => {
        const env = {
          ...baseEnv,
          DCR_MODE: 'self-hosted',
          DCR_STORE_URI: 'file://old-path/store.json',
        };

        const config = setup.parseConfig(['--auth=dcr', '--dcr-store-uri=file://new-path/store.json'], env);

        assert.strictEqual(config.auth, 'dcr');
        assert.ok(config.dcrConfig);
        assert.strictEqual(config.dcrConfig.storeUri, 'file://new-path/store.json');
      });
    });

    describe('Invalid DCR mode', () => {
      const invalid = { name: 'Error', message: 'Invalid DCR mode "invalid": only self-hosted is supported' };

      it('throws error for invalid --dcr-mode value', () => {
        assert.throws(() => setup.parseConfig(['--auth=dcr', '--dcr-mode=invalid'], baseEnv), invalid);
      });

      it('throws error for invalid DCR_MODE env var', () => {
        assert.throws(() => setup.parseConfig(['--auth=dcr'], { ...baseEnv, DCR_MODE: 'invalid' }), invalid);
      });
    });
  });

  describe('Server configuration', () => {
    it('includes server metadata', () => {
      const config = setup.parseConfig([], baseEnv);

      assert.ok(config.name);
      assert.ok(config.version);
      assert.ok(config.repositoryUrl);
      assert.ok(config.baseDir);
      assert.ok(config.resourceStoreUri);
    });

    it('parses transport configuration', () => {
      const config = setup.parseConfig([], baseEnv);

      assert.ok(config.transport);
      assert.strictEqual(config.transport.type, 'stdio');
    });

    it('parses --port for HTTP transport', () => {
      const config = setup.parseConfig(['--port=3456'], baseEnv);

      assert.strictEqual(config.transport.type, 'http');
      assert.strictEqual(config.transport.port, 3456);
    });
  });

  it('defaults to stdio transport with no args or env', () => {
    const config = setup.parseConfig([], baseEnv);

    assert.strictEqual(config.transport.type, 'stdio');
  });

  it('defaults headless to true for tests', () => {
    const env = {
      ...baseEnv,
      HEADLESS: 'true', // Explicit HEADLESS env var (no NODE_ENV magic)
    };

    const config = setup.parseConfig([], env);

    assert.strictEqual(config.headless, true);
  });

  it('uses --headless CLI arg to override env var', () => {
    const env = {
      ...baseEnv,
      HEADLESS: 'false', // Env says false, but CLI arg overrides
    };

    const config = setup.parseConfig(['--headless'], env);

    // CLI arg --headless should override HEADLESS env var
    assert.strictEqual(config.headless, true);
  });

  it('parses config from env object parameter', () => {
    const testEnv = {
      MS_CLIENT_ID: 'test-client-id',
      MS_CLIENT_SECRET: 'test-client-secret',
      MS_TENANT_ID: 'test-tenant-id',
    };

    const config = setup.parseConfig([], testEnv);

    assert.strictEqual(config.clientId, 'test-client-id');
    assert.strictEqual(config.tenantId, 'test-tenant-id');
  });

  it('uses empty array for args when args parameter is undefined', () => {
    const config = setup.parseConfig([], baseEnv);

    // Should parse successfully without CLI args
    assert.strictEqual(config.transport.type, 'stdio');
  });

  it('parses HTTP port from env in test config', () => {
    const testPort = 5568; // Explicit test port
    const env = {
      ...baseEnv,
      PORT: testPort.toString(),
    };

    const config = setup.parseConfig([], env);

    assert.strictEqual(config.transport.type, 'http');
    assert.strictEqual(config.transport.port, testPort);
    // redirectUri is only set when explicitly provided via --redirect-uri
    assert.strictEqual(config.redirectUri, undefined);
  });

  it('parses HTTP port from CLI --port flag (overrides env)', () => {
    const envPort = 5568;
    const cliPort = 5569;
    const env = {
      ...baseEnv,
      PORT: envPort.toString(), // Env var should be overridden by CLI flag
    };

    const config = setup.parseConfig([`--port=${cliPort}`], env);

    assert.strictEqual(config.transport.type, 'http');
    assert.strictEqual(config.transport.port, cliPort); // CLI flag wins
    // redirectUri is only set when explicitly provided via --redirect-uri
    assert.strictEqual(config.redirectUri, undefined);
  });

  it('parses --redirect-uri when explicitly provided', () => {
    const config = setup.parseConfig(['--redirect-uri=https://example.com/callback'], baseEnv);

    assert.strictEqual(config.redirectUri, 'https://example.com/callback');
  });

  it('parses --stdio explicitly', () => {
    const config = setup.parseConfig(['--stdio'], baseEnv);

    assert.strictEqual(config.transport.type, 'stdio');
  });

  it('defaults to loopback-oauth auth mode', () => {
    const config = setup.parseConfig([], baseEnv);

    assert.strictEqual(config.auth, 'loopback-oauth');
  });

  it('defaults to loopback-oauth auth mode', () => {
    const config = setup.parseConfig([], baseEnv);

    assert.strictEqual(config.auth, 'loopback-oauth');
  });

  it('defaults logLevel to info', () => {
    const config = setup.parseConfig([], baseEnv);

    assert.strictEqual(config.logLevel, 'info');
  });

  it('parses LOG_LEVEL from env', () => {
    const env = {
      ...baseEnv,
      LOG_LEVEL: 'debug',
    };

    const config = setup.parseConfig([], env);

    assert.strictEqual(config.logLevel, 'debug');
  });

  it('parses --log-level from CLI', () => {
    const config = setup.parseConfig(['--log-level=error'], baseEnv);

    assert.strictEqual(config.logLevel, 'error');
  });

  it('CLI --log-level overrides LOG_LEVEL env var', () => {
    const env = {
      ...baseEnv,
      LOG_LEVEL: 'debug',
    };

    const config = setup.parseConfig(['--log-level=warn'], env);

    assert.strictEqual(config.logLevel, 'warn');
  });

  describe('server.json DCR settings', () => {
    type Entry = { name?: string };
    const manifest = JSON.parse(fs.readFileSync(path.resolve('server.json'), 'utf8')) as { packages: { environmentVariables: Entry[]; packageArguments: Entry[] }[] };

    it('lists no external DCR setting in either transport', () => {
      const names = manifest.packages.flatMap((p) => [...p.environmentVariables, ...p.packageArguments].map((e) => e.name));
      for (const setting of ['--dcr-mode', '--dcr-verify-url', 'DCR_MODE', 'DCR_VERIFY_URL']) assert.ok(!names.includes(setting), `server.json still lists ${setting}`);
    });
  });
});
