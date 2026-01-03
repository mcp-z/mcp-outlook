import assert from 'assert';
import { parseConfig } from '../../src/setup/config.ts';

describe('parseConfig', () => {
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

      const config = parseConfig([], env);

      assert.strictEqual(config.clientId, 'test-client-id');
      assert.strictEqual(config.clientSecret, 'test-client-secret');
      assert.strictEqual(config.tenantId, 'common');
      assert.strictEqual(config.auth, 'loopback-oauth');
    });

    it('parses config with optional client secret omitted', () => {
      const config = parseConfig([], baseEnv);

      assert.strictEqual(config.clientId, 'test-client-id');
      assert.strictEqual(config.clientSecret, undefined);
      assert.strictEqual(config.tenantId, 'common');
    });
  });

  describe('Authentication modes', () => {
    it('parses --auth=loopback-oauth', () => {
      const config = parseConfig(['--auth=loopback-oauth'], baseEnv);

      assert.strictEqual(config.auth, 'loopback-oauth');
      assert.strictEqual(config.dcrConfig, undefined);
    });

    it('parses --auth=device-code', () => {
      const config = parseConfig(['--auth=device-code'], baseEnv);

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

        const config = parseConfig(['--auth=dcr'], env);

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

        const config = parseConfig(['--auth=dcr', '--dcr-store-uri=file://custom-path/store.json'], env);

        assert.strictEqual(config.auth, 'dcr');
        assert.ok(config.dcrConfig);
        assert.strictEqual(config.dcrConfig.mode, 'self-hosted');
        assert.strictEqual(config.dcrConfig.storeUri, 'file://custom-path/store.json');
      });
    });

    describe('External DCR mode', () => {
      it('parses DCR mode with external configuration', () => {
        const env = {
          ...baseEnv,
          DCR_MODE: 'external',
          DCR_VERIFY_URL: 'https://auth.example.com/oauth/verify',
        };

        const config = parseConfig(['--auth=dcr'], env);

        assert.strictEqual(config.auth, 'dcr');
        assert.ok(config.dcrConfig);
        assert.strictEqual(config.dcrConfig.mode, 'external');
        assert.strictEqual(config.dcrConfig.verifyUrl, 'https://auth.example.com/oauth/verify');
        assert.strictEqual(config.dcrConfig.storeUri, undefined);
        assert.strictEqual(config.dcrConfig.clientId, 'test-client-id');
        assert.strictEqual(config.dcrConfig.tenantId, 'common');
      });

      it('parses DCR mode with CLI --dcr-mode=external', () => {
        const env = {
          ...baseEnv,
          DCR_VERIFY_URL: 'https://auth.example.com/oauth/verify',
        };

        const config = parseConfig(['--auth=dcr', '--dcr-mode=external'], env);

        assert.strictEqual(config.auth, 'dcr');
        assert.ok(config.dcrConfig);
        assert.strictEqual(config.dcrConfig.mode, 'external');
        assert.strictEqual(config.dcrConfig.verifyUrl, 'https://auth.example.com/oauth/verify');
      });

      it('parses DCR mode with CLI --dcr-verify-url', () => {
        const env = {
          ...baseEnv,
          DCR_MODE: 'external',
        };

        const config = parseConfig(['--auth=dcr', '--dcr-verify-url=https://new.example.com/verify'], env);

        assert.strictEqual(config.auth, 'dcr');
        assert.ok(config.dcrConfig);
        assert.strictEqual(config.dcrConfig.mode, 'external');
        assert.strictEqual(config.dcrConfig.verifyUrl, 'https://new.example.com/verify');
      });

      it('throws error when DCR_VERIFY_URL missing in external mode', () => {
        const env = {
          ...baseEnv,
          DCR_MODE: 'external',
        };

        assert.throws(() => parseConfig(['--auth=dcr'], env), {
          name: 'Error',
          message: 'DCR external mode requires --dcr-verify-url or DCR_VERIFY_URL environment variable',
        });
      });
    });

    describe('DCR mode defaults', () => {
      it('defaults to self-hosted mode when DCR_MODE not specified', () => {
        const env = {
          ...baseEnv,
          DCR_STORE_URI: 'file://.dcr.json',
        };

        const config = parseConfig(['--auth=dcr'], env);

        assert.strictEqual(config.auth, 'dcr');
        assert.ok(config.dcrConfig);
        assert.strictEqual(config.dcrConfig.mode, 'self-hosted');
        assert.strictEqual(config.dcrConfig.storeUri, 'file://.dcr.json');
      });
    });

    describe('DCR CLI overrides', () => {
      it('CLI --dcr-mode overrides DCR_MODE env var', () => {
        const env = {
          ...baseEnv,
          DCR_MODE: 'self-hosted',
          DCR_STORE_URI: 'file://.dcr.json',
          DCR_VERIFY_URL: 'https://auth.example.com/oauth/verify',
        };

        const config = parseConfig(['--auth=dcr', '--dcr-mode=external'], env);

        assert.strictEqual(config.auth, 'dcr');
        assert.ok(config.dcrConfig);
        assert.strictEqual(config.dcrConfig.mode, 'external');
        assert.strictEqual(config.dcrConfig.verifyUrl, 'https://auth.example.com/oauth/verify');
      });

      it('CLI --dcr-verify-url overrides DCR_VERIFY_URL env var', () => {
        const env = {
          ...baseEnv,
          DCR_MODE: 'external',
          DCR_VERIFY_URL: 'https://old.example.com/verify',
        };

        const config = parseConfig(['--auth=dcr', '--dcr-verify-url=https://new.example.com/verify'], env);

        assert.strictEqual(config.auth, 'dcr');
        assert.ok(config.dcrConfig);
        assert.strictEqual(config.dcrConfig.verifyUrl, 'https://new.example.com/verify');
      });

      it('CLI --dcr-store-uri overrides DCR_STORE_URI env var', () => {
        const env = {
          ...baseEnv,
          DCR_MODE: 'self-hosted',
          DCR_STORE_URI: 'file://old-path/store.json',
        };

        const config = parseConfig(['--auth=dcr', '--dcr-store-uri=file://new-path/store.json'], env);

        assert.strictEqual(config.auth, 'dcr');
        assert.ok(config.dcrConfig);
        assert.strictEqual(config.dcrConfig.storeUri, 'file://new-path/store.json');
      });
    });

    describe('Invalid DCR mode', () => {
      it('throws error for invalid --dcr-mode value', () => {
        const env = {
          ...baseEnv,
        };

        assert.throws(() => parseConfig(['--auth=dcr', '--dcr-mode=invalid'], env), {
          name: 'Error',
          message: 'Invalid --dcr-mode value: "invalid". Valid values: self-hosted, external',
        });
      });

      it('throws error for invalid DCR_MODE env var', () => {
        const env = {
          ...baseEnv,
          DCR_MODE: 'invalid',
        };

        assert.throws(() => parseConfig(['--auth=dcr'], env), {
          name: 'Error',
          message: 'Invalid --dcr-mode value: "invalid". Valid values: self-hosted, external',
        });
      });
    });
  });

  describe('Server configuration', () => {
    it('includes server metadata', () => {
      const config = parseConfig([], baseEnv);

      assert.ok(config.name);
      assert.ok(config.version);
      assert.ok(config.repositoryUrl);
      assert.ok(config.baseDir);
      assert.ok(config.storageDir);
    });

    it('parses transport configuration', () => {
      const config = parseConfig([], baseEnv);

      assert.ok(config.transport);
      assert.strictEqual(config.transport.type, 'stdio');
    });

    it('parses --port for HTTP transport', () => {
      const config = parseConfig(['--port=3456'], baseEnv);

      assert.strictEqual(config.transport.type, 'http');
      assert.strictEqual(config.transport.port, 3456);
    });
  });
});
