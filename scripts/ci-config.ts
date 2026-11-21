import type { CiConfig } from './lib/ci-setup.ts';

export const config = {
  provider: 'microsoft',
  flows: [
    {
      id: 'mcp-outlook',
      refreshTokenSecret: 'TEST_REFRESH_TOKEN',
      scopeSecret: 'TEST_SCOPE',
      clientIdSecret: 'MS_CLIENT_ID',
      clientSecretSecret: 'MS_CLIENT_SECRET',
      clientSecretRequired: false,
      tenantSecret: 'MS_TENANT_ID',
      storage: {
        kind: 'account',
        path: '.tokens/store.json',
        service: 'mcp-outlook',
        accountId: {
          env: 'TEST_ACCOUNT_ID',
        },
      },
    },
  ],
  requiredTestSecrets: [],
  serviceAccount: null,
} as const satisfies CiConfig;
