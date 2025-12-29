#!/usr/bin/env node

/**
 * Outlook Test Token Setup
 *
 * Generates OAuth token for Outlook server tests in local .tokens/ directory.
 *
 * Usage:
 *   npm run test:setup
 */

import { LoopbackOAuthProvider } from '@mcp-z/oauth-microsoft';
import { MS_SCOPE } from '../src/constants.ts';
import createStore from '../src/lib/create-store.ts';
import { createConfig } from '../src/setup/config.ts';

async function setupTest(): Promise<void> {
  const config = createConfig();
  console.log('🔐 Outlook Test Token Setup');
  console.log('');

  // Use local .tokens/ directory at package root (no subdirectories)
  const tokenStore = await createStore<unknown>('file://.//.tokens/store.json');

  const auth = new LoopbackOAuthProvider({
    service: config.name,
    clientId: config.clientId,
    clientSecret: config.clientSecret,
    tenantId: config.tenantId,
    scope: MS_SCOPE,
    headless: false,
    logger: console,
    tokenStore,
  });

  console.log('Starting OAuth flow...');
  console.log('');

  // Trigger OAuth flow - will fetch email and set as active account
  // OAuth flow will store token/account with email as accountId
  await auth.getAccessToken();

  console.log('✓ OAuth flow completed, fetching user email...');

  // Get email for display (from active account)
  const email = await auth.getUserEmail();

  console.log('');
  console.log('✅ OAuth token generated successfully!');
  console.log(`📧 Authenticated as: ${email}`);
  console.log('📁 Token saved to: .tokens/store.json');
  console.log(`   Token key: ${email}:outlook:token`);
  console.log('');
  console.log('Run `npm run test:unit` to verify Outlook API integration');
}

// Run if executed directly
if (import.meta.main) {
  setupTest()
    .then(() => {
      process.exit(0);
    })
    .catch((error) => {
      console.error('\n❌ Token setup failed:', error.message);
      process.exit(1);
    });
}
