#!/usr/bin/env node

/**
 * Outlook Mail Cleaner
 *
 * Uses loopback OAuth to authenticate and clean (move to trash) all emails.
 * Always requests new tokens, does not store them persistently.
 *
 * Usage:
 *   node scripts/clean.ts           # Preview mode: lists messages that would be deleted
 *   node scripts/clean.ts --force   # Force mode: actually deletes the messages
 *   node scripts/clean.ts --headless # Prints URL for CI/SSH environments
 */

import { LoopbackOAuthProvider } from '@mcp-z/oauth-microsoft';
import { Client } from '@microsoft/microsoft-graph-client';
import Keyv from 'keyv';
import { MS_SCOPE } from '../src/constants.ts';
import { createConfig } from '../src/setup/config.ts';

const CHUNK_SIZE = 100; // Process in chunks to avoid memory issues
const MAX_BATCH_SIZE = 10000; // Stop after this many messages to prevent runaway

async function cleanMail(): Promise<void> {
  // Parse command line arguments
  const args = process.argv.slice(2);
  const isForce = args.includes('--force');
  const isHeadless = args.includes('--headless');

  const config = createConfig();
  // Create modified config for headless
  const modifiedConfig = { ...config, headless: config.headless || isHeadless };

  console.log('🧹 Outlook Mail Cleaner');
  console.log('');

  if (isForce) {
    console.log('⚠️  WARNING: FORCE MODE - This will move ALL your emails to trash!');
    console.log('   Make sure you have a backup or are absolutely sure.');
  } else {
    console.log('👀 PREVIEW MODE: This will list messages that WOULD be deleted.');
    console.log('   Run with --force to actually delete them.');
  }
  console.log('');

  // Use in-memory store so tokens are not persisted
  const tokenStore = new Keyv();

  const auth = new LoopbackOAuthProvider({
    service: config.name,
    clientId: config.clientId,
    clientSecret: config.clientSecret,
    tenantId: config.tenantId || 'common',
    scope: MS_SCOPE,
    headless: modifiedConfig.headless,
    logger: console,
    tokenStore, // In-memory, not persistent
  });

  console.log('Starting OAuth flow...');
  console.log('');

  try {
    // Get fresh access token (always requests new since in-memory store)
    const token = await auth.getAccessToken('temp-user');

    console.log('✓ Authentication successful!');
    console.log('');

    // Create Graph client with simple token auth provider
    const graph = Client.initWithMiddleware({
      authProvider: {
        getAccessToken: async () => token,
      },
    });

    console.log('Searching for all messages...');

    // Collect all messages by paginating through all messages (include subject for preview)
    const allMessages: Array<{ id: string; subject: string }> = [];
    let pageCount = 0;
    let hasNextPage = true;
    let currentUrl = '/me/messages?$select=id,subject&$top=500';

    while (hasNextPage && allMessages.length < MAX_BATCH_SIZE) {
      pageCount++;
      console.log(`Fetching page ${pageCount}...`);

      const request = graph.api(currentUrl);
      const response = await request.get();
      const messages = response.value || [];

      for (const message of messages) {
        if (message.id && message.subject !== undefined) {
          allMessages.push({
            id: message.id,
            subject: message.subject || '(No subject)',
          });
        }
      }

      const nextLink = response['@odata.nextLink'];
      hasNextPage = !!nextLink;
      currentUrl = nextLink || currentUrl;

      console.log(`  Found ${messages.length} messages in this page (total: ${allMessages.length})`);

      // Safety check
      if (allMessages.length >= MAX_BATCH_SIZE) {
        console.log(`⚠️  Reached maximum batch size limit (${MAX_BATCH_SIZE}). Stopping at ${allMessages.length} messages.`);
        hasNextPage = false;
      }

      // Safety check: if no messages in last few pages but still hasNextPage, might be stuck
      if (pageCount > 10 && messages.length === 0) {
        console.log('⚠️  No more messages found. Stopping.');
        hasNextPage = false;
      }
    }

    console.log('');
    console.log(`📧 Found ${allMessages.length} messages total (${pageCount} pages)`);

    if (allMessages.length === 0) {
      console.log('✅ No messages to clean!');
      return;
    }

    if (!isForce) {
      // Preview mode: show first 10 messages and summary
      console.log('');
      console.log('📋 Preview of messages that would be deleted:');
      console.log('');

      const previewCount = Math.min(10, allMessages.length);
      for (let i = 0; i < previewCount; i++) {
        const msg = allMessages[i];
        console.log(`  ${i + 1}. [${msg.id}] ${msg.subject}`);
      }

      if (allMessages.length > 10) {
        console.log(`  ... and ${allMessages.length - 10} more messages`);
      }

      console.log('');
      console.log(`💡 To delete these ${allMessages.length} messages, run with --force flag:`);
      console.log('   node scripts/clean.ts --force');

      return;
    }

    // Force mode: actually delete
    console.log('');
    console.log(`🗑️  Deleting ${allMessages.length} messages...`);

    const messageIds = allMessages.map((m) => m.id);
    let totalSuccess = 0;
    let totalFailure = 0;

    for (let i = 0; i < messageIds.length; i += CHUNK_SIZE) {
      const chunk = messageIds.slice(i, i + CHUNK_SIZE);
      console.log(`Processing chunk ${Math.floor(i / CHUNK_SIZE) + 1}/${Math.ceil(messageIds.length / CHUNK_SIZE)}: ${chunk.length} messages`);

      const chunkResults = await Promise.allSettled(
        chunk.map(async (id) => {
          await graph.api(`/me/messages/${encodeURIComponent(id)}`).delete();
          return { id, success: true };
        })
      );

      const successCount = chunkResults.filter((r) => r.status === 'fulfilled').length;
      const failureCount = chunkResults.filter((r) => r.status === 'rejected').length;

      totalSuccess += successCount;
      totalFailure += failureCount;

      console.log(`  ✓ ${successCount} successful, ✗ ${failureCount} failed`);
    }

    console.log('');
    console.log('🧹 Cleanup complete!');
    console.log(`✅ Successfully moved ${totalSuccess} messages to trash`);
    if (totalFailure > 0) {
      console.log(`❌ Failed to move ${totalFailure} messages`);
    }
  } catch (error) {
    console.error('\n❌ Cleanup failed:', error instanceof Error ? error.message : String(error));
    throw error;
  }
}

// Run if executed directly
if (import.meta.main) {
  cleanMail()
    .then(() => {
      console.log('');
      console.log('Done!');
      process.exit(0);
    })
    .catch(() => {
      process.exit(1);
    });
}
