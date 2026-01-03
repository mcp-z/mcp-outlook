import type { Client as GraphClient } from '@microsoft/microsoft-graph-client';
import type { Logger } from '../../src/types.ts';

// Type for error objects that may have status/code properties
type ErrorWithStatus = {
  status?: number;
  statusCode?: number;
  code?: number | string;
};

interface EmailAddress {
  name?: string | undefined;
  address: string;
}

interface Recipient {
  emailAddress: EmailAddress;
}

interface MessageBody {
  contentType: 'Text' | 'HTML';
  content: string;
}

interface GraphMessage {
  subject: string;
  body: MessageBody;
  toRecipients: Recipient[];
  from?: {
    emailAddress: EmailAddress;
  };
  ccRecipients?: Recipient[];
  bccRecipients?: Recipient[];
  categories?: string[];
  importance?: 'low' | 'normal' | 'high';
}

export interface CreateTestMessageOptions {
  subject?: string;
  body?: string;
  from?: { name?: string; address: string };
  to: string | string[];
  cc?: string | string[];
  bcc?: string | string[];
  categories?: string[];
  importance?: 'low' | 'normal' | 'high';
}

/**
 * Create a minimal test message in the test account and return its id.
 * Keeps setup DRY for tests inside microsoft/servers/mcp-outlook.
 */
// Helpers now accept a pre-created Microsoft Graph client instance (do not take credentials/providers)
export async function createTestMessage(client: GraphClient, opts: CreateTestMessageOptions): Promise<string> {
  const subject = opts.subject || `ci-test-${Date.now()}`;
  const body = opts.body || 'Automated integration test message body';

  // Build recipient arrays
  const toRecipients = (Array.isArray(opts.to) ? opts.to : [opts.to]).map((address) => ({
    emailAddress: { address },
  }));

  const message: GraphMessage = {
    subject,
    body: {
      contentType: 'Text',
      content: body,
    },
    toRecipients,
  };

  // Add optional fields
  if (opts.from) {
    message.from = {
      emailAddress: {
        name: opts.from.name,
        address: opts.from.address,
      },
    };
  }

  if (opts.cc) {
    message.ccRecipients = (Array.isArray(opts.cc) ? opts.cc : [opts.cc]).map((address) => ({
      emailAddress: { address },
    }));
  }

  if (opts.bcc) {
    message.bccRecipients = (Array.isArray(opts.bcc) ? opts.bcc : [opts.bcc]).map((address) => ({
      emailAddress: { address },
    }));
  }

  if (opts.categories) {
    message.categories = opts.categories;
  }

  if (opts.importance) {
    message.importance = opts.importance;
  }

  // Use draft-send pattern to get message ID
  // 1. Create draft
  const draft = await client.api('/me/messages').post(message);
  const draftId = draft.id;

  if (!draftId) {
    throw new Error('createTestMessage: draft creation did not return an ID');
  }

  // 2. Send the draft
  await client.api(`/me/messages/${draftId}/send`).post({});

  // 3. After sending, query Sent Items to get the real sent message ID
  // (the draft ID becomes invalid after sending)
  const timeout = 10000; // 10 seconds to find in Sent Items
  const interval = 500;
  const start = Date.now();

  while (Date.now() - start < timeout) {
    const sentMessages = await client.api('/me/mailFolders/SentItems/messages').filter(`subject eq '${subject}'`).top(1).get();

    if (sentMessages?.value?.[0]?.id) {
      return sentMessages.value[0].id;
    }

    await new Promise((resolve) => setTimeout(resolve, interval));
  }

  throw new Error(`createTestMessage: sent message not found in Sent Items with subject "${subject}" after ${timeout}ms`);
}

/**
 * Create a minimal draft message (without sending) in the test account and return its id.
 * Use this instead of createTestMessage when you don't need to actually send the email,
 * which helps avoid hitting daily sending limits.
 */
export async function createTestDraftMessage(client: GraphClient, opts: CreateTestMessageOptions): Promise<string> {
  const subject = opts.subject || `ci-test-draft-${Date.now()}`;
  const body = opts.body || 'Automated integration test draft message body';

  // Build recipient arrays
  const toRecipients = (Array.isArray(opts.to) ? opts.to : [opts.to]).map((address) => ({
    emailAddress: { address },
  }));

  const message: GraphMessage = {
    subject,
    body: {
      contentType: 'Text',
      content: body,
    },
    toRecipients,
  };

  // Add optional fields (same as createTestMessage)
  if (opts.from) {
    message.from = {
      emailAddress: {
        name: opts.from.name,
        address: opts.from.address,
      },
    };
  }

  if (opts.cc) {
    message.ccRecipients = (Array.isArray(opts.cc) ? opts.cc : [opts.cc]).map((address) => ({
      emailAddress: { address },
    }));
  }

  if (opts.bcc) {
    message.bccRecipients = (Array.isArray(opts.bcc) ? opts.bcc : [opts.bcc]).map((address) => ({
      emailAddress: { address },
    }));
  }

  if (opts.categories) {
    message.categories = opts.categories;
  }

  if (opts.importance) {
    message.importance = opts.importance;
  }

  // Create draft only (no sending)
  const draft = await client.api('/me/messages').post(message);
  const draftId = draft.id;

  if (!draftId) {
    throw new Error('createTestDraftMessage: draft creation did not return an ID');
  }

  return draftId;
}

/**
 * Delete a test message created with createTestMessage or createTestDraftMessage.
 * Silently succeeds if the message is already deleted (e.g., moved to trash by the test).
 * Throws on other errors that indicate actual problems.
 */
export async function deleteTestMessage(client: GraphClient, id: string, logger: Logger): Promise<void> {
  try {
    await client.api(`/me/messages/${encodeURIComponent(id)}`).delete();
    logger.debug('Test message close successful', { messageId: id });
  } catch (e: unknown) {
    const statusCode = e && typeof e === 'object' && ('status' in e || 'statusCode' in e) ? (e as ErrorWithStatus).status || (e as ErrorWithStatus).statusCode : undefined;
    const errorCode = e && typeof e === 'object' && 'code' in e ? (e as ErrorWithStatus).code : undefined;

    // If message already deleted (404 or ErrorItemNotFound), that's fine - close succeeded
    if (statusCode === 404 || errorCode === 'ErrorItemNotFound') {
      logger.debug('Test message close: already deleted', { messageId: id });
      return;
    }

    // For other errors, log and throw - these indicate real problems
    logger.error('Test message close failed', {
      messageId: id,
      error: e instanceof Error ? e.message : String(e),
      status: statusCode,
      code: errorCode,
    });
    throw e;
  }
}
