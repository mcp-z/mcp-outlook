import { ComposeContentTypeSchema, createEmailRecipientsSchema, createMessageResultSchema } from '@mcp-z/email';
import type { EnrichedExtra } from '@mcp-z/oauth-microsoft';
import { schemas } from '@mcp-z/oauth-microsoft';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import { Client } from '@microsoft/microsoft-graph-client';
import { type CallToolResult, ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { z } from 'zod';
import { buildOutlookMessage } from '../../email/composition/outlook-message-builder.js';

const MessageResultSchema = createMessageResultSchema('outlook');

const inputSchema = z.object({
  to: createEmailRecipientsSchema('to', true),
  cc: createEmailRecipientsSchema('cc', false),
  bcc: createEmailRecipientsSchema('bcc', false),
  subject: z.string().describe('Email subject line').default(''),
  body: z.string().trim().min(1).describe('Email body content (plain text or HTML)'),
  contentType: ComposeContentTypeSchema,
});

// Success branch schema
const successBranchSchema = MessageResultSchema.extend({
  type: z.literal('success'),
});

// Output schema with auth_required support
const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Send an email message through Outlook',
  inputSchema: inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ to, cc, bcc, subject, body, contentType }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;

  try {
    const graph = Client.initWithMiddleware({ authProvider: extra.authContext.auth });

    // Use composition module to build the message
    const toArray = to
      ?.split(',')
      .map((addr) => addr.trim())
      .filter((addr) => addr);
    const ccArray = cc
      ?.split(',')
      .map((addr) => addr.trim())
      .filter((addr) => addr);
    const bccArray = bcc
      ?.split(',')
      .map((addr) => addr.trim())
      .filter((addr) => addr);

    const message = buildOutlookMessage({
      to: toArray,
      cc: ccArray && ccArray.length > 0 ? ccArray : undefined,
      bcc: bccArray && bccArray.length > 0 ? bccArray : undefined,
      subject: subject ?? '',
      body,
      contentType: contentType === 'html' ? 'HTML' : 'Text',
    });

    // Use draft-send pattern to get message ID
    // 1. Create draft message
    const draft = await graph.api('/me/messages').post(message);
    const messageId = draft.id;

    if (!messageId) {
      logger.error('outlook.message.send draft creation failed - no ID returned');
      throw new McpError(ErrorCode.InternalError, 'Failed to create draft message');
    }

    // 2. Send the draft
    await graph.api(`/me/messages/${messageId}/send`).post({});
    logger.info('Outlook: sent mail successfully', { messageId });

    const toCount = message.toRecipients?.length ?? 0;
    const ccCount = message.ccRecipients?.length ?? 0;
    const bccCount = message.bccRecipients?.length ?? 0;

    logger.info('outlook.message.send success', { toCount, ccCount, bccCount, subject: subject ?? '', messageId });

    const totalRecipients = toCount + ccCount + bccCount;
    const result: Output = {
      type: 'success' as const,
      id: messageId,
      sentAt: new Date().toISOString(),
      recipientCount: totalRecipients,
      webLink: 'https://outlook.live.com/mail/0/sentitems',
    };

    return {
      content: [
        {
          type: 'text' as const,
          text: JSON.stringify(result),
        },
      ],
      structuredContent: { result },
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error('outlook.message.send error', { error: message });

    throw new McpError(ErrorCode.InternalError, `Error sending message: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'message-send',
    config,
    handler,
  } satisfies ToolModule;
}
