import type { EnrichedExtra } from '@mcp-z/oauth-microsoft';
import { schemas } from '@mcp-z/oauth-microsoft';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import { Client } from '@microsoft/microsoft-graph-client';
import { type CallToolResult, ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { z } from 'zod';

const SuccessSchema = z.object({
  success: z.boolean().describe('Whether the operation completed successfully'),
  id: z.string().optional().describe('Message ID that was marked as read'),
});

const inputSchema = z.object({
  id: z.coerce.string().trim().min(1).describe('Outlook message ID to mark as read'),
});

// Success branch schema
const successBranchSchema = SuccessSchema.extend({
  type: z.literal('success'),
});

// Output schema with auth_required support
const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Mark an Outlook message as read',
  inputSchema: inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ id }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('outlook-message-mark-read called', { id });

  if (!id) {
    logger.info('outlook-message-mark-read missing id');
    throw new McpError(ErrorCode.InvalidParams, 'Missing id');
  }

  try {
    const graph = Client.initWithMiddleware({ authProvider: extra.authContext.auth });
    await graph.api(`/me/messages/${encodeURIComponent(id)}`).update({ isRead: true });

    logger.info('outlook-message-mark-read success', { id });

    const result: Output = {
      type: 'success' as const,
      success: true,
      id,
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
    logger.error('outlook-message-mark-read error', { error: message });

    throw new McpError(ErrorCode.InternalError, `Error marking message as read: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'message-mark-read',
    config,
    handler,
  } satisfies ToolModule;
}
