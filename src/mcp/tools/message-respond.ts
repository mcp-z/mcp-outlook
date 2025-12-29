import { ComposeContentTypeSchema } from '@mcp-z/email';
import type { EnrichedExtra } from '@mcp-z/oauth-microsoft';
import { schemas } from '@mcp-z/oauth-microsoft';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import { Client } from '@microsoft/microsoft-graph-client';
import { type CallToolResult, ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { z } from 'zod';

const SuccessSchema = z.object({
  success: z.boolean().describe('Whether the reply was sent successfully'),
  id: z.string().optional().describe('Original message ID that was replied to'),
});

const inputSchema = z.object({
  id: z.coerce.string().trim().min(1).describe('Outlook message ID to reply to'),
  body: z.coerce.string().trim().min(1).describe('Reply body content (plain text or HTML)'),
  contentType: ComposeContentTypeSchema,
});

// Success branch schema
const successBranchSchema = SuccessSchema.extend({
  type: z.literal('success'),
});

// Output schema with auth_required support
const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Reply to an Outlook message',
  inputSchema: inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ id, body, contentType }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('outlook.respond called', { id, hasBody: !!body });

  if (!id || !body) {
    logger.info('outlook.respond missing id or body');
    throw new McpError(ErrorCode.InvalidParams, 'Missing id or body');
  }

  try {
    const graph = Client.initWithMiddleware({ authProvider: extra.authContext.auth });
    await graph.api(`/me/messages/${encodeURIComponent(id)}/reply`).post({ message: { body: { contentType: contentType === 'html' ? 'HTML' : 'Text', content: body } } });

    logger.info('outlook.respond success', { id });

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
    logger.error('outlook.respond error', { error: message });

    throw new McpError(ErrorCode.InternalError, `Error replying to message: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'message-respond',
    config,
    handler,
  } satisfies ToolModule;
}
