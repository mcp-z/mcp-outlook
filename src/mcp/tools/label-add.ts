import type { EnrichedExtra } from '@mcp-z/oauth-microsoft';
import { schemas } from '@mcp-z/oauth-microsoft';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import { Client } from '@microsoft/microsoft-graph-client';
import { type CallToolResult, ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { z } from 'zod';

const LabelResultSchema = z.object({
  id: z.string().describe('Message ID the categories were added to'),
  labels: z.array(z.string()).optional().describe('Categories that were applied'),
});

const inputSchema = z.object({
  id: z.coerce.string().trim().min(1).describe('Outlook message ID to add categories to'),
  labels: z.array(z.coerce.string().trim()).min(1).describe('Category names to add (use outlook-categories-list to discover available categories)'),
});

// Success branch schema
const successBranchSchema = LabelResultSchema.extend({
  type: z.literal('success'),
});

// Output schema with auth_required support
const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Add a label/category or move message to a folder in Outlook',
  inputSchema: inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ id, labels }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('outlook.label.add called', { id, labels });

  try {
    const graph = Client.initWithMiddleware({ authProvider: extra.authContext.auth });
    await graph
      .api(`/me/messages/${encodeURIComponent(id)}`)
      .header('Content-Type', 'application/json')
      .patch({ categories: labels });

    const result: Output = {
      type: 'success' as const,
      id,
      labels,
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
    logger.error('outlook.label.add error', { error: message });

    throw new McpError(ErrorCode.InternalError, `Error adding label: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'label-add',
    config,
    handler,
  } satisfies ToolModule;
}
