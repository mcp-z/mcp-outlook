import type { EnrichedExtra } from '@mcp-z/oauth-microsoft';
import { schemas } from '@mcp-z/oauth-microsoft';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import { Client } from '@microsoft/microsoft-graph-client';
import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { type CallToolResult, ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { z } from 'zod';
import { OutlookCategorySchema } from '../../schemas/index.ts';

/**
 * Input schema for the categories-list tool
 *
 * Defines the validation schema for input parameters.
 */
export const inputSchema = z.object({});

// Success branch schema
const successBranchSchema = z.object({
  type: z.literal('success'),
  items: z.array(OutlookCategorySchema),
});

// Output schema with auth_required support
const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'List Outlook categories with colors and names.',
  inputSchema: inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

/**
 * Input type for the categories-list tool
 *
 * Represents the input parameters for listing Outlook categories.
 */
export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler(_: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('outlook.categories.list called');

  try {
    const graph = Client.initWithMiddleware({ authProvider: extra.authContext.auth });
    const started = Date.now();

    // Call Microsoft Graph API to get master categories
    const response = await graph.api('/me/outlook/masterCategories').get();
    const categories = Array.isArray(response.value) ? response.value : [];

    const durationMs = Date.now() - started;
    logger.info('outlook.categories.list returning', { categoriesCount: categories.length });
    logger.info('outlook.categories.list metrics', { durationMs });

    const result: Output = {
      type: 'success' as const,
      items: categories.map((category: MicrosoftGraph.OutlookCategory) => ({
        id: category.id || undefined,
        displayName: category.displayName || undefined,
        color: category.color || undefined,
      })),
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
    logger.error('outlook.categories.list error', { error: message });

    throw new McpError(ErrorCode.InternalError, `Error listing categories: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'categories-list',
    config,
    handler,
  } satisfies ToolModule;
}
