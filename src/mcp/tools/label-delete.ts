import type { EnrichedExtra } from '@mcp-z/oauth-microsoft';
import { schemas } from '@mcp-z/oauth-microsoft';

const { AuthRequiredBranchSchema } = schemas;

import type { ToolModule } from '@mcp-z/server';
import { Client } from '@microsoft/microsoft-graph-client';
import { type CallToolResult, ErrorCode, McpError } from '@modelcontextprotocol/sdk/types.js';
import { z } from 'zod';
import { CHUNK_SIZE, MAX_BATCH_SIZE } from '../../constants.ts';

const BatchResultSchema = z.object({
  totalRequested: z.number().describe('Total number of categories requested to delete'),
  successCount: z.number().describe('Number of categories successfully deleted'),
  failureCount: z.number().describe('Number of categories that failed to delete'),
  results: z
    .array(
      z.object({
        id: z.string().describe('Category ID'),
        success: z.boolean().describe('Whether the operation succeeded'),
        error: z.string().optional().describe('Error message if operation failed'),
      })
    )
    .describe('Individual results for each category'),
});

const inputSchema = z.object({
  ids: z.array(z.coerce.string().trim().min(1)).min(1).max(MAX_BATCH_SIZE).describe('Outlook category IDs to delete'),
});

// Success branch schema
const successBranchSchema = BatchResultSchema.extend({
  type: z.literal('success'),
});

// Output schema with auth_required support
const outputSchema = z.discriminatedUnion('type', [successBranchSchema, AuthRequiredBranchSchema]);

const config = {
  description: 'Delete Outlook categories permanently (irreversible). Categories organize messages by color and name.',
  inputSchema: inputSchema,
  outputSchema: z.object({
    result: outputSchema,
  }),
} as const;

export type Input = z.infer<typeof inputSchema>;
export type Output = z.infer<typeof outputSchema>;

async function handler({ ids }: Input, extra: EnrichedExtra): Promise<CallToolResult> {
  const logger = extra.logger;
  logger.info('outlook-label-delete called', { count: ids.length });

  // Validate IDs for Outlook-specific constraints
  // Note: Schema already validates non-empty, trimmed strings with min/max length
  // This check enforces business rule: Outlook category IDs cannot contain path separators or angle brackets
  const invalidIds: string[] = [];
  for (const id of ids) {
    if (id.includes('/') || id.includes('\\') || id.includes('<') || id.includes('>')) {
      invalidIds.push(id);
    }
  }

  if (invalidIds.length > 0) {
    logger.info('outlook-label-delete found invalid ids', { invalidIds, count: invalidIds.length });
    throw new McpError(ErrorCode.InvalidParams, `Found ${invalidIds.length} invalid IDs (contain invalid characters): ${invalidIds.join(', ')}`);
  }

  try {
    const graph = Client.initWithMiddleware({ authProvider: extra.authContext.auth });
    // Process deletions in chunks to prevent memory exhaustion
    const allResults: Array<{ id: string; success: boolean; error?: string }> = [];

    for (let i = 0; i < ids.length; i += CHUNK_SIZE) {
      const chunk = ids.slice(i, i + CHUNK_SIZE);
      logger.info('Processing chunk', { chunkIndex: i / CHUNK_SIZE + 1, chunkSize: chunk.length, totalChunks: Math.ceil(ids.length / CHUNK_SIZE) });

      // Delete categories sequentially to avoid API lock conflicts
      const chunkResults: Array<{ status: string; value?: unknown; reason?: unknown }> = [];
      for (const id of chunk) {
        try {
          await graph.api(`/me/outlook/masterCategories/${encodeURIComponent(id)}`).delete();
          chunkResults.push({ status: 'fulfilled', value: { id, success: true } });
        } catch (error) {
          chunkResults.push({ status: 'rejected', reason: error });
        }
      }

      // Process chunk results
      const processedChunkResults = chunkResults.map((result, chunkIndex) => {
        const id = chunk[chunkIndex];
        if (!id) {
          throw new Error(`Chunk index ${chunkIndex} is out of bounds for chunk of size ${chunk.length}`);
        }
        if (result.status === 'fulfilled') {
          return { id, success: true };
        }

        // Extract error message more robustly
        let errorMessage = 'Unknown error';
        if (result.reason) {
          if (typeof result.reason === 'string') {
            errorMessage = result.reason;
          } else if (result.reason instanceof Error) {
            errorMessage = result.reason.message;
          } else if (result.reason && typeof result.reason === 'object' && 'message' in result.reason) {
            errorMessage = String(result.reason.message);
          } else if (result.reason && typeof result.reason === 'object' && 'error' in result.reason) {
            errorMessage = String(result.reason.error);
          } else {
            errorMessage = String(result.reason);
          }
        }

        return { id, success: false, error: errorMessage };
      });

      allResults.push(...processedChunkResults);
    }

    const successCount = allResults.filter((r) => r.success).length;
    const failureCount = allResults.filter((r) => !r.success).length;

    const result: Output = {
      type: 'success' as const,
      totalRequested: ids.length,
      successCount,
      failureCount,
      results: allResults,
    };

    logger.info('outlook-label-delete completed', {
      totalRequested: ids.length,
      successCount,
      failureCount,
      chunksProcessed: Math.ceil(ids.length / CHUNK_SIZE),
    });

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
    logger.error('outlook-label-delete error', { error: message });

    throw new McpError(ErrorCode.InternalError, `Error deleting categories: ${message}`, {
      stack: error instanceof Error ? error.stack : undefined,
    });
  }
}

export default function createTool() {
  return {
    name: 'label-delete',
    config,
    handler,
  } satisfies ToolModule;
}
