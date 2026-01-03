import type { EnrichedExtra, MicrosoftAuthProvider } from '@mcp-z/oauth-microsoft';
import type { AnySchema, SchemaOutput } from '@modelcontextprotocol/sdk/server/zod-compat.js';
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import pino from 'pino';
import type { StorageContext, StorageExtra } from '../../src/types.ts';

/**
 * Typed handler signature for test files
 * Use with tool's Input type: `let handler: TypedHandler<Input>;`
 */
export type TypedHandler<I, E = EnrichedExtra> = (input: I, extra: E) => Promise<CallToolResult>;

/**
 * Create EnrichedExtra for testing
 *
 * In production, the middleware automatically creates and injects authContext and logger.
 * In tests, we call handlers directly, so we need to provide it ourselves.
 *
 * Note: The auth and logger here are just placeholders - the real auth/logger come from
 * the middleware wrapper created in create-middleware-context.ts
 */
export function createExtra(): EnrichedExtra;
export function createExtra(storageContext: StorageContext): EnrichedExtra & StorageExtra;
export function createExtra(storageContext?: StorageContext): EnrichedExtra {
  const extra: EnrichedExtra & Partial<StorageExtra> = {
    signal: new AbortController().signal,
    requestId: 'test-request-id',
    sendNotification: async () => {},
    sendRequest: async <U extends AnySchema>() => ({}) as SchemaOutput<U>,
    // Middleware injects these - placeholders for type compatibility
    authContext: {
      auth: {} as MicrosoftAuthProvider, // Placeholder auth client
      accountId: 'test-account',
    },
    logger: pino({ level: 'silent' }),
    ...(storageContext ? { storageContext } : {}),
  };

  return extra as EnrichedExtra;
}
