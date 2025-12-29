import { normalizeDateToISO as toIsoUtc } from '@mcp-z/email';
import type { EnrichedExtra } from '@mcp-z/oauth-microsoft';
import type { ResourceConfig, ResourceModule } from '@mcp-z/server';
import { Client } from '@microsoft/microsoft-graph-client';
import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { ResourceTemplate } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { RequestHandlerExtra } from '@modelcontextprotocol/sdk/shared/protocol.js';
import type { ReadResourceResult, ServerNotification, ServerRequest } from '@modelcontextprotocol/sdk/types.js';

export default function createResource() {
  const template = new ResourceTemplate('outlook://messages/{id}', {
    list: undefined,
  });
  const config: ResourceConfig = {
    description: 'Outlook message metadata (lightweight: id, subject, from, to, date)',
    mimeType: 'application/json',
  };

  const handler = async (uri: URL, variables: { id: string }, extra: RequestHandlerExtra<ServerRequest, ServerNotification>): Promise<ReadResourceResult> => {
    try {
      const { logger, authContext } = extra as unknown as EnrichedExtra;

      logger.info(variables, 'outlook-email resource fetch');

      const graph = Client.initWithMiddleware({ authProvider: authContext.auth });
      const message = (await graph.api(`/me/messages/${encodeURIComponent(variables.id)}`).get()) as MicrosoftGraph.Message;

      // Build message data
      const toAddresses = message.toRecipients?.map((r) => r.emailAddress?.address).filter(Boolean);
      const toStr = toAddresses && toAddresses.length > 0 ? toAddresses.join(', ') : '';

      // Return lightweight metadata only (no body/snippet)
      const metadata = {
        id: message.id ?? variables.id,
        subject: message.subject ?? '',
        from: message.from?.emailAddress?.address ?? '',
        to: toStr,
        date: toIsoUtc(message.receivedDateTime) || message.receivedDateTime || '',
      };

      return {
        contents: [
          {
            uri: uri.href,
            mimeType: 'application/json',
            text: JSON.stringify(metadata),
          },
        ],
      };
    } catch (e) {
      const { logger } = extra as unknown as EnrichedExtra;
      logger.error(e as Record<string, unknown>, 'outlook-email resource fetch failed');
      const error = e as { message?: unknown };
      return {
        contents: [
          {
            uri: uri.href,
            mimeType: 'application/json',
            text: JSON.stringify({ error: String(error?.message ?? e) }),
          },
        ],
      };
    }
  };

  return {
    name: 'email',
    template,
    config,
    handler,
  } satisfies ResourceModule;
}
