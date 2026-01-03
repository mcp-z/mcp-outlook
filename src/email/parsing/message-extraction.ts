import type { Client } from '@microsoft/microsoft-graph-client';
import type { OutlookMessage } from '../../lib/outlook/types.ts';
import { mapOutlookMessage, type NormalizedMessage } from './message-mapping.ts';

/** Fetch a single Outlook message. Throws on non-OK Graph responses. */
export async function fetchOutlookMessage(graph: Client, id: string): Promise<NormalizedMessage> {
  if (!id) throw new Error('fetchOutlookMessage: id required');
  const selectFields = ['id', 'conversationId', 'receivedDateTime', 'subject', 'from', 'toRecipients', 'ccRecipients', 'bccRecipients', 'bodyPreview', 'categories', 'body'];
  try {
    const data: unknown = await graph
      .api(`/me/messages/${encodeURIComponent(id)}`)
      .select(selectFields.join(','))
      .get();
    return mapOutlookMessage(data as OutlookMessage, { preferHtml: true });
  } catch (e: unknown) {
    const status = (e as Record<string, unknown>)?.status || (e as Record<string, unknown>)?.code || 'unknown';
    throw new Error(`outlook.fetchMessage: failed (${status})`);
  }
}
