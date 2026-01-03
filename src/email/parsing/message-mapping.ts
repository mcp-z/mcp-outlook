import { stripHtml, normalizeDateToISO as toIsoUtc } from '@mcp-z/email';
import type { OutlookMessage } from '../../lib/outlook/types.ts';
import { extractEmailsFromRecipients } from './header-parsing.ts';

export interface NormalizedAddress {
  address?: string | undefined;
  name?: string | undefined;
}
export interface NormalizedMessage {
  id: string;
  threadId?: string;
  date?: string;
  subject?: string;
  from?: NormalizedAddress;
  to?: string[];
  cc?: string[];
  bcc?: string[];
  snippet?: string;
  labels?: string[];
  body?: string;
}

export function mapOutlookMessage(msg: OutlookMessage, opts: { preferHtml?: boolean } = {}): NormalizedMessage {
  const fromAddr = msg.from?.emailAddress;
  const bodyContent = msg.body?.content ?? '';
  const bodyType = String(msg.body?.contentType ?? '').toLowerCase();

  let body: string | undefined;
  if (bodyContent) {
    if (opts.preferHtml && bodyType === 'html') body = stripHtml(bodyContent);
    else body = bodyType === 'html' ? stripHtml(bodyContent) : String(bodyContent);
  }

  const out: NormalizedMessage = { id: msg.id ?? '' };

  if (msg.conversationId) out.threadId = String(msg.conversationId);
  if (msg.receivedDateTime) out.date = toIsoUtc(msg.receivedDateTime) ?? String(msg.receivedDateTime);
  if (msg.subject) out.subject = String(msg.subject);

  const fromObj = fromAddr ? { address: fromAddr.address || undefined, name: fromAddr.name || undefined } : undefined;
  if (fromObj && (fromObj.address || fromObj.name)) out.from = fromObj;

  const toList = extractEmailsFromRecipients(msg.toRecipients ?? []);
  if (Array.isArray(toList) && toList.length) out.to = toList;

  const ccList = extractEmailsFromRecipients(msg.ccRecipients ?? []);
  if (Array.isArray(ccList) && ccList.length) out.cc = ccList;

  const bccList = extractEmailsFromRecipients(msg.bccRecipients ?? []);
  if (Array.isArray(bccList) && bccList.length) out.bcc = bccList;

  if (msg.bodyPreview) out.snippet = String(msg.bodyPreview);
  if (Array.isArray(msg.categories) && msg.categories.length) out.labels = msg.categories.map(String);
  if (body) out.body = body;

  return out;
}
