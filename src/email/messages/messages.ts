import { buildContentForItems, stripHtml, normalizeDateToISO as toIsoUtc } from '@mcp-z/email';
import type { OutlookMessage } from '../../lib/outlook/types.js';
import { extractFrom, formatAddressList } from '../parsing/header-parsing.js';

export { buildContentForItems };

export function toRowFromOutlook(msg: OutlookMessage, opts: { body?: boolean; addressFormat?: 'email' | 'name' | 'raw' } = { body: false, addressFormat: 'email' }) {
  const id = msg.id ?? '';
  const provider = 'outlook';
  const threadId = msg.conversationId ?? '';
  const to = msg.toRecipients ? formatAddressList(msg.toRecipients, opts.addressFormat) : '';
  const from = msg.from ? (extractFrom(msg.from) ?? '') : '';
  const cc = msg.ccRecipients ? formatAddressList(msg.ccRecipients, opts.addressFormat) : '';
  const bcc = msg.bccRecipients ? formatAddressList(msg.bccRecipients, opts.addressFormat) : '';
  const date = msg.receivedDateTime ? toIsoUtc(msg.receivedDateTime) || msg.receivedDateTime : '';
  const subject = msg.subject ?? '';
  const labels = msg.categories ? msg.categories.join(';') : '';
  const snippet = msg.bodyPreview ?? '';
  let bodyFull = '';
  if (opts.body && msg.body) {
    const raw = msg.body.content ?? '';
    if (raw) {
      bodyFull = msg.body.contentType && msg.body.contentType.toLowerCase() === 'html' ? stripHtml(raw) : String(raw);
    }
  }
  return [id, provider, threadId, to, from, cc, bcc, date, subject, labels, snippet, bodyFull];
}
