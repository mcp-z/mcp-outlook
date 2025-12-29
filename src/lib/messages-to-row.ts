import { extractCurrentMessageFromHtmlToText, extractCurrentMessageFromText, formatAddresses, normalizeDateToISO } from '@mcp-z/email';

interface OutlookRecipient {
  emailAddress?: {
    name?: string;
    address?: string;
  };
  name?: string;
  address?: string;
}

interface OutlookMessageBody {
  contentType?: 'text' | 'html';
  content?: string;
}

interface OutlookMessage {
  id?: string;
  conversationId?: string;
  toRecipients?: OutlookRecipient[];
  from?: OutlookRecipient;
  ccRecipients?: OutlookRecipient[];
  bccRecipients?: OutlookRecipient[];
  receivedDateTime?: string;
  subject?: string;
  bodyPreview?: string;
  body?: OutlookMessageBody;
}

interface FormatOptions {
  body?: boolean;
  addressFormat?: 'raw' | 'name' | 'email';
}

// Type guard for OutlookMessage
function isOutlookMessage(obj: unknown): obj is OutlookMessage {
  if (typeof obj !== 'object' || obj === null) {
    return false;
  }

  const msg = obj as Record<string, unknown>;

  if (msg.id !== undefined && typeof msg.id !== 'string') return false;
  if (msg.conversationId !== undefined && typeof msg.conversationId !== 'string') return false;
  if (msg.receivedDateTime !== undefined && typeof msg.receivedDateTime !== 'string') return false;
  if (msg.subject !== undefined && typeof msg.subject !== 'string') return false;
  if (msg.bodyPreview !== undefined && typeof msg.bodyPreview !== 'string') return false;

  // Check recipient arrays if they exist
  if (msg.toRecipients !== undefined && !Array.isArray(msg.toRecipients)) return false;
  if (msg.ccRecipients !== undefined && !Array.isArray(msg.ccRecipients)) return false;
  if (msg.bccRecipients !== undefined && !Array.isArray(msg.bccRecipients)) return false;

  // Check from field if it exists
  if (msg.from !== undefined && (typeof msg.from !== 'object' || msg.from === null)) return false;

  // Check body field if it exists
  if (msg.body !== undefined) {
    if (typeof msg.body !== 'object' || msg.body === null) return false;
    const body = msg.body as Record<string, unknown>;
    if (body.contentType !== undefined && !['text', 'html'].includes(String(body.contentType))) return false;
    if (body.content !== undefined && typeof body.content !== 'string') return false;
  }

  return true;
}

function formatRecipientList(list: OutlookRecipient[] = [], mode: 'raw' | 'name' | 'email' = 'email'): string {
  if (!Array.isArray(list) || list.length === 0) return '';
  const addresses = list
    .map((r) => ({
      name: r?.emailAddress?.name ?? r?.name ?? '',
      email: r?.emailAddress?.address ?? r?.address ?? '',
    }))
    .filter((x) => x.email);
  return formatAddresses(addresses, mode);
}

// Result type for the row transformation
type MessageRow = [
  string, // id
  string, // provider
  string, // threadId
  string, // to
  string, // from
  string, // cc
  string, // bcc
  string, // date
  string, // subject
  string, // labels
  string, // snippet
  string, // bodyFull
];

export function toRowFromOutlook(msg: unknown, opts: FormatOptions = { body: false, addressFormat: 'email' }): MessageRow {
  // Validate input using type guard
  if (!isOutlookMessage(msg)) {
    // Return empty row for invalid input
    return ['', 'outlook', '', '', '', '', '', '', '', '', '', ''];
  }

  const id = msg.id ?? '';
  const provider = 'outlook';
  const threadId = msg.conversationId ?? '';
  const fmt = opts.addressFormat || 'email';
  const to = formatRecipientList(msg.toRecipients || [], fmt);

  const from = (() => {
    const f = msg.from;
    if (!f) return '';
    const name = f.emailAddress?.name ?? f.name ?? '';
    const addr = f.emailAddress?.address ?? f.address ?? '';
    return formatRecipientList([{ name, address: addr }], fmt);
  })();

  const cc = formatRecipientList(msg.ccRecipients || [], fmt);
  const bcc = formatRecipientList(msg.bccRecipients || [], fmt);
  const date = normalizeDateToISO(msg.receivedDateTime) ?? '';
  const subject = msg.subject ?? '';
  const labels = '';
  const snippet = msg.bodyPreview ?? '';

  let bodyFull = '';
  if (opts.body && msg.body) {
    const contentType = msg.body.contentType || 'text';
    const content = msg.body.content ?? '';
    bodyFull = contentType === 'text' ? extractCurrentMessageFromText(content) : extractCurrentMessageFromHtmlToText(content);
  }

  return [id, provider, threadId, to, from, cc, bcc, date, subject, labels, snippet, bodyFull];
}

interface ClientSideFilters {
  subjectIncludes?: string[];
  bodyIncludes?: string[];
  textIncludes?: string[];
  fromIncludes?: string[];
  toIncludes?: string[];
  ccIncludes?: string[];
  bccIncludes?: string[];
}

interface FilterContent {
  subject?: string;
  snippetOrPreview?: string;
  fullBody?: string;
  from?: string;
  to?: string;
  cc?: string;
  bcc?: string;
}

export function filterClientSide(filters: ClientSideFilters, { subject = '', snippetOrPreview = '', fullBody = '', from = '', to = '', cc = '', bcc = '' }: FilterContent = {}) {
  const lower = (a: string[]) => a.map((t) => String(t).toLowerCase());
  const subjectTokens = lower(filters.subjectIncludes || []);
  const bodyTokens = lower(filters.bodyIncludes || []);
  const textTokens = lower(filters.textIncludes || []);
  const fromTokens = lower(filters.fromIncludes || []);
  const toCredentials = lower(filters.toIncludes || []);
  const ccTokens = lower(filters.ccIncludes || []);
  const bccTokens = lower(filters.bccIncludes || []);

  const s = String(subject ?? '').toLowerCase();
  const b = String((fullBody || snippetOrPreview) ?? '').toLowerCase();
  const f = String(from ?? '').toLowerCase();
  const t = String(to ?? '').toLowerCase();
  const c = String(cc ?? '').toLowerCase();
  const bc = String(bcc ?? '').toLowerCase();

  const anyIncludes = (val: string, tokens: string[]) => (tokens.length === 1 ? val.includes(tokens[0] ?? '') : tokens.some((token) => val.includes(token)));
  const subjectOk = subjectTokens.length ? anyIncludes(s, subjectTokens) : true;
  const bodyOk = bodyTokens.length ? anyIncludes(b, bodyTokens) : true;
  const textOk = textTokens.length ? textTokens.some((token) => s.includes(token) || b.includes(token)) : true;
  const fromOk = fromTokens.length ? anyIncludes(f, fromTokens) : true;
  const toOk = toCredentials.length ? anyIncludes(t, toCredentials) : true;
  const ccOk = ccTokens.length ? anyIncludes(c, ccTokens) : true;
  const bccOk = bccTokens.length ? anyIncludes(bc, bccTokens) : true;

  return subjectOk && bodyOk && textOk && fromOk && toOk && ccOk && bccOk;
}
