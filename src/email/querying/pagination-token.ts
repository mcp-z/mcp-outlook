const PREFIX = 'v2:';
export const FIRST_PAGE_SKIP_TOKEN = '__mcp_outlook_first_page__';

interface V2Payload {
  v: 2;
  skipToken: string;
  offset: number;
  mode?: string;
}

export interface DecodeNextPageTokenResult {
  skipToken?: string;
  offset: number;
  mode?: string;
  legacy: boolean;
}

export function encodeNextPageToken(params: { skipToken?: string; offset: number; mode?: string }): string {
  const sanitizedOffset = Math.max(0, Math.trunc(params.offset));
  const payload: V2Payload = {
    v: 2,
    skipToken: params.skipToken ?? FIRST_PAGE_SKIP_TOKEN,
    offset: sanitizedOffset,
  };
  if (typeof params.mode === 'string' && params.mode.length > 0) {
    payload.mode = params.mode;
  }
  const serialized = JSON.stringify(payload);
  return `${PREFIX}${base64UrlEncode(serialized)}`;
}

export function decodeNextPageToken(token: string | undefined): DecodeNextPageTokenResult {
  if (!token) {
    return { offset: 0, legacy: true };
  }

  if (!token.startsWith(PREFIX)) {
    return { skipToken: token, offset: 0, legacy: true };
  }

  const encoded = token.slice(PREFIX.length);
  try {
    const decodedJson = base64UrlDecode(encoded);
    const parsed = JSON.parse(decodedJson) as Record<string, unknown>;
    if (parsed.v !== 2) throw new Error('unsupported token version');
    const skipTokenRaw = parsed.skipToken;
    const offsetRaw = parsed.offset;
    if (typeof skipTokenRaw !== 'string' || skipTokenRaw.trim() === '') throw new Error('missing skipToken');
    if (typeof offsetRaw !== 'number' || !Number.isInteger(offsetRaw) || offsetRaw < 0) throw new Error('invalid offset');

    return {
      skipToken: skipTokenRaw === FIRST_PAGE_SKIP_TOKEN ? undefined : skipTokenRaw,
      offset: offsetRaw,
      mode: typeof parsed.mode === 'string' ? parsed.mode : undefined,
      legacy: false,
    };
  } catch {
    return { skipToken: token, offset: 0, legacy: true };
  }
}

function base64UrlEncode(value: string): string {
  return Buffer.from(value, 'utf8').toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

function base64UrlDecode(value: string): string {
  const padded = value.replace(/-/g, '+').replace(/_/g, '/');
  const padLength = (4 - (padded.length % 4)) % 4;
  const normalized = padded + '='.repeat(padLength);
  return Buffer.from(normalized, 'base64').toString('utf8');
}
