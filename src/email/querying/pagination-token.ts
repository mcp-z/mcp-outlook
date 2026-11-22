import { ProtocolError, ProtocolErrorCode } from '@mcp-z/server';

const PREFIX = 'v3:';

interface V3Payload {
  v: 3;
  link?: string;
  offset: number;
  mode?: string;
}

export interface DecodeNextPageTokenResult {
  /** Path and query of the Graph next-page link that fetches the page, absent for the first page. */
  link?: string;
  offset: number;
  mode?: string;
}

export function encodeNextPageToken(params: { link?: string; offset: number; mode?: string }): string {
  const payload: V3Payload = { v: 3, offset: Math.max(0, Math.trunc(params.offset)) };
  if (params.link) payload.link = params.link;
  if (typeof params.mode === 'string' && params.mode.length > 0) payload.mode = params.mode;
  return `${PREFIX}${Buffer.from(JSON.stringify(payload), 'utf8').toString('base64url')}`;
}

export function decodeNextPageToken(token: string | undefined): DecodeNextPageTokenResult {
  if (!token) return { offset: 0 };

  if (!token.startsWith(PREFIX)) {
    throw new ProtocolError(ProtocolErrorCode.InvalidParams, 'pageToken was issued by an earlier mcp-outlook version and cannot be used; repeat the search without pageToken');
  }

  try {
    const parsed = JSON.parse(Buffer.from(token.slice(PREFIX.length), 'base64url').toString('utf8')) as Record<string, unknown>;
    if (parsed.v !== 3) throw new Error('unsupported token version');
    const { link, offset, mode } = parsed;
    if (link !== undefined && (typeof link !== 'string' || link === '')) throw new Error('invalid link');
    if (typeof offset !== 'number' || !Number.isInteger(offset) || offset < 0) throw new Error('invalid offset');
    return {
      ...(typeof link === 'string' && { link }),
      offset,
      ...(typeof mode === 'string' && { mode }),
    };
  } catch {
    throw new ProtocolError(ProtocolErrorCode.InvalidParams, 'Invalid pageToken');
  }
}
