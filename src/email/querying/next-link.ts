import { ProtocolError, ProtocolErrorCode } from '@mcp-z/server';

const GRAPH_HOST = 'graph.microsoft.com';
const GRAPH_VERSION_PATH = '/v1.0';
const MESSAGES_PATH = `${GRAPH_VERSION_PATH}/me/messages`;

function refuse(reason: string): never {
  throw new ProtocolError(ProtocolErrorCode.InternalError, `Microsoft Graph returned a next-page link that cannot be followed (${reason}); results would be incomplete`);
}

/**
 * Check an `@odata.nextLink` and return its path and query after the version (`/me/messages?...`),
 * as the URL parser writes them. Microsoft says to follow the whole link and not to rebuild it, so nothing is decoded or reordered.
 */
export function checkNextLink(link: string): string {
  let url: URL;
  try {
    url = new URL(link);
  } catch {
    return refuse('not a URL');
  }
  if (url.protocol !== 'https:') return refuse(`scheme ${url.protocol.slice(0, -1)}`);
  if (url.host !== GRAPH_HOST) return refuse(`host ${url.host}`);
  if (url.pathname !== MESSAGES_PATH) return refuse(`path ${url.pathname}`);
  return url.pathname.slice(GRAPH_VERSION_PATH.length) + url.search;
}
