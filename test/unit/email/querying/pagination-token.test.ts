import '../../../lib/env-loader.ts';
import { ProtocolError, ProtocolErrorCode } from '@mcp-z/server';
import assert from 'assert';
import { decodeNextPageToken, encodeNextPageToken } from '../../../../src/email/querying/pagination-token.ts';

// A v2 token as the previous version issued it: base64url of {"v":2,"skipToken":"abc","offset":0,"mode":"search"}, written as a literal.
const V2_TOKEN = 'v2:eyJ2IjoyLCJza2lwVG9rZW4iOiJhYmMiLCJvZmZzZXQiOjAsIm1vZGUiOiJzZWFyY2gifQ';
const EARLIER_VERSION = /issued by an earlier mcp-outlook version and cannot be used; repeat the search without pageToken/;

describe('pagination-token', () => {
  it('round-trips link, offset, and mode', () => {
    const link = '/me/messages?$select=id&$top=3&$skip=3';
    const token = encodeNextPageToken({ link, offset: 5, mode: 'search' });
    assert.ok(token.startsWith('v3:'), 'should use v3 prefix');
    const decoded = decodeNextPageToken(token);
    assert.strictEqual(decoded.link, link);
    assert.strictEqual(decoded.offset, 5);
    assert.strictEqual(decoded.mode, 'search');
  });

  it('omits the link for the first page', () => {
    const token = encodeNextPageToken({ offset: 0 });
    assert.ok(token.startsWith('v3:'), 'should use v3 prefix');
    const decoded = decodeNextPageToken(token);
    assert.strictEqual(decoded.link, undefined);
    assert.strictEqual(decoded.offset, 0);
  });

  it('decodes no token as the first page', () => {
    const decoded = decodeNextPageToken(undefined);
    assert.strictEqual(decoded.link, undefined);
    assert.strictEqual(decoded.offset, 0);
  });

  it('refuses a raw string token as issued by an earlier version', () => {
    assert.throws(
      () => decodeNextPageToken('not-a-token'),
      (error: unknown) => {
        assert.ok(error instanceof ProtocolError);
        assert.strictEqual(error.code, ProtocolErrorCode.InvalidParams);
        assert.match(error.message, EARLIER_VERSION);
        return true;
      }
    );
  });

  it('refuses a v2 token as issued by an earlier version', () => {
    assert.throws(
      () => decodeNextPageToken(V2_TOKEN),
      (error: unknown) => {
        assert.ok(error instanceof ProtocolError);
        assert.strictEqual(error.code, ProtocolErrorCode.InvalidParams);
        assert.match(error.message, EARLIER_VERSION);
        return true;
      }
    );
  });

  it('preserves mode when provided', () => {
    const token = encodeNextPageToken({ link: '/me/messages?$skip=10', offset: 10, mode: 'search' });
    const decoded = decodeNextPageToken(token);
    assert.strictEqual(decoded.mode, 'search');
  });

  it('throws InvalidParams for a v3 token with bad encoding', () => {
    assert.throws(() => decodeNextPageToken('v3:!!!'), /Invalid pageToken/);
  });

  it('throws InvalidParams for a v3 token with the wrong version', () => {
    const payload = Buffer.from(JSON.stringify({ v: 1, link: '/me/messages?$skip=3', offset: 0 })).toString('base64url');
    assert.throws(() => decodeNextPageToken(`v3:${payload}`), /Invalid pageToken/);
  });
});
