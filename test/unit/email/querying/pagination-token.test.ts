import assert from 'assert';
import { decodeNextPageToken, encodeNextPageToken } from '../../../../src/email/querying/pagination-token.ts';

describe('pagination-token', () => {
  it('round-trips skipToken, offset, and mode', () => {
    const token = encodeNextPageToken({ skipToken: 'abc123', offset: 5, mode: 'search' });
    const decoded = decodeNextPageToken(token);
    assert.strictEqual(decoded.legacy, false);
    assert.strictEqual(decoded.skipToken, 'abc123');
    assert.strictEqual(decoded.offset, 5);
    assert.strictEqual(decoded.mode, 'search');
  });

  it('treats undefined skipToken as first-page sentinel', () => {
    const token = encodeNextPageToken({ offset: 0 });
    assert.ok(token.startsWith('v2:'), 'should use v2 prefix');
    const decoded = decodeNextPageToken(token);
    assert.strictEqual(decoded.skipToken, undefined);
    assert.strictEqual(decoded.offset, 0);
    assert.strictEqual(decoded.legacy, false);
  });

  it('returns legacy metadata for invalid tokens', () => {
    const decoded = decodeNextPageToken('not-a-token');
    assert.strictEqual(decoded.legacy, true);
    assert.strictEqual(decoded.skipToken, 'not-a-token');
    assert.strictEqual(decoded.offset, 0);
  });

  it('preserves mode when provided', () => {
    const token = encodeNextPageToken({ skipToken: 'page-2', offset: 10, mode: 'search' });
    const decoded = decodeNextPageToken(token);
    assert.strictEqual(decoded.mode, 'search');
  });
});
