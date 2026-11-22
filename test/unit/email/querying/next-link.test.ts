import '../../../lib/env-loader.ts';
import { ProtocolError, ProtocolErrorCode } from '@mcp-z/server';
import assert from 'assert';
import { checkNextLink } from '../../../../src/email/querying/next-link.ts';

// Link forms recorded from Microsoft Graph (outlook-nextlink-facts): page size 3, the $select list is the one this server sends.
const SELECT = '$select=id,conversationId,receivedDateTime,subject,bodyPreview,from,toRecipients,ccRecipients,bccRecipients';
// The service sent a 104-character $skiptoken of letters and digits. This is a 104-character PLACEHOLDER of the same shape, not a real token.
const PLACEHOLDER_SKIPTOKEN = 'A1b2C3d4E5f6G7h8I9j0K1l2M3n4O5p6Q7r8S9t0U1v2W3x4Y5z6A1b2C3d4E5f6G7h8I9j0K1l2M3n4O5p6Q7r8S9t0U1v2W3x4Y5z6';
assert.strictEqual(PLACEHOLDER_SKIPTOKEN.length, 104);

const SKIP_PATH = `/me/messages?${SELECT}&$top=3&$skip=3`;
const SKIP_LINK = `https://graph.microsoft.com/v1.0${SKIP_PATH}`;
// The $search value is quoted and percent-encoded as written here (placeholder shape; the facts record the parameter order, not its bytes).
const SKIPTOKEN_PATH = `/me/messages?$search=%22subject%3Atest%22&${SELECT}&$top=3&$skiptoken=${PLACEHOLDER_SKIPTOKEN}`;
const SKIPTOKEN_LINK = `https://graph.microsoft.com/v1.0${SKIPTOKEN_PATH}`;

function refused(link: string, reason: RegExp): void {
  assert.throws(
    () => checkNextLink(link),
    (error: unknown) => {
      assert.ok(error instanceof ProtocolError);
      assert.strictEqual(error.code, ProtocolErrorCode.InternalError);
      assert.match(error.message, /next-page link that cannot be followed/);
      assert.match(error.message, reason);
      assert.match(error.message, /results would be incomplete/);
      return true;
    }
  );
}

describe('checkNextLink', () => {
  it('accepts the $skip form and returns its path and query unchanged', () => {
    assert.strictEqual(checkNextLink(SKIP_LINK), SKIP_PATH);
  });

  it('accepts the $skiptoken form and returns its path and query byte for byte', () => {
    assert.strictEqual(checkNextLink(SKIPTOKEN_LINK), SKIPTOKEN_PATH);
  });

  it('refuses a link that is not a URL', () => {
    refused('not a url', /not a URL/);
  });

  it('refuses a link on another host', () => {
    refused('https://other.example/v1.0/me/messages?$skip=3', /host other\.example/);
  });

  it('refuses a link on another path', () => {
    refused('https://graph.microsoft.com/v1.0/users?$skip=3', /path \/v1\.0\/users/);
  });

  it('accepts the default https port written out and keeps the leading slash', () => {
    assert.strictEqual(checkNextLink('https://graph.microsoft.com:443/v1.0/me/messages?$skip=3'), '/me/messages?$skip=3');
  });

  it('refuses a non-https scheme and names the scheme, not the host', () => {
    refused('http://graph.microsoft.com/v1.0/me/messages?$skip=3', /scheme http/);
    refused('HTTP://graph.microsoft.com/v1.0/me/messages?$skip=3', /scheme http/);
  });

  it('refuses a link on another version', () => {
    refused('https://graph.microsoft.com/beta/me/messages?$skip=3', /path \/beta\/me\/messages/);
  });
});
