import assert from 'assert';
import { toRowFromOutlook } from '../../../src/lib/messages-to-row.js';

// --- Date normalization + address formatting ---------------------------------
it('toRowFromOutlook: normalizes receivedDateTime to ISO8601 UTC', () => {
  const rawDate = 'Sat, 30 Aug 2025 15:44:48 -0700';
  const msg = {
    id: 'mid',
    conversationId: 'cid',
    subject: 'Hello',
    toRecipients: [{ emailAddress: { name: 'Bob', address: 'bob@example.com' } }],
    from: { emailAddress: { name: 'Alice', address: 'alice@example.com' } },
    ccRecipients: [],
    bccRecipients: [],
    receivedDateTime: rawDate,
    bodyPreview: 'snippet',
  };
  const row = toRowFromOutlook(msg, { body: false });
  const date = row[7];
  assert.equal(date, '2025-08-30T22:44:48.000Z');
});

it('toRowFromOutlook: address formatting modes (default=email)', () => {
  const msg = {
    id: 'mid',
    conversationId: 'cid',
    subject: 'Hello',
    toRecipients: [{ emailAddress: { name: 'Bob', address: 'bob@example.com' } }, { emailAddress: { name: '', address: 'carol@example.com' } }],
    from: { emailAddress: { name: 'Alice Smith', address: 'alice@example.com' } },
    ccRecipients: [],
    bccRecipients: [],
    receivedDateTime: 'Sat, 30 Aug 2025 15:44:48 -0700',
    bodyPreview: 'snippet',
  };

  // default (email)
  let row = toRowFromOutlook(msg, { body: false });
  assert.equal(row[3], 'bob@example.com, carol@example.com');
  assert.equal(row[4], 'alice@example.com');

  // raw
  row = toRowFromOutlook(msg, { body: false, addressFormat: 'raw' });
  assert.equal(row[3], 'Bob <bob@example.com>, carol@example.com');
  assert.equal(row[4], 'Alice Smith <alice@example.com>');

  // name (fallback to email when missing)
  row = toRowFromOutlook(msg, { body: false, addressFormat: 'name' });
  assert.equal(row[3], 'Bob, carol@example.com');
  assert.equal(row[4], 'Alice Smith');
});

it('toRowFromOutlook: trims quoted history in text and html', () => {
  const textMsg = {
    id: 'mid',
    conversationId: 'cid',
    subject: 'Hello',
    toRecipients: [],
    from: {},
    ccRecipients: [],
    bccRecipients: [],
    receivedDateTime: 'Sat, 30 Aug 2025 15:44:48 -0700',
    bodyPreview: 'snippet',
    body: { contentType: 'text', content: 'Hi Bob\nLatest update\n\n> On Aug 24, 2025, at 11:16\u202FAM, Kevin wrote:\n> Previous content\n-----Original Message-----\nFrom: Alice' },
  };
  let row = toRowFromOutlook(textMsg, { body: true });
  let body = row[11] ?? '';
  assert.ok(body.includes('Latest update'));
  assert.ok(!/Original Message/i.test(body));

  const htmlMsg = {
    id: 'mid',
    conversationId: 'cid',
    subject: 'Hello',
    toRecipients: [],
    from: {},
    ccRecipients: [],
    bccRecipients: [],
    receivedDateTime: 'Sat, 30 Aug 2025 15:44:48 -0700',
    bodyPreview: 'snippet',
    body: { contentType: 'html', content: '<div>Hi Bob<br>Latest update</div><blockquote>Older quoted text</blockquote>' },
  };
  row = toRowFromOutlook(htmlMsg, { body: true });
  body = row[11] ?? '';
  assert.ok(body.includes('Latest update'));
  assert.ok(!/Older quoted/i.test(body));
});

it('toRowFromOutlook: trims when quoted block has only > prefixes', () => {
  const msg = {
    id: 'mid',
    conversationId: 'cid',
    subject: 'Hello',
    toRecipients: [],
    from: {},
    ccRecipients: [],
    bccRecipients: [],
    receivedDateTime: 'Sat, 30 Aug 2025 15:44:48 -0700',
    bodyPreview: 'snippet',
    body: { contentType: 'text', content: 'Intro\nDetails\n\n> q1\n> q2\n> q3' },
  };
  const row = toRowFromOutlook(msg, { body: true });
  const body = row[11] ?? '';
  assert.ok(body.includes('Details'));
  assert.ok(!/q2/.test(body));
});
