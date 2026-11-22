import '../../../lib/env-loader.ts';
import assert from 'assert';
import { decodeNextPageToken } from '../../../../src/email/querying/pagination-token.ts';
import { cutToMaxItems, midPageToken, resumeCollection } from '../../../../src/email/querying/paging-decisions.ts';

const LINK = '/me/messages?$select=id&$top=3&$skip=3';

describe('midPageToken', () => {
  it('resumes in the page that was being read, one past the last message taken', () => {
    const decoded = decodeNextPageToken(midPageToken({ link: LINK, pageStart: 0, index: 1, mode: 'search' }));
    assert.strictEqual(decoded.link, LINK);
    assert.strictEqual(decoded.offset, 2);
    assert.strictEqual(decoded.mode, 'search');
  });

  it('adds the offset the page was resumed at', () => {
    const decoded = decodeNextPageToken(midPageToken({ link: LINK, pageStart: 2, index: 0, mode: 'filter' }));
    assert.strictEqual(decoded.offset, 3);
    assert.strictEqual(decoded.mode, 'filter');
  });

  it('carries no link for the first page', () => {
    const decoded = decodeNextPageToken(midPageToken({ pageStart: 0, index: 0, mode: 'search' }));
    assert.strictEqual(decoded.link, undefined);
    assert.strictEqual(decoded.offset, 1);
  });
});

describe('resumeCollection', () => {
  it('resumes the filter collection for a filter token', () => {
    assert.strictEqual(resumeCollection('filter'), 'filter');
  });

  it('runs the search collection for a search token, a plain token and no token', () => {
    assert.strictEqual(resumeCollection('search'), 'search');
    assert.strictEqual(resumeCollection(undefined), 'search');
  });
});

describe('cutToMaxItems', () => {
  const rows = (n: number) => Array.from({ length: n }, (_, i) => i);

  it('keeps a batch that fits and reports truncated only when a link remains at the limit', () => {
    assert.deepStrictEqual(cutToMaxItems(rows(3), 0, 10, true), { rows: rows(3), truncated: false });
    assert.deepStrictEqual(cutToMaxItems(rows(3), 7, 10, true), { rows: rows(3), truncated: true });
    assert.deepStrictEqual(cutToMaxItems(rows(3), 7, 10, false), { rows: rows(3), truncated: false });
  });

  it('cuts a batch that overshoots maxItems and reports truncated, with or without a link', () => {
    assert.deepStrictEqual(cutToMaxItems(rows(50), 50, 51, true), { rows: [0], truncated: true });
    assert.deepStrictEqual(cutToMaxItems(rows(50), 50, 51, false), { rows: [0], truncated: true });
  });

  it('keeps nothing once maxItems is reached', () => {
    assert.deepStrictEqual(cutToMaxItems(rows(2), 5, 5, false), { rows: [], truncated: true });
  });
});
