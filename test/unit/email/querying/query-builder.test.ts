import assert from 'assert';
import { toGraphFilter, toOutlookFilter } from '../../../../src/email/querying/query-builder.ts';
import type { OutlookSystemCategory } from '../../../../src/schemas/outlook-query-schema.ts';

describe('toGraphFilter - basic field queries', () => {
  it('handles from field with email address', () => {
    const result = toGraphFilter({ from: 'alice@example.com' });
    assert.ok(result.filter?.includes('from/emailAddress/address'));
    assert.ok(result.filter?.includes('from/emailAddress/name'));
    assert.ok(result.filter?.includes('alice@example.com'));
    assert.ok(result.filter?.includes('or'));
  });

  it('handles from field without @ (name-only) - client-side filter', () => {
    const result = toGraphFilter({ from: 'Alice' });
    // Name-only from should skip server-side filter (handled client-side)
    assert.strictEqual(result.filter, null);
    assert.strictEqual(result.requireBodyClientFilter, false);
  });

  it('handles to field', () => {
    const result = toGraphFilter({ to: 'bob@example.com' });
    assert.ok(result.filter?.includes('toRecipients/any'));
    assert.ok(result.filter?.includes('bob@example.com'));
  });

  it('handles cc field', () => {
    const result = toGraphFilter({ cc: 'charlie@example.com' });
    assert.ok(result.filter?.includes('ccRecipients/any'));
    assert.ok(result.filter?.includes('charlie@example.com'));
  });

  it('handles bcc field', () => {
    const result = toGraphFilter({ bcc: 'dave@example.com' });
    assert.ok(result.filter?.includes('bccRecipients/any'));
    assert.ok(result.filter?.includes('dave@example.com'));
  });

  it('handles subject field (startswith limitation)', () => {
    const result = toGraphFilter({ subject: 'meeting' });
    assert.ok(result.filter?.includes('startswith(subject'));
    assert.ok(result.filter?.includes('meeting'));
  });

  it('handles body field (client-side filter flag)', () => {
    const result = toGraphFilter({ body: 'report' });
    // Body should skip server-side filter and set requireBodyClientFilter flag
    assert.strictEqual(result.filter, null);
    assert.strictEqual(result.requireBodyClientFilter, true);
    // Should add to search parameter
    assert.ok(result.search?.includes('report'));
  });

  it('handles text field (client-side filter flag)', () => {
    const result = toGraphFilter({ text: 'budget' });
    // Text should skip server-side filter and set requireBodyClientFilter flag
    assert.strictEqual(result.filter, null);
    assert.strictEqual(result.requireBodyClientFilter, true);
    // Should add to search parameter
    assert.ok(result.search?.includes('budget'));
  });
});

describe('toGraphFilter - field operators', () => {
  it('handles $any operator (OR logic) for from', () => {
    const result = toGraphFilter({ from: { $any: ['alice@example.com', 'bob@example.com'] } });
    assert.ok(result.filter?.includes('alice@example.com'));
    assert.ok(result.filter?.includes('bob@example.com'));
    assert.ok(result.filter?.includes('or'));
  });

  it('handles $all operator (AND logic) for to', () => {
    const result = toGraphFilter({ to: { $all: ['alice@example.com', 'bob@example.com'] } });
    assert.ok(result.filter?.includes('toRecipients/any'));
    assert.ok(result.filter?.includes('alice@example.com'));
    assert.ok(result.filter?.includes('bob@example.com'));
    assert.ok(result.filter?.includes('and'));
  });

  it('handles $none operator (NOT logic) for subject', () => {
    const result = toGraphFilter({ subject: { $none: ['spam', 'ads'] } });
    assert.ok(result.filter?.includes('not'));
    assert.ok(result.filter?.includes('startswith(subject'));
  });

  it('handles multiple field operators in same query', () => {
    const result = toGraphFilter({
      $and: [{ from: { $any: ['alice@example.com', 'bob@example.com'] } }, { to: { $all: ['charlie@example.com', 'dave@example.com'] } }],
    });
    assert.ok(result.filter?.includes('alice@example.com'));
    assert.ok(result.filter?.includes('bob@example.com'));
    assert.ok(result.filter?.includes('charlie@example.com'));
    assert.ok(result.filter?.includes('dave@example.com'));
  });
});

describe('toGraphFilter - category queries', () => {
  it('handles single category (work)', () => {
    const result = toGraphFilter({ categories: 'work' });
    assert.ok(result.filter?.includes('categories/any(c: c eq'));
    assert.ok(result.filter?.includes('Work'));
  });

  it('handles multiple categories with $any', () => {
    const result = toGraphFilter({ categories: { $any: ['work', 'personal', 'family'] } });
    assert.ok(result.filter?.includes('Work'));
    assert.ok(result.filter?.includes('Personal'));
    assert.ok(result.filter?.includes('Family'));
    assert.ok(result.filter?.includes('or'));
  });

  it('throws error for invalid category names (fail fast)', () => {
    // Should throw on the first invalid category
    // Cast through unknown to simulate untrusted runtime input that should trigger validation
    const invalidCategories = ['invalid', 'work', 'bad'] as unknown as OutlookSystemCategory[];
    assert.throws(() => toGraphFilter({ categories: { $any: invalidCategories } }), /Invalid Outlook category: "invalid"/);
  });

  it('maps all valid categories correctly', () => {
    const testCases: Array<{ category: 'personal' | 'work' | 'family' | 'travel' | 'important' | 'urgent'; expected: string }> = [
      { category: 'personal', expected: 'Personal' },
      { category: 'work', expected: 'Work' },
      { category: 'family', expected: 'Family' },
      { category: 'travel', expected: 'Travel' },
      { category: 'important', expected: 'Important' },
      { category: 'urgent', expected: 'Urgent' },
    ];

    for (const { category, expected } of testCases) {
      const result = toGraphFilter({ categories: category });
      assert.ok(result.filter?.includes(expected), `Category ${category} should map to ${expected}`);
    }
  });

  it('combines categories with other fields', () => {
    const result = toGraphFilter({
      $and: [{ categories: 'work' }, { from: 'alice@example.com' }],
    });
    assert.ok(result.filter?.includes('categories/any(c: c eq'));
    assert.ok(result.filter?.includes('Work'));
    assert.ok(result.filter?.includes('from/emailAddress'));
    assert.ok(result.filter?.includes('alice@example.com'));
  });
});

describe('toGraphFilter - exact phrase matching', () => {
  it('handles exactPhrase in search parameter', () => {
    const result = toGraphFilter({ exactPhrase: 'quarterly report' });
    // KQL escapes quotes, so we get \\"quarterly report\\"
    assert.ok(result.search?.includes('quarterly report'));
    assert.ok(result.search?.includes('\\"'));
    assert.strictEqual(result.filter, null);
  });

  it('escapes special characters in exactPhrase', () => {
    const result = toGraphFilter({ exactPhrase: 'meeting (urgent)' });
    assert.ok(result.search?.includes('meeting'));
    assert.ok(result.search?.includes('\\('));
    assert.ok(result.search?.includes('\\)'));
    assert.ok(result.search?.includes('\\"'));
  });

  it('combines exactPhrase with other filters', () => {
    const result = toGraphFilter({
      $and: [{ exactPhrase: 'quarterly report' }, { from: 'alice@example.com' }],
    });
    assert.ok(result.search?.includes('quarterly report'));
    assert.ok(result.search?.includes('\\"'));
    assert.ok(result.filter?.includes('from/emailAddress'));
    assert.ok(result.filter?.includes('alice@example.com'));
  });
});

describe('toGraphFilter - attachment flag', () => {
  it('handles hasAttachment = true', () => {
    const result = toGraphFilter({ hasAttachment: true });
    assert.strictEqual(result.filter, 'hasAttachments eq true');
  });

  it('combines hasAttachment with other queries', () => {
    const result = toGraphFilter({
      $and: [{ hasAttachment: true }, { from: 'alice@example.com' }],
    });
    assert.ok(result.filter?.includes('hasAttachments eq true'));
    assert.ok(result.filter?.includes('from/emailAddress'));
    assert.ok(result.filter?.includes('alice@example.com'));
  });
});

describe('toGraphFilter - date ranges', () => {
  it('handles date with $gte only (receivedDateTime ge)', () => {
    const result = toGraphFilter({ date: { $gte: '2025-01-15' } });
    assert.ok(result.filter?.includes('receivedDateTime ge 2025-01-15T00:00:00Z'));
  });

  it('handles date with $lt only (receivedDateTime lt)', () => {
    const result = toGraphFilter({ date: { $lt: '2025-01-20' } });
    assert.ok(result.filter?.includes('receivedDateTime lt 2025-01-20T00:00:00Z'));
  });

  it('handles date with both $gte and $lt (range)', () => {
    const result = toGraphFilter({ date: { $gte: '2025-01-15', $lt: '2025-01-20' } });
    assert.ok(result.filter?.includes('receivedDateTime ge 2025-01-15T00:00:00Z'));
    assert.ok(result.filter?.includes('receivedDateTime lt 2025-01-20T00:00:00Z'));
    assert.ok(result.filter?.includes('and'));
  });

  it('uses ISO 8601 format with timezone', () => {
    const result = toGraphFilter({ date: { $gte: '2025-01-15' } });
    assert.ok(result.filter?.includes('T00:00:00Z'));
  });
});

describe('toGraphFilter - logical operators', () => {
  it('handles $and operator with multiple conditions', () => {
    const result = toGraphFilter({
      $and: [{ from: 'alice@example.com' }, { subject: 'meeting' }, { hasAttachment: true }],
    });
    assert.ok(result.filter?.includes('from/emailAddress'));
    assert.ok(result.filter?.includes('alice@example.com'));
    assert.ok(result.filter?.includes('startswith(subject'));
    assert.ok(result.filter?.includes('meeting'));
    assert.ok(result.filter?.includes('hasAttachments eq true'));
    assert.ok(result.filter?.includes('and'));
  });

  it('handles $or operator with multiple conditions', () => {
    const result = toGraphFilter({
      $or: [{ from: 'alice@example.com' }, { from: 'bob@example.com' }],
    });
    assert.ok(result.filter?.includes('alice@example.com'));
    assert.ok(result.filter?.includes('bob@example.com'));
    assert.ok(result.filter?.includes('or'));
  });

  it('handles $not operator', () => {
    const result = toGraphFilter({
      $not: { subject: 'spam' },
    });
    assert.ok(result.filter?.includes('not'));
    assert.ok(result.filter?.includes('startswith(subject'));
  });

  it('handles nested logical operators', () => {
    const result = toGraphFilter({
      $and: [
        { from: 'alice@example.com' },
        {
          $or: [{ subject: 'meeting' }, { subject: 'conference' }],
        },
      ],
    });
    assert.ok(result.filter?.includes('from/emailAddress'));
    assert.ok(result.filter?.includes('alice@example.com'));
    assert.ok(result.filter?.includes('meeting'));
    assert.ok(result.filter?.includes('conference'));
    assert.ok(result.filter?.includes('or'));
    assert.ok(result.filter?.includes('and'));
  });

  it('handles complex nested query combinations', () => {
    const result = toGraphFilter({
      $and: [
        { from: 'alice@example.com' },
        {
          $or: [{ subject: 'meeting' }, { subject: 'conference' }],
        },
        {
          $not: { categories: 'personal' },
        },
      ],
    });
    assert.ok(result.filter?.includes('from/emailAddress'));
    assert.ok(result.filter?.includes('alice@example.com'));
    assert.ok(result.filter?.includes('meeting'));
    assert.ok(result.filter?.includes('conference'));
    assert.ok(result.filter?.includes('not'));
    assert.ok(result.filter?.includes('Personal'));
  });
});

describe('toGraphFilter - OData filter string generation', () => {
  it('generates valid OData filter syntax', () => {
    const result = toGraphFilter({ from: 'alice@example.com' });
    // Should have nested property access
    assert.ok(result.filter?.includes('from/emailAddress/address'));
    // Should use 'eq' operator
    assert.ok(result.filter?.includes(' eq '));
    // Should have quoted values
    assert.ok(result.filter?.includes("'alice@example.com'"));
  });

  it('uses eq/and/or/not operators correctly', () => {
    const result = toGraphFilter({
      $and: [
        { from: 'alice@example.com' },
        {
          $or: [{ to: 'bob@example.com' }, { to: 'charlie@example.com' }],
        },
        { $not: { subject: 'spam' } },
      ],
    });
    assert.ok(result.filter?.includes(' eq '));
    assert.ok(result.filter?.includes(' and '));
    assert.ok(result.filter?.includes(' or '));
    assert.ok(result.filter?.includes('not '));
  });

  it('escapes single quotes in values (doubled)', () => {
    const result = toGraphFilter({ from: "alice'o'malley@example.com" });
    assert.ok(result.filter?.includes("''"));
  });

  it('handles nested properties (from/emailAddress/address)', () => {
    const result = toGraphFilter({ from: 'alice@example.com' });
    assert.ok(result.filter?.includes('from/emailAddress/address'));
    assert.ok(result.filter?.includes('from/emailAddress/name'));
  });

  it('handles any() for collections (categories, recipients)', () => {
    const resultCat = toGraphFilter({ categories: 'work' });
    assert.ok(resultCat.filter?.includes('categories/any(c:'));

    const resultTo = toGraphFilter({ to: 'alice@example.com' });
    assert.ok(resultTo.filter?.includes('toRecipients/any(r:'));
  });
});

describe('toGraphFilter - KQL search parameter', () => {
  it('generates search parameter for subject/body/text', () => {
    const result = toGraphFilter({ subject: 'meeting' });
    assert.ok(result.search?.includes('meeting'));
  });

  it('escapes KQL special characters', () => {
    const result = toGraphFilter({ subject: 'test:value(with)special-chars_here' });
    // KQL special characters should be escaped
    assert.ok(result.search?.includes('\\:'));
    assert.ok(result.search?.includes('\\('));
    assert.ok(result.search?.includes('\\)'));
    assert.ok(result.search?.includes('\\-'));
    assert.ok(result.search?.includes('\\_'));
  });

  it('combines multiple search terms with OR', () => {
    const result = toGraphFilter({ subject: { $any: ['meeting', 'conference', 'event'] } });
    assert.ok(result.search?.includes('meeting'));
    assert.ok(result.search?.includes('conference'));
    assert.ok(result.search?.includes('event'));
    assert.ok(result.search?.includes(' OR '));
  });

  it('search is null when no searchable fields', () => {
    const result = toGraphFilter({ from: 'alice@example.com' });
    assert.strictEqual(result.search, null);
  });
});

describe('toGraphFilter - requireBodyClientFilter flag', () => {
  it('sets flag to true for body queries', () => {
    const result = toGraphFilter({ body: 'report' });
    assert.strictEqual(result.requireBodyClientFilter, true);
  });

  it('sets flag to true for text queries', () => {
    const result = toGraphFilter({ text: 'budget' });
    assert.strictEqual(result.requireBodyClientFilter, true);
  });

  it('flag is false for other query types', () => {
    const result1 = toGraphFilter({ from: 'alice@example.com' });
    assert.strictEqual(result1.requireBodyClientFilter, false);

    const result2 = toGraphFilter({ subject: 'meeting' });
    assert.strictEqual(result2.requireBodyClientFilter, false);
  });

  it('flag indicates server-side limitations', () => {
    // Body and text cannot be filtered server-side in Microsoft Graph
    const result = toGraphFilter({
      $and: [{ from: 'alice@example.com' }, { body: 'report' }],
    });
    assert.strictEqual(result.requireBodyClientFilter, true);
  });
});

describe('toGraphFilter - edge cases', () => {
  it('handles empty query object', () => {
    const result = toGraphFilter({});
    assert.strictEqual(result.filter, null);
    assert.strictEqual(result.search, null);
    assert.strictEqual(result.requireBodyClientFilter, false);
  });

  it('throws error on empty strings in field operators (fail fast)', () => {
    // Should throw on the first empty string
    assert.throws(() => toGraphFilter({ from: { $any: ['', '  ', 'alice@example.com'] } }), /Invalid from value: empty string/);
  });

  it('handles single-element field operator arrays', () => {
    const result = toGraphFilter({ categories: { $any: ['work'] } });
    assert.ok(result.filter?.includes('Work'));
    // Single element in $any should not generate multiple OR clauses at the operator level
    // (though internal field logic like from's address/name check may still use 'or')
    const orCount = (result.filter?.match(/ or /g) || []).length;
    assert.strictEqual(orCount, 0, 'single category should not have OR operator');
  });

  it('combines all query types in one complex query', () => {
    const result = toGraphFilter({
      $and: [{ from: { $any: ['alice@example.com', 'bob@example.com'] } }, { subject: 'meeting' }, { categories: 'work' }, { hasAttachment: true }, { date: { $gte: '2025-01-15' } }],
    });
    assert.ok(result.filter?.includes('alice@example.com'));
    assert.ok(result.filter?.includes('bob@example.com'));
    assert.ok(result.filter?.includes('Work'));
    assert.ok(result.filter?.includes('hasAttachments eq true'));
    assert.ok(result.filter?.includes('receivedDateTime ge'));
    assert.ok(result.search?.includes('meeting'));
  });
});

describe('toGraphFilter - OData limitations', () => {
  it('subject uses startswith (not contains)', () => {
    const result = toGraphFilter({ subject: 'meeting' });
    assert.ok(result.filter?.includes('startswith(subject'));
    assert.ok(!result.filter?.includes('contains('));
  });

  it('body/text skip server-side filter (client-side)', () => {
    const resultBody = toGraphFilter({ body: 'report' });
    assert.strictEqual(resultBody.filter, null);
    assert.strictEqual(resultBody.requireBodyClientFilter, true);

    const resultText = toGraphFilter({ text: 'budget' });
    assert.strictEqual(resultText.filter, null);
    assert.strictEqual(resultText.requireBodyClientFilter, true);
  });

  it('from name-only skips server filter', () => {
    const result = toGraphFilter({ from: 'Alice' });
    assert.strictEqual(result.filter, null);
  });

  it('no toLower() function support', () => {
    const result = toGraphFilter({ from: 'alice@example.com' });
    // Should not contain toLower() calls
    assert.ok(!result.filter?.includes('toLower('));
  });
});

describe('toGraphFilter - label queries', () => {
  it('toGraphFilter handles label queries with case-sensitive OData filters', () => {
    // Test single label
    const singleParsed = { label: 'Work' };
    const singleResult = toGraphFilter(singleParsed);

    assert.ok(singleResult.filter?.includes('categories/any(c: c eq'), 'expected categories/any OData filter');
    assert.ok(singleResult.filter?.includes('Work'), 'expected case-sensitive Work value');

    // Test case-sensitive handling (no normalization)
    const caseParsed = { label: { $any: ['Work', 'work', 'WORK'] } };
    const caseResult = toGraphFilter(caseParsed);

    assert.ok(caseResult.filter?.includes('Work'), 'expected case-sensitive Work');
    assert.ok(caseResult.filter?.includes('work'), 'expected case-sensitive work');
    assert.ok(caseResult.filter?.includes('WORK'), 'expected case-sensitive WORK');
    assert.ok(caseResult.filter?.includes(' or '), 'expected or for multiple labels');
  });

  it('toGraphFilter handles multiple label queries with OR logic', () => {
    const parsed = { label: { $any: ['Important', 'Urgent', 'Work'] } };
    const result = toGraphFilter(parsed);

    assert.ok(result.filter?.includes('Important'), 'expected Important label');
    assert.ok(result.filter?.includes('Urgent'), 'expected Urgent label');
    assert.ok(result.filter?.includes('Work'), 'expected Work label');
    assert.ok(result.filter?.includes(' or '), 'expected OR operator for multiple labels');
  });

  it('toGraphFilter handles label queries with AND logic', () => {
    const parsed = { label: { $all: ['Important', 'Work'] } };
    const result = toGraphFilter(parsed);

    assert.ok(result.filter?.includes('Important'), 'expected Important label');
    assert.ok(result.filter?.includes('Work'), 'expected Work label');
    assert.ok(result.filter?.includes(' and '), 'expected AND operator for $all labels');
  });

  it('toGraphFilter handles label queries with NOT logic', () => {
    const parsed = { label: { $none: ['Spam', 'Trash'] } };
    const result = toGraphFilter(parsed);

    assert.ok(result.filter?.includes('not'), 'expected NOT operator');
    assert.ok(result.filter?.includes('Spam'), 'expected Spam label in NOT clause');
    assert.ok(result.filter?.includes('Trash'), 'expected Trash label in NOT clause');
  });

  it('toGraphFilter combines label queries with other fields', () => {
    const parsed = {
      $and: [{ label: 'Work' }, { from: 'alice@example.com' }, { subject: 'meeting' }],
    };
    const result = toGraphFilter(parsed);

    assert.ok(result.filter?.includes('categories/any(c: c eq'), 'expected label OData filter');
    assert.ok(result.filter?.includes('Work'), 'expected Work label');
    assert.ok(result.filter?.includes('from/emailAddress'), 'expected from query');
    assert.ok(result.filter?.includes('alice@example.com'), 'expected alice email');
    assert.ok(result.search?.includes('meeting'), 'expected meeting in search');
  });

  it('toGraphFilter handles label queries with special characters and escaping', () => {
    const parsed = { label: { $any: ['project-2024', 'team@work', "label'with'quotes"] } };
    const result = toGraphFilter(parsed);

    assert.ok(result.filter?.includes('project-2024'), 'expected hyphenated label');
    assert.ok(result.filter?.includes('team@work'), 'expected label with @ symbol');
    assert.ok(result.filter?.includes("''"), 'expected escaped single quotes (doubled)');
  });

  it('throws error on empty label values (fail fast)', () => {
    const parsed = { label: { $any: ['', '  ', 'Valid'] } };

    // Should throw on the first empty string
    assert.throws(() => toGraphFilter(parsed), /Invalid label value: empty string/);
  });

  it("toGraphFilter label queries don't interfere with search parameter", () => {
    const parsed = {
      $and: [{ label: 'Work' }, { subject: 'meeting' }],
    };
    const result = toGraphFilter(parsed);

    assert.ok(result.filter?.includes('categories/any(c: c eq'), 'expected label in filter');
    assert.ok(result.filter?.includes('Work'), 'expected Work label');
    assert.ok(result.search?.includes('meeting'), 'expected meeting in search parameter');
    assert.ok(!result.search?.includes('Work'), 'Work should not be in search parameter');
  });
});

describe('toGraphFilter - real-world query examples', () => {
  it('finds emails from specific sender with attachment', () => {
    const result = toGraphFilter({
      $and: [{ from: 'alice@example.com' }, { hasAttachment: true }],
    });
    assert.ok(result.filter?.includes('from/emailAddress'));
    assert.ok(result.filter?.includes('alice@example.com'));
    assert.ok(result.filter?.includes('hasAttachments eq true'));
  });

  it('finds emails in work category with subject keyword', () => {
    const result = toGraphFilter({
      $and: [{ categories: 'work' }, { subject: 'invoice' }],
    });
    assert.ok(result.filter?.includes('categories/any(c: c eq'));
    assert.ok(result.filter?.includes('Work'));
    assert.ok(result.search?.includes('invoice'));
  });

  it('finds emails in date range from multiple senders', () => {
    const result = toGraphFilter({
      $and: [{ from: { $any: ['alice@example.com', 'bob@example.com'] } }, { date: { $gte: '2025-01-15', $lt: '2025-01-20' } }],
    });
    assert.ok(result.filter?.includes('alice@example.com'));
    assert.ok(result.filter?.includes('bob@example.com'));
    assert.ok(result.filter?.includes('receivedDateTime ge 2025-01-15'));
    assert.ok(result.filter?.includes('receivedDateTime lt 2025-01-20'));
  });

  it('complex query with all features combined', () => {
    const result = toGraphFilter({
      $and: [
        { from: { $any: ['alice@example.com', 'bob@example.com'] } },
        { subject: 'meeting' },
        { categories: 'work' },
        { hasAttachment: true },
        { date: { $gte: '2025-01-15' } },
        {
          $not: { label: 'archived' },
        },
      ],
    });
    assert.ok(result.filter?.includes('alice@example.com'));
    assert.ok(result.search?.includes('meeting'));
    assert.ok(result.filter?.includes('Work'));
    assert.ok(result.filter?.includes('hasAttachments eq true'));
    assert.ok(result.filter?.includes('receivedDateTime ge'));
    assert.ok(result.filter?.includes('not'));
    assert.ok(result.filter?.includes('archived'));
  });

  it('searches by exact phrase with filters', () => {
    const result = toGraphFilter({
      $and: [{ exactPhrase: 'quarterly report' }, { from: 'finance@example.com' }, { hasAttachment: true }],
    });
    // KQL escapes quotes
    assert.ok(result.search?.includes('quarterly report'));
    assert.ok(result.search?.includes('\\"'));
    assert.ok(result.filter?.includes('from/emailAddress'));
    assert.ok(result.filter?.includes('finance@example.com'));
    assert.ok(result.filter?.includes('hasAttachments eq true'));
  });
});

describe('toOutlookFilter - integration', () => {
  it('combines OData filter + KQL search + filters', () => {
    const result = toOutlookFilter({
      $and: [{ from: 'alice@example.com' }, { subject: 'meeting' }],
    });
    assert.ok(result.filter?.includes('from/emailAddress'));
    assert.ok(result.filter?.includes('alice@example.com'));
    assert.ok(result.search?.includes('meeting'));
    assert.ok(Array.isArray(result.filters.fromIncludes));
    assert.ok(result.filters.fromIncludes.includes('alice@example.com'));
    assert.ok(Array.isArray(result.filters.subjectIncludes));
    assert.ok(result.filters.subjectIncludes.includes('meeting'));
  });

  it('returns requireBodyClientFilter flag', () => {
    const result = toOutlookFilter({ body: 'report' });
    assert.strictEqual(result.requireBodyClientFilter, true);
  });

  it('returns extracted filters', () => {
    const result = toOutlookFilter({
      $and: [{ from: 'alice@example.com' }, { to: 'bob@example.com' }, { subject: 'meeting' }, { categories: 'work' }],
    });
    assert.deepStrictEqual(result.filters.fromIncludes, ['alice@example.com']);
    assert.deepStrictEqual(result.filters.toIncludes, ['bob@example.com']);
    assert.deepStrictEqual(result.filters.subjectIncludes, ['meeting']);
    assert.deepStrictEqual(result.filters.categoriesIncludes, ['work']);
  });

  it('all components work together', () => {
    const result = toOutlookFilter({
      $and: [{ from: 'alice@example.com' }, { body: 'report' }, { hasAttachment: true }],
    });
    // OData filter
    assert.ok(result.filter?.includes('from/emailAddress'));
    assert.ok(result.filter?.includes('hasAttachments eq true'));
    // KQL search
    assert.ok(result.search?.includes('report'));
    // Client filter flag
    assert.strictEqual(result.requireBodyClientFilter, true);
    // Extracted filters
    assert.deepStrictEqual(result.filters.fromIncludes, ['alice@example.com']);
    assert.deepStrictEqual(result.filters.bodyIncludes, ['report']);
    assert.strictEqual(result.filters.hasAttachment, true);
  });
});
