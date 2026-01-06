import assert from 'assert';
import { OutlookQueryParameterSchema } from '../../../src/schemas/outlook-query-schema.ts';

describe('OutlookQueryParameterSchema', () => {
  it('accepts structured query objects', () => {
    const result = OutlookQueryParameterSchema.safeParse({ from: 'alice@example.com', date: { $gte: '2025-01-01' } });
    assert.ok(result.success, 'expected structured query to parse successfully');
    if (result.success) {
      assert.strictEqual(result.data.from, 'alice@example.com');
      assert.deepStrictEqual(result.data.date, { $gte: '2025-01-01' });
    }
  });

  it('parses JSON string inputs', () => {
    const jsonString = JSON.stringify({ date: { $lt: '2025-12-31' } });
    const result = OutlookQueryParameterSchema.safeParse(jsonString);
    assert.ok(result.success, 'expected JSON string to parse and validate');
    if (result.success) {
      assert.deepStrictEqual(result.data.date, { $lt: '2025-12-31' });
    }
  });

  it('rejects invalid JSON strings with friendly message', () => {
    const result = OutlookQueryParameterSchema.safeParse('{not json}');
    assert.ok(!result.success, 'expected invalid JSON string to fail validation');
    if (!result.success) {
      const [issue] = result.error.issues;
      assert.ok(issue.message.includes('Query must be valid JSON'), 'expected helpful message');
    }
  });

  it('supports kqlQuery inside JSON string', () => {
    const rawString = JSON.stringify({ kqlQuery: 'from:"alice@example.com" AND subject:report' });
    const result = OutlookQueryParameterSchema.safeParse(rawString);
    assert.ok(result.success, 'expected kqlQuery string to parse');
    if (result.success) {
      assert.strictEqual(result.data.kqlQuery, 'from:"alice@example.com" AND subject:report');
    }
  });
});
