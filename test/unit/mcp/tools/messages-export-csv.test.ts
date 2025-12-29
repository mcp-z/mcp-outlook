import type { EnrichedExtra, Logger } from '@mcp-z/oauth-microsoft';
import assert from 'assert';
import { existsSync } from 'fs';
import { mkdir, readFile, rm } from 'fs/promises';
import * as path from 'path';
import createTool, { type Input, type Output } from '../../../../src/mcp/tools/messages-export-csv.js';
import type { StorageContext, StorageExtra } from '../../../../src/types.js';
import { createExtra, type TypedHandler } from '../../../lib/create-extra.js';
import createMiddlewareContext from '../../../lib/create-middleware-context.js';

/**
 * Simple CSV parser that handles quoted fields containing commas.
 * RFC 4180 compliant for basic CSV parsing.
 */
function parseCsvRow(row: string): string[] {
  const columns: string[] = [];
  let current = '';
  let inQuotes = false;

  for (let i = 0; i < row.length; i++) {
    const char = row[i];
    const nextChar = row[i + 1];

    if (char === '"') {
      if (inQuotes && nextChar === '"') {
        // Escaped quote (double quote)
        current += '"';
        i++; // Skip next quote
      } else {
        // Toggle quote mode
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      // Column separator (not inside quotes)
      columns.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }

  // Add last column
  columns.push(current.trim());

  return columns;
}

describe('Outlook messages export CSV tool (directory creation)', () => {
  let logger: Logger;
  let exportCsvHandler: TypedHandler<Input, EnrichedExtra & StorageExtra>;
  const tmpDir = path.join(process.cwd(), '.tmp');
  const testStorageDir = path.join(tmpDir, 'test-export-storage');
  const storageContext: StorageContext = {
    storageDir: testStorageDir,
    baseUrl: 'http://localhost:3000',
    transport: { type: 'http', port: 3000 },
  };
  const stdioStorageContext: StorageContext = {
    storageDir: testStorageDir,
    transport: { type: 'stdio' },
  };

  before(async () => {
    const middlewareContext = await createMiddlewareContext();
    const middleware = middlewareContext.middleware;

    const tool = createTool();
    const wrappedTool = middleware.withToolAuth(tool);
    exportCsvHandler = wrappedTool.handler as TypedHandler<Input>;

    // Ensure .tmp directory exists (parent for all test storage)
    await mkdir(tmpDir, { recursive: true });
  });

  after(async () => {
    // Clean up entire .tmp directory after all tests
    try {
      await rm(tmpDir, { recursive: true, force: true });
    } catch (err) {
      logger.warn({ err }, 'Failed to clean up .tmp directory');
    }
  });

  afterEach(async () => {
    // Clean up test storage directory after each test
    try {
      await rm(testStorageDir, { recursive: true, force: true });
    } catch (_err) {
      // Ignore errors if directory doesn't exist
    }
  });

  it('creates storage directory if it does not exist', async () => {
    // Ensure directory doesn't exist before test
    assert.strictEqual(existsSync(testStorageDir), false, 'Storage directory should not exist initially');

    // Export with minimal query (limit to 1 message for speed)
    const result = await exportCsvHandler(
      {
        query: {},
        maxItems: 1,
        filename: 'test-dir-creation.csv',
        contentType: 'text',
        excludeThreadHistory: false,
      },
      createExtra(stdioStorageContext)
    );

    // Validate success
    const structured = result?.structuredContent?.result as Output | undefined;
    assert.strictEqual(structured?.type, 'success', 'Expected success result');

    if (structured?.type === 'success') {
      // Verify directory was created
      assert.strictEqual(existsSync(testStorageDir), true, 'Storage directory should have been created');

      // Verify CSV file exists
      const csvPath = path.join(testStorageDir, structured.filename);
      assert.strictEqual(existsSync(csvPath), true, 'CSV file should exist');

      // Verify CSV has content (at least headers)
      const csvContent = await readFile(csvPath, 'utf-8');
      assert.ok(csvContent.includes('id'), 'CSV should contain header with id column');
    }
  });

  it('works when storage directory already exists', async () => {
    // Pre-create storage directory
    await mkdir(testStorageDir, { recursive: true });
    assert.strictEqual(existsSync(testStorageDir), true, 'Storage directory should exist before test');

    // Export with minimal query
    const result = await exportCsvHandler(
      {
        query: {},
        maxItems: 1,
        filename: 'test-existing-dir.csv',
        contentType: 'text',
        excludeThreadHistory: false,
      },
      createExtra(stdioStorageContext)
    );

    // Validate success
    const structured = result?.structuredContent?.result as Output | undefined;
    assert.strictEqual(structured?.type, 'success', 'Expected success result');

    if (structured?.type === 'success') {
      // Verify CSV file exists
      const csvPath = path.join(testStorageDir, structured.filename);
      assert.strictEqual(existsSync(csvPath), true, 'CSV file should exist');
    }
  });

  it('creates parent directories recursively if needed', async () => {
    // Use nested directory path
    const nestedStorageDir = path.join(tmpDir, 'deeply', 'nested', 'storage', 'dir');

    // Ensure nested path doesn't exist
    assert.strictEqual(existsSync(nestedStorageDir), false, 'Nested storage directory should not exist initially');

    const nestedStorageContext: StorageContext = {
      storageDir: nestedStorageDir,
      baseUrl: 'http://localhost:3000',
      transport: { type: 'http', port: 3000 },
    };

    const middlewareContext = await createMiddlewareContext();
    logger = middlewareContext.logger;
    const middleware = middlewareContext.middleware;
    const tool = createTool();
    const wrappedTool = middleware.withToolAuth(tool);
    const nestedHandler = wrappedTool.handler;

    // Export with minimal query
    const result = await nestedHandler(
      {
        query: {},
        maxItems: 1,
        filename: 'test-nested-dirs.csv',
        contentType: 'text',
        excludeThreadHistory: false,
      },
      createExtra(nestedStorageContext)
    );

    try {
      // Validate success
      const structured = (result as unknown as { structuredContent?: { result: Output } })?.structuredContent?.result as Output | undefined;
      assert.strictEqual(structured?.type, 'success', 'Expected success result');

      if (structured?.type === 'success') {
        // Verify all parent directories were created
        assert.strictEqual(existsSync(nestedStorageDir), true, 'Nested storage directory should have been created');
        assert.strictEqual(existsSync(path.join(tmpDir, 'deeply')), true, 'Parent directory should exist');
        assert.strictEqual(existsSync(path.join(tmpDir, 'deeply', 'nested')), true, 'Grandparent directory should exist');

        // Verify CSV file exists
        const csvPath = path.join(nestedStorageDir, structured.filename);
        assert.strictEqual(existsSync(csvPath), true, 'CSV file should exist in nested directory');
      }
    } finally {
      // Clean up nested directory structure
      await rm(path.join(tmpDir, 'deeply'), { recursive: true, force: true });
    }
  });

  it('exports valid CSV with headers and data', async () => {
    // Pre-create storage directory
    await mkdir(testStorageDir, { recursive: true });

    // Export with minimal query
    const result = await exportCsvHandler(
      {
        query: { from: 'noreply' }, // Common sender for testing
        maxItems: 5,
        filename: 'test-csv-content.csv',
        contentType: 'text',
        excludeThreadHistory: false,
      },
      createExtra(stdioStorageContext)
    );

    // Validate success
    const structured = result?.structuredContent?.result as Output | undefined;
    assert.strictEqual(structured?.type, 'success', 'Expected success result');

    if (structured?.type === 'success') {
      assert.ok(structured.rowCount >= 0, 'Should have row count');

      // Verify CSV file structure
      const csvPath = path.join(testStorageDir, structured.filename);
      const csvContent = await readFile(csvPath, 'utf-8');
      const lines = csvContent.split('\n').filter((line) => line.trim());

      // Verify header row
      assert.ok(lines.length > 0, 'CSV should have at least header row');
      const headerLine = lines[0];
      assert.ok(headerLine, 'Header line should exist');
      assert.ok(headerLine.includes('id'), 'Header should include id');
      assert.ok(headerLine.includes('subject'), 'Header should include subject');
      assert.ok(headerLine.includes('from'), 'Header should include from');
      assert.ok(headerLine.includes('to'), 'Header should include to');
      assert.ok(headerLine.includes('date'), 'Header should include date');

      // If messages were found, verify data rows have actual values
      if (structured.rowCount > 0) {
        assert.ok(lines.length > 1, 'CSV should have data rows when messages found');

        // Parse CSV headers to get column indices using proper CSV parser
        const headers = parseCsvRow(headerLine);
        const idIndex = headers.indexOf('id');
        const subjectIndex = headers.indexOf('subject');
        const fromIndex = headers.indexOf('from');
        const dateIndex = headers.indexOf('date');

        // Validate all required columns exist in header
        assert.ok(idIndex >= 0, 'id column should exist in header');
        assert.ok(subjectIndex >= 0, 'subject column should exist in header');
        assert.ok(fromIndex >= 0, 'from column should exist in header');
        assert.ok(dateIndex >= 0, 'date column should exist in header');

        // Check first data row has correct number of columns and validates critical fields
        const firstDataRow = lines[1];
        assert.ok(firstDataRow, 'Should have first data row');
        const columns = parseCsvRow(firstDataRow);

        // CSV rows should have same number of columns as headers
        assert.ok(columns.length === headers.length, `CSV row should have ${headers.length} columns, got ${columns.length}`);

        // Validate critical fields that must always have values
        assert.ok(columns[idIndex] !== undefined, 'id column should exist in data row');
        assert.ok(columns[idIndex].length > 0, 'id field must have a non-empty value');

        assert.ok(columns[dateIndex] !== undefined, 'date column should exist in data row');
        assert.ok(columns[dateIndex].length > 0, 'date field must have a non-empty value for real messages');

        // Validate optional fields exist as columns (but can be empty strings)
        assert.ok(columns[subjectIndex] !== undefined, 'subject column should exist in data row');
        assert.ok(columns[fromIndex] !== undefined, 'from column should exist in data row');
      }
    }
  });

  it('returns absolute file:// URI for stdio transport', async () => {
    // Pre-create storage directory
    await mkdir(testStorageDir, { recursive: true });

    // Export with minimal query
    const result = await exportCsvHandler(
      {
        query: {},
        maxItems: 1,
        filename: 'test-uri-format.csv',
        contentType: 'text',
        excludeThreadHistory: false,
      },
      createExtra(stdioStorageContext)
    );

    // Validate success
    const structured = result?.structuredContent?.result as Output | undefined;
    assert.strictEqual(structured?.type, 'success', 'Expected success result');

    if (structured?.type === 'success') {
      // Verify URI format
      const uri = structured.uri;
      assert.ok(uri.startsWith('file://'), 'URI should start with file://');
      assert.ok(path.isAbsolute(uri.replace('file://', '')), 'URI should contain absolute path');
      assert.ok(uri.includes(testStorageDir), 'URI should include storage directory path');
      assert.ok(uri.includes(structured.filename), 'URI should include filename');

      // Verify file exists at the URI path
      const filePath = uri.replace('file://', '');
      assert.strictEqual(existsSync(filePath), true, 'File should exist at URI path');
    }
  });

  describe('negative test cases - error handling', () => {
    beforeEach(async () => {
      // Ensure storage directory exists for negative tests
      await mkdir(testStorageDir, { recursive: true });
    });

    it('handles empty query results gracefully', async () => {
      // Query that should match no messages
      const result = await exportCsvHandler(
        {
          query: { from: `nonexistent-${Date.now()}@invalid-domain.com` },
          maxItems: 10,
          filename: 'test-empty-results.csv',
          contentType: 'text',
          excludeThreadHistory: false,
        },
        createExtra(storageContext)
      );

      const structured = result?.structuredContent?.result as Output | undefined;
      assert.strictEqual(structured?.type, 'success', 'Should succeed even with no results');

      if (structured?.type === 'success') {
        // Verify CSV file exists but has only headers
        const csvPath = path.join(testStorageDir, structured.filename);
        assert.strictEqual(existsSync(csvPath), true, 'CSV file should exist');

        const csvContent = await readFile(csvPath, 'utf-8');
        const lines = csvContent.split('\n').filter((line) => line.trim());

        assert.ok(lines.length >= 1, 'Should have at least header row');
        assert.strictEqual(structured.rowCount, 0, 'Should report 0 rows for empty results');
      }
    });

    it('handles invalid query gracefully', async () => {
      // Query with invalid field
      const result = await exportCsvHandler(
        {
          query: { invalidField: 'test' },
          maxItems: 5,
          filename: 'test-invalid-query.csv',
          contentType: 'text',
          excludeThreadHistory: false,
        } as Input,
        createExtra(storageContext)
      );

      // Should either succeed with empty/filtered results or return error
      const structured = result?.structuredContent?.result as Output | undefined;
      assert.ok(structured, 'Should return structured result');
      assert.ok(['success', 'auth_required'].includes(structured?.type), 'Should handle invalid query gracefully');

      if (structured?.type === 'success') {
        // File should exist even if no results
        const csvPath = path.join(testStorageDir, structured.filename);
        assert.strictEqual(existsSync(csvPath), true, 'CSV file should exist');
      }
    });

    it('validates CSV structure with minimal data', async () => {
      // Export with very restrictive query
      const result = await exportCsvHandler(
        {
          query: { date: { $gte: '2099-01-01', $lt: '2099-01-02' } }, // Future date
          maxItems: 1,
          filename: 'test-minimal-data.csv',
          contentType: 'text',
          excludeThreadHistory: false,
        },
        createExtra(storageContext)
      );

      const structured = result?.structuredContent?.result as Output | undefined;
      assert.strictEqual(structured?.type, 'success', 'Should succeed');

      if (structured?.type === 'success') {
        // Verify CSV has proper structure even with no data
        const csvPath = path.join(testStorageDir, structured.filename);
        const csvContent = await readFile(csvPath, 'utf-8');
        const lines = csvContent.split('\n').filter((line) => line.trim());

        // Must have header row
        assert.ok(lines.length >= 1, 'Should have header row');
        const headerLine = lines[0];
        assert.ok(headerLine, 'Header line should exist');

        // Verify all expected columns are in header
        assert.ok(headerLine.includes('id'), 'Header should include id');
        assert.ok(headerLine.includes('from'), 'Header should include from');
        assert.ok(headerLine.includes('to'), 'Header should include to');
        assert.ok(headerLine.includes('subject'), 'Header should include subject');
        assert.ok(headerLine.includes('date'), 'Header should include date');
      }
    });

    it('handles maxItems=0 edge case', async () => {
      const result = await exportCsvHandler(
        {
          query: {},
          maxItems: 0,
          filename: 'test-max-items-zero.csv',
          contentType: 'text',
          excludeThreadHistory: false,
        },
        createExtra(storageContext)
      );

      const structured = result?.structuredContent?.result as Output | undefined;
      assert.strictEqual(structured?.type, 'success', 'Should succeed with maxItems=0');

      if (structured?.type === 'success') {
        // Should create file with just headers
        const csvPath = path.join(testStorageDir, structured.filename);
        assert.strictEqual(existsSync(csvPath), true, 'CSV file should exist');

        const csvContent = await readFile(csvPath, 'utf-8');
        const lines = csvContent.split('\n').filter((line) => line.trim());

        assert.ok(lines.length >= 1, 'Should have header row');
        assert.strictEqual(structured.rowCount, 0, 'Should report 0 rows');
      }
    });

    it('validates data integrity with malformed date query', async () => {
      // Query with invalid date format
      // Malformed dates cause Graph API to throw invalid filter clause errors
      try {
        await exportCsvHandler(
          {
            query: { date: { $gte: 'invalid-date', $lt: 'also-invalid' } },
            maxItems: 5,
            filename: 'test-malformed-date.csv',
            contentType: 'text',
            excludeThreadHistory: false,
          } as Input,
          createExtra(storageContext)
        );
        assert.fail('Expected McpError to be thrown for malformed date query');
      } catch (err) {
        assert.ok(err instanceof Error, 'Error should be an Error instance');
        assert.ok(err.message.includes('Error exporting messages') || err.message.includes('Invalid filter'), 'Error should indicate invalid query');
      }
    });
  });
});
