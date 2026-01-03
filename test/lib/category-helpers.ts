import type { Client } from '@microsoft/microsoft-graph-client';
import { setTimeout as delay } from 'timers/promises';
import type { Logger } from '../../src/types.ts';

// Type for error objects that may have status/code properties
type ErrorWithStatus = {
  status?: number;
  statusCode?: number;
  code?: number | string;
};

/**
 * Create a test category in Outlook
 */
export async function createTestCategory(graph: Client, opts: { displayName?: string; color?: string } = {}): Promise<string> {
  const displayName = opts.displayName || `ci-test-category-${Date.now()}`;
  const color = opts.color || 'preset0';

  const response = await graph.api('/me/outlook/masterCategories').post({
    displayName,
    color,
  });

  const categoryId = response.id;
  if (!categoryId) throw new Error('createTestCategory: expected created category id');
  return categoryId;
}

/**
 * Delete a test category created with createTestCategory.
 * Throws on any error - close failures indicate test problems that need to be visible.
 */
export async function deleteTestCategory(graph: Client, id: string, logger: Logger): Promise<void> {
  const startTime = Date.now();
  try {
    await graph.api(`/me/outlook/masterCategories/${encodeURIComponent(id)}`).delete();
    logger.debug('Test category close successful', { categoryId: id, duration: Date.now() - startTime });
  } catch (e: unknown) {
    const duration = Date.now() - startTime;
    const errorDetails = {
      categoryId: id,
      duration,
      error: e instanceof Error ? e.message : String(e),
      status: e && typeof e === 'object' && ('status' in e || 'statusCode' in e) ? (e as ErrorWithStatus).status || (e as ErrorWithStatus).statusCode : undefined,
      code: e && typeof e === 'object' && 'code' in e ? (e as ErrorWithStatus).code : undefined,
    };

    logger.error('Test category close failed', errorDetails);
    throw e; // Always throw - if we're deleting it, it should exist
  }
}

/**
 * Check if a category exists
 */
export async function categoryExists(graph: Client, id: string): Promise<boolean> {
  try {
    await graph.api(`/me/outlook/masterCategories/${encodeURIComponent(id)}`).get();
    return true;
  } catch (e: unknown) {
    // Handle both status codes and error codes from Microsoft Graph
    if (e && typeof e === 'object') {
      const status = 'status' in e ? (e as ErrorWithStatus).status : undefined;
      const statusCode = 'statusCode' in e ? (e as ErrorWithStatus).statusCode : undefined;
      const code = 'code' in e ? (e as ErrorWithStatus).code : undefined;
      if (status === 404 || statusCode === 404 || code === 404 || code === 'CategoryNotFound') {
        return false;
      }
    }
    throw e;
  }
}

/**
 * Wait for a category to be deleted (eventual consistency)
 */
export async function waitForCategoryDeleted(graph: Client, id: string, opts: { interval?: number; timeout?: number } = {}): Promise<void> {
  const initialInterval = typeof opts.interval === 'number' ? opts.interval : 100;
  const timeout = typeof opts.timeout === 'number' ? opts.timeout : 5000;
  const maxInterval = 1000;
  const start = Date.now();
  let currentInterval = initialInterval;

  while (true) {
    if (Date.now() - start > timeout) {
      throw new Error(`waitForCategoryDeleted: timeout waiting for category ${id} to be deleted`);
    }
    const exists = await categoryExists(graph, id);
    if (!exists) {
      return; // Successfully deleted
    }

    await delay(currentInterval);
    // Exponential backoff with cap
    currentInterval = Math.min(currentInterval * 1.5, maxInterval);
  }
}

/**
 * Batch delete multiple test categories with enhanced error reporting.
 * Returns close summary for test diagnostics.
 */
export async function batchDeleteTestCategories(graph: Client, ids: string[], logger: Logger): Promise<{ successful: number; failed: number; errors: Array<{ id: string; error: string }> }> {
  const startTime = Date.now();
  const results = { successful: 0, failed: 0, errors: [] as Array<{ id: string; error: string }> };

  logger.debug('Starting batch category close', { count: ids.length });

  // Sequential deletion required: Microsoft Graph API returns 409 conflicts
  // when deleting categories in parallel. See investigation report in
  // .agents/reports/batch-test-actual-investigation.md
  for (const id of ids) {
    try {
      await deleteTestCategory(graph, id, logger);
      results.successful++;
    } catch (e: unknown) {
      results.failed++;
      results.errors.push({
        id,
        error: e instanceof Error ? e.message : String(e),
      });
    }
  }

  const duration = Date.now() - startTime;
  logger.info('Batch category close completed', {
    total: ids.length,
    successful: results.successful,
    failed: results.failed,
    duration,
  });

  return results;
}
