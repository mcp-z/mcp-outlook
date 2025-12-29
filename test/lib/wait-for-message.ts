import { setTimeout as delay } from 'timers/promises';

// Type for error objects that may have status/code properties
type ErrorWithStatus = {
  status?: number;
  statusCode?: number;
  code?: number | string;
};

export default async function waitForMessage(client: { api: (path: string) => { get: () => Promise<{ id?: string }> } }, id: string, opts: { interval?: number; timeout?: number } = {}): Promise<unknown> {
  const initialInterval = typeof opts.interval === 'number' ? opts.interval : 100;
  const timeout = typeof opts.timeout === 'number' ? opts.timeout : 10000;
  const maxInterval = 1000;
  const start = Date.now();
  let currentInterval = initialInterval;

  while (true) {
    if (Date.now() - start > timeout) throw new Error('waitForMessage: timeout waiting for message');
    try {
      const resp = await client.api(`/me/messages/${id}`).get();
      if (resp && typeof (resp as Record<string, unknown>).id === 'string') return resp;
    } catch (e: unknown) {
      // Only retry on 404 (message not indexed yet)
      // Fail fast on auth errors, rate limits, or other issues
      const status = e && typeof e === 'object' && ('status' in e || 'statusCode' in e || 'code' in e) ? (e as ErrorWithStatus).status || (e as ErrorWithStatus).statusCode || (e as ErrorWithStatus).code : undefined;
      if (status !== 404 && status !== 'ENOTFOUND') {
        throw e; // Real error - don't hide it
      }
      // 404 is expected during indexing - continue retry loop
    }
    await delay(currentInterval);
    // Exponential backoff with cap
    currentInterval = Math.min(currentInterval * 1.5, maxInterval);
  }
}
