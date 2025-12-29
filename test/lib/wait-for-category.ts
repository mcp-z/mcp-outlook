import { setTimeout as delay } from 'timers/promises';

export default async function waitForCategory(client: { api: (path: string) => { get: () => Promise<{ id?: string }> } }, id: string, opts: { interval?: number; timeout?: number } = {}): Promise<unknown> {
  const initialInterval = typeof opts.interval === 'number' ? opts.interval : 100;
  const timeout = typeof opts.timeout === 'number' ? opts.timeout : 5000;
  const maxInterval = 1000;
  const start = Date.now();
  let currentInterval = initialInterval;

  while (true) {
    if (Date.now() - start > timeout) throw new Error('waitForCategory: timeout waiting for category');
    try {
      const resp = await client.api(`/me/outlook/masterCategories/${encodeURIComponent(id)}`).get();
      if (resp && typeof (resp as Record<string, unknown>).id === 'string') return resp;
    } catch {
      // ignore and retry
    }
    await delay(currentInterval);
    // Exponential backoff with cap
    currentInterval = Math.min(currentInterval * 1.5, maxInterval);
  }
}
