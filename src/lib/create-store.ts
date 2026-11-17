import type { Logger } from '@mcp-z/server';
import type Keyv from 'keyv';
import keyvRegistry, { migrateKeyv5Keys } from 'keyv-registry';

// keyv-file is a direct dependency so the default file:// URIs resolve; keyv-registry
// loads adapters with require() rather than depending on any of them itself.

export default async function createStore<T>(uri: string, logger: Logger): Promise<Keyv<T>> {
  const store = await keyvRegistry<T>(uri);
  // Name only the protocol: the URI can hold a password.
  if (!store) throw new Error(`Failed to create store for protocol: ${uri.split(':')[0]}:`);
  // Saved sign-ins written by keyv 5 sit under old keys; move them once. The URI is never logged.
  try {
    const moved = await migrateKeyv5Keys(store as Keyv);
    if (moved > 0) logger.info(`Moved ${moved} saved sign-in entries to the keyv 6 format`);
  } catch (error) {
    // Keep the migration error as the one reported, even if closing the store fails too.
    await store.disconnect().catch(() => {});
    throw error;
  }
  return store;
}
