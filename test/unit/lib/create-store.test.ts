import { getToken } from '@mcp-z/oauth';
import assert from 'assert';
import * as fs from 'fs';
import { KeyvFile } from 'keyv-file';
import * as path from 'path';
import { pathToFileURL } from 'url';
import createStore from '../../../src/lib/create-store.ts';

const fixture = path.join(process.cwd(), 'test', 'fixtures', 'keyv5-tokens.json');

describe('createStore keyv 5 migration (outlook)', () => {
  let dir: string;
  const messages: string[] = [];
  const logger = {
    debug: () => {},
    info: (message: string) => messages.push(message),
    warn: () => {},
    error: () => {},
  };

  beforeEach(() => {
    messages.length = 0;
    fs.mkdirSync(path.join(process.cwd(), '.tmp'), { recursive: true });
    dir = fs.mkdtempSync(path.join(process.cwd(), '.tmp', 'create-store-'));
    fs.copyFileSync(fixture, path.join(dir, 'tokens.json'));
  });

  afterEach(() => {
    fs.rmSync(dir, { recursive: true, force: true });
  });

  it('finds a token saved by keyv 5 and logs the count without the URI', async () => {
    const uri = pathToFileURL(path.join(dir, 'tokens.json')).href;
    const store = await createStore<unknown>(uri, logger);

    const token = await getToken<{ accessToken: string }>(store, { accountId: 'user@example.com', service: 'outlook' });
    assert.strictEqual(token?.accessToken, 'dummy-access');

    assert.strictEqual(messages.length, 1);
    assert.strictEqual(messages[0], 'Moved 4 saved sign-in entries to the keyv 6 format');
    assert.ok(!messages[0]?.includes(dir));
  });

  it('logs nothing when there is nothing to move', async () => {
    const uri = pathToFileURL(path.join(dir, 'tokens.json')).href;
    await createStore<unknown>(uri, logger);
    messages.length = 0;
    const store = await createStore<unknown>(uri, logger);

    assert.strictEqual(messages.length, 0);
    const token = await getToken<{ accessToken: string }>(store, { accountId: 'user@example.com', service: 'outlook' });
    assert.strictEqual(token?.accessToken, 'dummy-access');
  });

  it('rejects, so the server does not start, when a read fails during the move', async () => {
    const uri = pathToFileURL(path.join(dir, 'tokens.json')).href;
    const original = KeyvFile.prototype.get;
    let failed = false;
    KeyvFile.prototype.get = function (this: KeyvFile, ...args: Parameters<KeyvFile['get']>) {
      if (!failed) {
        failed = true;
        throw new Error('simulated read failure');
      }
      return original.apply(this, args);
    } as KeyvFile['get'];
    try {
      await assert.rejects(createStore<unknown>(uri, logger), /migrateKeyv5Keys/);
      assert.ok(failed);
    } finally {
      KeyvFile.prototype.get = original;
    }
  });
});
