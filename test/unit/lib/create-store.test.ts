import { getToken } from '@mcp-z/oauth';
import assert from 'assert';
import * as fs from 'fs';
import { safeRmSync } from 'fs-remove-compat';
import { KeyvFile } from 'keyv-file';
import * as path from 'path';
import { pathToFileURL } from 'url';
import createStore from '../../../src/lib/create-store.ts';
import { createDcrStore, createTokenStore } from '../../../src/setup/runtime.ts';

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

describe('create-store default stores (outlook)', () => {
  const root = path.resolve('.tmp', `create-store-default-${process.pid}`);
  const logger = { debug: () => {}, info: () => {}, warn: () => {}, error: () => {} };
  const saved = { token: process.env.TOKEN_STORE_URI, dcr: process.env.DCR_STORE_URI };

  before(() => {
    delete process.env.TOKEN_STORE_URI;
    delete process.env.DCR_STORE_URI;
  });
  after(() => {
    if (saved.token !== undefined) process.env.TOKEN_STORE_URI = saved.token;
    if (saved.dcr !== undefined) process.env.DCR_STORE_URI = saved.dcr;
    safeRmSync(root, { recursive: true, force: true });
  });

  function scratch(name: string): string {
    const folder = path.join(root, name);
    fs.mkdirSync(folder, { recursive: true });
    return folder;
  }
  const accessToken = async (store: { get(key: string): Promise<unknown> }, key: string) => ((await store.get(key)) as { accessToken?: string } | undefined)?.accessToken;

  it('two token stores opened from the default URI keep each other sign-ins and see each other changes', async () => {
    const folder = scratch('shared');
    const first = await createTokenStore(folder, logger);
    const second = await createTokenStore(folder, logger);
    try {
      await first.set('alice@example.com:outlook:token', { accessToken: 'alice' });
      await second.set('bob@example.com:outlook:token', { accessToken: 'bob' });
      assert.strictEqual(await accessToken(first, 'bob@example.com:outlook:token'), 'bob');
      assert.strictEqual(await accessToken(second, 'alice@example.com:outlook:token'), 'alice');
      await first.delete('alice@example.com:outlook:token');
      assert.strictEqual(await second.get('alice@example.com:outlook:token'), undefined);
      assert.strictEqual(await accessToken(second, 'bob@example.com:outlook:token'), 'bob');
    } finally {
      await first.disconnect();
      await second.disconnect();
    }
  });

  it('two DCR stores opened from the default URI keep each other registrations', async () => {
    const folder = scratch('shared-dcr');
    const first = await createDcrStore(folder, true, logger);
    const second = await createDcrStore(folder, true, logger);
    assert.ok(first && second);
    try {
      await first.set('dcr:client:a', { accessToken: 'a' });
      await second.set('dcr:client:b', { accessToken: 'b' });
      assert.strictEqual(await accessToken(first, 'dcr:client:b'), 'b');
      assert.strictEqual(await accessToken(second, 'dcr:client:a'), 'a');
    } finally {
      await first.disconnect();
      await second.disconnect();
    }
  });

  it('reads a file written through an explicit file:// URI', async () => {
    const folder = scratch('written-by-keyv-file');
    const file = path.join(folder, 'tokens.json');
    const writer = await createStore<unknown>(pathToFileURL(file).href, logger);
    await writer.set('alice@example.com:outlook:token', { accessToken: 'from-keyv-file' });
    await writer.disconnect();
    const store = await createTokenStore(folder, logger);
    try {
      assert.strictEqual(await accessToken(store, 'alice@example.com:outlook:token'), 'from-keyv-file');
    } finally {
      await store.disconnect();
    }
  });

  it('rejects naming the file when the default file is damaged', async () => {
    const folder = scratch('damaged');
    const file = path.join(folder, 'tokens.json');
    const damaged = '{"cache":[["a"';
    fs.writeFileSync(file, damaged);
    await assert.rejects(createTokenStore(folder, logger), (err: Error) => err.message.includes(file));
    assert.strictEqual(fs.readFileSync(file, 'utf8'), damaged);
  });
});
