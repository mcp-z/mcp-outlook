import assert from 'assert';
import spawn from 'cross-spawn-cb';
import * as fs from 'fs';
import { safeRmSync } from 'fs-remove-compat';
import getPort from 'get-port';
import * as path from 'path';
import { pathToFileURL } from 'url';

// Real server processes: stdin is closed so a server that starts when it should refuse fails the spec instead of hanging it.
describe('DCR settings (outlook)', () => {
  const projectDir = path.resolve('.tmp', `dcr-settings-${process.pid}`);
  const serverPath = path.resolve('bin/server.js');
  const baseEnv = () => ({ PATH: process.env.PATH ?? '', MS_CLIENT_ID: 'test-client-id', MS_TENANT_ID: 'common', LOG_LEVEL: 'error' });

  before(() => {
    fs.mkdirSync(projectDir, { recursive: true });
    fs.writeFileSync(path.join(projectDir, '.mcp.json'), '{ "mcpServers": {} }\n');
  });
  after(() => safeRmSync(projectDir, { recursive: true, force: true }));

  // Resolves with stderr when the server stops with a non-zero exit; fails when it exits cleanly.
  async function refused(args: string[], env: NodeJS.ProcessEnv): Promise<string> {
    try {
      await spawn(process.execPath, [serverPath, ...args], { cwd: projectDir, env: { ...baseEnv(), ...env }, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });
    } catch (error) {
      return String((error as { stderr?: unknown }).stderr ?? '');
    }
    throw new Error('the server exited cleanly but should have refused to start');
  }

  // stdio on purpose: the refusal is at config parse, and if it regressed the process exits on closed stdin instead of hanging.
  const dcr = ['--auth=dcr', '--base-url=http://127.0.0.1:1'];
  const message = 'external DCR mode is not supported';

  function assertOneError(stderr: string, text: string) {
    assert.strictEqual(stderr.split(text).length - 1, 1, `stderr should name the problem once: ${stderr}`);
  }

  it('refuses --dcr-mode=external', async () => {
    assertOneError(await refused([...dcr, '--dcr-mode=external', '--dcr-verify-url=http://127.0.0.1:9/verify'], {}), message);
  });

  it('refuses DCR_MODE=external', async () => {
    assertOneError(await refused(dcr, { DCR_MODE: 'external', DCR_VERIFY_URL: 'http://127.0.0.1:9/verify' }), message);
  });

  it('refuses a DCR verify URL', async () => {
    const stderr = await refused(dcr, { DCR_VERIFY_URL: 'http://127.0.0.1:9/verify' });
    assert.ok(stderr.includes('DCR_VERIFY_URL'), stderr);
    assert.ok(stderr.includes('is not supported'), stderr);
  });

  it('writes registrations to the file --dcr-store-uri names', async () => {
    const port = await getPort({ host: '127.0.0.1' });
    const stateDir = path.join(projectDir, 'flag-state');
    const flagFile = path.join(stateDir, 'flag-dcr.json');
    const envFile = path.join(stateDir, 'env-dcr.json');
    fs.mkdirSync(stateDir, { recursive: true });
    const base = `http://127.0.0.1:${port}`;
    const args = [serverPath, '--auth=dcr', `--port=${port}`, `--base-url=${base}`, `--dcr-store-uri=${pathToFileURL(flagFile).href.replace(/^file:/, 'file+shared:')}`, '--dcr-mode=self-hosted'];
    const env = { ...baseEnv(), DCR_STORE_URI: pathToFileURL(envFile).href.replace(/^file:/, 'file+shared:'), TOKEN_STORE_URI: pathToFileURL(path.join(stateDir, 'tokens.json')).href };
    const child = spawn.crossSpawn(process.execPath, args, { cwd: projectDir, env, stdio: ['ignore', 'ignore', 'pipe'] });
    let stderr = '';
    child.stderr?.on('data', (chunk) => {
      stderr += String(chunk);
    });
    const exited = new Promise<number | null>((resolve) => child.once('exit', resolve));
    try {
      let status = 0;
      while (status !== 201) {
        if (child.exitCode !== null) throw new Error(`the server exited before registering: ${stderr}`);
        try {
          const response = await fetch(`${base}/oauth/register`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ client_name: 'dcr-settings', redirect_uris: ['http://127.0.0.1:53682/callback'] }) });
          status = response.status;
        } catch {
          // not listening yet
        }
      }
      // file+shared writes the file before the registration is answered.
      assert.ok(fs.existsSync(flagFile) && fs.readFileSync(flagFile, 'utf8').includes('dcr-settings'), `no registration in ${flagFile}`);
      assert.ok(!fs.existsSync(envFile) || !fs.readFileSync(envFile, 'utf8').includes('dcr-settings'), 'the environment file must not receive the registration');
    } finally {
      child.kill();
      await exited;
    }
  });

  // The refusal is for --auth=dcr only: elsewhere the settings never had an effect, so stale ones must not stop a working install.
  describe('stale external DCR settings outside --auth=dcr', () => {
    const stale = { DCR_MODE: 'external', DCR_VERIFY_URL: 'http://127.0.0.1:9/verify' };

    // stdio: the settings are read at config parse, so this proves the start without needing a port.
    async function startsOnStdio(args: string[], tokenFile: string) {
      const env = { ...baseEnv(), ...stale, TOKEN_STORE_URI: pathToFileURL(path.join(projectDir, tokenFile)).href };
      const child = spawn.crossSpawn(process.execPath, [serverPath, '--stdio', ...args], { cwd: projectDir, env, stdio: ['pipe', 'pipe', 'pipe'] });
      let stderr = '';
      child.stderr?.on('data', (chunk) => {
        stderr += String(chunk);
      });
      const exited = new Promise<number | null>((resolve) => child.once('exit', resolve));
      const answered = new Promise<string>((resolve, reject) => {
        let out = '';
        child.stdout?.on('data', (chunk) => {
          out += String(chunk);
          if (out.includes('\n')) resolve(out);
        });
        child.once('exit', () => reject(new Error(`the server stopped instead of starting: ${stderr}`)));
      });
      try {
        child.stdin?.write(`${JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'dcr-settings', version: '0.0.0-test' } } })}\n`);
        assert.ok((await answered).includes('"result"'));
      } finally {
        child.kill();
        await exited;
      }
    }

    it('starts in loopback-oauth mode', async () => {
      await startsOnStdio(['--auth=loopback-oauth', '--headless'], 'loopback-tokens.json');
    });

    it('starts in device-code mode', async () => {
      await startsOnStdio(['--auth=device-code', '--headless'], 'device-code-tokens.json');
    });

    it('still refuses with --auth=dcr', async () => {
      assertOneError(await refused(dcr, stale), message);
    });
  });
});
