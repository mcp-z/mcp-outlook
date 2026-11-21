import assert from 'assert';
import { handleVersionHelp } from '../../../src/setup/version-help.ts';

describe('setup/version-help', () => {
  it('help names the private configuration file location', () => {
    const { handled, output } = handleVersionHelp(['--help']);
    assert.strictEqual(handled, true);
    assert.ok(output?.includes('~/.config/mcp-z/mcp-outlook/.env'));
  });

  it('help no longer lists the external DCR settings', () => {
    const { output } = handleVersionHelp(['--help']);
    for (const setting of ['--dcr-mode', '--dcr-verify-url', 'DCR_MODE', 'DCR_VERIFY_URL']) assert.ok(!output?.includes(setting), `help still lists ${setting}`);
  });
});
