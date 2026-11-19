import '../../../lib/env-loader.ts';
import { ProtocolError, ProtocolErrorCode } from '@mcp-z/server';
import assert from 'assert';
import { allFailedError } from '../../../../src/mcp/tools/label-delete.ts';

describe('label-delete allFailedError', () => {
  it('is an InternalError naming every failure when every delete failed', () => {
    const error = allFailedError([
      { id: 'a', success: false, error: 'Forbidden' },
      { id: 'b', success: false },
    ]);
    assert.ok(error instanceof ProtocolError);
    assert.strictEqual(error.code, ProtocolErrorCode.InternalError);
    assert.match(error.message, /Failed to delete all 2 categories/);
    assert.match(error.message, /a: Forbidden/);
    assert.match(error.message, /b: Unknown error/);
  });

  it('is null when only some deletes failed (partial failure keeps the success shape)', () => {
    assert.strictEqual(
      allFailedError([
        { id: 'a', success: true },
        { id: 'b', success: false, error: 'Forbidden' },
      ]),
      null
    );
  });

  it('is null when every delete succeeded', () => {
    assert.strictEqual(allFailedError([{ id: 'a', success: true }]), null);
  });
});
