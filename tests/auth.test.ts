import assert from 'node:assert/strict';
import test from 'node:test';
import { validateOAuthState } from '../server.ts';

test('accepts matching OAuth state values', () => {
  assert.equal(validateOAuthState('state-123', 'state-123'), true);
});

test('rejects missing or mismatched OAuth state values', () => {
  assert.equal(validateOAuthState(undefined, 'state-123'), false);
  assert.equal(validateOAuthState('state-123', undefined), false);
  assert.equal(validateOAuthState('state-123', 'state-456'), false);
});
