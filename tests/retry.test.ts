import assert from 'node:assert/strict';
import test from 'node:test';
import { backoffDelayMs, isTransientStatus, parseRetryAfter } from '../src/download/retry.ts';

test('parses Retry-After seconds', () => {
  assert.equal(parseRetryAfter('7'), 7);
});

test('parses Retry-After date', () => {
  const future = new Date(Date.now() + 5000).toUTCString();
  const parsed = parseRetryAfter(future);
  assert.ok(parsed !== undefined && parsed > 0 && parsed <= 6);
});

test('detects transient HTTP statuses', () => {
  assert.equal(isTransientStatus(429), true);
  assert.equal(isTransientStatus(503), true);
  assert.equal(isTransientStatus(404), false);
});

test('uses Retry-After before exponential backoff', () => {
  assert.equal(backoffDelayMs(3, 2), 2000);
  assert.ok(backoffDelayMs(2) >= 2000);
});
