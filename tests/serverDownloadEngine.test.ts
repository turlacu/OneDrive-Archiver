import assert from 'node:assert/strict';
import test from 'node:test';
import path from 'node:path';
import { resolveInsideRoot } from '../src/server/serverDownloadEngine.ts';

test('resolves server download paths inside the configured root', () => {
  const root = path.resolve('/tmp/syncpoint-root');
  assert.equal(
    resolveInsideRoot(root, 'folder/report.txt'),
    path.join(root, 'folder', 'report.txt'),
  );
});

test('rejects server download paths escaping the configured root', () => {
  const root = path.resolve('/tmp/syncpoint-root');
  assert.throws(() => resolveInsideRoot(root, '../escape.txt'), /escapes/);
  assert.throws(() => resolveInsideRoot(root, '/etc/passwd'), /escapes/);
});
