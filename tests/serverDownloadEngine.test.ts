import assert from 'node:assert/strict';
import test from 'node:test';
import path from 'node:path';
import { relativePathPartsFromGraphPath, resolveInsideRoot, sanitizeUserFolder } from '../src/server/serverDownloadEngine.ts';

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

test('sanitizes user emails for isolated server folders', () => {
  assert.equal(sanitizeUserFolder('Owner@Example.COM'), 'owner_example.com');
  assert.equal(sanitizeUserFolder('bad/user@example.com'), 'bad_user_example.com');
  assert.equal(sanitizeUserFolder('***'), 'user');
});

test('extracts OneDrive relative paths for nested selections', () => {
  assert.deepEqual(
    relativePathPartsFromGraphPath('/drive/root:/Imagini/Emergency'),
    ['Imagini', 'Emergency'],
  );
  assert.deepEqual(
    relativePathPartsFromGraphPath('/drive/root:/Imagini/Emergency/photo.jpg', 'photo.jpg'),
    ['Imagini', 'Emergency'],
  );
});
