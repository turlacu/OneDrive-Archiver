import assert from 'node:assert/strict';
import test from 'node:test';
import { isTransientOfficeLockFile, normalizePathSegment, normalizeRelativePath, shouldIncludeFile } from '../src/download/pathTools.ts';

test('normalizes invalid path characters and reserved names', () => {
  assert.equal(normalizePathSegment('a<b>c:d"e/f\\g|h?i*.txt'), 'a_b_c_d_e_f_g_h_i_.txt');
  assert.equal(normalizePathSegment('CON'), '_CON');
  assert.equal(normalizePathSegment('report. '), 'report');
});

test('normalizes a nested relative path', () => {
  assert.equal(normalizeRelativePath(['Folder ', 'aux.txt', 'file:name.docx']), 'Folder/_aux.txt/file_name.docx');
});

test('detects transient Office lock files', () => {
  assert.equal(isTransientOfficeLockFile('~$document.docx'), true);
  assert.equal(isTransientOfficeLockFile('document.docx'), false);
});

test('applies include and exclude extension filters', () => {
  assert.equal(shouldIncludeFile('a.docx', [], ['tmp']), true);
  assert.equal(shouldIncludeFile('a.tmp', [], ['tmp']), false);
  assert.equal(shouldIncludeFile('a.pdf', ['pdf'], []), true);
  assert.equal(shouldIncludeFile('a.docx', ['pdf'], []), false);
});
