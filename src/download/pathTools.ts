import type { ConflictStrategy } from './types';

const windowsReservedNames = /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(\..*)?$/i;
const unsafeCharacters = /[<>:"/\\|?*\u0000-\u001F]/g;

export function isTransientOfficeLockFile(name: string) {
  return name.startsWith('~$');
}

export function normalizePathSegment(segment: string) {
  const normalized = segment
    .normalize('NFC')
    .replace(unsafeCharacters, '_')
    .replace(/[. ]+$/g, '')
    .trim();
  const fallback = normalized || 'unnamed';
  return windowsReservedNames.test(fallback) ? `_${fallback}` : fallback;
}

export function normalizeRelativePath(parts: string[]) {
  return parts.map(normalizePathSegment).filter(Boolean).join('/');
}

export function extensionOf(name: string) {
  const dot = name.lastIndexOf('.');
  return dot > -1 ? name.slice(dot + 1).toLowerCase() : '';
}

export function shouldIncludeFile(
  name: string,
  includeExtensions: string[],
  excludeExtensions: string[],
) {
  const ext = extensionOf(name);
  const includes = includeExtensions.map(value => value.toLowerCase().replace(/^\./, ''));
  const excludes = excludeExtensions.map(value => value.toLowerCase().replace(/^\./, ''));

  if (excludes.includes(ext)) return false;
  if (includes.length === 0) return true;
  return includes.includes(ext);
}

export async function getUniqueFileName(
  directory: FileSystemDirectoryHandle,
  requestedName: string,
  strategy: ConflictStrategy,
) {
  const normalized = normalizePathSegment(requestedName);
  if (strategy === 'overwrite') return normalized;

  try {
    await directory.getFileHandle(normalized);
  } catch {
    return normalized;
  }

  if (strategy === 'skip_existing' || strategy === 'keep_newest') {
    return null;
  }

  const dot = normalized.lastIndexOf('.');
  const base = dot > 0 ? normalized.slice(0, dot) : normalized;
  const ext = dot > 0 ? normalized.slice(dot) : '';
  let index = 1;

  while (index < 10000) {
    const candidate = `${base} (${index})${ext}`;
    try {
      await directory.getFileHandle(candidate);
    } catch {
      return candidate;
    }
    index += 1;
  }

  throw new Error(`Unable to allocate conflict-free filename for ${requestedName}`);
}
