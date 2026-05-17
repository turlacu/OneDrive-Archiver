export const quickXorWidthInBits = 160;
export const quickXorShift = 11;

export function bytesToBase64(bytes: Uint8Array) {
  if (typeof btoa === 'function') {
    let binary = '';
    for (const byte of bytes) binary += String.fromCharCode(byte);
    return btoa(binary);
  }
  return Buffer.from(bytes).toString('base64');
}

export function base64ToBytes(value: string) {
  if (typeof atob === 'function') {
    const binary = atob(value);
    return Uint8Array.from(binary, char => char.charCodeAt(0));
  }
  return Uint8Array.from(Buffer.from(value, 'base64'));
}

export function hashMatches(actual: string, expected: string) {
  return actual.trim().toLowerCase() === expected.trim().toLowerCase();
}

export function isValidQuickXorHash(value?: string) {
  if (!value) return false;
  try {
    return base64ToBytes(value).length === quickXorWidthInBits / 8;
  } catch {
    return false;
  }
}
