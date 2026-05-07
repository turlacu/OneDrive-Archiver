const widthInBits = 160;
const shift = 11;

function base64ToBytes(value: string) {
  const binary = atob(value);
  return Uint8Array.from(binary, char => char.charCodeAt(0));
}

function bytesToBase64(bytes: Uint8Array) {
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
}

export async function quickXorHash(file: File | Blob) {
  const hash = new Uint8Array(widthInBits / 8);
  const reader = file.stream().getReader();
  let length = 0;

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;

    for (let i = 0; i < value.length; i += 1) {
      const byte = value[i];
      const bitOffset = ((length + i) * shift) % widthInBits;
      const byteOffset = Math.floor(bitOffset / 8);
      const bitRemainder = bitOffset % 8;

      hash[byteOffset] ^= byte << bitRemainder;
      if (bitRemainder > 0) {
        hash[(byteOffset + 1) % hash.length] ^= byte >> (8 - bitRemainder);
      }
    }

    length += value.length;
  }

  const lengthBytes = new Uint8Array(8);
  const view = new DataView(lengthBytes.buffer);
  view.setBigUint64(0, BigInt(length), true);
  for (let i = 0; i < lengthBytes.length; i += 1) {
    hash[hash.length - lengthBytes.length + i] ^= lengthBytes[i];
  }

  return bytesToBase64(hash);
}

export async function sha1Hex(file: File | Blob) {
  const digest = await crypto.subtle.digest('SHA-1', await file.arrayBuffer());
  return Array.from(new Uint8Array(digest)).map(byte => byte.toString(16).padStart(2, '0')).join('').toUpperCase();
}

export function hashMatches(actual: string, expected: string) {
  return actual.trim().toLowerCase() === expected.trim().toLowerCase();
}

export function isValidQuickXorHash(value?: string) {
  if (!value) return false;
  try {
    return base64ToBytes(value).length === 20;
  } catch {
    return false;
  }
}
