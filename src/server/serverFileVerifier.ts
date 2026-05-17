import { createHash } from 'node:crypto';
import { createReadStream } from 'node:fs';
import fs from 'node:fs/promises';
import { pipeline } from 'node:stream/promises';
import { bytesToBase64, hashMatches, isValidQuickXorHash, quickXorShift, quickXorWidthInBits } from '../download/hashTools';
import type { RemoteItemMetadata } from '../download/types';

export interface ServerVerificationResult {
  ok: boolean;
  cryptographic: boolean;
  message: string;
}

export async function verifyServerFile(filePath: string, item: RemoteItemMetadata): Promise<ServerVerificationResult> {
  const stat = await fs.stat(filePath);
  if (stat.size !== item.size) {
    return {
      ok: false,
      cryptographic: false,
      message: `Size mismatch for ${item.remotePath}: expected ${item.size} bytes, wrote ${stat.size} bytes.`,
    };
  }

  if (item.hashes.sha1Hash) {
    const actual = await sha1File(filePath);
    const ok = hashMatches(actual, item.hashes.sha1Hash);
    return {
      ok,
      cryptographic: true,
      message: ok
        ? `SHA-1 verification passed: ${item.remotePath}`
        : `SHA-1 verification failed for ${item.remotePath}.`,
    };
  }

  if (isValidQuickXorHash(item.hashes.quickXorHash)) {
    const actual = await quickXorHashFile(filePath);
    const ok = hashMatches(actual, item.hashes.quickXorHash || '');
    return {
      ok,
      cryptographic: true,
      message: ok
        ? `QuickXorHash verification passed: ${item.remotePath}`
        : `QuickXorHash verification failed for ${item.remotePath}.`,
    };
  }

  return {
    ok: true,
    cryptographic: false,
    message: `No OneDrive hash was available; verified by size for ${item.remotePath}.`,
  };
}

export async function sha1File(filePath: string) {
  const hash = createHash('sha1');
  await pipeline(createReadStream(filePath), hash);
  return hash.digest('hex').toUpperCase();
}

export async function quickXorHashFile(filePath: string) {
  const hash = new Uint8Array(quickXorWidthInBits / 8);
  let length = 0;
  for await (const chunk of createReadStream(filePath)) {
    const bytes = chunk as Buffer;
    for (let i = 0; i < bytes.length; i += 1) {
      const bitOffset = ((length + i) * quickXorShift) % quickXorWidthInBits;
      const byteOffset = Math.floor(bitOffset / 8);
      const bitRemainder = bitOffset % 8;
      hash[byteOffset] ^= bytes[i] << bitRemainder;
      if (bitRemainder > 0) {
        hash[(byteOffset + 1) % hash.length] ^= bytes[i] >> (8 - bitRemainder);
      }
    }
    length += bytes.length;
  }

  const lengthBytes = new Uint8Array(8);
  new DataView(lengthBytes.buffer).setBigUint64(0, BigInt(length), true);
  for (let i = 0; i < lengthBytes.length; i += 1) {
    hash[hash.length - lengthBytes.length + i] ^= lengthBytes[i];
  }
  return bytesToBase64(hash);
}
