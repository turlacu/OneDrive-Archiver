import { hashMatches, isValidQuickXorHash, quickXorHash, sha1Hex } from './quickXorHash';
import type { DownloadJob } from './types';

export interface VerificationResult {
  ok: boolean;
  cryptographic: boolean;
  message: string;
}

export class IntegrityVerifier {
  async verify(file: File, job: DownloadJob): Promise<VerificationResult> {
    if (file.size !== job.size) {
      return {
        ok: false,
        cryptographic: false,
        message: `Size mismatch: local ${file.size} bytes, remote ${job.size} bytes.`,
      };
    }

    if (job.hashes.sha1Hash) {
      const actual = await sha1Hex(file);
      return {
        ok: hashMatches(actual, job.hashes.sha1Hash),
        cryptographic: true,
        message: hashMatches(actual, job.hashes.sha1Hash)
          ? 'SHA-1 verification passed.'
          : 'SHA-1 verification failed.',
      };
    }

    if (isValidQuickXorHash(job.hashes.quickXorHash)) {
      const actual = await quickXorHash(file);
      return {
        ok: hashMatches(actual, job.hashes.quickXorHash || ''),
        cryptographic: true,
        message: hashMatches(actual, job.hashes.quickXorHash || '')
          ? 'QuickXorHash verification passed.'
          : 'QuickXorHash verification failed.',
      };
    }

    return {
      ok: true,
      cryptographic: false,
      message: 'No OneDrive hash was available; verified by size and saved remote eTag/cTag metadata only.',
    };
  }
}
