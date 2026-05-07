export class HttpDownloadError extends Error {
  status?: number;
  retryAfterSeconds?: number;

  constructor(message: string, status?: number, retryAfterSeconds?: number) {
    super(message);
    this.name = 'HttpDownloadError';
    this.status = status;
    this.retryAfterSeconds = retryAfterSeconds;
  }
}

export function parseRetryAfter(value: string | null) {
  if (!value) return undefined;
  const seconds = Number(value);
  if (Number.isFinite(seconds)) return Math.max(seconds, 0);
  const retryTime = Date.parse(value);
  if (Number.isFinite(retryTime)) return Math.max((retryTime - Date.now()) / 1000, 0);
  return undefined;
}

export function isTransientStatus(status?: number) {
  return status === 408 || status === 429 || status === 500 || status === 502 || status === 503 || status === 504;
}

export function backoffDelayMs(attempt: number, retryAfterSeconds?: number) {
  if (retryAfterSeconds !== undefined) return retryAfterSeconds * 1000;
  const base = Math.min(30000, 1000 * 2 ** Math.max(attempt - 1, 0));
  return base + Math.floor(Math.random() * Math.min(base, 1000));
}

export function sleep(ms: number, signal?: AbortSignal) {
  return new Promise<void>((resolve, reject) => {
    if (signal?.aborted) {
      reject(new DOMException('Operation cancelled', 'AbortError'));
      return;
    }
    const id = window.setTimeout(resolve, ms);
    signal?.addEventListener('abort', () => {
      window.clearTimeout(id);
      reject(new DOMException('Operation cancelled', 'AbortError'));
    }, { once: true });
  });
}
