import assert from 'node:assert/strict';
import test from 'node:test';
import { OneDriveClient } from '../src/download/oneDriveClient.ts';

class GraphRequest {
  constructor(
    private readonly endpoint: string,
    private readonly calls: string[],
  ) {}

  select() {
    return this;
  }

  top() {
    return this;
  }

  async get() {
    this.calls.push(this.endpoint);
    return {
      value: [],
      '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/me/drive/root/delta?token=next',
    };
  }
}

class GraphClient {
  calls: string[] = [];

  api(endpoint: string) {
    return new GraphRequest(endpoint, this.calls);
  }
}

test('uses the me drive delta endpoint for personal OneDrive baseline scans', async () => {
  const graph = new GraphClient();
  const client = new OneDriveClient(graph as any);

  await client.delta('me');

  assert.equal(graph.calls[0], '/me/drive/root/delta');
});

test('uses a saved delta link directly when present', async () => {
  const graph = new GraphClient();
  const client = new OneDriveClient(graph as any);
  const token = 'https://graph.microsoft.com/v1.0/me/drive/root/delta?token=next';

  await client.delta('me', token);

  assert.equal(graph.calls[0], token);
});
