import type { Client } from '@microsoft/microsoft-graph-client';
import type { RemoteHashes, RemoteItemMetadata } from './types';

export interface GraphDriveItem {
  id: string;
  name: string;
  size?: number;
  eTag?: string;
  cTag?: string;
  lastModifiedDateTime?: string;
  folder?: unknown;
  file?: {
    hashes?: RemoteHashes;
  };
  package?: unknown;
  remoteItem?: GraphDriveItem & {
    parentReference?: {
      driveId?: string;
      path?: string;
    };
  };
  parentReference?: {
    driveId?: string;
    path?: string;
  };
  '@microsoft.graph.downloadUrl'?: string;
  deleted?: unknown;
}

const selectFields = 'id,name,size,eTag,cTag,lastModifiedDateTime,folder,file,package,remoteItem,parentReference';

export class OneDriveClient {
  constructor(
    private readonly graphClient: Client,
    private readonly refreshAccessToken?: () => Promise<string | null>,
  ) {}

  async getMe() {
    return this.withAuthRetry(() => this.graphClient.api('/me').get());
  }

  async listChildren(folderId: string) {
    const endpoint = folderId === 'root'
      ? '/me/drive/root/children'
      : `/me/drive/items/${folderId}/children`;
    return this.readPagedItems(endpoint);
  }

  async listChildrenPaged(
    folderId: string,
    onPage: (items: GraphDriveItem[]) => void,
  ) {
    const endpoint = folderId === 'root'
      ? '/me/drive/root/children'
      : `/me/drive/items/${folderId}/children`;
    return this.readPagedItems(endpoint, onPage);
  }

  async getItem(itemId: string) {
    const response = await this.withAuthRetry(() => this.graphClient.api(`/me/drive/items/${itemId}`).get());
    return this.normalizeItem(response);
  }

  async refreshDownloadUrl(itemId: string) {
    const metadata = await this.getItem(itemId);
    if (metadata.downloadUrl) return metadata;

    const accessToken = await this.refreshAccessToken?.();
    if (!accessToken) return metadata;
    const response = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${itemId}/content`, {
      headers: { Authorization: `Bearer ${accessToken}` },
      redirect: 'manual',
    });
    const location = response.headers.get('Location');
    if (location) {
      return { ...metadata, downloadUrl: location };
    }

    return metadata;
  }

  async delta(driveId: string, token?: string) {
    const endpoint = token || `/drives/${driveId}/root/delta`;
    let response = await this.withAuthRetry(() => this.graphClient.api(endpoint).select(selectFields).top(200).get());
    const items: RemoteItemMetadata[] = [];
    let deltaToken: string | undefined;

    while (response) {
      for (const item of response.value || []) {
        if (!item.deleted && !item.folder) items.push(this.normalizeItem(item));
      }
      deltaToken = response['@odata.deltaLink'];
      const nextLink = response['@odata.nextLink'];
      response = nextLink ? await this.withAuthRetry(() => this.graphClient.api(nextLink).get()) : null;
    }

    return { items, deltaToken };
  }

  async fetchRange(downloadUrl: string, start: number, end: number, signal?: AbortSignal) {
    const response = await fetch(downloadUrl, {
      headers: { Range: `bytes=${start}-${end}` },
      signal,
    });

    return response;
  }

  private async readPagedItems(endpoint: string, onPage?: (items: GraphDriveItem[]) => void) {
    let response = await this.withAuthRetry(() => this.graphClient.api(endpoint).select(selectFields).top(200).get());
    const items: GraphDriveItem[] = [];
    while (response) {
      const page = response.value || [];
      items.push(...page);
      onPage?.(page);
      const nextLink = response['@odata.nextLink'];
      response = nextLink ? await this.withAuthRetry(() => this.graphClient.api(nextLink).get()) : null;
    }
    return items;
  }

  private async withAuthRetry<T>(operation: () => Promise<T>): Promise<T> {
    try {
      return await operation();
    } catch (error) {
      if (!this.isAuthError(error) || !this.refreshAccessToken) {
        throw error;
      }
      const token = await this.refreshAccessToken();
      if (!token) throw error;
      return operation();
    }
  }

  private isAuthError(error: unknown) {
    const candidate = error as { statusCode?: number; code?: string; body?: { message?: string }; message?: string };
    const message = candidate.body?.message || candidate.message || '';
    return candidate.statusCode === 401
      || candidate.statusCode === 403
      || message.includes('JWT is not well formed')
      || message.includes('InvalidAuthenticationToken')
      || message.includes('Access token has expired');
  }

  normalizeItem(item: GraphDriveItem): RemoteItemMetadata {
    const remoteItem = item.remoteItem;
    const source = remoteItem || item;
    const driveId = source.parentReference?.driveId || item.parentReference?.driveId || 'me';
    const path = source.parentReference?.path || item.parentReference?.path || '';
    const remotePath = `${path}/${source.name || item.name}`.replace(/^\/+/, '');

    return {
      driveId,
      itemId: source.id || item.id,
      name: source.name || item.name,
      remotePath,
      size: source.size || item.size || 0,
      eTag: source.eTag || item.eTag,
      cTag: source.cTag || item.cTag,
      lastModifiedDateTime: source.lastModifiedDateTime || item.lastModifiedDateTime,
      hashes: source.file?.hashes || item.file?.hashes || {},
      downloadUrl: source['@microsoft.graph.downloadUrl'] || item['@microsoft.graph.downloadUrl'],
    };
  }
}
