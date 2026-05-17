import React, { useState, useEffect, useRef } from 'react';
import { 
  Cloud, 
  AlertCircle, 
  LogIn,
  LogOut,
  ChevronRight,
  HardDrive,
  Moon,
  RotateCcw,
  Search,
  Sun,
  X
} from 'lucide-react';
import { motion } from 'motion/react';
import axios from 'axios';
import { Client } from '@microsoft/microsoft-graph-client';
import { ActivityLog } from './components/ActivityLog';
import { ExplorerTable } from './components/ExplorerTable';
import { DownloadEngine, type SourceSelection } from './download/downloadEngine';
import { FileCommitter } from './download/fileCommitter';
import { DownloadJobStore } from './download/jobStore';
import { OneDriveClient } from './download/oneDriveClient';
import { defaultDownloadSettings, type DownloadProgressSnapshot, type DownloadSettings, type ReporterEvent } from './download/types';

// Types
interface SyncItem {
  id: string;
  name: string;
  size: number;
  path: string;
  type: 'file' | 'folder';
  status: 'pending' | 'scanning' | 'downloading' | 'completed' | 'error' | 'skipped' | 'verifying' | 'queued' | 'paused' | 'retrying' | 'throttled' | 'failed' | 'stale_remote_changed' | 'insufficient_space' | 'cancelled' | 'interrupted';
  progress: number;
  sha1Hash?: string;
  error?: string;
}

interface LocalHandle {
  id: string;
  name: string;
  handle: FileSystemDirectoryHandle;
}

interface FolderCrumb {
  id: string;
  name: string;
}

interface GraphDriveItem {
  id: string;
  name: string;
  size?: number;
  folder?: unknown;
  file?: {
    hashes?: {
      sha1Hash?: string;
    };
  };
  parentReference?: {
    path?: string;
  };
  '@microsoft.graph.downloadUrl'?: string;
}

interface ActiveFile {
  id: string;
  name: string;
  status: SyncItem['status'];
  progress: number;
}

interface LogEntry {
  id: string;
  time: string;
  stage: string;
  message: string;
  jobName?: string;
  jobPath?: string;
  error?: string;
}

interface FailedFile {
  id: string;
  name: string;
  path: string;
  reason: string;
}

interface ServerArchiveJob {
  id: string;
  mode: 'start' | 'dry-run' | 'repair' | 'incremental';
  status: 'queued' | 'scanning' | 'downloading' | 'completed' | 'failed' | 'cancelled' | 'interrupted';
  targetRoot: string;
  log: string[];
  snapshot: DownloadProgressSnapshot;
}

export default function App() {
  const [accessToken, setAccessToken] = useState<string | null>(null);
  const [user, setUser] = useState<any>(null);
  const [rootItems, setRootItems] = useState<SyncItem[]>([]);
  const [currentFolder, setCurrentFolder] = useState<FolderCrumb>({ id: 'root', name: 'OneDrive' });
  const [breadcrumbs, setBreadcrumbs] = useState<FolderCrumb[]>([{ id: 'root', name: 'OneDrive' }]);
  const [isLoadingItems, setIsLoadingItems] = useState(false);
  const [itemsError, setItemsError] = useState<string | null>(null);
  const [selection, setSelection] = useState<Set<string>>(new Set());
  const [localDir, setLocalDir] = useState<LocalHandle | null>(null);
  const [isSyncing, setIsSyncing] = useState(false);
  const [syncStats, setSyncStats] = useState({ total: 0, completed: 0, errors: 0, size: 0, downloadedSize: 0 });
  const [downloadSpeed, setDownloadSpeed] = useState(0);
  const [downloadSettings, setDownloadSettings] = useState<DownloadSettings>(defaultDownloadSettings);
  const [progressSnapshot, setProgressSnapshot] = useState<DownloadProgressSnapshot | null>(null);
  const [activeFiles, setActiveFiles] = useState<ActiveFile[]>([]);
  const [logs, setLogs] = useState<LogEntry[]>([]);
  const [failedFiles, setFailedFiles] = useState<FailedFile[]>([]);
  const [mainTab, setMainTab] = useState<'explorer' | 'activity'>('explorer');
  const [resetLogsOnStart, setResetLogsOnStart] = useState(true);
  const [theme, setTheme] = useState<'light' | 'dark'>(() => localStorage.getItem('theme') === 'dark' ? 'dark' : 'light');
  const [tableScrollTop, setTableScrollTop] = useState(0);
  const [searchQuery, setSearchQuery] = useState('');
  const [activeSearchQuery, setActiveSearchQuery] = useState('');
  const [targetMode, setTargetMode] = useState<'browser' | 'server'>('browser');
  const [serverArchiveEnabled, setServerArchiveEnabled] = useState(false);
  const [serverTargetRoot, setServerTargetRoot] = useState<string | null>(null);
  const [serverJobs, setServerJobs] = useState<ServerArchiveJob[]>([]);
  const [activeServerJobId, setActiveServerJobId] = useState<string | null>(null);

  const oneDriveClient = useRef<OneDriveClient | null>(null);
  const downloadEngine = useRef<DownloadEngine | null>(null);
  const accessTokenRef = useRef<string | null>(null);
  const folderLoadRequestRef = useRef(0);

  useEffect(() => {
    document.documentElement.classList.toggle('dark', theme === 'dark');
    localStorage.setItem('theme', theme);
  }, [theme]);

  const refreshServerConfig = async () => {
    try {
      const res = await axios.get('/api/server-jobs/config');
      const enabled = Boolean(res.data?.enabled);
      setServerArchiveEnabled(enabled);
      setServerTargetRoot(enabled ? (res.data?.targetRoot || null) : null);
      if (enabled) setTargetMode('server');
      else setTargetMode('browser');
    } catch {
      setServerArchiveEnabled(false);
      setServerTargetRoot(null);
      setTargetMode('browser');
    }
  };

  useEffect(() => {
    refreshServerConfig();
  }, []);

  function isUsableAccessToken(value: string | null | undefined) {
    if (!value) return false;
    const token = value.trim();
    return token.length > 20
      && !token.includes(' ')
      && token !== 'undefined'
      && token !== 'null'
      && !token.startsWith('M.');
  }

  const clearStoredToken = () => {
    // Remove legacy readable tokens from older app versions. New auth keeps tokens in memory only.
    localStorage.removeItem('ms_token');
    localStorage.removeItem('ms_token_signal');
  };

  const clearAuthState = async () => {
    clearStoredToken();
    try {
      await axios.post('/api/auth/logout');
    } catch {
      // Server session cleanup is best-effort; local token cleanup is enough to stop client auth.
    }
  };

  const refreshAccessToken = async () => {
    try {
      const res = await axios.get('/api/auth/token');
      const token = res.data?.accessToken;
      if (!isUsableAccessToken(token)) {
        await clearAuthState();
        setAccessToken(null);
        addLog('Microsoft session expired. Please sign in again.');
        return null;
      }
      accessTokenRef.current = token;
      setAccessToken(token);
      if (res.data?.user) setUser(res.data.user);
      await refreshServerConfig();
      addLog('Microsoft access token refreshed.');
      return token;
    } catch (error: any) {
      await clearAuthState();
      setAccessToken(null);
      addLog(error.response?.data?.error || 'Microsoft session expired. Please sign in again.');
      return null;
    }
  };

  const completeAuthFromSession = async () => {
    try {
      const res = await axios.get('/api/auth/status');
      const token = res.data?.authenticated ? res.data?.accessToken : null;
      if (!isUsableAccessToken(token)) return false;
      accessTokenRef.current = token;
      setAccessToken(token);
      if (res.data?.user) setUser(res.data.user);
      await refreshServerConfig();
      addLog('Authentication successful');
      return true;
    } catch {
      return false;
    }
  };

  const mapDriveItem = (item: GraphDriveItem): SyncItem => ({
    id: item.id,
    name: item.name,
    size: item.size || 0,
    path: `${item.parentReference?.path || ''}/${item.name}`,
    type: item.folder ? 'folder' : 'file',
    status: 'pending',
    progress: 0,
    sha1Hash: item.file?.hashes?.sha1Hash
  });

  const sortSyncItems = (items: SyncItem[]) => {
    return [...items].sort((a, b) => {
      if (a.type !== b.type) return a.type === 'folder' ? -1 : 1;
      return a.name.localeCompare(b.name);
    });
  };

  const formatBytes = (bytes: number) => {
    if (bytes < 1024) return `${bytes} B`;
    const units = ['KB', 'MB', 'GB', 'TB'];
    let value = bytes / 1024;
    let unitIndex = 0;
    while (value >= 1024 && unitIndex < units.length - 1) {
      value /= 1024;
      unitIndex += 1;
    }
    return `${value.toFixed(value >= 10 ? 1 : 2)} ${units[unitIndex]}`;
  };

  const formatSpeed = (bytesPerSecond: number) => {
    if (bytesPerSecond <= 0) return '0.00KB/s';
    const units = ['B/s', 'KB/s', 'MB/s', 'GB/s', 'TB/s'];
    let value = bytesPerSecond;
    let unitIndex = 0;
    while (value >= 1024 && unitIndex < units.length - 1) {
      value /= 1024;
      unitIndex += 1;
    }
    return `${value.toFixed(2)}${units[unitIndex]}`;
  };

  const formatDuration = (seconds: number | undefined) => {
    if (!seconds || !Number.isFinite(seconds) || seconds <= 0) return '-';
    const total = Math.ceil(seconds);
    const days = Math.floor(total / 86400);
    const hours = Math.floor((total % 86400) / 3600);
    const minutes = Math.floor((total % 3600) / 60);
    const remainingSeconds = total % 60;
    const parts = [
      days > 0 ? `${days}d` : '',
      hours > 0 ? `${hours}h` : '',
      minutes > 0 ? `${minutes}m` : '',
      `${remainingSeconds}s`,
    ].filter(Boolean);
    return parts.join(' ');
  };

  // Initialize Graph Client
  useEffect(() => {
    if (accessToken) {
      if (!isUsableAccessToken(accessToken)) {
        clearAuthState();
        setAccessToken(null);
        addLog('Stored Microsoft token was invalid. Please sign in again.');
        return;
      }
      accessTokenRef.current = accessToken;
      const graphClient = Client.init({
        authProvider: (done) => done(null, accessTokenRef.current || accessToken)
      });
      oneDriveClient.current = new OneDriveClient(graphClient, refreshAccessToken);
      fetchUser();
      openFolder({ id: 'root', name: 'OneDrive' }, [{ id: 'root', name: 'OneDrive' }]);
    } else {
      oneDriveClient.current = null;
      accessTokenRef.current = null;
    }
  }, [accessToken]);

  const fetchUser = async () => {
    try {
      const me = await oneDriveClient.current?.getMe();
      setUser(me);
    } catch (e) {
      console.error('Fetch user error', e);
      const message = (e as any).body?.message || (e as Error).message || '';
      if ((e as any).statusCode === 401 || message.includes('JWT is not well formed')) handleLogout();
    }
  };

  const fetchFolderItems = async (folderId: string = 'root', folderName = currentFolder.name) => {
    const requestId = folderLoadRequestRef.current + 1;
    folderLoadRequestRef.current = requestId;
    setIsLoadingItems(true);
    setItemsError(null);
    setSelection(new Set());
    setRootItems([]);
    setTableScrollTop(0);
    setActiveSearchQuery('');
    try {
      const endpoint = folderId === 'root'
        ? 'root'
        : folderId;
      const driveItems: GraphDriveItem[] = [];
      await oneDriveClient.current?.listChildrenPaged(endpoint, page => {
        if (folderLoadRequestRef.current !== requestId) return;
        driveItems.push(...page);
        setRootItems(sortSyncItems(driveItems.map(mapDriveItem)));
      });
      if (folderLoadRequestRef.current !== requestId) return;

      const items = sortSyncItems(driveItems.map(mapDriveItem));
      setRootItems(items);
      addLog(`Loaded ${items.length} item${items.length === 1 ? '' : 's'} from ${folderName}`);
    } catch (e) {
       if (folderLoadRequestRef.current !== requestId) return;
       const message = (e as any).body?.message || (e as Error).message || 'Microsoft Graph rejected the file list request';
       setItemsError(message);
       setRootItems([]);
       addLog(`Error fetching items: ${message}`);
    } finally {
      if (folderLoadRequestRef.current === requestId) {
        setIsLoadingItems(false);
      }
    }
  };

  const searchFolderItems = async () => {
    const query = searchQuery.trim();
    if (!query) return;
    if (!oneDriveClient.current) return alert('Sign in to OneDrive first');

    const requestId = folderLoadRequestRef.current + 1;
    folderLoadRequestRef.current = requestId;
    setIsLoadingItems(true);
    setItemsError(null);
    setSelection(new Set());
    setRootItems([]);
    setTableScrollTop(0);
    setActiveSearchQuery(query);

    try {
      const driveItems: GraphDriveItem[] = [];
      await oneDriveClient.current.searchItems(currentFolder.id, query, page => {
        if (folderLoadRequestRef.current !== requestId) return;
        driveItems.push(...page);
        setRootItems(sortSyncItems(driveItems.map(mapDriveItem)));
      });
      if (folderLoadRequestRef.current !== requestId) return;

      const items = sortSyncItems(driveItems.map(mapDriveItem));
      setRootItems(items);
      addLog(`Found ${items.length} item${items.length === 1 ? '' : 's'} matching "${query}" in ${currentFolder.name}`);
    } catch (e) {
      if (folderLoadRequestRef.current !== requestId) return;
      const message = (e as any).body?.message || (e as Error).message || 'Microsoft Graph rejected the search request';
      setItemsError(message);
      setRootItems([]);
      addLog(`Search failed: ${message}`);
    } finally {
      if (folderLoadRequestRef.current === requestId) {
        setIsLoadingItems(false);
      }
    }
  };

  const clearSearch = () => {
    setSearchQuery('');
    if (activeSearchQuery) {
      fetchFolderItems(currentFolder.id, currentFolder.name);
    }
  };

  const openFolder = async (folder: FolderCrumb, nextBreadcrumbs?: FolderCrumb[]) => {
    setCurrentFolder(folder);
    setBreadcrumbs(nextBreadcrumbs || [...breadcrumbs, folder]);
    await fetchFolderItems(folder.id, folder.name);
  };

  const openFolderFromItem = async (item: { id: string; name: string; type: 'file' | 'folder' }) => {
    if (item.type !== 'folder' || isSyncing) return;
    await openFolder({ id: item.id, name: item.name });
  };

  // Global Auth Listeners
  useEffect(() => {
    const onSuccess = (token: string) => {
      if (!isUsableAccessToken(token)) {
        clearAuthState();
        setAccessToken(null);
        addLog('Rejected invalid Microsoft token. Please sign in again.');
        return;
      }
      setAccessToken(token);
      accessTokenRef.current = token;
      addLog('Authentication successful');
    };

    const onSessionSuccess = async () => {
      const ok = await completeAuthFromSession();
      if (!ok) addLog('Authentication completed, but the server session was not available. Please try signing in again.');
    };

    // Poll the server-side session so browser storage never contains Microsoft access tokens.
    const pollInterval = setInterval(async () => {
      if (!accessToken) {
        try {
          const res = await axios.get('/api/auth/status');
          if (res.data.authenticated && res.data.accessToken) {
            console.log('[AUTH-DEBUG] Authenticated via server session');
            onSuccess(res.data.accessToken);
            return;
          }
        } catch (e) {
          // Silent fail for polling
        }
      }
    }, 2000);

    // 3. Listen for BroadcastChannel
    const channel = new BroadcastChannel('onedrive_auth_channel');
    channel.onmessage = (event) => {
      if (event.data?.type === 'OAUTH_AUTH_SUCCESS') {
        onSessionSuccess();
      }
    };

    // 3. Listen for window messages
    const handleMessage = (event: MessageEvent) => {
      if (event.origin !== window.location.origin) return;
      if (event.data?.type === 'OAUTH_AUTH_SUCCESS') {
        onSessionSuccess();
      }
    };

    window.addEventListener('message', handleMessage);

    return () => {
      window.removeEventListener('message', handleMessage);
      clearInterval(pollInterval);
      channel.close();
    };
  }, [accessToken]);

  const handleLogin = async () => {
    try {
      addLog('Initiating Microsoft Authentication...');
      const res = await axios.get('/api/auth/url');
      
      if (!res.data.url) {
        throw new Error('Server failed to provide an authentication URL. Check your Microsoft Client ID configuration in Secrets.');
      }
      const authWindow = window.open(res.data.url, 'ms_login', 'width=600,height=700');
      
      if (!authWindow) {
        alert('Popup blocked! Please allow popups for this site to sign in.');
      }
    } catch (e: any) {
      const msg = e.response?.data?.error || e.message;
      alert(`Login error: ${msg}`);
      addLog(`Login error: ${msg}`);
    }
  };

  const handleLogout = () => {
    setAccessToken(null);
    setUser(null);
    clearAuthState();
    setRootItems([]);
    setSelection(new Set());
    setCurrentFolder({ id: 'root', name: 'OneDrive' });
    setBreadcrumbs([{ id: 'root', name: 'OneDrive' }]);
  };

  const selectLocalFolder = async () => {
    try {
      if (!('showDirectoryPicker' in window)) {
        alert('Folder download requires a Chromium-based browser such as Microsoft Edge or Chrome.');
        return;
      }
      const handle = await (window as any).showDirectoryPicker({ mode: 'readwrite' });
      const permission = await ensureDirectoryPermission(handle);
      if (!permission) {
        alert('The selected folder could not be used for writing. Choose a local folder again and allow access.');
        addLog(`Local folder could not be used: ${handle.name}`);
        return;
      }

      setLocalDir({ id: 'local', name: handle.name, handle });
      addLog(`Local folder selected: ${handle.name}`);
    } catch (e) {
      console.error('Picker error', e);
    }
  };

  const ensureDirectoryPermission = async (handle: FileSystemDirectoryHandle): Promise<boolean> => {
    try {
      const writableHandle = handle as FileSystemDirectoryHandle & {
        queryPermission?: (descriptor: { mode: 'readwrite' }) => Promise<PermissionState>;
        requestPermission?: (descriptor: { mode: 'readwrite' }) => Promise<PermissionState>;
      };

      if (!writableHandle.queryPermission || !writableHandle.requestPermission) {
        return await probeDirectoryAccess(handle);
      }

      const permission = await writableHandle.queryPermission({ mode: 'readwrite' });
      if (permission === 'granted') {
        return await probeDirectoryAccess(handle);
      }

      const requested = await writableHandle.requestPermission({ mode: 'readwrite' });
      return requested === 'granted' && await probeDirectoryAccess(handle);
    } catch (error) {
      if (isInvalidDirectoryHandleError(error)) {
        setLocalDir(null);
        addLog('The selected local folder handle could not be used. Choose the folder again and retry.');
        return false;
      }
      throw error;
    }
  };

  const probeDirectoryAccess = async (handle: FileSystemDirectoryHandle) => {
    const probeName = `.syncpoint-access-${Date.now()}-${Math.random().toString(36).slice(2)}.tmp`;
    try {
      const probe = await handle.getFileHandle(probeName, { create: true });
      const writable = await probe.createWritable();
      await writable.write('ok');
      await writable.close();
      await handle.removeEntry(probeName);
      return true;
    } catch (error) {
      try {
        await handle.removeEntry(probeName);
      } catch {
        // Best-effort cleanup for failed access probes.
      }
      if (isInvalidDirectoryHandleError(error)) {
        setLocalDir(null);
        addLog('The selected local folder handle could not be used. Choose the folder again and retry.');
        return false;
      }
      throw error;
    }
  };

  const isInvalidDirectoryHandleError = (error: unknown) => {
    const name = error instanceof DOMException ? error.name : '';
    const message = error instanceof Error ? error.message : String(error);
    return name === 'NotAllowedError'
      || name === 'NotFoundError'
      || message.includes('FileSystem')
      || message.includes('directory handle');
  };

  const operationErrorMessage = (error: unknown) => {
    const message = error instanceof Error ? error.message : String(error);
    if (isInvalidDirectoryHandleError(error)) {
      setLocalDir(null);
      return `A browser object handle failed (${message}). Choose the local folder again and retry.`;
    }
    return message;
  };

  const addLog = (message: string, stage = 'app', details?: Partial<LogEntry>) => {
    const entry: LogEntry = {
      id: `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      time: new Date().toLocaleTimeString(),
      stage,
      message,
      ...details,
    };
    setLogs(prev => [entry, ...prev]);
  };

  const refreshServerJobs = async () => {
    if (!serverArchiveEnabled) return;
    try {
      const res = await axios.get('/api/server-jobs');
      setServerJobs(res.data?.jobs || []);
    } catch {
      // Server-side downloads are optional.
    }
  };

  useEffect(() => {
    refreshServerJobs();
    const interval = setInterval(refreshServerJobs, 2000);
    return () => clearInterval(interval);
  }, [serverArchiveEnabled]);

  const handleEngineEvent = (event: ReporterEvent) => {
    if (event.type === 'log') {
      addLog(event.message, event.stage, {
        jobName: event.job?.name,
        jobPath: event.job?.remotePath,
        error: event.job?.errorMessage,
      });
      return;
    }

    if (event.type === 'progress') {
      setProgressSnapshot(event.snapshot);
      setSyncStats({
        total: event.snapshot.queuedFiles + event.snapshot.completedFiles + event.snapshot.failedFiles,
        completed: event.snapshot.completedFiles,
        errors: event.snapshot.failedFiles,
        size: event.snapshot.totalBytes,
        downloadedSize: event.snapshot.downloadedBytes,
      });
      setDownloadSpeed(event.snapshot.speedBytesPerSecond);
      return;
    }

    if (event.type === 'job') {
      updateItemStatus(event.job.itemId, event.job.status === 'failed' ? 'error' : event.job.status as SyncItem['status'], event.job.errorMessage);
      if (event.job.size > 0) {
        updateItemProgress(event.job.itemId, Math.min((event.job.downloadedBytes / event.job.size) * 100, 100));
      }
      if (event.job.status === 'failed') {
        const failed: FailedFile = {
          id: event.job.id,
          name: event.job.name,
          path: event.job.remotePath || event.job.localPath,
          reason: event.job.errorMessage || 'Unknown error',
        };
        setFailedFiles(prev => {
          const next = new Map(prev.map(file => [file.id, file]));
          next.set(failed.id, failed);
          return Array.from(next.values());
        });
      } else if (['queued', 'downloading', 'verifying', 'retrying', 'throttled', 'completed', 'skipped'].includes(event.job.status)) {
        setFailedFiles(prev => prev.filter(file => file.id !== event.job.id));
      }
      const activeStatuses: SyncItem['status'][] = ['downloading', 'verifying', 'retrying', 'throttled'];
      setActiveFiles(prev => {
        const next = new Map(prev.map(file => [file.id, file]));
        if (activeStatuses.includes(event.job.status as SyncItem['status'])) {
          next.set(event.job.id, {
            id: event.job.id,
            name: event.job.name,
            status: event.job.status as SyncItem['status'],
            progress: event.job.size > 0 ? Math.min((event.job.downloadedBytes / event.job.size) * 100, 100) : 100,
          });
        } else {
          next.delete(event.job.id);
        }
        return Array.from(next.values()).slice(0, downloadSettings.maxGlobalConcurrentDownloads + 2);
      });
    }
  };

  const updateDownloadSetting = <K extends keyof DownloadSettings>(key: K, value: DownloadSettings[K]) => {
    setDownloadSettings(prev => ({ ...prev, [key]: value }));
  };

  const toggleSelection = (id: string) => {
    setSelection(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const startSync = async () => {
    if (!localDir) return alert('Select a local folder first');
    if (selection.size === 0) return alert('Select items to sync');
    if (!oneDriveClient.current) return alert('Sign in to OneDrive first');
    if (!await ensureDirectoryPermission(localDir.handle)) {
      alert('Write permission is required to download files into the selected folder. Choose the folder again and allow access.');
      addLog(`Write permission missing for local folder: ${localDir.name}`);
      return;
    }

    setIsSyncing(true);
    setDownloadSpeed(0);
    setActiveFiles([]);
    setFailedFiles([]);
    if (resetLogsOnStart) setLogs([]);
    
    const itemsToProcess = rootItems.filter(i => selection.has(i.id));
    addLog(`Preparing ${itemsToProcess.length} selected item${itemsToProcess.length === 1 ? '' : 's'}...`);

    try {
      const engine = new DownloadEngine(oneDriveClient.current, handleEngineEvent, downloadSettings);
      downloadEngine.current = engine;
      const selections: SourceSelection[] = itemsToProcess.map(item => ({
        id: item.id,
        name: item.name,
        type: item.type,
      }));
      await engine.start(selections, localDir.handle);
    } catch (e) {
      addLog(`Sync failed: ${operationErrorMessage(e)}`);
    } finally {
      setIsSyncing(false);
      setDownloadSpeed(0);
      setActiveFiles([]);
      downloadEngine.current = null;
    }
  };

  const dryRunSelected = async () => {
    if (!localDir) return alert('Select a local folder first');
    if (selection.size === 0) return alert('Select items to preflight');
    if (!oneDriveClient.current) return alert('Sign in to OneDrive first');
    if (!await ensureDirectoryPermission(localDir.handle)) {
      alert('Read/write permission is required to inspect the selected target folder. Choose the folder again and allow access.');
      addLog(`Permission missing for local folder: ${localDir.name}`);
      return;
    }

    setIsSyncing(true);
    setDownloadSpeed(0);
    setActiveFiles([]);
    if (resetLogsOnStart) setLogs([]);

    const itemsToProcess = rootItems.filter(i => selection.has(i.id));
    addLog(`Preflighting ${itemsToProcess.length} selected item${itemsToProcess.length === 1 ? '' : 's'}...`);

    try {
      const engine = new DownloadEngine(oneDriveClient.current, handleEngineEvent, downloadSettings);
      downloadEngine.current = engine;
      const selections: SourceSelection[] = itemsToProcess.map(item => ({
        id: item.id,
        name: item.name,
        type: item.type,
      }));
      await engine.dryRun(selections, localDir.handle);
    } catch (e) {
      addLog(`Dry run failed: ${operationErrorMessage(e)}`);
    } finally {
      setIsSyncing(false);
      setDownloadSpeed(0);
      setActiveFiles([]);
      downloadEngine.current = null;
    }
  };

  const startIncrementalArchive = async () => {
    if (!localDir) return alert('Select a local folder first');
    if (!oneDriveClient.current) return alert('Sign in to OneDrive first');
    if (!await ensureDirectoryPermission(localDir.handle)) {
      alert('Write permission is required to archive changed files into the selected folder. Choose the folder again and allow access.');
      addLog(`Write permission missing for local folder: ${localDir.name}`);
      return;
    }

    setIsSyncing(true);
    setDownloadSpeed(0);
    setActiveFiles([]);
    setFailedFiles([]);
    if (resetLogsOnStart) setLogs([]);
    addLog('Starting incremental archive from OneDrive delta state...');

    try {
      const engine = new DownloadEngine(oneDriveClient.current, handleEngineEvent, downloadSettings);
      downloadEngine.current = engine;
      await engine.startIncremental(localDir.handle);
    } catch (e) {
      addLog(`Incremental archive failed: ${operationErrorMessage(e)}`);
    } finally {
      setIsSyncing(false);
      setDownloadSpeed(0);
      setActiveFiles([]);
      downloadEngine.current = null;
    }
  };

  const repairArchive = async () => {
    if (!localDir) return alert('Select a local folder first');
    if (!oneDriveClient.current) return alert('Sign in to OneDrive first');
    if (!await ensureDirectoryPermission(localDir.handle)) {
      alert('Write permission is required to repair files in the selected folder. Choose the folder again and allow access.');
      addLog(`Write permission missing for local folder: ${localDir.name}`);
      return;
    }

    setIsSyncing(true);
    setDownloadSpeed(0);
    setActiveFiles([]);
    setFailedFiles([]);
    if (resetLogsOnStart) setLogs([]);
    addLog(selection.size > 0
      ? 'Checking selected OneDrive items for missing or changed local files...'
      : 'Checking latest saved archive manifest for missing or changed local files...');

    try {
      const engine = new DownloadEngine(oneDriveClient.current, handleEngineEvent, downloadSettings);
      downloadEngine.current = engine;
      if (selection.size > 0) {
        const itemsToProcess = rootItems.filter(i => selection.has(i.id));
        const selections: SourceSelection[] = itemsToProcess.map(item => ({
          id: item.id,
          name: item.name,
          type: item.type,
        }));
        await engine.repairSelections(selections, localDir.handle);
      } else {
        await engine.repairFromLatestManifest(localDir.handle);
      }
    } catch (e) {
      addLog(`Repair failed: ${operationErrorMessage(e)}`);
    } finally {
      setIsSyncing(false);
      setDownloadSpeed(0);
      setActiveFiles([]);
      downloadEngine.current = null;
    }
  };

  const selectedSourceSelections = () => {
    const itemsToProcess = rootItems.filter(i => selection.has(i.id));
    return itemsToProcess.map(item => ({
      id: item.id,
      name: item.name,
      type: item.type,
    }));
  };

  const startServerArchiveJob = async (mode: 'start' | 'dry-run' | 'repair' | 'incremental') => {
    if (!serverTargetRoot) return alert('SERVER_DOWNLOAD_ROOT is not configured on the server.');
    if (!oneDriveClient.current) return alert('Sign in to OneDrive first');
    if (mode !== 'incremental' && selection.size === 0) return alert('Select OneDrive files or folders first');
    const endpoint = mode === 'start'
      ? '/api/server-jobs/start'
      : mode === 'dry-run'
        ? '/api/server-jobs/dry-run'
        : mode === 'repair'
          ? '/api/server-jobs/repair'
          : '/api/server-jobs/incremental';
    addLog(`Starting server ${mode === 'start' ? 'archive' : mode} job...`);
    try {
      const res = await axios.post(endpoint, {
        selections: selectedSourceSelections(),
        settings: downloadSettings,
      });
      const job = res.data?.job as ServerArchiveJob;
      setActiveServerJobId(job.id);
      await refreshServerJobs();
      addLog(`Server job started: ${job.id}`);
    } catch (error: any) {
      addLog(error.response?.data?.error || `Server job failed to start: ${error.message}`);
    }
  };

  const retryFailedDownloads = async (jobIds?: string[]) => {
    if (!localDir) return alert('Select a local folder first');
    if (!oneDriveClient.current) return alert('Sign in to OneDrive first');
    const failedJobIds = jobIds || failedFiles.map(file => file.id);
    if (failedJobIds.length === 0) return;
    if (!await ensureDirectoryPermission(localDir.handle)) {
      alert('Write permission is required to retry downloads into the selected folder. Choose the folder again and allow access.');
      addLog(`Write permission missing for local folder: ${localDir.name}`);
      return;
    }

    setIsSyncing(true);
    setDownloadSpeed(0);
    setActiveFiles([]);

    addLog(`Retrying ${failedJobIds.length} failed file${failedJobIds.length === 1 ? '' : 's'}...`);

    try {
      const engine = new DownloadEngine(oneDriveClient.current, handleEngineEvent, downloadSettings);
      downloadEngine.current = engine;
      await engine.retryJobs(failedJobIds, localDir.handle);
    } catch (e) {
      addLog(`Retry failed: ${operationErrorMessage(e)}`);
    } finally {
      setIsSyncing(false);
      setDownloadSpeed(0);
      setActiveFiles([]);
      downloadEngine.current = null;
    }
  };

  const resetStoredDownloadState = async () => {
    if (isSyncing) return;
    const confirmed = window.confirm('Reset saved download resume state? Completed files are not deleted.');
    if (!confirmed) return;
    await new DownloadJobStore().resetJobs();
    setFailedFiles([]);
    setProgressSnapshot(null);
    addLog('Saved download resume state was reset.');
  };

  const cleanupStaleDownloadState = async () => {
    if (isSyncing) return;
    const removed = await new DownloadJobStore().cleanupStaleJobs();
    addLog(`Cleaned ${removed} stale saved download job${removed === 1 ? '' : 's'}.`);
  };

  const clearStalePartials = async () => {
    if (isSyncing) return;
    if (!localDir) return alert('Select a local folder first');
    if (!await ensureDirectoryPermission(localDir.handle)) {
      alert('Write permission is required to remove stale partial files. Choose the folder again and allow access.');
      addLog(`Write permission missing for local folder: ${localDir.name}`);
      return;
    }

    const store = new DownloadJobStore();
    const jobs = await store.getJobsByStatus(['failed', 'cancelled', 'stale_remote_changed']);
    const committer = new FileCommitter();
    await Promise.all(jobs.map(job => committer.clearPartialByPath(localDir.handle, job.localPath)));
    addLog(`Cleared partial files for ${jobs.length} recoverable job${jobs.length === 1 ? '' : 's'}.`);
  };

  const updateItemStatus = (id: string, status: SyncItem['status'], error?: string) => {
    // This is expensive for large sets, but for root it's fine.
    setRootItems(prev => prev.map(i => i.id === id ? { ...i, status, error } : i));
  };

  const updateItemProgress = (id: string, progress: number) => {
    setRootItems(prev => prev.map(i => i.id === id ? { ...i, progress } : i));
  };

  const stopSync = () => {
    const activeServerJob = serverJobs.find(job => job.id === activeServerJobId);
    if (targetMode === 'server' && activeServerJob && ['queued', 'scanning', 'downloading'].includes(activeServerJob.status)) {
      axios.post(`/api/server-jobs/${activeServerJob.id}/cancel`)
        .then(refreshServerJobs)
        .catch((error: any) => addLog(error.response?.data?.error || `Cancel failed: ${error.message}`));
      addLog('Server job cancellation requested.');
      return;
    }
    downloadEngine.current?.cancel();
    setIsSyncing(false);
    setActiveFiles([]);
    addLog('Sync cancelled by user');
  };

  const activeServerJob = serverJobs.find(job => job.id === activeServerJobId)
    || serverJobs.find(job => ['queued', 'scanning', 'downloading'].includes(job.status));
  const isServerBusy = targetMode === 'server' && Boolean(activeServerJob && ['queued', 'scanning', 'downloading'].includes(activeServerJob.status));
  const activeSnapshot = targetMode === 'server' && activeServerJob ? activeServerJob.snapshot : progressSnapshot;
  const globalProgress = targetMode === 'server' && activeServerJob
    ? activeServerJob.snapshot.stagePercent
    : syncStats.size > 0
    ? Math.min((syncStats.downloadedSize / syncStats.size) * 100, 100)
    : 0;
  const remainingItems = targetMode === 'server' && activeServerJob
    ? activeServerJob.snapshot.queuedFiles
    : Math.max(syncStats.total - syncStats.completed, 0);
  const activeStage = activeSnapshot?.status || (isSyncing || isServerBusy ? 'downloading' : 'idle');
  const activeStageLabel = activeStage.replaceAll('_', ' ');
  const activeSpeed = targetMode === 'server' ? activeSnapshot?.speedBytesPerSecond || 0 : downloadSpeed;
  const activeStagePercent = Math.round(
    ['downloading', 'retrying', 'throttled'].includes(activeStage)
      ? globalProgress
      : activeSnapshot?.stagePercent ?? globalProgress
  );
  const rowHeight = 49;
  const tableViewportHeight = 640;
  const overscanRows = 8;
  const firstVisibleRow = Math.max(0, Math.floor(tableScrollTop / rowHeight) - overscanRows);
  const visibleRowCount = Math.ceil(tableViewportHeight / rowHeight) + overscanRows * 2;
  const visibleItems = rootItems.slice(firstVisibleRow, firstVisibleRow + visibleRowCount);
  const topSpacerHeight = firstVisibleRow * rowHeight;
  const bottomSpacerHeight = Math.max(0, (rootItems.length - firstVisibleRow - visibleItems.length) * rowHeight);
  const activeServerLogEntries: LogEntry[] = targetMode === 'server' && activeServerJob
    ? activeServerJob.log.map((line, index) => {
        const match = line.match(/^(\d{1,2}:\d{2}:\d{2}(?:\s+[AP]M)?)\s+(.+)$/i);
        return {
          id: `server-${activeServerJob.id}-${index}`,
          time: match?.[1] || '',
          stage: 'server',
          message: match?.[2] || line,
          jobName: activeServerJob.status,
          jobPath: activeServerJob.targetRoot,
        };
      })
    : [];
  const visibleLogs = [...activeServerLogEntries, ...logs];

  return (
    <div className="flex flex-col h-screen w-full bg-slate-50 text-slate-900 font-sans overflow-hidden dark:bg-neutral-950 dark:text-neutral-100">
      {/* Header: Geometric Alignment */}
      <header className="h-16 bg-white border-b border-slate-200 px-6 flex items-center justify-between shadow-sm shrink-0 dark:bg-neutral-900 dark:border-neutral-800">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 bg-blue-600 rounded flex items-center justify-center">
            <Cloud className="text-white" size={20} />
          </div>
          <h1 className="text-xl font-bold tracking-tight text-slate-800 uppercase dark:text-neutral-100">
            OneDrive <span className="text-blue-600 font-light dark:text-blue-400">Archiver</span>
          </h1>
        </div>
        
        <div className="flex items-center gap-4">
          <button
            onClick={() => setTheme(theme === 'dark' ? 'light' : 'dark')}
            className="h-9 w-9 border border-slate-200 rounded flex items-center justify-center text-slate-500 hover:text-blue-600 hover:bg-slate-50 transition-colors dark:border-neutral-700 dark:text-neutral-300 dark:hover:bg-neutral-800 dark:hover:text-blue-400"
            title={theme === 'dark' ? 'Switch to light mode' : 'Switch to dark mode'}
          >
            {theme === 'dark' ? <Sun size={16} /> : <Moon size={16} />}
          </button>
          {user ? (
            <div className="flex items-center gap-3 px-4 py-1.5 bg-slate-100 dark:bg-neutral-800 border border-slate-200 dark:border-neutral-700 rounded-md text-slate-700 dark:text-neutral-100">
              <div className="w-6 h-6 bg-blue-200 dark:bg-blue-500/30 rounded-full flex items-center justify-center">
                <span className="text-[10px] font-bold text-blue-700 dark:text-blue-100">{user.displayName?.charAt(0)}</span>
              </div>
              <span className="text-sm font-medium">{user.displayName || user.userPrincipalName}</span>
              <button
                onClick={handleLogout}
                title="Sign out of the current Microsoft session."
                className="ml-2 text-slate-400 dark:text-neutral-500 hover:text-red-500 transition-colors"
              >
                <LogOut size={14} />
              </button>
            </div>
          ) : (
            <button 
              onClick={handleLogin}
              title="Sign in with Microsoft so the app can read OneDrive files."
              className="px-6 py-2 bg-blue-600 text-white text-sm font-semibold rounded hover:bg-blue-500 disabled:bg-neutral-800 disabled:text-neutral-500 disabled:border disabled:border-neutral-700 disabled:opacity-100 transition-all flex items-center gap-2"
            >
              <LogIn size={16} />
              Sign In
            </button>
          )}
        </div>
      </header>

      {/* Main Viewport */}
      <main className="flex flex-1 overflow-hidden">
        {/* Left Sidebar: Account & Source */}
        <nav className="w-64 border-r border-slate-200 dark:border-neutral-700 bg-white dark:bg-neutral-900 p-6 flex flex-col gap-8 shrink-0 dark:bg-neutral-900 dark:border-neutral-800">
          <div>
            <p className="text-[10px] font-bold text-slate-400 dark:text-neutral-500 uppercase tracking-widest mb-4">Storage Source</p>
            <ul className="space-y-1">
              <li className="px-3 py-2 bg-blue-50 dark:bg-neutral-800 text-blue-700 dark:text-neutral-100 rounded border-r-4 border-blue-600 dark:border-blue-400 font-medium text-sm cursor-pointer">
                OneDrive Personal
              </li>
            </ul>
          </div>
          
          <div>
            <p className="text-[10px] font-bold text-slate-400 dark:text-neutral-500 uppercase tracking-widest mb-4">Archive Target</p>
            <div className="grid grid-cols-2 gap-1 mb-3">
              <button
                onClick={() => setTargetMode('browser')}
                className={`py-1.5 rounded text-[10px] font-bold uppercase ${targetMode === 'browser' ? 'bg-blue-50 dark:bg-neutral-800 text-blue-700 dark:text-neutral-100' : 'text-slate-500 dark:text-neutral-400 border border-slate-200 dark:border-neutral-700'}`}
                title="Write downloads to a folder selected by this browser."
              >
                Browser
              </button>
              <button
                onClick={() => setTargetMode('server')}
                disabled={!serverArchiveEnabled}
                className={`py-1.5 rounded text-[10px] font-bold uppercase ${targetMode === 'server' ? 'bg-blue-50 dark:bg-neutral-800 text-blue-700 dark:text-neutral-100' : 'text-slate-500 dark:text-neutral-400 border border-slate-200 dark:border-neutral-700'}`}
                title={serverArchiveEnabled ? 'Write downloads on the server under SERVER_DOWNLOAD_ROOT.' : 'Server archive mode is disabled on this deployment.'}
              >
                Server
              </button>
            </div>
            {targetMode === 'browser' ? (
              <div 
                onClick={selectLocalFolder}
                title="Choose the local folder where archived files will be written."
                className={`p-4 border rounded-lg cursor-pointer transition-all ${localDir ? 'border-slate-300 bg-slate-50 dark:bg-neutral-800/60 dark:border-neutral-700' : 'border-slate-200 dark:border-neutral-700 hover:border-slate-400 dark:border-neutral-700 dark:hover:border-neutral-500'}`}
              >
                <div className="flex items-center gap-2 mb-1">
                  <HardDrive size={16} className={localDir ? 'text-blue-600 dark:text-blue-400' : 'text-slate-400 dark:text-neutral-500'} />
                  <span className="text-xs font-bold truncate">{localDir ? localDir.name : 'Choose Folder'}</span>
                </div>
                <p className="text-[10px] text-slate-500 dark:text-neutral-400 leading-tight">
                  {localDir ? 'Browser target ready' : 'Select browser destination'}
                </p>
              </div>
            ) : serverArchiveEnabled ? (
              <div className={`p-4 border rounded-lg ${serverTargetRoot ? 'border-slate-300 bg-slate-50 dark:bg-neutral-800/60 dark:border-neutral-700' : 'border-red-200 dark:border-red-900 bg-red-50 dark:bg-red-950/30'}`}>
                <div className="flex items-center gap-2 mb-1">
                  <HardDrive size={16} className={`shrink-0 ${serverTargetRoot ? 'text-blue-600 dark:text-blue-400' : 'text-red-500'}`} />
                  <span className="min-w-0 text-xs font-bold break-words leading-snug">{serverTargetRoot || 'SERVER_DOWNLOAD_ROOT missing'}</span>
                </div>
                <p className="text-[10px] text-slate-500 dark:text-neutral-400 leading-tight">
                  {serverTargetRoot ? 'Server target configured' : 'Set SERVER_DOWNLOAD_ROOT in the server environment'}
                </p>
              </div>
            ) : (
              <div className="p-4 border rounded-lg border-slate-200 dark:border-neutral-700 bg-slate-50 dark:bg-neutral-800/60">
                <div className="flex items-center gap-2 mb-1">
                  <HardDrive size={16} className="text-slate-400 dark:text-neutral-500" />
                  <span className="text-xs font-bold truncate">Browser-only deployment</span>
                </div>
                <p className="text-[10px] text-slate-500 dark:text-neutral-400 leading-tight">
                  Server downloads are disabled
                </p>
              </div>
            )}
          </div>

          <div className="mt-auto p-4 bg-slate-50 dark:bg-neutral-800 rounded-lg border border-slate-100 dark:border-neutral-800">
            <p className="text-xs text-slate-500 dark:text-neutral-400 leading-tight">
              Status: <br/>
              <span className="font-bold text-slate-700 dark:text-neutral-100 uppercase">
                {isSyncing || isServerBusy ? `${activeStageLabel} ${activeStagePercent}%` : 'Idle / Ready'}
              </span>
            </p>
          </div>
        </nav>

        {/* Center: File Explorer Selection */}
        <section className="flex-1 bg-white dark:bg-neutral-900 flex flex-col overflow-hidden">
          <div className="p-6 border-b border-slate-100 dark:border-neutral-800 flex justify-between items-center bg-slate-50 dark:bg-neutral-800/50">
            <div>
              <div className="flex items-center gap-2">
                <button
                  onClick={() => setMainTab('explorer')}
                  className={`px-3 py-1.5 rounded text-xs font-bold uppercase tracking-wide ${mainTab === 'explorer' ? 'bg-blue-50 dark:bg-neutral-800 text-blue-700 dark:text-neutral-100' : 'text-slate-500 dark:text-neutral-400 hover:bg-slate-100 dark:hover:bg-neutral-700 dark:bg-neutral-800'}`}
                >
                  OneDrive explorer
                </button>
                <button
                  onClick={() => setMainTab('activity')}
                  className={`px-3 py-1.5 rounded text-xs font-bold uppercase tracking-wide ${mainTab === 'activity' ? 'bg-blue-50 dark:bg-neutral-800 text-blue-700 dark:text-neutral-100' : 'text-slate-500 dark:text-neutral-400 hover:bg-slate-100 dark:hover:bg-neutral-700 dark:bg-neutral-800'}`}
                >
                  Activity log
                  {failedFiles.length > 0 && <span className="ml-2 text-red-600 dark:text-red-400">{failedFiles.length}</span>}
                </button>
              </div>
              {mainTab === 'explorer' ? (
                <div className="flex items-center gap-1 mt-2 text-xs text-slate-500 dark:text-neutral-400">
                  {breadcrumbs.map((crumb, index) => (
                    <React.Fragment key={crumb.id}>
                      {index > 0 && <ChevronRight size={12} className="text-slate-300" />}
                      <button
                        onClick={() => openFolder(crumb, breadcrumbs.slice(0, index + 1))}
                        disabled={isSyncing || crumb.id === currentFolder.id}
                        className="hover:text-blue-600 dark:hover:text-sky-300 disabled:text-slate-700 dark:disabled:text-neutral-200 disabled:font-semibold"
                      >
                        {crumb.name}
                      </button>
                    </React.Fragment>
                  ))}
                  {activeSearchQuery && (
                    <span className="ml-2 text-blue-600 dark:text-sky-300">
                      Search: "{activeSearchQuery}"
                    </span>
                  )}
                </div>
              ) : (
                <div className="mt-2 text-xs text-slate-500 dark:text-neutral-400">{logs.length} log entries retained for this session</div>
              )}
            </div>
            <div className="flex items-center gap-2">
              {mainTab === 'explorer' ? (
                <>
                  <form
                    onSubmit={event => {
                      event.preventDefault();
                      searchFolderItems();
                    }}
                    className="flex items-center"
                  >
                    <div className="flex h-8 items-center rounded border border-slate-200 bg-white dark:border-neutral-700 dark:bg-neutral-900">
                      <Search size={14} className="ml-2 text-slate-400 dark:text-neutral-500" />
                      <input
                        value={searchQuery}
                        onChange={event => setSearchQuery(event.target.value)}
                        disabled={!accessToken || isLoadingItems || isSyncing}
                        placeholder="Search files"
                        className="h-full w-40 bg-transparent px-2 text-xs text-slate-700 outline-none placeholder:text-slate-400 disabled:text-slate-400 dark:text-neutral-100 dark:placeholder:text-neutral-500"
                      />
                      {searchQuery && (
                        <button
                          type="button"
                          onClick={clearSearch}
                          disabled={isLoadingItems || isSyncing}
                          className="mr-1 flex h-6 w-6 items-center justify-center rounded text-slate-400 hover:bg-slate-100 hover:text-slate-700 disabled:opacity-40 dark:hover:bg-neutral-800 dark:hover:text-neutral-100"
                          title="Clear search"
                        >
                          <X size={13} />
                        </button>
                      )}
                    </div>
                    <button
                      type="submit"
                      disabled={!accessToken || !searchQuery.trim() || isLoadingItems || isSyncing}
                      className="ml-1 h-8 w-8 border border-slate-200 dark:border-neutral-700 text-slate-700 dark:text-neutral-100 text-xs font-semibold rounded hover:bg-white dark:hover:bg-neutral-800 dark:bg-neutral-900 transition-colors disabled:text-slate-400 dark:disabled:text-neutral-500 flex items-center justify-center"
                      title="Search current folder"
                    >
                      <Search size={14} />
                    </button>
                  </form>
                  <button
                    onClick={() => fetchFolderItems(currentFolder.id)}
                    title="Reload the current OneDrive folder."
                    className="px-3 py-1 border border-slate-200 dark:border-neutral-700 text-slate-700 dark:text-neutral-100 text-xs font-semibold rounded hover:bg-white dark:hover:bg-neutral-800 dark:bg-neutral-900 transition-colors disabled:text-slate-400 dark:disabled:text-neutral-500"
                    disabled={!accessToken || isLoadingItems || isSyncing}
                  >
                    Refresh
                  </button>
                  <button
                    onClick={() => setSelection(new Set(rootItems.map(i => i.id)))}
                    title="Select every visible item in the current folder or search result."
                    className="px-3 py-1 border border-slate-200 dark:border-neutral-700 text-slate-700 dark:text-neutral-100 text-xs font-semibold rounded hover:bg-white dark:hover:bg-neutral-800 dark:bg-neutral-900 transition-colors disabled:text-slate-400 dark:disabled:text-neutral-500"
                    disabled={!accessToken}
                  >
                    Select All
                  </button>
                  <button
                    onClick={() => setSelection(new Set())}
                    title="Clear the current item selection."
                    className="px-3 py-1 border border-slate-200 dark:border-neutral-700 text-slate-700 dark:text-neutral-100 text-xs font-semibold rounded hover:bg-white dark:hover:bg-neutral-800 dark:bg-neutral-900 transition-colors disabled:text-slate-400 dark:disabled:text-neutral-500"
                    disabled={!accessToken}
                  >
                    Deselect
                  </button>
                </>
              ) : (
                <>
                  <label className="flex items-center gap-2 text-xs text-slate-600 dark:text-neutral-300">
                    <input
                      type="checkbox"
                      checked={resetLogsOnStart}
                      onChange={event => setResetLogsOnStart(event.target.checked)}
                    />
                    Reset on new download
                  </label>
                  <button
                    onClick={() => {
                      setLogs([]);
                      setFailedFiles([]);
                    }}
                    title="Clear visible logs and the current failed-file list. Saved resume state is not changed."
                    className="px-3 py-1 border border-slate-200 dark:border-neutral-700 text-slate-700 dark:text-neutral-100 text-xs font-semibold rounded hover:bg-white dark:hover:bg-neutral-800 dark:bg-neutral-900 transition-colors"
                  >
                    Clear Logs
                  </button>
                  <button
                    onClick={cleanupStaleDownloadState}
                    disabled={isSyncing}
                    title="Delete old completed, skipped, failed, cancelled, and stale saved download jobs from browser storage."
                    className="px-3 py-1 border border-slate-200 dark:border-neutral-700 text-slate-700 dark:text-neutral-100 text-xs font-semibold rounded hover:bg-white dark:hover:bg-neutral-800 dark:bg-neutral-900 transition-colors disabled:text-slate-400 dark:disabled:text-neutral-500"
                  >
                    Clean Stale State
                  </button>
                  <button
                    onClick={clearStalePartials}
                    disabled={isSyncing || !localDir}
                    title="Remove leftover .partial files for failed, cancelled, or stale jobs in the selected local folder."
                    className="px-3 py-1 border border-slate-200 dark:border-neutral-700 text-slate-700 dark:text-neutral-100 text-xs font-semibold rounded hover:bg-white dark:hover:bg-neutral-800 dark:bg-neutral-900 transition-colors disabled:text-slate-400 dark:disabled:text-neutral-500"
                  >
                    Clear Partials
                  </button>
                  <button
                    onClick={resetStoredDownloadState}
                    disabled={isSyncing}
                    title="Forget all saved download jobs and resume progress. Completed local files are not deleted."
                    className="px-3 py-1 border border-red-200 dark:border-red-900 text-red-700 dark:text-red-300 text-xs font-semibold rounded hover:bg-red-50 dark:hover:bg-red-950/40 dark:bg-neutral-900 transition-colors disabled:text-slate-400 dark:disabled:text-neutral-500"
                  >
                    Reset Resume State
                  </button>
                </>
              )}
            </div>
          </div>
          
          <div className="flex-1 overflow-hidden p-6 flex flex-col">
            {mainTab === 'explorer' ? (
              <ExplorerTable
                accessToken={accessToken}
                isLoadingItems={isLoadingItems}
                isSyncing={isSyncing}
                itemsError={itemsError}
                rootItemsLength={rootItems.length}
                emptyMessage={activeSearchQuery ? 'No matching files found' : 'This folder is empty'}
                visibleItems={visibleItems}
                selection={selection}
                rowHeight={rowHeight}
                topSpacerHeight={topSpacerHeight}
                bottomSpacerHeight={bottomSpacerHeight}
                onScroll={setTableScrollTop}
                onToggleSelection={toggleSelection}
                onOpenFolder={openFolderFromItem}
                onRetry={() => fetchFolderItems(currentFolder.id)}
              />
            ) : (
              <ActivityLog
                logs={visibleLogs}
                failedFiles={failedFiles}
                isSyncing={isSyncing}
                canRetry={Boolean(localDir)}
                onRetryFile={id => retryFailedDownloads([id])}
              />
            )}
          </div>
        </section>

        {/* Right Sidebar: Controls & Progress */}
        <aside className="w-80 border-l border-slate-200 dark:border-neutral-700 bg-white dark:bg-neutral-900 p-6 flex flex-col shadow-inner shrink-0 overflow-hidden min-h-0">
          <div className="flex-1 min-h-0 overflow-y-auto pr-1 space-y-6">
            <div className="space-y-4">
              <p className="text-[10px] font-bold text-slate-400 dark:text-neutral-500 uppercase tracking-widest">Task metrics</p>
              <div className="grid grid-cols-2 gap-3">
                <div className="p-3 bg-slate-50 dark:bg-neutral-800 border border-slate-200 dark:border-neutral-700 rounded">
                  <p className="text-[8px] font-bold text-slate-400 dark:text-neutral-500 uppercase mb-1">Completed</p>
                  <div className="text-xl font-bold text-slate-800 dark:text-neutral-100">{targetMode === 'server' && activeServerJob ? activeServerJob.snapshot.completedFiles : syncStats.completed}</div>
                </div>
                <div className="p-3 bg-slate-50 dark:bg-neutral-800 border border-slate-200 dark:border-neutral-700 rounded">
                  <p className="text-[8px] font-bold text-slate-400 dark:text-neutral-500 uppercase mb-1">Failures</p>
                  <div className="text-xl font-bold text-red-600 dark:text-red-400">{targetMode === 'server' && activeServerJob ? activeServerJob.snapshot.failedFiles : syncStats.errors}</div>
                </div>
              </div>
              {failedFiles.length > 0 && (
                <div className="border border-red-100 dark:border-red-900/60 bg-red-50 dark:bg-red-950/40 rounded p-3 space-y-2">
                  <p className="text-[9px] font-bold uppercase text-red-700 dark:text-red-300">Failed file details</p>
                  <div className="space-y-2 max-h-28 overflow-y-auto">
                    {failedFiles.map(file => (
                      <div key={file.id} className="text-xs flex items-start gap-2">
                        <div className="min-w-0 flex-1">
                          <div className="font-semibold text-red-900 dark:text-red-200 truncate" title={file.name}>{file.name}</div>
                          <div className="text-red-700 dark:text-red-300 truncate" title={file.reason}>{file.reason}</div>
                        </div>
                        <button
                          onClick={() => retryFailedDownloads([file.id])}
                          disabled={isSyncing || !localDir}
                          className="mt-0.5 flex h-6 w-6 shrink-0 items-center justify-center rounded border border-red-200 bg-white text-red-700 hover:bg-red-100 disabled:opacity-40 disabled:hover:bg-white dark:border-red-900 dark:bg-neutral-900 dark:text-red-300 dark:hover:bg-red-950"
                          title="Retry this file"
                        >
                          <RotateCcw size={12} />
                        </button>
                      </div>
                    ))}
                  </div>
                  <div className="flex items-center justify-between gap-2">
                    <button
                      onClick={() => setMainTab('activity')}
                      className="text-[10px] font-bold uppercase text-red-700 dark:text-red-300 hover:text-red-900 dark:hover:text-red-200"
                    >
                      View full activity
                    </button>
                    <button
                      onClick={() => retryFailedDownloads()}
                      disabled={isSyncing || !localDir}
                      className="text-[10px] font-bold uppercase text-white bg-red-600 hover:bg-red-700 disabled:opacity-40 disabled:hover:bg-red-600 rounded px-2 py-1"
                    >
                      Retry Failed
                    </button>
                  </div>
                </div>
              )}
            </div>

            <div className="space-y-4">
              <p className="text-[10px] font-bold text-slate-400 dark:text-neutral-500 uppercase tracking-widest">Global Progress</p>
              <div className="p-4 border border-slate-200 dark:border-neutral-700 rounded bg-white dark:bg-neutral-900 shadow-sm">
              <div className="flex justify-between items-center mb-3">
                <span className={`text-[10px] font-bold uppercase flex items-center gap-1.5 ${isSyncing || isServerBusy ? 'text-blue-600 dark:text-blue-400' : 'text-slate-400 dark:text-neutral-500'}`}>
                  <span className={`w-2 h-2 rounded-full ${isSyncing || isServerBusy ? 'bg-blue-500 animate-pulse' : 'bg-slate-300'}`}></span>
                  {isSyncing || isServerBusy ? `${activeStageLabel} ${activeStagePercent}%` : 'Idle'}
                </span>
                <span className="text-xs font-mono font-bold">
                  {Math.round(globalProgress)}%
                </span>
              </div>
              
              <div className="h-1.5 w-full bg-slate-100 dark:bg-neutral-800 rounded-full overflow-hidden mb-4">
                <motion.div 
                  className="h-full bg-blue-600"
                  initial={{ width: 0 }}
                  animate={{ width: `${globalProgress}%` }}
                />
              </div>

              <div>
                {!isSyncing && !isServerBusy ? (
                  <div className="space-y-2">
                    <button 
                      onClick={targetMode === 'server' ? () => startServerArchiveJob('start') : startSync}
                      disabled={!accessToken || selection.size === 0 || (targetMode === 'browser' ? !localDir : !serverTargetRoot)}
                      title={targetMode === 'server' ? 'Start a server-side archive job that continues after this browser closes.' : 'Download the selected OneDrive files or folders into the selected local archive folder.'}
                      className="w-full py-2.5 bg-blue-600 text-white rounded font-bold text-xs uppercase tracking-wider shadow-lg shadow-blue-100 dark:shadow-none hover:bg-blue-700 disabled:bg-slate-100 disabled:text-slate-500 disabled:border disabled:border-slate-200 dark:disabled:bg-neutral-800 dark:disabled:text-neutral-500 dark:disabled:border-neutral-700 disabled:opacity-100 disabled:shadow-none transition-all"
                    >
                      {syncStats.completed > 0 && syncStats.completed < syncStats.total ? 'Resume Archive' : 'Start Archive'}
                    </button>
                    <div className="grid grid-cols-3 gap-2">
                      <button
                        onClick={targetMode === 'server' ? () => startServerArchiveJob('dry-run') : dryRunSelected}
                        disabled={!accessToken || selection.size === 0 || (targetMode === 'browser' ? !localDir : !serverTargetRoot)}
                        title={targetMode === 'server' ? 'Preview the server-side archive job without writing files.' : 'Preview the selected archive run without downloading files.'}
                        className="min-w-0 px-1 py-2 border border-slate-200 dark:border-neutral-700 text-slate-700 dark:text-neutral-100 rounded font-bold text-[9px] uppercase leading-tight hover:bg-slate-50 dark:hover:bg-neutral-800 disabled:text-slate-400 dark:disabled:text-neutral-500 whitespace-normal break-words"
                      >
                        Dry Run
                      </button>
                      <button
                        onClick={targetMode === 'server' ? () => startServerArchiveJob('incremental') : startIncrementalArchive}
                        disabled={!accessToken || (targetMode === 'browser' ? !localDir : !serverTargetRoot)}
                        title={targetMode === 'server' ? 'Archive changed OneDrive files on the server from the saved delta state.' : 'Archive only OneDrive files that changed since the last saved incremental scan.'}
                        className="min-w-0 px-1 py-2 border border-slate-200 dark:border-neutral-700 text-slate-700 dark:text-neutral-100 rounded font-bold text-[9px] uppercase leading-tight hover:bg-slate-50 dark:hover:bg-neutral-800 disabled:text-slate-400 dark:disabled:text-neutral-500 whitespace-normal break-words"
                      >
                        Incremental
                      </button>
                      <button
                        onClick={targetMode === 'server' ? () => startServerArchiveJob('repair') : repairArchive}
                        disabled={!accessToken || (targetMode === 'browser' ? !localDir : (!serverTargetRoot || selection.size === 0))}
                        title={targetMode === 'server' ? 'Server-side repair scans the selected OneDrive items and redownloads missing or size-mismatched files.' : 'With a selection, scan selected OneDrive items and redownload missing/unverified local files. Without a selection, use saved repair records.'}
                        className="min-w-0 px-1 py-2 border border-slate-200 dark:border-neutral-700 text-slate-700 dark:text-neutral-100 rounded font-bold text-[9px] uppercase leading-tight hover:bg-slate-50 dark:hover:bg-neutral-800 disabled:text-slate-400 dark:disabled:text-neutral-500 whitespace-normal break-words"
                      >
                        Repair
                      </button>
                    </div>
                  </div>
                ) : (
                  <button 
                    onClick={stopSync}
                    title="Cancel the current archive operation and keep recoverable resume state."
                    className="w-full py-2.5 bg-white dark:bg-neutral-900 border border-slate-300 dark:border-neutral-700 text-slate-700 dark:text-neutral-100 rounded font-bold text-xs uppercase tracking-wider hover:bg-slate-50 dark:hover:bg-neutral-800 transition-all"
                  >
                    Suspend Download
                  </button>
                )}
              </div>
              </div>
            </div>

            <div className="space-y-3">
              <p className="text-[10px] font-bold text-slate-400 dark:text-neutral-500 uppercase tracking-widest">Downloader settings</p>
              <div className="grid grid-cols-2 gap-3">
              <label className="text-[10px] font-bold text-slate-500 dark:text-neutral-400 uppercase">
                Global
                <input
                  type="number"
                  min={1}
                  max={8}
                  value={downloadSettings.maxGlobalConcurrentDownloads}
                  onChange={event => updateDownloadSetting('maxGlobalConcurrentDownloads', Number(event.target.value))}
                  disabled={isSyncing}
                  className="mt-1 w-full border border-slate-200 dark:border-neutral-700 rounded px-2 py-1 text-xs font-normal bg-white dark:bg-neutral-950 text-slate-700 dark:text-neutral-100 disabled:bg-slate-50 dark:disabled:bg-neutral-900 disabled:text-slate-400 dark:disabled:text-neutral-500"
                />
              </label>
              <label className="text-[10px] font-bold text-slate-500 dark:text-neutral-400 uppercase">
                Chunks/file
                <input
                  type="number"
                  min={1}
                  max={6}
                  value={downloadSettings.maxChunksPerFile}
                  onChange={event => updateDownloadSetting('maxChunksPerFile', Number(event.target.value))}
                  disabled={isSyncing}
                  className="mt-1 w-full border border-slate-200 dark:border-neutral-700 rounded px-2 py-1 text-xs font-normal bg-white dark:bg-neutral-950 text-slate-700 dark:text-neutral-100 disabled:bg-slate-50 dark:disabled:bg-neutral-900 disabled:text-slate-400 dark:disabled:text-neutral-500"
                />
              </label>
              <label className="text-[10px] font-bold text-slate-500 dark:text-neutral-400 uppercase">
                Chunk MB
                <input
                  type="number"
                  min={1}
                  max={128}
                  value={Math.round(downloadSettings.chunkSize / (1024 * 1024))}
                  onChange={event => updateDownloadSetting('chunkSize', Number(event.target.value) * 1024 * 1024)}
                  disabled={isSyncing}
                  className="mt-1 w-full border border-slate-200 dark:border-neutral-700 rounded px-2 py-1 text-xs font-normal bg-white dark:bg-neutral-950 text-slate-700 dark:text-neutral-100 disabled:bg-slate-50 dark:disabled:bg-neutral-900 disabled:text-slate-400 dark:disabled:text-neutral-500"
                />
              </label>
              <label className="text-[10px] font-bold text-slate-500 dark:text-neutral-400 uppercase">
                Retries
                <input
                  type="number"
                  min={0}
                  max={10}
                  value={downloadSettings.retryCount}
                  onChange={event => updateDownloadSetting('retryCount', Number(event.target.value))}
                  disabled={isSyncing}
                  className="mt-1 w-full border border-slate-200 dark:border-neutral-700 rounded px-2 py-1 text-xs font-normal bg-white dark:bg-neutral-950 text-slate-700 dark:text-neutral-100 disabled:bg-slate-50 dark:disabled:bg-neutral-900 disabled:text-slate-400 dark:disabled:text-neutral-500"
                />
              </label>
            </div>
            <label className="text-[10px] font-bold text-slate-500 dark:text-neutral-400 uppercase block">
              Conflict
              <select
                value={downloadSettings.conflictStrategy}
                onChange={event => updateDownloadSetting('conflictStrategy', event.target.value as DownloadSettings['conflictStrategy'])}
                disabled={isSyncing}
                className="mt-1 w-full border border-slate-200 dark:border-neutral-700 rounded px-2 py-1 text-xs font-normal bg-white dark:bg-neutral-950 text-slate-700 dark:text-neutral-100 disabled:bg-slate-50 dark:disabled:bg-neutral-900 disabled:text-slate-400 dark:disabled:text-neutral-500"
              >
                <option value="skip_existing">Skip existing</option>
                <option value="overwrite">Overwrite</option>
                <option value="rename_new">Rename new file</option>
                <option value="keep_both">Keep both</option>
                <option value="keep_newest">Keep newest</option>
              </select>
              </label>
              <label className="flex items-center gap-2 text-xs text-slate-600 dark:text-neutral-300">
                <input
                  type="checkbox"
                  checked={downloadSettings.verifyAfterDownload}
                  onChange={event => updateDownloadSetting('verifyAfterDownload', event.target.checked)}
                  disabled={isSyncing}
                />
                Verify after download
              </label>
            </div>
          </div>

          <div className="shrink-0 border-t border-slate-100 dark:border-neutral-800 pt-4 mt-4">
            {activeFiles.length > 0 && (
              <div className="text-xs mb-3 text-slate-600 dark:text-neutral-300 font-medium">
                <div className="text-slate-400 dark:text-neutral-500 uppercase text-[9px] font-bold mb-1">Current files</div>
                <div className="space-y-1.5">
                  {activeFiles.map(file => (
                    <div key={file.id}>
                      <div className="flex justify-between gap-2">
                        <span className="truncate">{file.name}</span>
                        <span className="font-mono text-slate-400 dark:text-neutral-500 shrink-0">{Math.round(file.progress)}%</span>
                      </div>
                      <div className="mt-1 h-1 bg-slate-100 dark:bg-neutral-800 rounded-full overflow-hidden">
                        <div className="h-full bg-blue-600" style={{ width: `${file.progress}%` }}></div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}
            {(isSyncing || isServerBusy) && (
              <div className="flex justify-between text-xs mb-2 text-slate-600 dark:text-neutral-300 font-medium">
                <span className="capitalize">{activeStageLabel}:</span>
                <span>{activeStagePercent}%</span>
              </div>
            )}
            <div className="flex justify-between text-xs mb-2 text-slate-600 dark:text-neutral-300 font-medium">
              <span>Selected Total:</span>
              <span>{formatBytes(targetMode === 'server' && activeServerJob ? activeServerJob.snapshot.totalBytes : syncStats.size)}</span>
            </div>
            <div className="flex justify-between text-xs mb-2 text-slate-600 dark:text-neutral-300 font-medium">
              <span>Downloaded:</span>
              <span>{formatBytes(targetMode === 'server' && activeServerJob ? activeServerJob.snapshot.downloadedBytes : syncStats.downloadedSize)}</span>
            </div>
            <div className="flex justify-between text-xs mb-2 text-slate-600 dark:text-neutral-300 font-medium">
              <span>Items in Queue:</span>
              <span>{remainingItems}</span>
            </div>
            <div className="flex justify-between text-xs text-slate-600 dark:text-neutral-300 font-medium">
              <span>Download Speed:</span>
              <span className="text-blue-600 dark:text-blue-400 italic">{isSyncing || isServerBusy ? formatSpeed(activeSpeed) : '0.00KB/s'}</span>
            </div>
            <div className="flex justify-between text-xs mt-2 text-slate-600 dark:text-neutral-300 font-medium">
              <span>ETA:</span>
              <span>{formatDuration(activeSnapshot?.etaSeconds)}</span>
            </div>
          </div>
        </aside>
      </main>

      {/* Footer: System Status */}
      <footer className="h-10 bg-neutral-900 text-neutral-400 dark:bg-neutral-950 dark:text-neutral-500 px-6 flex items-center justify-between text-[10px] uppercase font-bold tracking-widest shrink-0">
        <div className="flex gap-6">
          <span className="flex items-center gap-2">
            <span className={`w-1.5 h-1.5 rounded-full ${accessToken ? 'bg-green-500' : 'bg-red-500'}`}></span> 
            Microsoft API {accessToken ? 'Connected' : 'Offline'}
          </span>
          <span className="flex items-center gap-2">
            <span className={`w-1.5 h-1.5 rounded-full ${(targetMode === 'browser' ? localDir : serverTargetRoot) ? 'bg-green-500' : 'bg-slate-500'}`}></span> 
            Target {targetMode === 'browser' ? (localDir ? 'Browser Mounted' : 'Unassigned') : (serverTargetRoot ? 'Server Ready' : 'Server Missing')}
          </span>
        </div>
        <div className="flex gap-6">
          <span>Version 1.0.0-Stable</span>
          <span className="text-slate-500 dark:text-neutral-400">{new Date().toLocaleDateString()}</span>
        </div>
      </footer>
    </div>
  );
}
