import React, { useState, useEffect, useRef, useCallback } from 'react';
import { 
  Cloud, 
  Folder, 
  File, 
  Download, 
  CheckCircle2, 
  AlertCircle, 
  Play, 
  Pause, 
  RotateCcw, 
  Settings,
  LogIn,
  LogOut,
  ChevronRight,
  ChevronDown,
  Github,
  HardDrive,
  Activity,
  Verified
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import axios from 'axios';
import { Client } from '@microsoft/microsoft-graph-client';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

// Types
interface SyncItem {
  id: string;
  name: string;
  size: number;
  path: string;
  type: 'file' | 'folder';
  status: 'pending' | 'downloading' | 'completed' | 'error' | 'skipped' | 'verifying';
  progress: number;
  sha1Hash?: string;
  error?: string;
}

interface LocalHandle {
  id: string;
  name: string;
  handle: FileSystemDirectoryHandle;
}

export default function App() {
  const [accessToken, setAccessToken] = useState<string | null>(() => {
    const fromStorage = localStorage.getItem('ms_token');
    if (fromStorage) return fromStorage;
    
    // Check cookies
    const match = document.cookie.match(/(^| )ms_token=([^;]+)/);
    if (match) return match[2];
    
    return null;
  });
  const [user, setUser] = useState<any>(null);
  const [rootItems, setRootItems] = useState<SyncItem[]>([]);
  const [selection, setSelection] = useState<Set<string>>(new Set());
  const [localDir, setLocalDir] = useState<LocalHandle | null>(null);
  const [isSyncing, setIsSyncing] = useState(false);
  const [syncQueue, setSyncQueue] = useState<SyncItem[]>([]);
  const [syncStats, setSyncStats] = useState({ total: 0, completed: 0, errors: 0, size: 0, downloadedSize: 0 });
  const [logs, setLogs] = useState<string[]>([]);

  const graphClient = useRef<Client | null>(null);
  const syncRef = useRef<boolean>(false);

  // Initialize Graph Client
  useEffect(() => {
    if (accessToken) {
      graphClient.current = Client.init({
        authProvider: (done) => done(null, accessToken)
      });
      fetchUser();
      fetchRootItems();
    }
  }, [accessToken]);

  const fetchUser = async () => {
    try {
      const me = await graphClient.current?.api('/me').get();
      setUser(me);
    } catch (e) {
      console.error('Fetch user error', e);
      if ((e as any).statusCode === 401) handleLogout();
    }
  };

  const fetchRootItems = async (folderId: string = 'root') => {
    try {
      const response = await graphClient.current?.api(`/me/drive/${folderId}/children`).get();
      const items: SyncItem[] = response.value.map((item: any) => ({
        id: item.id,
        name: item.name,
        size: item.size,
        path: item.parentReference.path + '/' + item.name,
        type: item.folder ? 'folder' : 'file',
        status: 'pending',
        progress: 0,
        sha1Hash: item.file?.hashes?.sha1Hash
      }));
      setRootItems(items);
    } catch (e) {
       addLog(`Error fetching items: ${(e as Error).message}`);
    }
  };

  // Global Auth Listeners
  useEffect(() => {
    const onSuccess = (token: string) => {
      setAccessToken(token);
      localStorage.setItem('ms_token', token);
      addLog('Authentication successful');
    };

    // 1. Listen for Storage Events (Best for cross-tab if opener is lost)
    const handleStorage = (e: StorageEvent) => {
      console.log('[AUTH-DEBUG] Storage event:', e.key);
      if ((e.key === 'ms_token_signal' || e.key === 'ms_token') && e.newValue) {
        onSuccess(e.newValue);
        if (e.key === 'ms_token_signal') localStorage.removeItem('ms_token_signal');
      }
    };

    // 2. Poll for token (Fallback for environments where storage events are unreliable)
    const pollInterval = setInterval(async () => {
      const signal = localStorage.getItem('ms_token_signal');
      const token = localStorage.getItem('ms_token');
      const cookieMatch = document.cookie.match(/(^| )ms_token=([^;]+)/);
      const cookieToken = cookieMatch ? cookieMatch[2] : null;

      if (signal) {
        console.log('[AUTH-DEBUG] Token signal detected via polling');
        onSuccess(signal);
        localStorage.removeItem('ms_token_signal');
        return;
      } 
      
      if (!accessToken) {
        // Try server-side session polling
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

        if (token) {
          setAccessToken(token);
        } else if (cookieToken) {
          setAccessToken(cookieToken);
        }
      }
    }, 2000);

    // 3. Listen for BroadcastChannel
    const channel = new BroadcastChannel('onedrive_auth_channel');
    channel.onmessage = (event) => {
      if (event.data?.type === 'OAUTH_AUTH_SUCCESS') {
        onSuccess(event.data.accessToken);
      }
    };

    // 3. Listen for window messages
    const handleMessage = (event: MessageEvent) => {
      if (event.data?.type === 'OAUTH_AUTH_SUCCESS') {
        onSuccess(event.data.accessToken);
      }
    };

    window.addEventListener('storage', handleStorage);
    window.addEventListener('message', handleMessage);

    return () => {
      window.removeEventListener('storage', handleStorage);
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

      // Explicitly clear signal before opening
      localStorage.removeItem('ms_token_signal');

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
    localStorage.removeItem('ms_token');
    setRootItems([]);
  };

  const selectLocalFolder = async () => {
    try {
      const handle = await (window as any).showDirectoryPicker();
      setLocalDir({ id: 'local', name: handle.name, handle });
      addLog(`Local folder selected: ${handle.name}`);
    } catch (e) {
      console.error('Picker error', e);
    }
  };

  const addLog = (message: string) => {
    setLogs(prev => [new Date().toLocaleTimeString() + ': ' + message, ...prev].slice(0, 50));
  };

  const verifyItemIntegrity = async (item: SyncItem, parentHandle: FileSystemDirectoryHandle): Promise<boolean> => {
    try {
      if (item.type === 'folder') {
        const directoryHandle = await parentHandle.getDirectoryHandle(item.name);
        const childrenResponse = await graphClient.current?.api(`/me/drive/items/${item.id}/children`).get();
        for (const child of childrenResponse.value) {
           const childSyncItem: SyncItem = { 
             id: child.id, 
             name: child.name, 
             size: child.size, 
             path: child.parentReference.path + '/' + child.name, 
             type: child.folder ? 'folder' : 'file',
             status: 'pending',
             progress: 0,
             sha1Hash: child.file?.hashes?.sha1Hash 
           };
           await verifyItemIntegrity(childSyncItem, directoryHandle);
        }
        return true;
      }

      updateItemStatus(item.id, 'verifying');
      const fileHandle = await parentHandle.getFileHandle(item.name);
      const file = await fileHandle.getFile();
      
      if (file.size !== item.size) {
        updateItemStatus(item.id, 'error', 'Size mismatch');
        addLog(`Integrity Failure: ${item.name} (Size mismatch)`);
        return false;
      }

      if (item.sha1Hash) {
        const buffer = await file.arrayBuffer();
        const hash = await crypto.subtle.digest('SHA-1', buffer);
        const hashHex = Array.from(new Uint8Array(hash)).map(b => b.toString(16).padStart(2, '0')).join('').toUpperCase();
        
        if (hashHex === item.sha1Hash) {
          updateItemStatus(item.id, 'completed');
          addLog(`Verified: ${item.name}`);
          return true;
        } else {
          updateItemStatus(item.id, 'error', 'Hash mismatch');
          addLog(`Integrity Failure: ${item.name} (Hash mismatch)`);
          return false;
        }
      }
      updateItemStatus(item.id, 'completed');
      return true;
    } catch (e: any) {
      updateItemStatus(item.id, 'error', 'Missing local file');
      return false;
    }
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

    setIsSyncing(true);
    syncRef.current = true;
    
    // Flatten selected items (dfs is better but we'll start with selected roots)
    const itemsToProcess = rootItems.filter(i => selection.has(i.id));
    setSyncQueue(itemsToProcess);
    setSyncStats({ total: itemsToProcess.length, completed: 0, errors: 0, size: itemsToProcess.reduce((a,b) => a+b.size, 0), downloadedSize: 0 });

    for (const item of itemsToProcess) {
      if (!syncRef.current) break;
      await processItem(item, localDir.handle);
    }

    setIsSyncing(false);
    syncRef.current = false;
    addLog('Sync complete');
  };

  const processItem = async (item: SyncItem, parentHandle: FileSystemDirectoryHandle) => {
    try {
      updateItemStatus(item.id, 'downloading');
      
      if (item.type === 'folder') {
        const directoryHandle = await parentHandle.getDirectoryHandle(item.name, { create: true });
        const childrenResponse = await graphClient.current?.api(`/me/drive/items/${item.id}/children`).get();
        const children = childrenResponse.value;
        
        for (const child of children) {
          if (!syncRef.current) break;
          const childSyncItem: SyncItem = {
            id: child.id,
            name: child.name,
            size: child.size,
            path: child.parentReference.path + '/' + child.name,
            type: child.folder ? 'folder' : 'file',
            status: 'pending',
            progress: 0,
            sha1Hash: child.file?.hashes?.sha1Hash
          };
          await processItem(childSyncItem, directoryHandle);
        }
      } else {
        await downloadFile(item, parentHandle);
      }
      
      updateItemStatus(item.id, 'completed');
      setSyncStats(prev => ({ ...prev, completed: prev.completed + 1 }));
    } catch (e) {
      updateItemStatus(item.id, 'error', (e as Error).message);
      setSyncStats(prev => ({ ...prev, errors: prev.errors + 1 }));
      addLog(`Error processing ${item.name}: ${(e as Error).message}`);
    }
  };

  const downloadFile = async (item: SyncItem, parentHandle: FileSystemDirectoryHandle) => {
    // Integrity Check Replacement / Resumable Logic
    let existingFile: FileSystemFileHandle | null = null;
    try {
      existingFile = await parentHandle.getFileHandle(item.name);
      
      // Verify integrity if already exists
      if (item.sha1Hash) {
        updateItemStatus(item.id, 'verifying');
        const file = await existingFile.getFile();
        const buffer = await file.arrayBuffer();
        const hash = await crypto.subtle.digest('SHA-1', buffer);
        const hashHex = Array.from(new Uint8Array(hash)).map(b => b.toString(16).padStart(2, '0')).join('').toUpperCase();
        
        if (hashHex === item.sha1Hash) {
          addLog(`Skipped (Already exists & Valid): ${item.name}`);
          updateItemStatus(item.id, 'skipped');
          return;
        } else {
          addLog(`Hash mismatch for ${item.name}. Re-downloading.`);
        }
      }
    } catch (e) {
      // File doesn't exist, proceed to download
    }

    const fileHandle = await parentHandle.getFileHandle(item.name, { create: true });
    const writable = await fileHandle.createWritable();
    
    const downloadUrl = await graphClient.current?.api(`/me/drive/items/${item.id}/content`).get();
    // Graph SDK returns the raw stream if we ask for it correctly, or we can fetch the @microsoft.graph.downloadUrl
    const itemFull = await graphClient.current?.api(`/me/drive/items/${item.id}`).select('id,@microsoft.graph.downloadUrl').get();
    const url = itemFull['@microsoft.graph.downloadUrl'];

    const response = await fetch(url);
    const reader = response.body?.getReader();
    if (!reader) throw new Error('Failed to get download reader');

    let receivedLength = 0;
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      await writable.write(value);
      receivedLength += value.length;
      updateItemProgress(item.id, (receivedLength / item.size) * 100);
    }
    await writable.close();
    addLog(`Downloaded: ${item.name}`);
  };

  const updateItemStatus = (id: string, status: SyncItem['status'], error?: string) => {
    // This is expensive for large sets, but for root it's fine.
    setRootItems(prev => prev.map(i => i.id === id ? { ...i, status, error } : i));
  };

  const updateItemProgress = (id: string, progress: number) => {
    setRootItems(prev => prev.map(i => i.id === id ? { ...i, progress } : i));
  };

  const stopSync = () => {
    syncRef.current = false;
    setIsSyncing(false);
    addLog('Sync cancelled by user');
  };

  return (
    <div className="flex flex-col h-screen w-full bg-slate-50 text-slate-900 font-sans overflow-hidden">
      {/* Header: Geometric Alignment */}
      <header className="h-16 bg-white border-b border-slate-200 px-6 flex items-center justify-between shadow-sm shrink-0">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 bg-blue-600 rounded flex items-center justify-center">
            <Cloud className="text-white" size={20} />
          </div>
          <h1 className="text-xl font-bold tracking-tight text-slate-800 uppercase">
            SyncPoint <span className="text-blue-600 font-light">Archiver</span>
          </h1>
        </div>
        
        <div className="flex items-center gap-4">
          {user ? (
            <div className="flex items-center gap-3 px-4 py-1.5 bg-slate-100 border border-slate-200 rounded-md">
              <div className="w-6 h-6 bg-blue-200 rounded-full flex items-center justify-center">
                <span className="text-[10px] font-bold text-blue-700">{user.displayName?.charAt(0)}</span>
              </div>
              <span className="text-sm font-medium">{user.displayName || user.userPrincipalName}</span>
              <button onClick={handleLogout} className="ml-2 text-slate-400 hover:text-red-500 transition-colors">
                <LogOut size={14} />
              </button>
            </div>
          ) : (
            <button 
              onClick={handleLogin}
              className="px-6 py-2 bg-blue-600 text-white text-sm font-semibold rounded hover:bg-blue-700 transition-all flex items-center gap-2"
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
        <nav className="w-64 border-r border-slate-200 bg-white p-6 flex flex-col gap-8 shrink-0">
          <div>
            <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-4">Storage Source</p>
            <ul className="space-y-1">
              <li className="px-3 py-2 bg-blue-50 text-blue-700 rounded border-r-4 border-blue-600 font-medium text-sm cursor-pointer">
                OneDrive Personal
              </li>
              <li className="px-3 py-2 text-slate-500 rounded hover:bg-slate-50 font-medium text-sm cursor-not-allowed opacity-50">
                SharePoint Sites
              </li>
              <li className="px-3 py-2 text-slate-500 rounded hover:bg-slate-50 font-medium text-sm cursor-not-allowed opacity-50">
                Shared with Me
              </li>
            </ul>
          </div>
          
          <div>
            <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-4">Local Target</p>
            <div 
              onClick={selectLocalFolder}
              className={`p-4 border-2 border-dashed rounded-lg cursor-pointer transition-all ${localDir ? 'border-blue-600 bg-blue-50/30' : 'border-slate-200 hover:border-blue-400'}`}
            >
              <div className="flex items-center gap-2 mb-1">
                <HardDrive size={16} className={localDir ? 'text-blue-600' : 'text-slate-400'} />
                <span className="text-xs font-bold truncate">{localDir ? localDir.name : 'Choose Folder'}</span>
              </div>
              <p className="text-[10px] text-slate-500 leading-tight">
                {localDir ? 'Drive Mapped & Ready' : 'Select sync destination'}
              </p>
            </div>
          </div>

          <div className="mt-auto p-4 bg-slate-50 rounded-lg border border-slate-100">
            <p className="text-xs text-slate-500 leading-tight">
              Status: <br/>
              <span className="font-bold text-slate-700 uppercase">
                {isSyncing ? 'Actively Syncing' : 'Idle / Ready'}
              </span>
            </p>
          </div>
        </nav>

        {/* Center: File Explorer Selection */}
        <section className="flex-1 bg-white flex flex-col overflow-hidden">
          <div className="p-6 border-b border-slate-100 flex justify-between items-center bg-slate-50/50">
            <h2 className="text-sm font-bold text-slate-700 uppercase tracking-wide">OneDrive explorer</h2>
            <div className="flex gap-2">
              <button 
                onClick={() => setSelection(new Set(rootItems.map(i => i.id)))}
                className="px-3 py-1 border border-slate-200 text-xs font-semibold rounded hover:bg-white transition-colors"
                disabled={!accessToken}
              >
                Select All
              </button>
              <button 
                onClick={() => setSelection(new Set())}
                className="px-3 py-1 border border-slate-200 text-xs font-semibold rounded hover:bg-white transition-colors"
                disabled={!accessToken}
              >
                Deselect
              </button>
            </div>
          </div>
          
          <div className="flex-1 overflow-hidden p-6 flex flex-col">
            <div className="flex-1 border border-slate-200 rounded overflow-hidden flex flex-col">
              <div className="grid grid-cols-12 bg-slate-50 border-b border-slate-200 p-3 text-[10px] font-bold text-slate-500 uppercase tracking-tighter shrink-0">
                <div className="col-span-1"></div>
                <div className="col-span-6">Name</div>
                <div className="col-span-2">Size</div>
                <div className="col-span-3 text-right">Status</div>
              </div>
              
              <div className="flex-1 overflow-y-auto min-h-0 bg-white">
                {!accessToken ? (
                  <div className="h-full flex flex-col items-center justify-center p-12 text-center text-slate-400">
                    <Cloud size={48} className="opacity-10 mb-4" />
                    <p className="text-sm font-medium">Please sign in to browse cloud storage</p>
                  </div>
                ) : rootItems.length === 0 ? (
                  <div className="p-8 space-y-4">
                    {[1,2,3,4,5].map(i => (
                      <div key={i} className="h-4 bg-slate-100 animate-pulse rounded w-full"></div>
                    ))}
                  </div>
                ) : (
                  rootItems.map((item) => (
                    <div 
                      key={item.id}
                      onClick={() => !isSyncing && toggleSelection(item.id)}
                      className={`grid grid-cols-12 p-3 border-b border-slate-100 items-center text-sm transition-colors cursor-pointer hover:bg-slate-50 ${selection.has(item.id) ? 'bg-blue-50/50' : ''}`}
                    >
                      <div className="col-span-1 flex justify-center">
                        <input 
                          type="checkbox" 
                          checked={selection.has(item.id)} 
                          onChange={() => {}} 
                          className="w-4 h-4 rounded border-slate-300 text-blue-600 focus:ring-blue-500"
                        />
                      </div>
                      <div className="col-span-6 flex items-center gap-3">
                        {item.type === 'folder' ? (
                          <Folder size={16} className="text-blue-500" />
                        ) : (
                          <File size={16} className="text-slate-400" />
                        )}
                        <span className={`font-medium ${selection.has(item.id) ? 'text-blue-700' : 'text-slate-700'} truncate`}>
                          {item.name}
                        </span>
                      </div>
                      <div className="col-span-2 text-slate-500 font-mono text-xs">
                        {(item.size / (1024*1024)).toFixed(2)} MB
                      </div>
                      <div className="col-span-3 flex justify-end items-center gap-2">
                        {item.status === 'completed' && <Verified size={16} className="text-green-500" />}
                        {item.status === 'skipped' && <CheckCircle2 size={16} className="text-blue-400" />}
                        {item.status === 'verifying' && <RotateCcw size={16} className="animate-spin text-blue-500" />}
                        {item.status === 'downloading' && (
                          <div className="flex flex-col items-end gap-1 w-20">
                            <div className="h-1 w-full bg-slate-200 rounded-full overflow-hidden">
                              <div className="h-full bg-blue-600" style={{ width: `${item.progress}%` }}></div>
                            </div>
                            <span className="text-[8px] font-mono opacity-50">{Math.round(item.progress)}%</span>
                          </div>
                        )}
                        {item.status === 'error' && <AlertCircle size={16} className="text-red-500" />}
                        {item.status === 'pending' && selection.has(item.id) && <span className="text-[10px] text-slate-400 italic">Queued</span>}
                      </div>
                    </div>
                  )
                ))}
              </div>
            </div>
          </div>

          {/* Console / Log Footer */}
          <div className="h-32 bg-slate-900 text-slate-300 p-4 font-mono text-[10px] overflow-y-auto shrink-0 border-t border-slate-800">
            <div className="flex items-center gap-2 mb-2 text-slate-500 uppercase tracking-tighter font-bold">
              <Activity size={10} />
              <span>Real-time Sync Activity Log</span>
            </div>
            {logs.length > 0 ? logs.map((log, i) => (
              <div key={i} className="mb-0.5 leading-relaxed flex gap-2">
                <span className="opacity-30">[{new Date().toLocaleTimeString()}]</span>
                <span className="text-blue-400">SYNC_INF:</span>
                <span>{log}</span>
              </div>
            )) : (
              <p className="opacity-20 italic">Awaiting operations...</p>
            )}
          </div>
        </section>

        {/* Right Sidebar: Controls & Progress */}
        <aside className="w-80 border-l border-slate-200 bg-white p-6 flex flex-col gap-6 shadow-inner shrink-0">
          <div className="space-y-4">
            <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Task metrics</p>
            <div className="grid grid-cols-2 gap-3">
              <div className="p-3 bg-slate-50 border border-slate-200 rounded">
                <p className="text-[8px] font-bold text-slate-400 uppercase mb-1">Completed</p>
                <div className="text-xl font-bold text-slate-800">{syncStats.completed}</div>
              </div>
              <div className="p-3 bg-slate-50 border border-slate-200 rounded">
                <p className="text-[8px] font-bold text-slate-400 uppercase mb-1">Failures</p>
                <div className="text-xl font-bold text-red-600">{syncStats.errors}</div>
              </div>
            </div>
          </div>

          <div className="space-y-4">
            <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Global Progress</p>
            <div className="p-4 border border-slate-200 rounded bg-white shadow-sm">
              <div className="flex justify-between items-center mb-3">
                <span className={`text-[10px] font-bold uppercase flex items-center gap-1.5 ${isSyncing ? 'text-blue-600' : 'text-slate-400'}`}>
                  <span className={`w-2 h-2 rounded-full ${isSyncing ? 'bg-blue-500 animate-pulse' : 'bg-slate-300'}`}></span>
                  {isSyncing ? 'Actively Syncing' : 'Idle'}
                </span>
                <span className="text-xs font-mono font-bold">
                  {Math.round((syncStats.completed / (syncStats.total || 1)) * 100)}%
                </span>
              </div>
              
              <div className="h-1.5 w-full bg-slate-100 rounded-full overflow-hidden mb-4">
                <motion.div 
                  className="h-full bg-blue-600"
                  initial={{ width: 0 }}
                  animate={{ width: `${(syncStats.completed / (syncStats.total || 1)) * 100}%` }}
                />
              </div>

              <div className="space-y-2">
                {!isSyncing ? (
                  <button 
                    onClick={startSync}
                    disabled={!accessToken || !localDir || selection.size === 0}
                    className="w-full py-2.5 bg-blue-600 text-white rounded font-bold text-xs uppercase tracking-wider shadow-lg shadow-blue-100 hover:bg-blue-700 disabled:opacity-30 disabled:shadow-none transition-all"
                  >
                    {syncStats.completed > 0 && syncStats.completed < syncStats.total ? 'Resume Sync Process' : 'Start Sync Process'}
                  </button>
                ) : (
                  <button 
                    onClick={stopSync}
                    className="w-full py-2.5 bg-white border border-slate-300 text-slate-700 rounded font-bold text-xs uppercase tracking-wider hover:bg-slate-50 transition-all"
                  >
                    Suspend Download
                  </button>
                )}
                <button 
                  onClick={async () => {
                    if (!localDir) return;
                    setIsSyncing(true);
                    syncRef.current = true;
                    addLog('Starting Integrity Verification...');
                    const items = rootItems.filter(i => selection.has(i.id));
                    for (const item of items) {
                      if (!syncRef.current) break;
                      await verifyItemIntegrity(item, localDir.handle);
                    }
                    setIsSyncing(false);
                    syncRef.current = false;
                    addLog('Integrity check complete');
                  }}
                  disabled={!localDir || selection.size === 0 || isSyncing}
                  className="w-full py-2.5 bg-slate-50 border border-slate-200 text-slate-700 rounded font-bold text-xs uppercase tracking-wider hover:bg-slate-100 transition-all disabled:opacity-30"
                >
                  Verify Integrity (SHA-1)
                </button>
              </div>
            </div>
          </div>

          <div className="mt-auto border-t border-slate-100 pt-6">
            <div className="flex justify-between text-xs mb-2 text-slate-600 font-medium">
              <span>Selected Total:</span>
              <span>{(syncStats.size / (1024*1024)).toFixed(1)} MB</span>
            </div>
            <div className="flex justify-between text-xs mb-2 text-slate-600 font-medium">
              <span>Items in Queue:</span>
              <span>{syncStats.total - syncStats.completed}</span>
            </div>
            <div className="flex justify-between text-xs text-slate-600 font-medium">
              <span>Download Speed:</span>
              <span className="text-blue-600 italic">{isSyncing ? 'Optimizing...' : '0 KB/s'}</span>
            </div>
          </div>
        </aside>
      </main>

      {/* Footer: System Status */}
      <footer className="h-10 bg-slate-800 text-slate-400 px-6 flex items-center justify-between text-[10px] uppercase font-bold tracking-widest shrink-0">
        <div className="flex gap-6">
          <span className="flex items-center gap-2">
            <span className={`w-1.5 h-1.5 rounded-full ${accessToken ? 'bg-green-500' : 'bg-red-500'}`}></span> 
            Microsoft API {accessToken ? 'Connected' : 'Offline'}
          </span>
          <span className="flex items-center gap-2">
            <span className={`w-1.5 h-1.5 rounded-full ${localDir ? 'bg-green-500' : 'bg-slate-500'}`}></span> 
            Target Volume {localDir ? 'Mounted' : 'Unassigned'}
          </span>
        </div>
        <div className="flex gap-6">
          <span>Version 1.0.0-Stable</span>
          <span className="text-slate-500">{new Date().toLocaleDateString()}</span>
        </div>
      </footer>
    </div>
  );
}
