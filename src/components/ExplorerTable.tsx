import { AlertCircle, CheckCircle2, ChevronRight, Cloud, File, Folder, RotateCcw, Verified } from 'lucide-react';

export interface ExplorerItem {
  id: string;
  name: string;
  size: number;
  type: 'file' | 'folder';
  status: 'pending' | 'scanning' | 'downloading' | 'completed' | 'error' | 'skipped' | 'verifying' | 'queued' | 'paused' | 'retrying' | 'throttled' | 'failed' | 'stale_remote_changed' | 'insufficient_space' | 'cancelled' | 'interrupted';
  progress: number;
}

interface ExplorerTableProps {
  accessToken: string | null;
  isLoadingItems: boolean;
  isSyncing: boolean;
  itemsError: string | null;
  rootItemsLength: number;
  emptyMessage: string;
  visibleItems: ExplorerItem[];
  selection: Set<string>;
  rowHeight: number;
  topSpacerHeight: number;
  bottomSpacerHeight: number;
  onScroll: (scrollTop: number) => void;
  onToggleSelection: (id: string) => void;
  onOpenFolder: (item: ExplorerItem) => void;
  onRetry: () => void;
}

export function ExplorerTable({
  accessToken,
  isLoadingItems,
  isSyncing,
  itemsError,
  rootItemsLength,
  emptyMessage,
  visibleItems,
  selection,
  rowHeight,
  topSpacerHeight,
  bottomSpacerHeight,
  onScroll,
  onToggleSelection,
  onOpenFolder,
  onRetry,
}: ExplorerTableProps) {
  return (
    <div className="flex-1 border border-slate-200 dark:border-neutral-700 rounded overflow-hidden flex flex-col">
      <div className="grid grid-cols-12 bg-slate-50 dark:bg-neutral-800 border-b border-slate-200 dark:border-neutral-700 p-3 text-[10px] font-bold text-slate-500 dark:text-neutral-400 uppercase tracking-tighter shrink-0">
        <div className="col-span-1"></div>
        <div className="col-span-6">Name</div>
        <div className="col-span-2">Size</div>
        <div className="col-span-3 text-right">Status</div>
      </div>

      <div
        className="flex-1 overflow-y-auto min-h-0 bg-white dark:bg-neutral-900"
        onScroll={event => onScroll(event.currentTarget.scrollTop)}
      >
        {!accessToken ? (
          <div className="h-full flex flex-col items-center justify-center p-12 text-center text-slate-400 dark:text-neutral-500">
            <Cloud size={48} className="opacity-10 mb-4" />
            <p className="text-sm font-medium">Please sign in to browse cloud storage</p>
          </div>
        ) : isLoadingItems && rootItemsLength === 0 ? (
          <div className="p-8 space-y-4">
            {[1, 2, 3, 4, 5].map(i => (
              <div key={i} className="h-4 bg-slate-100 dark:bg-neutral-800 animate-pulse rounded w-full"></div>
            ))}
          </div>
        ) : itemsError ? (
          <div className="h-full flex flex-col items-center justify-center p-12 text-center text-slate-500 dark:text-neutral-400">
            <AlertCircle size={40} className="text-red-500 mb-4" />
            <p className="text-sm font-bold text-slate-700 dark:text-neutral-100 mb-2">Could not load OneDrive files</p>
            <p className="text-xs max-w-md leading-relaxed">{itemsError}</p>
            <button
              onClick={onRetry}
              className="mt-5 px-4 py-2 bg-blue-600 text-white rounded text-xs font-bold uppercase tracking-wide hover:bg-blue-700"
            >
              Retry
            </button>
          </div>
        ) : rootItemsLength === 0 ? (
          <div className="h-full flex flex-col items-center justify-center p-12 text-center text-slate-400 dark:text-neutral-500">
            <Folder size={48} className="opacity-10 mb-4" />
            <p className="text-sm font-medium">{emptyMessage}</p>
          </div>
        ) : (
          <>
            <div style={{ height: topSpacerHeight }} />
            {visibleItems.map(item => (
              <div
                key={item.id}
                onClick={() => !isSyncing && onToggleSelection(item.id)}
                onDoubleClick={() => onOpenFolder(item)}
                className={`grid grid-cols-12 px-3 border-b border-slate-100 dark:border-neutral-800 items-center text-sm transition-colors cursor-pointer hover:bg-slate-50 dark:hover:bg-neutral-800/80 ${selection.has(item.id) ? 'bg-blue-50 dark:bg-neutral-800' : ''}`}
                style={{ height: rowHeight }}
              >
                <div className="col-span-1 flex justify-center">
                  <input
                    type="checkbox"
                    checked={selection.has(item.id)}
                    onChange={() => {}}
                    className="w-4 h-4 rounded border-slate-300 text-blue-600 dark:text-blue-400 focus:ring-blue-500"
                  />
                </div>
                <div className="col-span-6 flex items-center gap-3">
                  {item.type === 'folder' ? (
                    <Folder size={16} className="text-blue-500 dark:text-sky-300" />
                  ) : (
                    <File size={16} className="text-slate-400 dark:text-neutral-500" />
                  )}
                  <span className={`font-medium ${selection.has(item.id) ? 'text-blue-700 dark:text-neutral-100' : 'text-slate-700 dark:text-neutral-200'} truncate`}>
                    {item.name}
                  </span>
                  {item.type === 'folder' && (
                    <button
                      onClick={event => {
                        event.stopPropagation();
                        onOpenFolder(item);
                      }}
                      disabled={isSyncing}
                      className="ml-auto text-slate-300 dark:text-neutral-500 hover:text-blue-600 dark:hover:text-sky-300 disabled:hover:text-slate-300 dark:disabled:hover:text-neutral-500"
                      title="Open folder"
                    >
                      <ChevronRight size={14} />
                    </button>
                  )}
                </div>
                <div className="col-span-2 text-slate-500 dark:text-neutral-400 font-mono text-xs">
                  {(item.size / (1024 * 1024)).toFixed(2)} MB
                </div>
                <div className="col-span-3 flex justify-end items-center gap-2">
                  {item.status === 'completed' && <Verified size={16} className="text-green-500" />}
                  {item.status === 'skipped' && <CheckCircle2 size={16} className="text-blue-400" />}
                  {item.status === 'verifying' && <RotateCcw size={16} className="animate-spin text-blue-500" />}
                  {(item.status === 'retrying' || item.status === 'throttled') && <RotateCcw size={16} className="animate-spin text-amber-500" />}
                  {item.status === 'downloading' && (
                    <div className="flex flex-col items-end gap-1 w-20">
                      <div className="h-1 w-full bg-slate-200 rounded-full overflow-hidden">
                        <div className="h-full bg-blue-600" style={{ width: `${item.progress}%` }}></div>
                      </div>
                      <span className="text-[8px] font-mono opacity-50">{Math.round(item.progress)}%</span>
                    </div>
                  )}
                  {(item.status === 'error' || item.status === 'failed' || item.status === 'stale_remote_changed') && <AlertCircle size={16} className="text-red-500" />}
                  {item.status === 'pending' && selection.has(item.id) && <span className="text-[10px] text-slate-400 dark:text-neutral-500 italic">Queued</span>}
                  {(item.status === 'queued' || item.status === 'retrying' || item.status === 'throttled') && (
                    <span className="text-[10px] text-slate-400 dark:text-neutral-500 italic capitalize">{item.status.replaceAll('_', ' ')}</span>
                  )}
                </div>
              </div>
            ))}
            <div style={{ height: bottomSpacerHeight }} />
            {isLoadingItems && (
              <div className="p-3 text-center text-[10px] font-bold uppercase tracking-widest text-slate-400 dark:text-neutral-500 border-t border-slate-100 dark:border-neutral-800">
                Loading more items...
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}
