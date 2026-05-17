import { AlertCircle, RotateCcw } from 'lucide-react';

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
  mode: string;
  status: string;
  targetRoot: string;
  log: string[];
}

interface ActivityLogProps {
  logs: LogEntry[];
  failedFiles: FailedFile[];
  activeServerJob?: ServerArchiveJob;
  isSyncing: boolean;
  canRetry: boolean;
  onRetryFile: (id: string) => void;
}

export function ActivityLog({ logs, failedFiles, activeServerJob, isSyncing, canRetry, onRetryFile }: ActivityLogProps) {
  return (
    <div className="flex-1 border border-slate-200 dark:border-neutral-700 rounded overflow-hidden flex flex-col bg-white dark:bg-neutral-900">
      {activeServerJob && (
        <div className="border-b border-slate-200 dark:border-neutral-700 bg-slate-50 dark:bg-neutral-800 p-4">
          <div className="flex items-center justify-between gap-3 mb-2">
            <div className="min-w-0">
              <div className="text-xs font-bold uppercase tracking-wide text-slate-700 dark:text-neutral-100">Server job</div>
              <div className="text-[10px] text-slate-500 dark:text-neutral-400 break-words">{activeServerJob.targetRoot}</div>
            </div>
            <div className="shrink-0 text-right">
              <div className="text-[9px] font-bold uppercase text-blue-600 dark:text-blue-400">{activeServerJob.status}</div>
              <div className="text-[9px] uppercase text-slate-400 dark:text-neutral-500">{activeServerJob.mode}</div>
            </div>
          </div>
          <div className="max-h-32 overflow-y-auto space-y-1 font-mono text-[10px] text-slate-600 dark:text-neutral-300">
            {activeServerJob.log.length > 0 ? activeServerJob.log.map((line, index) => (
              <div key={`${activeServerJob.id}-${index}`} className="break-words">{line}</div>
            )) : (
              <div className="text-slate-400 dark:text-neutral-500 italic">No server job log entries yet.</div>
            )}
          </div>
        </div>
      )}
      {failedFiles.length > 0 && (
        <div className="border-b border-red-100 dark:border-red-900/60 bg-red-50 dark:bg-red-950/40 p-4">
          <div className="flex items-center gap-2 text-xs font-bold uppercase tracking-wide text-red-700 dark:text-red-300 mb-3">
            <AlertCircle size={14} />
            Failed files
          </div>
          <div className="space-y-2 max-h-36 overflow-y-auto">
            {failedFiles.map(file => (
              <div key={file.id} className="text-xs grid grid-cols-12 gap-3 items-center">
                <div className="col-span-3 font-semibold text-red-900 dark:text-red-200 truncate" title={file.name}>{file.name}</div>
                <div className="col-span-3 text-red-700 dark:text-red-300 truncate" title={file.path}>{file.path}</div>
                <div className="col-span-4 text-red-600 dark:text-red-400 truncate" title={file.reason}>{file.reason}</div>
                <div className="col-span-2 flex justify-end">
                  <button
                    onClick={() => onRetryFile(file.id)}
                    disabled={isSyncing || !canRetry}
                    className="inline-flex h-7 w-7 items-center justify-center rounded border border-red-200 bg-white text-red-700 hover:bg-red-100 disabled:opacity-40 disabled:hover:bg-white dark:border-red-900 dark:bg-neutral-900 dark:text-red-300 dark:hover:bg-red-950"
                    title="Retry this file"
                  >
                    <RotateCcw size={13} />
                  </button>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}
      <div className="grid grid-cols-12 bg-slate-50 dark:bg-neutral-800 border-b border-slate-200 dark:border-neutral-700 p-3 text-[10px] font-bold text-slate-500 dark:text-neutral-400 uppercase tracking-tighter shrink-0">
        <div className="col-span-2">Time</div>
        <div className="col-span-2">Stage</div>
        <div className="col-span-3">File</div>
        <div className="col-span-5">Message</div>
      </div>
      <div className="flex-1 overflow-y-auto min-h-0 font-mono text-[10px]">
        {logs.length > 0 ? logs.map(log => (
          <div key={log.id} className={`grid grid-cols-12 gap-3 px-3 py-2 border-b border-slate-100 dark:border-neutral-800 ${log.stage === 'downloader' && log.message.startsWith('Failed') ? 'bg-red-50 dark:bg-red-950/40 text-red-800 dark:text-red-200' : 'text-slate-600 dark:text-neutral-300'}`}>
            <div className="col-span-2 text-slate-400 dark:text-neutral-500">{log.time}</div>
            <div className="col-span-2 text-blue-600 dark:text-blue-400 uppercase">{log.stage}</div>
            <div className="col-span-3 truncate" title={log.jobPath || log.jobName}>{log.jobName || '-'}</div>
            <div className="col-span-5 truncate" title={log.error || log.message}>{log.message}</div>
          </div>
        )) : (
          <div className="h-full flex items-center justify-center text-slate-300 italic">Awaiting operations...</div>
        )}
      </div>
    </div>
  );
}
