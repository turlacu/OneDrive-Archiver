import type { DownloadJob, DownloadProgressSnapshot, DownloadSummary, Reporter, ReporterEvent } from './types';

export class DownloadReporter {
  private summary: DownloadSummary = {
    downloaded: 0,
    resumed: 0,
    verified: 0,
    skipped: 0,
    failed: 0,
    changedRemotely: 0,
    conflicts: 0,
    insufficientDiskSpace: 0,
  };

  constructor(private readonly reporter: Reporter) {}

  log(stage: string, message: string, job?: DownloadJob) {
    this.reporter({ type: 'log', stage, message, job });
  }

  job(job: DownloadJob) {
    this.reporter({ type: 'job', job });
  }

  increment(key: keyof DownloadSummary) {
    this.summary[key] += 1;
  }

  progress(snapshot: Omit<DownloadProgressSnapshot, 'summary'>) {
    this.reporter({ type: 'progress', snapshot: { ...snapshot, summary: { ...this.summary } } });
  }

  emit(event: ReporterEvent) {
    this.reporter(event);
  }

  getSummary() {
    return { ...this.summary };
  }
}
