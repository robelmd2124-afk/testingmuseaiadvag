/**
 * SyncLogger Utility
 * Tracks synchronization success/failure history, operation latency, and state transitions.
 * Persists history to localStorage and renders the last N logs to help diagnose state inconsistencies.
 */

export type SyncLogStatus = 'SUCCESS' | 'FAILURE' | 'REVERTED' | 'INFO' | 'WARNING';

export type SyncLogType =
  | 'SHEET_FETCH'
  | 'SHEET_PUSH'
  | 'APPS_SCRIPT_SYNC'
  | 'RPC_READ'
  | 'RPC_WRITE'
  | 'DELETE_ROW'
  | 'UPDATE_ROW'
  | 'WEBHOOK_TEST'
  | 'LOCAL_CACHE'
  | 'STATE_REVERT';

export interface SyncLogEntry {
  id: string;
  timestamp: string; // ISO string
  timeFormatted: string; // Human readable time e.g. 02:45:10 PM BST
  type: SyncLogType;
  status: SyncLogStatus;
  message: string;
  details?: Record<string, any> | string;
  durationMs?: number;
}

const STORAGE_KEY = 'adverse_sync_logs_v1';
const MAX_STORED_LOGS = 30;

class SyncLoggerService {
  private logs: SyncLogEntry[] = [];

  constructor() {
    this.loadFromStorage();
  }

  private loadFromStorage() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) {
          this.logs = parsed.slice(0, MAX_STORED_LOGS);
        }
      }
    } catch (e) {
      this.logs = [];
    }
  }

  private saveToStorage() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(this.logs.slice(0, MAX_STORED_LOGS)));
    } catch (e) {
      // Ignore quota errors
    }
  }

  private formatTime(d: Date = new Date()): string {
    try {
      return new Intl.DateTimeFormat('en-US', {
        timeZone: 'Asia/Dhaka',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: true
      }).format(d) + ' BST';
    } catch (_) {
      return d.toLocaleTimeString();
    }
  }

  public log(entry: Omit<SyncLogEntry, 'id' | 'timestamp' | 'timeFormatted'>): SyncLogEntry {
    const now = new Date();
    const fullEntry: SyncLogEntry = {
      id: 'log_' + Date.now() + '_' + Math.random().toString(36).substring(2, 7),
      timestamp: now.toISOString(),
      timeFormatted: this.formatTime(now),
      ...entry
    };

    this.logs.unshift(fullEntry);
    if (this.logs.length > 50) {
      this.logs = this.logs.slice(0, 50);
    }
    this.saveToStorage();

    // Trigger update if container exists in DOM
    const container = document.getElementById('syncLogsContainer');
    if (container) {
      this.renderLogsToElement(container, 5);
    }

    return fullEntry;
  }

  public logSuccess(
    type: SyncLogType,
    message: string,
    details?: Record<string, any> | string,
    durationMs?: number
  ): SyncLogEntry {
    return this.log({
      type,
      status: 'SUCCESS',
      message,
      details,
      durationMs
    });
  }

  public logFailure(
    type: SyncLogType,
    error: any,
    details?: Record<string, any> | string,
    durationMs?: number
  ): SyncLogEntry {
    const errorMsg = error instanceof Error ? error.message : String(error?.message || error || 'Unknown sync error');
    return this.log({
      type,
      status: 'FAILURE',
      message: errorMsg,
      details,
      durationMs
    });
  }

  public logReverted(
    type: SyncLogType,
    reason: string = 'Reverted to last known state to prevent 0 or empty values',
    details?: Record<string, any> | string
  ): SyncLogEntry {
    return this.log({
      type,
      status: 'REVERTED',
      message: `⚠️ ${reason}`,
      details
    });
  }

  public logWarning(
    type: SyncLogType,
    message: string,
    details?: Record<string, any> | string
  ): SyncLogEntry {
    return this.log({
      type,
      status: 'WARNING',
      message,
      details
    });
  }

  public getRecentLogs(limit: number = 5): SyncLogEntry[] {
    return this.logs.slice(0, limit);
  }

  public getAllLogs(): SyncLogEntry[] {
    return [...this.logs];
  }

  public clearLogs() {
    this.logs = [];
    try {
      localStorage.removeItem(STORAGE_KEY);
    } catch (e) {}

    const container = document.getElementById('syncLogsContainer');
    if (container) {
      this.renderLogsToElement(container, 5);
    }
  }

  /**
   * Generates a clean HTML representation for the last N sync log entries
   */
  public renderLogsHtml(limit: number = 5): string {
    const recent = this.getRecentLogs(limit);

    if (!recent || recent.length === 0) {
      return `
        <div style="padding: 14px 16px; text-align: center; color: var(--txt-muted); font-size: 12px; background: var(--surface-hover); border-radius: 8px; border: 1px dashed var(--bor);">
          <span class="ms" style="font-size: 18px; vertical-align: middle; opacity: 0.6; margin-right: 4px;">history</span>
          No sync operations recorded in this session yet.
        </div>
      `;
    }

    const rowsHtml = recent.map((item) => {
      let badgeBg = 'rgba(16, 185, 129, 0.15)';
      let badgeColor = '#10B981';
      let icon = 'check_circle';
      let label = 'SUCCESS';

      if (item.status === 'FAILURE') {
        badgeBg = 'rgba(239, 68, 68, 0.15)';
        badgeColor = '#EF4444';
        icon = 'cancel';
        label = 'FAILED';
      } else if (item.status === 'REVERTED') {
        badgeBg = 'rgba(245, 158, 11, 0.18)';
        badgeColor = '#D97706';
        icon = 'history';
        label = 'REVERTED SAFE';
      } else if (item.status === 'WARNING') {
        badgeBg = 'rgba(245, 158, 11, 0.15)';
        badgeColor = '#F59E0B';
        icon = 'warning';
        label = 'WARNING';
      }

      let detailsText = '';
      if (item.details) {
        if (typeof item.details === 'string') {
          detailsText = item.details;
        } else {
          try {
            detailsText = Object.entries(item.details)
              .map(([k, v]) => `${k}: ${typeof v === 'object' ? JSON.stringify(v) : v}`)
              .join(' • ');
          } catch (_) {
            detailsText = '';
          }
        }
      }

      const durationTag = item.durationMs ? `<span style="font-size:10.5px;color:var(--txt-muted);margin-left:6px">(${item.durationMs}ms)</span>` : '';

      return `
        <div style="display: flex; flex-direction: column; gap: 4px; padding: 9px 12px; background: var(--surface-hover); border: 1px solid var(--bor); border-radius: 8px; font-size: 12px; transition: background 0.15s ease;">
          <div style="display: flex; justify-content: space-between; align-items: center; gap: 8px;">
            <div style="display: flex; align-items: center; gap: 6px; font-weight: 700;">
              <span class="ms" style="font-size: 15px; color: ${badgeColor}">${icon}</span>
              <span style="display: inline-block; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: 800; text-transform: uppercase; background: ${badgeBg}; color: ${badgeColor}">
                ${label}
              </span>
              <span style="font-size: 11px; font-weight: 700; color: var(--pri); background: rgba(99,102,241,0.08); padding: 1px 5px; border-radius: 4px;">
                ${item.type}
              </span>
              ${durationTag}
            </div>
            <span style="font-size: 11px; color: var(--txt-muted); font-family: monospace; white-space: nowrap;">
              ${item.timeFormatted}
            </span>
          </div>
          <div style="color: var(--txt); font-weight: 600; padding-left: 21px; word-break: break-word;">
            ${item.message}
          </div>
          ${detailsText ? `<div style="color: var(--txt-muted); font-size: 11px; padding-left: 21px; word-break: break-word;">${detailsText}</div>` : ''}
        </div>
      `;
    }).join('');

    return `
      <div style="display: flex; flex-direction: column; gap: 6px;">
        ${rowsHtml}
      </div>
    `;
  }

  /**
   * Directly mounts or updates the HTML in a DOM container
   */
  public renderLogsToElement(containerEl: HTMLElement, limit: number = 5) {
    if (!containerEl) return;
    containerEl.innerHTML = this.renderLogsHtml(limit);
  }
}

export const SyncLogger = new SyncLoggerService();

// Expose globally to window for Vanilla JS and React environments
if (typeof window !== 'undefined') {
  (window as any).SyncLogger = SyncLogger;
}

export default SyncLogger;
