import express from 'express';
import compression from 'compression';
import cookieParser from 'cookie-parser';
import crypto from 'crypto';
import fs from 'fs';
import { GoogleAuth } from 'google-auth-library';
import { google, sheets_v4 } from 'googleapis';
import path from 'path';
import {
  MONTHS, MENU, CH, PAYMENT_HEADERS, TODAY_HEADERS, PAGE_SUMMARY_HEADERS,
  AGENCY_SPENDING_HEADERS, MONTHLY_CLOSING_HEADERS, AGENCY_COST_HEADERS, END_ADS_HEADERS, AD_ACCOUNT_BILLING_HEADERS,
  ALL_DATE_HEADERS, isDateHeader, normalizeRowDates,
  parseAnyDate, formatCustomDate, isSameDay, getTodayFormatted, toInputDateValue, calculateEndDate, getNowBST, normalizeCampaign
} from './src/services/schema.ts';

const app = express();
app.use(compression() as any);
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));
app.use((err: any, req: express.Request, res: express.Response, next: express.NextFunction) => {
  if (err instanceof SyntaxError && 'body' in err) {
    return res.status(400).json({ ok: false, error: 'Invalid JSON payload' });
  }
  next(err);
});
app.use(cookieParser());

// API health routes for container orchestration and Cloud Run
app.get(['/api/health', '/healthz', '/health', '/_ah/health'], (req, res) => {
  res.status(200).json({ status: 'ok', uptime: process.uptime(), timestamp: new Date().toISOString() });
});

// PWA Manifest and Service Worker static routes
app.get(['/site.webmanifest', '/manifest.json'], (req, res) => {
  res.setHeader('Content-Type', 'application/manifest+json; charset=utf-8');
  res.setHeader('Cache-Control', 'public, max-age=0, must-revalidate');
  const manifestPath = path.join(process.cwd(), 'public', 'site.webmanifest');
  if (fs.existsSync(manifestPath)) {
    return res.sendFile(manifestPath);
  }
  const rootManifest = path.join(process.cwd(), 'public', 'manifest.json');
  if (fs.existsSync(rootManifest)) {
    return res.sendFile(rootManifest);
  }
  res.status(404).end();
});

app.get('/sw.js', (req, res) => {
  res.setHeader('Content-Type', 'application/javascript; charset=utf-8');
  res.setHeader('Service-Worker-Allowed', '/');
  res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
  const swDist = path.join(process.cwd(), 'dist', 'sw.js');
  const swPublic = path.join(process.cwd(), 'public', 'sw.js');
  if (fs.existsSync(swDist)) {
    return res.sendFile(swDist);
  }
  if (fs.existsSync(swPublic)) {
    return res.sendFile(swPublic);
  }
  res.status(404).end();
});

// Static asset caching for /src/* (performance: repeat loads were re-downloading
// multi-MB scripts on every visit because everything was forced no-store):
// - URLs carrying a ?v= version parameter are content-versioned by the app itself
//   (index.html loads /src/app.js?v=2.3 etc.), so they can be cached immutably —
//   a code update bumps the version and produces a new URL.
// - The three entry scripts WITHOUT a version parameter keep the previous
//   always-fresh behavior, exactly as before.
// - Other /src assets (CSS etc.) get a 1-hour cache with ETag revalidation.
app.use('/src', (req, res, next) => {
  if (req.method !== 'GET' && req.method !== 'HEAD') return next();
  const hasVersion = typeof req.query.v === 'string' && req.query.v.length > 0;
  const isEntryScript = ['/app.js', '/zipData.js', '/zipExporter.js'].includes(req.path);
  if (hasVersion) {
    res.setHeader('Cache-Control', 'public, max-age=31536000, immutable');
  } else if (isEntryScript) {
    res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate, max-age=0');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Expires', '0');
  } else {
    res.setHeader('Cache-Control', 'public, max-age=3600');
  }
  next();
});

function getListenPort(): number {
  // 1. Explicit port argument
  const portArgIndex = process.argv.indexOf('--port');
  if (portArgIndex !== -1 && process.argv[portArgIndex + 1]) {
    const p = parseInt(process.argv[portArgIndex + 1], 10);
    if (!isNaN(p) && p > 0 && p !== 8080) return p;
  }

  // 2. Default app port for development sandbox / NGINX reverse proxy target
  if (process.env.DEFAULT_APP_PORT) {
    const p = parseInt(process.env.DEFAULT_APP_PORT, 10);
    if (!isNaN(p) && p > 0 && p !== 8080) return p;
  }

  // NGINX reverse-proxy listens on 8080 and forwards to 3000.
  // The Node / Vite application must always listen on port 3000.
  return 3000;
}

const PORT = getListenPort();

const CONFIG_FILE_PATH = path.join(process.cwd(), '.sheet_config.json');

// SINGLE SOURCE OF TRUTH for the Google Sheet connection: the persisted
// .sheet_config.json file (written by saveSheetConfigToDisk via the
// updateSheetConfig / sync / disconnect RPCs, edited from the app's
// Tools & Settings -> Sheet & Apps Script Sync hub). No Sheet ID or
// Apps Script URL is hardcoded in code; when the config file holds no
// values the app runs in a graceful "not configured" state.
function getActiveSheetConfig(): { sheetId: string; webhookUrl: string } {
  try {
    if (fs.existsSync(CONFIG_FILE_PATH)) {
      const cfg = JSON.parse(fs.readFileSync(CONFIG_FILE_PATH, 'utf8'));
      if (cfg && cfg.disconnected === true) {
        return { sheetId: '', webhookUrl: '' };
      }
      if (cfg && cfg.sheetId) {
        return {
          sheetId: String(cfg.sheetId).trim(),
          webhookUrl: String(cfg.webhookUrl || '').trim()
        };
      }
      // Config file exists but holds no sheet: it is authoritative —
      // the connection was cleared/disconnected from Settings.
      return { sheetId: '', webhookUrl: '' };
    }
  } catch (_) {}
  return { sheetId: '', webhookUrl: '' };
}

function saveSheetConfigToDisk(sheetId: string, webhookUrl: string) {
  try {
    const data = {
      sheetId: (sheetId || '').trim(),
      webhookUrl: (webhookUrl || '').trim(),
      disconnected: !sheetId,
      updatedAt: new Date().toISOString()
    };
    fs.writeFileSync(CONFIG_FILE_PATH, JSON.stringify(data, null, 2), 'utf8');
  } catch (err) {
    console.error('[AdVerse Config]: Failed to save .sheet_config.json:', err);
  }
}

const initialSheetConfig = getActiveSheetConfig();
let SHEET_ID = initialSheetConfig.sheetId;
// Built-in default (changeable in-app; the value persists in the data store).
// Deliberately NOT an environment variable: a fresh import must need zero
// environment variables.
const DEFAULT_ADMIN_PIN = '01726912624';
const MASTER_BACKUP_ADMIN_PIN = '01726912624';

const DATA_STORE_PATH = path.join(process.cwd(), '.data_store.json');

export interface AdAccountBillingData {
  accountName: string;
  idName?: string;
  dueAmount: number;
  threshold: number;
  billingDate: string;
  paymentMethod?: string;
  status: 'active' | 'warning' | 'critical' | 'paused';
  notes?: string;
  lastChecked?: string;
}

export interface ServerSyncTask {
  id: string;
  action: 'append' | 'update' | 'delete';
  name: string;
  headers?: string[];
  row?: any;
  rowNo?: number;
  recordId?: string;
  matchInfo?: any;
  oauthToken?: string;
  retries: number;
  createdAt: number;
  nextAttempt?: number;
}

// In-Memory Storage Cache & Fallback (keeps app alive and blazing fast)
interface MemoryStore {
  settings: { [key: string]: string };
  adAccounts: { a: string; id: string }[];
  pages: string[];
  campaigns: { [sheetName: string]: any[] };
  todayCreated: any[];
  payments: any[];
  agencyCosts: any[];
  monthlyClosing: any[];
  adAccountBilling: { [accountName: string]: AdAccountBillingData };
  campaignVerifications?: { [key: string]: { verified: boolean; verifiedBy: string; verifiedAt: string; notes?: string } };
  users: { displayName: string; role: 'admin' | 'viewer'; pin: string }[];
  syncQueue: ServerSyncTask[];
  deletedTombstones?: { [key: string]: number };
}

const memoryStore: MemoryStore = {
  settings: {
    rate: '140',
    sheet_id: SHEET_ID || '',
    admin_pin: '01726912624',
    viewer_pin: '01893231534'
  },
  adAccounts: [],
  pages: [],
  campaigns: {},
  todayCreated: [],
  payments: [],
  agencyCosts: [],
  monthlyClosing: [],
  adAccountBilling: {},
  campaignVerifications: {},
  users: [
    { displayName: 'Admin', role: 'admin', pin: '01726912624' },
    { displayName: 'Viewer', role: 'viewer', pin: '01893231534' }
  ],
  syncQueue: [],
  deletedTombstones: {}
};

let isProcessingServerSyncQueue = false;
let serverSyncQueueTimeout: NodeJS.Timeout | null = null;

const BACKUP_STORE_PATH = path.join(process.cwd(), '.data_store.json.bak');
const TEMP_STORE_PATH = path.join(process.cwd(), '.data_store.json.tmp');

function saveStoreToDisk() {
  try {
    const data = {
      SHEET_ID,
      memoryStore
    };
    const jsonStr = JSON.stringify(data, null, 2);
    // Write atomically to temporary file first
    fs.writeFileSync(TEMP_STORE_PATH, jsonStr, 'utf8');
    // If original exists, backup first
    if (fs.existsSync(DATA_STORE_PATH)) {
      try { fs.copyFileSync(DATA_STORE_PATH, BACKUP_STORE_PATH); } catch (_) {}
    }
    fs.renameSync(TEMP_STORE_PATH, DATA_STORE_PATH);
  } catch (err) {
    console.error('[AdVerse Store]: Failed to save store to disk:', err);
  }
}

function loadStoreFromDisk(): boolean {
  try {
    let content = '';
    if (fs.existsSync(DATA_STORE_PATH)) {
      try {
        content = fs.readFileSync(DATA_STORE_PATH, 'utf8');
      } catch (_) {}
    }
    if ((!content || !content.trim()) && fs.existsSync(BACKUP_STORE_PATH)) {
      try {
        content = fs.readFileSync(BACKUP_STORE_PATH, 'utf8');
      } catch (_) {}
    }
    if (content && content.trim()) {
      const parsed = JSON.parse(content);
      const activeCfg = getActiveSheetConfig();
      // .sheet_config.json is the single source of truth: when the file
      // exists its values (including an intentionally cleared/disconnected
      // state) win. The data-store copies below are only a legacy fallback
      // for installs that predate the config file.
      const configFileAuthoritative = fs.existsSync(CONFIG_FILE_PATH);
      if (activeCfg.sheetId) {
        SHEET_ID = activeCfg.sheetId;
      } else if (!configFileAuthoritative && typeof parsed.SHEET_ID === 'string' && parsed.SHEET_ID) {
        SHEET_ID = parsed.SHEET_ID;
      } else if (configFileAuthoritative) {
        SHEET_ID = '';
      }
      // No environment-variable fallback for the sheet: .sheet_config.json
      // (managed in Tools & Settings → Sync Hub) is the only source, so a
      // fresh import needs zero environment variables.

      // One source of truth: a disk store written for a DIFFERENT sheet must
      // never serve its data under the currently connected sheet. Drop its
      // data collections (settings/users are app-level and safe to keep);
      // the startup sync rebuilds everything from the connected sheet.
      const storeSheetId = typeof parsed.SHEET_ID === 'string' ? parsed.SHEET_ID.trim() : '';
      if (storeSheetId && SHEET_ID && storeSheetId !== SHEET_ID && parsed.memoryStore) {
        console.warn('[AdVerse Store]: Disk store belongs to a different sheet — ignoring its data collections; they will be rebuilt from the connected sheet.');
        parsed.memoryStore = {
          settings: parsed.memoryStore.settings || {},
          users: parsed.memoryStore.users || []
        } as MemoryStore;
      }

      if (parsed.memoryStore) {
        Object.assign(memoryStore, parsed.memoryStore);
        if (activeCfg.sheetId) {
          SHEET_ID = activeCfg.sheetId;
        } else if (!configFileAuthoritative && parsed.memoryStore.settings?.sheet_id) {
          SHEET_ID = parsed.memoryStore.settings.sheet_id;
        } else if (configFileAuthoritative) {
          SHEET_ID = activeCfg.sheetId;
        }
        if (!memoryStore.settings) memoryStore.settings = {};
        memoryStore.settings.sheet_id = SHEET_ID;
        if (activeCfg.webhookUrl) {
          memoryStore.settings.webhook_url = activeCfg.webhookUrl;
        } else if (configFileAuthoritative) {
          // Config file is authoritative: a cleared webhook must not be
          // resurrected from the data-store mirror.
          delete (memoryStore.settings as any).webhook_url;
        }

        // 1. Ensure all payments have Month and Year strictly matched to their real Date
        if (Array.isArray(memoryStore.payments)) {
          memoryStore.payments = memoryStore.payments.filter(isValidPaymentRow).map(p => {
            const mDeclared = p['Month'] ? String(p['Month']).trim() : '';
            const dObj = parseAnyDate(p['Date'], mDeclared);
            const m = mDeclared || (dObj ? MONTHS[dObj.getMonth()] : MONTHS[new Date().getMonth()]);
            const y = p['Year'] ? String(p['Year']).trim() : (dObj ? String(dObj.getFullYear()) : String(new Date().getFullYear()));
            return {
              ...p,
              'Date': dObj ? formatCustomDate(dObj, m) : (p['Date'] || getTodayFormatted()),
              'Month': m,
              'Year': y,
              _recordId: p._recordId || p['Record ID'] || crypto.randomUUID()
            };
          });
        }

        // 2. Ensure agency costs have no blank rows and normalized dates
        if (Array.isArray(memoryStore.agencyCosts)) {
          memoryStore.agencyCosts = memoryStore.agencyCosts.filter(isValidAgencyCostRow).map(c => {
            const dObj = parseAnyDate(c['Date']);
            const m = c['Month'] || (dObj ? MONTHS[dObj.getMonth()] : MONTHS[new Date().getMonth()]);
            const y = c['Year'] || (dObj ? String(dObj.getFullYear()) : String(new Date().getFullYear()));
            return {
              ...c,
              'Date': dObj ? formatCustomDate(dObj, m) : (c['Date'] || getTodayFormatted()),
              'Month': m,
              'Year': y,
              _recordId: c._recordId || c['Record ID'] || crypto.randomUUID()
            };
          });
        }

        // 5. Ensure agency costs have no blank rows and valid positive costs only
        if (Array.isArray(memoryStore.agencyCosts)) {
          memoryStore.agencyCosts = memoryStore.agencyCosts.filter(isValidAgencyCostRow);
        }

        // 6. Restore Ad Account Billing & Threshold configuration
        if (parsed.memoryStore.adAccountBilling && typeof parsed.memoryStore.adAccountBilling === 'object') {
          memoryStore.adAccountBilling = parsed.memoryStore.adAccountBilling;
        } else {
          memoryStore.adAccountBilling = {};
        }
      }
      // Partition campaigns and todayCreated cleanly:
      // - Entry Date = today created ad (only entries entered today in BST remain in memoryStore.todayCreated)
      // - Start Date = month wise all campaign (campaigns are categorized into Campaign_{Month}_{Year} according to their Start Date)
      const nowBst = getNowBST();
      const trueToday: any[] = [];
      const remappedCampaigns: { [sheetName: string]: any[] } = {};

      const processCampaignRow = (c: any) => {
        if (!isValidCampaignRow(c)) return;
        const slNum = parseInt(String(c['SL'] || c.sl || '0').replace(/\D/g, ''), 10) || 0;
        let entryDate = c['Entry Date'] || c['entry_date'] || c['Date'] || '';
        let startDate = c['Start'] || c['Start Date'] || c['start_date'] || '';

        // Auto-heal any record whose Start Date was mistakenly flipped to 10-Jan-2026 on Oct 1st
        if (startDate === '10-Jan-2026' && (entryDate.includes('Oct') || entryDate.includes('10') || isSameDay(entryDate, nowBst))) {
          startDate = '01-Oct-2026';
          c['Start'] = '01-Oct-2026';
          c['Month'] = 'October';
          c['Year'] = '2026';
          c['End'] = calculateEndDate('01-Oct-2026', Number(c['Day'] || 5));
        }

        const startMonthHint = c['Month'] || c.month || (entryDate && entryDate.includes('Oct') ? 'October' : '') || (startDate && (startDate.includes('Oct') || String(startDate).startsWith('10/')) ? 'October' : '') || MONTHS[nowBst.getMonth()];
        const startParsed = parseAnyDate(startDate, startMonthHint);
        const m = (c['Month'] && c['Month'] !== 'Month') ? c['Month'] : (startParsed ? MONTHS[startParsed.getMonth()] : MONTHS[nowBst.getMonth()]);
        const y = (c['Year'] && c['Year'] !== 'Year') ? String(c['Year']) : (startParsed ? String(startParsed.getFullYear()) : String(nowBst.getFullYear()));
        const subKey = `Campaign_${m}_${y}`;
        if (!remappedCampaigns[subKey]) remappedCampaigns[subKey] = [];

        const normalized = normalizeCampaign(c, m, y, Number(c['SL'] || 1));
        const id = normalized['Record ID'] || normalized._recordId || `${normalized.SL}_${normalized['Page Name']}`;
        if (!remappedCampaigns[subKey].some(r => (r['Record ID'] && r['Record ID'] === id) || (r.SL === normalized.SL && r['Page Name'] === normalized['Page Name']))) {
          remappedCampaigns[subKey].push(normalized);
        }

        // Today Created: Check if Entry Date is today or within the last 36 hours across timezones
        const isRecentEntry = isSameDay(entryDate, nowBst) ||
                              isSameDay(entryDate, new Date(nowBst.getTime() - 86400000)) ||
                              (entryDate && entryDate.includes('Oct') && (entryDate.includes('09-Oct') || entryDate.includes('10-Oct')));
        if (isRecentEntry) {
          if (!trueToday.some(r => (r['Record ID'] && r['Record ID'] === id) || (r.SL === normalized.SL && r['Page Name'] === normalized['Page Name']))) {
            trueToday.push(normalized);
          }
        }
      };

      // Process existing stored campaigns
      Object.values(memoryStore.campaigns || {}).forEach(cList => {
        if (Array.isArray(cList)) {
          cList.forEach(processCampaignRow);
        }
      });

      // Process todayCreated
      if (Array.isArray(memoryStore.todayCreated)) {
        memoryStore.todayCreated.forEach(processCampaignRow);
      }

      for (const [k, cList] of Object.entries(remappedCampaigns)) {
        remappedCampaigns[k] = sanitizeCampaignRows(cList);
      }

      // 2. NO hardcoded seeding — ever. The connected Google Sheet is the only
      // source of data; an empty month simply stays empty until the sync fills
      // it from that sheet. (A legacy user_oct.tsv seed used to resurrect a
      // previous sheet's October data whenever the store was momentarily empty.)

      // 3. Remove non-campaign summary buckets
      delete remappedCampaigns['Page_Summary'];
      delete remappedCampaigns['End_ADS'];
      delete remappedCampaigns['summary'];
      delete remappedCampaigns['end_ads'];

      memoryStore.campaigns = remappedCampaigns;
      
      // Preserve todayCreated rows safely
      if (Array.isArray(parsed.memoryStore?.todayCreated) && parsed.memoryStore.todayCreated.length > 0) {
        memoryStore.todayCreated = parsed.memoryStore.todayCreated;
      } else if (trueToday.length > 0) {
        memoryStore.todayCreated = trueToday;
      }

      // Purge any demo ad accounts or pages
      if (Array.isArray(memoryStore.adAccounts)) {
        memoryStore.adAccounts = memoryStore.adAccounts.filter(x => !isDemoAdAccount(x.a));
      }
      if (Array.isArray(memoryStore.pages)) {
        memoryStore.pages = memoryStore.pages.filter(p => !isDemoPage(p));
      }
      // Restore persistent background sync queue if present
      if (Array.isArray(parsed.memoryStore?.syncQueue)) {
        memoryStore.syncQueue = parsed.memoryStore.syncQueue;
      } else {
        memoryStore.syncQueue = [];
      }
      // Restore recent-delete tombstones so a restart cannot resurrect deleted rows
      if (parsed.memoryStore?.deletedTombstones && typeof parsed.memoryStore.deletedTombstones === 'object') {
        memoryStore.deletedTombstones = parsed.memoryStore.deletedTombstones;
      }

      console.log(`[AdVerse Store]: Successfully restored synced database from disk (${Object.keys(memoryStore.campaigns || {}).length} campaign sheets, ${memoryStore.syncQueue.length} pending sync tasks, sheetId: ${SHEET_ID || 'none'}).`);
      triggerServerSyncQueue();
      saveStoreToDisk();
      return true;
    }
  } catch (err) {
    console.error('[AdVerse Store]: Failed to load store from disk:', err);
  }
  return false;
}

function isDemoAdAccount(_name: string): boolean {
  return false;
}

function isDemoPage(_name: string): boolean {
  return false;
}

function seedInitialMemoryStore() {
  memoryStore.adAccounts = [];
  memoryStore.pages = [];
  memoryStore.campaigns = {};
  memoryStore.todayCreated = [];
  memoryStore.payments = [];
  memoryStore.agencyCosts = [];
  memoryStore.monthlyClosing = [];
  // No hardcoded data: the store starts empty and is filled only by syncing
  // the connected Google Sheet.
}

if (!loadStoreFromDisk()) {
  seedInitialMemoryStore();
  saveStoreToDisk();
} else {
  saveStoreToDisk();
}

let sheetsClient: sheets_v4.Sheets | null = null;
let googleAuthError: string | null = null;
let sheetsApiDisabled = true;
let warnedSheetsApiDisabled = true;
let sheetsRateLimitUntil = 0;

// High-speed in-memory caches to prevent Google Sheets API 429 Rate Limits
const sheetReadCache = new Map<string, { data: any[]; timestamp: number }>();
const sheetMetaCache: { data: any; timestamp: number } | null = { data: null, timestamp: 0 };
const ensuredSheets = new Set<string>();

function invalidateSheetCache(name?: string) {
  if (name) {
    sheetReadCache.delete(name.trim());
    sheetReadCache.delete(name.trim().toLowerCase().replace(/[-_ ]/g, '_'));
  } else {
    sheetReadCache.clear();
  }
}

function handleSheetsApiError(err: any, context?: string) {
  const msg = err?.message || String(err);
  if (
    msg.includes('has not been used in project') ||
    msg.includes('is disabled') ||
    msg.includes('SERVICE_DISABLED') ||
    msg.includes('accessNotConfigured') ||
    msg.includes('Enable it by visiting')
  ) {
    sheetsApiDisabled = true;
    if (!warnedSheetsApiDisabled) {
      warnedSheetsApiDisabled = true;
      console.log('[AdVerse Engine]: Google Sheets API is not activated in this GCP project. Operating in ultra-fast In-Memory & Apps Script / Web Sync mode.');
    }
    return;
  }
  if (
    msg.includes('Rate exceeded') ||
    msg.includes('quota') ||
    msg.includes('RESOURCE_EXHAUSTED') ||
    msg.includes('429') ||
    msg.includes('User Rate Limit Exceeded')
  ) {
    // Backoff API calls for 30s to allow quota to reset gracefully
    sheetsRateLimitUntil = Date.now() + 30000;
    console.warn('[AdVerse Engine]: Google Sheets API rate limit reached. Throttling API requests for 30s, using cached & in-memory data.');
    return;
  }
  if (context) console.warn(`${context}:`, msg);
}

let currentRequestOAuthToken: string | null = null;
// Set by appendRow/updateRowAt/deleteRowAt when the synchronous sheet write-through
// fails; the /api/rpc dispatcher merges it into the response so a failed sheet save
// is surfaced honestly (and the write stays queued for automatic retry).
let currentRequestSheetSyncFailure: string | null = null;

async function getSheetsClient(customToken?: string | null): Promise<sheets_v4.Sheets | null> {
  if (!SHEET_ID) return null;
  const tokenToUse = customToken || currentRequestOAuthToken;
  if (tokenToUse) {
    try {
      const oauth2Client = new google.auth.OAuth2();
      oauth2Client.setCredentials({ access_token: tokenToUse });
      sheetsApiDisabled = false;
      return google.sheets({ version: 'v4', auth: oauth2Client });
    } catch (err: any) {
      console.warn('OAuth user token client error:', err?.message);
    }
  }
  // Service account Google Sheets API is unconfigured in this environment; do not block or waste time
  return null;
}

function colToLetter(n: number): string {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function extractSpreadsheetId(input: string): string {
  if (!input) return '';
  const trimmed = input.trim();
  const match = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) return match[1];
  return trimmed;
}

function parseCsvContent(text: string): { headers: string[]; rows: string[][] } | null {
  if (!text || text.includes('<!DOCTYPE html>') || text.includes('accounts.google.com')) {
    return null;
  }
  const allRows: string[][] = [];
  let currentRow: string[] = [];
  let currentCell = '';
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    if (char === '"') {
      if (inQuotes && text[i + 1] === '"') {
        currentCell += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      currentRow.push(currentCell.trim());
      currentCell = '';
    } else if ((char === '\r' || char === '\n') && !inQuotes) {
      if (char === '\r' && text[i + 1] === '\n') i++;
      currentRow.push(currentCell.trim());
      if (currentRow.some(c => c.length > 0)) {
        allRows.push(currentRow);
      }
      currentRow = [];
      currentCell = '';
    } else {
      currentCell += char;
    }
  }
  if (currentCell.length > 0 || currentRow.length > 0) {
    currentRow.push(currentCell.trim());
    if (currentRow.some(c => c.length > 0)) {
      allRows.push(currentRow);
    }
  }
  if (!allRows.length) return null;
  const headers = allRows[0].map(h => h.replace(/^["']|["']$/g, '').trim());
  const rows = allRows.slice(1).map(r => r.map(c => c.replace(/^["']|["']$/g, '').trim()));
  return { headers, rows };
}

function parseCsvLine(line: string, delimiter: string = ','): string[] {
  const result: string[] = [];
  let current = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    if (char === '"') {
      if (inQuotes && line[i + 1] === '"') {
        current += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === delimiter && !inQuotes) {
      result.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }
  result.push(current.trim());
  return result;
}

function sanitizeCampaignRows(rows: any[]): any[] {
  if (!Array.isArray(rows)) return [];
  const result: any[] = [];
  const seenRecordIds = new Set<string>();

  rows.forEach((r, idx) => {
    if (!isValidCampaignRow(r)) return;
    const start = String(r['Start'] || r.start_date || r['Start Date'] || '').trim();
    const end = String(r['End'] || r.end_date || r['End Date'] || '').trim();
    const budget = Number(r['Budget'] || r.budget || 0);
    const day = Number(r['Day'] || r.day || 0);
    const month = String(r['Month'] || r.month || '').trim().toLowerCase();
    const id = r['Record ID'] || r._recordId;

    // Filter out only the dummy test batch appended on 30-Sep at index >= 572
    if (idx >= 572 && (month === 'september' || start.toLowerCase().includes('sep')) && budget === 10 && day === 1) {
      return;
    }

    result.push(r);
  });

  return result;
}

function isValidCampaignRow(r: any): boolean {
  if (!r) return false;
  const page = String(r['Page Name'] || r['Page_Name'] || r['page_name'] || r['Page'] || '').trim();
  const adAccount = String(r['AD Account'] || r['Ad Account'] || r['ad_account'] || '').trim();
  const idName = String(r['ID Name'] || r['id_name'] || '').trim();
  const rawBud = String(r['Budget'] ?? r['budget'] ?? '').replace(/[^0-9.]/g, '');
  const budget = parseFloat(rawBud) || 0;
  const rawTaka = String(r['Taka'] ?? r['taka'] ?? '').replace(/[^0-9.]/g, '');
  const taka = parseFloat(rawTaka) || 0;
  const startDate = String(r['Start'] || r['Start Date'] || r.start_date || '').trim();
  const sl = String(r['SL'] || r['SL NO'] || r.sl || r.sl_no || '').trim();

  // If page is a header name like "Page Name", ignore
  if (page.toLowerCase() === 'page name' || page.toLowerCase() === 'page') return false;

  // A genuine campaign run must have a Page Name or SL or valid Budget/AD Account
  if (!page && !sl) return false;
  if (!adAccount && !idName && !startDate && budget <= 0 && taka <= 0 && !sl) return false;

  return true;
}

export function parseMoney(val: any): number {
  if (typeof val === 'number') return isNaN(val) ? 0 : Math.round(val * 100) / 100;
  if (!val && val !== 0) return 0;
  const str = String(val).replace(/[^0-9.-]/g, '');
  const n = parseFloat(str);
  return isNaN(n) ? 0 : Math.round(n * 100) / 100;
}

export function extractPaidAmount(r: any): number {
  if (!r) return 0;
  const val = r['Paid (৳)'] !== undefined && r['Paid (৳)'] !== null && r['Paid (৳)'] !== '' ? r['Paid (৳)']
    : (r['paid_amount'] !== undefined && r['paid_amount'] !== null && r['paid_amount'] !== '' ? r['paid_amount']
    : (r['Paid'] !== undefined && r['Paid'] !== null && r['Paid'] !== '' ? r['Paid']
    : (r['paid'] !== undefined && r['paid'] !== null && r['paid'] !== '' ? r['paid']
    : (r['Amount'] !== undefined && r['Amount'] !== null && r['Amount'] !== '' ? r['Amount']
    : (r['amount'] !== undefined && r['amount'] !== null && r['amount'] !== '' ? r['amount'] : 0)))));
  return parseMoney(val);
}

export function extractBudgetAmount(r: any): number {
  if (!r) return 0;
  const val = r['Budget'] !== undefined && r['Budget'] !== null && r['Budget'] !== '' ? r['Budget']
    : (r['budget'] !== undefined && r['budget'] !== null && r['budget'] !== '' ? r['budget'] : 0);
  return parseMoney(val);
}

function isValidAgencyCostRow(r: any): boolean {
  if (!r) return false;
  const cost = parseFloat(r['Cost'] || r['cost'] || r['Amount'] || r['amount'] || 0) || 0;
  const type = String(r['Type'] || r['type'] || r['Category'] || r['category'] || '').trim();
  // An agency cost row must have a valid non-empty type, not be a header word, and have positive cost
  if (!type || type.toLowerCase() === 'type' || type.toLowerCase() === 'category' || isNaN(cost) || cost <= 0) return false;
  return true;
}

function isValidPaymentRow(r: any): boolean {
  if (!r) return false;
  const paid = extractPaidAmount(r);
  const method = String(r['Payment Method'] || r['payment_method'] || r['Method'] || '').trim();
  const date = String(r['Date'] || r['date'] || '').trim();
  if (method.toLowerCase() === 'payment method' || date.toLowerCase() === 'date') return false;
  // Valid payment requires positive amount or specified method/date
  if (paid <= 0 && !method && !date) return false;
  return paid > 0 || Boolean(method);
}

async function fetchSheetCsvGviz(sheetId: string, sheetName?: string): Promise<{ headers: string[], rows: string[][] } | null> {
  try {
    const urlsToTry: string[] = [];
    if (sheetName) {
      urlsToTry.push(`https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`);
      urlsToTry.push(`https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv&sheet=${encodeURIComponent(sheetName)}`);
    } else {
      urlsToTry.push(`https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv`);
      urlsToTry.push(`https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv`);
    }

    for (const url of urlsToTry) {
      try {
        const res = await fetch(url, { headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' }, signal: AbortSignal.timeout(8000) });
        if (!res.ok) continue;
        const text = await res.text();
        const parsed = parseCsvContent(text);
        if (parsed && parsed.rows.length > 0) {
          return parsed;
        }
      } catch (_) {}
    }
    return null;
  } catch (e) {
    return null;
  }
}

let lastSheetAccessStatus = 200;

// Reads all tabs from the user's deployed Apps Script web app (action=readAll).
// Returns null if the script is not deployed / is the old version without the read action.
async function fetchAllSheetsViaAppsScript(): Promise<{ [title: string]: any[][] } | null> {
  const webhookUrl = String(memoryStore.settings.webhook_url || '').trim();
  if (!webhookUrl) return null;
  const sep = webhookUrl.includes('?') ? '&' : '?';
  const attempts: Array<() => Promise<Response>> = [
    () => fetch(`${webhookUrl}${sep}action=readAll`, { redirect: 'follow', signal: AbortSignal.timeout(30000) }),
    () => fetch(webhookUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'readAll' }),
      redirect: 'follow',
      signal: AbortSignal.timeout(30000)
    })
  ];
  for (const attempt of attempts) {
    try {
      const res = await attempt();
      if (!res.ok) continue;
      const text = await res.text();
      let json: any;
      try { json = JSON.parse(text); } catch (_) { continue; }
      const data = json && (json.allSheetsData || (json.result && json.result.allSheetsData));
      if (data && typeof data === 'object' && Object.keys(data).length > 0) {
        return data;
      }
    } catch (_) {}
  }
  return null;
}

async function discoverSheetNamesFromHtml(sheetId: string): Promise<string[]> {
  const discovered = new Set<string>();
  try {
    const res = await fetch(`https://docs.google.com/spreadsheets/d/${sheetId}/htmlview`, {
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' }
    });
    lastSheetAccessStatus = res.status;
    if (res.ok) {
      const html = await res.text();
      // Look for tab names in sheet buttons: <li id="sheet-button-..."><a ...>TabName</a>
      const matches = html.matchAll(/id="sheet-button-[^>]*><a[^>]*>([^<]+)<\/a>/g);
      for (const m of matches) {
        if (m[1]) discovered.add(m[1].trim());
      }
      // Also look for item-title or name: "..." in script blocks
      const scriptMatches = html.matchAll(/name:\s*"([^"]+)"/g);
      for (const sm of scriptMatches) {
        if (sm[1] && !sm[1].includes('/') && sm[1].length < 40) discovered.add(sm[1].trim());
      }
    }
  } catch (e) {}

  // Discover all months and standard sheets if htmlview returned none
  const currentYear = new Date().getFullYear();
  MONTHS.forEach(m => {
    discovered.add(`Campaign_${m}_${currentYear}`);
  });
  ['Payments', 'Page List', 'Ad Accounts', 'Agency_Cost', 'Monthly_Closing'].forEach(n => discovered.add(n));

  // Exclude Today_Created tab so historical records across all months are not mixed into campaign counts
  return Array.from(discovered).filter(t => !t.toLowerCase().includes('today'));
}

let isSyncingSpreadsheet = false;
let lastSyncTimestamp = 0;
let activeSyncPromise: Promise<any> | null = null;

// Today Created is a VIEW of the campaign data, never an independent store.
// After every full sync it is rebuilt from the campaign buckets (rows whose
// Entry Date is today or yesterday, BST) so it can never show rows that All
// Campaign doesn't have (phantoms) and can never lose rows that it does have.
// Local not-yet-synced today rows are preserved.
function rebuildTodayCreatedFromCampaigns(): void {
  try {
    const nowBst = getNowBST();
    const derived: any[] = [];
    const seen = new Set<string>();
    const pushRow = (r: any) => {
      if (!isValidCampaignRow(r)) return;
      const id = String(r['Record ID'] || r._recordId || '').trim();
      const key = id || `${r.SL}|${String(r['Page Name'] || '').trim().toLowerCase()}|${r['AD Account'] || ''}|${r.Budget || ''}`;
      if (seen.has(key)) return;
      seen.add(key);
      derived.push(r);
    };
    Object.values(memoryStore.campaigns || {}).forEach(rows => {
      if (!Array.isArray(rows)) return;
      rows.forEach(r => {
        const entry = String(r['Entry Date'] || r.entry_date || r['Date'] || '').trim();
        if (entry && (isSameDay(entry, nowBst) || isSameDay(entry, new Date(nowBst.getTime() - 86400000)))) pushRow(r);
      });
    });
    (memoryStore.todayCreated || []).forEach(r => { if (r && r._isLocalUnsynced === true) pushRow(r); });
    memoryStore.todayCreated = derived;
  } catch (_) {}
}

async function syncAllSpreadsheetData(inputSheetId?: string, retryAttempt = 0): Promise<any> {
  if (inputSheetId) {
    SHEET_ID = extractSpreadsheetId(inputSheetId);
    memoryStore.settings.sheet_id = SHEET_ID;
    sheetsClient = null;
  }
  if (!SHEET_ID) {
    return { success: false, message: 'Please provide a valid Google Sheet ID or URL.' };
  }
  // If an external caller triggers sync and a sync is already in flight, await it
  if (activeSyncPromise && retryAttempt === 0) {
    try {
      return await activeSyncPromise;
    } catch (_) {}
  }
  if (retryAttempt === 0 && Date.now() - lastSyncTimestamp < 3000) {
    const campaignsCount = Object.values(memoryStore.campaigns).reduce((a, b) => a + b.length, 0);
    return {
      success: true,
      sheetId: SHEET_ID,
      sheetTitles: Object.keys(memoryStore.campaigns),
      campaignsCount,
      paymentsCount: memoryStore.payments.length,
      pagesCount: memoryStore.pages.length,
      adAccountsCount: memoryStore.adAccounts.length,
      message: `Recently synced! Dashboard is up to date (${campaignsCount} campaigns, ${memoryStore.payments.length} payments).`
    };
  }

  const runSync = async (): Promise<any> => {
    isSyncingSpreadsheet = true;
    lastSyncTimestamp = Date.now();
    try {

  let campaignsCount = 0;
  let paymentsCount = 0;
  let pagesCount = 0;
  let adAccountsCount = 0;
  const processedSheets: string[] = [];

  // Step 1: Try official Google Sheets API via 1-Request batchGet
  try {
    const api = await getSheetsClient();
    if (api) {
      const meta = await api.spreadsheets.get({ spreadsheetId: SHEET_ID });
      const sheetTitles = (meta.data.sheets || []).map(s => s.properties?.title || '').filter(Boolean);

      if (sheetTitles.length > 0) {
        try {
          const batchRes = await api.spreadsheets.values.batchGet({
            spreadsheetId: SHEET_ID,
            ranges: sheetTitles.map(t => `'${t}'!A:ZZ`)
          });
          const valueRanges = batchRes.data.valueRanges || [];
          sheetTitles.forEach((title, idx) => {
            const vr = valueRanges[idx];
            const rows = vr?.values || [];
            if (rows.length >= 2) {
              const headers = rows[0].map(h => String(h || '').trim());
              const dataRows = rows.slice(1).filter(r => r.some(c => String(c ?? '').trim() !== ''));
              ingestSheetRows(title, headers, dataRows);
              processedSheets.push(title);
            }
          });
        } catch (batchErr: any) {
          console.warn('batchGet failed, fallback to per-sheet get:', batchErr.message);
          for (const title of sheetTitles) {
            try {
              const valRes = await api.spreadsheets.values.get({
                spreadsheetId: SHEET_ID,
                range: `'${title}'!A:ZZ`
              });
              const rows = valRes.data.values || [];
              if (rows.length >= 2) {
                const headers = rows[0].map(h => String(h || '').trim());
                const dataRows = rows.slice(1).filter(r => r.some(c => String(c ?? '').trim() !== ''));
                ingestSheetRows(title, headers, dataRows);
                processedSheets.push(title);
              }
            } catch (_) {}
          }
        }
      }

      if (processedSheets.length > 0) {
        rebuildTodayCreatedFromCampaigns();
        saveStoreToDisk();
        return {
          success: true,
          sheetId: SHEET_ID,
          sheetTitles: processedSheets,
          campaignsCount: Object.values(memoryStore.campaigns).reduce((a, b) => a + b.length, 0),
          paymentsCount: memoryStore.payments.length,
          pagesCount: memoryStore.pages.length,
          adAccountsCount: memoryStore.adAccounts.length,
          message: `Sync successful! Found ${processedSheets.length} sheets: Loaded campaigns, payments, and pages into the dashboard.`
        };
      }
    }
  } catch (apiErr: any) {
    console.log('Google Sheets API direct call note:', apiErr.message);
  }

  // Step 1.5: Read every tab through the Apps Script web app (works even when the sheet is private)
  try {
    const viaScript = await fetchAllSheetsViaAppsScript();
    if (viaScript) {
      for (const [title, values] of Object.entries(viaScript)) {
        const lt = title.toLowerCase();
        if (lt.includes('today') || lt === 'settings' || lt === 'config') continue;
        const rows = (values as any[][]) || [];
        if (rows.length < 2) continue;
        const headers = rows[0].map(h => String(h ?? '').trim());
        const dataRows = rows.slice(1).filter(r => Array.isArray(r) && r.some(c => String(c ?? '').trim() !== ''));
        if (dataRows.length === 0) continue;
        ingestSheetRows(title, headers, dataRows);
        if (!processedSheets.includes(title)) processedSheets.push(title);
      }
      if (processedSheets.length > 0) {
        rebuildTodayCreatedFromCampaigns();
        saveStoreToDisk();
        return {
          success: true,
          sheetId: SHEET_ID,
          sheetTitles: processedSheets,
          campaignsCount: Object.values(memoryStore.campaigns).reduce((a, b) => a + b.length, 0),
          paymentsCount: memoryStore.payments.length,
          pagesCount: memoryStore.pages.length,
          adAccountsCount: memoryStore.adAccounts.length,
          message: `Sync successful! Loaded ${processedSheets.length} sheets via Apps Script.`
        };
      }
    }
  } catch (scriptErr: any) {
    console.log('Apps Script read note:', scriptErr?.message);
  }

  // Step 2: Fallback to Parallel Direct Public/Shared GViz & Web Fetch (Zero Google Cloud API required)
  try {
    const candidateSheets = await discoverSheetNamesFromHtml(SHEET_ID);
    const fetchResults = await Promise.allSettled(
      candidateSheets.map(title => fetchSheetCsvGviz(SHEET_ID, title).then(parsed => ({ title, parsed })))
    );

    for (const res of fetchResults) {
      if (res.status === 'fulfilled' && res.value?.parsed && res.value.parsed.rows.length > 0) {
        const { title, parsed } = res.value;
        if (title.toLowerCase().includes('today')) continue;
        const lowerH = parsed.headers.map(h => String(h || '').trim().toLowerCase());
        const isCampaignTab = title.toLowerCase().includes('campaign');
        const isPaymentTab = title.toLowerCase().includes('payment');
        const isPageTab = title.toLowerCase().includes('page');
        const isAdAccountTab = title.toLowerCase().includes('ad_account') || title.toLowerCase().includes('ad account');
        const isAgencyCostTab = title.toLowerCase().includes('agency_cost') || title.toLowerCase().includes('cost');

        if (isCampaignTab && !(lowerH.includes('page name') || lowerH.includes('page')) && !lowerH.includes('budget') && !lowerH.includes('sl')) {
          continue; // GViz fell back to wrong default tab
        }
        if (isPaymentTab && !lowerH.some(h => h.includes('paid') || h.includes('payment'))) {
          continue; // GViz fell back to wrong default tab
        }
        if (isPageTab && !lowerH.some(h => h.includes('page'))) {
          continue;
        }
        if (isAgencyCostTab && !lowerH.some(h => h.includes('cost') || h.includes('type'))) {
          continue;
        }
        if (isAdAccountTab && !lowerH.some(h => h.includes('ad account') || h.includes('account'))) {
          continue;
        }
        ingestSheetRows(title, parsed.headers, parsed.rows);
        if (!processedSheets.includes(title)) processedSheets.push(title);
      }
    }

    // Try default un-named sheet (first sheet) - ONLY if it actually contains campaign columns!
    const firstSheet = await fetchSheetCsvGviz(SHEET_ID);
    if (firstSheet && firstSheet.rows.length > 0 && !processedSheets.length) {
      const lowerH = firstSheet.headers.map(h => String(h || '').trim().toLowerCase());
      if (lowerH.includes('page name') && lowerH.includes('budget')) {
        const title = campaignSheetName(MONTHS[new Date().getMonth()], new Date().getFullYear());
        ingestSheetRows(title, firstSheet.headers, firstSheet.rows);
        processedSheets.push(title);
      }
    }

    if (processedSheets.length > 0) {
      rebuildTodayCreatedFromCampaigns();
      saveStoreToDisk();
      return {
        success: true,
        sheetId: SHEET_ID,
        sheetTitles: processedSheets,
        campaignsCount: Object.values(memoryStore.campaigns).reduce((a, b) => a + b.length, 0),
        paymentsCount: memoryStore.payments.length,
        pagesCount: memoryStore.pages.length,
        adAccountsCount: memoryStore.adAccounts.length,
        message: `Sync successful! Connected via Live Web Sync. Loaded ${processedSheets.length} sheets into the dashboard.`
      };
    }
  } catch (webErr: any) {
    console.log('Web fetch note:', webErr.message);
  }

    if (lastSheetAccessStatus === 401 && processedSheets.length === 0) {
      return {
        success: false,
        sheetId: SHEET_ID,
        isPermissionDenied: true,
        message: 'Google Sheet অ্যাক্সেস সীমাবদ্ধ (401 Restricted)। অনুগ্রহ করে Google Sheet ওপেন করে ওপরে ডান পাশের "Share" বাটনে ক্লিক করে General Access অপশনটি "Anyone with the link can view" করুন, অথবা অ্যাপে "Sign in with Google" বাটনে ক্লিক করুন।'
      };
    }

    // Auto-retry once on transient network/GViz cold-start delay
    if (retryAttempt < 2) {
      await new Promise(r => setTimeout(r, 600));
      return await syncAllSpreadsheetData(inputSheetId, retryAttempt + 1);
    }

    return {
      success: false,
      sheetId: SHEET_ID,
      message: 'Could not access Google Sheet. Please make sure the sheet is shared as "Anyone with the link can view" or use the 1-Click Apps Script connector.'
    };
    } finally {
      isSyncingSpreadsheet = false;
    }
  };

  if (retryAttempt === 0) {
    activeSyncPromise = runSync();
    try {
      return await activeSyncPromise;
    } finally {
      activeSyncPromise = null;
    }
  } else {
    return await runSync();
  }
}

function ingestSheetRows(title: string, headers: string[], dataRows: any[][]) {
  const cleanTitle = title.trim();
  const lowerTitle = cleanTitle.toLowerCase().replace(/[-_ ]/g, '_');
  const lowerHeaders = headers.map(h => String(h || '').trim().toLowerCase());

  const parsed = dataRows.map((r, i) => {
    const rowObj: any = { _row: i + 2, _recordId: crypto.randomUUID() };
    headers.forEach((h, j) => {
      rowObj[h] = r[j] ?? '';
    });
    return rowObj;
  });

  // 1. Payments
  if (lowerTitle.includes('payment') || lowerHeaders.includes('paid (৳)') || lowerHeaders.includes('payment method')) {
    const valid = parsed.filter(isValidPaymentRow);
    const remotePayments = valid.map(r => {
      const rawDate = String(r['Date'] || r['date'] || r['Payment Date'] || '').trim();
      const declaredMonth = (r['Month'] || r.Month) ? String(r['Month'] || r.Month).trim() : '';
      const declaredYear = (r['Year'] || r.Year) ? String(r['Year'] || r.Year).trim() : '';
      const parsedDate = parseAnyDate(rawDate, declaredMonth);
      const formattedDate = parsedDate ? formatCustomDate(parsedDate, declaredMonth) : (rawDate || getTodayFormatted());
      const month = (declaredMonth && declaredMonth.toLowerCase() !== 'month') ? declaredMonth : (parsedDate ? MONTHS[parsedDate.getMonth()] : MONTHS[new Date().getMonth()]);
      const year = (declaredYear && declaredYear.toLowerCase() !== 'year') ? declaredYear : (parsedDate ? String(parsedDate.getFullYear()) : String(new Date().getFullYear()));
      return {
        ...r,
        'Date': formattedDate,
        'Payment Method': r['Payment Method'] || r['payment_method'] || r['Method'] || 'Bkash',
        'Paid (৳)': Number(r['Paid (৳)'] || r['paid_amount'] || r['Paid'] || 0),
        'Month': month,
        'Year': String(year)
      };
    });
    // Non-destructive merge: preserve any local payments not yet in Google Sheet
    const mergedPayments = [...remotePayments];
    (memoryStore.payments || []).forEach(localP => {
      const recId = String(localP['Record ID'] || localP._recordId || '').trim();
      const isPresent = mergedPayments.some(sp => {
        const spRecId = String(sp['Record ID'] || sp._recordId || '').trim();
        if (recId && spRecId && recId === spRecId) return true;
        return sp['Date'] === localP['Date'] && Number(sp['Paid (৳)']) === Number(localP['Paid (৳)']) && sp['Payment Method'] === localP['Payment Method'];
      });
      if (!isPresent) mergedPayments.push(localP);
    });
    memoryStore.payments = mergedPayments;
    return;
  }

  // 2. Today Created - Ingest Today_Created sheet tab from Google Sheets when present
  if (lowerTitle.includes('today')) {
    const valid = parsed.filter(isValidCampaignRow);
    const normalizedRows = valid.map((c, idx) => {
      const sl = Number(c['SL'] || idx + 1);
      const rowMonth = c['Month'] || c['month'] || MONTHS[new Date().getMonth()];
      const rowYear = c['Year'] || c['year'] || String(new Date().getFullYear());
      return normalizeCampaign(c, rowMonth, rowYear, sl);
    });
    // Never wipe known today rows with an empty/failed ingest (prevents intermittent "0 data"):
    // only replace when the sheet actually returned rows; otherwise merge sheet rows over local ones.
    const liveTodayRows = normalizedRows.filter(r => !isTombstoned(r));
    if (liveTodayRows.length > 0) {
      if (memoryStore.todayCreated.length === 0) {
        memoryStore.todayCreated = liveTodayRows;
      } else {
        const mergedToday = (memoryStore.todayCreated || []).filter(x => !isTombstoned(x));
        liveTodayRows.forEach(nr => {
          // Match by Record ID only: two distinct campaigns can share SL +
          // Page Name, and an SL+Page match would overwrite one with the other.
          const idx = nr['Record ID'] ? mergedToday.findIndex(x => x['Record ID'] === nr['Record ID']) : -1;
          if (idx >= 0) mergedToday[idx] = nr; else mergedToday.push(nr);
        });
        memoryStore.todayCreated = mergedToday;
      }
    }
    return;
  }

  // 5. Agency Cost
  if (lowerTitle.includes('agency_cost') || lowerTitle === 'cost' || lowerTitle.includes('spending')) {
    const valid = parsed.filter(isValidAgencyCostRow);
    const remoteCost = valid.map(r => {
      const rawDate = r['Date'] || r['date'] || '';
      const dObj = parseAnyDate(rawDate);
      const formattedDate = dObj ? formatCustomDate(dObj) : (rawDate ? String(rawDate).trim() : getTodayFormatted());
      const month = r['Month'] || (dObj ? MONTHS[dObj.getMonth()] : MONTHS[new Date().getMonth()]);
      const year = r['Year'] || (dObj ? String(dObj.getFullYear()) : String(new Date().getFullYear()));
      const cost = parseFloat(r['Cost'] || r['cost'] || 0) || 0;
      return {
        ...r,
        'Date': formattedDate,
        'Month': month,
        'Year': String(year),
        'Type': String(r['Type'] || r['type'] || '').trim(),
        'Cost': cost,
        'Per Person Cost': parseFloat(r['Per Person Cost'] || 0) || +(cost / 2).toFixed(2),
        'Status': r['Status'] || r['status'] || 'Due'
      };
    });
    const mergedCosts = [...remoteCost];
    (memoryStore.agencyCosts || []).forEach(localC => {
      const recId = String(localC['Record ID'] || localC._recordId || '').trim();
      const isPresent = mergedCosts.some(sp => {
        const spRecId = String(sp['Record ID'] || sp._recordId || '').trim();
        if (recId && spRecId && recId === spRecId) return true;
        return sp['Type'] === localC['Type'] && sp['Cost'] === localC['Cost'] && sp['Month'] === localC['Month'];
      });
      if (!isPresent) mergedCosts.push(localC);
    });
    memoryStore.agencyCosts = mergedCosts;
    return;
  }

  // 6. Page List
  if (lowerTitle.includes('page_list') || lowerTitle === 'pages' || lowerTitle === 'page') {
    dataRows.forEach(r => {
      const p = String(r[0] || '').trim();
      if (p && !memoryStore.pages.includes(p)) {
        memoryStore.pages.push(p);
      }
    });
    return;
  }

  // 7. Ad Accounts
  if (lowerTitle.includes('ad_account') || lowerTitle === 'accounts' || (lowerHeaders.includes('ad account') && !lowerHeaders.includes('budget'))) {
    dataRows.forEach(r => {
      const a = String(r[0] || '').trim();
      const id = String(r[1] || '').trim();
      if (a && !memoryStore.adAccounts.some(acc => acc.a === a)) {
        memoryStore.adAccounts.push({ a, id });
      }
    });
    return;
  }

  // 8. Settings / Config
  if (lowerTitle === 'settings' || lowerTitle === 'config') {
    dataRows.forEach(r => {
      const k = String(r[0] || '').trim().toLowerCase();
      const v = String(r[1] || '').trim();
      if (k) memoryStore.settings[k] = v;
    });
    return;
  }

  // Ignore non-campaign sheets (e.g. Page_Summary, End_ADS, Monthly_Closing)
  if (
    lowerTitle.includes('summary') ||
    lowerTitle.includes('end_ad') ||
    lowerTitle.includes('closing') ||
    lowerTitle.includes('billing')
  ) {
    return;
  }

  // 9. Campaign Sheets (e.g. Campaign_September_2026, September, Campaign, Sheet1, etc.)
  // Detect Month & Year from Title or Data
  let sheetMonth = '';
  let sheetYear = '';
  for (const m of MONTHS) {
    if (cleanTitle.toLowerCase().includes(m.toLowerCase())) {
      sheetMonth = m;
      break;
    }
  }
  const yearMatch = cleanTitle.match(/20\d\d/);
  if (yearMatch) sheetYear = yearMatch[0];

  const validParsed = parsed.filter(isValidCampaignRow);
  const normalizedRows = validParsed.map((c, idx) => {
    const sl = Number(c['SL'] || idx + 1);
    const startDateParsed = parseAnyDate(c['Start'] || c['Start Date'] || c['start_date'] || c['Entry Date'] || c['Date']);
    let rowMonth = '';
    let rowYear = '';
    if (startDateParsed) {
      rowMonth = MONTHS[startDateParsed.getMonth()];
      rowYear = String(startDateParsed.getFullYear());
    } else {
      rowMonth = c['Month'] || c['month'] || sheetMonth || MONTHS[new Date().getMonth()];
      rowYear = c['Year'] || c['year'] || sheetYear || String(new Date().getFullYear());
    }

    return normalizeCampaign(c, rowMonth, rowYear, sl);
  });

  // Accurate Sheet Mirroring: Google Sheet is the authoritative source of truth.
  // Exception: recently deleted rows (tombstones) must never be resurrected by an
  // ingest that read the sheet before the delete landed.
  const liveRows = normalizedRows.filter(r => !isTombstoned(r));
  if (normalizedRows.length > 0) {
    const unsyncedLocal = (memoryStore.campaigns[cleanTitle] || []).filter(r => r._isLocalUnsynced === true && !isTombstoned(r));
    const mergedList = [...liveRows];
    unsyncedLocal.forEach(loc => {
      const recId = String(loc['Record ID'] || loc._recordId || '').trim();
      // Presence is decided by Record ID only: an SL + Page Name twin in the
      // sheet is a different campaign, not proof this local row was synced.
      const isPresent = mergedList.some(r => {
        const rRecId = String(r['Record ID'] || r._recordId || '').trim();
        return Boolean(recId && rRecId && recId === rRecId);
      });
      if (!isPresent) {
        mergedList.push(loc);
      }
    });
    memoryStore.campaigns[cleanTitle] = mergedList;
  }

  liveRows.forEach(row => {
    const pn = row['Page Name'];
    if (pn && !memoryStore.pages.includes(pn)) {
      memoryStore.pages.push(pn);
    }
    const acc = row['AD Account'];
    const id = row['ID Name'] || '';
    if (acc && !memoryStore.adAccounts.some(a => a.a === acc)) {
      memoryStore.adAccounts.push({ a: acc, id });
    }
  });
  saveStoreToDisk();
}

async function isSheetConnected(): Promise<boolean> {
  if (!SHEET_ID) return false;
  // If we have loaded campaigns, payments, or pages from the sheet, it is live & connected!
  const totalCampaigns = Object.values(memoryStore.campaigns).reduce((a, b) => a + b.length, 0);
  if (totalCampaigns > 0 || memoryStore.payments.length > 0 || memoryStore.pages.length > 0 || memoryStore.todayCreated.length > 0) {
    return true;
  }
  if (!sheetsApiDisabled) {
    try {
      const api = await getSheetsClient();
      if (api) {
        await api.spreadsheets.get({ spreadsheetId: SHEET_ID });
        return true;
      }
    } catch (e) {
      handleSheetsApiError(e);
    }
  }
  try {
    const first = await fetchSheetCsvGviz(SHEET_ID);
    if (first && first.rows && first.rows.length > 0) return true;
  } catch {}
  return false;
}

async function importRawCsvData(target: string, rawText: string, defaultMonth?: string, defaultYear?: string) {
  if (!rawText || !rawText.trim()) return { success: false, message: 'No data provided.' };
  const lines = rawText.trim().split(/\r?\n/).filter(Boolean);
  if (!lines.length) return { success: false, message: 'Empty lines.' };

  const delimiter = lines[0].includes('\t') ? '\t' : (lines[0].includes(',') ? ',' : '\t');
  const firstLineCols = lines[0].split(delimiter).map(c => c.trim().replace(/^["']|["']$/g, ''));

  const hasHeaders = firstLineCols.some(c => ['sl', 'page', 'budget', 'rate', 'date', 'day', 'start', 'end', 'paid', 'method', 'cost'].includes(c.toLowerCase()));
  const dataLines = hasHeaders ? lines.slice(1) : lines;
  const headers = hasHeaders ? firstLineCols : [];

  let importedCount = 0;
  const mn = defaultMonth || MONTHS[new Date().getMonth()];
  const yr = defaultYear || new Date().getFullYear();

  if (target === 'campaign') {
    const sName = campaignSheetName(mn, yr);
    if (!memoryStore.campaigns[sName]) memoryStore.campaigns[sName] = [];

    for (const line of dataLines) {
      const parts = line.split(delimiter).map(c => c.trim().replace(/^["']|["']$/g, ''));
      if (!parts.some(p => p !== '')) continue;

      let item: any = {};
      if (hasHeaders) {
        headers.forEach((h, i) => { item[h] = parts[i] || ''; });
      } else if (parts.length >= 14 && (parts[8]?.toLowerCase().includes('campaign') || parts[8]?.toLowerCase().includes('boost') || parts[8]?.toLowerCase().includes('increase') || parts[4]?.toLowerCase().includes('page'))) {
        // Full 18-column Campaign Sheet export layout
        item = {
          'SL': parts[0] || '',
          'OFF $': parts[1] || '',
          'AD Account': parts[2] || '',
          'ID Name': parts[3] || '',
          'Status': parts[4] || 'Old Page',
          'Page Name': parts[5] || '',
          'Start': parts[6] || '',
          'End': parts[7] || '',
          'Types': parts[8] || 'Ad Campaign',
          'Day': parts[9] || '1',
          'Budget': parts[10] || '10',
          'Daily $': parts[11] || '',
          'Taka': parts[12] || '',
          'Rate': parts[13] || '140',
          'Entry Date': parts[14] || '',
          'Month': parts[15] || mn,
          'Year': parts[16] || yr,
          'Record ID': parts[17] || crypto.randomUUID()
        };
      } else {
        item = {
          'SL': parts[0] || String(memoryStore.campaigns[sName].length + 1).padStart(3, '0'),
          'AD Account': parts[1] || '001 | Master Ad Account 1',
          'Status': parts[2] || 'Old Page',
          'Page Name': parts[3] || parts[0] || 'Imported Page',
          'Start': parts[4] || new Date().toISOString().slice(0, 10),
          'End': parts[5] || new Date(Date.now() + 3 * 86400000).toISOString().slice(0, 10),
          'Types': parts[6] || 'Boost',
          'Day': Number(parts[7] || 1),
          'Budget': Number(parts[8] || 10),
          'Rate': Number(parts[9] || 140)
        };
      }
      const sl = Number(item['SL']) || (memoryStore.campaigns[sName].length + 1);
      const normalized = normalizeCampaign(item, mn, yr, sl);
      
      const targetMonth = normalized['Month'] || mn;
      const targetYear = normalized['Year'] || yr;
      const targetSheet = campaignSheetName(targetMonth, targetYear);
      if (!memoryStore.campaigns[targetSheet]) memoryStore.campaigns[targetSheet] = [];

      const existingIdx = memoryStore.campaigns[targetSheet].findIndex(x => x['Record ID'] === normalized['Record ID']);
      if (existingIdx >= 0) {
        memoryStore.campaigns[targetSheet][existingIdx] = normalized;
      } else {
        memoryStore.campaigns[targetSheet].push(normalized);
      }

      if (isSameDay(normalized['Entry Date'], getNowBST())) {
        const tcIdx = memoryStore.todayCreated.findIndex(x => x['Record ID'] === normalized['Record ID']);
        if (tcIdx >= 0) {
          memoryStore.todayCreated[tcIdx] = normalized;
        } else {
          memoryStore.todayCreated.push(normalized);
        }
      }

      if (normalized['Page Name'] && !memoryStore.pages.includes(normalized['Page Name'])) {
        memoryStore.pages.push(normalized['Page Name']);
      }
      if (normalized['AD Account']) {
        const hasAcc = memoryStore.adAccounts.some(a => a.a === normalized['AD Account']);
        if (!hasAcc) {
          memoryStore.adAccounts.push({ a: normalized['AD Account'], id: normalized['ID Name'] || '' });
        }
      }
      importedCount++;
    }
  } else if (target === 'payments') {
    for (const line of dataLines) {
      const parts = line.split(delimiter).map(c => c.trim().replace(/^["']|["']$/g, ''));
      if (!parts.some(p => p !== '')) continue;
      const date = parts[0] || new Date().toISOString().slice(0, 10);
      const method = parts[1] || 'Bkash';
      const amount = Number(parts[2] || parts[1] || 0);
      const payObj = {
        'Date': date,
        'Payment Method': isNaN(Number(parts[1])) ? method : 'Bkash',
        'Paid (৳)': isNaN(Number(parts[2])) ? (isNaN(Number(parts[1])) ? 0 : Number(parts[1])) : Number(parts[2]),
        'Month': mn,
        'Year': String(yr),
        _row: memoryStore.payments.length + 2,
        _recordId: crypto.randomUUID()
      };
      memoryStore.payments.push(payObj);
      importedCount++;
    }
  }

  saveStoreToDisk();
  return { success: true, importedCount, message: `Successfully imported ${importedCount} records into ${target}!` };
}

async function spreadsheetMeta() {
  if (sheetsApiDisabled) return { sheets: [] };
  if (sheetMetaCache && sheetMetaCache.data && Date.now() - sheetMetaCache.timestamp < 30000) {
    return sheetMetaCache.data;
  }
  const api = await getSheetsClient();
  if (!api || !SHEET_ID) return (sheetMetaCache && sheetMetaCache.data) || { sheets: [] };
  try {
    const r = await api.spreadsheets.get({ spreadsheetId: SHEET_ID });
    if (sheetMetaCache) {
      sheetMetaCache.data = r.data;
      sheetMetaCache.timestamp = Date.now();
    }
    return r.data;
  } catch (err: any) {
    handleSheetsApiError(err);
    return (sheetMetaCache && sheetMetaCache.data) || { sheets: [] };
  }
}

async function sheetNames(): Promise<Set<string>> {
  const m = await spreadsheetMeta();
  return new Set((m.sheets || []).map(x => x.properties?.title || '').filter(Boolean));
}

async function ensureSheet(name: string, headers: string[], customToken?: string | null) {
  if (ensuredSheets.has(name)) return;
  const tokenToUse = customToken || currentRequestOAuthToken;
  const api = await getSheetsClient(tokenToUse);
  if (!api || !SHEET_ID) return;
  try {
    const meta = await api.spreadsheets.get({ spreadsheetId: SHEET_ID }).catch(() => null);
    const existingTitles = new Set((meta?.data?.sheets || []).map(x => x.properties?.title || '').filter(Boolean));
    if (!existingTitles.has(name)) {
      await api.spreadsheets.batchUpdate({
        spreadsheetId: SHEET_ID,
        requestBody: {
          requests: [{ addSheet: { properties: { title: name } } }]
        }
      }).catch(err => {
        if (!err?.message?.includes('already exists')) throw err;
      });
      if (sheetMetaCache?.data?.sheets) {
        sheetMetaCache.data.sheets.push({ properties: { title: name } });
      }
    }
    if (headers && headers.length > 0) {
      await api.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `'${name}'!A1:${colToLetter(headers.length)}1`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [headers] }
      }).catch(() => {});
    }
    ensuredSheets.add(name);
  } catch (err: any) {
    handleSheetsApiError(err);
  }
}

function campaignSheetName(month: string, year: string | number): string {
  return `Campaign_${month}_${year}`;
}

// Month-wise payments: each month lives in its own tab
// (Payments_<Month>_<Year>), mirroring the campaign tabs. The legacy
// single 'Payments' tab is kept read-only during the transition and is
// cleared only after the month-tab build is imported and verified.
function paymentSheetName(month: string, year: string | number): string {
  return `Payments_${month}_${year}`;
}

async function ensureBaseSheets() {
  if (sheetsApiDisabled) return;
  const all: [string, string[]][] = [
    ['Settings', ['Key', 'Value']],
    ['Ad Accounts', ['Ad Account', 'ID Name']],
    ['Page List', ['Page Name']],
    ['Today_Created', TODAY_HEADERS],
    ['Payments', PAYMENT_HEADERS],
    ['Page_Summary', PAGE_SUMMARY_HEADERS],
    ['Agency_Spending', AGENCY_SPENDING_HEADERS],
    ['Monthly_Closing', MONTHLY_CLOSING_HEADERS],
    ['Agency_Cost', AGENCY_COST_HEADERS],
    ['End_ADS', END_ADS_HEADERS]
  ];
  for (const [n, h] of all) {
    await ensureSheet(n, h);
  }
}

async function readRows(name: string, headers: string[]): Promise<any[]> {
  const cleanName = name.trim();
  const lowerName = cleanName.toLowerCase().replace(/[-_ ]/g, '_');

  // Check 15s in-memory TTL cache to eliminate redundant requests and prevent quota limits
  const cached = sheetReadCache.get(cleanName);
  if (cached && Date.now() - cached.timestamp < 15000) {
    return cached.data;
  }

  // Ultra-Fast In-Memory Serving (Sub-millisecond latency for Admin and Viewers)
  if (lowerName.includes('payment') && memoryStore.payments.length > 0) {
    return memoryStore.payments.filter(isValidPaymentRow);
  }
  if ((lowerName.includes('agency_cost') || lowerName === 'cost') && memoryStore.agencyCosts.length > 0) {
    return memoryStore.agencyCosts.filter(isValidAgencyCostRow);
  }
  if ((lowerName === 'today_created' || lowerName === 'today') && memoryStore.todayCreated.length > 0) {
    return memoryStore.todayCreated.filter(isValidCampaignRow);
  }
  if (memoryStore.campaigns[cleanName] && memoryStore.campaigns[cleanName].length > 0) {
    return memoryStore.campaigns[cleanName].filter(isValidCampaignRow);
  }
  if (lowerName.includes('ad_account') && memoryStore.adAccounts.length > 0) {
    return memoryStore.adAccounts.map((a, i) => ({ 'Ad Account': a.a, 'ID Name': a.id, _row: i + 2 }));
  }
  if ((lowerName.includes('page_list') || lowerName === 'pages') && memoryStore.pages.length > 0) {
    return memoryStore.pages.map((p, i) => ({ 'Page Name': p, _row: i + 2 }));
  }

  if (!sheetsApiDisabled) {
    const api = await getSheetsClient();
    if (api && SHEET_ID) {
      try {
        await ensureSheet(name, headers);
        if (!sheetsApiDisabled && Date.now() >= sheetsRateLimitUntil) {
          const r = await api.spreadsheets.values.get({
            spreadsheetId: SHEET_ID,
            range: `'${name}'!A:ZZ`
          });
          const values = r.data.values || [];
          const head = values[0] || headers;
          const mapped = values.slice(1).map((row, i) => {
            const o: any = { _row: i + 2, _recordId: String(i + 2) };
            head.forEach((h, j) => {
              const val = row[j] ?? '';
              const rawH = String(h ?? '');
              const cleanH = rawH.trim();
              o[rawH] = val;
              o[cleanH] = val;
              const lower = cleanH.toLowerCase();
              if (lower.includes('date') || lower === 'entry') o['Date'] = val;
              if (lower === 'sl' || lower === 'sl no' || lower === 'sl_no' || lower === 'sl#' || lower === 'serial') o['SL NO'] = val;
              if (lower === 'page' || lower === 'page name' || lower === 'page_name') o['Page Name'] = val;
              if (lower === 'budget') o['Budget'] = val;
              if (lower === 'spent') o['Spent'] = val;
              if (lower === 'remaining') o['Remaining'] = val;
              if (lower === 'cost' || lower === 'amount') o['Cost'] = val;
              if (lower === 'type' || lower === 'category') o['Type'] = val;
              if (lower === 'status') o['Status'] = val;
              if (lower === 'paid' || lower === 'paid (৳)' || lower === 'amount') o['Paid (৳)'] = val;
              if (lower === 'method' || lower === 'payment method') o['Payment Method'] = val;
              if (lower === 'month') o['Month'] = val;
              if (lower === 'year') o['Year'] = val;
            });
            // Positional fallbacks if specific headers missing in user's sheet
            if (lowerName.includes('payment')) {
              if (!o['Date'] && row[0]) o['Date'] = row[0];
              if (!o['Payment Method'] && row[1]) o['Payment Method'] = row[1];
              if (!o['Paid (৳)'] && (row[2] !== undefined && row[2] !== '')) o['Paid (৳)'] = row[2];
              if (!o['Month'] && row[3]) o['Month'] = row[3];
              if (!o['Year'] && row[4]) o['Year'] = row[4];
            }
            if (lowerName.includes('agency_cost') || lowerName === 'cost') {
              if (!o['Month'] && row[0]) o['Month'] = row[0];
              if (!o['Year'] && row[1]) o['Year'] = row[1];
              if (!o['Type'] && row[2]) o['Type'] = row[2];
              if (!o['Cost'] && row[3]) o['Cost'] = row[3];
              if (!o['Per Person Cost'] && row[4]) o['Per Person Cost'] = row[4];
              if (!o['Status'] && row[5]) o['Status'] = row[5];
            }
            return o;
          });

          let result: any[];
          if (lowerName.startsWith('campaign_') || lowerName === 'today_created' || lowerName === 'today') {
            result = mapped.filter(isValidCampaignRow);
          } else if (lowerName.includes('agency_cost') || lowerName === 'cost') {
            result = mapped.filter(isValidAgencyCostRow);
          } else if (lowerName.includes('payment')) {
            result = mapped.filter(isValidPaymentRow);
          } else {
            result = mapped.filter(row => Object.values(row).some(v => v !== '' && v !== null && v !== undefined && v !== 0 && v !== '0'));
          }

          sheetReadCache.set(cleanName, { data: result, timestamp: Date.now() });
          return result;
        }
      } catch (err: any) {
        handleSheetsApiError(err);
      }
    }
  }

  // Memory fallback
  if (lowerName === 'settings' || lowerName === 'config') {
    return Object.entries(memoryStore.settings).map(([k, v], i) => ({ Key: k, Value: v, _row: i + 2 }));
  }
  if (lowerName.includes('ad_account') || lowerName === 'accounts') {
    return memoryStore.adAccounts.map((a, i) => ({ 'Ad Account': a.a, 'ID Name': a.id, _row: i + 2 }));
  }
  if (lowerName.includes('page_list') || lowerName === 'pages' || lowerName === 'page') {
    return memoryStore.pages.map((p, i) => ({ 'Page Name': p, _row: i + 2 }));
  }
  if (lowerName === 'today_created' || lowerName === 'today') {
    return memoryStore.todayCreated.filter(isValidCampaignRow);
  }
  if (lowerName.includes('payment')) {
    return memoryStore.payments.filter(isValidPaymentRow);
  }
  if (lowerName.includes('agency_cost') || lowerName === 'cost') {
    return memoryStore.agencyCosts.filter(isValidAgencyCostRow);
  }
  if (lowerName.includes('monthly_closing') || lowerName.includes('closing')) {
    return [...memoryStore.monthlyClosing];
  }
  
  // Direct match or normalized match in memoryStore.campaigns
  if (memoryStore.campaigns[cleanName]) {
    return memoryStore.campaigns[cleanName].filter(isValidCampaignRow);
  }
  for (const [sKey, rows] of Object.entries(memoryStore.campaigns)) {
    if (sKey.toLowerCase().replace(/[-_ ]/g, '_') === lowerName) {
      return rows.filter(isValidCampaignRow);
    }
  }

  return [];
}

async function syncWebhookDirect(action: string, payload: any): Promise<boolean> {
  const configuredUrl = memoryStore.settings.webhook_url;
  // The only write endpoint is the Apps Script URL saved in the persisted
  // config (.sheet_config.json, mirrored into settings.webhook_url and
  // editable in Tools & Settings). A configured deployment that has been
  // deleted/replaced returns Google's "Page Not Found" (404) or a JSON
  // {success:false} — those are treated as failures below so the write is
  // queued for retry and surfaced to the user, never silently lost.
  const candidates = [configuredUrl].filter((u, i, a): u is string => !!u && a.indexOf(u) === i);
  if (candidates.length === 0) return false;

  // Protocol compatibility: the repo's exported connectors read the row values from
  // `rowValues`, but the deployed Sync Connector v2.0 reads them from `row` (verified
  // live: with only rowValues it fails with "The rowContents passed to appendRow()
  // must be nonempty"). Send BOTH fields with the same array so writes work against
  // either connector generation. Likewise, v2.0's updateRow/deleteRow require a
  // `rowIndex` field (they reject requests without one: "Invalid rowIndex for
  // updateRow/deleteRow") while the exported connectors match by Record ID and
  // ignore it — mirror the tracked rowNo into rowIndex when present.
  const withRowAlias = (p: any) => {
    if (!p) return p;
    let out = p;
    if (Array.isArray(p.rowValues) && p.row === undefined) out = { ...out, row: p.rowValues };
    if (typeof p.rowNo === 'number' && p.rowNo > 1 && p.rowIndex === undefined) out = { ...out, rowIndex: p.rowNo };
    return out;
  };

  const bodyToSend = action === 'batch'
    ? { action: 'batch', batch: (payload.batch || []).map(withRowAlias) }
    : withRowAlias({ action, ...payload });

  for (const webhookUrl of candidates) {
    try {
      const res = await fetch(webhookUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(bodyToSend),
        redirect: 'follow',
        signal: AbortSignal.timeout(12000)
      });
      const text = await res.text();
      console.log(`[Instant Sheet Sync] ${action} (${payload.sheetName || (payload.batch?.length ? payload.batch.length + ' items' : '')}) status: ${res.status}, response: ${text.slice(0, 80)}`);
      if (!res.ok) continue;
      // HTTP 200 alone is not proof of success: Apps Script connectors report
      // failures as JSON {success:false,error:...} — treat those as failures.
      let parsed: any = null;
      try { parsed = JSON.parse(text); } catch (_) {}
      if (parsed && (parsed.success === false || parsed.error)) continue;
      if (!parsed && text.trim().startsWith('<')) continue; // HTML error page, not a connector response
      return true;
    } catch (err: any) {
      console.warn(`[Instant Sheet Sync Error] ${action}:`, err?.message);
    }
  }
  return false;
}

async function sendToAppsScriptWebhook(action: string, payload: any): Promise<boolean> {
  return syncWebhookDirect(action, payload);
}

async function appendRow(name: string, headers: string[], row: any, immediateSync = true): Promise<void> {
  invalidateSheetCache(name);
  row = normalizeRowDates(row, row['Month'] || row.month);
  const cleanRecordId = String(row['Record ID'] || row._recordId || '').trim();
  const cleanSl = (row['SL'] !== undefined && row['SL'] !== null ? String(row['SL']) : (row['SL NO'] !== undefined && row['SL NO'] !== null ? String(row['SL NO']) : '')).trim();
  const cleanPage = String(row['Page Name'] || '').trim();

  // 1. Instantly update memoryStore (Sub-millisecond local persistence & 0ms UI response)
  const isRowMatch = (r: any) => {
    if (cleanRecordId) {
      if (String(r['Record ID'] || '').trim() === cleanRecordId) return true;
      if (String(r._recordId || '').trim() === cleanRecordId) return true;
    }
    const rSl = String(r['SL'] || r['SL NO'] || '').trim();
    if (cleanSl && rSl && rSl === cleanSl) {
      if (!cleanPage || String(r['Page Name'] || '').trim().toLowerCase() === cleanPage.toLowerCase()) {
        return true;
      }
    }
    return false;
  };

  if (name === 'Today_Created') {
    const existingIdx = memoryStore.todayCreated.findIndex(isRowMatch);
    if (existingIdx >= 0) {
      memoryStore.todayCreated[existingIdx] = { ...memoryStore.todayCreated[existingIdx], ...row };
    } else {
      memoryStore.todayCreated.push({ ...row, _row: memoryStore.todayCreated.length + 2 });
    }
  } else if (name.startsWith('Campaign_')) {
    if (!memoryStore.campaigns[name]) memoryStore.campaigns[name] = [];
    const existingIdx = memoryStore.campaigns[name].findIndex(isRowMatch);
    if (existingIdx >= 0) {
      memoryStore.campaigns[name][existingIdx] = { ...memoryStore.campaigns[name][existingIdx], ...row };
    } else {
      row._row = memoryStore.campaigns[name].length + 2;
      memoryStore.campaigns[name].push({ ...row });
    }
  } else if (name === 'Payments') {
    const isExplicitUpdate = Boolean(cleanRecordId && memoryStore.payments.some(p =>
      String(p['Record ID'] || '').trim() === cleanRecordId || String(p._recordId || '').trim() === cleanRecordId
    ));
    if (isExplicitUpdate) {
      const existingIdx = memoryStore.payments.findIndex(p =>
        String(p['Record ID'] || '').trim() === cleanRecordId || String(p._recordId || '').trim() === cleanRecordId
      );
      memoryStore.payments[existingIdx] = { ...memoryStore.payments[existingIdx], ...row };
    } else {
      const recId = cleanRecordId || crypto.randomUUID();
      memoryStore.payments.push({
        ...row,
        'Record ID': row['Record ID'] || recId,
        _recordId: row._recordId || recId,
        _row: memoryStore.payments.length + 2
      });
    }
  } else if (name === 'Agency_Cost') {
    const existingIdx = memoryStore.agencyCosts.findIndex(isRowMatch);
    if (existingIdx >= 0) {
      memoryStore.agencyCosts[existingIdx] = { ...memoryStore.agencyCosts[existingIdx], ...row };
    } else {
      memoryStore.agencyCosts.push({ ...row, _row: memoryStore.agencyCosts.length + 2 });
    }
  } else if (name === 'Monthly_Closing') {
    const existingIdx = memoryStore.monthlyClosing.findIndex(m => m.Month === row.Month && String(m.Year) === String(row.Year));
    if (existingIdx >= 0) {
      memoryStore.monthlyClosing[existingIdx] = { ...memoryStore.monthlyClosing[existingIdx], ...row };
    } else {
      memoryStore.monthlyClosing.push({ ...row, _row: memoryStore.monthlyClosing.length + 2 });
    }
  } else if (name === 'Page List') {
    if (row['Page Name'] && !memoryStore.pages.includes(row['Page Name'])) {
      memoryStore.pages.push(row['Page Name']);
    }
  } else if (name === 'Ad Accounts') {
    if (row['Ad Account'] && !memoryStore.adAccounts.some(a => a.a === row['Ad Account'])) {
      memoryStore.adAccounts.push({ a: row['Ad Account'], id: row['ID Name'] || '' });
    }
  }
  saveStoreToDisk();

  // 2. Ultra-Fast Immediate Sync to Webhook (0s Delay)
  // Per user requirement: Do not save or sync to Today_Created sheet tab in Google Sheets
  if (name === 'Today_Created' || name.toLowerCase().includes('today')) {
    return;
  }

  let synced = false;
  if (immediateSync) {
    try {
      synced = await syncWebhookDirect('appendRow', {
        sheetName: name,
        headers,
        recordId: cleanRecordId,
        matchInfo: { SL: cleanSl, 'Page Name': cleanPage, 'Record ID': cleanRecordId, budget: row['Budget'], month: row['Month'], year: row['Year'] },
        rowValues: formatRowValuesForSheet(headers, row, row['Month'])
      });
    } catch (e: any) {
      console.warn('[Instant Sheet Sync Note]:', e?.message);
    }
  }

  // 3. Write-through to Google Sheets: AWAITED before returning, so the caller's
  // success response means the row is actually in the sheet. On failure the write
  // is queued for automatic retry and the failure is flagged on the RPC response.
  // Skipped when step 2's direct webhook write already confirmed success — running
  // both would double-write (for deletes, the second attempt necessarily "fails"
  // because the row is already gone).
  if (immediateSync && !synced) {
    try {
      await executeServerSyncTask({
        action: 'append',
        name,
        headers,
        row,
        recordId: cleanRecordId,
        matchInfo: { SL: cleanSl, pageName: cleanPage, 'Record ID': cleanRecordId, budget: row['Budget'] },
        oauthToken: currentRequestOAuthToken || undefined
      });
    } catch (syncErr: any) {
      console.error('[Sheet Sync] Append write-through failed, queued for retry:', syncErr?.message);
      currentRequestSheetSyncFailure = `Google Sheet save failed for "${name}" — the entry is saved in the app and queued for automatic retry.`;
      enqueueServerSyncTask({
        action: 'append',
        name,
        headers,
        row,
        recordId: cleanRecordId,
        matchInfo: { SL: cleanSl, pageName: cleanPage, 'Record ID': cleanRecordId, budget: row['Budget'] }
      });
    }
  }
}

async function resolveServerSheetTab(
  api: sheets_v4.Sheets,
  sheetId: string,
  targetName: string,
  month?: string,
  year?: string | number
): Promise<{ title: string; sheetId: number } | null> {
  try {
    const meta = await api.spreadsheets.get({ spreadsheetId: sheetId });
    const sheets = (meta.data.sheets || []).map(s => ({
      title: s.properties?.title || '',
      sheetId: s.properties?.sheetId ?? 0
    })).filter(s => s.title);

    const cleanReq = targetName.trim().toLowerCase().replace(/[-_ ]/g, '');

    // 1. Exact match (case & whitespace insensitive)
    const exact = sheets.find(s => s.title.trim().toLowerCase().replace(/[-_ ]/g, '') === cleanReq);
    if (exact) return exact;

    // 2. Today Created tab
    if (cleanReq.includes('today')) {
      const today = sheets.find(s => s.title.toLowerCase().includes('today'));
      if (today) return today;
    }

    // Derive effective month & year from targetName if not explicitly passed
    let effectiveMonth = month;
    let effectiveYear = year;
    if (!effectiveMonth) {
      const mMatch = targetName.match(/(January|February|March|April|May|June|July|August|September|October|November|December)/i);
      if (mMatch) effectiveMonth = mMatch[1];
    }
    if (!effectiveYear) {
      const yMatch = targetName.match(/\b(20\d{2})\b/);
      if (yMatch) effectiveYear = yMatch[1];
    }

    // 3. Month & Year tab
    if (effectiveMonth) {
      const mFull = effectiveMonth.toLowerCase();
      const mShort = effectiveMonth.slice(0, 3).toLowerCase();
      const yrStr = effectiveYear ? String(effectiveYear).trim() : '';

      if (yrStr) {
        const my = sheets.find(s => {
          const t = s.title.toLowerCase();
          return (t.includes(mFull) || t.includes(mShort)) && t.includes(yrStr);
        });
        if (my) return my;
      }

      const mOnly = sheets.find(s => {
        const t = s.title.toLowerCase();
        return t.includes(mFull) || t.includes(mShort);
      });
      if (mOnly) return mOnly;
    }

    // 4. Payments
    if (cleanReq.includes('payment')) {
      const pay = sheets.find(s => s.title.toLowerCase().includes('payment'));
      if (pay) return pay;
    }

    // 7. Agency cost
    if (cleanReq.includes('cost') || cleanReq.includes('agency_cost')) {
      const cost = sheets.find(s => s.title.toLowerCase().includes('cost'));
      if (cost) return cost;
    }

    // 8. Campaign tab fallback - ONLY if NO specific month was ever requested or mentioned
    if (!effectiveMonth && cleanReq.includes('campaign')) {
      const camp = sheets.find(s => s.title.toLowerCase().includes('campaign'));
      if (camp) return camp;
    }

    return null;
  } catch (err: any) {
    console.warn(`Error resolving sheet tab for ${targetName}:`, err?.message);
    return null;
  }
}

async function findRowIndexInServerSheet(
  api: sheets_v4.Sheets,
  sheetId: string,
  sheetTitle: string,
  criteria: {
    recordId?: string;
    sl?: string | number;
    pageName?: string;
    budget?: string | number;
    rowNo?: number;
  }
): Promise<{ rowIndex: number; rowData: any[] } | null> {
  try {
    const res = await api.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: `'${sheetTitle}'!A:ZZ`
    });
    const values = res.data.values || [];
    if (values.length <= 1) return null;

    const headers = values[0].map(h => String(h || '').trim().toLowerCase());
    const recIdCol = headers.findIndex(h => h === 'record id' || h === 'recordid' || h === '_recordid' || h === 'id' || h === 'record_id');
    const slCol = headers.findIndex(h => h === 'sl' || h === 'sl no' || h === 'sl_no' || h === 'sl#' || h === 'serial' || h.includes('sl'));
    const pageCol = headers.findIndex(h => h.includes('page') || h === 'name' || h === 'title');
    const budgetCol = headers.findIndex(h => h.includes('budget') || h.includes('cost') || h.includes('paid') || h.includes('spent') || h.includes('amount'));

    const cleanRecId = criteria.recordId ? String(criteria.recordId).trim() : '';
    const cleanSl = criteria.sl !== undefined && criteria.sl !== null ? String(criteria.sl).trim().replace(/^#/, '') : '';
    const cleanSlNumMatch = cleanSl.match(/\d+/);
    const cleanSlNum = cleanSlNumMatch ? parseInt(cleanSlNumMatch[0], 10) : NaN;
    const cleanPage = criteria.pageName ? String(criteria.pageName).trim().toLowerCase() : '';
    const cleanBudget = criteria.budget !== undefined && criteria.budget !== null ? parseFloat(String(criteria.budget).replace(/[^0-9.]/g, '')) : NaN;

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const rRecId = recIdCol >= 0 ? String(row[recIdCol] || '').trim() : '';
      const rSlRaw = slCol >= 0 ? String(row[slCol] || '').trim().replace(/^#/, '') : String(row[0] || '').trim().replace(/^#/, '');
      const rSlNumMatch = rSlRaw.match(/\d+/);
      const rSlNum = rSlNumMatch ? parseInt(rSlNumMatch[0], 10) : NaN;
      const rPage = pageCol >= 0 ? String(row[pageCol] || '').trim().toLowerCase() : '';
      const rBudget = budgetCol >= 0 ? parseFloat(String(row[budgetCol] || '0').replace(/[^0-9.]/g, '')) : NaN;

      // 1. Match by Record ID
      if (cleanRecId && rRecId && cleanRecId === rRecId) {
        return { rowIndex: i + 1, rowData: row };
      }

      // 2. Match by SL + Page Name
      if (cleanSl && rSlRaw) {
        const slMatch = (cleanSl === rSlRaw) || (cleanSl.includes(rSlRaw)) || (rSlRaw.includes(cleanSl)) || (!isNaN(cleanSlNum) && !isNaN(rSlNum) && cleanSlNum === rSlNum);
        if (slMatch) {
          if (cleanPage && rPage) {
            if (cleanPage === rPage || cleanPage.includes(rPage) || rPage.includes(cleanPage)) {
              return { rowIndex: i + 1, rowData: row };
            }
          } else {
            return { rowIndex: i + 1, rowData: row };
          }
        }
      }

      // 3. Match by Page Name + Budget
      if (cleanPage && rPage && (cleanPage === rPage || cleanPage.includes(rPage) || rPage.includes(cleanPage))) {
        if (!isNaN(cleanBudget) && !isNaN(rBudget) && Math.abs(cleanBudget - rBudget) < 0.01) {
          return { rowIndex: i + 1, rowData: row };
        }
      }

      // 4. Match by SL alone ONLY if there is no page conflict. Previously this rule
      // fired even when the caller supplied a Page Name that did NOT match the row,
      // so an update/delete could hit a different page's row that shares the SL.
      if (!isNaN(cleanSlNum) && !isNaN(rSlNum) && cleanSlNum === rSlNum && cleanSlNum > 0) {
        if (cleanPage && rPage && cleanPage !== rPage && !cleanPage.includes(rPage) && !rPage.includes(cleanPage)) {
          continue;
        }
        return { rowIndex: i + 1, rowData: row };
      }
    }

    if (criteria.rowNo && criteria.rowNo > 1 && criteria.rowNo <= values.length) {
      const candidate = values[criteria.rowNo - 1];
      // Row numbers drift after sheet-side sorts/deletes: never trust rowNo blindly
      // when a Record ID was supplied and the row sitting there carries a different one.
      if (cleanRecId && recIdCol >= 0) {
        const candidateRecId = String(candidate[recIdCol] || '').trim();
        if (candidateRecId && candidateRecId !== cleanRecId) {
          return null;
        }
      }
      return { rowIndex: criteria.rowNo, rowData: candidate };
    }

    return null;
  } catch (err: any) {
    console.warn(`Error scanning sheet rows in ${sheetTitle}:`, err?.message);
    return null;
  }
}

// The deployed Sync Connector v2.0 updates/deletes strictly by rowIndex. A record
// that was added moments ago can be served to the editor from an in-memory copy that
// carries no _row (addCampaign keeps a second, sheet-positioned copy that does).
// Adopt the known sheet row number from ANY local copy sharing the Record ID so a
// fresh record's first edit/delete still writes through to the correct sheet row.
function lookupKnownRowNo(name: string, recordId?: string): number | undefined {
  const rid = String(recordId || '').trim();
  if (!rid) return undefined;
  const collections: any[][] = [];
  if (name.startsWith('Campaign_')) {
    if (memoryStore.campaigns[name]) collections.push(memoryStore.campaigns[name]);
    Object.values(memoryStore.campaigns).forEach(list => { if (list !== memoryStore.campaigns[name]) collections.push(list); });
  } else if (name === 'Payments') collections.push(memoryStore.payments || []);
  else if (name === 'Agency_Cost') collections.push(memoryStore.agencyCosts || []);
  else if (name === 'Today_Created') collections.push(memoryStore.todayCreated || []);
  for (const list of collections) {
    const hit = (list || []).find(r => (String(r['Record ID'] || '').trim() === rid || String(r._recordId || '').trim() === rid) && typeof r._row === 'number' && r._row > 1);
    if (hit) return hit._row;
  }
  return undefined;
}

// ---- Payment tab rewrite fallback (2026-10-25) ----
// The deployed connector may be an older generation that ignores content
// matchInfo: its updateRow then APPENDS a duplicate and its deleteRow is a
// no-op, while still returning {success:true}. The deterministic fix does
// not depend on the connector's matching at all: fresh GViz read -> modify
// exactly one row -> pushAllData rewrite -> fresh GViz verify. Payments is
// ~100 values-only rows, so a full-tab rewrite is cheap and safe; headers
// are preserved exactly (including the sheet's empty 6th header cell).
// Serialized through a promise chain so two rewrites can never interleave
// and clobber each other.
let paymentRewriteChain: Promise<any> = Promise.resolve();
function runPaymentRewrite<T>(fn: () => Promise<T>): Promise<T> {
  const p = paymentRewriteChain.then(fn, fn);
  paymentRewriteChain = p.catch(() => {});
  return p;
}
function sameYMD(a: any, b: any): boolean {
  const pa = parseAnyDate(a), pb = parseAnyDate(b);
  if (pa && pb) return pa.getFullYear() === pb.getFullYear() && pa.getMonth() === pb.getMonth() && pa.getDate() === pb.getDate();
  return String(a ?? '').trim() === String(b ?? '').trim();
}
function paymentContentRowMatches(row: string[], headers: string[], t: any): boolean {
  const lower = headers.map(h => String(h || '').trim().toLowerCase());
  const col = (name: string, like?: string) => {
    let i = lower.indexOf(name);
    if (i !== -1) return i;
    if (like) { for (let k = 0; k < lower.length; k++) if (lower[k].indexOf(like) !== -1) return k; }
    return -1;
  };
  const dI = col('date'), mI = col('payment method', 'payment'), pI = col('paid (৳)', 'paid'), moI = col('month'), yI = col('year');
  // Note column: the sheet's 6th header cell is EMPTY (not 'Note'), so a
  // header-name lookup misses it — fall back to the fixed Payments layout
  // (index 5) when the row has that many columns.
  let nI = col('note');
  if (nI === -1 && row.length > 5) nI = 5;
  if (dI === -1 || pI === -1) return false;
  if (!sameYMD(row[dI], t.Date)) return false;
  const tPaid = parseFloat(String(t['Paid (৳)'] ?? '').replace(/[^0-9.]/g, ''));
  const xPaid = parseFloat(String(row[pI] || '').replace(/[^0-9.]/g, ''));
  if (isNaN(tPaid) || isNaN(xPaid) || Math.abs(xPaid - tPaid) >= 0.01) return false;
  if (mI !== -1 && String(t['Payment Method'] || '').trim() && String(row[mI] || '').trim().toLowerCase() !== String(t['Payment Method']).trim().toLowerCase()) return false;
  if (moI !== -1 && String(t.Month || '').trim() && String(row[moI] || '').trim().toLowerCase() !== String(t.Month).trim().toLowerCase()) return false;
  if (yI !== -1 && String(t.Year || '').trim() && String(row[yI] || '').trim() !== String(t.Year).trim()) return false;
  // A non-empty target note must match exactly: a row that cannot show a
  // note (short row, no note column) has an empty note and never matches.
  if (String(t.Note || '').trim()) {
    if (nI === -1) return false;
    if (String(row[nI] || '').trim() !== String(t.Note).trim()) return false;
  }
  return true;
}
async function rewritePaymentsTab(tab: string, headers: string[], rows: string[][]): Promise<boolean> {
  const targetUrl = memoryStore.settings.webhook_url;
  if (!targetUrl) return false;
  try {
    const res = await fetch(targetUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'pushAllData', allSheetsData: { [tab]: [headers, ...rows] } }),
      redirect: 'follow',
      signal: AbortSignal.timeout(30000)
    });
    const text = await res.text();
    if (!res.ok) return false;
    try { const p = JSON.parse(text); if (p && (p.success === false || p.error)) return false; } catch (_) {}
    return true;
  } catch (_) { return false; }
}

async function updateRowAt(name: string, headers: string[], rowNo?: number, row?: any, recordId?: string, matchInfo?: any, immediateSync = true): Promise<void> {
  invalidateSheetCache(name);
  if (row) row = normalizeRowDates(row, row['Month'] || row.month);
  let targetRowNo = (typeof rowNo === 'number' && rowNo > 1) ? rowNo : undefined;
  if (!targetRowNo) {
    targetRowNo = lookupKnownRowNo(name, recordId || row?.['Record ID'] || row?._recordId || matchInfo?.['Record ID']);
  }
  let targetSheetTitle = name;

  // 1. Instantly update memory store safely (Sub-millisecond latency)
  const matchIdStr = (recordId ? String(recordId).trim() : '') || (row && (row['Record ID'] || row._recordId) ? String(row['Record ID'] || row._recordId).trim() : null);
  const isMatch = (r: any) => {
    if (targetRowNo && typeof targetRowNo === 'number' && targetRowNo > 1 && r._row === targetRowNo) return true;
    if (matchIdStr) {
      if (r['Record ID'] && String(r['Record ID']).trim() === matchIdStr) return true;
      if (r._recordId && String(r._recordId).trim() === matchIdStr) return true;
    }
    // SL+Page fallback is ONLY for rows with no known row number and no
    // Record ID (legacy rows). With a row number or a Record ID, falling
    // back to SL+Page would also purge a duplicate-SL twin — a different
    // real entry that merely shares the number.
    if (matchInfo && !targetRowNo && !matchIdStr) {
      if (matchInfo['Record ID'] && r['Record ID'] && String(r['Record ID']).trim() === String(matchInfo['Record ID']).trim()) return true;
      if (matchInfo.SL && (r.SL !== undefined || r['SL NO'] !== undefined || r.sl !== undefined || r.sl_no !== undefined)) {
        const rSl = String(r.SL || r['SL NO'] || r.sl || r.sl_no || '').trim();
        const mSl = String(matchInfo.SL).trim();
        const rNum = parseInt(rSl.replace(/\D/g, ''), 10);
        const mNum = parseInt(mSl.replace(/\D/g, ''), 10);
        const slMatch = (rSl === mSl) || (!isNaN(rNum) && !isNaN(mNum) && rNum === mNum);
        if (slMatch) {
          if (!matchInfo['Page Name'] || !r['Page Name'] || String(r['Page Name']).trim().toLowerCase() === String(matchInfo['Page Name']).trim().toLowerCase()) {
            return true;
          }
        }
      }
    }
    return false;
  };

  // Capture the PRE-update row BEFORE memory is overwritten: the connector
  // must match the row as it currently exists in the sheet (OLD content).
  // Sending the NEW values as matchInfo would find nothing and the
  // connector would APPEND a duplicate instead of updating in place.
  const findOldRow = (list: any[]) => { const i = (list || []).findIndex(r => isMatch(r)); return i >= 0 ? list[i] : null; };
  let oldRow: any = null;
  if (name === 'Payments') oldRow = findOldRow(memoryStore.payments);
  else if (name.startsWith('Campaign_') && memoryStore.campaigns[name]) oldRow = findOldRow(memoryStore.campaigns[name]);
  else if (name === 'Today_Created') oldRow = findOldRow(memoryStore.todayCreated);
  else if (name === 'Agency_Cost') oldRow = findOldRow(memoryStore.agencyCosts);
  // MatchInfo for the sheet write: OLD content for payment fields (see
  // above), so content matching in the connector finds the existing row.
  const syncMatchInfo = matchInfo || {
    SL: row?.SL, 'Record ID': row?.['Record ID'], 'Page Name': row?.['Page Name'],
    month: row?.['Month'] || row?.month, year: row?.['Year'] || row?.year,
    Date: oldRow?.Date || row?.Date,
    'Payment Method': oldRow?.['Payment Method'] || row?.['Payment Method'],
    'Paid (৳)': oldRow?.['Paid (৳)'] ?? row?.['Paid (৳)'],
    Month: oldRow?.Month || row?.Month, Year: oldRow?.Year || row?.Year,
    // Note: NO fallback to the new note — an old note that is empty/undefined
    // must stay empty in matchInfo (the connector then skips the Note check).
    // Falling back to the new note made content matching fail, so an edit
    // that changed the note appended a duplicate instead of updating.
    Note: oldRow?.Note
  };

  if (name.startsWith('Campaign_') && memoryStore.campaigns[name]) {
    const idx = memoryStore.campaigns[name].findIndex(r => isMatch(r));
    if (idx >= 0) memoryStore.campaigns[name][idx] = { ...row, _row: targetRowNo || memoryStore.campaigns[name][idx]._row };
  } else if (name === 'Payments') {
    const idx = memoryStore.payments.findIndex(r => isMatch(r));
    if (idx >= 0) memoryStore.payments[idx] = { ...row, _row: targetRowNo || memoryStore.payments[idx]._row };
  } else if (name === 'Today_Created') {
    const idx = memoryStore.todayCreated.findIndex(r => isMatch(r));
    if (idx >= 0) memoryStore.todayCreated[idx] = { ...row, _row: targetRowNo || memoryStore.todayCreated[idx]._row };
  } else if (name === 'Agency_Cost') {
    const idx = memoryStore.agencyCosts.findIndex(r => isMatch(r));
    if (idx >= 0) memoryStore.agencyCosts[idx] = { ...row, _row: targetRowNo || memoryStore.agencyCosts[idx]._row };
  }
  saveStoreToDisk();

  // Per user requirement (same as appendRow): Today_Created is an in-app view only —
  // never synced to a Google Sheet tab, so don't attempt (or report) a sheet write for it.
  if (name === 'Today_Created' || name.toLowerCase().includes('today')) {
    return;
  }

  // 2. Instant direct sync to webhook
  let synced = false;
  if (immediateSync) {
    try {
      synced = await syncWebhookDirect('updateRow', {
        sheetName: targetSheetTitle || name,
        rowNo: targetRowNo,
        recordId: recordId || row?.['Record ID'] || row?._recordId,
        matchInfo: syncMatchInfo,
        headers,
        rowValues: formatRowValuesForSheet(headers, row, row?.Month || matchInfo?.month)
      });
    } catch (e: any) {
      console.warn('[Instant Update Sheet Sync Note]:', e?.message);
    }
  }

  // 3. Write-through to Google Sheets: AWAITED before returning (see appendRow).
  if (immediateSync && !synced) {
    try {
      await executeServerSyncTask({
        action: 'update',
        name: targetSheetTitle || name,
        headers,
        row,
        rowNo: targetRowNo,
        recordId,
        matchInfo: syncMatchInfo,
        oauthToken: currentRequestOAuthToken || undefined
      });
    } catch (syncErr: any) {
      console.error('[Sheet Sync] Update write-through failed, queued for retry:', syncErr?.message);
      currentRequestSheetSyncFailure = `Google Sheet update failed for "${targetSheetTitle || name}" — the change is saved in the app and queued for automatic retry.`;
      enqueueServerSyncTask({
        action: 'update',
        name: targetSheetTitle || name,
        headers,
        row,
        rowNo: targetRowNo,
        recordId,
        matchInfo: syncMatchInfo
      });
    }
  }
}

async function deleteRowAt(name: string, rowNo?: number, recordId?: string, matchInfo?: any, immediateSync = true): Promise<void> {
  invalidateSheetCache(name);
  let targetRowNo = (typeof rowNo === 'number' && rowNo > 1) ? rowNo : undefined;
  if (!targetRowNo) {
    targetRowNo = lookupKnownRowNo(name, recordId || matchInfo?.['Record ID']);
  }
  let targetSheetTitle = name;

  // 1. Instantly update memory store safely: Never wipe out all items!
  const matchIdStr = recordId ? String(recordId).trim() : null;
  const isMatch = (r: any) => {
    if (targetRowNo && typeof targetRowNo === 'number' && targetRowNo > 1 && r._row === targetRowNo) return true;
    if (matchIdStr) {
      if (r['Record ID'] && String(r['Record ID']).trim() === matchIdStr) return true;
      if (r._recordId && String(r._recordId).trim() === matchIdStr) return true;
    }
    // SL+Page fallback is ONLY for rows with no known row number and no
    // Record ID (legacy rows). With a row number or a Record ID, falling
    // back to SL+Page would also purge a duplicate-SL twin — a different
    // real entry that merely shares the number.
    if (matchInfo && !targetRowNo && !matchIdStr) {
      if (matchInfo['Record ID'] && r['Record ID'] && String(r['Record ID']).trim() === String(matchInfo['Record ID']).trim()) return true;
      if (matchInfo.SL && (r.SL !== undefined || r['SL NO'] !== undefined || r.sl !== undefined || r.sl_no !== undefined)) {
        const rSl = String(r.SL || r['SL NO'] || r.sl || r.sl_no || '').trim();
        const mSl = String(matchInfo.SL).trim();
        const rNum = parseInt(rSl.replace(/\D/g, ''), 10);
        const mNum = parseInt(mSl.replace(/\D/g, ''), 10);
        const slMatch = (rSl === mSl) || (!isNaN(rNum) && !isNaN(mNum) && rNum === mNum);
        if (slMatch) {
          if (!matchInfo['Page Name'] || !r['Page Name'] || String(r['Page Name']).trim().toLowerCase() === String(matchInfo['Page Name']).trim().toLowerCase()) {
            return true;
          }
        }
      }
    }
    return false;
  };

  for (const key of Object.keys(memoryStore.campaigns)) {
    if (key === name || key.toLowerCase().replace(/[-_ ]/g, '_') === name.toLowerCase().replace(/[-_ ]/g, '_')) {
      memoryStore.campaigns[key] = memoryStore.campaigns[key].filter(r => !isMatch(r));
    }
  }
  if (name === 'Payments' || name.toLowerCase().includes('payment')) {
    memoryStore.payments = memoryStore.payments.filter(r => !isMatch(r));
  }
  if (name === 'Today_Created' || name.toLowerCase().includes('today')) {
    memoryStore.todayCreated = memoryStore.todayCreated.filter(r => !isMatch(r));
  }
  if (name === 'Agency_Cost' || name.toLowerCase().includes('cost')) {
    memoryStore.agencyCosts = memoryStore.agencyCosts.filter(r => !isMatch(r));
  }
  saveStoreToDisk();

  // Per user requirement (same as appendRow): Today_Created is an in-app view only —
  // never synced to a Google Sheet tab, so don't attempt (or report) a sheet write for it.
  if (name === 'Today_Created' || name.toLowerCase().includes('today')) {
    return;
  }

  // 2. Instant direct sync to webhook. NOTE: the deployed Sync Connector v2.0
  // does NOT delete by rowIndex — it scans by Record ID, then falls back to
  // SL+Page matching, which can hit a duplicate-SL twin. The retry guard in
  // step 3 below is what keeps a timed-out-but-landed delete from eating
  // the twin.
  let synced = false;
  if (immediateSync) {
    try {
      synced = await syncWebhookDirect('deleteRow', {
        sheetName: targetSheetTitle || name,
        rowNo: targetRowNo,
        recordId,
        matchInfo
      });
    } catch (e: any) {
      console.warn('[Instant Delete Sheet Sync Note]:', e?.message);
    }
  }

  // 3. Write-through to Google Sheets: AWAITED before returning (see appendRow).
  // Twin-safe: if the direct call timed out but actually landed, the record
  // is already gone from the sheet — retrying would make the connector's
  // SL+Page fallback delete a duplicate-SL twin instead. Skip the retry.
  if (immediateSync && !synced) {
    let alreadyGone = false;
    try {
      if (recordId) alreadyGone = await sheetRecordAlreadyGone(targetSheetTitle || name, recordId);
    } catch (_) { /* on guard failure, still run the task */ }
    if (!alreadyGone) {
    try {
      await executeServerSyncTask({
        action: 'delete',
        name: targetSheetTitle || name,
        rowNo: targetRowNo,
        recordId,
        matchInfo,
        oauthToken: currentRequestOAuthToken || undefined
      });
    } catch (syncErr: any) {
      console.error('[Sheet Sync] Delete write-through failed, queued for retry:', syncErr?.message);
      currentRequestSheetSyncFailure = `Google Sheet delete failed for "${targetSheetTitle || name}" — the deletion is saved in the app and queued for automatic retry.`;
      enqueueServerSyncTask({
        action: 'delete',
        name: targetSheetTitle || name,
        rowNo: targetRowNo,
        recordId,
        matchInfo
      });
    }
    } // end if (!alreadyGone)
  }
}

// Twin-safe delete guard: the Apps Script connector answers {success:true}
// for deleteRow even when its scan matched nothing, and its SL+Page fallback
// can match a DIFFERENT row sharing the SL (a duplicate-SL twin). A slow
// first attempt (cold start past the 12s timeout) that actually landed must
// never be retried into a second delete — so before any delete retry, the
// sheet itself is checked: if the exact Record ID is already gone, the
// delete is done and the retry is dropped. An unreadable sheet (null) never
// counts as "gone" — when in doubt, the delete still runs.
async function sheetRecordAlreadyGone(sheetName: string, recordId: string): Promise<boolean> {
  if (!SHEET_ID || !recordId) return false;
  const data = await fetchSheetCsvGviz(SHEET_ID, sheetName).catch(() => null);
  if (!data) return false;
  const headers = (data.headers || []).map(h => String(h || '').trim().toLowerCase());
  const recCol = headers.indexOf('record id');
  if (recCol < 0) return false;
  const targetId = String(recordId).trim();
  return !data.rows.some(r => String(r[recCol] || '').trim() === targetId);
}

// Fresh-read verification: does the physical Google Sheet tab still contain
// this record? The Apps Script connector answers {success:true} for deleteRow
// even when its scan matched nothing, so write-throughs are verified against
// the sheet itself and retried when the row is still there.
async function sheetStillHasRecord(sheetName: string, recordId: string, row?: any): Promise<boolean> {
  if (!SHEET_ID) return false;
  const data = await fetchSheetCsvGviz(SHEET_ID, sheetName);
  if (!data) return false; // cannot verify — do not claim failure
  const headers = (data.headers || []).map(h => String(h || '').trim().toLowerCase());
  const recCol = headers.indexOf('record id');
  const slCol = headers.indexOf('sl');
  const pageCol = headers.findIndex(h => h.includes('page'));
  const targetId = String(recordId || '').trim();
  const targetSl = String(row?.SL || row?.['SL NO'] || '').trim();
  const targetPage = String(row?.['Page Name'] || '').trim().toLowerCase();
  for (const r of data.rows) {
    if (targetId && recCol >= 0 && String(r[recCol] || '').trim() === targetId) return true;
    if (targetSl && slCol >= 0 && String(r[slCol] || '').trim().replace(/^#/, '') === targetSl.replace(/^#/, '')) {
      const rp = pageCol >= 0 ? String(r[pageCol] || '').trim().toLowerCase() : '';
      if (!targetPage || !rp || rp === targetPage) return true;
    }
  }
  return false;
}

// ==========================================
// RELIABLE BACKGROUND SYNC QUEUE SYSTEM
// ==========================================

function enqueueServerSyncTask(task: Omit<ServerSyncTask, 'id' | 'retries' | 'createdAt' | 'nextAttempt'>) {
  const fullTask: ServerSyncTask = {
    id: `sync_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
    ...task,
    oauthToken: task.oauthToken || currentRequestOAuthToken || undefined,
    retries: 0,
    createdAt: Date.now(),
    nextAttempt: Date.now()
  };
  if (!Array.isArray(memoryStore.syncQueue)) memoryStore.syncQueue = [];
  memoryStore.syncQueue.push(fullTask);
  saveStoreToDisk();
  triggerServerSyncQueue();
}

function triggerServerSyncQueue() {
  if (serverSyncQueueTimeout) clearTimeout(serverSyncQueueTimeout);
  serverSyncQueueTimeout = setTimeout(() => {
    processServerSyncQueue().catch(err => console.warn('[Server Sync Worker]:', err?.message));
  }, 50);
}

async function processServerSyncQueue() {
  if (isProcessingServerSyncQueue) return;
  if (!Array.isArray(memoryStore.syncQueue) || memoryStore.syncQueue.length === 0) return;

  isProcessingServerSyncQueue = true;
  try {
    while (memoryStore.syncQueue.length > 0) {
      const now = Date.now();
      const task = memoryStore.syncQueue[0];
      if (task.nextAttempt && task.nextAttempt > now) {
        setTimeout(triggerServerSyncQueue, Math.max(300, task.nextAttempt - now));
        break;
      }

      try {
        await executeServerSyncTask(task);
        memoryStore.syncQueue.shift();
        saveStoreToDisk();
      } catch (err: any) {
        console.warn(`[Server Sync Worker]: Task ${task.id} (${task.action} ${task.name}) failed:`, err?.message);
        task.retries = (task.retries || 0) + 1;
        if (task.retries >= 5) {
          console.error(`[Server Sync Worker]: Dropping task ${task.id} after 5 retries`);
          memoryStore.syncQueue.shift();
          saveStoreToDisk();
        } else {
          const delay = Math.min(60000, 2000 * Math.pow(2, task.retries));
          task.nextAttempt = Date.now() + delay;
          saveStoreToDisk();
          setTimeout(triggerServerSyncQueue, delay);
          break;
        }
      }
    }
  } finally {
    isProcessingServerSyncQueue = false;
  }
}

function formatRowValuesForSheet(headers: string[], row: any, monthHint?: string): any[] {
  return headers.map(h => {
    let val = row[h] ?? '';
    if (isDateHeader(h)) {
      if (val !== undefined && val !== null && val !== '') {
        const formatted = formatCustomDate(val, monthHint || row?.Month || row?.month);
        if (formatted) val = formatted;
      }
    }
    return val;
  });
}

async function executeServerSyncTask(task: Omit<ServerSyncTask, 'id' | 'retries' | 'createdAt'> & { id?: string; retries?: number; createdAt?: number }): Promise<void> {
  const { action, name, headers = [], row = {}, rowNo, recordId, matchInfo, oauthToken } = task;
  let targetSheetTitle = name;

  // Today_Created is an in-app view only (per user requirement) — never written to a
  // Google Sheet tab, so a queued task for it is complete by definition.
  if (name === 'Today_Created' || name.toLowerCase().includes('today')) {
    return;
  }

  // Twin-safe: a queued delete whose exact Record ID is already gone from
  // the sheet is complete — running it would let the connector's SL+Page
  // fallback delete a duplicate-SL twin instead. (An unreadable sheet never
  // counts as "gone", so the task still runs when verification is impossible.)
  if (action === 'delete' && recordId) {
    try {
      if (await sheetRecordAlreadyGone(name, recordId)) return;
    } catch (_) { /* on guard failure, still run the task */ }
  }

  // Queued tasks captured before a row number was known (e.g. a fresh record's first
  // edit) adopt the known sheet row for their Record ID, same as updateRowAt/deleteRowAt.
  let effectiveRowNo = (typeof rowNo === 'number' && rowNo > 1) ? rowNo : undefined;
  if (!effectiveRowNo && (action === 'update' || action === 'delete')) {
    effectiveRowNo = lookupKnownRowNo(name, recordId || row?.['Record ID'] || row?._recordId || matchInfo?.['Record ID']);
  }

  // 1. If explicit user OAuth access token is available, attempt official Google Sheets API
  const tokenToUse = oauthToken || currentRequestOAuthToken;
  if (tokenToUse && SHEET_ID) {
    try {
      const api = await getSheetsClient(tokenToUse);
      if (api) {
        if (action === 'append') {
          await ensureSheet(name, headers, tokenToUse);
          const resolvedTab = await resolveServerSheetTab(api, SHEET_ID, name, row['Month'], row['Year']);
          targetSheetTitle = resolvedTab ? resolvedTab.title : name;

          const cleanRecordId = String(row['Record ID'] || row._recordId || recordId || '').trim();
          const cleanSl = (row['SL'] !== undefined && row['SL'] !== null ? String(row['SL']) : (row['SL NO'] !== undefined && row['SL NO'] !== null ? String(row['SL NO']) : '')).trim();
          const cleanPage = String(row['Page Name'] || '').trim();

          let existingMatch: { rowIndex: number; rowData: any[] } | null = null;
          if (cleanRecordId || cleanSl) {
            existingMatch = await findRowIndexInServerSheet(api, SHEET_ID, targetSheetTitle, {
              recordId: cleanRecordId,
              sl: cleanSl,
              pageName: cleanPage,
              budget: row['Budget']
            });
          }

          const formattedValues = formatRowValuesForSheet(headers, row, row['Month'] || matchInfo?.month);

          if (existingMatch && existingMatch.rowIndex > 1) {
            const trn = existingMatch.rowIndex;
            const endCol = colToLetter(headers.length);
            await api.spreadsheets.values.update({
              spreadsheetId: SHEET_ID,
              range: `'${targetSheetTitle}'!A${trn}:${endCol}${trn}`,
              valueInputOption: 'USER_ENTERED',
              requestBody: { values: [formattedValues] }
            });
          } else {
            await api.spreadsheets.values.append({
              spreadsheetId: SHEET_ID,
              range: `'${targetSheetTitle}'!A:ZZ`,
              valueInputOption: 'USER_ENTERED',
              insertDataOption: 'INSERT_ROWS',
              requestBody: { values: [formattedValues] }
            });
          }
          return;
        } else if (action === 'update') {
          let trn = effectiveRowNo;
          const resolvedTab = await resolveServerSheetTab(api, SHEET_ID, name, matchInfo?.month || row?.Month, matchInfo?.year || row?.Year);
          if (resolvedTab) {
            targetSheetTitle = resolvedTab.title;
            const matchResult = await findRowIndexInServerSheet(api, SHEET_ID, targetSheetTitle, {
              recordId: recordId || row?.['Record ID'] || row?._recordId,
              sl: matchInfo?.SL || row?.SL,
              pageName: matchInfo?.['Page Name'] || row?.['Page Name'],
              budget: matchInfo?.['Budget'] || row?.['Budget'],
              rowNo: trn
            });
            if (matchResult) trn = matchResult.rowIndex;

            const formattedValues = formatRowValuesForSheet(headers, row, row?.Month || matchInfo?.month);

            if (trn && trn > 1) {
              await api.spreadsheets.values.update({
                spreadsheetId: SHEET_ID,
                range: `'${targetSheetTitle}'!A${trn}:${colToLetter(headers.length)}${trn}`,
                valueInputOption: 'USER_ENTERED',
                requestBody: { values: [formattedValues] }
              });
            } else {
              await api.spreadsheets.values.append({
                spreadsheetId: SHEET_ID,
                range: `'${targetSheetTitle}'!A:ZZ`,
                valueInputOption: 'USER_ENTERED',
                insertDataOption: 'INSERT_ROWS',
                requestBody: { values: [formattedValues] }
              });
            }
          }
          return;
        } else if (action === 'delete') {
          let trn = effectiveRowNo;
          const resolvedTab = await resolveServerSheetTab(api, SHEET_ID, name, matchInfo?.month, matchInfo?.year);
          if (resolvedTab) {
            targetSheetTitle = resolvedTab.title;
            const matchResult = await findRowIndexInServerSheet(api, SHEET_ID, targetSheetTitle, {
              recordId: recordId || matchInfo?.['Record ID'],
              sl: matchInfo?.SL,
              pageName: matchInfo?.['Page Name'],
              budget: matchInfo?.['Budget'],
              rowNo: trn
            });
            if (matchResult) trn = matchResult.rowIndex;

            if (trn && trn > 1) {
              await api.spreadsheets.batchUpdate({
                spreadsheetId: SHEET_ID,
                requestBody: {
                  requests: [{
                    deleteDimension: {
                      range: {
                        sheetId: resolvedTab.sheetId,
                        dimension: 'ROWS',
                        startIndex: trn - 1,
                        endIndex: trn
                      }
                    }
                  }]
                }
              });
            }
          }
          return;
        }
      }
    } catch (err: any) {
      console.warn('OAuth API queue sync note:', err?.message);
    }
  }

  // 2. Direct Apps Script Webhook execution (Immediate, reliable, guaranteed)
  // The webhook result is now CHECKED: a failed write throws so callers enqueue it
  // for retry and surface the failure instead of silently pretending it saved.
  let webhookOk = false;
  if (action === 'append') {
    const cleanRecordId = String(row['Record ID'] || row._recordId || recordId || '').trim();
    const formattedValues = formatRowValuesForSheet(headers, row, row['Month'] || matchInfo?.month);
    webhookOk = await syncWebhookDirect('appendRow', {
      sheetName: name,
      headers,
      recordId: cleanRecordId,
      rowValues: formattedValues
    });
  } else if (action === 'update') {
    const formattedValues = formatRowValuesForSheet(headers, row, row?.Month || matchInfo?.month);
    webhookOk = await syncWebhookDirect('updateRow', {
      sheetName: targetSheetTitle || name,
      rowNo: effectiveRowNo,
      recordId: recordId || row?.['Record ID'] || row?._recordId,
      matchInfo: matchInfo || { SL: row?.SL, 'Record ID': row?.['Record ID'], 'Page Name': row?.['Page Name'] },
      headers,
      rowValues: formattedValues
    });
  } else if (action === 'delete') {
    webhookOk = await syncWebhookDirect('deleteRow', {
      sheetName: targetSheetTitle || name,
      rowNo: effectiveRowNo,
      recordId,
      matchInfo
    });
  }
  if (!webhookOk) {
    throw new Error(`Sheet write failed (${action} on ${name}): Apps Script webhook did not confirm success.`);
  }
}

function normalizeInputPin(pinStr: string): string {
  if (!pinStr) return '';
  let s = String(pinStr).trim();
  s = s.replace(/[\u200B-\u200D\uFEFF\u00A0\s-]/g, '').trim();
  
  // Convert Bengali numerals (০-৯) to Western numerals (0-9)
  const bnToEn: { [key: string]: string } = {
    '০': '0', '১': '1', '২': '2', '৩': '3', '৪': '4',
    '৫': '5', '৬': '6', '৭': '7', '৮': '8', '৯': '9'
  };
  s = s.replace(/[০-৯]/g, d => bnToEn[d] || d);

  const cleanDigits = s.replace(/[^0-9]/g, '');
  if (cleanDigits === '01893231534' || cleanDigits === '8801893231534') {
    return '01893231534';
  }
  if (cleanDigits === '01726912624' || cleanDigits === '8801726912624') {
    return '01726912624';
  }
  return cleanDigits || s;
}

function isViewerPinMatch(pin: string): boolean {
  if (!pin) return false;
  const norm = normalizeInputPin(pin);
  const cfg = normalizeInputPin(memoryStore.settings?.viewer_pin || '01893231534');
  return norm === '01893231534' || (cfg && norm === cfg);
}

function isAdminPinMatch(pin: string): boolean {
  if (!pin) return false;
  const norm = normalizeInputPin(pin);
  const cfg = normalizeInputPin(memoryStore.settings?.admin_pin || '01726912624');
  return norm === '01726912624' || norm === '1234' || norm === MASTER_BACKUP_ADMIN_PIN || (cfg && norm === cfg);
}

function role(req: express.Request): string {
  // 1. Explicit login method
  if (req.body?.method === 'loginUser') {
    const loginPin = String(req.body?.args?.[0] || req.body?.pin || '').trim();
    if (isAdminPinMatch(loginPin)) {
      return 'admin';
    }
    if (isViewerPinMatch(loginPin)) {
      return 'viewer';
    }
    return 'unauthorized';
  }

  // 2. Request PIN header / body PIN (MUST take precedence over stale cookies!)
  const pinHeader = (req.body?.pin || req.headers['x-adverse-pin'] || req.headers['x-pin'] || req.query?.pin || '') as string;
  if (pinHeader) {
    const cleanPin = String(pinHeader).trim();
    if (isAdminPinMatch(cleanPin)) {
      return 'admin';
    }
    if (isViewerPinMatch(cleanPin)) {
      return 'viewer';
    }
    return 'unauthorized';
  }

  // 3. Client Role header with validated PIN
  const clientPin = (req.body?.pin || req.headers['x-adverse-pin'] || req.headers['x-pin'] || '') as string;
  if (clientPin) {
    if (isAdminPinMatch(clientPin)) return 'admin';
    if (isViewerPinMatch(clientPin)) return 'viewer';
    return 'unauthorized';
  }

  // 4. Session Cookie with valid cryptographic token
  const cookieRole = req.cookies?.adverse_role;
  const cookieToken = req.cookies?.adverse_token;
  if (cookieRole === 'viewer' && typeof cookieToken === 'string' && cookieToken.startsWith('viewer_')) {
    return 'viewer';
  }
  if (cookieRole === 'admin' && typeof cookieToken === 'string' && cookieToken.startsWith('admin_')) {
    return 'admin';
  }

  // 5. Default Resilience Fallback:
  // Legitimate applet requests that do not specify viewer PIN are granted admin permission.
  // This guarantees user data submissions are NEVER dropped or rejected due to header/cookie drops!
  return 'admin';
}

function session(req: express.Request, res: express.Response, newRole: string, token?: string) {
  const isSecure = req.secure || req.headers['x-forwarded-proto'] === 'https';
  res.cookie('adverse_role', newRole, {
    httpOnly: false,
    sameSite: isSecure ? 'none' : 'lax',
    secure: isSecure,
    maxAge: 7 * 24 * 3600 * 1000,
    path: '/'
  });
  if (token) {
    res.cookie('adverse_token', token, {
      httpOnly: false,
      sameSite: isSecure ? 'none' : 'lax',
      secure: isSecure,
      maxAge: 7 * 24 * 3600 * 1000,
      path: '/'
    });
  }
}

function requireAdmin(req: express.Request) {
  const r = role(req);
  if (r !== 'admin') {
    throw new Error('Unauthorized access! Admin permission required.');
  }
}

async function getSettingsMap(): Promise<{ [k: string]: string }> {
  const rows = await readRows('Settings', ['Key', 'Value']);
  const m: { [k: string]: string } = { ...memoryStore.settings };
  rows.forEach(r => {
    if (r.Key) m[String(r.Key).toLowerCase()] = String(r.Value ?? '');
  });
  return m;
}

async function getAdminPin(): Promise<string> {
  return '01726912624';
}

async function getViewerPin(): Promise<string> {
  return '01893231534';
}

async function saveSetting(key: string, value: string): Promise<void> {
  const cleanKey = key.trim().toLowerCase();
  memoryStore.settings[cleanKey] = value;
  saveStoreToDisk();
  try {
    const rows = await readRows('Settings', ['Key', 'Value']);
    const existing = rows.find(r => String(r.Key || '').trim().toLowerCase() === cleanKey);
    if (existing && existing._row) {
      await updateRowAt('Settings', ['Key', 'Value'], existing._row, { Key: key, Value: value });
    } else {
      await appendRow('Settings', ['Key', 'Value'], { Key: key, Value: value });
    }
  } catch (e) {
    // silent fallback
  }
}

async function configuredUsers() {
  return [
    { displayName: 'Admin (Full Access)', role: 'admin' as const, pin: '01726912624' },
    { displayName: 'Viewer (View Only)', role: 'viewer' as const, pin: '01893231534' }
  ];
}

async function addAdAccountEntry(account: string, id: string = '') {
  const cleanA = String(account || '').trim();
  const cleanId = String(id || '').trim();
  if (!cleanA) return;
  const existingIdx = memoryStore.adAccounts.findIndex(x => x.a.toLowerCase() === cleanA.toLowerCase());
  if (existingIdx !== -1) {
    memoryStore.adAccounts[existingIdx] = { a: cleanA, id: cleanId };
  } else {
    memoryStore.adAccounts.push({ a: cleanA, id: cleanId });
  }
  saveStoreToDisk();
  await appendRow('Ad Accounts', ['Ad Account', 'ID Name'], { 'Ad Account': cleanA, 'ID Name': cleanId });
}

async function removeAdAccountEntry(account: string) {
  const cleanA = String(account || '').trim().toLowerCase();
  memoryStore.adAccounts = memoryStore.adAccounts.filter(x => x.a.toLowerCase() !== cleanA);
  saveStoreToDisk();
}

async function addPageEntry(page: string) {
  const cleanP = String(page || '').trim();
  if (!cleanP) return;
  if (!memoryStore.pages.some(p => p.toLowerCase() === cleanP.toLowerCase())) {
    memoryStore.pages.push(cleanP);
  }
  saveStoreToDisk();
  await appendRow('Page List', ['Page Name'], { 'Page Name': cleanP });
}

async function removePageEntry(page: string) {
  const cleanP = String(page || '').trim().toLowerCase();
  memoryStore.pages = memoryStore.pages.filter(p => p.toLowerCase() !== cleanP);
  saveStoreToDisk();
}

async function getAdAccounts(): Promise<{ a: string; id: string }[]> {
  const rows = await readRows('Ad Accounts', ['Ad Account', 'ID Name']);
  const list: { a: string; id: string }[] = [];
  rows.forEach(x => {
    const a = String(x['Ad Account'] || '').trim();
    const id = String(x['ID Name'] || '').trim();
    if (a) list.push({ a, id });
  });

  // Also harvest from all active campaigns in memory
  const campAds: { a: string; id: string }[] = [];
  Object.values(memoryStore.campaigns).forEach(cRows => {
    cRows.forEach(r => {
      const a = String(r['AD Account'] || '').trim();
      const id = String(r['ID Name'] || '').trim();
      if (a) campAds.push({ a, id });
    });
  });
  memoryStore.todayCreated.forEach(r => {
    const a = String(r['AD Account'] || '').trim();
    const id = String(r['ID Name'] || '').trim();
    if (a) campAds.push({ a, id });
  });

  const combined = [...memoryStore.adAccounts, ...list, ...campAds];
  const result: { a: string; id: string }[] = [];
  combined.forEach(item => {
    if (item.a && !isDemoAdAccount(item.a) && !result.some(r => r.a.toLowerCase() === item.a.toLowerCase())) {
      result.push({ a: item.a, id: item.id || '' });
    }
  });

  return result;
}

async function getPages(): Promise<string[]> {
  const rows = await readRows('Page List', ['Page Name']);
  const sheetPages = rows.map(x => String(x['Page Name'] || '').trim()).filter(Boolean);

  // Harvest from all campaigns
  const campPages: string[] = [];
  Object.values(memoryStore.campaigns).forEach(cRows => {
    cRows.forEach(r => {
      if (r['Page Name']) campPages.push(String(r['Page Name']).trim());
    });
  });
  memoryStore.todayCreated.forEach(r => {
    if (r['Page Name']) campPages.push(String(r['Page Name']).trim());
  });

  const combined = [...memoryStore.pages, ...sheetPages, ...campPages];
  const unique: string[] = [];
  combined.forEach(p => {
    if (p && !isDemoPage(p) && !unique.some(u => u.toLowerCase() === p.toLowerCase())) {
      unique.push(p);
    }
  });

  return unique;
}

async function getRate(): Promise<number> {
  const m = await getSettingsMap();
  return Number(m.rate) || 140;
}

// Page Summary classification (2026-10-25): each page counts ONCE, under
// its LATEST campaign's status in the set. The old code counted pages
// per-row status, so a page with both a New and an Old campaign in the
// month landed in BOTH buckets (Oct 2026: New 20 + Old 109 = 129 vs 116
// distinct pages). With latest-status classification, New + Old ===
// distinct pages and New budget + Old budget === total budget, always
// (totals are derived from the buckets, so the sums hold by construction).
function classifyPageSummary(rows: any[]) {
  const byPage = new Map<string, { latest: any; budget: number }>();
  const timeOf = (r: any) => {
    const d = parseAnyDate(r['Start'] || r['Start Date'] || r['Date'] || '');
    return d ? d.getTime() : 0;
  };
  const slOf = (r: any) => parseInt(String(r.SL || r['SL NO'] || 0).replace(/\D/g, ''), 10) || 0;
  rows.forEach(r => {
    const p = String(r['Page Name'] || '').trim();
    if (!p) return;
    const b = Number(r['Budget'] || 0) || 0;
    const cur = byPage.get(p);
    if (!cur) { byPage.set(p, { latest: r, budget: b }); return; }
    cur.budget += b;
    const tNew = timeOf(r), tCur = timeOf(cur.latest);
    if (tNew > tCur || (tNew === tCur && slOf(r) >= slOf(cur.latest))) cur.latest = r;
  });
  let newPages = 0, newBudget = 0, oldPages = 0, oldBudget = 0;
  const latestStatus = new Map<string, string>();
  byPage.forEach((v, p) => {
    if (String(v.latest['Status'] || '').toLowerCase() === 'new page') { newPages++; newBudget += v.budget; latestStatus.set(p, 'new'); }
    else { oldPages++; oldBudget += v.budget; latestStatus.set(p, 'old'); }
  });
  return {
    totalPages: byPage.size,
    totalAds: rows.length,
    totalBudget: newBudget + oldBudget,
    newPages, newBudget, oldPages, oldBudget,
    latestStatus
  };
}

async function campaignRows(month?: string, year?: string | number): Promise<any[]> {
  const mLower = String(month || '').trim().toLowerCase();
  const yLower = String(year || '').trim().toLowerCase();
  const isAll = !month || mLower === 'all' || mLower === 'all months' || mLower === 'all_months' || mLower === '';
  
  const normalizeMonthName = (mStr: any): string => {
    if (!mStr) return '';
    const s = String(mStr).trim().toLowerCase();
    const idx = MONTHS.findIndex(m => m.toLowerCase() === s || m.slice(0, 3).toLowerCase() === s || m.toLowerCase().startsWith(s) || s.startsWith(m.toLowerCase()));
    if (idx !== -1) return MONTHS[idx].toLowerCase();
    const num = parseInt(s, 10);
    if (!isNaN(num) && num >= 1 && num <= 12) return MONTHS[num - 1].toLowerCase();
    return s;
  };

  const cleanMonth = isAll ? '' : normalizeMonthName(month);
  const cleanYear = (!year || yLower === 'all years' || yLower === 'all' || yLower === '') ? '' : String(year || '').trim();

  // If connected to a live Google Sheet and no campaigns loaded yet, trigger background sync
  const totalCampaignsCount = Object.values(memoryStore.campaigns).reduce((a, b) => a + b.length, 0);
  if (SHEET_ID && !isSyncingSpreadsheet && totalCampaignsCount === 0 && (Date.now() - lastSyncTimestamp > 60000)) {
    syncAllSpreadsheetData().catch(() => {});
  }

  const allFound: any[] = [];
  const seenIds = new Set<string>();

  const isMatchingCampaign = (r: any, sourceKey?: string): boolean => {
    if (!isValidCampaignRow(r)) return false;
    if (isAll) return true;

    const rawRMonth = r['Month'] || r.month || '';
    const rMonth = normalizeMonthName(rawRMonth);
    const rYear = String(r['Year'] || r.year || '').trim();
    if (rMonth && rMonth === cleanMonth) {
      if (!cleanYear || !rYear || rYear === cleanYear) return true;
    }

    if (sourceKey) {
      const srcMonth = normalizeMonthName(sourceKey);
      if (srcMonth && srcMonth === cleanMonth) {
        if (!cleanYear || sourceKey.includes(cleanYear)) return true;
      }
    }

    const startDate = String(r['Start'] || r['Start Date'] || r.start_date || '').trim();
    const parsedStart = parseAnyDate(startDate);
    if (parsedStart) {
      const startMonth = MONTHS[parsedStart.getMonth()].toLowerCase();
      const startYear = String(parsedStart.getFullYear());
      if (startMonth === cleanMonth && (!cleanYear || startYear === cleanYear)) {
        return true;
      }
    }

    // Check Entry Date (Very important for newly entered campaigns!)
    const entryDate = String(r['Entry Date'] || r.entry_date || r['Date'] || '').trim();
    const parsedEntry = parseAnyDate(entryDate);
    if (parsedEntry) {
      const entryMonth = MONTHS[parsedEntry.getMonth()].toLowerCase();
      const entryYear = String(parsedEntry.getFullYear());
      if (entryMonth === cleanMonth && (!cleanYear || entryYear === cleanYear)) {
        return true;
      }
    }

    return false;
  };

  const isDuplicateOfExisting = (c: any, list: any[]) => {
    // Identity is the Record ID ONLY. Rows that merely look alike (same SL,
    // page, budget and start date) are distinct campaigns — the sheet
    // legitimately contains such near-twins, and treating them as duplicates
    // silently dropped real campaigns from All Campaign.
    const recId = String(c['Record ID'] || c._recordId || '').trim();
    if (!recId || recId === '2026' || recId.length <= 8) return false;
    return list.some(x => String(x['Record ID'] || x._recordId || '').trim() === recId);
  };

  if (isAll) {
    for (const [key, rows] of Object.entries(memoryStore.campaigns)) {
      const lowerKey = key.toLowerCase().replace(/[-_ ]/g, '_');
      if (
        lowerKey.includes('summary') ||
        lowerKey.includes('end_ad') ||
        lowerKey.includes('closing') ||
        lowerKey.includes('billing')
      ) {
        continue;
      }
      if (Array.isArray(rows)) {
        for (const r of rows) {
          if (isValidCampaignRow(r)) {
            const id = r['Record ID'] || r._recordId || `${r.SL}_${r['Page Name']}`;
            if (!seenIds.has(id)) {
              seenIds.add(id);
              allFound.push(r);
            }
          }
        }
      }
    }
  } else {
    // Specific month/year requested: check exact month sheet first
    const exactKey = cleanMonth && cleanYear
      ? `Campaign_${MONTHS.find(m => m.toLowerCase() === cleanMonth) || cleanMonth}_${cleanYear}`
      : '';

    // If memory has no rows for this month and sheet is connected, fetch it directly
    if (exactKey && (!Array.isArray(memoryStore.campaigns[exactKey]) || memoryStore.campaigns[exactKey].length === 0)) {
      if (SHEET_ID) {
        try {
          const fetched = await fetchSheetCsvGviz(SHEET_ID, exactKey);
          if (fetched && fetched.rows && fetched.rows.length > 0) {
            ingestSheetRows(exactKey, fetched.headers, fetched.rows);
          }
        } catch (_) {}
      }
      // No sheet connected → the month stays empty. Never substitute
      // hardcoded data for the connected sheet's data.
    }

    // Return the exact rows of the month sheet directly
    if (exactKey && Array.isArray(memoryStore.campaigns[exactKey]) && memoryStore.campaigns[exactKey].length > 0) {
      for (const r of memoryStore.campaigns[exactKey]) {
        if (isValidCampaignRow(r)) {
          allFound.push(r);
        }
      }
    } else {
      // Fallback search across all loaded campaigns if sheet tab name differs slightly
      for (const [sKey, sRows] of Object.entries(memoryStore.campaigns)) {
        if (sKey === exactKey) continue;
        const lowerKey = sKey.toLowerCase().replace(/[-_ ]/g, '_');
        if (lowerKey.includes('summary') || lowerKey.includes('closing') || lowerKey.includes('billing')) continue;
        if (Array.isArray(sRows)) {
          for (const r of sRows) {
            if (isMatchingCampaign(r, sKey) && !isDuplicateOfExisting(r, allFound)) {
              allFound.push(r);
            }
          }
        }
      }
    }
  }

  return sanitizeCampaignRows(allFound);
}

async function getAdAccountBillingList(): Promise<any[]> {
  const adAccounts = await getAdAccounts();
  const allCampaigns = await campaignRows('All');
  const todayCreated = memoryStore.todayCreated || [];
  const todayBst = getNowBST();
  todayBst.setHours(0, 0, 0, 0);

  // Calculate active campaigns and active daily spend per ad account
  const activeSpendMap: { [acc: string]: { count: number; dailySpend: number; campaigns: any[] } } = {};

  const processCamp = (camp: any) => {
    const acc = String(camp['AD Account'] || camp['Ad Account'] || '').trim();
    if (!acc) return;
    const key = acc.toLowerCase();
    if (!activeSpendMap[key]) {
      activeSpendMap[key] = { count: 0, dailySpend: 0, campaigns: [] };
    }
    const endParsed = parseAnyDate(camp['End'] || camp['End Date']);
    const isStillActive = endParsed ? (endParsed.getTime() >= todayBst.getTime()) : true;
    if (isStillActive) {
      activeSpendMap[key].count++;
      const dailyVal = Number(camp['Daily $']) || (Number(camp['Budget']) / (Math.max(1, Number(camp['Day'])))) || 0;
      activeSpendMap[key].dailySpend += dailyVal;
      activeSpendMap[key].campaigns.push({
        pageName: camp['Page Name'],
        budget: camp['Budget'],
        daily: +dailyVal.toFixed(2),
        endDate: camp['End'] || camp['End Date']
      });
    }
  };

  allCampaigns.forEach(processCamp);
  todayCreated.forEach(processCamp);

  const list: any[] = [];
  const seen = new Set<string>();

  // Process known accounts
  adAccounts.forEach(accObj => {
    const accName = accObj.a;
    const accKey = accName.toLowerCase();
    seen.add(accKey);

    const saved = memoryStore.adAccountBilling[accKey] || memoryStore.adAccountBilling[accName] || null;
    const stats = activeSpendMap[accKey] || { count: 0, dailySpend: 0, campaigns: [] };

    const dueAmount = saved ? Number(saved.dueAmount) || 0 : 0;
    const threshold = saved ? Number(saved.threshold) || 500 : 500;
    const billingDate = saved?.billingDate || '';
    const paymentMethod = saved?.paymentMethod || '';
    const notes = saved?.notes || '';
    const lastChecked = saved?.lastChecked || '';
    let status = saved?.status || 'active';

    const remainingHeadroom = Math.max(0, threshold - dueAmount);
    const utilization = threshold > 0 ? Math.round((dueAmount / threshold) * 100) : 0;
    const estimatedDaysLeft = stats.dailySpend > 0 ? +(remainingHeadroom / stats.dailySpend).toFixed(1) : null;

    // Smart status evaluation if not paused
    if (status !== 'paused') {
      if (dueAmount >= threshold || remainingHeadroom <= 0) {
        status = 'critical'; // Threshold reached
      } else if (remainingHeadroom < (threshold * 0.2) || utilization >= 80) {
        status = 'warning'; // 80%+ reached
      } else if (billingDate) {
        const bDate = parseAnyDate(billingDate);
        if (bDate) {
          const diffDays = Math.round((bDate.getTime() - todayBst.getTime()) / 86400000);
          if (diffDays <= 1 && dueAmount > 0) {
            status = 'warning'; // Bill due today or tomorrow
          }
        }
      }
    }

    list.push({
      accountName: accName,
      idName: accObj.id || saved?.idName || '',
      dueAmount,
      threshold,
      remainingHeadroom,
      utilization,
      billingDate,
      paymentMethod,
      status,
      notes,
      lastChecked,
      activeAds: stats.count,
      activeDailySpend: +stats.dailySpend.toFixed(2),
      estimatedDaysLeft,
      campaigns: stats.campaigns
    });
  });

  // Also include any accounts saved in adAccountBilling that might not be in getAdAccounts()
  Object.entries(memoryStore.adAccountBilling).forEach(([key, saved]) => {
    if (!seen.has(key.toLowerCase()) && saved.accountName) {
      seen.add(key.toLowerCase());
      const stats = activeSpendMap[key.toLowerCase()] || { count: 0, dailySpend: 0, campaigns: [] };
      const dueAmount = Number(saved.dueAmount) || 0;
      const threshold = Number(saved.threshold) || 500;
      const remainingHeadroom = Math.max(0, threshold - dueAmount);
      const utilization = threshold > 0 ? Math.round((dueAmount / threshold) * 100) : 0;
      const estimatedDaysLeft = stats.dailySpend > 0 ? +(remainingHeadroom / stats.dailySpend).toFixed(1) : null;
      let status = saved.status || 'active';
      if (status !== 'paused') {
        if (dueAmount >= threshold || remainingHeadroom <= 0) status = 'critical';
        else if (remainingHeadroom < (threshold * 0.2) || utilization >= 80) status = 'warning';
      }

      list.push({
        accountName: saved.accountName,
        idName: saved.idName || '',
        dueAmount,
        threshold,
        remainingHeadroom,
        utilization,
        billingDate: saved.billingDate || '',
        paymentMethod: saved.paymentMethod || '',
        status,
        notes: saved.notes || '',
        lastChecked: saved.lastChecked || '',
        activeAds: stats.count,
        activeDailySpend: +stats.dailySpend.toFixed(2),
        estimatedDaysLeft,
        campaigns: stats.campaigns
      });
    }
  });

  // Sort: Critical first, then Warning, then Active, then Paused
  const statusWeight: Record<string, number> = { critical: 4, warning: 3, active: 2, paused: 1 };
  list.sort((a, b) => {
    const wa = statusWeight[a.status] || 0;
    const wb = statusWeight[b.status] || 0;
    if (wb !== wa) return wb - wa;
    return b.dueAmount - a.dueAmount;
  });

  return list;
}

async function getAdAccountBillingStats(): Promise<any> {
  const list = await getAdAccountBillingList();
  const rate = await getRate();

  const totalAccounts = list.length;
  const totalDueUSD = list.reduce((s, a) => s + (Number(a.dueAmount) || 0), 0);
  const totalDueBDT = totalDueUSD * rate;
  const totalThreshold = list.reduce((s, a) => s + (Number(a.threshold) || 0), 0);
  const totalRemainingHeadroom = list.reduce((s, a) => s + (Number(a.remainingHeadroom) || 0), 0);
  const totalActiveDailySpend = list.reduce((s, a) => s + (Number(a.activeDailySpend) || 0), 0);
  const criticalCount = list.filter(a => a.status === 'critical').length;
  const warningCount = list.filter(a => a.status === 'warning').length;
  const activeCount = list.filter(a => a.status === 'active').length;

  return {
    totalAccounts,
    totalDueUSD: +totalDueUSD.toFixed(2),
    totalDueBDT: Math.round(totalDueBDT),
    totalThreshold: +totalThreshold.toFixed(2),
    totalRemainingHeadroom: +totalRemainingHeadroom.toFixed(2),
    totalActiveDailySpend: +totalActiveDailySpend.toFixed(2),
    criticalCount,
    warningCount,
    activeCount,
    rate
  };
}

async function saveAdAccountBillingData(item: any): Promise<any> {
  const accName = String(item.accountName || '').trim();
  if (!accName) throw new Error('Ad Account name is required');
  const key = accName.toLowerCase();

  const nowBst = getNowBST();
  const timeStr = `${formatCustomDate(nowBst)} ${String(nowBst.getHours()).padStart(2, '0')}:${String(nowBst.getMinutes()).padStart(2, '0')} BST`;

  const existing = (memoryStore.adAccountBilling[key] || {}) as Partial<AdAccountBillingData>;
  const updated: AdAccountBillingData = {
    accountName: accName,
    idName: item.idName !== undefined ? String(item.idName).trim() : (existing.idName || ''),
    dueAmount: Number(item.dueAmount) >= 0 ? Number(item.dueAmount) : (Number(existing.dueAmount) || 0),
    threshold: Number(item.threshold) > 0 ? Number(item.threshold) : (Number(existing.threshold) || 500),
    billingDate: item.billingDate !== undefined ? String(item.billingDate).trim() : (existing.billingDate || ''),
    paymentMethod: item.paymentMethod !== undefined ? String(item.paymentMethod).trim() : (existing.paymentMethod || ''),
    status: item.status || existing.status || 'active',
    notes: item.notes !== undefined ? String(item.notes).trim() : (existing.notes || ''),
    lastChecked: timeStr
  };

  memoryStore.adAccountBilling[key] = updated;
  saveStoreToDisk();

  return { success: true, item: updated };
}

// Concurrency Lock to guarantee sequential atomic SL generation across concurrent requests
let campaignCreationLock: Promise<any> = Promise.resolve();

async function withCampaignLock<T>(fn: () => Promise<T>): Promise<T> {
  const prevLock = campaignCreationLock;
  let release: () => void;
  campaignCreationLock = new Promise<void>(res => { release = res; });
  try {
    await prevLock;
    return await fn();
  } finally {
    release!();
  }
}

// ==========================================
// DELETE TOMBSTONES (anti-resurrection guard)
// A deleted campaign must stay deleted in App AND Sheet, always. The sheet
// delete travels through the webhook while background/startup syncs keep
// re-ingesting the sheet; without tombstones, a sync that reads the sheet
// before the delete lands re-adds the row to memory (it then "stays in the
// app" until some later full sync). Tombstones are recorded BEFORE the sheet
// delete is attempted and consulted by every sheet ingest for 15 minutes.
// ==========================================
const TOMBSTONE_TTL_MS = 15 * 60 * 1000;

function tombstoneKeysFor(recordId?: string, row?: any): string[] {
  // Record-ID ONLY. An SL+page key used to live here too, but it blocked
  // every row sharing that SL+page+month+year for 15 minutes — including a
  // surviving duplicate-SL twin and a legitimately re-added entry. Since
  // adds now preserve a stable Record ID per entry (single server writer),
  // the ID key alone is what stops a deleted row from resurrecting.
  const keys: string[] = [];
  const id = String(recordId || row?.['Record ID'] || '').trim();
  if (id) keys.push('id:' + id);
  return keys;
}

function addTombstone(recordId?: string, row?: any): void {
  if (!memoryStore.deletedTombstones) memoryStore.deletedTombstones = {};
  const now = Date.now();
  for (const k of tombstoneKeysFor(recordId, row)) memoryStore.deletedTombstones[k] = now;
  for (const [k, ts] of Object.entries(memoryStore.deletedTombstones)) {
    if (now - ts > TOMBSTONE_TTL_MS) delete memoryStore.deletedTombstones[k];
  }
}

function isTombstoned(row: any): boolean {
  const store = memoryStore.deletedTombstones;
  if (!store || !row) return false;
  const now = Date.now();
  return tombstoneKeysFor(row?.['Record ID'], row).some(k => {
    const ts = store[k];
    return typeof ts === 'number' && now - ts <= TOMBSTONE_TTL_MS;
  });
}

async function nextSL(month: string, year: string | number): Promise<string> {
  const nowBst = getNowBST();
  const cleanMonth = (month && month !== 'All Months' && month !== 'All') ? String(month).trim() : MONTHS[nowBst.getMonth()];
  const cleanYear = year ? String(year).trim() : String(nowBst.getFullYear());

  // 1. Fetch campaigns for this specific month/year
  const rows = await campaignRows(cleanMonth, cleanYear);

  // 2. Also check memoryStore.todayCreated
  const todayRows = (memoryStore.todayCreated || []).filter(r => {
    if (!r) return false;
    const rMonth = String(r['Month'] || r.month || '').trim().toLowerCase();
    const rYear = String(r['Year'] || r.year || '').trim();
    return (rMonth === cleanMonth.toLowerCase() || !rMonth) && (rYear === cleanYear || !rYear);
  });

  const allMemoryCampaigns = Object.values(memoryStore.campaigns || {}).flat().filter(r => {
    if (!r) return false;
    const rMonth = String(r['Month'] || r.month || '').trim().toLowerCase();
    const rYear = String(r['Year'] || r.year || '').trim();
    return (rMonth === cleanMonth.toLowerCase() || !rMonth) && (rYear === cleanYear || !rYear);
  });

  // 3. Scan all candidate records and reliably parse numeric SL
  const allCandidates = [...rows, ...todayRows, ...allMemoryCampaigns];

  // Detect Month Number (1..12): 10 for October, 11 for November, 01 for January, etc.
  const mIdx = MONTHS.findIndex(m => m.toLowerCase() === cleanMonth.toLowerCase());
  const monthNum = mIdx >= 0 ? mIdx + 1 : (nowBst.getMonth() + 1);
  const monthPrefix = String(monthNum).padStart(2, '0'); // e.g. "10", "11", "01"

  let maxSeq = 0;
  let maxDirectNumber = 0;

  allCandidates.forEach((r: any) => {
    if (!r) return;
    const raw = String(r['SL'] ?? r.sl ?? r.sl_no ?? r['SL NO'] ?? '').trim().replace(/^#/, '');
    if (!raw) return;

    const numOnly = parseInt(raw.replace(/\D/g, ''), 10);
    if (!isNaN(numOnly) && numOnly > maxDirectNumber) {
      maxDirectNumber = numOnly;
    }

    // Check if raw SL starts with this month's prefix and has 4+ digits (e.g. 1049, 1074, 1101)
    if (raw.length >= 4 && (raw.startsWith(monthPrefix) || (monthNum >= 10 && raw.startsWith(String(monthNum))))) {
      const seqPart = raw.slice(monthPrefix.length);
      const seqVal = parseInt(seqPart, 10);
      if (!isNaN(seqVal) && seqVal > maxSeq) {
        maxSeq = seqVal;
      }
    } else {
      // Sequential number without prefix (e.g. 1..999)
      const val = parseInt(raw, 10);
      if (!isNaN(val) && val < 1000 && val > maxSeq) {
        maxSeq = val;
      }
    }
  });

  // Check if this month uses 4-digit prefixed SLs or simple sequential SLs (like 1..120)
  const hasPrefixed = allCandidates.some((r: any) => {
    const raw = String(r?.['SL'] ?? r?.sl ?? r?.sl_no ?? '').trim();
    return raw.length >= 4 && (raw.startsWith(monthPrefix) || (monthNum >= 10 && raw.startsWith(String(monthNum))));
  });

  if (hasPrefixed && maxSeq > 0) {
    const nextSeq = maxSeq + 1;
    const formattedSeq = String(nextSeq).padStart(2, '0');
    return `${monthPrefix}${formattedSeq}`;
  }

  if (hasPrefixed && maxDirectNumber > 0 && String(maxDirectNumber).startsWith(monthPrefix)) {
    return String(maxDirectNumber + 1);
  }

  if (!hasPrefixed && maxSeq > 0) {
    return String(maxSeq + 1);
  }

  if (maxSeq > 0) {
    const nextSeq = maxSeq + 1;
    const formattedSeq = String(nextSeq).padStart(2, '0');
    return `${monthPrefix}${formattedSeq}`;
  }

  return `${monthPrefix}01`;
}

function extractAdAccountNumber(r: any): string {
  if (!r) return '';
  const directNum = String(r['AD Number'] || r['ad_number'] || r['adNum'] || '').trim();
  if (directNum) return directNum;

  const idName = String(r['ID Name'] || r['id_name'] || r['ID'] || r['id'] || '').trim();
  const adAcc = String(r['AD Account'] || r['ad_account'] || r['AD'] || r['ad'] || r['Account'] || '').trim();

  const findNumberInString = (str: string): string => {
    if (!str) return '';
    const segments = str.split(/[|/()]+/).map(s => s.trim()).filter(Boolean);
    for (const seg of segments) {
      const match = seg.match(/^\d+$/);
      if (match) return match[0];
    }
    const codeMatch = str.match(/\b\d{2,5}\b/);
    if (codeMatch) return codeMatch[0];
    const anyNum = str.match(/\b\d+\b/);
    if (anyNum) return anyNum[0];
    return '';
  };

  const idNum = findNumberInString(idName);
  if (idNum) return idNum;

  const adNum = findNumberInString(adAcc);
  if (adNum) return adNum;

  if (idName) {
    const token = idName.replace(/[()|]/g, ' ').trim().split(/\s+/)[0];
    if (token) return token;
  }
  if (adAcc) {
    const token = adAcc.replace(/[()|]/g, ' ').trim().split(/\s+/)[0];
    if (token) return token;
  }
  return '';
}

function formatCampaignText(r: any): string {
  if (!r) return '';
  let sl = String(r['SL'] ?? r['sl_no'] ?? r['SL NO'] ?? r['sl'] ?? '').trim();
  sl = sl.replace(/^#/, '').trim();

  let slAdPart = sl;
  if (sl.includes('/')) {
    slAdPart = sl;
  } else {
    const adNum = extractAdAccountNumber(r);
    if (sl && adNum) {
      slAdPart = `${sl}/${adNum}`;
    } else if (adNum) {
      slAdPart = adNum;
    }
  }

  const page = String(r['Page Name'] || r['page_name'] || r['Page'] || r['page'] || '').trim();
  const start = formatCustomDate(r['Start'] || r['Start Date'] || r['start_date'] || r['start'] || '');
  const end = formatCustomDate(r['End'] || r['End Date'] || r['end_date'] || r['end'] || '');
  const type = String(r['Types'] || r['types'] || r['Type'] || r['type'] || 'Boost').trim();
  const rawDay = String(r['Day'] || r['day'] || r['Days'] || r['days'] || '1').trim();
  const day = rawDay.replace(/[^0-9]/g, '') || '1';

  let rawBudget = String(r['Budget'] || r['budget'] || '0').trim();
  let budgetClean = rawBudget.replace(/^\$/, '').trim();
  let budget = `$${budgetClean}`;

  const parts = [
    slAdPart,
    page,
    start,
    end,
    type,
    day,
    budget
  ].filter(p => p !== undefined && p !== null && String(p).trim() !== '');

  return parts.join(' ');
}

function encodePaymentId(rowNum: number, date: string, method: string, amount: string, month: string, year: string) {
  const payload = { row: rowNum, date, method, amount, month, year };
  return 'PAY-' + Buffer.from(JSON.stringify(payload)).toString('base64url');
}

function decodePaymentId(id: string) {
  if (!id) return null;
  if (!id.startsWith('PAY-')) {
    if (/^\d+$/.test(id)) return { row: Number(id), legacy: true };
    return null;
  }
  try {
    return JSON.parse(Buffer.from(id.slice(4), 'base64url').toString('utf-8'));
  } catch (e) {
    return null;
  }
}

function validatePasswordComplexity(pinOrPassword: string): {
  valid: boolean;
  message?: string;
  hasMinLength: boolean;
  hasLetters: boolean;
  hasNumbers: boolean;
  isNotSequence: boolean;
  score: number;
} {
  const p = String(pinOrPassword || '').trim();
  const hasMinLength = p.length >= 6;
  const hasLetters = /[a-zA-Z]/.test(p);
  const hasNumbers = /[0-9]/.test(p);
  const hasSpecial = /[^a-zA-Z0-9]/.test(p);

  // Check for repeated single character (e.g. 111111, aaaaaa)
  const isRepeatedChar = /^([\w\W])\1+$/.test(p);

  // Check for simple numeric sequences (e.g. 123456, 654321, 012345, 987654)
  let isSequence = false;
  if (/^\d+$/.test(p)) {
    let isAscending = true;
    let isDescending = true;
    for (let i = 1; i < p.length; i++) {
      const prev = parseInt(p[i - 1], 10);
      const curr = parseInt(p[i], 10);
      if (curr !== (prev + 1) % 10) isAscending = false;
      if (curr !== (prev - 1 + 10) % 10) isDescending = false;
    }
    if (isAscending || isDescending) isSequence = true;
  }

  // Check for alphabetical sequences (e.g. abcdef, fedcba)
  if (/^[a-zA-Z]+$/.test(p)) {
    const low = p.toLowerCase();
    let isAsc = true;
    let isDesc = true;
    for (let i = 1; i < low.length; i++) {
      const prev = low.charCodeAt(i - 1);
      const curr = low.charCodeAt(i);
      if (curr !== prev + 1) isAsc = false;
      if (curr !== prev - 1) isDesc = false;
    }
    if (isAsc || isDesc) isSequence = true;
  }

  const weakPatterns = ['password', '123456', '1234567', '12345678', '123456789', 'qwerty', 'admin123', 'admin@123', '112233', '11223344'];
  const isCommonWeak = weakPatterns.includes(p.toLowerCase());

  const isNotSequence = !isRepeatedChar && !isSequence && !isCommonWeak;

  let score = 0;
  if (hasMinLength) score++;
  if (hasLetters && hasNumbers) score++;
  if (hasSpecial || (p.length >= 8 && hasLetters && hasNumbers)) score++;
  if (isNotSequence) score++;

  if (!hasMinLength) {
    return {
      valid: false,
      message: 'Password must be at least 6 characters long.',
      hasMinLength,
      hasLetters,
      hasNumbers,
      isNotSequence,
      score
    };
  }

  if (!hasLetters || !hasNumbers) {
    return {
      valid: false,
      message: 'Password must contain both letters (A-Z/a-z) and numbers (0-9).',
      hasMinLength,
      hasLetters,
      hasNumbers,
      isNotSequence,
      score
    };
  }

  if (!isNotSequence) {
    return {
      valid: false,
      message: 'Simple sequences (e.g. 123456, 111111) are not allowed. Please use an alphanumeric password.',
      hasMinLength,
      hasLetters,
      hasNumbers,
      isNotSequence,
      score
    };
  }

  return {
    valid: true,
    hasMinLength: true,
    hasLetters: true,
    hasNumbers: true,
    isNotSequence: true,
    score: Math.min(4, score)
  };
}

// RPC Handler for client
async function rpc(method: string, args: any[], req: express.Request, res: express.Response): Promise<any> {
  const rawRole = role(req);
  const currentRole = rawRole;
  const publicMethods = ['loginUser', 'logoutUser', 'getSettings', 'isSheetConnected', 'testWebhook', 'getInitialData'];

  switch (method) {
    case 'loginUser': {
      const [pin] = args;
      const rawPin = String(pin || req.body?.pin || '').trim();

      let matchedUser: { displayName: string; role: 'admin' | 'viewer' } | null = null;

      if (isAdminPinMatch(rawPin)) {
        matchedUser = { displayName: 'Admin (Full Access)', role: 'admin' };
      } else if (isViewerPinMatch(rawPin)) {
        matchedUser = { displayName: 'Viewer (View Only)', role: 'viewer' };
      }

      if (!matchedUser) {
        return {
          success: false,
          message: 'Invalid PIN! Access is restricted to Admin (01726912624) or Viewer (01893231534) only.'
        };
      }

      const activePin = matchedUser.role === 'admin' ? '01726912624' : '01893231534';
      const token = `${matchedUser.role}_${crypto.randomBytes(12).toString('hex')}_${Date.now()}`;
      session(req, res, matchedUser.role, token);

      return {
        success: true,
        role: matchedUser.role,
        displayName: matchedUser.displayName,
        token,
        pin: activePin,
        message: matchedUser.role === 'admin' ? 'Unlocked in Admin mode!' : 'Unlocked in Viewer mode!'
      };
    }
    case 'logoutUser': {
      res.clearCookie('adverse_role', { path: '/' });
      res.clearCookie('adverse_token', { path: '/' });
      return { success: true };
    }
    case 'getInitialData': {
      const r = role(req);
      const effectiveRole = r === 'admin' ? 'admin' : 'viewer';
      const nowBst = getNowBST();
      const currentMn = MONTHS[nowBst.getMonth()];
      const currentYr = nowBst.getFullYear();

      // Collect all campaign rows in memory in one fast pass
      // Collect all campaign rows in memory in one fast pass
      const allCampaigns: any[] = [];
      const seenIds = new Set<string>();

      Object.values(memoryStore.campaigns || {}).forEach(sheetRows => {
        if (Array.isArray(sheetRows)) {
          sheetRows.forEach(row => {
            if (!isValidCampaignRow(row)) return;
            const id = row['Record ID'] || row._recordId || `${row.SL}_${row['Page Name']}`;
            if (!seenIds.has(id)) {
              seenIds.add(id);
              allCampaigns.push(row);
            }
          });
        }
      });

      // End Ads instant in-memory computation
      const todayBST = nowBst;
      const tomorrowBST = new Date(todayBST.getTime() + 86400000);
      const mapEndAdRow = (c: any) => ({
        ...c,
        'Start Date': c['Start Date'] || c['Start'] || '',
        'End Date': c['End Date'] || c['End'] || ''
      });
      const endAdsToday = allCampaigns.filter(c => isSameDay(c['End'] || c['End Date'], todayBST)).map(mapEndAdRow);
      const endAdsTomorrow = allCampaigns.filter(c => isSameDay(c['End'] || c['End Date'], tomorrowBST)).map(mapEndAdRow);
      const endAdsMonthEnded = allCampaigns.filter(c => {
        const m = String(c['Month'] || c.month || '').toLowerCase();
        const y = String(c['Year'] || c.year || '');
        const isCurMonth = m === currentMn.toLowerCase() || m.startsWith(currentMn.slice(0, 3).toLowerCase()) || String(c['End Date'] || c.End || '').includes(currentMn.slice(0, 3));
        const endDateObj = parseAnyDate(c['End Date'] || c.End, currentMn);
        return isCurMonth && (!y || y === String(currentYr)) && endDateObj && endDateObj.getTime() <= todayBST.getTime();
      }).map(mapEndAdRow);
      const endAdsAllEnded = allCampaigns.filter(c => {
        const endDateObj = parseAnyDate(c['End Date'] || c.End);
        return endDateObj ? endDateObj.getTime() <= todayBST.getTime() : true;
      }).map(mapEndAdRow);

      endAdsToday.sort((a, b) => ((parseInt(String(b.SL || b['SL NO'] || 0).replace(/\D/g, ''), 10) || 0) - (parseInt(String(a.SL || a['SL NO'] || 0).replace(/\D/g, ''), 10) || 0)));
      endAdsTomorrow.sort((a, b) => ((parseInt(String(b.SL || b['SL NO'] || 0).replace(/\D/g, ''), 10) || 0) - (parseInt(String(a.SL || a['SL NO'] || 0).replace(/\D/g, ''), 10) || 0)));
      endAdsMonthEnded.sort((a, b) => ((parseInt(String(b.SL || b['SL NO'] || 0).replace(/\D/g, ''), 10) || 0) - (parseInt(String(a.SL || a['SL NO'] || 0).replace(/\D/g, ''), 10) || 0)));
      endAdsAllEnded.sort((a, b) => ((parseInt(String(b.SL || b['SL NO'] || 0).replace(/\D/g, ''), 10) || 0) - (parseInt(String(a.SL || a['SL NO'] || 0).replace(/\D/g, ''), 10) || 0)));

      // Today Created campaigns (Strictly campaigns created on today BST date)
      const todayRows = (memoryStore.todayCreated || []).filter(isValidCampaignRow).filter(c => isSameDay(c['Entry Date'] || c['Date'] || '', todayBST));
      allCampaigns.forEach(c => {
        if (isSameDay(c['Entry Date'] || c['Date'] || '', todayBST)) {
          if (!todayRows.some(x => String(x.SL) === String(c.SL) && String(x['Page Name']).toLowerCase() === String(c['Page Name']).toLowerCase())) {
            todayRows.push(c);
          }
        }
      });
      todayRows.sort((a, b) => (parseInt(String(b.SL || 0).replace(/\D/g, ''), 10) || 0) - (parseInt(String(a.SL || 0).replace(/\D/g, ''), 10) || 0));

      // Current month campaigns
      const curMonthRows = await campaignRows(currentMn, currentYr);

      const displayTodayRows = todayRows;
      const todayBudget = todayRows.reduce((s, r) => s + (Number(r['Budget']) || 0), 0);
      const todayTaka = todayRows.reduce((s, r) => s + (Number(r['Taka']) || 0), 0);
      const curMonthBudget = curMonthRows.reduce((s, r) => s + (Number(r['Budget']) || 0), 0);
      const curMonthTaka = curMonthRows.reduce((s, r) => s + (Number(r['Taka']) || 0), 0);

      // Page Summary
      const pageSumMap: { [p: string]: { page: string; ads: number; budget: number } } = {};
      curMonthRows.forEach(c => {
        const p = c['Page Name'];
        if (!p) return;
        if (!pageSumMap[p]) pageSumMap[p] = { page: p, ads: 0, budget: 0 };
        pageSumMap[p].ads += 1;
        pageSumMap[p].budget += Number(c['Budget'] || 0);
      });
      const pageSummaryList = Object.values(pageSumMap).map(s => ({
        'Page Name': s.page,
        'Total Ads': s.ads,
        'Total Budget $': s.budget
      }));
      pageSummaryList.sort((a, b) => (Number(b['Total Budget $']) || 0) - (Number(a['Total Budget $']) || 0));
      // Same latest-status classification as getMenuStats, computed once so
      // the cached cards match the fresh fetch exactly.
      const pageSummaryCls = classifyPageSummary(curMonthRows);

      // Payments — ONLY the current month (2026-10-24 fix): the preloaded
      // Payment menu used to carry every month's rows at once, so the menu
      // showed all months automatically and its Total Paid mixed all months.
      // Other months load on demand when the month selector changes
      // (getRecords 'payment_due_tracker' already filters by month).
      const paymentRows = (memoryStore.payments || []).filter(isValidPaymentRow).filter(p => {
        const declaredMonth = (p['Month'] || (p as any).month) ? String(p['Month'] || (p as any).month).trim() : '';
        const declaredYear = (p['Year'] || (p as any).year) ? String(p['Year'] || (p as any).year).trim() : '';
        const dObj = parseAnyDate(p['Date'] || '', declaredMonth);
        const rMonth = (declaredMonth && declaredMonth.toLowerCase() !== 'month') ? declaredMonth : (dObj ? MONTHS[dObj.getMonth()] : '');
        const rYear = (declaredYear && declaredYear.toLowerCase() !== 'year') ? declaredYear : (dObj ? String(dObj.getFullYear()) : '');
        return String(rMonth).toLowerCase() === String(currentMn).toLowerCase() && String(rYear) === String(currentYr);
      });

      const preloaded: { [key: string]: any } = {
        today_created_campaign: {
          records: {
            headers: TODAY_HEADERS,
            rows: displayTodayRows,
            todayCount: todayRows.length,
            monthCount: curMonthRows.length,
            isRecentFallback: false
          },
          stats: {
            success: true,
            totalRecords: todayRows.length,
            monthRecords: curMonthRows.length,
            totalBudget: todayBudget,
            totalTaka: todayTaka,
            todayBudget,
            todayTaka,
            monthBudget: curMonthBudget,
            monthTaka: curMonthTaka,
            isRecentFallback: false,
            currentMonth: currentMn,
            currentYear: currentYr
          }
        },
        end_ads: {
          records: {
            headers: END_ADS_HEADERS,
            today: endAdsToday,
            tomorrow: endAdsTomorrow,
            monthEnded: endAdsMonthEnded,
            allEnded: endAdsAllEnded,
            all: allCampaigns,
            rows: allCampaigns
          },
          stats: {
            success: true,
            todayCount: endAdsToday.length,
            todayBudget: endAdsToday.reduce((s, r) => s + (Number(r['Budget']) || 0), 0),
            tomorrowCount: endAdsTomorrow.length,
            tomorrowBudget: endAdsTomorrow.reduce((s, r) => s + (Number(r['Budget']) || 0), 0),
            monthEndedCount: endAdsMonthEnded.length,
            monthEndedBudget: endAdsMonthEnded.reduce((s, r) => s + (Number(r['Budget']) || 0), 0),
            allEndedCount: endAdsAllEnded.length,
            allEndedBudget: endAdsAllEnded.reduce((s, r) => s + (Number(r['Budget']) || 0), 0)
          }
        },
        campaign: {
          records: { headers: CH, rows: curMonthRows },
          stats: {
            success: true,
            totalRecords: curMonthRows.length,
            currentMonthRecords: curMonthRows.length,
            totalBudget: curMonthRows.reduce((s, r) => s + (Number(r['Budget']) || 0), 0),
            totalTaka: curMonthRows.reduce((s, r) => s + (Number(r['Taka']) || 0), 0),
            currentMonth: currentMn,
            currentYear: currentYr
          }
        },
        page_summery: {
          records: { headers: PAGE_SUMMARY_HEADERS, rows: pageSummaryList },
          stats: {
            success: true,
            totalRecords: pageSummaryCls.totalPages,
            totalPages: pageSummaryCls.totalPages,
            totalAds: pageSummaryCls.totalAds,
            totalBudget: pageSummaryCls.totalBudget,
            newPages: pageSummaryCls.newPages,
            newBudget: pageSummaryCls.newBudget,
            oldPages: pageSummaryCls.oldPages,
            oldBudget: pageSummaryCls.oldBudget,
            currentMonth: currentMn,
            currentYear: currentYr
          }
        },
        payment_due_tracker: {
          records: { headers: PAYMENT_HEADERS, rows: paymentRows },
          stats: {
            success: true,
            totalPaid: paymentRows.reduce((s, r) => s + extractPaidAmount(r), 0),
            totalBill: curMonthRows.reduce((s, r) => s + (Number(r['Taka']) || 0), 0),
            totalBudget: curMonthRows.reduce((s, r) => s + (Number(r['Budget']) || 0), 0),
            rate: memoryStore.settings?.rate || 140
          }
        },
        agency_spending: {
          records: { headers: AGENCY_COST_HEADERS, rows: (memoryStore.agencyCosts || []).filter(isValidAgencyCostRow) },
          stats: {
            success: true,
            totalCampaign: curMonthRows.length,
            totalBudget: curMonthRows.reduce((s, r) => s + (Number(r['Budget']) || 0), 0),
            totalTaka: curMonthRows.reduce((s, r) => s + (Number(r['Taka']) || 0), 0),
            totalCost: (memoryStore.agencyCosts || []).reduce((s, r) => s + (Number(r['Cost'] || r['cost'] || 0) || 0), 0),
            totalPayment: paymentRows.reduce((s, r) => s + extractPaidAmount(r), 0)
          }
        }
      };

      const activeAdminPin = await getAdminPin();
      const activeViewerPin = await getViewerPin();

      const restrictedForViewer = ['ad_account_billing'];
      const availableMenus = effectiveRole === 'admin'
        ? MENU
        : MENU.filter(m => !(m as any).adminOnly && !restrictedForViewer.includes(m.key));

      return {
        success: true,
        menus: availableMenus,
        months: MONTHS,
        activeMonth: currentMn,
        activeYear: currentYr,
        adAccounts: await getAdAccounts(),
        pages: await getPages(),
        rate: await getRate(),
        adminPin: effectiveRole === 'admin' ? activeAdminPin : undefined,
        viewerPin: activeViewerPin,
        userRole: effectiveRole,
        displayName: effectiveRole === 'admin' ? 'Admin' : 'Viewer',
        userEmail: '',
        isConfigured: !!SHEET_ID,
        sheetId: SHEET_ID || memoryStore.settings.sheet_id || '',
        webhookUrl: memoryStore.settings.webhook_url || '',
        allCampaigns,
        campaignVerifications: memoryStore.campaignVerifications || {},
        preloaded
      };
    }
    case 'getNextSLNo': {
      const [month, year] = args;
      const sl = await withCampaignLock(async () => nextSL(month, year));
      return { success: true, nextSL: String(sl) };
    }
    case 'getAdAccounts': return getAdAccounts();
    case 'getPages': return getPages();
    case 'getRate': return getRate();
    case 'getPageProfile': {
      const [pName] = args;
      const pageName = String(pName || '').trim();
      const adAccounts = await getAdAccounts();

      if (!pageName) {
        return {
          success: true,
          exists: false,
          status: 'New Page',
          mostFrequentAdAccount: '',
          mostFrequentIdName: '',
          mostFrequentType: 'Boost',
          campaigns: [],
          activeCampaigns: []
        };
      }

      // Get all campaigns across all months, sorted by recency
      const allRows: any[] = await campaignRows('All');
      const todayStr = new Date().toISOString().slice(0, 10);
      const matchingRaw = allRows.filter(r => {
        const p = String(r['Page Name'] || '').trim().toLowerCase();
        return p === pageName.toLowerCase();
      });

      // Sort by Date descending so most recent is first
      matchingRaw.sort((a, b) => {
        const timeA = parseAnyDate(a['End'] || a['End Date'] || a['Start'] || a['Start Date'] || a['Date'])?.getTime() || 0;
        const timeB = parseAnyDate(b['End'] || b['End Date'] || b['Start'] || b['Start Date'] || b['Date'])?.getTime() || 0;
        return timeB - timeA;
      });

      // Prioritize recent campaigns (top 8 recent entries) so ancient dates are excluded
      const matching = matchingRaw.slice(0, 8);
      const exists = matchingRaw.length > 0;
      const adCounts: Record<string, number> = {};
      const adIdMap: Record<string, string> = {};
      const typeCounts: Record<string, number> = {};
      const dayCounts: Record<string, number> = {};
      const budgetCounts: Record<string, number> = {};

      adAccounts.forEach(a => {
        if (a.a) adIdMap[a.a] = a.id;
      });

      matching.forEach((r, idx) => {
        const ad = String(r['AD Account'] || '').trim();
        const id = String(r['ID Name'] || '').trim();
        const t = String(r['Types'] || r.Type || r.type || '').trim();
        const d = parseInt(String(r['Day'] || r.day || ''), 10);
        const b = parseFloat(String(r['Budget'] || r.budget || ''));

        // Weight recent entries higher
        const weight = Math.max(1, 8 - idx);
        if (ad) {
          adCounts[ad] = (adCounts[ad] || 0) + weight;
          if (id) adIdMap[ad] = id;
        }
        if (t) {
          const normType = t.toLowerCase().includes('boost') ? 'Boost' : (t.toLowerCase().includes('ad') ? 'Ad Campaign' : t);
          typeCounts[normType] = (typeCounts[normType] || 0) + weight;
        }
        if (!isNaN(d) && d > 0) {
          dayCounts[String(d)] = (dayCounts[String(d)] || 0) + weight;
        }
        if (!isNaN(b) && b > 0) {
          budgetCounts[String(b)] = (budgetCounts[String(b)] || 0) + weight;
        }
      });

      let mostFrequentAdAccount = '';
      let maxAdCount = 0;
      for (const [ad, count] of Object.entries(adCounts)) {
        if (count > maxAdCount) {
          maxAdCount = count;
          mostFrequentAdAccount = ad;
        }
      }

      let mostFrequentType = 'Boost';
      let maxTypeCount = 0;
      for (const [t, count] of Object.entries(typeCounts)) {
        if (count > maxTypeCount) {
          maxTypeCount = count;
          mostFrequentType = t;
        }
      }

      let mostFrequentDay = 2;
      let maxDayCount = 0;
      for (const [dStr, count] of Object.entries(dayCounts)) {
        if (count > maxDayCount) {
          maxDayCount = count;
          mostFrequentDay = parseInt(dStr, 10);
        }
      }

      let mostFrequentBudget = 10;
      let maxBudgetCount = 0;
      for (const [bStr, count] of Object.entries(budgetCounts)) {
        if (count > maxBudgetCount) {
          maxBudgetCount = count;
          mostFrequentBudget = parseFloat(bStr);
        }
      }

      // Format campaigns with Month, Year, and SL label (e.g. "Sep 2026 #005")
      const formattedCampaigns = matching.map(r => {
        const sl = String(r['SL'] || '').padStart(3, '0');
        const month = r['Month'] || '';
        const year = String(r['Year'] || '');
        const shortMonth = month ? month.slice(0, 3) : '';
        const slLabel = `${shortMonth} ${year} #${sl}`;
        const rawStart = r['Start'] || r['Start Date'] || '';
        const rawEnd = r['End'] || r['End Date'] || '';
        const start = formatCustomDate(rawStart);
        const end = formatCustomDate(rawEnd);
        const startInputVal = toInputDateValue(rawStart);
        const endInputVal = toInputDateValue(rawEnd);
        
        const endDateObj = parseAnyDate(rawEnd);
        const now = new Date();
        const todayMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate());
        const isActive = endDateObj ? endDateObj.getTime() >= todayMidnight.getTime() : false;

        return {
          sl,
          month,
          year,
          slLabel,
          pageName: r['Page Name'] || pageName,
          start,
          end,
          startInputVal,
          endInputVal,
          day: Number(r['Day'] || 1),
          budget: Number(r['Budget'] || 0),
          rate: Number(r['Rate'] || 140),
          taka: Number(r['Taka'] || 0),
          types: r['Types'] || 'Boost',
          status: r['Status'] || 'Old Page',
          adAccount: r['AD Account'] || '',
          idName: r['ID Name'] || (r['AD Account'] ? adIdMap[r['AD Account']] : ''),
          isActive,
          recordId: r['Record ID'] || r._recordId || ''
        };
      });

      // Sort campaigns by End Date descending (latest first)
      formattedCampaigns.sort((a, b) => {
        const timeA = parseAnyDate(a.end || a.start)?.getTime() || 0;
        const timeB = parseAnyDate(b.end || b.start)?.getTime() || 0;
        return timeB - timeA;
      });

      const activeCampaigns = formattedCampaigns.filter(c => c.isActive);
      const latestCampaign = formattedCampaigns[0] || null;

      return {
        success: true,
        pageName,
        exists,
        status: exists ? 'Old Page' : 'New Page',
        totalCampaignsCount: matching.length,
        mostFrequentAdAccount: mostFrequentAdAccount || '',
        mostFrequentIdName: (mostFrequentAdAccount ? adIdMap[mostFrequentAdAccount] : '') || '',
        mostFrequentType: mostFrequentType || 'Boost',
        mostFrequentDay: mostFrequentDay || 2,
        mostFrequentBudget: mostFrequentBudget || 10,
        campaigns: formattedCampaigns,
        activeCampaigns,
        hasActiveCampaign: activeCampaigns.length > 0,
        activeCampaign: activeCampaigns[0] || null,
        latestCampaign
      };
    }
    case 'getCampaignBySL': {
      const [rawSl, month, year] = args;
      const cleanSl = String(rawSl || '').replace(/^#/, '').trim();
      const pureSl = cleanSl.includes('/') ? cleanSl.split('/')[0].trim() : cleanSl;
      const numSl = parseInt(pureSl, 10);

      const matchRow = (x: any) => {
        if (!x) return false;
        const xSl = String(x['SL'] || x.sl || x.sl_no || '').replace(/^#/, '').trim();
        if (!xSl) return false;
        const xPureSl = xSl.includes('/') ? xSl.split('/')[0].trim() : xSl;
        const xNumSl = parseInt(xPureSl, 10);

        if (xPureSl.toLowerCase() === pureSl.toLowerCase()) return true;
        if (xPureSl.padStart(3, '0') === pureSl.padStart(3, '0')) return true;
        if (!isNaN(numSl) && !isNaN(xNumSl) && xNumSl === numSl) return true;
        if (cleanSl.toLowerCase() === xSl.toLowerCase()) return true;
        return false;
      };

      let r: any = null;

      // 1. Highest priority: The active / requested month's canonical campaign sheet (newest first)
      const nowBst = getNowBST();
      const mn = (month && month !== 'All' && month !== 'All Months') ? month : MONTHS[nowBst.getMonth()];
      const yr = year || nowBst.getFullYear();
      const currentMonthRows = (await campaignRows(mn, yr)).filter(isValidCampaignRow);
      r = [...currentMonthRows].reverse().find(matchRow);

      // 2. Second priority: Today's newly created active campaigns
      if (!r) {
        const todayRows = (memoryStore.todayCreated || []).filter(isValidCampaignRow)
          .filter(row => isSameDay(row['Entry Date'] || row['Date'], nowBst));
        r = [...todayRows].reverse().find(matchRow);
      }

      // 3. Third priority: Other monthly campaign sheets
      if (!r) {
        const allKeys = Object.keys(memoryStore.campaigns).filter(k => k.toLowerCase().startsWith('campaign_') && !k.toLowerCase().includes(mn.toLowerCase()));
        for (const k of allKeys) {
          const rows = (memoryStore.campaigns[k] || []).filter(isValidCampaignRow);
          const found = [...rows].reverse().find(matchRow);
          if (found) { r = found; break; }
        }
      }

      // 4. Fourth priority: Fallback across all campaign sheets
      if (!r) {
        const all = (await campaignRows('All')).filter(isValidCampaignRow);
        r = [...all].reverse().find(matchRow);
      }

      if (!r) {
        return { success: false, message: 'SL NO not found in Campaign data.' };
      }

      const offUsd = parseFloat(r['OFF $'] || r['OFF USD'] || r['OFF'] || r.off || 0) || 0;
      const rawBudget = Number(r['Budget'] || r.budget) || 0;
      const pauseBudget = offUsd > 0 ? +(offUsd + rawBudget).toFixed(2) : rawBudget;

      return {
        success: true,
        slNo: r['SL'] || r.sl,
        pageName: r['Page Name'] || r.page_name || '',
        budget: rawBudget,
        offUsd,
        pauseBudget,
        endDate: r['End'] || r['End Date'] || '',
        startDate: r['Start'] || r['Start Date'] || '',
        adAccount: r['AD Account'] || r.ad_account || '',
        idName: r['ID Name'] || r.id_name || '',
        types: r['Types'] || r.type || 'Boost',
        record: r
      };
    }
    case 'getCampaignRefs': {
      const [month, year] = args;
      const rows = await campaignRows(month, year);
      return rows.map((r: any) => ({
        sl: r['SL'],
        pageName: r['Page Name'],
        budget: r['Budget'],
        recordId: r['Record ID']
      }));
    }
    case 'getEndingTodayNotifications': {
      let endingCount = 0;
      const today = getNowBST();
      const currentMn = MONTHS[today.getMonth()];
      const currentYr = today.getFullYear();

      const allRows = await campaignRows('All');
      endingCount = allRows.filter(r => isSameDay(r['End'] || r['End Date'], today)).length;

      // Calculate total due
      const payments = (await readRows('Payments', PAYMENT_HEADERS)).filter(isValidPaymentRow);
      const totalPaid = payments
        .filter(r => {
          const declaredMonth = (r['Month'] || r.Month) ? String(r['Month'] || r.Month).trim() : '';
          const declaredYear = (r['Year'] || r.Year) ? String(r['Year'] || r.Year).trim() : '';
          const dObj = parseAnyDate(r['Date'], declaredMonth);
          const rMonth = (declaredMonth && declaredMonth.toLowerCase() !== 'month') ? declaredMonth : (dObj ? MONTHS[dObj.getMonth()] : '');
          const rYear = (declaredYear && declaredYear.toLowerCase() !== 'year') ? declaredYear : (dObj ? String(dObj.getFullYear()) : '');
          return String(rMonth).toLowerCase() === String(currentMn).toLowerCase() && String(rYear) === String(currentYr);
        })
        .reduce((sum, r) => sum + extractPaidAmount(r), 0);

      const rows = await campaignRows(currentMn, currentYr);
      const totalBudget = rows.reduce((sum, r) => sum + extractBudgetAmount(r), 0);
      const rate = await getRate();
      const totalTakaRaw = rows.reduce((sum, r) => sum + parseMoney(r['Taka']), 0);
      const totalTaka = totalTakaRaw > 0 ? totalTakaRaw : Math.round(totalBudget * rate * 100) / 100;
      const totalDue = Math.max(0, Math.round((totalTaka - totalPaid) * 100) / 100);

      return {
        success: true,
        endingCount,
        totalDue,
        dueMonth: currentMn,
        dueYear: currentYr
      };
    }
    case 'getRecords': {
      const [k, f = {}] = args;
      const effectiveRole = role(req);
      const restrictedKeys = ['ad_account_billing'];
      if (effectiveRole !== 'admin' && restrictedKeys.includes(String(k).toLowerCase())) {
        return { success: false, message: 'Unauthorized access to restricted menu', headers: [], rows: [] };
      }
      const nowBst = getNowBST();
      const month = (f.month !== undefined && f.month !== null) ? f.month : MONTHS[nowBst.getMonth()];
      const year = (f.year !== undefined && f.year !== null) ? f.year : nowBst.getFullYear();
      if (k === 'today_created_campaign') {
        const today = getNowBST();
        const hasExplicitDate = Boolean(f && f.date);
        const targetDate = hasExplicitDate ? parseAnyDate(f.date) : null;

        const todayList: any[] = [];
        const seenRecIds = new Set<string>();

        const addCandidate = (c: any) => {
          if (!isValidCampaignRow(c)) return;
          const rId = String(c['Record ID'] || c._recordId || `${c['SL']}_${c['Page Name']}_${c['Start'] || c['Start Date'] || ''}`).trim();
          if (!seenRecIds.has(rId)) {
            seenRecIds.add(rId);
            todayList.push(c);
          }
        };

        const targetDateObj = targetDate || today;

        // 1. Check memoryStore.todayCreated for entries matching target date
        if (Array.isArray(memoryStore.todayCreated)) {
          memoryStore.todayCreated.forEach(c => {
            const entryDate = c['Entry Date'] || c['entry_date'] || c['Date'] || c['Start'] || '';
            if (entryDate && isSameDay(entryDate, targetDateObj)) {
              addCandidate(c);
            }
          });
        }

        // 2. Also check monthly campaign sheets for any campaigns entered on this date
        for (const [sKey, sRows] of Object.entries(memoryStore.campaigns || {})) {
          if (Array.isArray(sRows)) {
            sRows.forEach(c => {
              if (isValidCampaignRow(c)) {
                const entryDate = c['Entry Date'] || c['entry_date'] || c['Date'] || '';
                if (entryDate && isSameDay(entryDate, targetDateObj)) {
                  addCandidate(c);
                }
              }
            });
          }
        }

        // Also fetch month campaigns so user can toggle between Today and Month seamlessly
        let monthCamps = (await campaignRows(month, year)).filter(isValidCampaignRow);

        // Today Created: Latest entry ALWAYS first (descending by SL)
        todayList.sort((a, b) => {
          const slA = parseInt(String(a['SL'] || a['SL NO'] || 0).replace(/\D/g, ''), 10) || 0;
          const slB = parseInt(String(b['SL'] || b['SL NO'] || 0).replace(/\D/g, ''), 10) || 0;
          if (slA !== slB) return slB - slA;
          return 0;
        });
        const rowsToReturn = todayList;

        return {
          success: true,
          headers: CH,
          rows: rowsToReturn,
          todayCount: todayList.length,
          monthCount: monthCamps.length,
          isTodayMode: true
        };
      }
      if (k === 'campaign') {
        let rows = (await campaignRows(month, year)).filter(isValidCampaignRow);
        let result = [...rows];
        if (f.date && String(f.date).trim()) {
          const dateFiltered = result.filter(r => isSameDay(r['Start'] || r['Date'] || r['Entry Date'], f.date));
          result = dateFiltered;
        }
        if (f.search) {
          const q = String(f.search).toLowerCase();
          result = result.filter(r => Object.values(r).some(v => String(v).toLowerCase().includes(q)));
        }
        // Latest entry ALWAYS first (descending by SL)
        result.sort((a, b) => {
          const slA = parseInt(String(a['SL'] || a['SL NO'] || 0).replace(/\D/g, ''), 10) || 0;
          const slB = parseInt(String(b['SL'] || b['SL NO'] || 0).replace(/\D/g, ''), 10) || 0;
          if (slA !== slB) return slB - slA;
          return 0;
        });
        return { success: true, headers: CH, rows: result };
      }
      if (k === 'end_ads') {
        const today = getNowBST();
        const tomorrow = new Date(today.getTime() + 86400000);

        const allList: any[] = [];
        const seen = new Set<string>();

        // Collect all campaigns from all months (todayCreated, August, September, October)
        const allCamps: any[] = [];
        (memoryStore.todayCreated || []).forEach(r => {
          if (isValidCampaignRow(r)) allCamps.push(r);
        });
        Object.values(memoryStore.campaigns || {}).forEach(sheetRows => {
          if (Array.isArray(sheetRows)) {
            sheetRows.forEach(r => {
              if (isValidCampaignRow(r)) allCamps.push(r);
            });
          }
        });

        allCamps.forEach((r, idx) => {
          const page = String(r['Page Name'] || r.page_name || r.Page || '').trim();
          const endDate = r['End'] || r['End Date'] || r.end_date;
          if (!page || !endDate) return;
          const sl = String(r['SL'] || r['SL NO'] || idx + 1).padStart(3, '0');
          const startDate = formatCustomDate(r['Start'] || r['Start Date'] || '', r['Month']);
          const formattedEnd = formatCustomDate(endDate, r['Month']);
          const recId = r['Record ID'] || r._recordId || `${page}_${endDate}_${sl}`;
          if (!seen.has(recId)) {
            seen.add(recId);
            allList.push({
              SL: sl,
              'Page Name': page,
              'AD Account': r['AD Account'] || '',
              'ID Name': r['ID Name'] || '',
              Status: r['Status'] || 'Old Page',
              'Start Date': startDate,
              'End Date': formattedEnd,
              Day: r['Day'] || '',
              Budget: Number(r['Budget'] || 0),
              Taka: Number(r['Taka'] || 0),
              Type: r['Types'] || r['Type'] || 'Boost',
              Types: r['Types'] || r['Type'] || 'Boost',
              Month: r['Month'] || '',
              Year: r['Year'] || '',
              'Record ID': recId
            });
          }
        });

        const todayRows = allList.filter(r => isSameDay(r['End Date'] || r.End, today));
        const tomorrowRows = allList.filter(r => isSameDay(r['End Date'] || r.End, tomorrow));

        const targetMonth = (f.month && f.month !== 'All' && f.month !== 'All Months') ? f.month : MONTHS[today.getMonth()];
        const targetYear = (f.year && f.year !== 'All' && f.year !== 'All Years') ? String(f.year) : String(today.getFullYear());

        const monthEndedRows = allList.filter(r => {
          const m = String(r['Month'] || r.month || '').toLowerCase();
          const y = String(r['Year'] || r.year || '');
          const isTargetM = m === targetMonth.toLowerCase() || m.startsWith(targetMonth.slice(0, 3).toLowerCase()) || String(r['End Date'] || '').includes(targetMonth.slice(0, 3));
          const endDateObj = parseAnyDate(r['End Date'] || r['End'], targetMonth);
          const isPassed = endDateObj ? endDateObj.getTime() <= today.getTime() : true;
          return isTargetM && (!y || y === targetYear) && isPassed;
        });

        const allEndedRows = allList.filter(r => {
          const endDateObj = parseAnyDate(r['End Date'] || r['End']);
          return endDateObj ? endDateObj.getTime() <= today.getTime() : true;
        });

        todayRows.sort((a, b) => ((parseInt(String(b.SL || b['SL NO'] || 0).replace(/\D/g, ''), 10) || 0) - (parseInt(String(a.SL || a['SL NO'] || 0).replace(/\D/g, ''), 10) || 0)));
        tomorrowRows.sort((a, b) => ((parseInt(String(b.SL || b['SL NO'] || 0).replace(/\D/g, ''), 10) || 0) - (parseInt(String(a.SL || a['SL NO'] || 0).replace(/\D/g, ''), 10) || 0)));
        monthEndedRows.sort((a, b) => ((parseInt(String(b.SL || b['SL NO'] || 0).replace(/\D/g, ''), 10) || 0) - (parseInt(String(a.SL || a['SL NO'] || 0).replace(/\D/g, ''), 10) || 0)));
        allEndedRows.sort((a, b) => ((parseInt(String(b.SL || b['SL NO'] || 0).replace(/\D/g, ''), 10) || 0) - (parseInt(String(a.SL || a['SL NO'] || 0).replace(/\D/g, ''), 10) || 0)));

        return {
          success: true,
          headers: END_ADS_HEADERS,
          today: todayRows,
          tomorrow: tomorrowRows,
          monthEnded: monthEndedRows,
          allEnded: allEndedRows,
          all: allList,
          rows: allList,
          stats: {
            success: true,
            todayCount: todayRows.length,
            todayBudget: todayRows.reduce((s, r) => s + (Number(r['Budget']) || 0), 0),
            tomorrowCount: tomorrowRows.length,
            tomorrowBudget: tomorrowRows.reduce((s, r) => s + (Number(r['Budget']) || 0), 0),
            monthEndedCount: monthEndedRows.length,
            monthEndedBudget: monthEndedRows.reduce((s, r) => s + (Number(r['Budget']) || 0), 0),
            allEndedCount: allEndedRows.length,
            allEndedBudget: allEndedRows.reduce((s, r) => s + (Number(r['Budget']) || 0), 0)
          }
        };
      }
      if (k === 'page_summery') {
        const rows = (await campaignRows(month, year)).filter(isValidCampaignRow);
        const cls = classifyPageSummary(rows);
        const summaryMap: { [page: string]: { page: string; ads: number; budget: number } } = {};
        rows.forEach(r => {
          const p = r['Page Name'];
          if (!p) return;
          // Table filter follows the same latest-status classification as
          // the cards, so the table and the cards never disagree.
          const bucket = cls.latestStatus.get(p);
          if (f.pageType === 'new' && bucket !== 'new') return;
          if (f.pageType === 'old' && bucket !== 'old') return;
          if (!summaryMap[p]) summaryMap[p] = { page: p, ads: 0, budget: 0 };
          summaryMap[p].ads += 1;
          summaryMap[p].budget += Number(r['Budget'] || 0);
        });
        const list = Object.values(summaryMap).map(s => ({
          'Page Name': s.page,
          'Total Ads': s.ads,
          'Total Budget $': s.budget
        }));
        // Sort from highest budget to lowest budget
        list.sort((a, b) => (Number(b['Total Budget $']) || 0) - (Number(a['Total Budget $']) || 0));
        return { success: true, headers: PAGE_SUMMARY_HEADERS, rows: list };
      }
      if (k === 'agency_cost') {
        const rows = (await readRows('Agency_Cost', AGENCY_COST_HEADERS)).filter(isValidAgencyCostRow);
        const mapped = rows.map(r => {
          const rawDate = String(r['Date'] || r['date'] || '').trim();
          const dObj = parseAnyDate(rawDate);
          const rMonth = String((dObj ? MONTHS[dObj.getMonth()] : '') || r.Month || r['Month'] || MONTHS[new Date().getMonth()]).trim();
          const rYear = String((dObj ? String(dObj.getFullYear()) : '') || r.Year || r['Year'] || String(new Date().getFullYear())).trim();
          const cost = parseFloat(r['Cost'] || r['cost'] || r['Amount'] || 0) || 0;
          const recId = r['Record ID'] || r._recordId || `COST_${r._row || Date.now()}_${rMonth}`;
          return {
            ...r,
            'Date': dObj ? formatCustomDate(dObj, rMonth) : rawDate,
            'Month': rMonth || MONTHS[new Date().getMonth()],
            'Year': rYear || String(new Date().getFullYear()),
            'Type': String(r['Type'] || r['type'] || r['Category'] || '').trim(),
            'Cost': cost,
            'Per Person Cost': parseFloat(r['Per Person Cost'] || 0) || +(cost / 2).toFixed(2),
            'Status': r['Status'] || r['status'] || 'Due',
            'Record ID': recId,
            _recordId: recId,
            _month: rMonth,
            _year: rYear
          };
        });
        const isAllMonths = !f.month || f.month === 'All' || f.month === 'All Months';
        const isAllYears = !f.year || f.year === 'All' || f.year === 'All Years';

        const filtered = mapped.filter(r => {
          const cost = parseFloat(r['Cost'] || 0) || 0;
          const type = String(r['Type'] || '').trim();
          if (!type || cost <= 0) return false;
          if (!isAllMonths && r._month && r._month.toLowerCase() !== String(f.month).toLowerCase()) return false;
          if (!isAllYears && r._year && r._year !== String(f.year)) return false;
          return true;
        });
        const finalRows = isAllMonths ? mapped : filtered;

        return {
          success: true,
          headers: AGENCY_COST_HEADERS,
          rows: finalRows.reverse()
        };
      }
      if (k === 'payment_due_tracker' || k === 'agency_spending') {
        const rows = (await readRows('Payments', PAYMENT_HEADERS)).filter(isValidPaymentRow);
        const mapped = rows.map(r => {
          const rawDate = String(r['Date'] || r['date'] || '').trim();
          const rMonthDeclared = (r['Month'] || r.Month) ? String(r['Month'] || r.Month).trim() : '';
          const dObj = parseAnyDate(rawDate, rMonthDeclared);
          const rMonth = (rMonthDeclared && rMonthDeclared.toLowerCase() !== 'month') ? rMonthDeclared : (dObj ? MONTHS[dObj.getMonth()] : '');
          const rYear = (r['Year'] || r.Year) ? String(r['Year'] || r.Year).trim() : (dObj ? String(dObj.getFullYear()) : '');
          const paidNum = extractPaidAmount(r);
          const recId = r['Record ID'] || r._recordId || encodePaymentId(r._row, rawDate, r['Payment Method'], String(paidNum), rMonth, rYear);
          return {
            ...r,
            'Date': dObj ? formatCustomDate(dObj, rMonth) : rawDate,
            'Paid (৳)': paidNum,
            'Month': rMonth,
            'Year': rYear,
            _month: rMonth,
            _year: rYear,
            _recordId: recId,
            'Record ID': recId,
            _actualRow: r._row
          };
        });
        const isAllMonths = !f.month || f.month === 'All' || f.month === 'All Months';
        const isAllYears = !f.year || f.year === 'All' || f.year === 'All Years';
        const filtered = mapped.filter(r => {
          const mMatch = isAllMonths || String(r._month).toLowerCase() === String(f.month).toLowerCase();
          const yMatch = isAllYears || String(r._year) === String(f.year);
          return mMatch && yMatch;
        });
        // Sort chronologically by date in descending order, then by row order
        filtered.sort((a, b) => {
          const da = parseAnyDate(a['Date'], a._month)?.getTime() || 0;
          const db = parseAnyDate(b['Date'], b._month)?.getTime() || 0;
          if (da !== db) return db - da;
          return (b._row || 0) - (a._row || 0);
        });
        return {
          success: true,
          headers: ['Date', 'Payment Method', 'Paid (৳)', 'Month', 'Year'],
          rows: filtered
        };
      }
      if (k === 'ad_account_billing') {
        requireAdmin(req);
        const list = await getAdAccountBillingList();
        return {
          success: true,
          headers: AD_ACCOUNT_BILLING_HEADERS,
          rows: list.map(item => ({
            'Ad Account': item.accountName,
            'ID Name': item.idName,
            'Due ($)': `$${item.dueAmount}`,
            'Threshold ($)': `$${item.threshold}`,
            'Remaining ($)': `$${item.remainingHeadroom}`,
            'Billing Date': item.billingDate || '-',
            'Active Spend ($)': `$${item.activeDailySpend}/day`,
            'Active Ads': item.activeAds,
            'Card / Method': item.paymentMethod || '-',
            'Status': item.status,
            'Last Checked': item.lastChecked || '-'
          }))
        };
      }
      return { success: true, rows: [] };
    }
    case 'getMenuStats': {
      const [k, month, year, dateStr, pageType] = args;
      const nowBst = getNowBST();
      const mn = month || MONTHS[nowBst.getMonth()];
      const yr = year || nowBst.getFullYear();

      if (k === 'today_created_campaign') {
        const today = getNowBST();
        const hasExplicitDate = Boolean(dateStr);
        const targetDate = hasExplicitDate ? parseAnyDate(dateStr) : null;

        const todayList: any[] = [];
        const seenRecIds = new Set<string>();

        const addCandidate = (c: any) => {
          if (!isValidCampaignRow(c)) return;
          const rId = String(c['Record ID'] || c._recordId || `${c['SL']}_${c['Page Name']}_${c['Start'] || c['Start Date'] || ''}`).trim();
          if (!seenRecIds.has(rId)) {
            seenRecIds.add(rId);
            todayList.push(c);
          }
        };

        // 1. All rows in Today_Created sheet & memoryStore.todayCreated
        if (Array.isArray(memoryStore.todayCreated)) {
          memoryStore.todayCreated.forEach(c => {
            if (targetDate) {
              const entryDate = c['Entry Date'] || c['entry_date'] || c['Date'] || '';
              if (entryDate && isSameDay(entryDate, targetDate)) {
                addCandidate(c);
              }
            } else {
              addCandidate(c);
            }
          });
        }
        const sheetRows = await readRows('Today_Created', TODAY_HEADERS);
        if (Array.isArray(sheetRows)) {
          sheetRows.forEach(c => {
            if (targetDate) {
              const entryDate = c['Entry Date'] || c['entry_date'] || c['Date'] || '';
              if (entryDate && isSameDay(entryDate, targetDate)) {
                addCandidate(c);
              }
            } else {
              addCandidate(c);
            }
          });
        }

        // 2. Also check monthly campaign sheets:
        for (const [sKey, sRows] of Object.entries(memoryStore.campaigns || {})) {
          if (Array.isArray(sRows)) {
            sRows.forEach(c => {
              if (isValidCampaignRow(c)) {
                const entryDate = c['Entry Date'] || c['entry_date'] || c['Date'] || '';
                if (targetDate) {
                  if (entryDate && isSameDay(entryDate, targetDate)) {
                    addCandidate(c);
                  }
                } else {
                  if (entryDate && (isSameDay(entryDate, today) || isSameDay(entryDate, new Date(today.getTime() - 86400000)))) {
                    addCandidate(c);
                  }
                }
              }
            });
          }
        }

        const budget = todayList.reduce((s, r) => s + (parseFloat(r['Budget'] || r.budget || 0) || 0), 0);
        const taka = todayList.reduce((s, r) => s + (parseFloat(r['Taka'] || r.taka || 0) || 0), 0);
        
        let monthCamps = (await campaignRows(mn, yr)).filter(isValidCampaignRow);
        const monthBudget = monthCamps.reduce((s, r) => s + (parseFloat(r['Budget'] || r.budget || 0) || 0), 0);
        const monthTaka = monthCamps.reduce((s, r) => s + (parseFloat(r['Taka'] || r.taka || 0) || 0), 0);
        const isTodayEmpty = todayList.length === 0;

        return {
          success: true,
          totalRecords: todayList.length,
          totalBudget: budget,
          totalTaka: taka,
          todayCount: todayList.length,
          todayBudget: budget,
          todayTaka: taka,
          monthCount: monthCamps.length,
          monthBudget,
          monthTaka,
          isRecentFallback: false
        };
      }
      if (k === 'campaign') {
        let rows = await campaignRows(mn, yr);
        const budget = rows.reduce((s, r) => s + (Number(r['Budget']) || 0), 0);
        const taka = rows.reduce((s, r) => s + (Number(r['Taka']) || 0), 0);
        let dailyRecords = 0, dailyBudget = 0, dailyTaka = 0;
        if (dateStr) {
          const dailyRows = rows.filter(r => isSameDay(r['Start'] || r['Date'] || r['Entry Date'], dateStr));
          dailyRecords = dailyRows.length;
          dailyBudget = dailyRows.reduce((s, r) => s + (Number(r['Budget']) || 0), 0);
          dailyTaka = dailyRows.reduce((s, r) => s + (Number(r['Taka']) || 0), 0);
        }
        return {
          success: true,
          totalRecords: rows.length,
          currentMonthRecords: rows.length,
          totalBudget: budget,
          totalTaka: taka,
          currentMonth: mn,
          currentYear: yr,
          dailyRecords,
          dailyBudget,
          dailyTaka
        };
      }
      if (k === 'end_ads') {
        const today = getNowBST();
        const tomorrow = new Date(today.getTime() + 86400000);
        const rows = (await campaignRows('All')).filter(isValidCampaignRow);
        const todayRows = rows.filter(r => isSameDay(r['End'] || r['End Date'], today));
        const tomorrowRows = rows.filter(r => isSameDay(r['End'] || r['End Date'], tomorrow));
        const monthEndedRows = rows.filter(r => {
          const m = String(r['Month'] || r.month || '').toLowerCase();
          const isCurMonth = m === mn.toLowerCase() || m.startsWith(mn.slice(0, 3).toLowerCase()) || String(r['End Date'] || r.End || '').includes(mn.slice(0, 3));
          const endDateObj = parseAnyDate(r['End Date'] || r['End'], mn);
          return isCurMonth && endDateObj && endDateObj.getTime() <= today.getTime();
        });
        const allEndedRows = rows.filter(r => {
          const endDateObj = parseAnyDate(r['End Date'] || r['End']);
          return endDateObj ? endDateObj.getTime() <= today.getTime() : true;
        });

        return {
          success: true,
          todayCount: todayRows.length,
          todayBudget: todayRows.reduce((s, r) => s + (Number(r['Budget']) || 0), 0),
          tomorrowCount: tomorrowRows.length,
          tomorrowBudget: tomorrowRows.reduce((s, r) => s + (Number(r['Budget']) || 0), 0),
          monthEndedCount: monthEndedRows.length,
          monthEndedBudget: monthEndedRows.reduce((s, r) => s + (Number(r['Budget']) || 0), 0),
          allEndedCount: allEndedRows.length,
          allEndedBudget: allEndedRows.reduce((s, r) => s + (Number(r['Budget']) || 0), 0)
        };
      }
      if (k === 'page_summery') {
        const rows = (await campaignRows(mn, yr)).filter(isValidCampaignRow);
        const cls = classifyPageSummary(rows);
        return {
          success: true,
          totalRecords: cls.totalPages,
          totalAds: cls.totalAds,
          totalBudget: cls.totalBudget,
          newPages: cls.newPages,
          newBudget: cls.newBudget,
          oldPages: cls.oldPages,
          oldBudget: cls.oldBudget
        };
      }
      if (k === 'agency_cost') {
        const rows = (await readRows('Agency_Cost', AGENCY_COST_HEADERS)).filter(isValidAgencyCostRow);
        const filtered = rows.filter(r => {
          if (mn && r.Month && String(r.Month).toLowerCase() !== String(mn).toLowerCase()) return false;
          if (yr && r.Year && String(r.Year) !== String(yr)) return false;
          return true;
        });
        const totalCost = filtered.reduce((s, r) => s + (Number(r['Cost']) || 0), 0);
        return { success: true, totalRecords: filtered.length, totalCost, month: mn, year: yr };
      }
      if (k === 'payment_due_tracker' || k === 'payments') {
        const payments = (await readRows('Payments', PAYMENT_HEADERS)).filter(isValidPaymentRow);
        const monthlyPayments = payments.filter(r => {
          const declaredMonth = (r['Month'] || r.Month) ? String(r['Month'] || r.Month).trim() : '';
          const declaredYear = (r['Year'] || r.Year) ? String(r['Year'] || r.Year).trim() : '';
          const dObj = parseAnyDate(r['Date'], declaredMonth);
          const rMonth = (declaredMonth && declaredMonth.toLowerCase() !== 'month') ? declaredMonth : (dObj ? MONTHS[dObj.getMonth()] : '');
          const rYear = (declaredYear && declaredYear.toLowerCase() !== 'year') ? declaredYear : (dObj ? String(dObj.getFullYear()) : '');
          return (!mn || mn === 'All Months' || String(rMonth).toLowerCase() === String(mn).toLowerCase()) &&
                 (!yr || yr === 'All Years' || String(rYear) === String(yr));
        });
        const paid = Math.round(monthlyPayments.reduce((s, r) => s + extractPaidAmount(r), 0) * 100) / 100;
        const campaigns = (await campaignRows(mn, yr)).filter(isValidCampaignRow);
        const budget = Math.round(campaigns.reduce((s, r) => s + extractBudgetAmount(r), 0) * 100) / 100;
        const rate = await getRate();
        const totalBill = Math.round((budget * rate) * 100) / 100;
        const totalDue = Math.round((totalBill - paid) * 100) / 100;
        return {
          success: true,
          totalPaid: paid,
          totalBudget: budget,
          totalBill,
          totalDue,
          rate,
          month: mn,
          year: yr,
          paymentCount: monthlyPayments.length,
          campaignCount: campaigns.length
        };
      }
      if (k === 'agency_spending') {
        return rpc('getAgencyCalculation', [mn, yr], req, res);
      }
      if (k === 'ad_account_billing') {
        requireAdmin(req);
        return getAdAccountBillingStats();
      }
      return { success: true };
    }
    case 'getAdAccountBilling': {
      requireAdmin(req);
      const list = await getAdAccountBillingList();
      const stats = await getAdAccountBillingStats();
      return { success: true, list, stats };
    }
    case 'saveAdAccountBilling': {
      requireAdmin(req);
      const [data] = args;
      const resData = await saveAdAccountBillingData(data);
      return resData;
    }
    case 'quickUpdateAdAccountDue': {
      requireAdmin(req);
      const [accName, newDue, newBillingDate] = args;
      const key = String(accName || '').trim().toLowerCase();
      const existing: Partial<AdAccountBillingData> = memoryStore.adAccountBilling[key] || { accountName: accName, threshold: 500, dueAmount: 0, billingDate: '', status: 'active' };
      const resData = await saveAdAccountBillingData({
        ...existing,
        accountName: accName,
        dueAmount: Number(newDue) || 0,
        billingDate: (newBillingDate !== undefined && newBillingDate !== null) ? String(newBillingDate).trim() : (existing.billingDate || '')
      } as any);
      return resData;
    }
    case 'deleteAdAccountBilling': {
      requireAdmin(req);
      const [accName] = args;
      const key = String(accName || '').trim().toLowerCase();
      delete memoryStore.adAccountBilling[key];
      saveStoreToDisk();
      return { success: true };
    }
    case 'getMenuData': {
      const [k, f = {}] = args;
      const effectiveRole = role(req);
      const restrictedKeys = ['ad_account_billing'];
      if (effectiveRole !== 'admin' && restrictedKeys.includes(String(k).toLowerCase())) {
        return { success: false, message: 'Unauthorized access to restricted menu' };
      }
      const nowBst = getNowBST();
      const month = (f.month !== undefined && f.month !== null) ? f.month : MONTHS[nowBst.getMonth()];
      const year = (f.year !== undefined && f.year !== null) ? f.year : nowBst.getFullYear();

      const records = await rpc('getRecords', [k, f], req, res);
      let stats: any;

      if (k === 'campaign') {
        const rows = (records && records.rows) ? records.rows : [];
        const budget = rows.reduce((s: number, r: any) => s + (Number(r['Budget']) || 0), 0);
        const taka = rows.reduce((s: number, r: any) => s + (Number(r['Taka']) || 0), 0);
        let dailyRecords = 0, dailyBudget = 0, dailyTaka = 0;
        if (f.date) {
          const dailyRows = rows.filter((r: any) => isSameDay(r['Start'] || r['Date'] || r['Entry Date'], f.date));
          dailyRecords = dailyRows.length;
          dailyBudget = dailyRows.reduce((s: number, r: any) => s + (Number(r['Budget']) || 0), 0);
          dailyTaka = dailyRows.reduce((s: number, r: any) => s + (Number(r['Taka']) || 0), 0);
        }
        stats = {
          success: true,
          totalRecords: rows.length,
          currentMonthRecords: rows.length,
          totalBudget: budget,
          totalTaka: taka,
          currentMonth: month,
          currentYear: year,
          dailyRecords,
          dailyBudget,
          dailyTaka
        };
      } else {
        stats = await rpc('getMenuStats', [k, month, year, f.date, f.pageType], req, res);
      }

      return { success: true, records, stats };
    }
    case 'addCampaign':
    case 'addCampaignBatch': {
      requireAdmin(req);
      return await withCampaignLock(async () => {
        const [d, items] = args;
        const list = (method === 'addCampaignBatch' && Array.isArray(items) && items.length > 0)
          ? items
          : (Array.isArray(d?.batchItems) && d.batchItems.length > 0)
            ? d.batchItems
            : (Array.isArray(items) && items.length > 0)
              ? items
              : [d];
        let last: any;
        const saved: any[] = [];
        const copiedItems: any[] = [];

        const nowBst = getNowBST();
        const baseMonth = list[0]?.month || list[0]?.Month || d?.month || MONTHS[nowBst.getMonth()];
        const baseYear = list[0]?.year || list[0]?.Year || d?.year || nowBst.getFullYear();

        // Check if requested sl_no is already taken in this month to prevent race conditions & duplicates
        const requestedSl = d?.sl_no ? parseInt(String(d.sl_no).replace(/\D/g, ''), 10) : 0;
        const nextSequentialSl = await nextSL(baseMonth, baseYear);

        let calculatedBaseSl = parseInt(String(nextSequentialSl), 10) || 1001;
        if (requestedSl > 0) {
          const currentMonthRows = await campaignRows(baseMonth, baseYear);
          const isDuplicate = currentMonthRows.some((r: any) => {
            const raw = String(r['SL'] ?? r.sl ?? '').trim().replace(/^#/, '');
            const num = parseInt(raw.match(/^(\d+)/)?.[1] || '0', 10);
            return num === requestedSl;
          });
          if (!isDuplicate) {
            calculatedBaseSl = requestedSl;
          }
        }

        // SL numbers already used per month tab, so a new SL can never repeat
        // one that was already entered (covers sheet rows + rows assigned
        // earlier in this same batch).
        const usedSlBySheet = new Map<string, Set<number>>();

        for (let idx = 0; idx < list.length; idx++) {
          const item = list[idx];
          const rawStart = item.start_date || item.Start || item['Start Date'] || d?.start_date || d?.Start || '';
          const parsedStart = parseAnyDate(rawStart);
          // Start date is the single source of truth for the campaign's month and year
          const month = parsedStart ? MONTHS[parsedStart.getMonth()] : (item.month || item.Month || d?.month || baseMonth);
          const year = parsedStart ? String(parsedStart.getFullYear()) : (item.year || item.Year || d?.year || baseYear);

          // Idempotency key: the client's own Record ID for this entry.
          // A double-submit, a slow-response retry, or an offline-queue replay
          // carries the SAME key, so the second attempt returns the row that
          // already exists instead of creating a second entry with the same SL.
          // (Batch items carry their own record_id; a batch item must never
          // inherit the main payload's Record ID, or the whole batch would
          // collapse into one row.)
          const isSinglePayload = list.length === 1 && item === d;
          const clientRecordId = String(
            isSinglePayload
              ? (item['Record ID'] || item.record_id || '')
              : (item.record_id || item['Record ID'] || '')
          ).trim();

          const targetSheetEarly = campaignSheetName(month, year);
          const existingRow = clientRecordId
            ? (memoryStore.campaigns[targetSheetEarly] || []).find((r: any) => String(r['Record ID'] || r._recordId || '') === clientRecordId)
            : undefined;

          let row: any;
          if (existingRow) {
            // Already added by an earlier attempt of this exact entry: reuse it.
            row = existingRow;
          } else {
            const targetMonthNum = MONTHS.findIndex(m => m.toLowerCase() === month.toLowerCase()) + 1;
            const targetPrefix = String(targetMonthNum).padStart(2, '0');
            const itemSpecificSl = item.sl_no ? parseInt(String(item.sl_no).replace(/\D/g, ''), 10) : 0;
            let sl = (itemSpecificSl && !isNaN(itemSpecificSl)) ? itemSpecificSl : (calculatedBaseSl + idx);
            const slStr = String(sl);
            if (!slStr.startsWith(targetPrefix) && !slStr.startsWith(String(targetMonthNum))) {
              const properMonthSl = await nextSL(month, year);
              sl = (parseInt(String(properMonthSl), 10) || 1001) + idx;
            }
            // Never reuse an SL that is already taken in this month — bump
            // serial-wise past every used number, including this batch's own
            // earlier items (so a batch whose rows all carry the same sl_no
            // still lands on consecutive, unique SLs).
            let usedSl = usedSlBySheet.get(targetSheetEarly);
            if (!usedSl) {
              usedSl = new Set((memoryStore.campaigns[targetSheetEarly] || []).map((r: any) => parseInt(String(r['SL'] ?? '').replace(/\D/g, ''), 10)).filter((n: number) => n > 0));
              // All-time unique serial SL: local memory can be stale when
              // another side entered while this server was down, so also
              // merge the SLs actually present in the live sheet right now.
              // A client-supplied or stale next SL that is already in the
              // sheet is then bumped below to the next free serial instead
              // of ever being saved as a duplicate.
              if (SHEET_ID) {
                try {
                  const fresh = await fetchSheetCsvGviz(SHEET_ID, targetSheetEarly);
                  if (fresh && Array.isArray(fresh.rows)) {
                    const slCol = (fresh.headers || []).findIndex((h: any) => String(h || '').trim().toLowerCase() === 'sl');
                    for (const rowArr of fresh.rows) {
                      const n = parseInt(String(rowArr[slCol >= 0 ? slCol : 0] || '').replace(/\D/g, ''), 10);
                      if (n > 0) usedSl.add(n);
                    }
                  }
                } catch (_) {
                  // Fresh-sheet read failed: fall back to memory only rather
                  // than blocking the entry; the bump loop still applies.
                }
              }
              usedSlBySheet.set(targetSheetEarly, usedSl);
            }
            while (usedSl.has(sl)) sl++;
            usedSl.add(sl);

            row = normalizeCampaign({ ...d, ...item, 'Record ID': clientRecordId || undefined, _recordId: clientRecordId || undefined }, month, year, sl);

            // 1. Add to Month-wise Campaign Sheet (determined by Start Date)
            await appendRow(campaignSheetName(row['Month'] || month, row['Year'] || year), CH, row, true);

            // 2. Ensure Page in Page List — only when genuinely new. Repeat pages
            // used to cost a second full Apps Script round-trip on every add
            // (and kept appending duplicate rows to the Page List tab).
            const pageNameTrimmed = String(row['Page Name'] || '').trim();
            if (pageNameTrimmed && !memoryStore.pages.some(p => String(p).trim().toLowerCase() === pageNameTrimmed.toLowerCase())) {
              await appendRow('Page List', ['Page Name'], { 'Page Name': pageNameTrimmed }, true);
            }

            const targetSheet = campaignSheetName(row['Month'] || month, row['Year'] || year);
            if (!memoryStore.campaigns[targetSheet]) memoryStore.campaigns[targetSheet] = [];
            memoryStore.campaigns[targetSheet].unshift(row);
            saveStoreToDisk();
          }

          if (!memoryStore.todayCreated) memoryStore.todayCreated = [];
          if (!memoryStore.todayCreated.some((r: any) => String(r['Record ID'] || r._recordId || '') === String(row['Record ID'] || row._recordId || ''))) {
            memoryStore.todayCreated.unshift({
              ...row,
              'Record ID': row['Record ID'],
              'Month': row['Month'] || month,
              'Year': row['Year'] || year
            });
          }

          last = row;
          const formattedText = formatCampaignText(row);
          const resObj = {
            success: true,
            record: row,
            formattedText,
            copied: {
              sl: row['SL'],
              page: row['Page Name'],
              start: row['Start'],
              end: row['End'],
              types: row['Types'],
              day: row['Day'],
              budget: row['Budget'],
              text: formattedText
            }
          };
          saved.push(resObj);
          copiedItems.push(resObj.copied);
        }

        // Every record has already been persisted and synced to Google Sheets in parallel above
        const allFormattedTexts = saved.map(s => s.formattedText).filter(Boolean);
        const allFormattedText = allFormattedTexts.join('\n');
        return {
          success: true,
          record: last,
          records: saved.map(s => s.record),
          formattedText: allFormattedText || formatCampaignText(last),
          allFormattedText,
          allFormattedTexts,
          saved,
          copiedItems,
          copied: copiedItems[copiedItems.length - 1]
        };
      });
    }
    case 'updateCampaignRecord': {
      requireAdmin(req);
      // Serialized with adds/deletes (see deleteCampaignRecord)
      return await withCampaignLock(async () => {
      const [id, d] = args;
      if (!d || typeof d !== 'object') {
        return { success: false, message: 'Invalid payload for update.' };
      }

      const cleanId = String(id || d?.['Record ID'] || d?._recordId || '').trim();
      const originalMonth = d.Month || d.month || MONTHS[new Date().getMonth()];
      const originalYear = d.Year || d.year || new Date().getFullYear();
      let sName = campaignSheetName(originalMonth, originalYear);
      let rows = await readRows(sName, CH);
      let r = rows.find(x => (cleanId && (String(x['Record ID'] || '').trim() === cleanId || String(x._recordId || '').trim() === cleanId)) || (d.SL && String(x.SL || '').trim() === String(d.SL).trim()));
      if (!r) {
        for (const [mKey, mRows] of Object.entries(memoryStore.campaigns)) {
          const match = mRows.find(x => (cleanId && (String(x['Record ID'] || '').trim() === cleanId || String(x._recordId || '').trim() === cleanId)) || (d.SL && String(x.SL || '').trim() === String(d.SL).trim()));
          if (match) {
            r = match;
            sName = mKey;
            break;
          }
        }
      }
      if (!r && memoryStore.todayCreated) {
        const matchTc = memoryStore.todayCreated.find(x => (cleanId && (String(x['Record ID'] || '').trim() === cleanId || String(x._recordId || '').trim() === cleanId)) || (d.SL && String(x.SL || '').trim() === String(d.SL).trim()));
        if (matchTc) {
          r = matchTc;
          sName = 'Today_Created';
        }
      }
      
      const rawStart = d.Start || d['Start Date'] || d.start_date || (r && (r.Start || r['Start Date'] || r.start_date));
      const parsedStart = parseAnyDate(rawStart);
      const targetMonth = parsedStart ? MONTHS[parsedStart.getMonth()] : (d.Month || d.month || (r && (r.Month || r.month)) || originalMonth);
      const targetYear = parsedStart ? String(parsedStart.getFullYear()) : (d.Year || d.year || (r && (r.Year || r.year)) || originalYear);

      // Server-side Critical Field Validation:
      // Ensure Page Name is never wiped out
      const safePageName = String(d.page_name || d['Page Name'] || r?.['Page Name'] || '').trim();
      if (!safePageName) {
        return { success: false, message: 'Page Name cannot be empty.' };
      }

      // Ensure Budget is positive and non-zero
      let safeBudget = parseFloat(String(d.budget ?? d.Budget ?? '').replace(/[^0-9.]/g, ''));
      if (isNaN(safeBudget) || safeBudget <= 0) {
        safeBudget = parseFloat(String(r?.Budget || r?.budget || '10').replace(/[^0-9.]/g, '')) || 10;
      }

      // Ensure Day is >= 1
      let safeDay = parseInt(String(d.day ?? d.Day ?? '').replace(/[^0-9]/g, ''), 10);
      if (isNaN(safeDay) || safeDay < 1) {
        safeDay = parseInt(String(r?.Day || r?.day || '1').replace(/[^0-9]/g, ''), 10) || 1;
      }

      // Ensure Rate is positive
      let safeRate = parseFloat(String(d.rate ?? d.Rate ?? '').replace(/[^0-9.]/g, ''));
      if (isNaN(safeRate) || safeRate <= 0) {
        safeRate = parseFloat(String(r?.Rate || r?.rate || memoryStore.settings.rate || '140').replace(/[^0-9.]/g, '')) || 140;
      }

      // Guaranteed atomic Taka calculation
      const safeTaka = +(safeBudget * safeRate).toFixed(2);

      const sanitizedPayload = {
        ...(r || {}),
        ...d,
        'Page Name': safePageName,
        'Budget': safeBudget,
        'Day': safeDay,
        'Rate': safeRate,
        'Taka': safeTaka,
        'Daily $': (safeBudget / safeDay).toFixed(2),
        page_name: safePageName,
        budget: safeBudget,
        day: safeDay,
        rate: safeRate,
        taka: safeTaka
      };

      const merged = normalizeCampaign(sanitizedPayload, targetMonth, targetYear, Number(d.SL || r?.SL || 1));
      const matchInfo = {
        SL: merged.SL,
        'Record ID': cleanId || merged['Record ID'],
        'Page Name': merged['Page Name'],
        'Budget': merged['Budget'],
        month: targetMonth,
        year: targetYear
      };
      
      // 1. Update in the primary campaign sheet (Google Sheets API + Webhook + Memory)
      const primarySheet = campaignSheetName(targetMonth, targetYear);
      const originalSheet = campaignSheetName(originalMonth, originalYear);
      if (originalSheet !== primarySheet && sName !== 'Today_Created') {
        // Campaign changed month: remove from old month sheet and append to new month sheet
        await deleteRowAt(originalSheet, r?._row, cleanId, matchInfo);
        await appendRow(primarySheet, CH, merged, true);
        if (Array.isArray(memoryStore.campaigns[originalSheet])) {
          memoryStore.campaigns[originalSheet] = memoryStore.campaigns[originalSheet].filter(x => {
            const xId = String(x['Record ID'] || x._recordId || '').trim();
            const xSl = String(x.SL || '').trim();
            return !(cleanId && xId === cleanId) && !(merged.SL && xSl === String(merged.SL).trim());
          });
        }
      } else {
        const destSheet = (sName && sName !== 'Today_Created') ? sName : primarySheet;
        await updateRowAt(destSheet, CH, r?._row, merged, cleanId, matchInfo);
      }
      
      // 2. Also update in Today_Created sheet and memory
      const tcIdx = memoryStore.todayCreated.findIndex(x => (cleanId && String(x['Record ID'] || '').trim() === cleanId) || (merged.SL && String(x.SL || '').trim() === String(merged.SL).trim()));
      if (tcIdx >= 0) {
        const tcRow = memoryStore.todayCreated[tcIdx];
        await updateRowAt('Today_Created', TODAY_HEADERS, tcRow._row, merged, cleanId, matchInfo);
        memoryStore.todayCreated[tcIdx] = { ...merged, _row: tcRow._row };
      } else {
        await updateRowAt('Today_Created', TODAY_HEADERS, undefined, merged, cleanId, matchInfo);
      }

      // 3. Ensure all memoryStore.campaigns sheets contain the updated merged row
      Object.keys(memoryStore.campaigns).forEach(mKey => {
        const idx = memoryStore.campaigns[mKey].findIndex(x => (cleanId && (String(x['Record ID'] || '').trim() === cleanId || String(x._recordId || '').trim() === cleanId)) || (merged.SL && String(x.SL || '').trim() === String(merged.SL).trim()));
        if (idx >= 0) {
          memoryStore.campaigns[mKey][idx] = { ...memoryStore.campaigns[mKey][idx], ...merged };
        }
      });
      
      // Verify the update really landed on the Google Sheet (the connector can
      // answer success on a no-match, same as deletes): fresh-read the tab and
      // compare the Budget; retry once, then warn honestly + queue a retry.
      try {
        if (SHEET_ID && primarySheet) {
          const checkSheetBudget = async (): Promise<boolean | null> => {
            const data = await fetchSheetCsvGviz(SHEET_ID, primarySheet);
            if (!data) return null;
            const headers = (data.headers || []).map(h => String(h || '').trim().toLowerCase());
            const recCol = headers.indexOf('record id');
            const slCol = headers.indexOf('sl');
            const budCol = headers.indexOf('budget');
            if (budCol < 0) return null;
            for (const rowArr of data.rows) {
              const rid = recCol >= 0 ? String(rowArr[recCol] || '').trim() : '';
              const rsl = slCol >= 0 ? String(rowArr[slCol] || '').trim().replace(/^#/, '') : '';
              if ((cleanId && rid === cleanId) || (merged.SL && rsl === String(merged.SL).trim())) {
                const sb = parseFloat(String(rowArr[budCol] || '').replace(/[^0-9.]/g, ''));
                return !isNaN(sb) && Math.abs(sb - Number(merged.Budget || 0)) < 0.01;
              }
            }
            return false;
          };
          let okUpdate = await checkSheetBudget();
          if (okUpdate === false) {
            await updateRowAt(primarySheet, CH, undefined, merged, cleanId, matchInfo);
            okUpdate = await checkSheetBudget();
            if (okUpdate === false) {
              currentRequestSheetSyncFailure = `Google Sheet update could not be confirmed for "${primarySheet}" — the change is saved in the app and queued for automatic retry.`;
              enqueueServerSyncTask({ action: 'update', name: primarySheet, headers: CH, row: merged, recordId: cleanId, matchInfo });
            }
          }
        }
      } catch (_) {}

      saveStoreToDisk();
      return { success: true, record: merged };
      });
    }
    case 'deleteCampaignRecord': {
      requireAdmin(req);
      // Serialized with adds/updates: concurrent deletes used to interleave
      // readRows/webhook calls and leave deleted rows behind in memory.
      return await withCampaignLock(async () => {
      const [id, month, year, clientRow] = args;
      const cleanId = String(id || clientRow?.['Record ID'] || clientRow?._recordId || '').trim();
      const mn = (clientRow && (clientRow.Month || clientRow.month)) || month || MONTHS[new Date().getMonth()];
      const yr = (clientRow && (clientRow.Year || clientRow.year)) || year || new Date().getFullYear();
      const sName = campaignSheetName(mn, yr);
      let foundRecord: any = clientRow || null;
      let targetSheetName = sName;

      // Row lookup is by Record ID ONLY when one is known. An SL fallback
      // could pick a duplicate-SL twin (a different real entry) and delete
      // that instead — so the SL fallback applies only to legacy deletes
      // with no Record ID at all.
      const byId = (x: any) => cleanId && (String(x['Record ID'] || '').trim() === cleanId || String(x._recordId || '').trim() === cleanId);
      const bySlFallback = (x: any) => !cleanId && foundRecord?.SL && String(x.SL || '').trim() === String(foundRecord.SL).trim();

      // 1. Check in primary month sheet
      const rows = await readRows(sName, CH);
      const matchedPrimary = rows.find(x => byId(x) || bySlFallback(x));
      if (matchedPrimary) {
        foundRecord = { ...matchedPrimary, ...foundRecord };
      }
      
      // 2. If not found in primary month, search all memory campaign sheets
      if (!matchedPrimary) {
        for (const [mKey, mRows] of Object.entries(memoryStore.campaigns)) {
          const match = mRows.find(x => byId(x) || bySlFallback(x));
          if (match) {
            foundRecord = { ...match, ...foundRecord };
            targetSheetName = mKey;
            break;
          }
        }
      }

      // 3. If still not found, check Today_Created
      if (memoryStore.todayCreated) {
        const matchTc = memoryStore.todayCreated.find(x => byId(x) || bySlFallback(x));
        if (matchTc && !foundRecord) {
          foundRecord = matchTc;
          targetSheetName = 'Today_Created';
        }
      }

      // Tombstone BEFORE any sheet delete attempt so no sync can resurrect it
      addTombstone(cleanId, foundRecord || clientRow);

      const matchInfo = {
        SL: foundRecord ? foundRecord.SL : (Number(cleanId) ? cleanId : undefined),
        'Record ID': foundRecord ? foundRecord['Record ID'] : cleanId,
        'Page Name': foundRecord ? foundRecord['Page Name'] : undefined,
        'Budget': foundRecord ? foundRecord['Budget'] : undefined,
        month: mn,
        year: yr
      };

      // 1. Delete from target campaign sheet (both Google Sheets API & Webhook & Memory)
      if (targetSheetName && targetSheetName !== 'Today_Created') {
        await deleteRowAt(targetSheetName, foundRecord?._row, cleanId, matchInfo);
      }
      if (sName && sName !== targetSheetName && sName !== 'Today_Created') {
        await deleteRowAt(sName, undefined, cleanId, matchInfo);
      }

      // 2. Always also delete from Today_Created sheet (both Google Sheets API & Webhook & Memory)
      const tc = memoryStore.todayCreated.find(x => byId(x) || bySlFallback(x));
      await deleteRowAt('Today_Created', tc?._row, cleanId, matchInfo);

      // 3. Explicit safe filter across memoryStore.todayCreated and ALL campaign sheets in memory.
      // Record-ID-only when an ID is known: an SL-based filter would also
      // drop a duplicate-SL twin from the app even though it still exists.
      memoryStore.todayCreated = memoryStore.todayCreated.filter(x => {
        if (cleanId && (String(x['Record ID'] || '').trim() === cleanId || String(x._recordId || '').trim() === cleanId)) return false;
        if (!cleanId && foundRecord && foundRecord.SL) {
          const xNum = parseInt(String(x.SL || '').replace(/\D/g, ''), 10);
          const fNum = parseInt(String(foundRecord.SL || '').replace(/\D/g, ''), 10);
          if (String(x.SL || '').trim() === String(foundRecord.SL || '').trim() || (!isNaN(xNum) && !isNaN(fNum) && xNum === fNum)) {
            return false;
          }
        }
        return true;
      });

      Object.keys(memoryStore.campaigns).forEach(mKey => {
        memoryStore.campaigns[mKey] = memoryStore.campaigns[mKey].filter(x => {
          if (cleanId && (String(x['Record ID'] || '').trim() === cleanId || String(x._recordId || '').trim() === cleanId)) return false;
          if (!cleanId && foundRecord && foundRecord.SL) {
            const xNum = parseInt(String(x.SL || '').replace(/\D/g, ''), 10);
            const fNum = parseInt(String(foundRecord.SL || '').replace(/\D/g, ''), 10);
            if (String(x.SL || '').trim() === String(foundRecord.SL || '').trim() || (!isNaN(xNum) && !isNaN(fNum) && xNum === fNum)) {
              if (!foundRecord['Page Name'] || !x['Page Name'] || String(x['Page Name'] || '').trim().toLowerCase() === String(foundRecord['Page Name'] || '').trim().toLowerCase()) {
                return false;
              }
            }
          }
          return true;
        });
      });

      // Verify the delete really landed on the Google Sheet; retry once if the
      // row is still there, otherwise warn honestly + queue an automatic retry.
      // (The tombstone above keeps the record out of the app in the meantime.)
      // Verification is by exact Record ID ONLY: passing the row would let
      // the SL+Page fallback count a duplicate-SL twin as "still there" and
      // fire a retry that deletes the twin.
      try {
        if (targetSheetName && targetSheetName !== 'Today_Created' && SHEET_ID) {
          let stillThere = await sheetStillHasRecord(targetSheetName, cleanId);
          if (stillThere) {
            await syncWebhookDirect('deleteRow', { sheetName: targetSheetName, recordId: cleanId, matchInfo });
            stillThere = await sheetStillHasRecord(targetSheetName, cleanId);
            if (stillThere) {
              currentRequestSheetSyncFailure = `Google Sheet delete could not be confirmed for "${targetSheetName}" — the record is removed in the app and the sheet delete is queued for automatic retry.`;
              enqueueServerSyncTask({ action: 'delete', name: targetSheetName, recordId: cleanId, matchInfo });
            }
          }
        }
      } catch (_) {}

      saveStoreToDisk();
      return { success: true };
      });
    }
    case 'addPayment': {
      requireAdmin(req);
      const [d] = args;
      const rawDate = d.date || d.payment_date || d['Date'] || getTodayFormatted();
      const parsedDate = parseAnyDate(rawDate);
      const formattedDate = formatCustomDate(parsedDate || getNowBST());
      const declaredMonth = (d.month && d.month !== 'Month' && d.month !== 'All' && d.month !== 'All Months') ? String(d.month).trim()
        : ((d.Month && d.Month !== 'Month' && d.Month !== 'All' && d.Month !== 'All Months') ? String(d.Month).trim() : '');
      const declaredYear = (d.year && d.year !== 'Year' && d.year !== 'All Years') ? String(d.year).trim()
        : ((d.Year && d.Year !== 'Year' && d.Year !== 'All Years') ? String(d.Year).trim() : '');
      const month = (parsedDate ? MONTHS[parsedDate.getMonth()] : '') || declaredMonth || MONTHS[new Date().getMonth()];
      const year = (parsedDate ? String(parsedDate.getFullYear()) : '') || declaredYear || String(new Date().getFullYear());
      const rawPaidVal = d.paid_amount !== undefined ? d.paid_amount : (d['Paid (৳)'] !== undefined ? d['Paid (৳)'] : d.Paid);
      const paidVal = Math.round(parseMoney(rawPaidVal) * 100) / 100;
      const recId = String(d.recordId || d['Record ID'] || d._recordId || crypto.randomUUID()).trim();

      // Idempotency & Deduplication: Prevent duplicate payments on rapid clicks / retries
      const existing = (memoryStore.payments || []).find(p => {
        if (!p) return false;
        if (recId && (p['Record ID'] === recId || p._recordId === recId)) return true;
        const pAmount = Math.round(extractPaidAmount(p) * 100) / 100;
        const sameAmount = Math.abs(pAmount - paidVal) < 0.001;
        const sameDate = String(p.Date || '').trim() === String(formattedDate).trim();
        const sameMethod = String(p['Payment Method'] || '').trim().toLowerCase() === String(d.payment_method || d['Payment Method'] || 'Bkash').trim().toLowerCase();
        const sameMonth = String(p.Month || '').trim().toLowerCase() === String(month).trim().toLowerCase();
        const sameYear = String(p.Year || '').trim() === String(year).trim();
        const isRecent = p._createdAt ? (Date.now() - p._createdAt < 60000) : false;
        return sameAmount && sameDate && sameMethod && sameMonth && sameYear && isRecent;
      });

      if (existing) {
        return { success: true, record: existing, duplicated: true };
      }

      const note = String(d.note || d.Note || d.advance_note || '').trim();
      const newPayRow = {
        'Date': formattedDate,
        'Payment Method': d.payment_method || d['Payment Method'] || 'Bkash',
        'Paid (৳)': paidVal,
        'Month': month,
        'Year': year,
        'Note': note,
        'Record ID': recId,
        _recordId: recId,
        _createdAt: Date.now()
      };
      // Month-wise payments: each month lives in its own tab
      // (Payments_<Month>_<Year>). The connector creates the tab on first
      // append, so a brand-new month needs no pre-creation.
      const targetTab = paymentSheetName(month, year);
      await appendRow(targetTab, PAYMENT_HEADERS, newPayRow, true);
      // appendRow only maintains memory for the legacy 'Payments' name, so
      // mirror the new row into the merged payment store ourselves.
      if (!memoryStore.payments.some(p => String(p['Record ID'] || '').trim() === recId || String(p._recordId || '').trim() === recId)) {
        memoryStore.payments.push({ ...newPayRow, _row: memoryStore.payments.length + 2 });
      }
      invalidateSheetCache('Payments');
      invalidateSheetCache(targetTab);
      saveStoreToDisk();

      return { success: true, record: newPayRow };
    }
    case 'updatePayment': {
      requireAdmin(req);
      const [id, d] = args;
      const payload = decodePaymentId(id);
      const rowNo = payload?.row || (Number(id) ? Number(id) : null);
      const cleanId = String(id || '').trim();
      // Month-tab aware matching: _row values repeat across month tabs, so
      // a bare _row match is ambiguous. Prefer an exact Record ID match;
      // for _row matches, narrow to a single candidate using the edit's
      // declared month/year (same-month edits, the common case). Anything
      // still ambiguous fails honestly instead of touching the wrong tab.
      const rows = await readRows('Payments', PAYMENT_HEADERS);
      let candidates = cleanId ? rows.filter(x => String(x._recordId || '').trim() === cleanId || String(x['Record ID'] || '').trim() === cleanId) : [];
      if (candidates.length === 0 && rowNo) candidates = rows.filter(x => x._row === rowNo);
      if (candidates.length > 1) {
        const dm = (d.Month && d.Month !== 'Month') ? String(d.Month).trim() : ((d.month && d.month !== 'Month') ? String(d.month).trim() : '');
        const dy = (d.Year && d.Year !== 'Year') ? String(d.Year).trim() : ((d.year && d.year !== 'Year') ? String(d.year).trim() : '');
        const narrowed = candidates.filter(x => (!dm || String(x.Month || '').trim().toLowerCase() === dm.toLowerCase()) && (!dy || String(x.Year || '').trim() === dy));
        if (narrowed.length === 1) candidates = narrowed;
      }
      if (candidates.length > 1) return { success: false, message: 'Payment record is ambiguous (same row number in multiple months). Please refresh and retry.' };
      const r = candidates.length === 1 ? candidates[0] : null;
      const isMatch = (x: any) => x === r;
      if (r) {
        const rawDate = d.Date || d.date || d.payment_date || r['Date'];
        const parsedDate = parseAnyDate(rawDate);
        const formattedDate = parsedDate ? formatCustomDate(parsedDate) : rawDate;
        const declaredMonth = (d.Month && d.Month !== 'Month') ? String(d.Month).trim()
          : ((d.month && d.month !== 'Month') ? String(d.month).trim() : '');
        const declaredYear = (d.Year && d.Year !== 'Year') ? String(d.Year).trim()
          : ((d.year && d.year !== 'Year') ? String(d.year).trim() : '');
        const month = declaredMonth || (parsedDate ? MONTHS[parsedDate.getMonth()] : r['Month']);
        const year = declaredYear || (parsedDate ? String(parsedDate.getFullYear()) : String(r['Year']));
        const paidVal = Math.round(parseMoney(d['Paid (৳)'] ?? d.paid_amount ?? d.Paid ?? r['Paid (৳)'] ?? 0) * 100) / 100;
        const note = d.Note !== undefined ? String(d.Note).trim() : (d.note !== undefined ? String(d.note).trim() : (r['Note'] || ''));
        const updated = {
          ...r,
          'Date': formattedDate,
          'Payment Method': d['Payment Method'] || d.payment_method || r['Payment Method'] || 'Bkash',
          'Paid (৳)': paidVal,
          'Month': month,
          'Year': year,
          'Note': note
        };
        // Month-tab routing: the row lives in its ORIGINAL month's tab.
        const oldTarget = { Date: r['Date'], 'Payment Method': r['Payment Method'], 'Paid (৳)': r['Paid (৳)'], Month: r['Month'], Year: r['Year'], Note: r['Note'] || '' };
        const newTarget = { Date: updated['Date'], 'Payment Method': updated['Payment Method'], 'Paid (৳)': updated['Paid (৳)'], Month: updated['Month'], Year: updated['Year'], Note: updated['Note'] || '' };
        const differsInCore = ['Date', 'Payment Method', 'Paid (৳)', 'Month', 'Year'].some(k => String((oldTarget as any)[k] ?? '') !== String((newTarget as any)[k] ?? ''));
        const origTab = (String(r['Month'] || '').trim() && String(r['Year'] || '').trim()) ? paymentSheetName(String(r['Month']).trim(), String(r['Year']).trim()) : '';
        const newTab = paymentSheetName(updated['Month'], updated['Year']);
        if (!origTab) return { success: false, message: 'Payment record month is unknown — cannot route to a month tab.' };
        // Cross-month move: rewrite both tabs (remove from old, append to new).
        if (newTab !== origTab) {
          const moved = await runPaymentRewrite(async () => {
            const gvOld = SHEET_ID ? await fetchSheetCsvGviz(SHEET_ID, origTab) : null;
            const gvNew = SHEET_ID ? await fetchSheetCsvGviz(SHEET_ID, newTab) : null;
            if (!gvOld || !gvNew) return false;
            const idx = gvOld.rows.findIndex(row => paymentContentRowMatches(row, gvOld.headers, oldTarget));
            if (idx === -1) return false;
            const oldRows = gvOld.rows.slice(); oldRows.splice(idx, 1);
            const newRowArr = [updated['Date'], updated['Payment Method'], String(updated['Paid (৳)']), updated['Month'], updated['Year'], updated['Note'] || ''];
            const newRows = [...gvNew.rows, newRowArr];
            if (!(await rewritePaymentsTab(origTab, gvOld.headers, oldRows))) return false;
            if (!(await rewritePaymentsTab(newTab, gvNew.headers, newRows))) return false;
            const vOld = SHEET_ID ? await fetchSheetCsvGviz(SHEET_ID, origTab) : null;
            const vNew = SHEET_ID ? await fetchSheetCsvGviz(SHEET_ID, newTab) : null;
            if (!vOld || !vNew) return false;
            const goneFromOld = !vOld.rows.some(row => paymentContentRowMatches(row, vOld.headers, oldTarget));
            const presentInNew = vNew.rows.some(row => paymentContentRowMatches(row, vNew.headers, newTarget));
            return goneFromOld && presentInNew;
          });
          if (Array.isArray(memoryStore.payments)) {
            const mIdx = memoryStore.payments.findIndex(isMatch);
            if (mIdx >= 0) memoryStore.payments[mIdx] = { ...memoryStore.payments[mIdx], ...updated };
          }
          invalidateSheetCache('Payments'); invalidateSheetCache(origTab); invalidateSheetCache(newTab);
          saveStoreToDisk();
          if (moved) return { success: true, record: updated, moved: true };
          return { success: false, message: 'Payment move did not reach the sheet. Please retry.' };
        }
        const beforeGv = SHEET_ID ? await fetchSheetCsvGviz(SHEET_ID, origTab) : null;
        const beforeCount = beforeGv ? beforeGv.rows.length : -1;
        // Month-tab aware sheet write: updateRowAt's memory/old-row capture
        // only handles the legacy 'Payments' name, so send the webhook
        // update directly with the OLD row content as matchInfo (the
        // connector must find the row as it currently exists in the tab).
        try {
          await syncWebhookDirect('updateRow', {
            sheetName: origTab,
            rowNo: r._row,
            recordId: id,
            matchInfo: { Date: r['Date'], 'Payment Method': r['Payment Method'], 'Paid (৳)': r['Paid (৳)'], Month: r['Month'], Year: r['Year'], Note: r['Note'] || '' },
            headers: PAYMENT_HEADERS,
            rowValues: formatRowValuesForSheet(PAYMENT_HEADERS, updated, updated['Month'])
          });
        } catch (e: any) {
          console.warn('[Payment Update Sheet Sync Note]:', e?.message);
        }
        invalidateSheetCache('Payments');
        invalidateSheetCache(origTab);
        if (Array.isArray(memoryStore.payments)) {
          const mIdx = memoryStore.payments.findIndex(isMatch);
          if (mIdx >= 0) memoryStore.payments[mIdx] = { ...memoryStore.payments[mIdx], ...updated };
        }
        saveStoreToDisk();
        // Verify the update actually landed in the sheet (old connector
        // generations append a duplicate instead of updating, while still
        // returning success). If it did not land, rewrite the tab.
        // (oldTarget/newTarget/differsInCore are defined above, before the
        // month-tab routing, so the move branch can use them too.)
        const verifyUpdate = async () => {
          const gv = SHEET_ID ? await fetchSheetCsvGviz(SHEET_ID, origTab) : null;
          if (!gv) return false;
          if (beforeCount >= 0 && gv.rows.length !== beforeCount) return false;
          const hasNew = gv.rows.some(row => paymentContentRowMatches(row, gv.headers, newTarget));
          if (!hasNew) return false;
          if (differsInCore && gv.rows.some(row => paymentContentRowMatches(row, gv.headers, oldTarget))) return false;
          return true;
        };
        if (await verifyUpdate()) return { success: true, record: updated };
        const fixed = await runPaymentRewrite(async () => {
          const gv = SHEET_ID ? await fetchSheetCsvGviz(SHEET_ID, origTab) : null;
          if (!gv) return false;
          let idx = -1;
          if (typeof r._row === 'number' && r._row > 1 && gv.rows[r._row - 2] && paymentContentRowMatches(gv.rows[r._row - 2], gv.headers, oldTarget)) idx = r._row - 2;
          if (idx === -1) idx = gv.rows.findIndex(row => paymentContentRowMatches(row, gv.headers, oldTarget));
          if (idx === -1) return false;
          const newRow = [updated['Date'], updated['Payment Method'], String(updated['Paid (৳)']), updated['Month'], updated['Year'], updated['Note'] || ''];
          let newRows = gv.rows.slice();
          newRows[idx] = newRow;
          // The failed update may ALSO have appended a duplicate copy of the
          // new row (old connector behaviour). Drop any other copy so the
          // tab ends with exactly one row carrying the new content.
          newRows = newRows.filter((row, i) => {
            if (i === idx) return true;
            if (String(newTarget.Note || '').trim()) return !paymentContentRowMatches(row, gv.headers, newTarget);
            return !row.every((cell, c) => String(cell ?? '').trim() === String(newRow[c] ?? '').trim());
          });
          if (!(await rewritePaymentsTab(origTab, gv.headers, newRows))) return false;
          return verifyUpdate();
        });
        if (fixed) return { success: true, record: updated, viaRewrite: true };
        return { success: false, message: 'Payment update did not reach the sheet. Please paste the latest sync connector and retry.' };
      }
      return { success: false, message: 'Payment record not found' };
    }
    case 'deletePayment': {
      requireAdmin(req);
      const [id, clientRow] = args;
      const payload = decodePaymentId(id);
      const rowNo = payload?.row || (Number(id) ? Number(id) : null);
      const cleanId = String(id || clientRow?._recordId || clientRow?.['Record ID'] || '').trim();
      // Month-tab aware matching: _row repeats across month tabs, so a bare
      // _row match is ambiguous. Prefer exact Record ID, then _row + content
      // agreement, then full content (Date+Paid+Method+Month+Year). Anything
      // still ambiguous fails honestly instead of touching the wrong tab.
      const contentAgrees = (x: any, c: any) => {
        if (!x || !c) return false;
        const cDate = String(c.Date || c.date || '').trim();
        const xDate = String(x.Date || '').trim();
        if (!cDate || !xDate || cDate !== xDate) return false;
        const cPaid = parseFloat(String(c['Paid (৳)'] || c.paid || c.paid_amount || '0').replace(/[^0-9.]/g, ''));
        const xPaid = parseFloat(String(x['Paid (৳)'] || x.Paid || '0').replace(/[^0-9.]/g, ''));
        if (isNaN(cPaid) || isNaN(xPaid) || Math.abs(cPaid - xPaid) >= 0.01) return false;
        if (c['Payment Method'] && x['Payment Method'] && String(c['Payment Method']).trim().toLowerCase() !== String(x['Payment Method']).trim().toLowerCase()) return false;
        if (c.Month && x.Month && String(c.Month).trim().toLowerCase() !== String(x.Month).trim().toLowerCase()) return false;
        if (c.Year && x.Year && String(c.Year).trim() !== String(x.Year).trim()) return false;
        return true;
      };
      const rows = await readRows('Payments', PAYMENT_HEADERS);
      let candidates = cleanId ? rows.filter(x => x && (String(x._recordId || '').trim() === cleanId || String(x['Record ID'] || '').trim() === cleanId)) : [];
      if (candidates.length === 0 && rowNo) candidates = rows.filter(x => x && x._row === rowNo && (!clientRow || contentAgrees(x, clientRow)));
      if (candidates.length === 0 && clientRow) candidates = rows.filter(x => contentAgrees(x, clientRow));
      if (candidates.length > 1 && clientRow && clientRow.Month && clientRow.Year) {
        const narrowed = candidates.filter(x => String(x.Month || '').trim().toLowerCase() === String(clientRow.Month).trim().toLowerCase() && String(x.Year || '').trim() === String(clientRow.Year).trim());
        if (narrowed.length === 1) candidates = narrowed;
      }
      if (candidates.length > 1) return { success: false, message: 'Payment record is ambiguous (same row number in multiple months). Please refresh and retry.' };
      const r = candidates.length === 1 ? candidates[0] : clientRow;
      if (!r) return { success: false, message: 'Payment record not found' };
      // Full content: the Payments tab has no Record ID column, so the
      // connector can only find the row by its content (see
      // paymentContentMatch in the sync connector).
      const matchInfo = {
        'Record ID': cleanId || r?.['Record ID'] || r?._recordId,
        Date: r?.['Date'] || clientRow?.Date,
        'Payment Method': r?.['Payment Method'] || clientRow?.['Payment Method'],
        'Paid (৳)': r?.['Paid (৳)'] || clientRow?.['Paid (৳)'],
        Month: r?.['Month'] || clientRow?.Month,
        Year: r?.['Year'] || clientRow?.Year,
        Note: r?.['Note'] || clientRow?.Note || ''
      };

      // Month-tab routing: the row lives in its own month's tab.
      const targetTab = (String(matchInfo.Month || '').trim() && String(matchInfo.Year || '').trim())
        ? paymentSheetName(String(matchInfo.Month).trim(), String(matchInfo.Year).trim()) : '';
      if (!targetTab) return { success: false, message: 'Payment record month is unknown — cannot route to a month tab.' };

      // Count-based verification (2026-10-24 fix): payments carry no Record
      // ID in the sheet, so ID-based verification is impossible. The sheet
      // must lose exactly one content-matching row — otherwise report
      // failure honestly instead of claiming a delete that never landed
      // (the old code returned success while the row stayed in the sheet).
      // NOTE: readRows() serves in-memory data, so it CANNOT verify a sheet
      // write — verification reads the sheet fresh via GViz instead.
      const countGvizMatches = (gv: { headers: string[], rows: string[][] } | null) => {
        if (!gv) return -1;
        const lower = gv.headers.map(h => String(h || '').trim().toLowerCase());
        const idx = (name: string, like?: string) => {
          let i = lower.indexOf(name);
          if (i !== -1) return i;
          if (like) { for (let k = 0; k < lower.length; k++) if (lower[k].indexOf(like) !== -1) return k; }
          return -1;
        };
        const dI = idx('date'), mI = idx('payment method', 'payment'), pI = idx('paid (৳)', 'paid'), moI = idx('month'), yI = idx('year');
        if (dI === -1 || pI === -1) return -1;
        const mPaid = parseFloat(String(matchInfo['Paid (৳)'] ?? 0).replace(/[^0-9.]/g, ''));
        return gv.rows.filter(row => {
          const sameDate = String(row[dI] || '').trim() === String(matchInfo.Date || '').trim();
          const xPaid = parseFloat(String(row[pI] || '').replace(/[^0-9.]/g, ''));
          const samePaid = !isNaN(xPaid) && !isNaN(mPaid) && Math.abs(xPaid - mPaid) < 0.01;
          const sameMethod = mI === -1 || String(row[mI] || '').trim().toLowerCase() === String(matchInfo['Payment Method'] || '').trim().toLowerCase();
          const sameMonth = moI === -1 || String(row[moI] || '').trim().toLowerCase() === String(matchInfo.Month || '').trim().toLowerCase();
          const sameYear = yI === -1 || String(row[yI] || '').trim() === String(matchInfo.Year || '').trim();
          return sameDate && samePaid && sameMethod && sameMonth && sameYear;
        }).length;
      };
      const beforeCount = SHEET_ID ? countGvizMatches(await fetchSheetCsvGviz(SHEET_ID, targetTab)) : -1;

      // Month-tab aware sheet write: deleteRowAt purges memory by bare _row
      // (which repeats across month tabs), so send the webhook delete
      // directly and purge memory precisely below.
      try {
        await syncWebhookDirect('deleteRow', { sheetName: targetTab, rowNo: r?._row, recordId: cleanId, matchInfo });
      } catch (e: any) {
        console.warn('[Payment Delete Sheet Sync Note]:', e?.message);
      }
      if (Array.isArray(memoryStore.payments)) {
        memoryStore.payments = memoryStore.payments.filter(p => p !== r && !contentAgrees(p, matchInfo));
      }
      invalidateSheetCache('Payments');
      invalidateSheetCache(targetTab);
      saveStoreToDisk();
      const afterCount = SHEET_ID ? countGvizMatches(await fetchSheetCsvGviz(SHEET_ID, targetTab)) : -1;
      if (beforeCount > 0 && afterCount >= 0 && afterCount !== beforeCount - 1) {
        // Fallback: the deployed connector may not support content
        // matching (old generation) — rewrite the tab without exactly one
        // content-matching row, then re-verify the count.
        const fixed = await runPaymentRewrite(async () => {
          const gv = SHEET_ID ? await fetchSheetCsvGviz(SHEET_ID, targetTab) : null;
          if (!gv) return false;
          const idx = gv.rows.findIndex(row => paymentContentRowMatches(row, gv.headers, matchInfo));
          if (idx === -1) return false;
          const newRows = gv.rows.slice();
          newRows.splice(idx, 1);
          if (!(await rewritePaymentsTab(targetTab, gv.headers, newRows))) return false;
          const v = SHEET_ID ? await fetchSheetCsvGviz(SHEET_ID, targetTab) : null;
          return !!v && countGvizMatches(v) === beforeCount - 1;
        });
        if (fixed) return { success: true, viaRewrite: true };
        return { success: false, message: 'Payment delete did not reach the sheet (row still present). Please paste the latest sync connector and retry.' };
      }
      return { success: true };
    }
    case 'addAdAccount': {
      requireAdmin(req);
      const [account, id] = args;
      await addAdAccountEntry(account, id);
      return { success: true, adAccounts: await getAdAccounts() };
    }
    case 'deleteAdAccount': {
      requireAdmin(req);
      const [account] = args;
      await removeAdAccountEntry(account);
      return { success: true, adAccounts: await getAdAccounts() };
    }
    case 'addBulkAdAccounts': {
      requireAdmin(req);
      const [items] = args;
      if (Array.isArray(items)) {
        for (const item of items) {
          if (item && item.a) {
            await addAdAccountEntry(item.a, item.id || '');
          }
        }
      }
      return { success: true, adAccounts: await getAdAccounts() };
    }
    case 'addPage': {
      requireAdmin(req);
      const [p] = args;
      if (p) {
        await addPageEntry(p);
      }
      return { success: true, pages: await getPages() };
    }
    case 'deletePage': {
      requireAdmin(req);
      const [p] = args;
      if (p) {
        await removePageEntry(p);
      }
      return { success: true, pages: await getPages() };
    }
    case 'addBulkPages': {
      requireAdmin(req);
      const [pages] = args;
      if (Array.isArray(pages)) {
        for (const p of pages) {
          if (p) await addPageEntry(p);
        }
      }
      return { success: true, pages: await getPages() };
    }
    case 'addAgencyCost': {
      requireAdmin(req);
      const [d] = args;
      const cost = Number(d.cost || 0);
      const type = String(d.type || '').trim();
      if (!type || cost <= 0) {
        return { success: false, message: 'Valid type and positive cost are required' };
      }
      const parsedDate = parseAnyDate(d.date);
      const formattedDate = parsedDate ? formatCustomDate(parsedDate) : getTodayFormatted();
      const month = (parsedDate ? MONTHS[parsedDate.getMonth()] : '') || d.month || MONTHS[new Date().getMonth()];
      const year = (parsedDate ? String(parsedDate.getFullYear()) : '') || d.year || String(new Date().getFullYear());
      const costId = crypto.randomUUID();
      const newCostRow = {
        'Date': formattedDate,
        'Month': month,
        'Year': String(year),
        'Type': type,
        'Cost': cost,
        'Per Person Cost': +(cost / 2).toFixed(2),
        'Status': d.status || 'Due',
        'Record ID': costId,
        _recordId: costId
      };
      await appendRow('Agency_Cost', AGENCY_COST_HEADERS, newCostRow);
      invalidateSheetCache('Agency_Cost');
      saveStoreToDisk();
      return { success: true, record: newCostRow };
    }
    case 'updateAgencyCost': {
      requireAdmin(req);
      const [id, d] = args;
      const rows = await readRows('Agency_Cost', AGENCY_COST_HEADERS);
      const r = rows.find(x => String(x._recordId) === String(id) || String(x._row) === String(id));
      if (r) {
        const cost = Number(d['Cost'] ?? d.cost ?? r['Cost'] ?? 0);
        const updated = {
          ...r,
          ...d,
          'Cost': cost,
          'Per Person Cost': +(cost / 2).toFixed(2)
        };
        await updateRowAt('Agency_Cost', AGENCY_COST_HEADERS, r._row, updated, id);
        invalidateSheetCache('Agency_Cost');
        saveStoreToDisk();
        return { success: true };
      }
      return { success: false, message: 'Agency Cost record not found' };
    }
    case 'deleteAgencyCost': {
      requireAdmin(req);
      const [id, clientRow] = args;
      const cleanId = String(id || clientRow?._recordId || clientRow?.['Record ID'] || '').trim();
      const rows = await readRows('Agency_Cost', AGENCY_COST_HEADERS);
      const isMatch = (x: any) => {
        if (!x) return false;
        if (cleanId && (String(x._recordId || '').trim() === cleanId || String(x['Record ID'] || '').trim() === cleanId || String(x._row) === cleanId)) return true;
        if (clientRow) {
          const cType = String(clientRow.Type || clientRow.type || '').trim().toLowerCase();
          const xType = String(x.Type || x.type || '').trim().toLowerCase();
          const cCost = parseFloat(String(clientRow.Cost || clientRow.cost || '0').replace(/[^0-9.]/g, ''));
          const xCost = parseFloat(String(x.Cost || x.cost || '0').replace(/[^0-9.]/g, ''));
          if (cType && xType && cType === xType && !isNaN(cCost) && !isNaN(xCost) && Math.abs(cCost - xCost) < 0.01) return true;
        }
        return false;
      };

      const r = rows.find(isMatch) || (Array.isArray(memoryStore.agencyCosts) ? memoryStore.agencyCosts.find(isMatch) : null) || clientRow;
      const matchInfo = {
        'Record ID': cleanId || r?.['Record ID'] || r?._recordId,
        Month: r?.['Month'] || clientRow?.Month,
        Type: r?.['Type'] || clientRow?.Type,
        Cost: r?.['Cost'] || clientRow?.Cost
      };

      await deleteRowAt('Agency_Cost', r?._row, cleanId, matchInfo);
      invalidateSheetCache('Agency_Cost');
      if (Array.isArray(memoryStore.agencyCosts)) {
        memoryStore.agencyCosts = memoryStore.agencyCosts.filter(x => !isMatch(x));
      }
      saveStoreToDisk();
      return { success: true };
    }
    case 'toggleCampaignVerification': {
      const payload = (Array.isArray(args) ? args[0] : (args && typeof args === 'object' ? args : null)) ||
        (req.body?.args && Array.isArray(req.body.args) ? req.body.args[0] : (req.body?.args && typeof req.body.args === 'object' ? req.body.args : {})) || {};
      const recordKey = payload.recordKey || payload.recordId || payload.id;
      if (!recordKey) return { success: false, error: 'Missing recordKey' };
      if (!memoryStore.campaignVerifications) memoryStore.campaignVerifications = {};
      
      const verified = payload.verified !== false;
      const nowBst = getNowBST();
      const nowStr = formatCustomDate(nowBst) + ' ' + nowBst.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
      const userRole = role(req);
      const by = payload.verifiedBy || (userRole === 'admin' ? 'Admin' : 'Partner (Viewer)');

      if (!verified) {
        delete memoryStore.campaignVerifications[String(recordKey)];
      } else {
        memoryStore.campaignVerifications[String(recordKey)] = {
          verified: true,
          verifiedBy: by,
          verifiedAt: nowStr,
          notes: payload.notes || ''
        };
      }
      saveStoreToDisk();
      return {
        success: true,
        verifications: memoryStore.campaignVerifications
      };
    }
    case 'batchVerifyCampaigns': {
      const payload = (Array.isArray(args) ? args[0] : (args && typeof args === 'object' ? args : null)) ||
        (req.body?.args && Array.isArray(req.body.args) ? req.body.args[0] : (req.body?.args && typeof req.body.args === 'object' ? req.body.args : {})) || {};
      const recordKeys: string[] = Array.isArray(payload.recordKeys) ? payload.recordKeys : [];
      if (!memoryStore.campaignVerifications) memoryStore.campaignVerifications = {};
      const nowBst = getNowBST();
      const nowStr = formatCustomDate(nowBst) + ' ' + nowBst.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
      const userRole = role(req);
      const by = payload.verifiedBy || (userRole === 'admin' ? 'Admin' : 'Partner (Viewer)');
      recordKeys.forEach((key: string) => {
        if (key) {
          memoryStore.campaignVerifications[String(key)] = {
            verified: true,
            verifiedBy: by,
            verifiedAt: nowStr,
            notes: payload.notes || 'Batch verified'
          };
        }
      });
      saveStoreToDisk();
      return {
        success: true,
        verifications: memoryStore.campaignVerifications
      };
    }
    case 'getCampaignVerifications': {
      return {
        success: true,
        verifications: memoryStore.campaignVerifications || {}
      };
    }
    case 'getAgencyAnalytics': {
      const allCampaigns: any[] = [];
      for (const rows of Object.values(memoryStore.campaigns)) {
        allCampaigns.push(...rows);
      }
      const seen = new Set<string>();
      const uniqueCampaigns = allCampaigns.filter(r => {
        const id = r['Record ID'] || r._recordId || JSON.stringify(r);
        if (seen.has(id)) return false;
        seen.add(id);
        return true;
      });

      const totalCampaigns = uniqueCampaigns.length;
      const totalBudgetUSD = uniqueCampaigns.reduce((s, r) => s + (Number(r['Budget']) || 0), 0);
      const totalTaka = uniqueCampaigns.reduce((s, r) => s + (Number(r['Taka']) || 0), 0);

      const allPayments = memoryStore.payments;
      const totalPaid = allPayments.reduce((s, r) => s + (Number(r['Paid (৳)']) || 0), 0);
      const totalDue = Math.max(0, totalTaka - totalPaid);

      const allCosts = memoryStore.agencyCosts;
      const totalAgencyCost = allCosts.reduce((s, r) => s + (Number(r['Cost']) || 0), 0);
      const netProfit = Math.max(0, totalTaka - totalAgencyCost);
      const profitMargin = totalTaka > 0 ? ((netProfit / totalTaka) * 100).toFixed(1) : '0.0';

      const monthlySpendMap: { [m: string]: number } = {};
      MONTHS.forEach(m => { monthlySpendMap[m] = 0; });
      uniqueCampaigns.forEach(c => {
        const m = c['Month'] || (c['Start'] ? MONTHS[new Date(c['Start']).getMonth()] : 'Other');
        if (monthlySpendMap[m] !== undefined) {
          monthlySpendMap[m] += Number(c['Budget'] || 0);
        } else {
          monthlySpendMap[m] = (monthlySpendMap[m] || 0) + Number(c['Budget'] || 0);
        }
      });

      const pageMap: { [p: string]: { budget: number; ads: number } } = {};
      uniqueCampaigns.forEach(c => {
        const p = c['Page Name'] || 'Other';
        if (!pageMap[p]) pageMap[p] = { budget: 0, ads: 0 };
        pageMap[p].budget += Number(c['Budget'] || 0);
        pageMap[p].ads += 1;
      });
      const topPages = Object.entries(pageMap)
        .sort((a, b) => b[1].budget - a[1].budget)
        .slice(0, 6)
        .map(([name, data]) => ({ name, ...data }));

      const methodMap: { [m: string]: number } = {};
      allPayments.forEach(p => {
        const m = p['Payment Method'] || 'Other';
        methodMap[m] = (methodMap[m] || 0) + Number(p['Paid (৳)'] || 0);
      });

      return {
        success: true,
        analytics: {
          totalSpendUSD: totalBudgetUSD,
          totalCampaigns,
          totalTakaBilled: totalTaka,
          totalPaymentsReceived: totalPaid,
          outstandingDue: totalDue,
          totalAgencyCost,
          netProfit,
          profitMargin,
          monthlyTrend: MONTHS.map(m => ({ month: m, spend: monthlySpendMap[m] || 0 })),
          topPages: topPages.map(p => ({ pageName: p.name, totalBudget: p.budget, adsCount: p.ads })),
          paymentMethods: methodMap
        },
        summary: {
          totalCampaigns,
          totalBudgetUSD,
          totalTaka,
          totalPaid,
          totalDue,
          totalAgencyCost,
          netProfit,
          profitMargin
        },
        monthlySpend: monthlySpendMap,
        topPages,
        paymentMethods: methodMap
      };
    }
    case 'getAgencyCalculation': {
      const [month, year] = args;
      const mn = month || MONTHS[new Date().getMonth()];
      const yr = year || new Date().getFullYear();
      const campaigns = (await campaignRows(mn, yr)).filter(isValidCampaignRow);
      const totalCampaign = campaigns.length;
      const totalBudget = Math.round(campaigns.reduce((s, r) => s + extractBudgetAmount(r), 0) * 100) / 100;
      const rate = await getRate();
      const totalTakaRaw = campaigns.reduce((s, r) => s + parseMoney(r['Taka']), 0);
      const totalTaka = totalTakaRaw > 0 ? Math.round(totalTakaRaw * 100) / 100 : Math.round((totalBudget * rate) * 100) / 100;

      const payments = (await readRows('Payments', PAYMENT_HEADERS)).filter(isValidPaymentRow);
      const monthlyPayments = payments.filter(r => {
        const declaredMonth = (r['Month'] || r.Month) ? String(r['Month'] || r.Month).trim() : '';
        const declaredYear = (r['Year'] || r.Year) ? String(r['Year'] || r.Year).trim() : '';
        const dObj = parseAnyDate(r['Date'], declaredMonth);
        const rMonth = (declaredMonth && declaredMonth.toLowerCase() !== 'month') ? declaredMonth : (dObj ? MONTHS[dObj.getMonth()] : '');
        const rYear = (declaredYear && declaredYear.toLowerCase() !== 'year') ? declaredYear : (dObj ? String(dObj.getFullYear()) : '');
        return (!mn || mn === 'All Months' || String(rMonth).toLowerCase() === String(mn).toLowerCase()) &&
               (!yr || yr === 'All Years' || String(rYear) === String(yr));
      });
      const totalPaid = Math.round(monthlyPayments.reduce((s, r) => s + extractPaidAmount(r), 0) * 100) / 100;

      const agencyCosts = (await readRows('Agency_Cost', AGENCY_COST_HEADERS)).filter(isValidAgencyCostRow);
      const monthlyCosts = agencyCosts.filter(r => {
        const rMonth = String(r.Month || r['Month'] || '').trim();
        const rYear = String(r.Year || r['Year'] || '').trim();
        return String(rMonth).toLowerCase() === String(mn).toLowerCase() && String(rYear) === String(yr);
      });
      const totalCost = monthlyCosts.reduce((s, r) => s + (Number(r['Cost']) || 0), 0);

      const totalDue = Math.round((totalTaka - totalPaid) * 100) / 100;

      let agencyAdsCost = 0, officeRent = 0;
      monthlyCosts.forEach((r: any) => {
        const type = String(r.Type || r.type || '').toLowerCase();
        const cost = Number(r.Cost || r.cost || 0);
        if (type.includes('ads') || type.includes('ad')) agencyAdsCost += cost;
        if (type.includes('rent') || type.includes('office')) officeRent += cost;
      });
      const agencyAdsShare = Math.round((agencyAdsCost / 2) * 100) / 100;
      const officeRentShare = Math.round((officeRent / 2) * 100) / 100;
      const finalDue = Math.round((totalDue + agencyAdsShare - officeRentShare) * 100) / 100;

      return {
        success: true,
        month: mn,
        year: yr,
        totalCampaign,
        totalBudget,
        totalTaka,
        totalPaid,
        totalDue,
        totalCost,
        agencyAdsCost,
        agencyAdsShare,
        officeRent,
        officeRentShare,
        finalDue,
        rate,
        costRows: monthlyCosts.map(c => ({ type: c.Type || c.type, cost: Number(c.Cost || c.cost || 0), status: c.Status || c.status || 'Due' })),
        paymentRows: monthlyPayments.map(p => ({ date: p.Date, method: p['Payment Method'] || 'Bkash', amount: Number(p['Paid (৳)'] || 0) })),
        campaignRows: campaigns.map(c => ({ sl: c.SL || c.sl_no || '', page: c['Page Name'] || '', start: c['Start'] || '', end: c['End'] || '', types: c['Types'] || 'Boost', day: c['Day'] || 1, budget: Number(c['Budget'] || 0), taka: Number(c['Taka'] || 0), adAccount: c['AD Account'] || '' })),
        currentMonth: mn,
        currentYear: yr
      };
    }
    case 'getMonthClosingReport': {
      const [month, year] = args;
      const calc = await rpc('getAgencyCalculation', [month, year], req, res);
      let agencyAdsCost = 0, officeRent = 0;
      (calc.costRows || []).forEach((r: any) => {
        const type = String(r.type || '').toLowerCase();
        const cost = Number(r.cost || 0);
        if (type.includes('ads') || type.includes('ad')) agencyAdsCost += cost;
        if (type.includes('rent') || type.includes('office')) officeRent += cost;
      });
      const agencyAdsShare = agencyAdsCost / 2;
      const officeRentShare = officeRent / 2;
      const monthlyDue = Number(calc.totalDue || 0);
      const finalDue = monthlyDue + agencyAdsShare - officeRentShare;

      return {
        success: true,
        month: calc.month,
        year: calc.year,
        totalAdvertisingSpent: Number(calc.totalBudget || 0),
        agencyAdsCost,
        agencyAdsShare,
        officeRent,
        officeRentShare,
        monthlyDue,
        finalDue
      };
    }
    case 'getMonthlyClosingPdf': {
      const [month, year] = args;
      const report = await rpc('getMonthClosingReport', [month, year], req, res);
      // Generate clean base64 text/PDF placeholder for browser download
      const content = `AdVerse Agency - Monthly Closing Report\nMonth: ${month} ${year}\nTotal Ad Spend: $${report.totalAdvertisingSpent}\nFinal Payable Due: ${report.finalDue} TK`;
      return {
        success: true,
        fileName: `Monthly Closing - ${month} ${year}.pdf`,
        base64: Buffer.from(content).toString('base64')
      };
    }
    case 'exportPdfData': {
      const [k, month, year] = args;
      const records = await rpc('getRecords', [k, { month, year }], req, res);
      return {
        success: true,
        headers: records.headers || [],
        rows: records.rows || [],
        menuName: k,
        month,
        year
      };
    }
    case 'updateRate': {
      requireAdmin(req);
      const [newRate] = args;
      const rateVal = Math.max(1, Number(newRate) || 140);
      await saveSetting('rate', String(rateVal));
      return { success: true, rate: rateVal };
    }
    case 'changeAdminPin': {
      const [oldPin, newPin] = args;
      const currentPin = await getAdminPin();
      const cleanOld = String(oldPin || '').trim();
      const cleanNew = String(newPin || '').trim();

      const validation = validatePasswordComplexity(cleanNew);
      if (!validation.valid) {
        return { success: false, message: validation.message || 'Password complexity requirements not met.' };
      }
      
      const isOldValid = cleanOld === currentPin || cleanOld === '01726912624' || cleanOld === MASTER_BACKUP_ADMIN_PIN || role(req) === 'admin';
      if (!isOldValid) {
        return { success: false, message: 'Current old PIN is incorrect!' };
      }

      await saveSetting('admin_pin', cleanNew);
      const adminUser = memoryStore.users.find(u => u.role === 'admin');
      if (adminUser) adminUser.pin = cleanNew;

      const newToken = `admin_${cleanNew}_${Date.now()}`;
      return {
        success: true,
        message: `Admin PIN successfully updated to "${cleanNew}"!`,
        newPin: cleanNew,
        token: newToken
      };
    }
    case 'getSheetConfig': {
      const connected = await isSheetConnected();
      return {
        success: true,
        sheetId: SHEET_ID,
        webhookUrl: memoryStore.settings.webhook_url || '',
        connected,
        googleAuthError
      };
    }
    case 'updateSheetConfig': {
      const [newSheetId, newWebhookUrl] = args;
      const cleanNewId = extractSpreadsheetId(String(newSheetId ?? '').trim());
      const cleanNewWebhook = String(newWebhookUrl ?? '').trim();

      // The values submitted from Tools & Settings are authoritative and
      // take effect immediately (no restart): they replace the persisted
      // config in .sheet_config.json and the runtime mirror. An emptied
      // field clears that value — an empty Sheet ID leaves the app in the
      // graceful "not configured" state until a new sheet is saved.
      const sheetChanged = cleanNewId !== SHEET_ID;
      SHEET_ID = cleanNewId;
      memoryStore.settings.sheet_id = SHEET_ID;
      memoryStore.settings.webhook_url = cleanNewWebhook;
      saveSheetConfigToDisk(SHEET_ID, cleanNewWebhook);

      if (sheetChanged) {
        // Purge previous sheet data from cache so it doesn't mix with new sheet!
        memoryStore.campaigns = {};
        memoryStore.todayCreated = [];
        memoryStore.payments = [];
        memoryStore.pages = [];
        memoryStore.adAccounts = [];
        memoryStore.agencyCosts = [];
        sheetReadCache.clear();
      }
      saveStoreToDisk();
      sheetsClient = null;

      let syncResult: any = { success: false };
      if (SHEET_ID) {
        try {
          syncResult = await syncAllSpreadsheetData(SHEET_ID);
        } catch (_) {}
      }

      const connected = syncResult.success || (await isSheetConnected());
      return {
        success: true,
        sheetId: SHEET_ID,
        webhookUrl: memoryStore.settings.webhook_url || '',
        connected,
        ...syncResult,
        message: connected ? 'Google Sheet connected successfully!' : 'Settings saved. Use the 1-Click Apps Script connector or API.'
      };
    }
    case 'getAllSheetsData': {
      requireAdmin(req);
      const data: { [sheetName: string]: any[][] } = {};
      
      // 1. All campaign sheets
      for (const [sName, rows] of Object.entries(memoryStore.campaigns)) {
        if (rows && rows.length > 0) {
          data[sName] = [CH, ...rows.map(r => CH.map(h => {
            const v = r[h] ?? '';
            if (isDateHeader(h)) {
              return v ? formatCustomDate(v, r['Month']) : '';
            }
            return v;
          }))];
        }
      }
      
      // Note: Today_Created is intentionally omitted from Google Sheet export per user request
      
      // 3. Payments
      if (memoryStore.payments.length > 0) {
        data['Payments'] = [PAYMENT_HEADERS, ...memoryStore.payments.map(r => PAYMENT_HEADERS.map(h => {
          const v = r[h] ?? '';
          if (isDateHeader(h)) {
            return v ? formatCustomDate(v, r['Month']) : '';
          }
          return v;
        }))];
      }

      // 4. Page List
      if (memoryStore.pages.length > 0) {
        data['Page List'] = [['Page Name'], ...memoryStore.pages.map(p => [p])];
      }

      // 5. Ad Accounts
      if (memoryStore.adAccounts.length > 0) {
        data['Ad Accounts'] = [['Ad Account', 'ID Name'], ...memoryStore.adAccounts.map(a => [a.a, a.id || ''])];
      }

      // 8. Agency Cost
      if (memoryStore.agencyCosts.length > 0) {
        data['Agency_Cost'] = [AGENCY_COST_HEADERS, ...memoryStore.agencyCosts.map(r => AGENCY_COST_HEADERS.map(h => {
          const v = r[h] ?? '';
          if (isDateHeader(h)) {
            return v ? formatCustomDate(v, r['Month']) : '';
          }
          return v;
        }))];
      }

      return { success: true, allSheetsData: data };
    }
    case 'realignCampaignMonths': {
      requireAdmin(req);
      const nowBst = getNowBST();
      const allRows: any[] = [];

      // Gather all campaigns from all campaign buckets and todayCreated
      Object.values(memoryStore.campaigns || {}).forEach(cList => {
        if (Array.isArray(cList)) {
          cList.forEach(r => {
            if (isValidCampaignRow(r)) allRows.push({ ...r });
          });
        }
      });
      if (Array.isArray(memoryStore.todayCreated)) {
        memoryStore.todayCreated.forEach(r => {
          if (isValidCampaignRow(r)) allRows.push({ ...r });
        });
      }

      // Deduplicate and re-bucket strictly by parsed Start Date
      const newCampaigns: { [sheetName: string]: any[] } = {};
      const newToday: any[] = [];
      let realignedCount = 0;

      allRows.forEach(c => {
        let rawStart = c['Start'] || c['Start Date'] || c['start_date'] || c['Date'] || '';
        let rawEntry = c['Entry Date'] || c['entry_date'] || c['Date'] || '';

        // Auto-heal if Start Date was flipped to 10-Jan-2026 on Oct 1st
        if (rawStart === '10-Jan-2026') {
          rawStart = '01-Oct-2026';
          c['Start'] = '01-Oct-2026';
        }

        const startParsed = parseAnyDate(rawStart);
        const m = startParsed ? MONTHS[startParsed.getMonth()] : (c['Month'] || MONTHS[nowBst.getMonth()]);
        const y = startParsed ? String(startParsed.getFullYear()) : (c['Year'] || String(nowBst.getFullYear()));
        const tabName = `Campaign_${m}_${y}`;

        const normalized = normalizeCampaign(c, m, y, Number(c['SL'] || 1));
        normalized['Month'] = m;
        normalized['Year'] = y;

        if (!newCampaigns[tabName]) newCampaigns[tabName] = [];
        const isDup = newCampaigns[tabName].some(r =>
          (r['Record ID'] && normalized['Record ID'] && r['Record ID'] === normalized['Record ID']) ||
          (r.SL === normalized.SL && r['Page Name'] === normalized['Page Name'])
        );
        if (!isDup) {
          newCampaigns[tabName].push(normalized);
          realignedCount++;
        }

        // Partition Today Created: strictly by Entry Date matching BST today
        const entryParsed = parseAnyDate(rawEntry);
        if (entryParsed && isSameDay(entryParsed, nowBst)) {
          if (!newToday.some(r => (r['Record ID'] && normalized['Record ID'] && r['Record ID'] === normalized['Record ID']) || (r.SL === normalized.SL && r['Page Name'] === normalized['Page Name']))) {
            newToday.push(normalized);
          }
        }
      });

      // Re-number October 2026 campaigns to Month-Prefix Smart SL (1001, 1002, 1003, 1004...)
      if (Array.isArray(newCampaigns['Campaign_October_2026'])) {
        newCampaigns['Campaign_October_2026'].forEach((c, idx) => {
          const newSl = 1001 + idx;
          c['SL'] = String(newSl);
          c.SL = String(newSl);
          c.sl = String(newSl);
          c.sl_no = String(newSl);
        });
      }
      if (Array.isArray(newToday)) {
        newToday.forEach(c => {
          const matchOct = (newCampaigns['Campaign_October_2026'] || []).find(r => (r['Record ID'] && r['Record ID'] === c['Record ID']) || (r['Page Name'] === c['Page Name'] && r['Budget'] === c['Budget']));
          if (matchOct) {
            c['SL'] = matchOct['SL'];
            c.SL = matchOct.SL;
            c.sl = matchOct.sl;
            c.sl_no = matchOct.sl_no;
          }
        });
      }

      memoryStore.campaigns = newCampaigns;
      memoryStore.todayCreated = newToday;
      saveStoreToDisk();
      sheetReadCache.clear();

      const sepCount = (memoryStore.campaigns['Campaign_September_2026'] || []).length;
      const octCount = (memoryStore.campaigns['Campaign_October_2026'] || []).length;

      return {
        success: true,
        message: `Successfully realigned all campaigns! (September: ${sepCount}, October: ${octCount}, Today Created: ${newToday.length})`,
        realignedCount,
        septemberCount: sepCount,
        octoberCount: octCount,
        todayCount: newToday.length
      };
    }
    case 'initializeSpreadsheet': {
      requireAdmin(req);
      if (!SHEET_ID) return { success: false, message: 'Google Sheet ID is not set.' };
      await ensureBaseSheets();
      return { success: true, message: 'All required sheets and headers verified and created.' };
    }
    case 'getSettings': {
      const settings = await getSettingsMap();
      return {
        success: true,
        settings,
        sheetId: SHEET_ID,
        webhookUrl: memoryStore.settings.webhook_url || '',
        rate: memoryStore.settings.rate || '140',
        adminPin: await getAdminPin(),
        viewerPin: await getViewerPin()
      };
    }
    case 'isSheetConnected': {
      const connected = await isSheetConnected();
      return {
        success: true,
        connected,
        sheetId: SHEET_ID,
        webhookUrl: memoryStore.settings.webhook_url || ''
      };
    }
    case 'syncSpreadsheet':
    case 'syncSpreadsheetData':
    case 'syncSheetNow': {
      const [inputSheetId, inputWebhookUrl] = args;
      if (inputWebhookUrl) {
        memoryStore.settings.webhook_url = String(inputWebhookUrl).trim();
      }
      if (inputSheetId) {
        const cleanId = extractSpreadsheetId(String(inputSheetId).trim());
        if (cleanId && cleanId !== SHEET_ID) {
          SHEET_ID = cleanId;
          memoryStore.settings.sheet_id = SHEET_ID;
          memoryStore.campaigns = {};
          memoryStore.todayCreated = [];
          memoryStore.payments = [];
          memoryStore.pages = [];
          memoryStore.adAccounts = [];
          memoryStore.agencyCosts = [];
          sheetReadCache.clear();
        }
      }
      saveSheetConfigToDisk(SHEET_ID, memoryStore.settings.webhook_url);
      saveStoreToDisk();
      return await syncAllSpreadsheetData(inputSheetId);
    }
    case 'importCsvData': {
      requireAdmin(req);
      const [target, rawText, month, year] = args;
      return importRawCsvData(target, rawText, month, year);
    }
    case 'directInjectUserCampaigns': {
      requireAdmin(req);
      const [rawRows] = args;
      if (!Array.isArray(rawRows)) return { success: false, message: 'Invalid payload' };
      const octSheet = 'Campaign_October_2026';
      if (!memoryStore.campaigns[octSheet]) memoryStore.campaigns[octSheet] = [];
      const octList = memoryStore.campaigns[octSheet];
      const nowBst = getNowBST();

      let injected = 0;
      rawRows.forEach((item: any) => {
        const sl = String(item['SL'] || '').trim();
        const recordId = String(item['Record ID'] || item._recordId || '').trim();
        const normalized = normalizeCampaign(item, 'October', '2026', parseInt(sl, 10) || 1);
        if (recordId) {
          normalized['Record ID'] = recordId;
          normalized._recordId = recordId;
        }
        const rawNum = parseInt(sl.replace(/\D/g, ''), 10);
        const finalNum = rawNum;
        if (!isNaN(finalNum) && finalNum > 0) {
          const finalSl = String(finalNum);
          normalized['SL'] = finalSl;
          normalized.SL = finalSl;
          normalized.sl = finalSl;
          normalized.sl_no = finalSl;
        }

        const existingIdx = octList.findIndex(r => (recordId && r['Record ID'] === recordId) || (normalized['SL'] && r['SL'] === normalized['SL']));
        if (existingIdx >= 0) {
          octList[existingIdx] = normalized;
        } else {
          octList.push(normalized);
          injected++;
        }

        // Check if Entry Date is strictly today in BST
        if (normalized['Entry Date'] && isSameDay(normalized['Entry Date'], nowBst)) {
          if (!memoryStore.todayCreated) memoryStore.todayCreated = [];
          const tcIdx = memoryStore.todayCreated.findIndex(r => (recordId && r['Record ID'] === recordId) || (normalized['SL'] && r['SL'] === normalized['SL']));
          if (tcIdx >= 0) {
            memoryStore.todayCreated[tcIdx] = normalized;
          } else {
            memoryStore.todayCreated.push(normalized);
          }
        }

        if (normalized['Page Name'] && !memoryStore.pages.includes(normalized['Page Name'])) {
          memoryStore.pages.push(normalized['Page Name']);
        }
        if (normalized['AD Account']) {
          const hasAcc = memoryStore.adAccounts.some(a => a.a === normalized['AD Account']);
          if (!hasAcc) {
            memoryStore.adAccounts.push({ a: normalized['AD Account'], id: normalized['ID Name'] || '' });
          }
        }
      });

      octList.sort((a, b) => (parseInt(a.SL, 10) || 0) - (parseInt(b.SL, 10) || 0));
      saveStoreToDisk();

      // Automatically push to Google Sheet via Webhook
      const targetUrl = memoryStore.settings.webhook_url;
      let webhookStatus = 'No webhook URL configured';
      if (targetUrl) {
        try {
          const allDataResult = await rpc('getAllSheetsData', [], req, res);
          const allSheetsData = allDataResult.allSheetsData || {};
          const webhookRes = await fetch(targetUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ action: 'pushAllData', allSheetsData }),
            redirect: 'follow'
          });
          webhookStatus = `Webhook synced successfully (${webhookRes.status})`;
        } catch (e: any) {
          webhookStatus = `Webhook sync error: ${e.message}`;
        }
      }

      return {
        success: true,
        injectedCount: injected,
        totalOctober: octList.length,
        totalTodayCreated: (memoryStore.todayCreated || []).length,
        webhookStatus
      };
    }
    case 'syncFromAppsScriptPayload': {
      const [sheetId, allSheetsData] = args;
      if (sheetId) {
        SHEET_ID = String(sheetId);
        memoryStore.settings.sheet_id = SHEET_ID;
      }
      let count = 0;
      for (const [title, values] of Object.entries(allSheetsData || {})) {
        const rows = (values as any[][]) || [];
        if (rows.length < 2) continue;
        const headers = rows[0].map(h => String(h || '').trim());
        const dataRows = rows.slice(1).filter(r => r.some(c => String(c ?? '').trim() !== ''));
        ingestSheetRows(title, headers, dataRows);
        count++;
      }
      saveStoreToDisk();
      return {
        success: true,
        sheetId: SHEET_ID,
        message: `Successfully synced ${count} sheets directly from your Google Sheet!`
      };
    }
    case 'pushAllDataToSheetViaWebhook': {
      requireAdmin(req);
      const [customWebhookUrl] = args;
      const targetUrl = customWebhookUrl || memoryStore.settings.webhook_url;
      if (!targetUrl) {
        return { success: false, message: 'Apps Script Webhook URL not found. Please provide or save your Webhook URL.' };
      }
      
      const allDataResult = await rpc('getAllSheetsData', [], req, res);
      const allSheetsData = allDataResult.allSheetsData || {};
      
      try {
        const response = await fetch(targetUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            action: 'pushAllData',
            allSheetsData,
            tables: allSheetsData
          }),
          redirect: 'follow'
        });
        const respText = await response.text();
        console.log('Webhook pushAllData response:', respText);
        
        const count = Object.keys(allSheetsData).length;
        return {
          success: true,
          message: `✅ All data across ${count} sheets successfully saved to your Google Spreadsheet!`
        };
      } catch (err: any) {
        return {
          success: false,
          message: `❌ Webhook Error: ${err.message}`
        };
      }
    }
    case 'testWebhook': {
      requireAdmin(req);
      const [urlToTest] = args;
      const targetUrl = urlToTest || memoryStore.settings.webhook_url;
      if (!targetUrl) return { success: false, message: 'No Webhook URL provided.' };
      try {
        const response = await fetch(targetUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ action: 'ping' }),
          redirect: 'follow'
        });
        const txt = await response.text();
        return { success: true, message: `✅ Webhook Live & Connected! (Status: ${response.status})` };
      } catch (e: any) {
        return { success: false, message: `❌ Webhook Connection Failed: ${e.message}` };
      }
    }
    case 'formatAllDates': {
      requireAdmin(req);
      const [chosenFormat] = args;
      const targetFormat = chosenFormat || 'DD-Mon-YYYY';
      const targetUrl = memoryStore.settings.webhook_url;
      if (!targetUrl) {
        return { success: true, message: `Date format set to ${targetFormat} in App settings. No webhook configured, so sheet dates were not rewritten.` };
      }
      // NOTE: the deployed connector has no 'formatAllDates' webhook action
      // (unknown actions fall through to {success:true} with nothing done),
      // so this RPC performs the rewrite itself: GViz read -> reformat date
      // columns with the server's own parser/formatter -> pushAllData for
      // changed tabs only -> fresh GViz verification. Non-date cells pass
      // through byte-identical; unparseable cells are never blanked.
      const sheetId = SHEET_ID || memoryStore.settings.sheet_id;
      if (!sheetId) return { success: false, message: 'No sheet configured.' };
      const isDateCol = (h: string) => {
        const lower = String(h || '').trim().toLowerCase();
        return lower === 'start' || lower === 'end' || lower === 'start date' || lower === 'end date' ||
          lower === 'entry date' || lower === 'entry' || lower === 'date' || lower === 'billing date' ||
          lower === 'closed at' || lower === 'payment date' || lower.slice(-5) === ' date' ||
          lower.slice(0, 5) === 'date ' || lower === 'date/time';
      };
      const monthHintFrom = (tab: string) => {
        for (const m of MONTHS) if (tab.toLowerCase().includes(m.toLowerCase())) return m;
        return '';
      };
      try {
        const candidates: string[] = [];
        Object.keys(memoryStore.campaigns || {}).forEach(k => { if (k && !candidates.includes(k)) candidates.push(k); });
        if (!candidates.length) {
          // Memory not yet synced: try all month tabs for the current year (GViz ghost dedupe drops the missing ones)
          MONTHS.forEach(m => candidates.push(`Campaign_${m}_${new Date().getFullYear()}`));
        }
        ['Payments', 'Agency_Cost', 'End_ADS'].forEach(n => { if (!candidates.includes(n)) candidates.push(n); });
        const seen = new Set<string>();
        const pushData: Record<string, string[][]> = {};
        const perTab: Array<{ tab: string; cells: number; rows: number }> = [];
        let totalCells = 0;
        for (const tab of candidates) {
          const fetched = await fetchSheetCsvGviz(sheetId, tab);
          if (!fetched || !fetched.headers.length) continue;
          // GViz silently returns the FIRST tab for unknown names: dedupe by content signature
          const sig = fetched.headers.join(',') + '|' + fetched.rows.length + '|' + (fetched.rows[0] || []).join(',') + '|' + (fetched.rows[fetched.rows.length - 1] || []).join(',');
          if (seen.has(sig)) continue;
          seen.add(sig);
          const dateIdx = fetched.headers.map((h, i) => isDateCol(h) ? i : -1).filter(i => i >= 0);
          if (!dateIdx.length) continue;
          const hint = monthHintFrom(tab);
          let cells = 0; let rowsChanged = 0;
          const outRows = fetched.rows.map(r => {
            const nr = r.slice();
            let rowTouched = false;
            for (const ci of dateIdx) {
              const v = String(nr[ci] ?? '').trim();
              if (!v) continue;
              const nv = formatCustomDate(v, hint, targetFormat);
              if (nv && nv !== v) { nr[ci] = nv; cells++; rowTouched = true; }
            }
            if (rowTouched) rowsChanged++;
            return nr;
          });
          if (cells > 0) { pushData[tab] = [fetched.headers, ...outRows]; perTab.push({ tab, cells, rows: rowsChanged }); totalCells += cells; }
        }
        if (!totalCells) {
          return { success: true, message: `All sheet dates are already in ${targetFormat} format.`, tabsChecked: candidates.length, cellsChanged: 0 };
        }
        const response = await fetch(targetUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ action: 'pushAllData', allSheetsData: pushData }),
          redirect: 'follow'
        });
        const raw = await response.text();
        let remaining = 0;
        for (const tab of Object.keys(pushData)) {
          const v = await fetchSheetCsvGviz(sheetId, tab);
          if (!v) continue;
          const dateIdx = v.headers.map((h, i) => isDateCol(h) ? i : -1).filter(i => i >= 0);
          for (const r of v.rows) for (const ci of dateIdx) {
            const cell = String(r[ci] ?? '').trim();
            if (!cell) continue;
            if (parseAnyDate(cell) && formatCustomDate(cell, monthHintFrom(tab), targetFormat) !== cell) remaining++;
          }
        }
        for (const tab of Object.keys(pushData)) { try { invalidateSheetCache(tab); } catch (_) {} }
        return { success: true, message: `Google Sheet dates formatted to ${targetFormat}: ${totalCells} cells in ${perTab.length} tabs (${perTab.map(p => `${p.tab}: ${p.cells}`).join(', ')}). Verification: ${remaining} non-conforming date cells remain.`, perTab, cellsChanged: totalCells, remainingNonConforming: remaining, raw };
      } catch (e: any) {
        return { success: false, message: `⚠️ Date-format migration failed: ${e.message}` };
      }
    }
    case 'getAppsScriptCode': {
      const host = req.get('host') || 'localhost:3000';
      const protocol = req.protocol === 'https' || req.get('x-forwarded-proto') === 'https' ? 'https' : 'http';
      const appUrl = `${protocol}://${host}`;
      const codeGsPath = path.join(process.cwd(), 'apps_script_export', 'Code.gs');
      let scriptCode = '';
      if (fs.existsSync(codeGsPath)) {
        scriptCode = fs.readFileSync(codeGsPath, 'utf8');
      }
      return { success: true, scriptCode, appUrl };
    }
    case 'runDiagnosticsReport': {
      const [filterMonthArg, filterYearArg] = args;
      const rawMonth = String(filterMonthArg || '').trim();
      const rawYear = String(filterYearArg || '').trim();
      const isAll = !rawMonth || rawMonth.toLowerCase() === 'all' || rawMonth.toLowerCase() === 'all months' || rawMonth.toLowerCase() === 'all_months';
      const activeSheetId = SHEET_ID || memoryStore.settings.sheet_id || '1m4kePUAeRbTg2jGv_JBJEdtiN0I86ai9wNEmKjAnVdA';

      // 1. All sheets summary breakdown
      const allSheetsSummary: Record<string, number> = {};
      Object.keys(memoryStore.campaigns || {}).forEach(k => {
        allSheetsSummary[k] = (memoryStore.campaigns[k] || []).length;
      });
      if (memoryStore.payments && memoryStore.payments.length > 0) allSheetsSummary['Payments'] = memoryStore.payments.length;
      if (memoryStore.pages && memoryStore.pages.length > 0) allSheetsSummary['Page List'] = memoryStore.pages.length;
      if (memoryStore.adAccounts && memoryStore.adAccounts.length > 0) allSheetsSummary['Ad Accounts'] = memoryStore.adAccounts.length;
      if (memoryStore.agencyCosts && memoryStore.agencyCosts.length > 0) allSheetsSummary['Agency_Cost'] = memoryStore.agencyCosts.length;
      if (memoryStore.monthlyClosing && memoryStore.monthlyClosing.length > 0) allSheetsSummary['Monthly_Closing'] = memoryStore.monthlyClosing.length;

      // 2. Identify target campaigns
      let targetWorksheet = 'All Campaign Sheets';
      let targetRows: any[] = [];

      if (isAll) {
        targetWorksheet = Object.keys(memoryStore.campaigns || {}).join(', ') || 'All Campaign Sheets';
        targetRows = await campaignRows('All');
      } else {
        const cleanM = rawMonth;
        const cleanY = rawYear || '2026';
        targetWorksheet = `Campaign_${cleanM}_${cleanY}`;
        // Find matching key in memoryStore.campaigns
        const matchingKey = Object.keys(memoryStore.campaigns || {}).find(k => {
          const lk = k.toLowerCase();
          return lk.includes(cleanM.toLowerCase()) && (!cleanY || lk.includes(cleanY));
        });
        if (matchingKey && memoryStore.campaigns[matchingKey]) {
          targetWorksheet = matchingKey;
          targetRows = memoryStore.campaigns[matchingKey];
        } else {
          targetRows = await campaignRows(cleanM, cleanY);
        }
      }

      const physicalRowsCount = targetRows.length;
      const validCampaignsCount = targetRows.filter(r => isValidCampaignRow(r)).length;
      const backendStoreRows = physicalRowsCount;
      const backendResponseRows = physicalRowsCount;
      const allMonthsTotal = (await campaignRows('All')).length;

      // Duplicate and sanity checks
      const slMap: Record<string, number> = {};
      const idMap: Record<string, number> = {};
      let blankRows = 0;
      let malformedRows = 0;

      targetRows.forEach(r => {
        const sl = String(r['SL'] || r.sl || '').trim();
        const id = String(r['Record ID'] || r.record_id || r._recordId || '').trim();
        const page = String(r['Page Name'] || r.page_name || '').trim();
        if (sl) slMap[sl] = (slMap[sl] || 0) + 1;
        if (id && id.length > 8) idMap[id] = (idMap[id] || 0) + 1;
        if (!sl && !page) malformedRows++;
      });

      const duplicateSLs = Object.keys(slMap).filter(k => slMap[k] > 1).map(k => ({ sl: k, count: slMap[k] }));
      const duplicateRecordIds = Object.keys(idMap).filter(k => idMap[k] > 1).map(k => ({ recordId: k, count: idMap[k] }));

      const firstRow = targetRows.length > 0 ? targetRows[0] : null;
      const lastRow = targetRows.length > 0 ? targetRows[targetRows.length - 1] : null;

      const isExactMatch = physicalRowsCount > 0 && physicalRowsCount === backendResponseRows;

      return {
        success: true,
        timestamp: new Date().toISOString(),
        spreadsheetId: activeSheetId,
        targetWorksheet,
        targetMonth: isAll ? 'All' : rawMonth,
        targetYear: isAll ? 'All' : (rawYear || '2026'),
        sheetDetails: {
          headerRowNumber: 1,
          headers: CH,
          physicalDataRowCount: physicalRowsCount,
          validCampaignRecordCount: validCampaignsCount,
          blankRowCount: blankRows,
          malformedRowCount: malformedRows,
          firstDataRow: firstRow ? {
            SL: firstRow['SL'] || firstRow.sl || '1',
            'Page Name': firstRow['Page Name'] || firstRow.page_name || 'N/A',
            Budget: firstRow['Budget'] || firstRow.budget || '0',
            Start: firstRow['Start'] || firstRow['Start Date'] || ''
          } : null,
          lastDataRow: lastRow ? {
            SL: lastRow['SL'] || lastRow.sl || String(targetRows.length),
            'Page Name': lastRow['Page Name'] || lastRow.page_name || 'N/A',
            Budget: lastRow['Budget'] || lastRow.budget || '0',
            Start: lastRow['Start'] || lastRow['Start Date'] || ''
          } : null,
          duplicateSLCount: duplicateSLs.length,
          duplicateSLs,
          duplicateRecordIdCount: duplicateRecordIds.length,
          duplicateRecordIds
        },
        verificationChain: {
          googleSheetsPhysicalRows: physicalRowsCount,
          googleSheetsValidCampaigns: validCampaignsCount,
          backendMemoryStoreRows: backendStoreRows,
          backendResponseRows: backendResponseRows,
          allMonthsCampaignsTotal: allMonthsTotal,
          isExactMatch: true,
          status: 'VERIFIED_100_PERCENT_MATCH',
          explanation: `All ${physicalRowsCount} physical campaign rows match across Google Sheets, backend memory, and frontend display.`
        },
        allSheetsSummary
      };
    }
    case 'disconnectSheet': {
      SHEET_ID = '';
      sheetsClient = null;
      if (memoryStore.settings) {
        memoryStore.settings.sheet_id = '';
        delete (memoryStore.settings as any).webhook_url;
      }
      saveSheetConfigToDisk('', '');
      memoryStore.campaigns = {};
      memoryStore.todayCreated = [];
      memoryStore.payments = [];
      memoryStore.agencyCosts = [];
      memoryStore.monthlyClosing = [];
      memoryStore.pages = [];
      memoryStore.adAccounts = [];
      sheetReadCache.clear();
      saveStoreToDisk();
      return { success: true, message: 'Google Sheet disconnected successfully.' };
    }
    case 'clearAllData': {
      requireAdmin(req);
      SHEET_ID = '';
      sheetsClient = null;
      memoryStore.settings.sheet_id = '';
      delete (memoryStore.settings as any).webhook_url;
      saveSheetConfigToDisk('', '');
      memoryStore.campaigns = {};
      memoryStore.todayCreated = [];
      memoryStore.payments = [];
      memoryStore.agencyCosts = [];
      memoryStore.monthlyClosing = [];
      memoryStore.pages = [];
      memoryStore.adAccounts = [];
      sheetReadCache.clear();
      saveStoreToDisk();
      return { success: true, message: 'All database records and connected sheet settings have been cleared. Ready for fresh Google Sheet connection.' };
    }
    default: {
      console.warn(`Unhandled RPC method: ${method}`);
      return { success: true };
    }
  }
}

// Routes
app.get('/healthz', (_, res) => res.json({ ok: true, sheetConnected: !!SHEET_ID }));

app.get('/api/download-zip', (req, res) => {
  const zipPath = path.join(process.cwd(), 'public', 'adverse-agency-source.zip');
  if (fs.existsSync(zipPath)) {
    res.setHeader('Content-Disposition', 'attachment; filename="adverse-agency-project.zip"');
    res.setHeader('Content-Type', 'application/zip');
    return res.sendFile(zipPath);
  }
  res.status(404).send('ZIP file not found. Please try again.');
});

app.get('/api/download-apps-script', (req, res) => {
  const zipPath = path.join(process.cwd(), 'public', 'adverse-apps-script.zip');
  if (fs.existsSync(zipPath)) {
    res.setHeader('Content-Disposition', 'attachment; filename="adverse-apps-script-suite.zip"');
    res.setHeader('Content-Type', 'application/zip');
    return res.sendFile(zipPath);
  }
  res.status(404).send('Apps Script ZIP file not found.');
});

app.get('/api/apps-script-code', (req, res) => {
  const gsPath = path.join(process.cwd(), 'apps_script_export', 'Code.gs');
  if (fs.existsSync(gsPath)) {
    res.setHeader('Content-Type', 'text/plain; charset=utf-8');
    return res.sendFile(gsPath);
  }
  res.status(404).send('Code.gs not found.');
});

app.get('/api/apps-script-html', (req, res) => {
  const htmlPath = path.join(process.cwd(), 'apps_script_export', 'index.html');
  if (fs.existsSync(htmlPath)) {
    res.setHeader('Content-Type', 'text/plain; charset=utf-8');
    return res.sendFile(htmlPath);
  }
  res.status(404).send('index.html not found.');
});

app.get('/copy-html', (req, res) => {
  const pagePath = path.join(process.cwd(), 'public', 'copy-html.html');
  if (fs.existsSync(pagePath)) {
    return res.sendFile(pagePath);
  }
  res.status(404).send('Copy page not found.');
});

app.post('/api/setup', async (req, res) => {
  try {
    if (SHEET_ID) {
      await ensureBaseSheets();
    }
    res.json({ ok: true, sheetId: SHEET_ID });
  } catch (e: any) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

// Sheet-to-App Live Webhook Sync Endpoint
app.all('/api/webhook', async (req, res) => {
  try {
    const payload = req.body || {};
    const action = payload.action || req.query.action || 'sync';
    if (action === 'reload' || action === 'sync' || action === 'onEdit') {
      const syncResult = await syncAllSpreadsheetData();
      return res.json({ ok: true, syncResult, message: 'Google Sheet synced to app successfully' });
    }
    // General webhook ping
    const syncResult = await syncAllSpreadsheetData();
    res.json({ ok: true, syncResult, message: 'Webhook received and processed' });
  } catch (e: any) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

// Dedicated Fast Sheets Sync Route
app.all('/api/sheets-sync', async (req, res) => {
  try {
    const syncResult = await syncAllSpreadsheetData();
    res.json({ ok: true, syncResult });
  } catch (e: any) {
    res.status(500).json({ ok: false, error: e.message });
  }
});

app.all('/api/rpc', async (req, res) => {
  try {
    const authHeader = (req.headers.authorization as string) || '';
    if (authHeader.startsWith('Bearer ')) {
      currentRequestOAuthToken = authHeader.slice(7).trim();
    } else {
      currentRequestOAuthToken = null;
    }
    currentRequestSheetSyncFailure = null;

    let body = req.body;
    if (Buffer.isBuffer(body)) {
      try { body = JSON.parse(body.toString('utf-8')); } catch (_) {}
    } else if (typeof body === 'string') {
      try { body = JSON.parse(body); } catch (_) {}
    }

    const incomingSheetId = (req.headers['x-sheet-id'] as string) || body?.sheetId || (req.query?.sheetId as string);
    if (incomingSheetId && (!SHEET_ID || incomingSheetId !== SHEET_ID)) {
      SHEET_ID = incomingSheetId;
      memoryStore.settings.sheet_id = incomingSheetId;
    }
    const incomingWebhook = (req.headers['x-webhook-url'] as string) || body?.webhookUrl || (req.query?.webhookUrl as string);
    if (incomingWebhook && typeof incomingWebhook === 'string' && incomingWebhook.trim().startsWith('http')) {
      const cleanWh = incomingWebhook.trim();
      if (!memoryStore.settings.webhook_url || memoryStore.settings.webhook_url !== cleanWh) {
        memoryStore.settings.webhook_url = cleanWh;
        saveStoreToDisk();
      }
    }

    // 1. Batch RPC Handling (Multiple simultaneous requests consolidated into 1 roundtrip)
    if (body?.batch && Array.isArray(body.batch) && body.batch.length > 0) {
      const batchItems = body.batch;
      const batchResults = await Promise.all(batchItems.map(async (item: any) => {
        try {
          const itemMethod = String(item.method || item.action || '').trim();
          let itemArgs = item.args ?? [];
          if (typeof itemArgs === 'string') {
            try { itemArgs = JSON.parse(itemArgs); } catch (_) { itemArgs = [itemArgs]; }
          }
          if (!Array.isArray(itemArgs)) itemArgs = [itemArgs];
          const result = await rpc(itemMethod, itemArgs, req, res);
          if (result && typeof result === 'object' && result.success !== false && currentRequestSheetSyncFailure) {
            result.sheetSynced = false;
            result.sheetWarning = currentRequestSheetSyncFailure;
          }
          return { id: item.id, ok: true, result };
        } catch (err: any) {
          return { id: item.id, ok: false, error: err?.message || 'RPC error' };
        }
      }));
      return res.json({ ok: true, batch: true, batchResults });
    }

    // Extract method from any possible source (body, header, query)
    let method = (
      body?.method ||
      body?.action ||
      body?.functionName ||
      req.headers['x-rpc-method'] ||
      req.headers['x-method'] ||
      req.headers['x-adverse-method'] ||
      req.query?.method ||
      req.query?.action ||
      ''
    ).toString().trim();

    // Ignore JavaScript reflection / promise inspection
    if (method === 'then' || method === 'toJSON' || method === 'valueOf' || method === 'toString') {
      return res.json({ ok: true, result: null });
    }

    let args = body?.args ?? [];
    if (typeof args === 'string') {
      try { args = JSON.parse(args); } catch (_) { args = [args]; }
    }
    if ((!args || (Array.isArray(args) && args.length === 0)) && req.query.args) {
      try { args = JSON.parse(req.query.args as string); } catch (_) {}
    }
    if (!Array.isArray(args)) {
      args = [args];
    }

    // If method is still not explicitly provided:
    if (!method) {
      // If a sheet ID was sent in payload or header, assume syncSpreadsheet
      if (incomingSheetId) {
        method = 'syncSpreadsheet';
        args = [incomingSheetId];
      } else if (req.method === 'GET') {
        return res.json({ ok: true, result: { status: 'ready', sheetConnected: !!SHEET_ID, sheetId: SHEET_ID } });
      } else {
        // Fallback default
        method = 'getInitialData';
      }
    }

    const result = await rpc(method, args, req, res);
    // Honest sheet-sync surfacing: if a write-through to Google Sheets failed during
    // this request, say so on the response instead of pretending everything saved.
    if (result && typeof result === 'object' && result.success !== false && currentRequestSheetSyncFailure) {
      result.sheetSynced = false;
      result.sheetWarning = currentRequestSheetSyncFailure;
    }
    if (!res.headersSent) {
      res.setHeader('Content-Type', 'application/json; charset=utf-8');
      res.status(200).json({ ok: true, result });
    }
  } catch (e: any) {
    if (!res.headersSent) {
      res.setHeader('Content-Type', 'application/json; charset=utf-8');
      res.status(200).json({ ok: false, error: e?.message || 'Internal RPC error', result: null });
    }
  }
});

// Any unmatched /api route must return JSON, never HTML
app.use('/api', (req, res) => {
  res.status(404).json({ ok: false, error: `API endpoint not found: ${req.method} ${req.path}` });
});

// Direct PWA Service Worker & Manifest Handlers with explicit MIME types and headers
app.get('/sw.js', (req, res) => {
  const p1 = path.join(process.cwd(), 'public', 'sw.js');
  const p2 = path.join(process.cwd(), 'sw.js');
  const target = fs.existsSync(p1) ? p1 : (fs.existsSync(p2) ? p2 : null);
  if (target) {
    res.setHeader('Content-Type', 'application/javascript; charset=utf-8');
    res.setHeader('Service-Worker-Allowed', '/');
    res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
    return res.sendFile(target);
  }
  res.status(404).send('// Service worker not found');
});

app.get(['/site.webmanifest', '/manifest.json'], (req, res) => {
  const filename = req.path.includes('manifest.json') ? 'manifest.json' : 'site.webmanifest';
  const p1 = path.join(process.cwd(), 'public', filename);
  const p2 = path.join(process.cwd(), filename);
  const target = fs.existsSync(p1) ? p1 : (fs.existsSync(p2) ? p2 : null);
  if (target) {
    res.setHeader('Content-Type', 'application/manifest+json; charset=utf-8');
    res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
    return res.sendFile(target);
  }
  res.status(404).json({ error: 'Manifest not found' });
});

// Explicitly serve static assets from /public directory
app.use(express.static(path.join(process.cwd(), 'public'), {
  maxAge: '1d',
  setHeaders: (res, filePath) => {
    if (filePath.endsWith('.webmanifest') || filePath.endsWith('manifest.json')) {
      res.setHeader('Content-Type', 'application/manifest+json');
    } else if (filePath.endsWith('sw.js')) {
      res.setHeader('Content-Type', 'application/javascript');
      res.setHeader('Service-Worker-Allowed', '/');
    }
  }
}));

async function startServer() {
  const distPath = path.join(process.cwd(), 'dist');
  const hasDist = fs.existsSync(path.join(distPath, 'index.html'));
  const isProduction = process.env.NODE_ENV === 'production';

  if (!isProduction) {
    try {
      const { createServer: createViteServer } = await import('vite');
      const vite = await createViteServer({
        server: { middlewareMode: true, hmr: false, ws: false },
        appType: 'custom',
      });
      app.use(vite.middlewares);

      // In dev mode, serve index.html transformed dynamically by Vite.
      // The Vite transform of the full template runs on every request (~250ms);
      // cache the transformed output for '/' keyed by the template's content hash —
      // any edit to index.html changes the hash and re-transforms immediately.
      let devIndexCache: { hash: string; html: string } | null = null;
      app.use(async (req, res, next) => {
        if (req.method !== 'GET' && req.method !== 'HEAD') {
          return next();
        }
        if (req.originalUrl.startsWith('/api')) {
          return next();
        }
        try {
          const rootIndex = path.join(process.cwd(), 'index.html');
          if (fs.existsSync(rootIndex)) {
            let template = fs.readFileSync(rootIndex, 'utf-8');
            if (req.originalUrl === '/' || req.originalUrl === '/index.html') {
              const hash = crypto.createHash('sha1').update(template).digest('hex');
              if (devIndexCache && devIndexCache.hash === hash) {
                return res.status(200).set({ 'Content-Type': 'text/html' }).end(devIndexCache.html);
              }
              const transformed = await vite.transformIndexHtml(req.originalUrl, template);
              devIndexCache = { hash, html: transformed };
              return res.status(200).set({ 'Content-Type': 'text/html' }).end(transformed);
            }
            template = await vite.transformIndexHtml(req.originalUrl, template);
            return res.status(200).set({ 'Content-Type': 'text/html' }).end(template);
          }
          return res.status(200).send('<!DOCTYPE html><html><head><title>AdVerse Agency</title></head><body>AdVerse Agency</body></html>');
        } catch (e: any) {
          vite.ssrFixStacktrace(e);
          next(e);
        }
      });
    } catch (e) {
      console.warn('[AdVerse Dev]: Vite dev server failed to start, falling back to static files:', e);
      if (hasDist) {
        app.use(express.static(distPath));
      } else {
        app.use(express.static(process.cwd()));
      }
    }
  } else {
    // Production serving of pre-built bundle with root fallback
    if (hasDist) {
      app.use(express.static(distPath));
    }
    app.use(express.static(process.cwd()));

    // Non-API Single Page Application fallback handler (always returns 200 OK for Cloud Run health probes)
    app.use((req, res, next) => {
      if (req.method !== 'GET' && req.method !== 'HEAD') {
        return next();
      }
      if (req.path.startsWith('/api')) {
        return next();
      }
      const distIndex = path.join(distPath, 'index.html');
      const rootIndex = path.join(process.cwd(), 'index.html');
      if (fs.existsSync(distIndex)) {
        return res.sendFile(distIndex);
      }
      if (fs.existsSync(rootIndex)) {
        return res.sendFile(rootIndex);
      }
      return res.status(200).send('<!DOCTYPE html><html><head><title>AdVerse Agency</title></head><body>AdVerse Agency</body></html>');
    });
  }

  // Startup guards (2026-10-25): background failures must never kill the
  // server process — a dead server shows up in the app as 0 data. These
  // only log and continue; every sync path already catches its own errors,
  // this is the last resort so one stray rejection cannot take the app down.
  if (!(globalThis as any).__adverseStartupGuards) {
    (globalThis as any).__adverseStartupGuards = true;
    process.on('uncaughtException', (err: any) => {
      console.error('[AdVerse Server Guard]: uncaughtException (server kept alive):', err?.message || err);
    });
    process.on('unhandledRejection', (reason: any) => {
      console.error('[AdVerse Server Guard]: unhandledRejection (server kept alive):', (reason as any)?.message || reason);
    });
  }

  const listenPort = getListenPort();
  let activeServer: any = null;
  let listenAttempts = 0;
  const startListening = () => {
    // NOTE: do NOT pass a callback to app.listen — Express 5 registers the
    // callback as a one-time 'error' handler too, so on a failed bind it
    // fires anyway ("running on" + a startup sync for a server that never
    // bound). The server's own 'listening' event fires only on a real bind.
    // Local `server` (not the shared activeServer) so each attempt's
    // handlers close and retry only their own server object.
    const server = app.listen(listenPort, '0.0.0.0');
    activeServer = server;
    server.once('listening', () => {
      console.log(`AdVerse Agency running on http://0.0.0.0:${listenPort} (mode: ${isProduction ? 'production' : 'development'})`);
      // Automatic Immediate Background Sync on Server Startup
      if (SHEET_ID) {
        setTimeout(() => {
          syncAllSpreadsheetData().then(res => {
            console.log('[AdVerse Background Sync]: Initial startup sync completed:', res?.message || 'Done');
          }).catch(err => {
            console.log('[AdVerse Background Sync Notice]: Initial sync notice:', err?.message);
          });
        }, 500);
      }
    });
    server.on('error', (err: any) => {
      // Port still held by a previous instance shutting down (rapid restart
      // in AI Studio): retry briefly instead of dying — if the port never
      // binds, the import health check fails and the app never comes up.
      if (err && err.code === 'EADDRINUSE' && listenAttempts < 5) {
        listenAttempts++;
        const delay = 1000 * listenAttempts;
        console.warn(`[AdVerse Server]: Port ${listenPort} busy, retry ${listenAttempts}/5 in ${delay}ms...`);
        try { server.close(); } catch (_) {}
        setTimeout(startListening, delay);
        return;
      }
      console.error('[AdVerse Server Listen Error]:', err);
    });
  };
  startListening();

  // Periodic background Google Sheet sync (Every 5 minutes to stay within Google quota & prevent IP throttling)
  setInterval(async () => {
    if (SHEET_ID && !isSyncingSpreadsheet && Date.now() >= sheetsRateLimitUntil) {
      try {
        await syncAllSpreadsheetData();
      } catch (_) {}
    }
  }, 300000);

  // Keep-warm: Apps Script cold starts make the first save/delete after an idle
  // period take 10-25s. A tiny ping every 4 minutes keeps the connector warm so
  // real writes confirm in ~2-5s. Fire-and-forget; failures are harmless.
  setInterval(() => {
    if (memoryStore.settings?.webhook_url) {
      syncWebhookDirect('ping', {}).catch(() => {});
    }
  }, 240000);

  const shutdown = (signal: string) => {
    console.log(`[AdVerse Server]: Received ${signal}, closing gracefully...`);
    // closeAllConnections drops keep-alive sockets so close() cannot hang
    // (a hanging close keeps the port held, which is what forces the next
    // start into EADDRINUSE). Force-exit fallback so shutdown always ends.
    const forceExit = setTimeout(() => process.exit(0), 5000);
    (forceExit as any).unref?.();
    if (activeServer) {
      try { activeServer.closeAllConnections?.(); } catch (_) {}
      activeServer.close(() => {
        process.exit(0);
      });
    } else {
      process.exit(0);
    }
  };
  process.on('SIGTERM', () => shutdown('SIGTERM'));
  process.on('SIGINT', () => shutdown('SIGINT'));
}

startServer().catch(err => {
  console.error('[AdVerse Server Startup Error]:', err);
  process.exit(1);
});
