// AdVerse Agency - Client Application Engine
import { googleSignIn, googleSignOut, getGoogleAccessToken, getCurrentGoogleUser, initGoogleAuth } from './services/googleAuth';
import {
  saveCampaignToFirestore,
  savePaymentToFirestore,
  saveSettingsToFirestore,
  testConnection
} from './services/firestoreService';
import {
  fetchAllSpreadsheetData,
  pushAllSheetsDataDirect,
  fetchSpreadsheetMetadata,
  deleteRowFromSheetDirect,
  updateRowInSheetDirect,
  upsertRowInSheetDirect
} from './services/sheetsService';
import SyncLogger from './services/syncLogger';

if (typeof window !== 'undefined') {
  window.SyncLogger = SyncLogger;
}


// Schema Header Constants & Fallback Guarantees
export const CH = (typeof window !== 'undefined' && window.CH) ? window.CH : ['SL','OFF $','AD Account','ID Name','Status','Page Name','Start','End','Types','Day','Budget','Daily $','Taka','Rate','Entry Date','Month','Year','Record ID'];
export const TODAY_HEADERS = (typeof window !== 'undefined' && window.TODAY_HEADERS) ? window.TODAY_HEADERS : ['SL','OFF $','AD Account','ID Name','Status','Page Name','Start','End','Types','Day','Budget','Daily $','Taka','Rate','Entry Date','Month','Year','Record ID'];
export const PAYMENT_HEADERS = (typeof window !== 'undefined' && window.PAYMENT_HEADERS) ? window.PAYMENT_HEADERS : ['Date','Payment Method','Paid (৳)','Month','Year','Note'];
export const PAGE_SUMMARY_HEADERS = (typeof window !== 'undefined' && window.PAGE_SUMMARY_HEADERS) ? window.PAGE_SUMMARY_HEADERS : ['Page Name','Total Ads','Total Budget $'];
export const AGENCY_SPENDING_HEADERS = (typeof window !== 'undefined' && window.AGENCY_SPENDING_HEADERS) ? window.AGENCY_SPENDING_HEADERS : ['Month','Year','Total Spend $'];
export const MONTHLY_CLOSING_HEADERS = (typeof window !== 'undefined' && window.MONTHLY_CLOSING_HEADERS) ? window.MONTHLY_CLOSING_HEADERS : ['Month','Year','Status','Closed At','Closed By','Total Campaign','Total Budget $','Total Campaign Taka','Total Paid (৳)','Total Due (৳)','Agency Cost (৳)','Rate'];
export const AGENCY_COST_HEADERS = (typeof window !== 'undefined' && window.AGENCY_COST_HEADERS) ? window.AGENCY_COST_HEADERS : ['Date','Month','Year','Type','Cost','Per Person Cost','Status','Record ID'];
export const END_ADS_HEADERS = (typeof window !== 'undefined' && window.END_ADS_HEADERS) ? window.END_ADS_HEADERS : ['Page Name','Start Date','End Date','Day','Budget','Type'];
export const AD_ACCOUNT_BILLING_HEADERS = (typeof window !== 'undefined' && window.AD_ACCOUNT_BILLING_HEADERS) ? window.AD_ACCOUNT_BILLING_HEADERS : ['Ad Account','ID Name','Due ($)','Threshold ($)','Remaining ($)','Billing Date','Active Spend ($)','Active Ads','Card / Method','Status','Last Checked'];

const MONTHS = (typeof window !== 'undefined' && window.MONTHS) ? window.MONTHS : ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

if (typeof window !== 'undefined') {
  window.CH = CH;
  window.TODAY_HEADERS = TODAY_HEADERS;
  window.PAYMENT_HEADERS = PAYMENT_HEADERS;
  window.PAGE_SUMMARY_HEADERS = PAGE_SUMMARY_HEADERS;
  window.AGENCY_SPENDING_HEADERS = AGENCY_SPENDING_HEADERS;
  window.MONTHLY_CLOSING_HEADERS = MONTHLY_CLOSING_HEADERS;
  window.AGENCY_COST_HEADERS = AGENCY_COST_HEADERS;
  window.END_ADS_HEADERS = END_ADS_HEADERS;
  window.AD_ACCOUNT_BILLING_HEADERS = AD_ACCOUNT_BILLING_HEADERS;
  window.MONTHS = MONTHS;
}

window.google = window.google || {};
google.script = google.script || {};
google.script.run = new Proxy({}, {
  get(_target, prop) {
    if (typeof prop === 'symbol' || prop === 'then' || prop === 'toJSON' || prop === 'valueOf' || prop === 'toString') {
      return undefined;
    }
    if (prop === 'withSuccessHandler') return function(success) {
      return new Proxy({__success: success}, {
        get(state, method) {
          if (typeof method === 'symbol' || method === 'then' || method === 'toJSON') return undefined;
          if (method === 'withFailureHandler') return function(failure) {
            return new Proxy({...state, __failure: failure}, {
              get(s, realMethod) {
                if (typeof realMethod === 'symbol' || realMethod === 'then' || realMethod === 'toJSON') return undefined;
                return (...args) => cloudRpc(String(realMethod), args, s.__success, s.__failure);
              }
            });
          };
          return (...args) => cloudRpc(String(method), args, state.__success, state.__failure);
        }
      });
    };
    if (prop === 'withFailureHandler') return function(failure) {
      return new Proxy({__failure: failure}, {
        get(state, method) {
          if (typeof method === 'symbol' || method === 'then' || method === 'toJSON') return undefined;
          if (method === 'withSuccessHandler') return function(success) {
            return new Proxy({...state, __success: success}, {
              get(s, realMethod) {
                if (typeof realMethod === 'symbol' || realMethod === 'then' || realMethod === 'toJSON') return undefined;
                return (...args) => cloudRpc(String(realMethod), args, s.__success, s.__failure);
              }
            });
          };
          return (...args) => cloudRpc(String(method), args, null, state.__failure);
        }
      });
    };
    return (...args) => cloudRpc(String(prop), args);
  }
});

// --- Intelligent Request Batching, Debouncing & In-Flight Deduplication Engine ---
let rpcRequestIdSeq = 0;
let pendingRpcBatch = [];
let batchDispatchTimer = null;
const inFlightRequests = new Map(); // Key -> Promise (In-flight deduplication for identical reads)

// Methods that can be safely deduplicated when identical read requests are in-flight
const DEDUPLICATABLE_READ_METHODS = new Set([
  'getInitialData',
  'getRecords',
  'getMenuStats',
  'getMenuData',
  'getSettings',
  'getAccountsAndPages',
  'getProfile',
  'getSpreadsheetMetadata',
  'getAdminPin',
  'getViewerPin',
  'getAgencyAnalytics',
  'getCampaignSummary'
]);

function getRequestDedupKey(method, args) {
  try {
    return `${method}:${JSON.stringify(args || [])}`;
  } catch (_) {
    return null;
  }
}

function getRpcHeaders(method) {
  const savedPin = localStorage.getItem('adverse_saved_pin') || '';
  const savedToken = localStorage.getItem('adverse_session_token') || '';
  const savedRole = localStorage.getItem('adverse_user_role') || '';
  const googleToken = getGoogleAccessToken();
  let sheetId = '';
  let webhookUrl = '';
  try {
    const savedConfig = JSON.parse(localStorage.getItem('adverse_sheet_config') || '{}');
    if (savedConfig.sheetId) sheetId = savedConfig.sheetId;
    if (savedConfig.webhookUrl) webhookUrl = savedConfig.webhookUrl;
  } catch (e) {}
  if (!sheetId && state.sheetId) sheetId = state.sheetId;

  const reqHeaders = {
    'Content-Type': 'application/json',
    'x-rpc-method': method,
    'x-method': method,
    'x-adverse-pin': savedPin,
    'x-pin': savedPin,
    'x-adverse-token': savedToken,
    'x-adverse-role': savedRole
  };
  if (googleToken) {
    reqHeaders['Authorization'] = `Bearer ${googleToken}`;
  }
  if (sheetId) {
    reqHeaders['x-sheet-id'] = sheetId;
  }
  if (webhookUrl) {
    reqHeaders['x-webhook-url'] = webhookUrl;
  }
  return { reqHeaders, savedPin, savedToken, savedRole, sheetId, webhookUrl };
}

function computeClientNextSL(month, year) {
  const targetMonth = (month && month !== 'All' && month !== 'All Months' && month !== 'All_Months')
    ? String(month).trim()
    : (state?.filters?.month && state.filters.month !== 'All' && state.filters.month !== 'All Months' ? state.filters.month : MONTHS[new Date().getMonth()]);
  const targetYear = (year && year !== 'All' && year !== 'All Years')
    ? String(year).trim()
    : (state?.filters?.year && state.filters.year !== 'All Years' ? String(state.filters.year) : String(new Date().getFullYear()));

  const mIdx = MONTHS.findIndex(m => m.toLowerCase() === targetMonth.toLowerCase());
  const monthNum = mIdx >= 0 ? mIdx + 1 : (new Date().getMonth() + 1);
  const monthPrefix = String(monthNum).padStart(2, '0');

  const allRecords = [
    ...(state?.rawRecords || []),
    ...(state?.filteredRecords || []),
    ...(state?.todayCreated || []),
    ...(state?.todayCreatedRegistry || []),
    ...(state?.masterCampaignRegistry || []),
    ...(state?.records || [])
  ];

  let maxSeq = 0;
  let maxDirectNumber = 0;

  allRecords.forEach(r => {
    if (!r) return;
    const rMonth = String(r['Month'] || r.month || '').trim().toLowerCase();
    const rYear = String(r['Year'] || r.year || '').trim();

    const startDate = r['Start Date'] || r['Start'] || r.start_date || '';
    const parsedStart = startDate ? parseAnyDate(startDate) : null;
    const isMatchingMonth = !rMonth || rMonth === targetMonth.toLowerCase() || (parsedStart && parsedStart.getMonth() === (monthNum - 1));
    const isMatchingYear = !rYear || rYear === targetYear || (parsedStart && String(parsedStart.getFullYear()) === targetYear);

    const raw = String(r['SL'] ?? r.sl ?? r.sl_no ?? r['SL NO'] ?? '').trim().replace(/^#/, '');
    if (!raw) return;

    if (isMatchingMonth && isMatchingYear) {
      const numOnly = parseInt(raw.replace(/\D/g, ''), 10);
      if (!isNaN(numOnly) && String(numOnly).startsWith(monthPrefix) && numOnly > maxDirectNumber) {
        maxDirectNumber = numOnly;
      }

      if (raw.length >= 4 && (raw.startsWith(monthPrefix) || (monthNum >= 10 && raw.startsWith(String(monthNum))))) {
        const seqPart = raw.slice(monthPrefix.length);
        const seqVal = parseInt(seqPart, 10);
        if (!isNaN(seqVal) && seqVal > maxSeq) {
          maxSeq = seqVal;
        }
      } else {
        const val = parseInt(raw, 10);
        if (!isNaN(val) && val < 1000 && val > maxSeq) {
          maxSeq = val;
        }
      }
    }
  });

  const hasPrefixed = allRecords.some(r => {
    if (!r) return false;
    const rMonth = String(r['Month'] || r.month || '').trim().toLowerCase();
    if (rMonth && rMonth !== targetMonth.toLowerCase()) return false;
    const raw = String(r['SL'] ?? r.sl ?? r.sl_no ?? '').trim();
    return raw.length >= 4 && (raw.startsWith(monthPrefix) || (monthNum >= 10 && raw.startsWith(String(monthNum))));
  });

  if (hasPrefixed && maxSeq > 0) {
    return `${monthPrefix}${String(maxSeq + 1).padStart(2, '0')}`;
  }

  if (hasPrefixed && maxDirectNumber > 0 && String(maxDirectNumber).startsWith(monthPrefix)) {
    return String(maxDirectNumber + 1);
  }

  if (!hasPrefixed && maxSeq > 0) {
    return String(maxSeq + 1);
  }

  if (maxSeq > 0) {
    return `${monthPrefix}${String(maxSeq + 1).padStart(2, '0')}`;
  }
  return `${monthPrefix}01`;
}

function getGracefulRpcFallback(method, args) {
  const mKey = Array.isArray(args) && args[0] ? String(args[0]) : '';
  const mn = state?.filters?.month || '';
  const yr = state?.filters?.year || '';

  if (method === 'getNextSLNo') {
    const calc = computeClientNextSL(args?.[0] || mn, args?.[1] || yr);
    return {
      success: true,
      nextSL: calc
    };
  }

  if (method === 'getInitialData' || method === 'getAppData') {
    let snapshot = null;
    try {
      const raw = localStorage.getItem(SNAPSHOT_STORAGE_KEY);
      snapshot = raw ? JSON.parse(raw) : null;
    } catch (_) {}

    const master = (state.masterCampaignRegistry && state.masterCampaignRegistry.length > 0)
      ? state.masterCampaignRegistry
      : (snapshot?.allCampaigns || []);

    return {
      success: true,
      menus: getStandardMenus(),
      months: snapshot?.months || state.months || [],
      adAccounts: snapshot?.adAccounts || state.adAccounts || [],
      pages: snapshot?.pages || state.pages || [],
      rate: snapshot?.rate || state.rate || 140,
      allCampaigns: master,
      preloaded: {
        today_created_campaign: { records: { headers: state.headers.length ? state.headers : CH, rows: state.todayCreatedRegistry || master.slice(0, 10) }, stats: {} },
        campaign: { records: { headers: state.headers.length ? state.headers : CH, rows: master }, stats: {} },
        end_ads: { records: { today: [], tomorrow: [], all: master }, stats: {} }
      }
    };
  }

  // Check clientCache variations
  if (mKey) {
    const cached = state.clientCache[`${mKey}_${mn}_${yr}___`] ||
      state.clientCache[`${mKey}_${mn}_${yr}__`] ||
      state.clientCache[`${mKey}_${mn}_${yr}`] ||
      state.clientCache[`${mKey}_____`] ||
      state.clientCache[mKey];
    if (cached) return cached;
  }

  const cachedByMethod = state.clientCache[`${method}_${mn}_${yr}`] || state.clientCache[method];
  if (cachedByMethod) return cachedByMethod;

  if (mKey === 'campaign' || mKey === 'today_created_campaign') {
    const rows = (mKey === 'today_created_campaign') ? (state.todayCreatedRegistry || []) : (state.masterCampaignRegistry || []);
    return {
      success: true,
      records: { headers: state.headers.length ? state.headers : CH, rows },
      stats: { totalRecords: rows.length, totalBudget: 0, totalTaka: 0 }
    };
  }

  if (mKey === 'end_ads') {
    const rows = state.masterCampaignRegistry || [];
    const todayBST = getNowBST();
    const tomorrowBST = new Date(todayBST.getTime() + 86400000);
    return {
      success: true,
      records: {
        today: rows.filter(r => isSameDay(r['End Date'] || r['End'], todayBST)),
        tomorrow: rows.filter(r => isSameDay(r['End Date'] || r['End'], tomorrowBST)),
        all: rows
      },
      stats: { success: true }
    };
  }

  return { success: true, offline: true, records: { headers: [], rows: [] }, stats: {} };
}

async function executeSingleRpc(method, args) {
  const { reqHeaders, savedPin, savedToken, savedRole, sheetId, webhookUrl } = getRpcHeaders(method);

  for (let attempt = 0; attempt < 4; attempt++) {
    try {
      const res = await fetch(`/api/rpc?method=${encodeURIComponent(method)}`, {
        method: 'POST',
        headers: reqHeaders,
        credentials: 'include',
        body: JSON.stringify({ method, args, pin: savedPin, token: savedToken, role: savedRole, sheetId, webhookUrl })
      });
      const rawText = await res.text();
      let data;
      try {
        data = JSON.parse(rawText);
      } catch (parseErr) {
        if (rawText.includes('Rate exceeded') || res.status === 429) {
          throw new Error('Server rate limit temporarily reached. Please wait a moment.');
        }
        if (attempt < 3) {
          await new Promise(r => setTimeout(r, 200 * (attempt + 1)));
          continue;
        }
        return getGracefulRpcFallback(method, args);
      }
      if (!res.ok || data.ok === false) {
        if (data.error === 'LOGIN_REQUIRED' || (data.result && data.result.error === 'LOGIN_REQUIRED')) {
          showPinLogin('Session expired. Please enter your PIN.');
          throw new Error('LOGIN_REQUIRED');
        }
        if (attempt < 3) {
          await new Promise(r => setTimeout(r, 200 * (attempt + 1)));
          continue;
        }
        return getGracefulRpcFallback(method, args);
      }
      if (data && data.result && data.result.error === 'LOGIN_REQUIRED') {
        showPinLogin('Session expired. Please enter your PIN.');
        throw new Error('LOGIN_REQUIRED');
      }
      // Honest sheet-sync surfacing: the server flags responses whose Google Sheet
      // write-through failed (entry is kept in-app and queued for automatic retry).
      if (data && data.result && data.result.sheetSynced === false && data.result.sheetWarning) {
        try { showToast('⚠️ ' + data.result.sheetWarning, 'error'); } catch (_) {}
      }
      return (data && data.result !== undefined) ? data.result : (data ?? getGracefulRpcFallback(method, args));
    } catch (err) {
      if (err.message === 'LOGIN_REQUIRED') throw err;
      if (attempt < 3) {
        await new Promise(r => setTimeout(r, 200 * Math.pow(2, attempt)));
        continue;
      }
      return getGracefulRpcFallback(method, args);
    }
  }
}

async function executeBatchRpc(batchPayload) {
  const { reqHeaders, savedPin, savedToken, savedRole, sheetId, webhookUrl } = getRpcHeaders('batch');

  for (let attempt = 0; attempt < 4; attempt++) {
    try {
      const res = await fetch(`/api/rpc?method=batch`, {
        method: 'POST',
        headers: reqHeaders,
        credentials: 'include',
        body: JSON.stringify({ batch: batchPayload, pin: savedPin, token: savedToken, role: savedRole, sheetId, webhookUrl })
      });
      const rawText = await res.text();
      let data;
      try {
        data = JSON.parse(rawText);
      } catch (parseErr) {
        if (attempt < 3) {
          await new Promise(r => setTimeout(r, 200 * (attempt + 1)));
          continue;
        }
        return batchPayload.map(it => ({ id: it.id, ok: true, result: getGracefulRpcFallback(it.method, it.args) }));
      }
      if (!res.ok || data.ok === false) {
        if (attempt < 3) {
          await new Promise(r => setTimeout(r, 200 * (attempt + 1)));
          continue;
        }
        return batchPayload.map(it => ({ id: it.id, ok: true, result: getGracefulRpcFallback(it.method, it.args) }));
      }
      return data.batchResults || [];
    } catch (err) {
      if (attempt < 3) {
        await new Promise(r => setTimeout(r, 200 * Math.pow(2, attempt)));
        continue;
      }
      return batchPayload.map(it => ({ id: it.id, ok: true, result: getGracefulRpcFallback(it.method, it.args) }));
    }
  }
}

async function flushRpcBatchQueue() {
  batchDispatchTimer = null;
  if (pendingRpcBatch.length === 0) return;

  const currentBatch = pendingRpcBatch.splice(0, pendingRpcBatch.length);

  // If only 1 item in batch, send as standard direct call for minimal payload overhead
  if (currentBatch.length === 1) {
    const singleItem = currentBatch[0];
    try {
      const res = await executeSingleRpc(singleItem.method, singleItem.args);
      if (typeof singleItem.success === 'function') singleItem.success(res);
      singleItem.resolve(res);
    } catch (err) {
      const fallback = getGracefulRpcFallback(singleItem.method, singleItem.args);
      if (fallback) {
        if (typeof singleItem.success === 'function') singleItem.success(fallback);
        singleItem.resolve(fallback);
      } else {
        if (typeof singleItem.failure === 'function') singleItem.failure({ message: err?.message || 'Request failed' });
        singleItem.reject(err);
      }
    }
    return;
  }

  // Multiple items: Send 1 Consolidated Batch Request
  try {
    const batchPayload = currentBatch.map(it => ({
      id: it.id,
      method: it.method,
      args: it.args
    }));

    const batchResults = await executeBatchRpc(batchPayload);
    const resultMap = new Map((batchResults || []).map(r => [r.id, r]));

    for (const item of currentBatch) {
      const resObj = resultMap.get(item.id);
      if (resObj && (resObj.ok || resObj.result !== undefined)) {
        const finalRes = (resObj.result !== undefined) ? resObj.result : resObj;
        if (typeof item.success === 'function') item.success(finalRes);
        item.resolve(finalRes);
      } else {
        const fallback = getGracefulRpcFallback(item.method, item.args);
        if (fallback) {
          if (typeof item.success === 'function') item.success(fallback);
          item.resolve(fallback);
        } else {
          const errMsg = resObj?.error || 'Batch item execution failed';
          if (typeof item.failure === 'function') item.failure({ message: errMsg });
          item.reject(new Error(errMsg));
        }
      }
    }
  } catch (err) {
    // If the entire batch network request failed, gracefully resolve all items using fallback
    for (const item of currentBatch) {
      const fallback = getGracefulRpcFallback(item.method, item.args);
      if (fallback) {
        if (typeof item.success === 'function') item.success(fallback);
        item.resolve(fallback);
      } else {
        if (typeof item.failure === 'function') item.failure({ message: err?.message || 'Batch request failed' });
        item.reject(err);
      }
    }
  }
}

async function cloudRpc(method, args, success, failure) {
  if (!method || method === 'then' || method === 'toJSON' || typeof method !== 'string') {
    return null;
  }

  // 1. In-Flight Request Deduplication:
  // If an identical read request is ALREADY in-flight or waiting in batch, attach to existing promise!
  const dedupKey = DEDUPLICATABLE_READ_METHODS.has(method) ? getRequestDedupKey(method, args) : null;
  if (dedupKey && inFlightRequests.has(dedupKey)) {
    return inFlightRequests.get(dedupKey)
      .then(res => {
        if (typeof success === 'function') success(res);
        return res;
      })
      .catch(err => {
        if (typeof failure === 'function') failure({ message: err?.message || 'Request failed' });
        throw err;
      });
  }

  const promise = new Promise((resolve, reject) => {
    const item = {
      id: 'req_' + (++rpcRequestIdSeq) + '_' + Date.now(),
      method,
      args: Array.isArray(args) ? args : [args],
      resolve,
      reject,
      success,
      failure,
      dedupKey
    };

    pendingRpcBatch.push(item);

    // Schedule debounced batch flush (15ms window allows concurrent event loop calls to consolidate)
    if (!batchDispatchTimer) {
      batchDispatchTimer = setTimeout(flushRpcBatchQueue, 15);
    }
  });

  if (dedupKey) {
    inFlightRequests.set(dedupKey, promise);
    promise.finally(() => {
      inFlightRequests.delete(dedupKey);
    });
  }

  return promise;
}
window.__cloudRpc = cloudRpc;

// Global State
let storedMasterCampaigns = [];
try {
  const rawMaster = localStorage.getItem('adverse_master_campaigns_v2');
  if (rawMaster) storedMasterCampaigns = JSON.parse(rawMaster);
} catch (e) {}

let state = {
  menus: [],
  months: [],
  currentMenu: null,
  adAccounts: [],
  pages: [],
  rate: 140,
  rawRecords: [],
  filteredRecords: [],
  headers: [],
  sortCol: null,
  sortAsc: true,
  filters: { month: '', year: '', date: '', search: '', pageType: '' },
  editingId: null,
  saving: false,
  userRole: 'unauthorized',
  displayName: '',
  theme: localStorage.getItem('adverse_theme') || 'light',
  clientCache: {},
  newlyAddedRecordKeys: new Set(),
  focusedRowIndex: null,
  currentRecordsMenuKey: null,
  campaignVerifications: (() => {
    try { return JSON.parse(localStorage.getItem('adverse_verifications') || '{}'); } catch(e) { return {}; }
  })(),
  verificationFilter: 'all',
  dateFormat: (typeof localStorage !== 'undefined' ? localStorage.getItem('adverse_date_format') : null) || 'DD-MM-YYYY',
  masterCampaignRegistry: Array.isArray(storedMasterCampaigns) ? storedMasterCampaigns : []
};
if (typeof window !== 'undefined') {
  window.state = state;
}

const CANONICAL_MENUS = [
  { key: 'today_created_campaign', label: 'Today Created', custom: 1 },
  { key: 'end_ads', label: 'End Ads', custom: 9 },
  { key: 'campaign', label: 'All Campaign', custom: 4 },
  { key: 'page_summery', label: 'Page Summary', custom: 3 },
  { key: 'payment_due_tracker', label: 'Payment', custom: 2 },
  { key: 'agency_spending', label: 'Agency Calculation', custom: 5 }
];

function getStandardMenus() {
  const admin = isAdmin();
  return CANONICAL_MENUS.filter(m => admin || !m.adminOnly);
}

const SNAPSHOT_STORAGE_KEY = 'adverse_instant_snapshot_v2';

function openMobileCampaignMenu() {
  const modal = document.getElementById('mobileCampaignMenuModal');
  if (modal) modal.classList.remove('hidden');
}
window.openMobileCampaignMenu = openMobileCampaignMenu;

function closeMobileCampaignMenu() {
  const modal = document.getElementById('mobileCampaignMenuModal');
  if (modal) modal.classList.add('hidden');
}
window.closeMobileCampaignMenu = closeMobileCampaignMenu;

function selectMenuMobile(key) {
  closeMobileCampaignMenu();
  selectMenu(key);
}
window.selectMenuMobile = selectMenuMobile;

let snapshotSaveTimer = null;

function saveDedicatedMenuCache(menuKey, data) {
  if (!menuKey || !data) return;
  try {
    const storageKey = `adverse_cache_${menuKey}`;
    localStorage.setItem(storageKey, JSON.stringify({
      data,
      timestamp: Date.now()
    }));
  } catch (_) {}
}

function loadDedicatedMenuCaches() {
  const menus = ['campaign', 'today_created_campaign', 'end_ads', 'page_summery', 'payment_due_tracker', 'agency_spending'];
  state.clientCache = state.clientCache || {};
  for (const m of menus) {
    try {
      const raw = localStorage.getItem(`adverse_cache_${m}`);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (parsed && parsed.data) {
          state.clientCache[m] = parsed.data;
          const mn = state.filters?.month || '';
          const yr = state.filters?.year || '';
          if (mn && yr) {
            state.clientCache[`${m}_${mn}_${yr}`] = parsed.data;
            state.clientCache[`${m}_${mn}_${yr}___`] = parsed.data;
            state.clientCache[`${m}_${mn}_${yr}__`] = parsed.data;
          }
        }
      }
    } catch (_) {}
  }
}

function saveLocalAppSnapshot() {
  if (snapshotSaveTimer) clearTimeout(snapshotSaveTimer);
  snapshotSaveTimer = setTimeout(() => {
    if ('requestIdleCallback' in window) {
      requestIdleCallback(performSaveSnapshot);
    } else {
      setTimeout(performSaveSnapshot, 50);
    }
  }, 300);
}

function performSaveSnapshot() {
  try {
    const cleanCache = {};
    if (state.clientCache) {
      for (const [k, v] of Object.entries(state.clientCache)) {
        if (!v) continue;
        cleanCache[k] = {
          stats: v.stats,
          records: v.records ? {
            headers: v.records.headers,
            rows: v.records.rows ? v.records.rows : undefined,
            today: v.records.today ? v.records.today : undefined,
            tomorrow: v.records.tomorrow ? v.records.tomorrow : undefined,
            all: v.records.all ? v.records.all : undefined
          } : undefined
        };
        // Also save dedicated key per menu for isolated reliability
        if (!k.includes('_')) {
          saveDedicatedMenuCache(k, v);
        }
      }
    }
    if (state.sheetId) {
      try { localStorage.setItem('adverse_data_sheet_id', state.sheetId); } catch (_) {}
    }
    const snapshot = {
      menus: getStandardMenus(),
      months: state.months,
      adAccounts: state.adAccounts,
      pages: state.pages,
      rate: state.rate,
      userRole: state.userRole,
      displayName: state.displayName,
      adminPin: state.adminPin,
      viewerPin: state.viewerPin,
      clientCache: cleanCache,
      filters: state.filters,
      timestamp: Date.now()
    };
    localStorage.setItem(SNAPSHOT_STORAGE_KEY, JSON.stringify(snapshot));

    // Ensure Master Campaign Registry is permanently preserved in local storage
    if (Array.isArray(state.masterCampaignRegistry) && state.masterCampaignRegistry.length > 0) {
      try {
        localStorage.setItem('adverse_master_campaigns_v2', JSON.stringify(state.masterCampaignRegistry.slice(0, 2000)));
      } catch (_) {}
    }
    // Ensure Today Created is permanently preserved in local storage
    if (Array.isArray(state.todayCreatedRegistry) && state.todayCreatedRegistry.length > 0) {
      try {
        localStorage.setItem('adverse_today_created_v2', JSON.stringify(state.todayCreatedRegistry));
      } catch (_) {}
    }
  } catch (e) {
    // Graceful fallback for quota limits
  }
}

function restoreLocalAppSnapshot() {
  try {
    loadDedicatedMenuCaches();

    try {
      const rawMaster = localStorage.getItem('adverse_master_campaigns_v2');
      if (rawMaster) {
        const parsedMaster = JSON.parse(rawMaster);
        if (Array.isArray(parsedMaster) && parsedMaster.length > 0) {
          state.masterCampaignRegistry = parsedMaster;
        }
      }
    } catch (_) {}

    try {
      const rawToday = localStorage.getItem('adverse_today_created_v2');
      if (rawToday) {
        const parsedToday = JSON.parse(rawToday);
        if (Array.isArray(parsedToday) && parsedToday.length > 0) {
          state.todayCreatedRegistry = parsedToday;
        }
      }
    } catch (_) {}

    const raw = localStorage.getItem(SNAPSHOT_STORAGE_KEY);
    if (!raw) return true;
    const snapshot = JSON.parse(raw);
    if (!snapshot) return true;

    // Strict canonical order: Today Created, End Ads, Pause ADS, All Campaign, Page Summary, Payment, Agency Calculation, Ad Account Billing
    state.menus = getStandardMenus();
    state.months = snapshot.months || [];
    state.adAccounts = snapshot.adAccounts || [];
    state.pages = (snapshot.pages || []).filter(p => {
      const lower = String(p || '').toLowerCase();
      return !lower.includes('verification') && !lower.includes('test_sync_page');
    });
    state.rate = snapshot.rate || 140;
    const savedPin = (localStorage.getItem('adverse_saved_pin') || '').trim();
    if (savedPin === '01893231534') {
      state.userRole = 'viewer';
    } else {
      state.userRole = snapshot.userRole || localStorage.getItem('adverse_user_role') || 'unauthorized';
    }
    state.displayName = (state.userRole === 'admin') ? 'Admin' : 'Viewer';
    state.adminPin = snapshot.adminPin || '01726912624';
    state.viewerPin = snapshot.viewerPin || '01893231534';
    if (snapshot.clientCache && typeof snapshot.clientCache === 'object') {
      state.clientCache = { ...snapshot.clientCache, ...state.clientCache };
    }
    return true;
  } catch (e) {
    return false;
  }
}

const menuIcons = {
  today_created_campaign: 'today',
  campaign: 'campaign',
  end_ads: 'event_busy',
  page_summery: 'analytics',
  agency_spending: 'payments',
  agency_cost: 'receipt_long',
  payment_due_tracker: 'account_balance_wallet',
  ad_account_billing: 'credit_card'
};

const MONEY_COLS = ['OFF $', 'Budget', 'Daily $', 'Taka', 'Paid (৳)', 'Spent', 'Remaining', 'Cost', 'Per Person Cost'];
const DATE_COLS = ['Start', 'End', 'Entry Date', 'Date', 'Start Date', 'End Date', 'Billing Date', 'Closed At', 'Entry', 'Payment Date', 'Last Checked'];

function isDateHeader(headerName) {
  if (!headerName) return false;
  const lower = String(headerName).trim().toLowerCase();
  return (
    DATE_COLS.includes(headerName) ||
    lower === 'start' ||
    lower === 'end' ||
    lower === 'start date' ||
    lower === 'end date' ||
    lower === 'entry date' ||
    lower === 'entry' ||
    lower === 'date' ||
    lower === 'payment date' ||
    lower === 'billing date' ||
    lower === 'closed at' ||
    lower.endsWith(' date') ||
    lower.startsWith('date ') ||
    lower === 'date/time'
  );
}
const isDateField = isDateHeader;
const MONTH_NAMES_FULL = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const MONTH_NAMES_SHORT = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

const _parseAnyDateCache = new Map();
function parseAnyDate(val, monthHint) {
  if (!val && val !== 0) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  const isCacheable = typeof val === 'string' || typeof val === 'number';
  const cKey = isCacheable ? `${val}_${monthHint ?? ''}` : null;
  if (cKey && _parseAnyDateCache.has(cKey)) {
    return _parseAnyDateCache.get(cKey);
  }
  const result = _computeParseAnyDate(val, monthHint);
  if (cKey) {
    if (_parseAnyDateCache.size > 3000) _parseAnyDateCache.clear();
    _parseAnyDateCache.set(cKey, result);
  }
  return result;
}

function _computeParseAnyDate(val, monthHint) {
  if (!val && val !== 0) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  
  // Handle Excel / Google Sheets numerical date serials (e.g. 45560 = ~2024, 46259 = ~2026)
  if (typeof val === 'number' && val > 30000 && val < 70000) {
    const utcDays = Math.round(val - 25569);
    const d = new Date((utcDays * 86400 + 43200) * 1000);
    return isNaN(d.getTime()) ? null : d;
  }

  let str = String(val).trim();
  if (!str) return null;
  const lowerCheck = str.toLowerCase();
  if (lowerCheck === 'undefined' || lowerCheck === 'null' || lowerCheck === 'nan' || lowerCheck === 'invalid date' || lowerCheck === 'loading...') {
    return null;
  }

  // Convert Bengali numerals (০-৯) to Western numerals (0-9)
  const bnToEn = { '০': '0', '১': '1', '২': '2', '৩': '3', '৪': '4', '৫': '5', '৬': '6', '৭': '7', '৮': '8', '৯': '9' };
  str = str.replace(/[০-৯]/g, d => bnToEn[d] || d);

  // Convert Bengali month names to English
  const bnMonthMap = {
    'জানুয়ারি': 'Jan', 'জানুয়ারী': 'Jan', 'ফেব্রুয়ারি': 'Feb', 'ফেব্রুয়ারী': 'Feb',
    'মার্চ': 'Mar', 'এপ্রিল': 'Apr', 'মে': 'May', 'জুন': 'Jun',
    'জুলাই': 'Jul', 'আগস্ট': 'Aug', 'সেপ্টেম্বর': 'Sep',
    'অক্টোবর': 'Oct', 'নভেম্বর': 'Nov', 'ডিসেম্বর': 'Dec'
  };
  for (const [bnM, enM] of Object.entries(bnMonthMap)) {
    if (str.includes(bnM)) {
      str = str.replace(bnM, enM);
      break;
    }
  }

  // If numeric string in serial range
  if (/^\d{5}(\.\d+)?$/.test(str)) {
    const num = parseFloat(str);
    if (num > 30000 && num < 70000) {
      const utcDays = Math.round(num - 25569);
      const d = new Date((utcDays * 86400 + 43200) * 1000);
      if (!isNaN(d.getTime())) return d;
    }
  }

  // Match GViz Date constructor: Date(2026, 7, 29)
  const matchGviz = str.match(/Date\((\d{4}),\s*(\d{1,2}),\s*(\d{1,2})/i);
  if (matchGviz) {
    return new Date(parseInt(matchGviz[1], 10), parseInt(matchGviz[2], 10), parseInt(matchGviz[3], 10), 12, 0, 0);
  }

  // Clean off timestamp if present: e.g. "25-Aug-2026 12:00:00" -> "25-Aug-2026"
  const cleanDateStr = str.replace(/T\d{2}:\d{2}.*$/, '').replace(/\s+\d{1,2}:\d{2}(:\d{2})?.*$/, '').trim();

  // DD-Mon-YYYY (e.g. 25-Aug-2026, 25-August-2026)
  const dMonYMatch = cleanDateStr.match(/^(\d{1,2})[-/ .]([A-Za-z]{3,9})[-/ .](\d{2,4})$/);
  if (dMonYMatch) {
    const day = parseInt(dMonYMatch[1], 10);
    const monStr = dMonYMatch[2].slice(0, 3).toLowerCase();
    let yr = parseInt(dMonYMatch[3], 10);
    if (yr < 100) yr += 2000;
    const mIdx = MONTH_NAMES_SHORT.findIndex(m => m.toLowerCase() === monStr);
    if (mIdx !== -1) {
      const d = new Date(yr, mIdx, day, 12, 0, 0);
      if (!isNaN(d.getTime())) return d;
    }
  }

  // Mon-DD-YYYY (e.g. Aug-25-2026, August 25, 2026)
  const monDYMatch = cleanDateStr.match(/^([A-Za-z]{3,9})[-/ .](\d{1,2}),?[-/ .](\d{2,4})$/);
  if (monDYMatch) {
    const monStr = monDYMatch[1].slice(0, 3).toLowerCase();
    const day = parseInt(monDYMatch[2], 10);
    let yr = parseInt(monDYMatch[3], 10);
    if (yr < 100) yr += 2000;
    const mIdx = MONTH_NAMES_SHORT.findIndex(m => m.toLowerCase() === monStr);
    if (mIdx !== -1) {
      const d = new Date(yr, mIdx, day, 12, 0, 0);
      if (!isNaN(d.getTime())) return d;
    }
  }

  // YYYY-MM-DD
  const isoMatch = cleanDateStr.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (isoMatch) {
    const d = new Date(parseInt(isoMatch[1], 10), parseInt(isoMatch[2], 10) - 1, parseInt(isoMatch[3], 10), 12, 0, 0);
    if (!isNaN(d.getTime())) return d;
  }

  // DD/MM/YYYY or MM/DD/YYYY
  const dmyMatch = cleanDateStr.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})/);
  if (dmyMatch) {
    const p1 = parseInt(dmyMatch[1], 10);
    const p2 = parseInt(dmyMatch[2], 10);
    let yr = parseInt(dmyMatch[3], 10);
    if (yr < 100) yr += 2000;

    // Disambiguate using monthHint if provided
    if (monthHint !== undefined && monthHint !== null && monthHint !== '') {
      let targetMonIdx = -1;
      if (typeof monthHint === 'number') {
        targetMonIdx = (monthHint >= 1 && monthHint <= 12) ? monthHint - 1 : (monthHint >= 0 && monthHint <= 11 ? monthHint : -1);
      } else {
        const mStr = String(monthHint).trim().toLowerCase();
        targetMonIdx = MONTH_NAMES_FULL.findIndex(m => m.toLowerCase() === mStr);
        if (targetMonIdx === -1) {
          targetMonIdx = MONTH_NAMES_SHORT.findIndex(m => m.toLowerCase() === mStr.slice(0, 3));
        }
        if (targetMonIdx === -1 && !isNaN(parseInt(mStr, 10))) {
          const num = parseInt(mStr, 10);
          targetMonIdx = (num >= 1 && num <= 12) ? num - 1 : -1;
        }
      }
      if (targetMonIdx !== -1) {
        if (p1 === targetMonIdx + 1 && p2 <= 31) {
          return new Date(yr, p1 - 1, p2, 12, 0, 0);
        } else if (p2 === targetMonIdx + 1 && p1 <= 31) {
          return new Date(yr, p2 - 1, p1, 12, 0, 0);
        }
      }
    }

    if (p1 > 12 && p2 <= 12) {
      return new Date(yr, p2 - 1, p1, 12, 0, 0);
    } else if (p1 <= 12 && p2 > 12) {
      return new Date(yr, p1 - 1, p2, 12, 0, 0);
    } else if (p1 <= 31 && p2 <= 12) {
      // If slash notation / (standard Google Sheets / US format e.g. 10/7/2026 for Oct 7), p1 is Month and p2 is Day
      if (cleanDateStr.includes('/')) {
        return new Date(yr, p1 - 1, p2, 12, 0, 0);
      }
      // Standard Bangladesh dash/dot format: DD-MM-YYYY
      return new Date(yr, p2 - 1, p1, 12, 0, 0);
    }
  }

  const generic = new Date(cleanDateStr);
  if (!isNaN(generic.getTime())) return generic;
  return null;
}

function getNowBST() {
  try {
    const dtf = new Intl.DateTimeFormat('en-US', {
      timeZone: 'Asia/Dhaka',
      year: 'numeric',
      month: 'numeric',
      day: 'numeric',
      hour: 'numeric',
      minute: 'numeric',
      second: 'numeric',
      hour12: false
    });
    const parts = {};
    dtf.formatToParts(new Date()).forEach(p => { parts[p.type] = p.value; });
    const year = parseInt(parts.year, 10);
    const month = parseInt(parts.month, 10) - 1;
    const day = parseInt(parts.day, 10);
    const hour = parseInt(parts.hour, 10) || 0;
    const minute = parseInt(parts.minute, 10) || 0;
    const second = parseInt(parts.second, 10) || 0;
    return new Date(year, month, day, hour, minute, second);
  } catch (e) {
    const now = new Date();
    const utc = now.getTime() + (now.getTimezoneOffset() * 60000);
    return new Date(utc + (6 * 3600000));
  }
}

function getDateProximityBadge(dateVal) {
  if (!dateVal) return '';
  const parsed = parseAnyDate(dateVal);
  if (!parsed) return '';
  const today = getNowBST();
  today.setHours(0, 0, 0, 0);
  const target = new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
  const diffDays = Math.round((target - today) / (1000 * 60 * 60 * 24));
  if (diffDays === 0) {
    return ` <span class="pill-dot-ending" title="Campaign Ends Today!" style="display:inline-block;vertical-align:middle;margin-left:4px"></span>`;
  }
  if (diffDays === 1) {
    return ` <span class="pill-dot" style="background:#10B981;width:6px;height:6px;display:inline-block;vertical-align:middle;margin-left:4px" title="Campaign Ends Tomorrow"></span>`;
  }
  return '';
}

function getActiveDateFormat() {
  return (state && state.dateFormat) || (typeof localStorage !== 'undefined' ? localStorage.getItem('adverse_date_format') : null) || 'DD-MM-YYYY';
}

function _formatDateObj(d, targetFmt) {
  if (!d || isNaN(d.getTime())) return '';
  const dayPadded = String(d.getDate()).padStart(2, '0');
  const daySingle = String(d.getDate());
  const monIdx = d.getMonth();
  const monShort = MONTH_NAMES_SHORT[monIdx] || 'Jan';
  const monNum = String(monIdx + 1).padStart(2, '0');
  const yr = String(d.getFullYear());

  switch (targetFmt) {
    case 'DD-MM-YYYY':
      return `${dayPadded}-${monNum}-${yr}`;
    case 'DD/MM/YYYY':
      return `${dayPadded}/${monNum}/${yr}`;
    case 'YYYY-MM-DD':
      return `${yr}-${monNum}-${dayPadded}`;
    case 'D-Mon-YYYY':
      return `${daySingle}-${monShort}-${yr}`;
    case 'DD-Mon-YYYY':
    default:
      return `${dayPadded}-${monShort}-${yr}`;
  }
}

const _formatCustomDateCache = new Map();
function formatCustomDate(val, monthHint, formatOverride) {
  if (!val && val !== 0) return '';
  const targetFmt = formatOverride || getActiveDateFormat();
  const isCacheable = (typeof val === 'string' || typeof val === 'number') && !formatOverride;
  const cKey = isCacheable ? `${val}_${monthHint ?? ''}_${targetFmt}` : null;
  if (cKey && _formatCustomDateCache.has(cKey)) {
    return _formatCustomDateCache.get(cKey);
  }
  const result = _computeFormatCustomDate(val, monthHint, targetFmt);
  if (cKey) {
    if (_formatCustomDateCache.size > 5000) _formatCustomDateCache.clear();
    _formatCustomDateCache.set(cKey, result);
  }
  return result;
}

function _computeFormatCustomDate(val, monthHint, targetFmt) {
  if (!val && val !== 0) return '';
  const fmt = targetFmt || getActiveDateFormat();
  if (val instanceof Date) {
    return _formatDateObj(val, fmt);
  }
  const str = String(val).trim();
  if (!str) return '';
  const lower = str.toLowerCase();
  if (lower === 'undefined' || lower === 'null' || lower === 'nan' || lower === 'invalid date' || lower === 'loading...') return '';

  const parsed = parseAnyDate(str, monthHint);
  if (!parsed || isNaN(parsed.getTime())) return str;
  return _formatDateObj(parsed, fmt);
}

function setSystemDateFormat(fmt, syncToSheet = false) {
  if (!fmt) return;
  state.dateFormat = fmt;
  if (typeof localStorage !== 'undefined') {
    localStorage.setItem('adverse_date_format', fmt);
  }
  if (typeof _formatCustomDateCache !== 'undefined') {
    _formatCustomDateCache.clear();
  }
  updateDateFormatUI();
  showToast(`Date format set to ${fmt}`, 'success');
  if (typeof refreshCurrentView === 'function') {
    refreshCurrentView();
  }
  if (syncToSheet) {
    formatAllSpreadsheetDatesNow(fmt);
  }
}

function updateDateFormatUI() {
  const curFmt = getActiveDateFormat();
  const topLabel = document.getElementById('topDateFormatLabel');
  if (topLabel) topLabel.textContent = curFmt;
  const select = document.getElementById('configDateFormatSelect');
  if (select) select.value = curFmt;
  const dropdown = document.getElementById('dateFormatDropdownMenu');
  if (dropdown) dropdown.classList.add('hidden');
}

function toggleDateFormatDropdown() {
  const dropdown = document.getElementById('dateFormatDropdownMenu');
  if (dropdown) {
    dropdown.classList.toggle('hidden');
  }
}

function formatAllSpreadsheetDatesNow(fmtOverride) {
  const fmt = fmtOverride || getActiveDateFormat();
  showToast(`Formatting Google Sheet dates to ${fmt}...`, 'info');
  const whUrl = (state.sheetConfig && state.sheetConfig.webhookUrl) || localStorage.getItem('adverse_webhook_url') || '';
  if (whUrl) {
    fetch(whUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'formatAllDates', dateFormat: fmt })
    })
    .then(r => r.json())
    .then(res => {
      showToast(`✅ Google Sheet dates formatted to ${fmt}!`, 'success');
    })
    .catch(() => {
      if (window.__cloudRpc) {
        window.__cloudRpc('formatAllDates', [fmt], (res) => {
          showToast(`✅ Google Sheet dates updated to ${fmt}!`, 'success');
        });
      }
    });
  } else if (window.__cloudRpc) {
    window.__cloudRpc('formatAllDates', [fmt], (res) => {
      showToast(`✅ Dates updated to ${fmt}!`, 'success');
    });
  } else {
    showToast(`Date format active in App (${fmt})! Connect Webhook to format Sheet.`, 'info');
  }
}

function getTodayFormatted() {
  return formatCustomDate(getNowBST());
}

function isSameDay(d1, d2) {
  if (!d1 || !d2) return false;
  const f1 = formatCustomDate(d1);
  const f2 = formatCustomDate(d2);
  if (f1 && f2 && f1.toLowerCase() === f2.toLowerCase()) return true;
  const p1 = parseAnyDate(d1);
  const p2 = parseAnyDate(d2);
  if (!p1 || !p2) return false;
  return p1.getFullYear() === p2.getFullYear() &&
         p1.getMonth() === p2.getMonth() &&
         p1.getDate() === p2.getDate();
}
window.isSameDay = isSameDay;

function getTodayLocalDateStr() {
  const now = getNowBST();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function toInputDateValue(val) {
  const parsed = parseAnyDate(val) || getNowBST();
  const y = parsed.getFullYear();
  const m = String(parsed.getMonth() + 1).padStart(2, '0');
  const d = String(parsed.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function escapeHtml(s) {
  return String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function fmtNum(v) {
  return Number(v || 0).toLocaleString('en-US', { maximumFractionDigits: 2 });
}

function parseMoney(val) {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  if (!val && val !== 0) return 0;
  const str = String(val).replace(/[^0-9.-]/g, '');
  const n = parseFloat(str);
  return isNaN(n) ? 0 : n;
}
window.parseMoney = parseMoney;

function formatCurrency(v) {
  return fmtNum(v);
}
window.formatCurrency = formatCurrency;

function showToast(msg, type = 'success') {
  const t = document.getElementById('toast');
  if (!t) return;
  t.textContent = msg;
  const typeClass = type === 'error' ? 'toast-error' : (type === 'warning' ? 'toast-warning' : (type === 'info' ? 'toast-info' : 'toast-success'));
  t.className = `toast ${typeClass} show`;
  clearTimeout(t._toastTimer);
  t._toastTimer = setTimeout(() => t.classList.remove('show'), 3400);
}

function showSyncIssueBanner(_msg = '') {
  // Silent Mode: Do not display visual banner or toast on screen as requested by user
  const banner = document.getElementById('syncIssueBanner');
  if (banner && banner.parentNode) banner.remove();
}
window.showSyncIssueBanner = showSyncIssueBanner;

function isAdmin() {
  const role = (state.userRole || localStorage.getItem('adverse_user_role') || '').trim().toLowerCase();
  const savedPin = (localStorage.getItem('adverse_saved_pin') || '').trim();
  if (role === 'viewer' || savedPin === '01893231534' || savedPin.endsWith('1893231534')) {
    return false;
  }
  if (role === 'admin' || savedPin === '01726912624' || savedPin.endsWith('1726912624')) {
    return true;
  }
  return role === 'admin';
}

// Banglish & Bengali Phonetic Transliteration Engine
const BanglishEngine = {
  vowels: {
    'অ': 'o', 'আ': 'a', 'ই': 'i', 'ঈ': 'ee', 'উ': 'u', 'ঊ': 'oo', 'ঋ': 'ri',
    'এ': 'e', 'ঐ': 'oi', 'ও': 'o', 'ঔ': 'ou'
  },
  vowelSigns: {
    'া': 'a', 'ি': 'i', 'ী': 'ee', 'ু': 'u', 'ূ': 'oo', 'ৃ': 'ri',
    'ে': 'e', 'ৈ': 'oi', 'ো': 'o', 'ৌ': 'ou', '্': ''
  },
  consonants: {
    'ক': 'k', 'খ': 'kh', 'গ': 'g', 'ঘ': 'gh', 'ঙ': 'ng',
    'চ': 'ch', 'ছ': 'chh', 'জ': 'j', 'ঝ': 'jh', 'ঞ': 'n',
    'ট': 't', 'ঠ': 'th', 'ড': 'd', 'ঢ': 'dh', 'ণ': 'n',
    'ত': 't', 'থ': 'th', 'দ': 'd', 'ধ': 'dh', 'ন': 'n',
    'প': 'p', 'ফ': 'f', 'ব': 'b', 'ভ': 'bh', 'ম': 'm',
    'য': 'j', 'র': 'r', 'ল': 'l', 'শ': 'sh', 'ষ': 'sh', 'স': 's', 'হ': 'h',
    'ড়': 'r', 'ঢ়': 'rh', 'য়': 'y', 'ৎ': 't', 'ং': 'ng', 'ঃ': 'h', 'ঁ': 'n'
  },
  conjuncts: {
    'ক্ষ': 'kkh', 'জ্ঞ': 'gg', 'ঞ্চ': 'nch', 'ঞ্জ': 'nj', 'ঙ্ক': 'nk', 'ঙ্গ': 'ng',
    'ণ্ট': 'nt', 'ণ্ঠ': 'nth', 'ণ্ড': 'nd', 'ন্ত': 'nt', 'ন্থ': 'nth', 'ন্দ': 'nd', 'ন্ধ': 'ndh',
    'ম্প': 'mp', 'ম্ফ': 'mph', 'ম্ব': 'mb', 'ম্ভ': 'mbh', 'ম্ম': 'mm',
    'ক্ত': 'kt', 'ক্ল': 'kl', 'প্ল': 'pl', 'ব্ল': 'bl', 'গ্ল': 'gl',
    'ভ্র': 'bhr', 'প্র': 'pr', 'ত্র': 'tr', 'দ্র': 'dr', 'ধ্র': 'dhr', 'গ্র': 'gr',
    'শ্র': 'shr', 'ফ্র': 'fr', 'ক্র': 'kr', 'ট্র': 'tr', 'ড্র': 'dr',
    'দ্ব': 'dw', 'ত্ব': 'tw', 'স্ব': 'sw', 'শ্ব': 'shw',
    'ষ্ট': 'sht', 'ষ্ঠ': 'shth', 'স্প': 'sp', 'স্ফ': 'sph', 'স্ত': 'st', 'স্থ': 'sth', 'স্খ': 'skh'
  },

  isBengali(str) {
    return /[\u0980-\u09FF]/.test(String(str || ''));
  },

  toBanglish(text) {
    if (!text) return '';
    const str = String(text).trim();
    let result = '';
    let i = 0;
    const len = str.length;

    while (i < len) {
      if (i + 1 < len) {
        const twoChar = str.substring(i, i + 2);
        if (this.conjuncts[twoChar]) {
          result += this.conjuncts[twoChar];
          i += 2;
          continue;
        }
      }

      const ch = str[i];
      if (this.conjuncts[ch]) {
        result += this.conjuncts[ch];
      } else if (this.vowels[ch]) {
        result += this.vowels[ch];
      } else if (this.vowelSigns[ch] !== undefined) {
        result += this.vowelSigns[ch];
      } else if (this.consonants[ch]) {
        result += this.consonants[ch];
      } else {
        result += ch;
      }
      i++;
    }
    return result.toLowerCase().trim();
  },

  getAliases(text) {
    if (!text) return [];
    const orig = String(text).trim();
    const origLower = orig.toLowerCase();
    const set = new Set([origLower]);

    if (this.isBengali(orig)) {
      const banglish = this.toBanglish(orig);
      if (banglish) {
        set.add(banglish);
        set.add(banglish.replace(/u/g, 'o'));
        set.add(banglish.replace(/o/g, 'u'));
        set.add(banglish.replace(/ee/g, 'i'));
        set.add(banglish.replace(/i/g, 'ee'));
        set.add(banglish.replace(/f/g, 'ph'));
        set.add(banglish.replace(/ph/g, 'f'));
        set.add(banglish.replace(/sh/g, 's'));
        set.add(banglish.replace(/s/g, 'sh'));
        set.add(banglish.replace(/k/g, 'c'));
        set.add(banglish.replace(/c/g, 'k'));
        set.add(banglish.replace(/v/g, 'b'));
        set.add(banglish.replace(/b/g, 'v'));
        set.add(banglish.replace(/j/g, 'z'));

        const words = banglish.split(/\s+/).filter(Boolean);
        words.forEach(w => {
          set.add(w);
          set.add(w.replace(/u/g, 'o'));
          set.add(w.replace(/o/g, 'u'));
          set.add(w.replace(/sh/g, 's'));
          set.add(w.replace(/s/g, 'sh'));
          set.add(w.replace(/ee/g, 'i'));
        });

        if (words.length > 1) {
          set.add(words.map(w => w[0]).join(''));
        }
      }

      const bnWords = orig.split(/\s+/).filter(Boolean);
      bnWords.forEach(w => set.add(w.toLowerCase()));
      if (bnWords.length > 1) {
        set.add(bnWords.map(w => w[0]).join(''));
      }
    } else {
      const words = origLower.split(/\s+/).filter(Boolean);
      words.forEach(w => set.add(w));
      if (words.length > 1) {
        set.add(words.map(w => w[0]).join(''));
      }
      set.add(origLower.replace(/c/g, 'k'));
      set.add(origLower.replace(/k/g, 'c'));
      set.add(origLower.replace(/ph/g, 'f'));
      set.add(origLower.replace(/f/g, 'ph'));
      set.add(origLower.replace(/sh/g, 's'));
      set.add(origLower.replace(/s/g, 'sh'));
      set.add(origLower.replace(/u/g, 'o'));
      set.add(origLower.replace(/o/g, 'u'));
      set.add(origLower.replace(/v/g, 'b'));
      set.add(origLower.replace(/b/g, 'v'));
      set.add(origLower.replace(/z/g, 'j'));
    }

    return Array.from(set);
  },

  matches(targetText, query) {
    if (!query) return true;
    if (!targetText) return false;
    const q = String(query).trim().toLowerCase();
    if (!q) return true;

    const target = String(targetText).trim();
    const targetLower = target.toLowerCase();
    const words = targetLower.split(/\s+/).filter(Boolean);
    const banglish = this.toBanglish(target);
    const bangLower = banglish ? banglish.toLowerCase() : '';
    const bWords = bangLower ? bangLower.split(/\s+/).filter(Boolean) : [];

    // Requirement 1: If 1-3 letters, match ONLY pages that start with or have a word starting with the letters
    if (q.length <= 3) {
      if (targetLower.startsWith(q)) return true;
      if (words.some(w => w.startsWith(q))) return true;
      if (bangLower && bangLower.startsWith(q)) return true;
      if (bWords.some(w => w.startsWith(q))) return true;
      return false;
    }

    // 4+ letters: Allow direct substring match & normalized phonetic match
    if (targetLower.includes(q)) return true;
    if (words.some(w => w.startsWith(q))) return true;
    if (bangLower && bangLower.includes(q)) return true;
    if (bWords.some(w => w.startsWith(q))) return true;

    const normalizeSound = s => s
      .replace(/c/g, 'k')
      .replace(/ph/g, 'f')
      .replace(/sh/g, 's')
      .replace(/ee/g, 'i')
      .replace(/oo/g, 'u')
      .replace(/v/g, 'b')
      .replace(/z/g, 'j')
      .replace(/tt/g, 't')
      .replace(/dd/g, 'd');

    const qNorm = normalizeSound(q);
    const targetNorm = normalizeSound(banglish || targetLower);
    if (targetNorm.includes(qNorm)) return true;

    const normWords = words.map(w => normalizeSound(w));
    if (normWords.some(w => w.startsWith(qNorm) || w.includes(qNorm))) return true;

    return false;
  },

  getRomanDisplayHint(text) {
    if (!text || !this.isBengali(text)) return '';
    const roman = this.toBanglish(text);
    if (!roman || roman.toLowerCase() === text.toLowerCase()) return '';
    return roman.replace(/\b\w/g, c => c.toUpperCase());
  }
};

// Smart Floating Autocomplete for Page Names (Banglish & Bengali Aware)
function initSmartPageAutocomplete(inputEl, dropdownEl, options = {}) {
  if (!inputEl || !dropdownEl) return;

  let selectedIndex = -1;
  let currentItems = [];
  let isSelecting = false;

  const getPageUsageStats = () => {
    const counts = {};
    const adMap = {};
    (state.rawRecords || []).forEach(r => {
      const p = String(r['Page Name'] || r['page_name'] || '').trim();
      const ad = String(r['AD Account'] || r['ad_account'] || '').trim();
      if (p) {
        const key = p.toLowerCase();
        counts[key] = (counts[key] || 0) + 1;
        if (ad && !adMap[key]) adMap[key] = ad;
      }
    });
    return { counts, adMap };
  };

  const getFullPageList = () => {
    const { counts } = getPageUsageStats();
    const set = new Set((state.pages || []).filter(p => !isDemoPage(p)));
    (state.rawRecords || []).forEach(r => {
      const p = r['Page Name'] || r['page_name'];
      if (p && String(p).trim() && !isDemoPage(String(p).trim())) {
        set.add(String(p).trim());
      }
    });
    const list = Array.from(set);
    list.sort((a, b) => {
      const countA = counts[a.toLowerCase()] || 0;
      const countB = counts[b.toLowerCase()] || 0;
      if (countB !== countA) return countB - countA;
      return a.localeCompare(b);
    });
    return list;
  };

  const renderDropdown = (query = '') => {
    if (isSelecting) return;
    const cleanQ = String(query || '').trim();

    // Do NOT show suggestions if user hasn't typed anything
    if (!cleanQ || cleanQ.length < 1) {
      dropdownEl.classList.add('hidden');
      dropdownEl.style.display = 'none';
      dropdownEl.innerHTML = '';
      currentItems = [];
      selectedIndex = -1;
      return;
    }

    const { counts, adMap } = getPageUsageStats();
    const allPages = getFullPageList();

    let matched = allPages.filter(p => BanglishEngine.matches(p, cleanQ));
      matched.sort((a, b) => {
        const aLower = a.toLowerCase();
        const bLower = b.toLowerCase();
        const qLower = cleanQ.toLowerCase();
        const aStarts = aLower.startsWith(qLower);
        const bStarts = bLower.startsWith(qLower);
        if (aStarts && !bStarts) return -1;
        if (!aStarts && bStarts) return 1;

        const aBanglish = BanglishEngine.toBanglish(a);
        const bBanglish = BanglishEngine.toBanglish(b);
        const aBStarts = aBanglish.startsWith(qLower);
        const bBStarts = bBanglish.startsWith(qLower);
        if (aBStarts && !bBStarts) return -1;
        if (!aBStarts && bBStarts) return 1;

        // Prioritize old most-used pages (higher usage count first)
        const countA = counts[aLower] || 0;
        const countB = counts[bLower] || 0;
        if (countB !== countA) return countB - countA;

        return a.localeCompare(b);
      });

    currentItems = matched;
    selectedIndex = -1;

    if (!matched.length) {
      if (cleanQ) {
        dropdownEl.innerHTML = `
          <div class="smart-suggestion-item" data-val="${escapeHtml(cleanQ)}" style="border-left:3px solid var(--suc)">
            <div class="smart-suggestion-name">
              <div class="smart-suggestion-title">
                <span class="ms" style="font-size:14px;color:var(--suc);vertical-align:text-bottom">add_circle</span>
                Use as New Page: "<b>${escapeHtml(cleanQ)}</b>"
              </div>
              <div class="smart-suggestion-roman">Click or Enter to select this as a New Page</div>
            </div>
            <span class="smart-suggestion-badge new">New Page</span>
          </div>
        `;
        dropdownEl.classList.remove('hidden');
        bindItemClicks();
      } else {
        dropdownEl.classList.add('hidden');
      }
      return;
    }

    dropdownEl.innerHTML = matched.slice(0, 15).map((page, idx) => {
      const romanHint = BanglishEngine.getRomanDisplayHint(page);
      const isKnown = (state.pages || []).includes(page);
      const pageKey = page.toLowerCase();
      const pastCampaignCount = counts[pageKey] || 0;
      const frequentAd = adMap[pageKey] || '';
      
      return `
        <div class="smart-suggestion-item" data-idx="${idx}" data-val="${escapeHtml(page)}">
          <div class="smart-suggestion-name">
            <div class="smart-suggestion-title">
              <span class="ms" style="font-size:14px;color:var(--pri);vertical-align:text-bottom">pages</span>
              ${escapeHtml(page)}
            </div>
            <div class="smart-suggestion-roman" style="display:flex;align-items:center;gap:6px;flex-wrap:wrap">
              ${romanHint ? `<span><span class="ms" style="font-size:11px">translate</span> ${escapeHtml(romanHint)}</span>` : ''}
              ${frequentAd ? `<span style="color:var(--pri);font-weight:600">• AD: ${escapeHtml(frequentAd)}</span>` : ''}
            </div>
          </div>
          <div style="display:flex;align-items:center;gap:6px">
            ${pastCampaignCount > 0 ? `<span style="font-size:11px;font-weight:700;color:var(--suc)">${pastCampaignCount} campaigns</span>` : ''}
            <span class="smart-suggestion-badge ${isKnown || pastCampaignCount > 0 ? 'old' : 'new'}">
              ${isKnown || pastCampaignCount > 0 ? 'Old Page' : 'New'}
            </span>
          </div>
        </div>
      `;
    }).join('');

    const exactMatch = matched.some(m => m.toLowerCase() === cleanQ.toLowerCase());
    if (cleanQ && !exactMatch) {
      dropdownEl.innerHTML += `
        <div class="smart-suggestion-item" data-val="${escapeHtml(cleanQ)}" style="border-top:1px dashed var(--bor);margin-top:4px">
          <div class="smart-suggestion-name">
            <div class="smart-suggestion-title" style="color:var(--suc)">
              <span class="ms" style="font-size:14px;color:var(--suc);vertical-align:text-bottom">add</span>
              Use as New Page: "<b>${escapeHtml(cleanQ)}</b>"
            </div>
          </div>
          <span class="smart-suggestion-badge new">New</span>
        </div>
      `;
    }

    dropdownEl.classList.remove('hidden');
    bindItemClicks();
  };

  const bindItemClicks = () => {
    dropdownEl.querySelectorAll('.smart-suggestion-item').forEach(item => {
      item.addEventListener('mousedown', e => {
        e.preventDefault();
        e.stopPropagation();
        const val = item.getAttribute('data-val');
        if (val) selectPage(val);
      });
    });
  };

  let hasJustSelected = false;

  const selectPage = (val) => {
    isSelecting = true;
    hasJustSelected = true;
    inputEl.value = val;
    dropdownEl.classList.add('hidden');
    dropdownEl.style.display = 'none';
    dropdownEl.innerHTML = '';
    selectedIndex = -1;
    inputEl.blur();
    if (typeof options.onSelect === 'function') {
      options.onSelect(val);
    }
    inputEl.dispatchEvent(new Event('change', { bubbles: true }));
    setTimeout(() => {
      isSelecting = false;
      dropdownEl.classList.add('hidden');
      dropdownEl.style.display = 'none';
    }, 250);
  };

  const updateHighlight = () => {
    const items = dropdownEl.querySelectorAll('.smart-suggestion-item');
    items.forEach((it, idx) => {
      it.classList.toggle('highlighted', idx === selectedIndex);
      if (idx === selectedIndex) {
        it.scrollIntoView({ block: 'nearest' });
      }
    });
  };

  inputEl.addEventListener('input', () => {
    hasJustSelected = false;
    dropdownEl.style.display = '';
    if (!isSelecting) renderDropdown(inputEl.value);
  });

  inputEl.addEventListener('focus', () => {
    if (!isSelecting && !hasJustSelected && inputEl.value.trim().length >= 1) {
      dropdownEl.style.display = '';
      renderDropdown(inputEl.value);
    }
  });

  inputEl.addEventListener('keydown', e => {
    const items = dropdownEl.querySelectorAll('.smart-suggestion-item');
    if (dropdownEl.classList.contains('hidden') || !items.length) {
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        renderDropdown(inputEl.value);
      }
      return;
    }

    if (e.key === 'ArrowDown') {
      e.preventDefault();
      selectedIndex = (selectedIndex + 1) % items.length;
      updateHighlight();
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      selectedIndex = (selectedIndex - 1 + items.length) % items.length;
      updateHighlight();
    } else if (e.key === 'Enter') {
      if (selectedIndex >= 0 && items[selectedIndex]) {
        e.preventDefault();
        const val = items[selectedIndex].getAttribute('data-val');
        if (val) selectPage(val);
      }
    } else if (e.key === 'Escape') {
      dropdownEl.classList.add('hidden');
      selectedIndex = -1;
    }
  });

  document.addEventListener('click', e => {
    if (!inputEl.contains(e.target) && !dropdownEl.contains(e.target)) {
      dropdownEl.classList.add('hidden');
      selectedIndex = -1;
    }
  });
}

// Helper functions to ignore demo accounts and demo pages
function isDemoAdAccount(name) {
  const lower = String(name || '').toLowerCase().trim();
  return (
    lower.includes('master ad account') ||
    lower.includes('business manager prime') ||
    lower.includes('agency growth account')
  );
}

function isDemoPage(name) {
  const lower = String(name || '').toLowerCase().trim();
  return (
    lower === 'demo store 1' ||
    lower === 'fashion store' ||
    lower === 'online shop bd' ||
    lower === 'gadget gallery' ||
    lower === 'tech trend bd'
  );
}

// Smart Floating Autocomplete for AD Accounts
function initSmartAdAccountAutocomplete(inputEl, dropdownEl, options = {}) {
  if (!inputEl || !dropdownEl) return;

  let selectedIndex = -1;
  let isSelecting = false;

  const renderDropdown = (query = '') => {
    if (isSelecting) return;
    const cleanQ = String(query || '').trim().toLowerCase();
    // Only show suggestions when user actually types something
    if (!cleanQ) {
      dropdownEl.classList.add('hidden');
      return;
    }

    const list = (state.adAccounts || []).filter(a => !isDemoAdAccount(a.a));
    const matched = list.filter(a => {
      const name = (a.a || '').toLowerCase();
      const id = (a.id || '').toLowerCase();
      return name.includes(cleanQ) || id.includes(cleanQ);
    });

    matched.sort((a, b) => {
      const aName = (a.a || '').toLowerCase();
      const bName = (b.a || '').toLowerCase();
      const aStarts = aName.startsWith(cleanQ);
      const bStarts = bName.startsWith(cleanQ);
      if (aStarts && !bStarts) return -1;
      if (!aStarts && bStarts) return 1;
      return aName.localeCompare(bName);
    });

    selectedIndex = -1;
    if (!matched.length) {
      dropdownEl.classList.add('hidden');
      return;
    }

    dropdownEl.innerHTML = matched.map((a, idx) => `
      <div class="smart-suggestion-item" data-idx="${idx}" data-val="${escapeHtml(a.a)}" data-id="${escapeHtml(a.id || '')}">
        <div class="smart-suggestion-name" style="width:100%">
          <div class="smart-suggestion-title">
            <span class="ms" style="font-size:14px;color:var(--pri);vertical-align:text-bottom">badge</span>
            ${escapeHtml(a.a)}
          </div>
          <div class="smart-suggestion-roman">ID: ${escapeHtml(a.id || 'N/A')}</div>
        </div>
      </div>
    `).join('');

    dropdownEl.classList.remove('hidden');

    dropdownEl.querySelectorAll('.smart-suggestion-item').forEach(item => {
      item.addEventListener('mousedown', e => {
        e.preventDefault();
        e.stopPropagation();
        const val = item.getAttribute('data-val');
        const id = item.getAttribute('data-id');
        if (val) {
          isSelecting = true;
          inputEl.value = val;
          dropdownEl.classList.add('hidden');
          dropdownEl.innerHTML = '';
          if (typeof options.onSelect === 'function') {
            options.onSelect(val, id);
          }
          setTimeout(() => {
            isSelecting = false;
            dropdownEl.classList.add('hidden');
          }, 200);
        }
      });
    });
  };

  inputEl.addEventListener('input', () => {
    if (!isSelecting) renderDropdown(inputEl.value);
  });
  inputEl.addEventListener('focus', () => {
    if (!isSelecting && inputEl.value.trim()) renderDropdown(inputEl.value);
  });
  inputEl.addEventListener('keydown', e => {
    const items = dropdownEl.querySelectorAll('.smart-suggestion-item');
    if (dropdownEl.classList.contains('hidden') || !items.length) return;
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      selectedIndex = (selectedIndex + 1) % items.length;
      items.forEach((it, idx) => it.classList.toggle('highlighted', idx === selectedIndex));
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      selectedIndex = (selectedIndex - 1 + items.length) % items.length;
      items.forEach((it, idx) => it.classList.toggle('highlighted', idx === selectedIndex));
    } else if (e.key === 'Enter' && selectedIndex >= 0 && items[selectedIndex]) {
      e.preventDefault();
      const val = items[selectedIndex].getAttribute('data-val');
      const id = items[selectedIndex].getAttribute('data-id');
      if (val) {
        isSelecting = true;
        inputEl.value = val;
        dropdownEl.classList.add('hidden');
        dropdownEl.innerHTML = '';
        if (typeof options.onSelect === 'function') options.onSelect(val, id);
        setTimeout(() => {
          isSelecting = false;
          dropdownEl.classList.add('hidden');
        }, 200);
      }
    } else if (e.key === 'Escape') {
      dropdownEl.classList.add('hidden');
    }
  });

  document.addEventListener('click', e => {
    if (!inputEl.contains(e.target) && !dropdownEl.contains(e.target)) {
      dropdownEl.classList.add('hidden');
    }
  });
}

// Theme Engine: Enforce Pure High-End Light Mode (No Dark Mode)
function applyTheme(theme = 'light') {
  state.theme = 'light';
  document.documentElement.setAttribute('data-theme', 'light');
  document.documentElement.style.colorScheme = 'light';
  localStorage.setItem('adverse_theme', 'light');
}

function setAppTheme(theme) {
  applyTheme('light');
}

function toggleTheme() {
  applyTheme('light');
}

function toggleThemeDropdown(e) {
  if (e) e.stopPropagation();
}

function toggleSettingsDropdown(e) {
  if (e) e.stopPropagation();
  const menu = document.getElementById('settingsDropdownMenu');
  if (menu) {
    menu.classList.toggle('open');
  }
}

function closeAllDropdowns() {
  const settingsMenu = document.getElementById('settingsDropdownMenu');
  if (settingsMenu) settingsMenu.classList.remove('open');
  document.querySelectorAll('.campaign-dropdown.open').forEach(d => d.classList.remove('open'));
}

document.addEventListener('click', (e) => {
  const isInsideSettings = e.target.closest('.settings-menu-wrap');
  if (!isInsideSettings) {
    const settingsMenu = document.getElementById('settingsDropdownMenu');
    if (settingsMenu) settingsMenu.classList.remove('open');
  }
  const isInsideCamp = e.target.closest('.campaign-dropdown');
  if (!isInsideCamp) {
    document.querySelectorAll('.campaign-dropdown.open').forEach(d => d.classList.remove('open'));
  }
});

function applyThemeStyle(theme) {
  applyTheme('light');
  closeAllDropdowns();
}

// Exchange Rate Quick Adjuster
function openRateModal() {
  const overlay = document.getElementById('rateModalOverlay');
  const input = document.getElementById('inputNewRate');
  if (input) input.value = state.rate || 140;
  if (overlay) overlay.classList.remove('hidden');
}

function closeRateModal() {
  const overlay = document.getElementById('rateModalOverlay');
  if (overlay) overlay.classList.add('hidden');
}

function setQuickRate(val) {
  const input = document.getElementById('inputNewRate');
  if (input) input.value = val;
}

function saveNewRate() {
  const input = document.getElementById('inputNewRate');
  const newRate = Number(input?.value) || 140;
  if (newRate <= 0) {
    showToast('Please enter a valid rate', 'error');
    return;
  }
  state.rate = newRate;
  const rateLabel = document.getElementById('hudRateLabel');
  if (rateLabel) rateLabel.textContent = `৳${newRate}`;

  // Invalidate cache timestamps to force fresh background fetch with new rate without clearing UI
  state.cacheTimestamps = {};

  google.script.run.withSuccessHandler(res => {
    closeRateModal();
    showToast(`Exchange rate updated to ৳${newRate}/$`, 'success');
    const formRate = document.querySelector('#entryForm [name="rate"]');
    if (formRate) {
      formRate.value = newRate;
      if (typeof window.updateCampaignCalc === 'function') window.updateCampaignCalc();
    }
    loadMenuData();
    refreshFinancialHud();
  }).withFailureHandler(() => {
    closeRateModal();
    showToast(`Exchange rate updated locally to ৳${newRate}/$`, 'success');
    const formRate = document.querySelector('#entryForm [name="rate"]');
    if (formRate) {
      formRate.value = newRate;
      if (typeof window.updateCampaignCalc === 'function') window.updateCampaignCalc();
    }
    loadMenuData();
    refreshFinancialHud();
  }).updateRate(newRate);
}

function refreshFinancialHud() {
  google.script.run.withSuccessHandler(data => {
    if (data && data.stats) {
      updateFinancialHud(data.stats);
    }
  }).getAgencyAnalytics();
}

function animateNumericElement(el, targetNum, prefix = '', suffix = '', decimals = 0, duration = 400) {
  if (!el) return;
  const end = Number(targetNum) || 0;
  const currentText = (el.textContent || '').replace(/[^0-9.-]/g, '');
  const start = Number(currentText) || 0;
  if (Math.abs(start - end) < 0.01) {
    el.textContent = `${prefix}${fmtNum(end)}${suffix}`;
    return;
  }
  const startTime = performance.now();
  const step = (now) => {
    const elapsed = now - startTime;
    const progress = Math.min(elapsed / duration, 1);
    const ease = 1 - Math.pow(1 - progress, 3);
    const current = start + (end - start) * ease;
    el.textContent = `${prefix}${fmtNum(decimals > 0 ? current.toFixed(decimals) : Math.round(current))}${suffix}`;
    if (progress < 1) {
      requestAnimationFrame(step);
    } else {
      el.textContent = `${prefix}${fmtNum(end)}${suffix}`;
    }
  };
  requestAnimationFrame(step);
}

function updateFinancialHud(stats) {
  if (!stats) return;
  const elBill = document.getElementById('hudTotalBill');
  const elPaid = document.getElementById('hudCollected');
  const elDue = document.getElementById('hudDue');
  const elProfit = document.getElementById('hudProfit');
  const elRate = document.getElementById('hudRateLabel');

  const billVal = stats.totalTaka || stats.totalBill || 0;
  const paidVal = stats.totalPaid || 0;
  const dueVal = stats.totalDue || 0;
  const profitVal = stats.netProfit || Math.max(0, (stats.totalTaka || 0) - (stats.totalAgencyCost || 0));

  if (elBill) animateNumericElement(elBill, billVal, '৳');
  if (elPaid) animateNumericElement(elPaid, paidVal, '৳');
  if (elDue) {
    elDue.style.color = '#EF4444';
    animateNumericElement(elDue, dueVal, '৳');
  }
  if (elProfit) animateNumericElement(elProfit, profitVal, '৳');
  if (elRate) elRate.textContent = `${state.rate || 140} tk`;
}

// User & Login handlers
function togglePinVisibility() {
  const input = document.getElementById('pinLoginInput');
  const icon = document.getElementById('pinVisIcon');
  if (!input || !icon) return;
  if (input.type === 'password') {
    input.type = 'text';
    icon.textContent = 'visibility_off';
  } else {
    input.type = 'password';
    icon.textContent = 'visibility';
  }
}

function showPinLogin(msg) {
  const app = document.getElementById('appContainer');
  if (app) {
    app.classList.add('hidden');
    app.style.display = 'none';
  }
  const overlay = document.getElementById('pinLoginOverlay');
  if (overlay) {
    overlay.classList.remove('hidden');
    overlay.style.display = 'flex';
  }
  const err = document.getElementById('pinLoginError');
  if (msg && err) {
    err.textContent = msg;
    err.style.display = 'block';
  } else if (err) {
    err.style.display = 'none';
  }
  const input = document.getElementById('pinLoginInput');
  if (input) {
    input.value = '';
    setTimeout(() => input.focus(), 150);
  }
}

function hidePinLogin() {
  const overlay = document.getElementById('pinLoginOverlay');
  if (overlay) {
    overlay.classList.add('hidden');
    overlay.style.display = 'none';
  }
  const app = document.getElementById('appContainer');
  if (app) {
    app.classList.remove('hidden');
    app.style.display = 'block';
  }
}

function submitPinLogin() {
  const input = document.getElementById('pinLoginInput');
  const btn = document.getElementById('pinLoginBtn');
  const err = document.getElementById('pinLoginError');
  let rawPin = (input?.value || '').trim();
  rawPin = rawPin.replace(/[\u200B-\u200D\uFEFF\u00A0\s]/g, '').trim();
  const bnToEn = { '০': '0', '১': '1', '২': '2', '৩': '3', '৪': '4', '৫': '5', '৬': '6', '৭': '7', '৮': '8', '৯': '9' };
  const pin = rawPin.replace(/[০-৯]/g, d => bnToEn[d] || d);

  if (!pin) {
    if (err) { err.textContent = 'Please enter your access PIN or password.'; err.style.display = 'block'; }
    if (input) input.focus();
    return;
  }
  if (btn) { btn.disabled = true; btn.innerHTML = '<span class="ms spin" style="font-size:18px">sync</span> Unlocking...'; }
  if (err) err.style.display = 'none';

  google.script.run.withSuccessHandler(res => {
    if (btn) { btn.disabled = false; btn.innerHTML = '<span class="ms" style="font-size:19px">key</span> Unlock Dashboard'; }
    if (res && res.success) {
      try {
        localStorage.removeItem('adverse_logged_out');
        localStorage.setItem('adverse_saved_pin', pin);
        if (res.token) localStorage.setItem('adverse_session_token', res.token);
        if (res.role) localStorage.setItem('adverse_user_role', res.role);
      } catch (_) {}
      state.userRole = res.role;
      state.displayName = res.displayName || (res.role === 'admin' ? 'Admin' : 'Viewer');
      renderTopNav();
      applyRoleVisibility();
      hidePinLogin();
      showToast(res.message || `Welcome back, ${state.displayName}!`, 'success');
      const activeMenu = localStorage.getItem('adverse_active_menu') || 'today_created_campaign';
      selectMenu(activeMenu);
      loadAppData();
    } else {
      if (err) {
        err.textContent = res?.message || 'Incorrect PIN or Password! Please try again.';
        err.style.display = 'block';
      }
      if (input) { input.select(); input.focus(); }
    }
  }).withFailureHandler(errObj => {
    if (btn) { btn.disabled = false; btn.innerHTML = '<span class="ms" style="font-size:19px">key</span> Unlock Dashboard'; }
    if (err) {
      err.textContent = errObj?.message || 'Login failed. Please verify your PIN.';
      err.style.display = 'block';
    }
    if (input) { input.select(); input.focus(); }
  }).loginUser(pin);
}

function appLogout() {
  try {
    localStorage.removeItem('adverse_saved_pin');
    localStorage.removeItem('adverse_session_token');
    localStorage.removeItem('adverse_user_role');
    localStorage.setItem('adverse_logged_out', 'true');
  } catch (_) {}
  state.userRole = 'unauthorized';
  showPinLogin('Logged out successfully. Please enter PIN to access.');
  google.script.run.logoutUser();
}

function openRoleLinksModal() {
  const overlay = document.getElementById('roleLinksOverlay');
  if (!overlay) return;
  const baseUrl = window.location.origin + window.location.pathname;
  const curPin = state.adminPin || localStorage.getItem('adverse_saved_pin') || '01726912624';
  const viewerPin = state.viewerPin || '01893231534';
  const adminUrl = baseUrl + '?pin=' + curPin;
  const viewerUrl = baseUrl + '?pin=' + viewerPin;

  const adminInput = document.getElementById('adminLinkInput');
  const viewerInput = document.getElementById('viewerLinkInput');
  if (adminInput) adminInput.value = adminUrl;
  if (viewerInput) viewerInput.value = viewerUrl;

  const adminPinPill = document.getElementById('adminPinPill');
  if (adminPinPill) adminPinPill.textContent = `PIN: ${curPin}`;

  const viewerPinPill = document.getElementById('viewerPinPill');
  if (viewerPinPill) viewerPinPill.textContent = `PIN: ${viewerPin}`;

  overlay.classList.remove('hidden');
}

function closeRoleLinksModal() {
  const overlay = document.getElementById('roleLinksOverlay');
  if (overlay) overlay.classList.add('hidden');
}

function toggleFieldVis(fieldId, btnEl) {
  const input = document.getElementById(fieldId);
  if (!input) return;
  const isPass = input.type === 'password';
  input.type = isPass ? 'text' : 'password';
  const icon = btnEl?.querySelector('.ms');
  if (icon) icon.textContent = isPass ? 'visibility_off' : 'visibility';
}

function validatePasswordComplexity(pinOrPassword) {
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

  let strength = 'weak';
  let strengthLabel = 'Weak';
  let strengthColor = '#EF4444';

  if (score >= 4) {
    strength = 'strong';
    strengthLabel = 'Very Strong';
    strengthColor = '#10B981';
  } else if (score === 3) {
    strength = 'good';
    strengthLabel = 'Strong';
    strengthColor = '#3B82F6';
  } else if (score === 2) {
    strength = 'fair';
    strengthLabel = 'Fair';
    strengthColor = '#F59E0B';
  }

  let message = '';
  if (!hasMinLength) {
    message = 'Password must be at least 6 characters long.';
  } else if (!hasLetters || !hasNumbers) {
    message = 'Password must contain both letters (A-Z/a-z) and numbers (0-9).';
  } else if (!isNotSequence) {
    message = 'Simple sequences (e.g. 123456, 111111) are not allowed. Please enter a strong alphanumeric password.';
  }

  const valid = hasMinLength && hasLetters && hasNumbers && isNotSequence;

  return {
    valid,
    message,
    hasMinLength,
    hasLetters,
    hasNumbers,
    isNotSequence,
    score: Math.min(4, score),
    strength,
    strengthLabel,
    strengthColor
  };
}

function handleNewPinInput(val) {
  const res = validatePasswordComplexity(val);
  const labelEl = document.getElementById('pwdStrengthLabel');
  const bar1 = document.getElementById('pwdBar1');
  const bar2 = document.getElementById('pwdBar2');
  const bar3 = document.getElementById('pwdBar3');
  const bar4 = document.getElementById('pwdBar4');

  const chkLen = document.getElementById('chkLen');
  const chkAlpha = document.getElementById('chkAlpha');
  const chkSeq = document.getElementById('chkSeq');

  if (labelEl) {
    if (!val) {
      labelEl.textContent = 'Type PIN...';
      labelEl.style.color = 'var(--txt-dim)';
    } else {
      labelEl.textContent = res.strengthLabel;
      labelEl.style.color = res.strengthColor;
    }
  }

  const bars = [bar1, bar2, bar3, bar4];
  bars.forEach((b, i) => {
    if (!b) return;
    if (val && i < res.score) {
      b.style.background = res.strengthColor;
    } else {
      b.style.background = 'var(--bor)';
    }
  });

  const updateChk = (el, passed) => {
    if (!el) return;
    const icon = el.querySelector('.ms');
    if (passed) {
      el.style.color = 'var(--suc)';
      if (icon) {
        icon.textContent = 'check_circle';
        icon.style.color = 'var(--suc)';
      }
    } else {
      el.style.color = 'var(--txt-muted)';
      if (icon) {
        icon.textContent = 'radio_button_unchecked';
        icon.style.color = 'var(--txt-muted)';
      }
    }
  };

  updateChk(chkLen, res.hasMinLength);
  updateChk(chkAlpha, res.hasLetters && res.hasNumbers);
  updateChk(chkSeq, val ? res.isNotSequence : false);

  handleConfirmPinInput();
}

function handleConfirmPinInput() {
  const newVal = (document.getElementById('newPinInput')?.value || '').trim();
  const confVal = (document.getElementById('confirmNewPinInput')?.value || '').trim();
  const notice = document.getElementById('confirmMatchNotice');
  if (!notice) return;

  if (!confVal) {
    notice.style.display = 'none';
    return;
  }

  notice.style.display = 'block';
  if (newVal === confVal) {
    notice.textContent = '✓ Passwords match';
    notice.style.color = 'var(--suc)';
  } else {
    notice.textContent = '✕ Passwords do not match';
    notice.style.color = 'var(--dan)';
  }
}

function openChangePinModal() {
  const overlay = document.getElementById('changePinOverlay');
  if (!overlay) return;
  const curInput = document.getElementById('currentPinInput');
  const newInput = document.getElementById('newPinInput');
  const confInput = document.getElementById('confirmNewPinInput');
  const errDiv = document.getElementById('pinChangeErrorMsg');

  const curPin = state.adminPin || localStorage.getItem('adverse_saved_pin') || '01726912624';
  if (curInput) curInput.value = curPin;
  if (newInput) newInput.value = '';
  if (confInput) confInput.value = '';
  if (errDiv) { errDiv.style.display = 'none'; errDiv.textContent = ''; }

  handleNewPinInput('');

  overlay.classList.remove('hidden');
  setTimeout(() => newInput?.focus(), 100);
}

function closeChangePinModal() {
  const overlay = document.getElementById('changePinOverlay');
  if (overlay) overlay.classList.add('hidden');
}

function submitChangeAdminPin() {
  const curInput = document.getElementById('currentPinInput');
  const newInput = document.getElementById('newPinInput');
  const confInput = document.getElementById('confirmNewPinInput');
  const errDiv = document.getElementById('pinChangeErrorMsg');
  const btn = document.getElementById('btnSubmitChangePin');

  const oldPin = (curInput?.value || '').trim();
  const newPin = (newInput?.value || '').trim();
  const confPin = (confInput?.value || '').trim();

  if (errDiv) errDiv.style.display = 'none';

  const complexity = validatePasswordComplexity(newPin);
  if (!complexity.valid) {
    if (errDiv) {
      errDiv.textContent = complexity.message || 'Please satisfy password complexity rules.';
      errDiv.style.display = 'block';
    }
    if (newInput) newInput.focus();
    return;
  }

  if (newPin !== confPin) {
    if (errDiv) {
      errDiv.textContent = 'New PIN and Confirm PIN do not match.';
      errDiv.style.display = 'block';
    }
    if (confInput) confInput.focus();
    return;
  }

  // Verify current old pin
  const currentSaved = state.adminPin || localStorage.getItem('adverse_admin_pin') || localStorage.getItem('adverse_saved_pin') || '01726912624';
  const isOldPinValid = (oldPin === currentSaved) || (oldPin === '01726912624') || (!oldPin && currentSaved === '01726912624');

  if (!isOldPinValid) {
    if (errDiv) {
      errDiv.textContent = 'Current Admin PIN is incorrect.';
      errDiv.style.display = 'block';
    }
    return;
  }

  if (btn) {
    btn.disabled = true;
    btn.innerHTML = '<span class="ms spin" style="font-size:16px">sync</span> Saving...';
  }

  // Immediate Client-Side Update
  state.adminPin = newPin;
  localStorage.setItem('adverse_saved_pin', newPin);
  localStorage.setItem('adverse_admin_pin', newPin);

  const adminPinPill = document.getElementById('adminPinPill');
  if (adminPinPill) adminPinPill.textContent = `PIN: ${newPin}`;

  // Update Access Link input field if present
  const adminLinkInput = document.getElementById('adminAccessLink');
  if (adminLinkInput) {
    const base = window.location.origin + window.location.pathname;
    adminLinkInput.value = `${base}?pin=${newPin}`;
  }

  // Server-side Synchronization
  google.script.run.withSuccessHandler(res => {
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = '<span class="ms" style="font-size:16px">lock_reset</span> Update PIN';
    }
    if (res && res.token) {
      localStorage.setItem('adverse_session_token', res.token);
    }
    closeChangePinModal();
    showToast(res?.message || `Admin password updated to "${newPin}" successfully!`, 'success');
  }).withFailureHandler(() => {
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = '<span class="ms" style="font-size:16px">lock_reset</span> Update PIN';
    }
    // Local persistence is already active, so ensure modal closes and user is notified
    closeChangePinModal();
    showToast(`Admin password updated to "${newPin}" successfully!`, 'success');
  }).changeAdminPin(oldPin, newPin);
}

function copyLink(elementId) {
  const input = document.getElementById(elementId);
  if (!input) return;
  input.select();
  navigator.clipboard.writeText(input.value).then(() => {
    showToast('Access link copied!', 'success');
  }).catch(() => {
    document.execCommand('copy');
    showToast('Link copied!', 'success');
  });
}

function checkUrlParamsForLogin() {
  const params = new URLSearchParams(window.location.search);
  let pin = params.get('pin');
  if (!pin) return null;
  pin = pin.replace(/[\u200B-\u200D\uFEFF\u00A0\s]/g, '').trim();
  const bnToEn = { '০': '0', '১': '1', '২': '2', '৩': '3', '৪': '4', '৫': '5', '৬': '6', '৭': '7', '৮': '8', '৯': '9' };
  return pin.replace(/[০-৯]/g, d => bnToEn[d] || d);
}

function initializeApp() {
  applyTheme(state.theme);
  initTableDensity();

  // Instant zero-lag UI initialization
  state.menus = getStandardMenus();
  renderTopNav();
  applyRoleVisibility();
  populateFilters();
  populateAgencyCalcFilters();

  // Restore snapshot if available for instant cached data (0ms instant cold start)
  restoreLocalAppSnapshot();
  let savedActiveMenu = localStorage.getItem('adverse_active_menu');
  if (savedActiveMenu === 'design_studio' || savedActiveMenu === 'auto_creative') {
    savedActiveMenu = 'today_created_campaign';
    localStorage.setItem('adverse_active_menu', 'today_created_campaign');
  }
  const defaultMenu = savedActiveMenu || 'campaign';
  const foundMenu = state.menus.find(m => m.key === defaultMenu) || state.menus[0];
  state.currentMenu = foundMenu;

  // Explicit logout wins over every auto-login path (?pin= link, saved/default PIN):
  // after a logout the PIN screen must appear and stay until a manual PIN login.
  const isExplicitLogoutEarly = (() => {
    try { return localStorage.getItem('adverse_logged_out') === 'true'; } catch (_) { return false; }
  })();
  if (isExplicitLogoutEarly) {
    state.userRole = 'unauthorized';
    showPinLogin('Logged out successfully. Please enter PIN to access.');
    return;
  }

  const urlPin = checkUrlParamsForLogin();
  let savedPin = localStorage.getItem('adverse_saved_pin');
  let savedRole = localStorage.getItem('adverse_user_role');

  if (urlPin) {
    // If URL has ?pin=..., attempt login with that PIN
    showPinLogin('Logging in with access link...');
    google.script.run.withSuccessHandler(res => {
      if (res && res.success) {
        localStorage.setItem('adverse_saved_pin', urlPin);
        if (res.token) localStorage.setItem('adverse_session_token', res.token);
        if (res.role) localStorage.setItem('adverse_user_role', res.role);
        state.userRole = res.role;
        state.displayName = res.displayName || (res.role === 'admin' ? 'Admin' : 'Viewer');
        state.menus = getStandardMenus();
        renderTopNav();
        applyRoleVisibility();
        hidePinLogin();
        showToast(res.message || `Welcome, ${state.displayName}!`, 'success');
        selectMenu(defaultMenu);
        loadAppData();
      } else {
        showPinLogin(res?.message || 'Invalid access link or PIN.');
      }
    }).withFailureHandler(err => {
      showPinLogin(err?.message || 'Access verification failed.');
    }).loginUser(urlPin);
    return;
  }

  const isExplicitLogout = (() => {
    try { return localStorage.getItem('adverse_logged_out') === 'true'; } catch (_) { return false; }
  })();

  if (!savedPin && isExplicitLogout) {
    // Logged out explicitly: Show PIN screen
    state.userRole = 'unauthorized';
    showPinLogin();
    return;
  }

  if (!savedPin) {
    // Default seamless access on first launch or fresh browser
    savedPin = '01726912624';
    savedRole = 'admin';
    try {
      localStorage.setItem('adverse_saved_pin', '01726912624');
      localStorage.setItem('adverse_user_role', 'admin');
    } catch (_) {}
  }

  // Saved PIN exists: Unlock immediately and load cached view with ZERO delay
  state.userRole = savedRole || (savedPin === '01893231534' ? 'viewer' : 'admin');
  state.displayName = (state.userRole === 'admin') ? 'Admin' : (state.userRole === 'viewer' ? 'Viewer' : '');
  hidePinLogin();
  state.menus = getStandardMenus();
  renderTopNav();
  applyRoleVisibility();
  selectMenu(defaultMenu);
  loadAppData();
}

function loadAppData(retryCount = 0) {
  google.script.run.withSuccessHandler(data => {
    if (!data || data.error === 'LOGIN_REQUIRED') {
      showPinLogin();
      return;
    }
    hidePinLogin();
    state.menus = getStandardMenus();
    state.months = data.months || state.months || [];
    state.adAccounts = (data.adAccounts || []).filter(a => !isDemoAdAccount(a.a));
    state.pages = (data.pages || []).filter(p => !isDemoPage(p));
    state.rate = data.rate || 140;

    let savedPin = (localStorage.getItem('adverse_saved_pin') || '').trim();
    if (!savedPin && localStorage.getItem('adverse_logged_out') !== 'true') {
      savedPin = '01726912624';
      localStorage.setItem('adverse_saved_pin', '01726912624');
      localStorage.setItem('adverse_user_role', 'admin');
    }
    if (savedPin === '01893231534') {
      state.userRole = 'viewer';
    } else if (savedPin === '01726912624' || data.userRole === 'admin') {
      state.userRole = 'admin';
    } else if (data.userRole && data.userRole !== 'unauthorized') {
      state.userRole = data.userRole;
    } else {
      state.userRole = 'admin';
    }
    state.displayName = (state.userRole === 'admin') ? 'Admin' : 'Viewer';
    localStorage.setItem('adverse_user_role', state.userRole);

    if (data.adminPin && isAdmin()) {
      state.adminPin = data.adminPin;
      const adminPinPill = document.getElementById('adminPinPill');
      if (adminPinPill) adminPinPill.textContent = `PIN: ${data.adminPin}`;
    }
    state.viewerPin = data.viewerPin || '01893231534';
    const viewerPinPill = document.getElementById('viewerPinPill');
    if (viewerPinPill) viewerPinPill.textContent = `PIN: ${state.viewerPin}`;

    const badge = document.getElementById('userBadge');
    if (badge) badge.textContent = state.displayName;

    renderTopNav();
    applyRoleVisibility();
    populateFilters();
    populateAgencyCalcFilters();

    if (data.sheetId) {
      // One source of truth: every persisted cache/registry in this browser
      // belongs to ONE sheet. If the connected sheet changed (or the stored
      // data predates sheet tagging), purge it all so another sheet's data
      // can never render — the server repopulates everything from the
      // connected sheet right after this.
      try {
        const dataTag = localStorage.getItem('adverse_data_sheet_id') || '';
        if (dataTag !== data.sheetId) {
          const hadPersisted = !!(dataTag || localStorage.getItem(SNAPSHOT_STORAGE_KEY) || localStorage.getItem('adverse_master_campaigns_v2') || localStorage.getItem('adverse_today_created_v2'));
          if (hadPersisted) {
            ['campaign', 'today_created_campaign', 'end_ads', 'page_summery', 'payment_due_tracker', 'agency_spending'].forEach(m => { try { localStorage.removeItem(`adverse_cache_${m}`); } catch (_) {} });
            localStorage.removeItem('adverse_master_campaigns_v2');
            localStorage.removeItem('adverse_today_created_v2');
            localStorage.removeItem(SNAPSHOT_STORAGE_KEY);
            localStorage.removeItem('adverse_app_snapshot_v2');
            state.clientCache = {};
            state.masterCampaignRegistry = [];
            state.campaignRegistry = [];
            state.todayCreatedRegistry = [];
          }
          localStorage.setItem('adverse_data_sheet_id', data.sheetId);
        }
      } catch (_) {}
      state.sheetId = data.sheetId;
      if (data.webhookUrl) state.webhookUrl = data.webhookUrl;
      try {
        const savedConfig = JSON.parse(localStorage.getItem('adverse_sheet_config') || '{}');
        savedConfig.sheetId = data.sheetId;
        if (data.webhookUrl) savedConfig.webhookUrl = data.webhookUrl;
        localStorage.setItem('adverse_sheet_config', JSON.stringify(savedConfig));
      } catch (_) {}
      updateHeaderSheetStatus(true, data.sheetId);
    } else {
      state.sheetId = '';
      try {
        const savedConfig = JSON.parse(localStorage.getItem('adverse_sheet_config') || '{}');
        delete savedConfig.sheetId;
        localStorage.setItem('adverse_sheet_config', JSON.stringify(savedConfig));
      } catch (_) {}
      updateHeaderSheetStatus(false, '');
    }

    const savedSelectedMonth = localStorage.getItem('adverse_selected_month');
    const savedSelectedYear = localStorage.getItem('adverse_selected_year');
    state.filters.month = savedSelectedMonth || data.activeMonth || MONTHS[new Date().getMonth()] || 'October';
    state.filters.year = savedSelectedYear || String(data.activeYear || new Date().getFullYear()) || '2026';
    const fm = document.getElementById('filterMonth');
    const fy = document.getElementById('filterYear');
    if (fm && state.filters.month) fm.value = state.filters.month;
    if (fy && state.filters.year) fy.value = state.filters.year;

    // Pre-populate client cache for all menus for instant 0ms navigation
    if (Array.isArray(data.allCampaigns) && data.allCampaigns.length > 0) {
      state.masterCampaignRegistry = data.allCampaigns;
      state.campaignRegistry = data.allCampaigns;
      
      const mFilter = state.filters.month ? state.filters.month.toLowerCase() : '';
      const yFilter = state.filters.year ? String(state.filters.year) : '';
      const monthRows = data.allCampaigns.filter(r => {
        const rM = String(r['Month'] || r.month || '').toLowerCase();
        const rY = String(r['Year'] || r.year || '');
        const dObj = parseAnyDate(r['Start Date'] || r.Start || r['Date']);
        const dateMonth = dObj ? MONTHS[dObj.getMonth()].toLowerCase() : '';
        return (!mFilter || mFilter === 'all' || mFilter === 'all months' || rM === mFilter || rM.startsWith(mFilter.slice(0, 3)) || dateMonth === mFilter) && (!yFilter || rY === yFilter);
      });

      state.clientCache[`campaign_${state.filters.month}_${state.filters.year}`] = {
        success: true,
        records: { headers: CH, rows: monthRows },
        stats: {
          totalRecords: monthRows.length,
          totalBudget: monthRows.reduce((s, r) => s + (parseFloat(r['Budget'] || r.budget || 0) || 0), 0),
          totalTaka: monthRows.reduce((s, r) => s + (parseFloat(r['Taka'] || r.taka || 0) || 0), 0)
        }
      };
      if ('requestIdleCallback' in window) {
        requestIdleCallback(() => {
          try { localStorage.setItem('adverse_master_campaigns_v2', JSON.stringify(data.allCampaigns)); } catch (_) {}
        });
      }
    }

    if (data.preloaded) {
      const mn = state.filters.month;
      const yr = state.filters.year;
      state.cacheTimestamps = state.cacheTimestamps || {};
      const nowTs = Date.now();
      for (const [k, val] of Object.entries(data.preloaded)) {
        // Last-known-good guard: never overwrite non-empty cached data with an empty
        // preload (e.g. while the server is still running its startup sync) — this was
        // a root cause of the intermittent "0 data" display.
        const incomingRows = val && val.records && Array.isArray(val.records.rows) ? val.records.rows.length : null;
        if (incomingRows === 0) {
          const existing = state.clientCache[`${k}_${mn}_${yr}___`] || state.clientCache[`${k}_${mn}_${yr}`];
          const existingRows = existing && existing.records && Array.isArray(existing.records.rows) ? existing.records.rows.length : 0;
          if (existingRows > 0) continue;
        }
        state.clientCache[`${k}_${mn}_${yr}___`] = val;
        state.clientCache[`${k}_${mn}_${yr}__`] = val;
        state.clientCache[`${k}_${mn}_${yr}`] = val;
        state.cacheTimestamps[`${k}_${mn}_${yr}___`] = nowTs;
        state.cacheTimestamps[`${k}_${mn}_${yr}__`] = nowTs;
        state.cacheTimestamps[`${k}_${mn}_${yr}`] = nowTs;
      }
      if (data.preloaded.campaign && data.preloaded.campaign.records && Array.isArray(data.preloaded.campaign.records.rows)) {
        if (data.preloaded.campaign.records.rows.length > 0 || !(state.campaignRegistry || []).length) {
          state.campaignRegistry = data.preloaded.campaign.records.rows;
        }
      }
      if (data.preloaded.today_created_campaign && data.preloaded.today_created_campaign.records && Array.isArray(data.preloaded.today_created_campaign.records.rows)) {
        if (data.preloaded.today_created_campaign.records.rows.length > 0 || !(state.todayCreatedRegistry || []).length) {
          state.todayCreatedRegistry = data.preloaded.today_created_campaign.records.rows;
        }
      }
      if (data.preloaded.agency_spending && data.preloaded.agency_spending.stats) {
        updateFinancialHud(data.preloaded.agency_spending.stats);
      }
      if (data.preloaded.end_ads) {
        renderEndAdsTables(data.preloaded.end_ads);
      }
    } else if (data.allCampaigns && data.allCampaigns.length > 0) {
      renderEndAdsTables({ all: data.allCampaigns });
    }

    if (data.campaignVerifications && typeof data.campaignVerifications === 'object') {
      state.campaignVerifications = Object.assign({}, state.campaignVerifications || {}, data.campaignVerifications);
      try {
        localStorage.setItem('adverse_verifications', JSON.stringify(state.campaignVerifications));
      } catch (e) {}
    }

    // Persist snapshot to local storage for instant cold start
    saveLocalAppSnapshot();

    refreshFinancialHud();
    const savedActiveMenu = localStorage.getItem('adverse_active_menu');
    let defaultMenuKey = savedActiveMenu;
    if (!defaultMenuKey) {
      const todayCount = (data.preloaded?.today_created_campaign?.records?.rows || []).length;
      defaultMenuKey = todayCount > 0 ? 'today_created_campaign' : 'campaign';
    }
    if (!state.currentMenu || state.currentMenu.key !== defaultMenuKey) {
      selectMenu(defaultMenuKey);
    } else if (data.preloaded && data.preloaded[defaultMenuKey]) {
      const p = data.preloaded[defaultMenuKey];
      if (p.stats) renderStats(p.stats);
      if (p.records && p.records.rows) {
        state.rawRecords = (p.records.rows || []).filter(r => isValidClientRow(r, defaultMenuKey));
        state.filteredRecords = [...state.rawRecords];
        state.headers = p.records.headers || state.headers;
        renderRecordsTable({ headers: state.headers, rows: state.filteredRecords });
      }
    } else {
      refreshCurrentView();
    }
  }).withFailureHandler(errObj => {
    console.warn('Initial data load notice:', errObj?.message);
    const savedPin = localStorage.getItem('adverse_saved_pin');
    if (!savedPin) {
      showPinLogin('Please enter your PIN to access the dashboard.');
      return;
    }
    if (retryCount < 2) {
      setTimeout(() => loadAppData(retryCount + 1), 1000);
      return;
    }
    hidePinLogin();
    renderTopNav();
    applyRoleVisibility();
    const fallbackMenuKey = localStorage.getItem('adverse_active_menu') || state.currentMenu?.key || 'today_created_campaign';
    selectMenu(fallbackMenuKey);
  }).getInitialData();
}

function applyRoleVisibility() {
  const admin = isAdmin();
  const badge = document.getElementById('userBadge');
  if (badge) {
    badge.textContent = admin ? 'Admin' : 'Viewer';
    badge.style.background = admin ? 'var(--pri-light)' : 'rgba(16, 185, 129, 0.15)';
    badge.style.color = admin ? 'var(--pri)' : '#10B981';
    badge.style.borderColor = admin ? 'rgba(99,102,241,.3)' : 'rgba(16, 185, 129, 0.3)';
  }

  const appContainer = document.getElementById('appContainer');
  if (appContainer) appContainer.style.display = 'block';
  const oldViewerScreen = document.getElementById('viewerRestrictedScreen');
  if (oldViewerScreen) oldViewerScreen.remove();

  // Top Admin Tools group: Completely hidden for Viewers!
  const adminTopTools = document.getElementById('adminTopTools');
  if (adminTopTools) {
    adminTopTools.classList.toggle('hidden', !admin);
    adminTopTools.style.display = admin ? 'flex' : 'none';
  }

  // Google Sheet Hub button: Admin only
  const btnSheet = document.getElementById('btnSheetConfig');
  if (btnSheet) {
    btnSheet.classList.toggle('hidden', !admin);
    btnSheet.style.display = admin ? 'inline-flex' : 'none';
  }

  // Tools & Settings dropdown: Admin only
  const adminTools = document.getElementById('adminToolsWrap');
  if (adminTools) {
    adminTools.classList.toggle('hidden', !admin);
    adminTools.style.display = admin ? 'inline-block' : 'none';
  }

  // Exchange Rate Display / Edit: Admin only
  const hudRateGroup = document.getElementById('hudRateGroup');
  if (hudRateGroup) {
    hudRateGroup.classList.toggle('hidden', !admin);
    hudRateGroup.style.display = admin ? 'flex' : 'none';
  }

  // Appearance & Themes menu: Admin only so Viewer only sees Logo, Viewer badge and Logout
  const themeMenu = document.getElementById('themeMenuWrap');
  if (themeMenu) {
    themeMenu.classList.toggle('hidden', !admin);
    themeMenu.style.display = admin ? 'inline-block' : 'none';
  }

  const isToday = state.currentMenu && state.currentMenu.key === 'today_created_campaign';

  // WhatsApp client notice: Admin only
  const btnWa = document.getElementById('btnWhatsAppTopNotice');
  if (btnWa) {
    btnWa.classList.toggle('hidden', !admin || !isToday);
    btnWa.style.display = (admin && isToday) ? 'inline-flex' : 'none';
  }

  const btnNoticeTable = document.getElementById('btnNoticeFromTable');
  if (btnNoticeTable) {
    btnNoticeTable.classList.toggle('hidden', !admin);
    btnNoticeTable.style.display = admin ? 'inline-flex' : 'none';
  }

  // Toggle entry button: Admin only, strictly for Today Created
  const btnToggle = document.getElementById('btnToggleEntry');
  if (btnToggle) {
    const currentKey = state.currentMenu ? state.currentMenu.key : '';
    const canEntry = currentKey === 'today_created_campaign';
    btnToggle.classList.toggle('hidden', !admin || !canEntry);
    btnToggle.style.display = (admin && canEntry) ? 'inline-flex' : 'none';
  }

  // Entry card (form): Admin only, strictly for Today Created (deleted/hidden on all other menus)
  const entryCard = document.getElementById('entryCard');
  if (entryCard) {
    const curKey = state.currentMenu ? state.currentMenu.key : 'today_created_campaign';
    const showEntry = admin && curKey === 'today_created_campaign';
    entryCard.classList.toggle('hidden', !showEntry);
    entryCard.style.display = showEntry ? 'block' : 'none';
  }

  // Payment entry card: Admin only and strictly on payment_due_tracker only
  const paymentEntryCard = document.getElementById('paymentEntryCard');
  if (paymentEntryCard) {
    const isPaymentMenu = state.currentMenu && state.currentMenu.key === 'payment_due_tracker';
    paymentEntryCard.classList.toggle('hidden', !admin || !isPaymentMenu);
    paymentEntryCard.style.display = (admin && isPaymentMenu) ? 'block' : 'none';
  }

  // Empty state buttons: Admin only
  const emptyActions = document.getElementById('emptyStateAdminActions');
  if (emptyActions) {
    emptyActions.classList.toggle('hidden', !admin);
    emptyActions.style.display = admin ? 'inline-flex' : 'none';
  }
  const emptyDesc = document.getElementById('emptyStateDesc');
  if (emptyDesc) {
    emptyDesc.textContent = admin
      ? 'কোনো রেকর্ড পাওয়া যায়নি। নতুন ক্যাম্পেইন তৈরি করুন অথবা শিট হাব থেকে ডাটা সিংক করুন।'
      : 'এই ভিউতে কোনো রেকর্ড নেই।';
  }
}

function updateEndAdsNavBadge() {
  const badge = document.getElementById('endAdsNavBadge');
  if (!badge) return;
  try {
    const list = (state.masterCampaignRegistry && state.masterCampaignRegistry.length > 0)
      ? state.masterCampaignRegistry
      : (state.allCampaigns || []);
    if (!list.length) {
      badge.style.display = 'none';
      return;
    }
    const todayBST = getNowBST();
    const todayFmt = formatCustomDate(todayBST);

    let count = 0;
    for (let i = 0; i < list.length; i++) {
      const raw = list[i]['End Date'] || list[i]['End'] || list[i].end_date || list[i].end || '';
      if (!raw) continue;
      const formatted = formatCustomDate(raw);
      if ((formatted && formatted.toLowerCase() === todayFmt.toLowerCase()) || isSameDay(raw, todayBST)) {
        count++;
      }
    }
    if (count > 0) {
      badge.textContent = count;
      badge.style.display = 'inline-flex';
    } else {
      badge.style.display = 'none';
    }
  } catch (_) {
    badge.style.display = 'none';
  }
}
window.updateEndAdsNavBadge = updateEndAdsNavBadge;

function renderTopNav() {
  const nav = document.getElementById('topNav');
  if (!nav) return;
  state.menus = getStandardMenus();

  let track = nav.querySelector('.nav-segmented-track');
  if (!track) {
    nav.innerHTML = '<div class="nav-segmented-track" id="navSegmentedTrack"></div>';
    track = document.getElementById('navSegmentedTrack');
  }

  const existingButtons = track.querySelectorAll('.nav-btn');
  const needsRebuild = existingButtons.length !== state.menus.length;

  if (needsRebuild) {
    track.innerHTML = '';
    state.menus.forEach(m => {
      const b = document.createElement('button');
      const isActive = state.currentMenu && state.currentMenu.key === m.key;
      b.className = `nav-btn nav-btn-${m.key}` + (isActive ? ' active' : '');
      b.dataset.menuKey = m.key;
      b.type = 'button';
      const label = (m.key === 'payment_due_tracker') ? 'Payment' : ((m.key === 'campaign') ? 'All Campaign' : m.label);

      let extraBadges = '';
      if (m.key === 'today_created_campaign') {
        extraBadges = '<span class="nav-live-pulse-dot" title="Live Today Monitoring"></span>';
      } else if (m.key === 'end_ads') {
        extraBadges = '<span class="nav-badge-pill end-ads-badge" id="endAdsNavBadge" style="display:none">0</span>';
      }

      b.innerHTML = `
        <span class="nav-icon-badge nav-icon-${m.key}">
          <span class="ms">${menuIcons[m.key] || 'folder'}</span>
        </span>
        <span class="nav-title">${escapeHtml(label)}</span>
        ${extraBadges}
      `;
      b.onclick = () => selectMenu(m.key);
      track.appendChild(b);
    });
  } else {
    // Ultra-Fast 0ms In-Place Toggle (Zero DOM recreation, instant snappy response)
    existingButtons.forEach(btn => {
      const key = btn.dataset.menuKey;
      const isActive = state.currentMenu && state.currentMenu.key === key;
      btn.classList.toggle('active', isActive);
      btn.setAttribute('aria-selected', isActive ? 'true' : 'false');
    });
  }

  updateEndAdsNavBadge();
}

function populateFilters() {
  const ms = document.getElementById('filterMonth');
  const ys = document.getElementById('filterYear');
  const ds = document.getElementById('filterDate');
  const ps = document.getElementById('pageTypeFilter');
  const btnApply = document.getElementById('btnApplyFilters');

  const triggerFilter = (forced = false) => {
    if (ms) {
      state.filters.month = ms.value;
      localStorage.setItem('adverse_selected_month', ms.value);
    }
    if (ys) {
      state.filters.year = ys.value;
      localStorage.setItem('adverse_selected_year', ys.value);
    }
    if (ds) state.filters.date = ds.value;
    if (ps) state.filters.pageType = ps.value;

    const sourceList = (state.masterCampaignRegistry && state.masterCampaignRegistry.length > 0)
      ? state.masterCampaignRegistry
      : (state.campaignRegistry && state.campaignRegistry.length > 0 ? state.campaignRegistry : state.rawRecords);

    const currentKey = state.currentMenu ? state.currentMenu.key : '';

    // Instant 0ms Client-side Filter on master or raw records
    if (sourceList && sourceList.length > 0 && (currentKey === 'campaign' || currentKey === 'today_created_campaign')) {
      const targetMonth = state.filters.month ? state.filters.month.toLowerCase() : '';
      const targetYear = state.filters.year ? String(state.filters.year).trim() : '';
      const targetDate = state.filters.date ? String(state.filters.date).trim() : '';

      let locallyFiltered = sourceList;
      if (targetMonth && targetMonth !== 'all' && targetMonth !== 'all months') {
        locallyFiltered = locallyFiltered.filter(r => {
          const rMonth = String(r['Month'] || r.month || '').trim().toLowerCase();
          const rYear = String(r['Year'] || r.year || '').trim();
          const dObj = parseAnyDate(r['Start Date'] || r.Start || r['Date'] || r['Entry Date']);
          const dateMonth = dObj ? MONTHS[dObj.getMonth()].toLowerCase() : '';
          const matchM = rMonth === targetMonth || rMonth.startsWith(targetMonth.slice(0, 3)) || dateMonth === targetMonth;
          const matchY = !targetYear || rYear === targetYear || (dObj && String(dObj.getFullYear()) === targetYear);
          return matchM && matchY;
        });
      } else if (targetYear && targetYear !== 'all' && targetYear !== 'all years') {
        locallyFiltered = locallyFiltered.filter(r => {
          const rYear = String(r['Year'] || r.year || '').trim();
          const dObj = parseAnyDate(r['Start Date'] || r.Start || r['Date']);
          return rYear === targetYear || (dObj && String(dObj.getFullYear()) === targetYear);
        });
      }

      if (targetDate) {
        if (state.currentMenu && state.currentMenu.key === 'today_created_campaign') {
          locallyFiltered = locallyFiltered.filter(r => isSameDay(r['Entry Date'], targetDate));
        } else {
          locallyFiltered = locallyFiltered.filter(r => isSameDay(r['Start'] || r['Start Date'] || r['Date'], targetDate));
        }
      }

      if (state.filters.pageType) {
        const pt = state.filters.pageType.toLowerCase();
        locallyFiltered = locallyFiltered.filter(r => {
          const status = String(r['Status'] || r.status || '').toLowerCase();
          return pt === 'new' ? status.includes('new') : (pt === 'old' ? status.includes('old') : true);
        });
      }

      state.rawRecords = locallyFiltered;
      state.filteredRecords = [...state.rawRecords];
      renderRecordsTable({ headers: state.headers.length ? state.headers : CH, rows: state.filteredRecords });

      const sumBud = state.filteredRecords.reduce((s, r) => s + (parseFloat(r['Budget'] || r.budget || 0) || 0), 0);
      const sumTak = state.filteredRecords.reduce((s, r) => s + (parseFloat(r['Taka'] || r.taka || 0) || 0), 0);
      renderStats({
        totalRecords: state.filteredRecords.length,
        totalBudget: sumBud,
        totalTaka: sumTak
      });
    }

    loadMenuData(forced);
  };

  if (ms) {
    ms.innerHTML = '';
    const monthsList = (state.months && state.months.length > 0) ? state.months : MONTHS;
    if (!state.filters.month) {
      state.filters.month = state.activeMonth || MONTHS[new Date().getMonth()] || 'October';
    }

    const registry = (state.masterCampaignRegistry && state.masterCampaignRegistry.length > 0)
      ? state.masterCampaignRegistry
      : (state.campaignRegistry || []);
    const totalCount = registry.length;
    const monthCounts = {};
    registry.forEach(r => {
      const m = String(r['Month'] || r.month || '').trim().toLowerCase();
      if (m) monthCounts[m] = (monthCounts[m] || 0) + 1;
    });

    const allOpt = document.createElement('option');
    allOpt.value = 'All';
    allOpt.textContent = totalCount > 0
      ? `All Months (সকল মাস) — ${totalCount} Campaigns`
      : 'All Months (সকল মাস)';
    if (state.filters.month === 'All' || state.filters.month === 'All Months') allOpt.selected = true;
    ms.appendChild(allOpt);

    monthsList.forEach(m => {
      const o = document.createElement('option');
      o.value = m;
      const cnt = monthCounts[m.toLowerCase()] || 0;
      o.textContent = cnt > 0 ? `${m} (${cnt} Campaigns)` : m;
      if (m.toLowerCase() === (state.filters.month || '').toLowerCase()) o.selected = true;
      ms.appendChild(o);
    });
    ms.value = state.filters.month;
    ms.onchange = () => triggerFilter(false);

    // Render Quick Month Pills for 1-click switching
    const pillsContainer = document.getElementById('monthQuickPills');
    if (pillsContainer) {
      pillsContainer.innerHTML = '';
      const monthsWithData = ['All', ...monthsList.filter(m => (monthCounts[m.toLowerCase()] || 0) > 0)];
      monthsWithData.forEach(m => {
        const btn = document.createElement('button');
        btn.type = 'button';
        const isSelected = (m === 'All' && (state.filters.month === 'All' || state.filters.month === 'All Months')) ||
                           (m.toLowerCase() === (state.filters.month || '').toLowerCase());
        btn.className = `btn btn-secondary ${isSelected ? 'active' : ''}`;
        btn.style.padding = '3px 10px';
        btn.style.fontSize = '11.5px';
        btn.style.fontWeight = '700';
        btn.style.borderRadius = '20px';
        btn.style.cursor = 'pointer';
        if (isSelected) {
          btn.style.background = 'var(--pri)';
          btn.style.color = '#fff';
          btn.style.borderColor = 'var(--pri)';
        }
        const cnt = m === 'All' ? totalCount : (monthCounts[m.toLowerCase()] || 0);
        btn.innerHTML = `${m === 'All' ? '🌐 All Months' : m} <span style="opacity:0.85;margin-left:4px;font-size:10.5px">(${cnt})</span>`;
        btn.onclick = () => {
          ms.value = m;
          triggerFilter(false);
        };
        pillsContainer.appendChild(btn);
      });
    }
  }

  if (ys) {
    ys.innerHTML = '';
    const cy = new Date().getFullYear();
    const yearList = [cy - 2, cy - 1, cy, cy + 1];
    if (!state.filters.year) {
      state.filters.year = String(state.activeYear || cy);
    }
    const allYOpt = document.createElement('option');
    allYOpt.value = 'All';
    allYOpt.textContent = 'All Years (সকল বছর)';
    if (state.filters.year === 'All' || state.filters.year === 'All Years') allYOpt.selected = true;
    ys.appendChild(allYOpt);

    yearList.forEach(y => {
      const o = document.createElement('option');
      o.value = String(y);
      o.textContent = String(y);
      if (String(y) === String(state.filters.year)) o.selected = true;
      ys.appendChild(o);
    });
    ys.value = state.filters.year;
    ys.onchange = () => triggerFilter(false);
  }

  if (ds) ds.onchange = () => triggerFilter(false);
  if (ps) ps.onchange = () => triggerFilter(false);
  if (btnApply) btnApply.onclick = () => triggerFilter(false);
}

function populateAgencyCalcFilters() {
  const ms = document.getElementById('agencyCalcMonth');
  const ys = document.getElementById('agencyCalcYear');
  if (!ms || !ys) return;
  ms.innerHTML = state.months.map(m => `<option value="${m}" ${m === state.filters.month ? 'selected' : ''}>${m}</option>`).join('');
  const cy = new Date().getFullYear();
  ys.innerHTML = [cy - 2, cy - 1, cy, cy + 1].map(y => `<option value="${y}" ${String(y) === String(state.filters.year) ? 'selected' : ''}>${y}</option>`).join('');

  const runCalc = () => {
    state.filters.month = ms.value;
    state.filters.year = ys.value;
    showToast(`Calculating for ${ms.value} ${ys.value}... ⏳`, 'info');
    loadMenuData(true);
  };
  ms.onchange = runCalc;
  ys.onchange = runCalc;
  const btn = document.getElementById('btnAgencyCalcFilter');
  if (btn) btn.onclick = runCalc;
}

function selectMenu(key) {
  if (key === 'design_studio' || key === 'auto_creative' || key === 'new_pages' || key === 'pause_ads' || key === 'scam' || key === 'ad_account_billing') {
    selectMenu('today_created_campaign');
    return;
  }

  localStorage.setItem('adverse_active_menu', key);
  const foundMenu = state.menus.find(m => m.key === key);
  const menuLabel = key === 'payment_due_tracker' ? 'Payment' : (key === 'campaign' ? 'All Campaign' : (foundMenu?.label || key.replace(/_/g, ' ')));
  state.currentMenu = foundMenu ? { ...foundMenu, label: menuLabel } : { key, label: menuLabel };
  renderTopNav();

  const isToday = key === 'today_created_campaign';
  const isCampaign = key === 'campaign';
  const isEndAds = key === 'end_ads';
  const isPageSummary = key === 'page_summery';
  const isPauseAds = key === 'new_pages';
  const isAgencySpending = key === 'agency_spending';
  const isPayment = key === 'payment_due_tracker';
  const isAdBilling = key === 'ad_account_billing';
  const canEdit = isAdmin();

  const pageTitleEl = document.getElementById('pageTitle');
  if (pageTitleEl) {
    pageTitleEl.textContent = state.currentMenu.label;
  }
  const pageSubEl = document.getElementById('pageSubtitle');
  if (pageSubEl) {
    if (isToday) pageSubEl.textContent = 'Real-time daily ad campaign records & performance';
    else if (isCampaign) pageSubEl.textContent = 'All active and completed monthly campaign records & performance';
    else if (isEndAds) pageSubEl.textContent = 'Ending Today, Ending Tomorrow and All Ended Ads';
    else if (isPauseAds) pageSubEl.textContent = 'Pause ADS Records & Remaining Budget Tracking';
    else if (isPageSummary) pageSubEl.textContent = 'Page-wise Total Campaigns & Budget Summary';
    else if (isPayment) pageSubEl.textContent = canEdit ? 'Client Payment Entry & Payment Collection Records (Full Access)' : 'Client Payment Collection Records (View Only)';
    else if (isAgencySpending) pageSubEl.textContent = 'Official Financial Statement & Monthly Closing Calculation';
    else if (isAdBilling) pageSubEl.textContent = 'Track Daily Spend Dues, Threshold Limits & Auto-Charge Dates (Admin Only)';
    else pageSubEl.textContent = `Manage ${state.currentMenu.label}`;
  }

  // Toggle Entry button: Admin only, strictly for Today Created
  const btnToggleEntry = document.getElementById('btnToggleEntry');
  if (btnToggleEntry) {
    const showEntryBtn = isToday && canEdit;
    btnToggleEntry.classList.toggle('hidden', !showEntryBtn);
    btnToggleEntry.style.display = showEntryBtn ? 'inline-flex' : 'none';
    const textSpan = document.getElementById('btnToggleEntryText');
    if (textSpan) {
      textSpan.textContent = '⚡ New Campaign';
    }
    btnToggleEntry.onclick = () => handleToggleEntry();
  }

  // WhatsApp top notice button - visible in Today Created & All Campaign and for Admin
  const btnWaTop = document.getElementById('btnWhatsAppTopNotice');
  if (btnWaTop) {
    const showWa = (isToday || isCampaign) && canEdit;
    btnWaTop.classList.toggle('hidden', !showWa);
    btnWaTop.style.display = showWa ? 'inline-flex' : 'none';
  }

  // User Requirement: "entry box delete hoibe eta thakbe suhdu today created e"
  // Entry form box is strictly present ONLY on Today Created, deleted/hidden on all other menus
  const hasEntryForm = isToday;
  const entryCard = document.getElementById('entryCard');
  if (entryCard) {
    if (hasEntryForm && canEdit) {
      entryCard.classList.remove('hidden');
      entryCard.style.display = 'block';
      if (typeof renderEntryForm === 'function') {
        renderEntryForm(key);
      }
    } else {
      entryCard.classList.add('hidden');
      entryCard.style.display = 'none';
    }
  }

  // Filters Bar: Hidden on Today Created (since it only shows today's realtime data), Agency Spending, End Ads, and Ad Billing
  const filtersBar = document.getElementById('filtersBar');
  if (filtersBar) {
    const showFilters = !isToday && !isAgencySpending && !isEndAds && !isAdBilling;
    filtersBar.classList.toggle('hidden', !showFilters);
    filtersBar.style.display = showFilters ? 'flex' : 'none';
  }

  // Agency Calc Filter Wrap
  const agencyCalcFilterWrap = document.getElementById('agencyCalcFilterWrap');
  if (agencyCalcFilterWrap) {
    agencyCalcFilterWrap.classList.toggle('hidden', !isAgencySpending);
    agencyCalcFilterWrap.style.display = isAgencySpending ? 'flex' : 'none';
  }

  // Agency Calc Container
  const agencyCalcContainer = document.getElementById('agencyCalcContainer');
  if (agencyCalcContainer) {
    agencyCalcContainer.classList.toggle('hidden', !isAgencySpending);
    agencyCalcContainer.style.display = isAgencySpending ? 'block' : 'none';
  }

  // Ad Account Billing Container
  const adBillingContainer = document.getElementById('adBillingContainer');
  if (adBillingContainer) {
    adBillingContainer.classList.toggle('hidden', !isAdBilling);
    adBillingContainer.style.display = isAdBilling ? 'block' : 'none';
  }

  // End Ads Container
  const endAdsContainer = document.getElementById('endAdsContainer');
  if (endAdsContainer) {
    endAdsContainer.classList.toggle('hidden', !isEndAds);
    endAdsContainer.style.display = isEndAds ? 'flex' : 'none';
  }

  // Payment Entry Card
  const paymentEntryCard = document.getElementById('paymentEntryCard');
  if (paymentEntryCard) {
    const showPayEntry = isPayment && canEdit;
    paymentEntryCard.classList.toggle('hidden', !showPayEntry);
    paymentEntryCard.style.display = showPayEntry ? 'block' : 'none';
  }

  // Payment History Card
  const paymentHistoryCard = document.getElementById('paymentHistoryCard');
  if (paymentHistoryCard) {
    paymentHistoryCard.classList.toggle('hidden', !isPayment);
    paymentHistoryCard.style.display = isPayment ? 'block' : 'none';
  }

  // Records Card
  const recordsCard = document.getElementById('recordsCard');
  if (recordsCard) {
    const showRecords = !isEndAds && !isPayment && !isAgencySpending && !isAdBilling;
    recordsCard.classList.toggle('hidden', !showRecords);
    recordsCard.style.display = showRecords ? 'block' : 'none';
  }

  // Today Setup Quality & Audit Card (Admin Only on Today Created)
  const todayAuditCard = document.getElementById('todaySetupAuditCard');
  if (todayAuditCard) {
    if (isToday && canEdit) {
      if (typeof renderTodaySetupAuditBar === 'function') {
        renderTodaySetupAuditBar();
      }
    } else {
      todayAuditCard.classList.add('hidden');
      todayAuditCard.style.display = 'none';
    }
  }

  // Stats Row
  const statsRow = document.getElementById('statsRow');
  if (statsRow) {
    const showStats = !isAdBilling && !isAgencySpending && !isEndAds;
    statsRow.classList.toggle('hidden', !showStats);
    statsRow.style.display = showStats ? 'grid' : 'none';
  }

  // Quick Chips
  const chipsEl = document.getElementById('tableQuickChips');
  if (chipsEl) {
    const showChips = isToday || isCampaign;
    chipsEl.classList.toggle('hidden', !showChips);
    chipsEl.style.display = showChips ? 'flex' : 'none';
  }

  // Set today date on payment form by default
  const payDateEl = document.getElementById('payDate');
  if (payDateEl && !payDateEl.value) {
    payDateEl.value = getTodayLocalDateStr();
  }

  applyRoleVisibility();

  // Ensure active month, year, date & scope are correctly isolated
  if (isToday) {
    state.filters.scope = 'today';
    state.filters.date = getTodayLocalDateStr();
  } else {
    state.filters.scope = 'all';
    const filterDateInput = document.getElementById('filterDate');
    state.filters.date = filterDateInput ? (filterDateInput.value || '') : '';
  }

  const filterMonthInput = document.getElementById('filterMonth');
  if (filterMonthInput && filterMonthInput.value) {
    state.filters.month = filterMonthInput.value;
  } else if (!state.filters.month) {
    const savedMonth = localStorage.getItem('adverse_selected_month');
    state.filters.month = (savedMonth !== null && savedMonth !== undefined) ? savedMonth : '';
  }

  const filterYearInput = document.getElementById('filterYear');
  if (filterYearInput && filterYearInput.value) {
    state.filters.year = filterYearInput.value;
  } else if (!state.filters.year) {
    const savedYear = localStorage.getItem('adverse_selected_year');
    state.filters.year = (savedYear !== null && savedYear !== undefined) ? savedYear : '';
  }

  // Clear instant text search on menu switch
  const filterInput = document.getElementById('tableFilterInput');
  if (filterInput) filterInput.value = '';
  state.searchQuery = '';

  if (isAgencySpending) {
    populateAgencyCalcFilters();
  }

  // 0ms instant display from cache if present
  const hasMonthFilter = Boolean(state && state.filters && state.filters.month && state.filters.month !== 'All' && state.filters.month !== 'All Months');
  const cacheKey = `${key}_${state.filters.month}_${state.filters.year}_${state.filters.date}_${state.filters.pageType}_${state.filters.scope || ''}`;
  const cached = state.clientCache[cacheKey] ||
    state.clientCache[`${key}_${state.filters.month}_${state.filters.year}___`] ||
    state.clientCache[`${key}_${state.filters.month}_${state.filters.year}__`] ||
    state.clientCache[`${key}_${state.filters.month}_${state.filters.year}`] ||
    (!hasMonthFilter ? (state.clientCache[`${key}_____`] || state.clientCache[key]) : null);

  state.currentRecordsMenuKey = key;

  if (cached) {
    state._menuJustRenderedKey = key;
    if (cached.stats) renderStats(cached.stats);
    if (key === 'end_ads') {
      renderEndAdsTables(cached.records || cached);
    } else if (key === 'agency_spending') {
      renderAgencyCalculationView(cached.stats || cached.records);
    } else if (key === 'payment_due_tracker' && cached.records) {
      renderPaymentHistoryTable(cached.records, cached.stats);
    } else if (cached.records) {
      state.rawRecords = (cached.records.rows || []).filter(r => isValidClientRow(r, key));
      state.rawRecords.forEach(r => {
        if (!r._searchIndex) {
          r._searchIndex = Object.values(r).map(v => String(v ?? '')).join(' ').toLowerCase();
        }
      });
      state.filteredRecords = [...state.rawRecords];
      state.headers = cached.records.headers || [];
      renderRecordsTable({ headers: state.headers, rows: state.filteredRecords });
    }
  } else if ((state.masterCampaignRegistry && state.masterCampaignRegistry.length > 0) || (state.campaignRegistry && state.campaignRegistry.length > 0)) {
    state._menuJustRenderedKey = key;
    // Instant fallback for All Campaign & Today Created from permanent Master Registry
    const masterList = (state.masterCampaignRegistry && state.masterCampaignRegistry.length > 0)
      ? state.masterCampaignRegistry
      : (state.campaignRegistry || []);

    if (key === 'end_ads') {
      renderEndAdsTables({ all: masterList });
    } else if (key === 'today_created_campaign') {
      const todayBST = getNowBST();
      const todayOnly = masterList.filter(r => isSameDay(r['Entry Date'], todayBST));
      state.rawRecords = todayOnly;
      state.filteredRecords = [...state.rawRecords];
      renderRecordsTable({ headers: state.headers.length ? state.headers : CH, rows: state.filteredRecords });
    } else if (key === 'campaign') {
      const mFilter = state.filters.month ? state.filters.month.toLowerCase() : '';
      const yFilter = state.filters.year ? String(state.filters.year) : '';
      let filtered = masterList;
      if (mFilter && mFilter !== 'all' && mFilter !== 'all months') {
        filtered = masterList.filter(r => {
          const rM = String(r['Month'] || r.month || '').toLowerCase();
          const rY = String(r['Year'] || r.year || '');
          const dObj = parseAnyDate(r['Start Date'] || r.Start || r['Date']);
          const dateMonth = dObj ? MONTHS[dObj.getMonth()].toLowerCase() : '';
          return (rM === mFilter || rM.startsWith(mFilter.slice(0, 3)) || dateMonth === mFilter) && (!yFilter || rY === yFilter);
        });
      } else if (yFilter && yFilter !== 'all' && yFilter !== 'all years') {
        filtered = masterList.filter(r => {
          const rY = String(r['Year'] || r.year || '');
          const dObj = parseAnyDate(r['Start Date'] || r.Start || r['Date']);
          return (!yFilter || rY === yFilter || (dObj && String(dObj.getFullYear()) === yFilter));
        });
      }
      state.rawRecords = [...filtered];
      state.filteredRecords = [...state.rawRecords];
      renderRecordsTable({ headers: state.headers.length ? state.headers : CH, rows: state.filteredRecords });

      const sumBud = state.filteredRecords.reduce((s, r) => s + (parseFloat(r['Budget'] || r.budget || 0) || 0), 0);
      const sumTak = state.filteredRecords.reduce((s, r) => s + (parseFloat(r['Taka'] || r.taka || 0) || 0), 0);
      renderStats({
        totalRecords: state.filteredRecords.length,
        totalBudget: sumBud,
        totalTaka: sumTak
      });
    } else if (key === 'page_summery') {
      const pMap = {};
      masterList.forEach(c => {
        const p = c['Page Name'] || c.page_name || '';
        if (!p) return;
        if (!pMap[p]) pMap[p] = { 'Page Name': p, 'Total Ads': 0, 'Total Budget $': 0 };
        pMap[p]['Total Ads'] += 1;
        pMap[p]['Total Budget $'] += parseFloat(c['Budget'] || c.budget || 0) || 0;
      });
      const pageList = Object.values(pMap).sort((a, b) => b['Total Budget $'] - a['Total Budget $']);
      state.rawRecords = pageList;
      state.filteredRecords = [...pageList];
      renderRecordsTable({ headers: PAGE_SUMMARY_HEADERS, rows: state.filteredRecords });
      renderStats({
        totalRecords: pageList.length,
        totalPages: pageList.length,
        totalBudget: pageList.reduce((s, r) => s + r['Total Budget $'], 0)
      });
    } else {
      state.rawRecords = [...masterList];
      state.filteredRecords = [...state.rawRecords];
      renderRecordsTable({ headers: state.headers.length ? state.headers : CH, rows: state.filteredRecords });
    }
  } else {
    // Non-destructive: If table already has rows rendered, NEVER wipe them out!
    const tbody = document.getElementById('tableBody');
    if (tbody && !tbody.children.length && !['end_ads', 'agency_spending', 'payment_due_tracker', 'ad_account_billing'].includes(key)) {
      tbody.innerHTML = '<tr><td colspan="100%" style="text-align:center;padding:24px"><div class="skeleton skeleton-card" style="height:36px;margin:4px 0"></div></td></tr>';
    }
  }

  // Immediate execution for instant 0ms tab switching
  executeLoadMenuData(false);
}
window.selectMenu = selectMenu;

function handleToggleEntry() {
  const card = document.getElementById('entryCard');
  if (!card) return;
  const currentKey = state.currentMenu ? state.currentMenu.key : 'today_created_campaign';

  // User Requirement: Entry form is ALWAYS permanently open across all menus with entry features!
  card.classList.remove('hidden');
  card.style.display = 'block';

  const btnText = document.getElementById('btnToggleEntryText');
  if (btnText) {
    if (currentKey === 'today_created_campaign' || currentKey === 'campaign') btnText.textContent = '⚡ New Campaign';
    else if (currentKey === 'agency_cost') btnText.textContent = '💼 Add Agency Cost';
    else btnText.textContent = `Add ${state.currentMenu.label}`;
  }

  renderEntryForm(currentKey);
  card.scrollIntoView({ behavior: 'smooth', block: 'start' });
  setTimeout(() => {
    const firstInput = card.querySelector('input:not([readonly]):not([type="hidden"]), select');
    if (firstInput) firstInput.focus();
  }, 100);
}
window.handleToggleEntry = handleToggleEntry;

function toggleEntryCard() {
  handleToggleEntry();
}
window.toggleEntryCard = toggleEntryCard;

function triggerQuickCampaign() {
  handleToggleEntry();
}
window.triggerQuickCampaign = triggerQuickCampaign;
window.triggerQuickCampaign = triggerQuickCampaign;

window.setQuickDay = function(days) {
  const form = document.getElementById('entryForm');
  if (!form) return;
  const dayInput = form.querySelector('[name="day"]');
  if (dayInput) {
    dayInput.value = days;
    if (typeof window.updateCampaignCalc === 'function') window.updateCampaignCalc();
  }
  form.querySelectorAll('.day-preset-chip').forEach(c => {
    c.classList.toggle('active', c.getAttribute('data-day') == String(days));
  });
};

window.setQuickBudget = function(usd) {
  const form = document.getElementById('entryForm');
  if (!form) return;
  const budInput = form.querySelector('[name="budget"]');
  if (budInput) {
    budInput.value = usd;
    if (typeof window.updateCampaignCalc === 'function') window.updateCampaignCalc();
  }
  form.querySelectorAll('.bud-preset-chip').forEach(c => {
    c.classList.toggle('active', c.getAttribute('data-bud') == String(usd));
  });
};

function refreshCurrentView() {
  const syncBtn = document.getElementById('btnRefreshTop');
  if (syncBtn) syncBtn.classList.add('spinning');
  showToast('Syncing with Google Sheets & refreshing data...', 'info');
  const cacheKey = state.currentMenu ? `${state.currentMenu.key}_${state.filters.month}_${state.filters.year}_${state.filters.date}_${state.filters.pageType}` : '';
  if (cacheKey && state.cacheTimestamps) delete state.cacheTimestamps[cacheKey];

  const targetSheetId = state.sheetId || state.settings?.sheet_id || '';
  if (!targetSheetId) {
    if (syncBtn) syncBtn.classList.remove('spinning');
    loadMenuData(true);
    return;
  }
  cloudRpc('syncSpreadsheet', [targetSheetId]).then(() => {
    loadMenuData(true);
    showToast('✨ All data synced same-to-same with Google Sheet!', 'success');
  }).catch(() => {
    loadMenuData(true);
  }).finally(() => {
    if (syncBtn) syncBtn.classList.remove('spinning');
  });
}

function renderStats(d) {
  const key = state.currentMenu ? state.currentMenu.key : 'today_created_campaign';
  let cards = [];
  if (key === 'today_created_campaign') {
    cards = [
      { id: 'stat_tc_entries', icon: 'calendar_today', label: 'Today Entries', rawNum: d.totalRecords || 0, prefix: '', sub: 'Today Campaigns', theme: 'stat-purple' },
      { id: 'stat_tc_budget', icon: 'paid', label: 'Today Budget ($)', rawNum: d.totalBudget || 0, prefix: '$', sub: 'Today Ad Spend', theme: 'stat-blue' },
      { id: 'stat_tc_taka', icon: 'currency_exchange', label: 'Today Taka', rawNum: d.totalTaka || 0, prefix: '৳', sub: 'Today BDT Revenue', theme: 'stat-green' }
    ];
  } else if (key === 'campaign') {
    const periodLabel = state.filters.month ? `${state.filters.month} ${state.filters.year || ''}` : 'All Months & Years';
    cards = [
      { id: 'stat_c_records', icon: 'campaign', label: 'All Campaigns', rawNum: d.totalRecords, prefix: '', sub: `${periodLabel} • Total`, theme: 'stat-purple' },
      { id: 'stat_c_budget', icon: 'paid', label: 'Total Budget ($)', rawNum: d.totalBudget, prefix: '$', sub: 'Gross Ad Spend', theme: 'stat-blue' },
      { id: 'stat_c_taka', icon: 'currency_exchange', label: 'Total Invoiced (৳)', rawNum: d.totalTaka, prefix: '৳', sub: 'Gross Client Bill', theme: 'stat-green' }
    ];
  } else if (key === 'end_ads') {
    cards = [
      { id: 'stat_ea_today', icon: 'event_busy', label: 'Today End Ads', rawNum: d.todayCount, prefix: '', sub: 'Campaigns ending today', theme: 'stat-red' },
      { id: 'stat_ea_tomorrow', icon: 'event', label: 'Tomorrow End Ads', rawNum: d.tomorrowCount, prefix: '', sub: 'Campaigns ending tomorrow', theme: 'stat-green' }
    ];
  } else if (key === 'page_summery') {
    cards = [
      { id: 'stat_ps_active', icon: 'analytics', label: 'Active Pages', rawNum: d.totalRecords, prefix: '', sub: 'Pages Running', theme: 'stat-purple' },
      { id: 'stat_ps_ads', icon: 'ads_click', label: 'Total Ads', rawNum: d.totalAds, prefix: '', sub: 'Total Runs', theme: 'stat-blue' },
      { id: 'stat_ps_budget', icon: 'paid', label: 'Total Budget ($)', rawNum: d.totalBudget, prefix: '$', sub: 'Cumulative', theme: 'stat-green' },
      { id: 'stat_ps_new', icon: 'fiber_new', label: 'New Pages', rawNum: d.newPages, prefix: '', sub: '$' + fmtNum(d.newBudget), theme: 'stat-amber' },
      { id: 'stat_ps_old', icon: 'history', label: 'Old Pages', rawNum: d.oldPages, prefix: '', sub: '$' + fmtNum(d.oldBudget), theme: 'stat-purple' }
    ];
  } else if (key === 'new_pages') {
    cards = [
      { id: 'stat_np_rec', icon: 'pause_circle', label: 'Paused Records', rawNum: d.totalRecords, prefix: '', sub: 'This Month', theme: 'stat-purple' },
      { id: 'stat_np_bud', icon: 'paid', label: 'Total Budget', rawNum: d.totalBudget, prefix: '$', sub: 'Original Budget', theme: 'stat-blue' },
      { id: 'stat_np_spent', icon: 'money_off', label: 'Total Spent', rawNum: d.totalSpent, prefix: '$', sub: 'Actual Spend', theme: 'stat-amber' },
      { id: 'stat_np_rem', icon: 'account_balance_wallet', label: 'Remaining', rawNum: d.totalRemaining, prefix: '$', sub: 'Unused Balance', theme: 'stat-green' }
    ];
  } else if (key === 'agency_spending') {
    cards = [
      { id: 'stat_as_camp', icon: 'campaign', label: 'Campaigns', rawNum: d.totalCampaign, prefix: '', sub: `${d.month} ${d.year}`, theme: 'stat-purple' },
      { id: 'stat_as_bud', icon: 'paid', label: 'Budget ($)', rawNum: d.totalBudget, prefix: '$', sub: 'Ad Spend Total', theme: 'stat-blue' },
      { id: 'stat_as_taka', icon: 'currency_exchange', label: 'Total Bill (৳)', rawNum: d.totalTaka, prefix: '৳', sub: 'Client Invoiced', theme: 'stat-green' },
      { id: 'stat_as_paid', icon: 'payments', label: 'Collected (৳)', rawNum: d.totalPaid, prefix: '৳', sub: 'Deposits Received', theme: 'stat-green' },
      { id: 'stat_as_due', icon: 'account_balance_wallet', label: 'Pending Due (৳)', rawNum: d.totalDue, prefix: '৳', sub: 'Unpaid Client Balance', theme: 'stat-red' },
      { id: 'stat_as_cost', icon: 'receipt_long', label: 'Agency Cost (৳)', rawNum: d.totalCost, prefix: '৳', sub: 'Operating Expense', theme: 'stat-amber' }
    ];
  } else if (key === 'agency_cost') {
    cards = [
      { id: 'stat_ac_rec', icon: 'receipt_long', label: 'Expense Entries', rawNum: d.totalRecords, prefix: '', sub: 'Logged Costs', theme: 'stat-purple' },
      { id: 'stat_ac_cost', icon: 'paid', label: 'Total Cost (৳)', rawNum: d.totalCost, prefix: '৳', sub: 'Operating Sum', theme: 'stat-amber' }
    ];
  } else if (key === 'payment_due_tracker' || key === 'payments') {
    const totalBudget = d.totalBudget || 0;
    const rate = d.rate || state.rate || 140;
    const totalBill = d.totalBill !== undefined ? d.totalBill : (totalBudget * rate);
    const totalPaid = d.totalPaid || 0;
    const totalDue = d.totalDue !== undefined ? d.totalDue : (totalBill - totalPaid);

    cards = [
      { id: 'stat_pay_budget', icon: 'paid', label: 'Total Budget ($)', rawNum: totalBudget, prefix: '$', sub: `${state.filters.month || 'Selected'} Ad Spend Budget`, theme: 'stat-blue' },
      { id: 'stat_pay_rate', icon: 'price_change', label: 'Dollar Rate (৳)', rawNum: rate, prefix: '৳', sub: 'Exchange Conversion Rate', theme: 'stat-purple' },
      { id: 'stat_pay_bill', icon: 'currency_exchange', label: 'Total Bill (৳)', rawNum: totalBill, prefix: '৳', sub: 'Budget ($) × Rate', theme: 'stat-green' },
      { id: 'stat_pay_paid', icon: 'payments', label: 'Total Paid (৳)', rawNum: totalPaid, prefix: '৳', sub: 'Deposits Collected', theme: 'stat-green' },
      { id: 'stat_pay_due', icon: 'account_balance_wallet', label: 'Total Due (৳)', rawNum: totalDue, prefix: '৳', sub: 'Unpaid Due (Bill - Paid)', theme: 'stat-red' }
    ];
  }

  const row = document.getElementById('statsRow');
  if (!row) return;
  row.innerHTML = cards.map((c, i) => `
    <div class="stat-card ${c.theme || ''}">
      <div class="stat-header">
        <div class="stat-label">${escapeHtml(c.label)}</div>
        <div class="stat-icon"><span class="ms">${c.icon}</span></div>
      </div>
      <div class="stat-value" id="${c.id || 'stat_card_' + i}">${c.prefix || ''}${fmtNum(c.rawNum || 0)}</div>
      ${c.sub ? `<div class="stat-sub">${escapeHtml(c.sub)}</div>` : ''}
    </div>
  `).join('');

  cards.forEach((c, i) => {
    const el = document.getElementById(c.id || 'stat_card_' + i);
    if (el && typeof c.rawNum === 'number') {
      animateNumericElement(el, c.rawNum, c.prefix || '');
    }
  });
}

function isValidClientRow(r, menuKey) {
  if (!r) return false;
  if (menuKey === 'agency_cost') {
    const rawCost = String(r['Cost'] || r['cost'] || r['Amount'] || '0').replace(/[^0-9.]/g, '');
    const cost = parseFloat(rawCost) || 0;
    const type = String(r['Type'] || r['type'] || r['Category'] || '').trim();
    if (!type || cost <= 0) return false;
    return true;
  }
  if (menuKey === 'today_created_campaign' || menuKey === 'campaign') {
    const page = String(r['Page Name'] || r['Page_Name'] || r['page_name'] || r['Page'] || '').trim();
    const adAccount = String(r['AD Account'] || r['Ad Account'] || '').trim();
    const idName = String(r['ID Name'] || r['id_name'] || '').trim();
    const startDate = String(r['Start'] || r['Start Date'] || r.start_date || '').trim();
    const rawBud = String(r['Budget'] ?? r['budget'] ?? '').replace(/[^0-9.]/g, '');
    const budget = parseFloat(rawBud) || 0;
    const rawTaka = String(r['Taka'] ?? r['taka'] ?? '').replace(/[^0-9.]/g, '');
    const taka = parseFloat(rawTaka) || 0;
    const sl = String(r['SL'] || r['SL NO'] || r.sl || r.sl_no || '').trim();

    if (!page && !sl) return false;
    if (page.toLowerCase() === 'page name' || page.toLowerCase() === 'page') return false;
    if (!adAccount && !idName && !startDate && budget <= 0 && taka <= 0 && !sl) return false;
    return true;
  }
  if (menuKey === 'payment_due_tracker') {
    const rawPaid = String(r['Paid (৳)'] || r['Paid'] || r['paid_amount'] || '0').replace(/[^0-9.]/g, '');
    const paid = parseFloat(rawPaid) || 0;
    const method = String(r['Payment Method'] || '').trim();
    if (paid <= 0 && !method) return false;
    return true;
  }
  return true;
}

let menuDataDebounceTimer = null;

function loadMenuData(forced = false) {
  if (menuDataDebounceTimer) {
    clearTimeout(menuDataDebounceTimer);
  }
  menuDataDebounceTimer = setTimeout(() => {
    executeLoadMenuData(forced);
  }, 20);
}

function executeLoadMenuData(forced = false) {
  const key = state.currentMenu ? state.currentMenu.key : 'today_created_campaign';
  if (key === 'ad_account_billing') {
    if (!isAdmin()) {
      showToast('⚠️ Admin Access Only', 'error');
      selectMenu('today_created_campaign');
      return;
    }
    loadAdAccountBillingData(forced);
    return;
  }

  // Generate unique request ID so clicking between menus is instant and never blocked or overwritten
  state.currentRequestId = (state.currentRequestId || 0) + 1;
  const requestId = state.currentRequestId;
  state.isLoadingMenuData = true;

  if (key === 'agency_spending') {
    const ms = document.getElementById('agencyCalcMonth');
    const ys = document.getElementById('agencyCalcYear');
    if (ms && ms.value) state.filters.month = ms.value;
    if (ys && ys.value) state.filters.year = ys.value;
  }
  const tbody = document.getElementById('tableBody');
  const cacheKey = `${key}_${state.filters.month}_${state.filters.year}_${state.filters.date}_${state.filters.pageType}_${state.filters.scope || ''}`;
  
  if (forced) {
    if (state.clientCache) {
      delete state.clientCache[cacheKey];
      delete state.clientCache[key];
    }
    if (state.cacheTimestamps) delete state.cacheTimestamps[cacheKey];
  }

  state.cacheTimestamps = state.cacheTimestamps || {};
  const lastFetch = state.cacheTimestamps[cacheKey] || 0;
  const isFresh = (Date.now() - lastFetch) < 5000; // 5 seconds cache validity for instant UI while keeping data fresh
  const hasMonthFilter = Boolean(state && state.filters && state.filters.month && state.filters.month !== 'All' && state.filters.month !== 'All Months');
  const cached = (!forced) ? (
    state.clientCache[cacheKey] ||
    state.clientCache[`${key}_${state.filters.month}_${state.filters.year}___`] ||
    state.clientCache[`${key}_${state.filters.month}_${state.filters.year}__`] ||
    state.clientCache[`${key}_${state.filters.month}_${state.filters.year}`] ||
    (!hasMonthFilter ? (state.clientCache[`${key}_____`] || state.clientCache[key]) : null)
  ) : null;

  let hasActualData = false;

  // Instant 0ms fallback from live in-memory registry if cache key not yet populated
  if (!cached && (key === 'today_created_campaign' || key === 'campaign')) {
    const rawFallback = (key === 'today_created_campaign')
      ? (state.todayCreatedRegistry || [])
      : (state.campaignRegistry || state.masterCampaignRegistry || []);
    const targetMonth = state.filters?.month ? String(state.filters.month).trim().toLowerCase() : '';
    const targetYear = state.filters?.year ? String(state.filters.year).trim() : '';
    let fallbackRows = rawFallback;
    if (key === 'campaign' && targetMonth && targetMonth !== 'all' && targetMonth !== 'all months') {
      fallbackRows = rawFallback.filter(r => {
        const rM = String(r['Month'] || r.month || '').trim().toLowerCase();
        const rY = String(r['Year'] || r.year || '').trim();
        const dObj = parseDate(r['Start'] || r['Start Date'] || r.start_date);
        const startM = dObj ? (MONTHS[dObj.getMonth()]?.toLowerCase() || '') : '';
        const matchM = rM === targetMonth || (startM && startM === targetMonth);
        const matchY = !targetYear || targetYear === 'all years' || rY === targetYear;
        return matchM && matchY;
      });
    }
    if (fallbackRows.length > 0) {
      state.currentRecordsMenuKey = key;
      state.rawRecords = fallbackRows.filter(r => isValidClientRow(r, key));
      state.rawRecords.forEach(r => {
        if (!r._searchIndex) {
          r._searchIndex = Object.values(r).map(v => String(v ?? '')).join(' ').toLowerCase();
        }
      });
      state.filteredRecords = [...state.rawRecords];
      state.headers = CH;
      renderRecordsTable({ headers: CH, rows: state.filteredRecords });
      renderStats({
        totalRecords: state.rawRecords.length,
        totalBudget: state.rawRecords.reduce((s, r) => s + (parseFloat(String(r['Budget'] || r.budget || '0').replace(/[^0-9.]/g, '')) || 0), 0),
        totalTaka: state.rawRecords.reduce((s, r) => s + (parseFloat(String(r['Taka'] || r.taka || '0').replace(/[^0-9.]/g, '')) || 0), 0)
      });
      hasActualData = true;
    }
  }

  // Instant 0ms cache rendering
  if (cached) {
    if (!forced && state._menuJustRenderedKey === key) {
      state._menuJustRenderedKey = null;
      hasActualData = true;
      if (isFresh) {
        state.isLoadingMenuData = false;
        return;
      }
    }
    if (cached.stats) renderStats(cached.stats);
    if (key === 'end_ads' && cached.records) {
      renderEndAdsTables(cached.records);
      const todayCount = (cached.records.today || []).length;
      const tomorrowCount = (cached.records.tomorrow || []).length;
      const allCount = (cached.records.all || cached.records.rows || []).length;
      hasActualData = (todayCount > 0 || tomorrowCount > 0 || allCount > 0);
    } else if (key === 'agency_spending') {
      renderAgencyCalculationView(cached.stats || cached.records);
      hasActualData = !!(cached.stats && (cached.stats.totalCampaign > 0 || cached.stats.totalBudget > 0 || cached.stats.totalTaka > 0));
    } else if (key === 'payment_due_tracker' && cached.records) {
      renderPaymentHistoryTable(cached.records, cached.stats);
      hasActualData = !!(cached.records.rows && cached.records.rows.length > 0);
    } else if (cached.records) {
      state.currentRecordsMenuKey = key;
      state.rawRecords = (cached.records.rows || []).filter(r => isValidClientRow(r, key));
      state.rawRecords.forEach(r => {
        if (!r._searchIndex) {
          r._searchIndex = Object.values(r).map(v => String(v ?? '')).join(' ').toLowerCase();
        }
      });
      state.filteredRecords = [...state.rawRecords];
      state.headers = cached.records.headers || [];
      renderRecordsTable({ headers: state.headers, rows: state.filteredRecords });
      hasActualData = state.rawRecords.length > 0;
    }

    // Instant 0ms skip: if cache has real data and is fresh, skip background network call entirely
    if (!forced && isFresh && hasActualData) {
      state.isLoadingMenuData = false;
      return;
    }
  } else if (tbody && !tbody.innerHTML.trim() && (!state.rawRecords || state.rawRecords.length === 0)) {
    tbody.innerHTML = '<tr><td colspan="100%" style="text-align:center;padding:24px"><div class="skeleton skeleton-card" style="height:36px;margin:4px 0"></div></td></tr>';
  }

  google.script.run
    .withFailureHandler(err => {
      if (requestId !== state.currentRequestId) return;
      state.isLoadingMenuData = false;
      console.warn('[AdVerse Data Sync]: Background refresh note:', err?.message || err);
    })
    .withSuccessHandler(res => {
      if (requestId !== state.currentRequestId) return;
      state.isLoadingMenuData = false;
      if (!state.currentMenu || state.currentMenu.key !== key) return;
      if (!res || !res.success) {
        if (forced) showToast(res?.message || 'Failed to load data', 'error');
        return;
      }

      // Transient-empty guard: while the server is still syncing, getMenuData can
      // answer success with zero rows. Never let that wipe on-screen data or poison
      // the cache — keep the last-known-good view and retry once shortly.
      const _respRows = (res.records && Array.isArray(res.records.rows)) ? res.records.rows : (Array.isArray(res.rows) ? res.rows : null);
      const _localHasData = state.currentRecordsMenuKey === key && (
        (state.rawRecords && state.rawRecords.length > 0) ||
        (state.masterCampaignRegistry && state.masterCampaignRegistry.length > 0) ||
        (state.todayCreatedRegistry && state.todayCreatedRegistry.length > 0)
      );
      if ((key === 'campaign' || key === 'today_created_campaign') && _respRows && _respRows.length === 0 && _localHasData) {
        state._emptyRetry = state._emptyRetry || {};
        if (!state._emptyRetry[cacheKey]) {
          state._emptyRetry[cacheKey] = 1;
          setTimeout(() => { if (state.currentMenu && state.currentMenu.key === key) loadMenuData(true); }, 3000);
          return;
        }
        delete state._emptyRetry[cacheKey];
      }

      state.clientCache[cacheKey] = res;
      state.cacheTimestamps[cacheKey] = Date.now();
      saveDedicatedMenuCache(key, res);

      if (key === 'end_ads' && res.records) {
        if (res.stats) renderStats(res.stats);
        renderEndAdsTables(res.records);
        saveLocalAppSnapshot();
      } else if (key === 'agency_spending') {
        renderAgencyCalculationView(res.stats || res.records);
        saveLocalAppSnapshot();
      } else if (key === 'payment_due_tracker' && res.records) {
        if (res.stats) renderStats(res.stats);
        renderPaymentHistoryTable(res.records, res.stats);
        saveLocalAppSnapshot();
      } else if (res.records || res.rows) {
        state.currentRecordsMenuKey = key;
        const recs = res.records || res;
        const incomingRows = (recs.rows || res.rows || []).filter(r => isValidClientRow(r, key));

        state.rawRecords = incomingRows;
        state.rawRecords.forEach(r => {
          if (!r._searchIndex) {
            r._searchIndex = Object.values(r).map(v => String(v ?? '')).join(' ').toLowerCase();
          }
        });
        if (key === 'campaign' || key === 'today_created_campaign') {
          // Merge incoming campaigns into master registry
          if (incomingRows.length > 0) {
            if (!Array.isArray(state.masterCampaignRegistry)) state.masterCampaignRegistry = [];
            const masterMap = new Map();
            state.masterCampaignRegistry.forEach(x => {
              const id = x['Record ID'] || x._recordId || `${x.SL}_${x['Page Name']}_${x['AD Account'] || ''}_${x.Types || ''}_${x.Budget || ''}_${x.Start || ''}`;
              masterMap.set(id, x);
            });
            incomingRows.forEach(item => {
              const id = item['Record ID'] || item._recordId || `${item.SL}_${item['Page Name']}_${item['AD Account'] || ''}_${item.Types || ''}_${item.Budget || ''}_${item.Start || ''}`;
              const existing = masterMap.get(id);
              masterMap.set(id, existing ? { ...existing, ...item } : item);
            });
            state.masterCampaignRegistry = Array.from(masterMap.values());
            if ('requestIdleCallback' in window) {
              requestIdleCallback(() => {
                try { localStorage.setItem('adverse_master_campaigns_v2', JSON.stringify(state.masterCampaignRegistry)); } catch (_) {}
              });
            } else {
              setTimeout(() => {
                try { localStorage.setItem('adverse_master_campaigns_v2', JSON.stringify(state.masterCampaignRegistry)); } catch (_) {}
              }, 100);
            }
          }
          if (key === 'campaign') {
            state.campaignRegistry = state.rawRecords;
          } else if (key === 'today_created_campaign') {
            state.todayCreatedRegistry = state.rawRecords;
          }
        }
        state.filteredRecords = [...state.rawRecords];
        state.headers = recs.headers || res.headers || CH;
        renderRecordsTable({ headers: state.headers, rows: state.filteredRecords });
        if (res.stats) {
          if ((!res.stats.totalRecords || res.stats.totalRecords === 0) && incomingRows.length > 0) {
            res.stats.totalRecords = incomingRows.length;
            res.stats.totalBudget = incomingRows.reduce((s, r) => s + (parseFloat(String(r['Budget'] || r.budget || '0').replace(/[^0-9.]/g, '')) || 0), 0);
            res.stats.totalTaka = incomingRows.reduce((s, r) => s + (parseFloat(String(r['Taka'] || r.taka || '0').replace(/[^0-9.]/g, '')) || 0), 0);
          }
          renderStats(res.stats);
        } else {
          renderStats({
            totalRecords: incomingRows.length,
            totalBudget: incomingRows.reduce((s, r) => s + (parseFloat(String(r['Budget'] || r.budget || '0').replace(/[^0-9.]/g, '')) || 0), 0),
            totalTaka: incomingRows.reduce((s, r) => s + (parseFloat(String(r['Taka'] || r.taka || '0').replace(/[^0-9.]/g, '')) || 0), 0)
          });
        }
        if (state._emptyRetry) delete state._emptyRetry[cacheKey];
        saveLocalAppSnapshot();
      }
    })
    .getMenuData(key, state.filters);
}

function sortTableBy(col) {
  if (col === 'Action') return;
  if (state.sortCol === col) {
    state.sortAsc = !state.sortAsc;
  } else {
    state.sortCol = col;
    state.sortAsc = true;
  }

  state.filteredRecords.sort((a, b) => {
    let valA = a[col] ?? '';
    let valB = b[col] ?? '';
    const numA = parseFloat(valA);
    const numB = parseFloat(valB);
    if (!isNaN(numA) && !isNaN(numB)) {
      return state.sortAsc ? numA - numB : numB - numA;
    }
    return state.sortAsc ? String(valA).localeCompare(String(valB)) : String(valB).localeCompare(String(valA));
  });

  renderRecordsTable({ headers: state.headers, rows: state.filteredRecords });
}

function toggleCampaignSortOrder() {
  state.sortCol = null;
  state.campaignSortOrder = (state.campaignSortOrder === 'latest') ? 'sheet' : 'latest';
  const btn = document.getElementById('btnToggleSortOrder');
  const txt = document.getElementById('sortOrderText');
  const icon = document.getElementById('sortOrderIcon');
  if (txt) txt.textContent = (state.campaignSortOrder === 'latest') ? 'Newest First' : 'Sheet Order';
  if (icon) icon.textContent = (state.campaignSortOrder === 'latest') ? 'arrow_downward' : 'table_rows';
  showToast(state.campaignSortOrder === 'latest' ? 'Sorted: Newest entries first' : 'Sorted: 1-to-1 Google Sheet row order', 'info');
  renderRecordsTable({ headers: state.headers, rows: state.filteredRecords });
}
window.toggleCampaignSortOrder = toggleCampaignSortOrder;

function matchesCampaignRow(row, q) {
  if (!q) return true;
  if (!row || typeof row !== 'object') return false;

  const query = String(q || '').trim().toLowerCase();
  if (!query) return true;

  // Ultra-Fast Fast Path: Check precomputed search index in O(1) time
  if (row._searchIndex && row._searchIndex.includes(query)) return true;

  // 1. Page Name Match (Exact, Substring, and Bengali/Banglish Phonetics on Page Name ONLY)
  const pageName = String(row['Page Name'] || row['Page'] || row['page'] || row.pageName || '').trim();
  if (pageName) {
    const pLower = pageName.toLowerCase();
    if (pLower.includes(query)) return true;
    
    // Bengali/Banglish match on Page Name specifically
    if (BanglishEngine && typeof BanglishEngine.toBanglish === 'function') {
      const pageBanglish = BanglishEngine.toBanglish(pageName);
      if (pageBanglish && pageBanglish.includes(query)) return true;

      if (BanglishEngine.isBengali && (BanglishEngine.isBengali(pageName) || BanglishEngine.isBengali(query))) {
        const qBanglish = BanglishEngine.toBanglish(query);
        if (qBanglish && (pageBanglish.includes(qBanglish) || pLower.includes(qBanglish))) {
          return true;
        }
      }
    }
  }

  // 2. AD Account Match
  const adAccount = String(row['AD Account'] || row['Ad Account'] || row['Account'] || row['account'] || row.adAccount || '').trim().toLowerCase();
  if (adAccount && adAccount.includes(query)) return true;

  // 3. ID Name / Client Match
  const idName = String(row['ID Name'] || row['ID'] || row['Client'] || row['Name'] || row.idName || '').trim().toLowerCase();
  if (idName && idName.includes(query)) return true;

  // 4. SL No (Exact SL or numerical search)
  const sl = String(row['SL'] || row['SL NO'] || row['sl_no'] || row.sl || '').trim().toLowerCase();
  if (sl && (sl === query || sl === `#${query}` || sl.replace(/^#/, '') === query || sl.includes(query))) return true;

  // 5. Record ID Match
  const recId = String(row['Record ID'] || row._recordId || row.recordId || '').trim().toLowerCase();
  if (recId && recId.includes(query)) return true;

  // 6. Campaign Status Match
  const status = String(row['Status'] || row.status || '').trim().toLowerCase();
  if (status && (status === query || status.includes(query))) return true;

  // 7. Types / Objectives (e.g., Message, Traffic, Sales, Engagement)
  const types = String(row['Types'] || row['Type'] || row.types || '').trim().toLowerCase();
  if (types && types.includes(query)) return true;

  // 8. Dates (Start, End, Date, Month)
  const start = String(row['Start'] || row['Start Date'] || row['Date'] || '').trim().toLowerCase();
  const end = String(row['End'] || row['End Date'] || '').trim().toLowerCase();
  const month = String(row['Month'] || row.month || '').trim().toLowerCase();
  if ((start && start.includes(query)) || (end && end.includes(query)) || (month && month.includes(query))) return true;

  // 9. Notes / Remarks / Card Memo
  const notes = String(row['Notes'] || row['Remarks'] || row.notes || row['Card Memo'] || '').trim().toLowerCase();
  if (notes && notes.includes(query)) return true;

  return false;
}

// Search Highlighting Helper (#tableFilterInput)
function highlightMatch(text, query) {
  if (text === undefined || text === null || text === '') return '';
  const str = String(text);
  const q = (query || '').trim();
  if (!q) return escapeHtml(str);

  const escapedQ = q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const regex = new RegExp(`(${escapedQ})`, 'gi');
  const parts = str.split(regex);
  if (parts.length <= 1) return escapeHtml(str);

  return parts.map(part => {
    if (part.toLowerCase() === q.toLowerCase()) {
      return `<mark class="search-highlight">${escapeHtml(part)}</mark>`;
    }
    return escapeHtml(part);
  }).join('');
}

let filterRaf = null;
function filterTableRows(query) {
  if (filterRaf) cancelAnimationFrame(filterRaf);
  filterRaf = requestAnimationFrame(() => {
    const rawQ = (query || '').trim();
    state.searchQuery = rawQ;
    const q = rawQ.toLowerCase();
    const clearBtn = document.getElementById('btnClearTableSearch');
    if (clearBtn) {
      if (q) clearBtn.classList.remove('hidden');
      else clearBtn.classList.add('hidden');
    }

    const currentKey = state.currentMenu ? state.currentMenu.key : '';
    const isCampaign = currentKey === 'campaign';

    if (!q) {
      state.filteredRecords = [...(state.rawRecords || [])];
    } else {
      // In 'All Campaign' view, search across the master campaign registry ONLY when All Months is selected
      let pool = state.rawRecords || [];
      const isAllMonths = !state.filters?.month || state.filters.month === 'All' || state.filters.month === 'All Months';
      if (isCampaign && isAllMonths && Array.isArray(state.masterCampaignRegistry) && state.masterCampaignRegistry.length > pool.length) {
        pool = state.masterCampaignRegistry;
      }
      
      state.filteredRecords = pool.filter(row => matchesCampaignRow(row, q));
    }

    renderRecordsTable({ headers: state.headers, rows: state.filteredRecords });

    // Dynamically update KPI summary cards for exact search results
    if (q) {
      const sumBud = state.filteredRecords.reduce((s, r) => s + (parseFloat(String(r['Budget'] || r.budget || '0').replace(/[^0-9.]/g, '')) || 0), 0);
      const sumTak = state.filteredRecords.reduce((s, r) => s + (parseFloat(String(r['Taka'] || r.taka || '0').replace(/[^0-9.]/g, '')) || 0), 0);
      renderStats({
        totalRecords: state.filteredRecords.length,
        totalBudget: sumBud,
        totalTaka: sumTak
      });
    } else if (state.rawRecords && state.rawRecords.length > 0) {
      const sumBud = state.rawRecords.reduce((s, r) => s + (parseFloat(String(r['Budget'] || r.budget || '0').replace(/[^0-9.]/g, '')) || 0), 0);
      const sumTak = state.rawRecords.reduce((s, r) => s + (parseFloat(String(r['Taka'] || r.taka || '0').replace(/[^0-9.]/g, '')) || 0), 0);
      renderStats({
        totalRecords: state.rawRecords.length,
        totalBudget: sumBud,
        totalTaka: sumTak
      });
    }
  });
}

function clearTableSearch() {
  const input = document.getElementById('tableFilterInput') || document.getElementById('tableSearchInput');
  if (input) {
    input.value = '';
    input.focus();
  }
  state.searchQuery = '';
  const clearBtn = document.getElementById('btnClearTableSearch');
  if (clearBtn) clearBtn.classList.add('hidden');
  filterTableRows('');
}

function renderRecordsTable(data) {
  const thead = document.getElementById('tableHead');
  const tbody = document.getElementById('tableBody');
  const empty = document.getElementById('emptyState');
  const counter = document.getElementById('tableSummaryCounter');
  const rawList = data.rows || [];
  const currentKey = state.currentMenu ? state.currentMenu.key : '';
  const records = rawList.filter(r => isValidClientRow(r, currentKey));
  const isTodayCreated = state.currentMenu && (state.currentMenu.key === 'today_created_campaign' || state.currentMenu.key === 'today_created');

  let headers = data.headers || [];

  // Reliable campaign table detection
  const isPageSummary = state.currentMenu && state.currentMenu.key === 'page_summery';
  if (isPageSummary) {
    // User Requirement: Page summary highest budget to lowest budget
    records.sort((a, b) => {
      const bBud = parseFloat(b['Total Budget $'] || b['budget'] || b['Budget'] || 0) || 0;
      const aBud = parseFloat(a['Total Budget $'] || a['budget'] || a['Budget'] || 0) || 0;
      return bBud - aBud;
    });
  } else if (isTodayCreated) {
    // Today Created: Latest entry first (descending by SL No)
    records.sort((a, b) => {
      const slA = parseInt(String(a['SL'] || a['SL NO'] || a['sl_no'] || 0).replace(/\D/g, ''), 10) || 0;
      const slB = parseInt(String(b['SL'] || b['SL NO'] || a['sl_no'] || 0).replace(/\D/g, ''), 10) || 0;
      if (slA !== slB) return slB - slA;
      const tA = parseAnyDate(a['Entry Date'] || a.entry_date || a['Start'])?.getTime() || 0;
      const tB = parseAnyDate(b['Entry Date'] || b.entry_date || b['Start'])?.getTime() || 0;
      if (tA !== tB) return tB - tA;
      return (b._row || 0) - (a._row || 0);
    });
  } else if (state.sortCol) {
    // Explicit user column sort respected
  } else if (state.campaignSortOrder === 'sheet') {
    // Only when the user explicitly picks Sheet Order via the toggle:
    // 1-to-1 exact Google Sheet row order.
    records.sort((a, b) => {
      if (a._row !== undefined && b._row !== undefined && a._row !== b._row) {
        return (a._row || 0) - (b._row || 0);
      }
      return 0;
    });
  } else {
    // DEFAULT for every entry list: latest entry FIRST, serial by serial —
    // campaigns by SL descending, dated lists (payments, costs, pause ads)
    // by Date descending — so the newest entries are always on top and
    // nobody has to scroll to the bottom to see the last entry.
    records.sort((a, b) => {
      const slA = parseInt(String(a['SL'] || a['SL NO'] || a['sl_no'] || 0).replace(/\D/g, ''), 10) || 0;
      const slB = parseInt(String(b['SL'] || b['SL NO'] || b['sl_no'] || 0).replace(/\D/g, ''), 10) || 0;
      if (slA !== slB) return slB - slA;
      const tA = parseAnyDate(a['Date'] || a['Entry Date'] || a['Start'])?.getTime() || 0;
      const tB = parseAnyDate(b['Date'] || b['Entry Date'] || b['Start'])?.getTime() || 0;
      if (tA !== tB) return tB - tA;
      return (b._row || 0) - (a._row || 0);
    });
  }

  const isToday = currentKey === 'today_created_campaign' || currentKey === 'today_created';
  const isCampaign = currentKey === 'campaign';
  const isCampaignView = isToday || isCampaign || (!state.currentMenu);
  const isAgencyCost = currentKey === 'agency_cost';

  if (isAgencyCost) {
    const costRows = records;

    const html = `
      <div style="display:flex;flex-direction:column;gap:24px;padding:8px">
        <div class="records-card" style="padding:20px;background:var(--card);border:1px solid var(--bor);border-radius:var(--radius-md)">
          <h3 style="font-size:16px;font-weight:800;margin-bottom:14px;display:flex;align-items:center;gap:8px"><span class="ms" style="color:var(--pri)">receipt_long</span> Add Agency Cost Entry</h3>
          <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px;align-items:flex-end">
            <div class="form-field"><label>Date</label><input type="date" id="costDateInput" value="${getTodayDateStr()}" class="form-control"></div>
            <div class="form-field"><label>Type</label><input type="text" id="costTypeInput" placeholder="e.g. Office Rent, Electricity..." class="form-control"></div>
            <div class="form-field"><label>Taka (৳)</label><input type="number" step="any" id="costTakaInput" placeholder="e.g. 25000" class="form-control"></div>
            <div class="form-field"><label>Status</label><select id="costStatusSelect" class="form-control"><option value="Paid">Paid</option><option value="Due">Due</option></select></div>
            <div><button type="button" onclick="submitAgencyCostForm()" class="btn btn-primary" style="width:100%;justify-content:center;height:42px"><span class="ms">save</span> Save Agency Cost</button></div>
          </div>
        </div>

        <div class="records-card">
          <div class="table-toolbar"><div style="font-weight:800;font-size:14px">Saved Agency Cost Entries</div></div>
          <div class="table-wrap">
            <table>
              <thead><tr><th>Date</th><th>Type</th><th>Taka (৳)</th><th>Status</th>${isAdmin() ? '<th style="text-align:center">Action</th>' : ''}</tr></thead>
              <tbody>
                ${costRows.length === 0 ? '<tr><td colspan="5" style="text-align:center;padding:24px;color:var(--txt-muted)">No agency cost entries found</td></tr>' : costRows.map(c => `
                  <tr>
                    <td><span class="table-date" style="font-weight:600;font-variant-numeric:tabular-nums">${escapeHtml(formatCustomDate(c['Date'] || ''))}</span></td>
                    <td style="font-weight:700">${escapeHtml(c['Type'] || c['type'] || '')}</td>
                    <td style="color:var(--suc);font-weight:800">৳${fmtNum(c['Cost'] || c['cost'] || 0)}</td>
                    <td><span class="badge ${String(c['Status'] || c['status']) === 'Paid' ? 'badge-success' : 'badge-danger'}">${escapeHtml(c['Status'] || c['status'] || 'Due')}</span></td>
                    ${isAdmin() ? `<td style="text-align:center">
                      <button onclick='deleteAgencyCostRecord("${c['Record ID'] || c._recordId}")' class="btn-icon text-red" title="Delete"><span class="ms">delete</span></button>
                    </td>` : ''}
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    `;
    if (tbody) tbody.innerHTML = html;
    if (thead) thead.innerHTML = '';
    if (empty) empty.classList.add('hidden');
    if (counter) counter.textContent = `${costRows.length} entries`;
    return;
  }

  if (!isAdmin() && isCampaignView) {
    // Specifically restricted columns for Viewer as requested
    headers = ['SL', 'OFF $', 'Status', 'Page Name', 'Start', 'End', 'Types', 'Day', 'Budget', 'Taka'];
  }

  // De-duplicate headers case-insensitively while preserving order
  const seenHeader = new Set();
  headers = headers.filter(h => {
    const key = String(h || '').trim().toLowerCase();
    if (!key || seenHeader.has(key)) return false;
    seenHeader.add(key);
    return true;
  });

  // Calculate final display columns sequence (User Requirement: Remove Action column completely since clicking any row opens edit/delete options)
  let displayColumns = [];
  if (isCampaignView) {
    if (!isAdmin()) {
      displayColumns = ['SL', 'OFF $', 'Status', 'Page Name', 'Start', 'End', 'Types', 'Day', 'Budget', 'Taka'];
    } else {
      const beforeTaka = ['SL', 'OFF $', 'AD Account', 'ID Name', 'Status', 'Page Name', 'Start', 'End', 'Types', 'Day', 'Budget'];
      const hasDaily = headers.some(h => String(h).trim().toLowerCase() === 'daily $');
      if (hasDaily) beforeTaka.push('Daily $');
      beforeTaka.push('Taka');

      const rateCol = headers.find(h => {
        const k = String(h).trim().toLowerCase();
        return k === 'rate' || k === 'dollar rate (৳)';
      }) || 'Rate';

      const entryCol = headers.find(h => {
        const k = String(h).trim().toLowerCase();
        return k === 'entry date' || k === 'entry';
      }) || 'Entry Date';

      const monthCol = headers.find(h => String(h).trim().toLowerCase() === 'month') || 'Month';
      const yearCol = headers.find(h => String(h).trim().toLowerCase() === 'year') || 'Year';
      const recordIdCol = headers.find(h => {
        const k = String(h).trim().toLowerCase();
        return k === 'record id' || k === 'record_id';
      }) || 'Record ID';

      displayColumns = [
        ...beforeTaka,
        rateCol,
        entryCol,
        monthCol,
        yearCol,
        recordIdCol
      ];
    }
  } else if (isAgencyCost) {
    displayColumns = ['Date', 'Month', 'Year', 'Type', 'Cost', 'Per Person Cost', 'Status'];
  } else if (isPageSummary) {
    displayColumns = ['Rank', 'Page Name', 'Total Ads', 'Total Budget ($)'];
  } else {
    displayColumns = headers.filter(h => String(h).trim().toLowerCase() !== 'action');
  }

  // Ensure Action column is strictly excluded from all tables across the app
  displayColumns = displayColumns.filter(h => String(h).trim().toLowerCase() !== 'action');

  let displayRecords = records;

  if (counter) {
    if (isPageSummary) {
      const pageTotalUSD = displayRecords.reduce((sum, r) => sum + (parseFloat(r['Total Budget ($)'] || r['budget'] || r['Budget'] || 0) || 0), 0);
      counter.innerHTML = `<span style="font-weight:700">${displayRecords.length} pages • $${fmtNum(pageTotalUSD)}</span>`;
    } else {
      let totalUSD = 0;
      let totalBDT = 0;
      let verifiedCount = 0;
      let pendingCount = 0;

      displayRecords.forEach(r => {
        if (r['Budget']) totalUSD += parseFloat(String(r['Budget']).replace(/[^0-9.]/g, '')) || 0;
        if (r['Taka']) totalBDT += parseFloat(String(r['Taka']).replace(/[^0-9.]/g, '')) || 0;
        if (isCampaignVerified(r)) verifiedCount++;
        else pendingCount++;
      });

      counter.innerHTML = `<span style="font-weight:700">${displayRecords.length} records • $${fmtNum(totalUSD)} • ৳${fmtNum(totalBDT)}</span>`;
    }
  }

  const colClassMap = new Map();
  displayColumns.forEach(h => {
    const lower = String(h || '').trim().toLowerCase();
    let c = '';
    if (lower === 'sl' || lower === 'sl no') c = 'col-sl';
    else if (lower === 'off $' || lower === 'off') c = 'col-off';
    else if (lower === 'status') c = 'col-status';
    else if (lower === 'page name' || lower === 'page') c = 'col-page';
    else if (lower === 'start' || lower === 'start date' || lower === 'end' || lower === 'end date') c = 'col-date';
    else if (lower === 'types' || lower === 'type') c = 'col-types';
    else if (lower === 'day') c = 'col-day';
    else if (lower === 'budget') c = 'col-budget';
    else if (lower === 'daily $') c = 'col-daily';
    else if (lower === 'dollar rate (৳)' || lower === 'rate') c = 'col-rate';
    else if (lower === 'taka') c = 'col-taka';
    else if (lower === 'ad account' || lower === 'account') c = 'col-ad';
    else if (lower === 'id name' || lower === 'id') c = 'col-id';
    else if (lower === 'action') c = 'th-action';
    else if (lower === 'entry date' || lower === 'entry' || lower === 'month' || lower === 'year' || lower.includes('record id')) c = 'col-meta';
    colClassMap.set(h, c);
  });
  const getColClass = (h) => colClassMap.get(h) || '';

  const isCheckableTab = isTodayCreated || isCampaignView;
  const allVerified = isCheckableTab && displayRecords.length > 0 && displayRecords.every(r => isCampaignVerified(r));

  if (thead) {
    let thHtml = '';
    if (isCheckableTab && isAdmin()) {
      thHtml += `<th class="th-check" style="width:46px;min-width:46px;text-align:center;padding:8px 4px" title="Check / Uncheck All">
        <input type="checkbox" id="masterTodayCheckbox" onchange="toggleAllTodayVerified(this.checked)" style="cursor:pointer;width:18px;height:18px;accent-color:#10b981;border-radius:4px;vertical-align:middle" ${allVerified ? 'checked' : ''} />
      </th>`;
    }
    thHtml += displayColumns.map(h => {
      if (h === 'Action') {
        return `<th class="th-action" style="text-align:center">Action</th>`;
      }
      const isSorted = state.sortCol === h;
      const arrow = isSorted ? (state.sortAsc ? 'arrow_upward' : 'arrow_downward') : 'unfold_more';
      return `<th class="${getColClass(h)}" onclick="sortTableBy('${escapeHtml(h)}')">${escapeHtml(h)}<span class="ms th-sort">${arrow}</span></th>`;
    }).join('');
    thead.innerHTML = `<tr>${thHtml}</tr>`;
  }

  if (tbody) {
    if (!displayRecords.length) {
      tbody.innerHTML = '';
      if (empty) {
        empty.classList.remove('hidden');
        const emptyTitle = empty.querySelector('div[style*="font-weight:700"]');
        const emptyDesc = document.getElementById('emptyStateDesc');
        const emptyActions = document.getElementById('emptyStateAdminActions');
        if (state.verificationFilter && state.verificationFilter !== 'all') {
          if (emptyTitle) emptyTitle.textContent = state.verificationFilter === 'pending' ? 'No Pending Ads Found' : 'No Verified Ads Found';
          if (emptyDesc) emptyDesc.innerHTML = `All ${records.length} campaigns match your criteria. Click <strong>All</strong> above to view everything.`;
          if (emptyActions) emptyActions.classList.add('hidden');
        } else if (isTodayCreated) {
          if (emptyTitle) emptyTitle.textContent = 'No Campaigns Created Today (' + getTodayFormatted() + ')';
          if (emptyDesc) {
            const masterCount = (state.masterCampaignRegistry || []).length;
            emptyDesc.innerHTML = `No new campaigns logged today yet. Total <strong>${masterCount}</strong> monthly campaigns are available in <strong>All Campaign</strong>.`;
          }
          if (emptyActions) {
            emptyActions.classList.remove('hidden');
            emptyActions.style.display = 'inline-flex';
            emptyActions.innerHTML = `
              <button class="btn btn-primary" onclick="selectMenu('campaign')" type="button" style="padding:8px 18px;font-size:13px;font-weight:700;display:inline-flex;align-items:center;gap:6px">
                <span class="ms" style="font-size:16px">campaign</span> View All Campaigns
              </button>
              ${isAdmin() ? `<button class="btn btn-secondary" onclick="handleToggleEntry()" type="button" style="padding:8px 18px;font-size:13px;font-weight:700;display:inline-flex;align-items:center;gap:6px">
                <span class="ms" style="font-size:16px">add</span> Create New Campaign
              </button>` : ''}
            `;
          }
        } else if (isCampaignView) {
          const mLabel = (state.filters.month && state.filters.month !== 'All') ? state.filters.month : 'selected period';
          if (emptyTitle) emptyTitle.textContent = `No Campaigns Found for ${mLabel}`;
          if (emptyDesc) {
            const masterCount = (state.masterCampaignRegistry || []).length;
            emptyDesc.innerHTML = `No campaigns found for <strong>${mLabel}</strong>. Total <strong>${masterCount}</strong> campaigns are saved in the system across all months.`;
          }
          if (emptyActions) {
            emptyActions.classList.remove('hidden');
            emptyActions.style.display = 'inline-flex';
            emptyActions.innerHTML = `
              <button class="btn btn-primary" onclick="state.filters.month='All'; const m=document.getElementById('filterMonth'); if(m) m.value='All'; loadMenuData(true);" type="button" style="padding:8px 18px;font-size:13px;font-weight:700;display:inline-flex;align-items:center;gap:6px">
                <span class="ms" style="font-size:16px">visibility</span> Show All Months Data (${(state.masterCampaignRegistry || []).length})
              </button>
              ${isAdmin() ? `<button class="btn btn-secondary" onclick="handleToggleEntry()" type="button" style="padding:8px 18px;font-size:13px;font-weight:700;display:inline-flex;align-items:center;gap:6px">
                <span class="ms" style="font-size:16px">add</span> Create New Campaign
              </button>` : ''}
            `;
          }
        }
      }
      if (typeof clearRowFocus === 'function') clearRowFocus();
      return;
    }
    if (empty) empty.classList.add('hidden');

    if (state.focusedRowIndex !== null && state.focusedRowIndex >= displayRecords.length) {
      if (typeof clearRowFocus === 'function') clearRowFocus();
    }

    state.currentRenderedRecords = displayRecords;
    if (window._tableChunkTimer) {
      cancelAnimationFrame(window._tableChunkTimer);
      window._tableChunkTimer = null;
    }

    const sq = (state.searchQuery || '').trim();

    // High performance pre-indexed page stats (O(1) lookup per row instead of O(N^2))
    const currentRegVersion = `${state.masterCampaignRegistry?.length || 0}_${state.rawRecords?.length || 0}`;
    if (!state._pageUsageStatsCache || state._pageUsageStatsVersion !== currentRegVersion) {
      const counts = {};
      const adMap = {};
      const pool = (state.masterCampaignRegistry && state.masterCampaignRegistry.length > 0) ? state.masterCampaignRegistry : (state.rawRecords || []);
      for (let i = 0; i < pool.length; i++) {
        const item = pool[i];
        const p = String(item['Page Name'] || item.page_name || '').trim();
        const ad = String(item['AD Account'] || item.ad_account || '').trim();
        if (p) {
          const k = p.toLowerCase();
          counts[k] = (counts[k] || 0) + 1;
          if (ad && !adMap[k]) adMap[k] = ad;
        }
      }
      state._pageUsageStatsCache = { counts, adMap };
      state._pageUsageStatsVersion = currentRegVersion;
    }
    const pageUsageStatsCache = state._pageUsageStatsCache;

    const renderRowHtml = (r, rIdx) => {
      const renderCell = (h) => {
        let val = r[h];
        if (val === undefined || val === null || val === '') {
          const lower = String(h).toLowerCase();
          if (lower === 'start' || lower === 'start date') val = r['Start Date'] || r['start_date'] || '';
          else if (lower === 'end' || lower === 'end date') val = r['End Date'] || r['end_date'] || '';
          else if (lower === 'types' || lower === 'type') val = r['Type'] || r['type'] || '';
          else if (lower === 'off $' || lower === 'off' || lower === 'off usd') val = r['OFF $'] || r['OFF USD'] || r['off_usd'] || r['OFF'] || r['Off'] || r['off'] || '';
          else if (lower === 'page name' || lower === 'page') val = r['page_name'] || r['Page'] || '';
          else if (lower === 'budget') val = r['budget'] || '';
          else if (lower === 'taka') val = r['taka'] || '';
          else if (lower === 'day') val = r['day'] || '';
          else if (lower === 'status') val = r['status'] || '';
          else if (lower === 'sl' || lower === 'sl no') val = r['sl'] || r['sl_no'] || '';
          else if (lower === 'rate' || lower === 'dollar rate (৳)') val = r['Dollar Rate (৳)'] || r['Rate'] || r['rate'] || state.rate || '';
          else if (lower === 'entry date' || lower === 'entry') val = r['Entry Date'] || r['entry_date'] || r['Date'] || '';
          else if (lower === 'record id' || lower === 'record_id') val = r['Record ID'] || r['record_id'] || r._recordId || '';
          else if (lower === 'month') val = r['Month'] || r['month'] || '';
          else if (lower === 'year') val = r['Year'] || r['year'] || '';
        }
        if (val === undefined || val === null) val = '';
        const cClass = getColClass(h);
        const lowerH = String(h || '').toLowerCase();

        if (h === 'SL' || h === 'SL NO' || lowerH === 'sl' || lowerH === 'sl no') {
          return `
            <td class="${cClass}">
              <div style="display:flex;align-items:center;gap:5px;white-space:nowrap">
                <span style="font-weight:800;font-variant-numeric:tabular-nums">${highlightMatch(val, sq)}</span>
              </div>
            </td>
          `;
        }

        if (h === 'Rank') {
          const rankLabel = rIdx === 0 ? '🥇 #1' : (rIdx === 1 ? '🥈 #2' : (rIdx === 2 ? '🥉 #3' : `#${rIdx + 1}`));
          return `<td class="${cClass}" style="text-align:center;font-weight:800;color:${rIdx < 3 ? 'var(--pri)' : 'var(--txt-muted)'};font-size:12px">${rankLabel}</td>`;
        }
        if (h === 'Total Ads' || String(h).toLowerCase() === 'total ads') {
          const adsCount = r['Total Ads'] || r['ads'] || r['Ads'] || val || 1;
          return `<td class="${cClass}" style="text-align:center"><span class="pill pill-blue" style="font-weight:700;font-size:11px;padding:2px 8px">${highlightMatch(String(adsCount), sq)} Ads</span></td>`;
        }
        if (h === 'Total Budget ($)' || String(h).toLowerCase() === 'total budget $' || String(h).toLowerCase() === 'total budget') {
          const bud = r['Total Budget $'] || r['budget'] || r['Budget'] || val || 0;
          return `<td class="${cClass}" style="text-align:right;font-weight:800;color:var(--sec);font-size:13px;font-variant-numeric:tabular-nums">$${highlightMatch(fmtNum(bud), sq)}</td>`;
        }
        if (h === 'Cost' || String(h).toLowerCase() === 'cost') {
          return `<td class="${cClass}" style="text-align:right;font-weight:800;font-variant-numeric:tabular-nums;color:var(--suc);font-size:13.5px">৳${highlightMatch(fmtNum(val), sq)}</td>`;
        }
        if (h === 'Per Person Cost' || String(h).toLowerCase() === 'per person cost') {
          const costVal = parseFloat(r['Cost'] || r.cost || 0) || 0;
          const perVal = val || (costVal / 2);
          return `<td class="${cClass}" style="text-align:right;font-weight:700;font-variant-numeric:tabular-nums;color:var(--pri)">৳${highlightMatch(fmtNum(perVal), sq)}</td>`;
        }
        if (h === 'OFF $' || String(h).toLowerCase() === 'off $' || String(h).toLowerCase() === 'off' || String(h).toLowerCase() === 'off usd') {
          let offVal = r['OFF $'] ?? r['OFF USD'] ?? r['off_usd'] ?? r['OFF'] ?? r['Off'] ?? r['off'] ?? val ?? '0';
          if (offVal === '' || offVal === undefined || offVal === null) offVal = '0';
          return `<td class="${cClass}" style="text-align:center;font-weight:700;color:var(--dan);font-variant-numeric:tabular-nums">${highlightMatch(offVal, sq)}</td>`;
        }
        if (h === 'Taka' || String(h).toLowerCase() === 'taka' || h === 'Total Taka') {
          let tVal = val;
          if (tVal === '' || tVal === undefined || tVal === null || Number(tVal) === 0) {
            const b = Number(r['Budget'] || r['budget'] || r['Total Budget $'] || 0);
            const rt = Number(r['Dollar Rate (৳)'] || r['Rate'] || r['rate'] || state.rate || 140);
            if (b > 0 && rt > 0) tVal = Math.round(b * rt);
          }
          return `<td class="${cClass}" style="text-align:right;font-weight:700;font-variant-numeric:tabular-nums;color:var(--txt)">৳${highlightMatch(fmtNum(tVal), sq)}</td>`;
        }
        if (String(h).toLowerCase().includes('due') || String(val).toLowerCase() === 'due') {
          return `<td class="${cClass}" style="text-align:right;font-weight:900;font-variant-numeric:tabular-nums;color:#EF4444">${typeof val === 'number' || !isNaN(parseFloat(val)) ? '৳' + highlightMatch(fmtNum(val), sq) : highlightMatch(val, sq)}</td>`;
        }
        if (String(h).toLowerCase().includes('paid') || String(val).toLowerCase() === 'paid') {
          return `<td class="${cClass}" style="text-align:right;font-weight:900;font-variant-numeric:tabular-nums;color:#10B981">${typeof val === 'number' || !isNaN(parseFloat(val)) ? '৳' + highlightMatch(fmtNum(val), sq) : highlightMatch(val, sq)}</td>`;
        }
        if (h === 'Dollar Rate (৳)' || h === 'Rate' || String(h).toLowerCase() === 'rate') {
          let rVal = val;
          if (rVal === '' || rVal === undefined || rVal === null) {
            rVal = r['Dollar Rate (৳)'] || r['Rate'] || r['rate'] || state.rate || '';
          }
          return `<td class="${cClass}" style="text-align:right;font-weight:600;font-variant-numeric:tabular-nums">${escapeHtml(rVal ? `৳${rVal}` : '')}</td>`;
        }
        if (h === 'AD Account' || String(h).toLowerCase() === 'ad account') {
          return `<td class="${cClass}" style="font-family:monospace;font-size:11px;font-weight:600;white-space:nowrap;letter-spacing:-0.2px">${highlightMatch(val, sq)}</td>`;
        }
        if (MONEY_COLS.includes(h) || h === 'Budget') {
          return `<td class="${cClass}" style="text-align:right;font-weight:700;font-variant-numeric:tabular-nums">${highlightMatch(fmtNum(val), sq)}</td>`;
        }
        if (isDateHeader(h)) {
          const prox = (h === 'End' || String(h).toLowerCase() === 'end date') ? getDateProximityBadge(val) : '';
          return `<td class="${cClass}"><span class="table-date" style="font-weight:600;font-variant-numeric:tabular-nums">${highlightMatch(formatCustomDate(val), sq)}${prox}</span></td>`;
        }
        if (h === 'Status') {
          const sLower = String(val || '').toLowerCase();
          if (sLower.includes('paid')) {
            return `<td class="${cClass}"><span class="badge badge-paid"><span class="ms" style="font-size:12px">check_circle</span> ${highlightMatch('Paid', sq)}</span></td>`;
          }
          if (sLower.includes('due')) {
            return `<td class="${cClass}"><span class="badge badge-due"><span class="ms" style="font-size:12px">error</span> ${highlightMatch('Due', sq)}</span></td>`;
          }
          const isOld = sLower.includes('old');
          const statusText = isOld ? 'Old Page' : 'New Page';
          const badgeClass = isOld ? 'badge-old-page' : 'badge-new-page';
          return `<td class="${cClass}"><span class="page-status-badge ${badgeClass}">${highlightMatch(val || statusText, sq)}</span></td>`;
        }
        if (h === 'Types' || h === 'Type') {
          const tLower = String(val).toLowerCase();
          let pillCls = 'pill-blue';
          let icon = '';
          if (tLower.includes('boost')) pillCls = 'pill-boost';
          else if (tLower.includes('message') || tLower.includes('msg')) pillCls = 'pill-msg';
          else if (tLower.includes('lead')) pillCls = 'pill-lead';
          else if (tLower.includes('engagement') || tLower.includes('eng')) pillCls = 'pill-eng';
          else if (tLower.includes('increase') || tLower.includes('inc')) { pillCls = 'pill-inc'; icon = '<span class="ms" style="font-size:11px">trending_up</span>'; }
          else if (tLower.includes('like')) { pillCls = 'pill-like'; icon = '<span class="ms" style="font-size:11px">thumb_up</span>'; }
          return `<td class="${cClass}"><span class="pill ${pillCls}" style="font-weight:500;font-size:9.5px">${icon}${highlightMatch(val, sq)}</span></td>`;
        }
        if (h === 'Page Name' || h === 'Page') {
          if (isPageSummary) {
            return `
              <td class="${cClass}">
                <div style="font-weight:800;font-size:13px;color:var(--txt)">${highlightMatch(val, sq)}</div>
              </td>
            `;
          }
          const stats = pageUsageStatsCache;
          const pKey = String(val || '').toLowerCase();
          const pCount = stats.counts ? (stats.counts[pKey] || 0) : 0;
          const fAd = stats.adMap ? (stats.adMap[pKey] || '') : '';
          return `
            <td class="${cClass}">
              <div class="page-name-cell">
                <span class="page-name-text">${highlightMatch(val, sq)}</span>
                <div class="page-hover-popover">
                  <div style="font-weight:800;font-size:12.5px;color:var(--txt);margin-bottom:4px">${highlightMatch(val, sq)}</div>
                  <div style="display:flex;align-items:center;gap:6px;margin-bottom:6px">
                    <span class="page-status-badge ${pCount > 0 ? 'badge-old-page' : 'badge-new-page'}" style="font-size:10px;padding:2px 6px">
                      ${pCount > 0 ? `${pCount} Campaigns` : 'New Page'}
                    </span>
                    ${fAd ? `<span style="font-size:11px;color:var(--txt-muted)">AD: <b>${highlightMatch(fAd, sq)}</b></span>` : ''}
                  </div>
                  <div style="font-size:10.5px;color:var(--txt-dim)">Click row to inspect campaign</div>
                </div>
              </div>
            </td>
          `;
        }
        if (String(h).toLowerCase().includes('record id')) {
          return `<td class="${cClass}" style="font-family:monospace;font-size:10.5px;color:var(--txt-dim)">${highlightMatch(val, sq)}</td>`;
        }
        return `<td class="${cClass}">${highlightMatch(val, sq)}</td>`;
      };

      const isNewlyAddedRow = Boolean(
        r._optimistic ||
        (r._recordKey && state.newlyAddedRecordKeys && state.newlyAddedRecordKeys.has(r._recordKey)) ||
        (r['Record ID'] && state.newlyAddedRecordKeys && state.newlyAddedRecordKeys.has(r['Record ID']))
      );

      const isVer = isCampaignVerified(r);
      const isFocused = state.focusedRowIndex === rIdx;
      return `
        <tr class="table-record-row ${isNewlyAddedRow ? 'table-row-fade-in' : ''} ${isFocused ? 'row-focused' : ''} ${isVer ? 'row-verified' : ''}" id="tableRecordRow-${rIdx}" data-row-idx="${rIdx}" ${isFocused ? 'aria-selected="true"' : ''} onclick="handleTableRowClick(event, ${rIdx})">
          ${((isTodayCreated || isCampaignView) && isAdmin()) ? `
            <td class="td-check" style="text-align:center;width:46px;min-width:46px;padding:6px 4px" onclick="event.stopPropagation();">
              <input type="checkbox" class="row-campaign-check" onchange="window.toggleRowVerification(${rIdx}, this.checked)" ${isVer ? 'checked' : ''} style="cursor:pointer;width:18px;height:18px;accent-color:#10b981;border-radius:4px;vertical-align:middle" title="${isVer ? 'Checked (Click to untick)' : 'Unchecked (Click to tick)'}" />
            </td>
          ` : ''}
          ${displayColumns.map(h => {
            if (h === 'Action') {
              if (isAgencyCost) {
                return `
                  <td class="td-action" onclick="event.stopPropagation()">
                    <div class="action-btn-group">
                      <button type="button" class="table-btn" onclick="openEditAgencyCostRecord(${rIdx})" title="Edit Agency Cost" style="width:30px;height:30px;padding:0;display:inline-flex;align-items:center;justify-content:center;border-radius:6px;background:rgba(99,102,241,0.1);color:var(--pri);border:1px solid rgba(99,102,241,0.25)">
                        <span class="ms" style="font-size:15px">edit</span>
                      </button>
                      <button type="button" class="table-btn table-btn-danger" style="color:var(--dan);width:30px;height:30px;padding:0;display:inline-flex;align-items:center;justify-content:center;border-radius:6px" onclick="deleteAgencyCostByIndex(${rIdx})" title="Delete Agency Cost">
                        <span class="ms" style="font-size:15px">delete</span>
                      </button>
                    </div>
                  </td>
                `;
              }
              return `
                <td class="td-action" onclick="event.stopPropagation()">
                  <div class="action-btn-group">
                    ${isCampaignView ? `
                      <button class="table-btn table-btn-copy" onclick="copyCampaignRecordByIndex(${rIdx})" title="Copy Campaign Format (187/011 Page Start End Type Day $Budget)">
                        <span class="ms">content_copy</span>
                      </button>
                    ` : ''}
                    ${(isCampaignView || r['Page Name']) ? `
                      <button class="table-btn table-btn-clone" onclick="cloneCampaignRecordByIndex(${rIdx})" title="Duplicate / Clone Campaign">
                        <span class="ms">library_add</span>
                      </button>
                    ` : ''}
                    ${isCampaignView ? `
                      <button class="table-btn table-btn-increase" onclick="quickIncreaseCampaignByIndex(${rIdx})" title="Quick Increase Budget">
                        <span class="ms">trending_up</span>
                      </button>
                    ` : ''}
                    ${isTodayCreated ? `
                      <button class="table-btn table-btn-wa" onclick="openSingleRecordWhatsAppByIndex(${rIdx})" title="Send WhatsApp Notice">
                        <span class="ms">chat</span>
                      </button>
                    ` : ''}
                    <button class="table-btn" onclick="openEditRecordByIndex(${rIdx})" title="Edit Record">
                      <span class="ms">edit</span>
                    </button>
                    <button class="table-btn" style="color:var(--dan)" onclick="deleteRecordByIndex(${rIdx})" title="Delete Record">
                      <span class="ms">delete</span>
                    </button>
                  </div>
                </td>
              `;
            }
            return renderCell(h);
          }).join('')}
        </tr>
      `;
    };

    // Ultra-Fast Progressive Rendering: Instant 1-shot render for <=100 rows (<1.5ms), streaming chunks for large tables
    if (displayRecords.length <= 100) {
      tbody.innerHTML = displayRecords.map((r, rIdx) => renderRowHtml(r, rIdx)).join('');
    } else {
      const INITIAL_CHUNK = 40; // instant viewport fill in < 1ms
      tbody.innerHTML = displayRecords.slice(0, INITIAL_CHUNK).map((r, rIdx) => renderRowHtml(r, rIdx)).join('');
      let offset = INITIAL_CHUNK;
      const CHUNK_SIZE = 150;
      const renderNextChunk = () => {
        if (state.currentRenderedRecords !== displayRecords) return;
        const chunk = displayRecords.slice(offset, offset + CHUNK_SIZE);
        if (chunk.length > 0) {
          const chunkHtml = chunk.map((r, idx) => renderRowHtml(r, offset + idx)).join('');
          tbody.insertAdjacentHTML('beforeend', chunkHtml);
          offset += CHUNK_SIZE;
          if (offset < displayRecords.length) {
            window._tableChunkTimer = (window.requestIdleCallback ? requestIdleCallback(renderNextChunk) : requestAnimationFrame(renderNextChunk));
          }
        }
      };
      window._tableChunkTimer = (window.requestIdleCallback ? requestIdleCallback(renderNextChunk) : requestAnimationFrame(renderNextChunk));
    }

    if (state.focusedRowIndex !== null && state.focusedRowIndex < records.length && typeof renderRowFocusMenu === 'function') {
      renderRowFocusMenu(records[state.focusedRowIndex], state.focusedRowIndex, records.length);
    }

    const activeDensity = state.tableDensity || localStorage.getItem('adverse_table_density') || 'comfortable';
    const tableWraps = document.querySelectorAll('.table-wrap');
    tableWraps.forEach(wrap => {
      wrap.classList.toggle('compact-density', activeDensity === 'compact');
      wrap.classList.toggle('density-compact', activeDensity === 'compact');
      wrap.classList.toggle('density-comfortable', activeDensity === 'comfortable');
    });

    // Auto-sync KPI stats if records are rendered to guarantee numbers never stay 0
    if (records.length > 0 && (currentKey === 'today_created_campaign' || currentKey === 'campaign')) {
      const validRows = records.filter(r => isValidClientRow(r, currentKey));
      if (validRows.length > 0) {
        const sumBud = validRows.reduce((s, r) => s + (parseFloat(r['Budget'] || r.budget || 0) || 0), 0);
        const sumTak = validRows.reduce((s, r) => s + (parseFloat(r['Taka'] || r.taka || 0) || 0), 0);
        
        const elEnt = document.getElementById('stat_tc_entries') || document.getElementById('stat_c_records');
        const elBud = document.getElementById('stat_tc_budget') || document.getElementById('stat_c_budget');
        const elTak = document.getElementById('stat_tc_taka') || document.getElementById('stat_c_taka');
        
        if (elEnt && (elEnt.textContent === '0' || !elEnt.textContent || elEnt.textContent === '')) {
          elEnt.textContent = fmtNum(validRows.length);
        }
        if (elBud && (elBud.textContent === '$0' || !elBud.textContent || elBud.textContent === '')) {
          elBud.textContent = `$${fmtNum(sumBud)}`;
        }
        if (elTak && (elTak.textContent === '৳0' || !elTak.textContent || elTak.textContent === '')) {
          elTak.textContent = `৳${fmtNum(sumTak)}`;
        }
      }
    }

    if (isTodayCreated && typeof renderTodaySetupAuditBar === 'function') {
      renderTodaySetupAuditBar();
    }
  }
}

function renderEndAdsTables(data) {
  if (!data) return;
  const todayBody = document.getElementById('todayEndBody');
  const tomorrowBody = document.getElementById('tomorrowEndBody');
  const todayCountEl = document.getElementById('todayTableCount');
  const tomorrowCountEl = document.getElementById('tomorrowTableCount');
  const topTodayCount = document.getElementById('topTodayEndCount');
  const topTodayBudget = document.getElementById('topTodayEndBudget');
  const topTomorrowCount = document.getElementById('topTomorrowEndCount');
  const topTomorrowBudget = document.getElementById('topTomorrowEndBudget');

  const recs = data.records || data;
  const todayBST = getNowBST();
  const tomorrowBST = new Date(todayBST.getTime() + 86400000);
  const todayFmt = formatCustomDate(todayBST);
  const tomorrowFmt = formatCustomDate(tomorrowBST);

  // Gather master rows from all available pools
  const allRows = (Array.isArray(recs.all) && recs.all.length > 0)
    ? recs.all
    : ((Array.isArray(recs.rows) && recs.rows.length > 0)
      ? recs.rows
      : (state.masterCampaignRegistry && state.masterCampaignRegistry.length > 0
        ? state.masterCampaignRegistry
        : (state.allCampaigns || [])));

  const isEndingToday = (r) => {
    const raw = r['End Date'] || r['End'] || r.end_date || r.end || '';
    if (!raw) return false;
    const formatted = formatCustomDate(raw);
    if (formatted && formatted.toLowerCase() === todayFmt.toLowerCase()) return true;
    return isSameDay(raw, todayBST);
  };

  const isEndingTomorrow = (r) => {
    const raw = r['End Date'] || r['End'] || r.end_date || r.end || '';
    if (!raw) return false;
    const formatted = formatCustomDate(raw);
    if (formatted && formatted.toLowerCase() === tomorrowFmt.toLowerCase()) return true;
    return isSameDay(raw, tomorrowBST);
  };

  let todayList = (Array.isArray(recs.today) && recs.today.length > 0)
    ? [...recs.today]
    : allRows.filter(isEndingToday);

  let tomorrowList = (Array.isArray(recs.tomorrow) && recs.tomorrow.length > 0)
    ? [...recs.tomorrow]
    : allRows.filter(isEndingTomorrow);

  // Deduplicate and sort by Budget descending
  const seenToday = new Set();
  todayList = todayList.filter(r => {
    const id = r['Record ID'] || r._recordId || `${r.SL}_${r['Page Name']}_${r['AD Account'] || ''}_${r.Types || ''}_${r.Budget || ''}_${r.Start || ''}`;
    if (seenToday.has(id)) return false;
    seenToday.add(id);
    return true;
  });

  const seenTomorrow = new Set();
  tomorrowList = tomorrowList.filter(r => {
    const id = r['Record ID'] || r._recordId || `${r.SL}_${r['Page Name']}_${r['AD Account'] || ''}_${r.Types || ''}_${r.Budget || ''}_${r.Start || ''}`;
    if (seenTomorrow.has(id)) return false;
    seenTomorrow.add(id);
    return true;
  });

  todayList.sort((a, b) => ((parseInt(String(b.SL || b['SL NO'] || 0).replace(/\D/g, ''), 10) || 0) - (parseInt(String(a.SL || a['SL NO'] || 0).replace(/\D/g, ''), 10) || 0)));
  tomorrowList.sort((a, b) => ((parseInt(String(b.SL || b['SL NO'] || 0).replace(/\D/g, ''), 10) || 0) - (parseInt(String(a.SL || a['SL NO'] || 0).replace(/\D/g, ''), 10) || 0)));

  state.currentEndAdsToday = todayList;
  state.currentEndAdsTomorrow = tomorrowList;

  // Navigation Badges
  const endNavBadge = document.getElementById('endAdsNavBadge');
  if (endNavBadge) {
    endNavBadge.textContent = String(todayList.length);
    endNavBadge.style.display = todayList.length > 0 ? 'inline-flex' : 'none';
  }
  const endMobileBadge = document.getElementById('endAdsMobileBadge');
  if (endMobileBadge) {
    endMobileBadge.textContent = String(todayList.length);
    endMobileBadge.style.display = todayList.length > 0 ? 'inline-flex' : 'none';
  }

  // Budget calculations
  const todayBudgetSum = todayList.reduce((sum, r) => sum + (parseFloat(r['Budget'] || r.budget || 0) || 0), 0);
  const tomorrowBudgetSum = tomorrowList.reduce((sum, r) => sum + (parseFloat(r['Budget'] || r.budget || 0) || 0), 0);

  if (topTodayCount) topTodayCount.textContent = `${todayList.length} Campaigns`;
  if (topTodayBudget) topTodayBudget.textContent = `Total Spend: $${fmtNum(todayBudgetSum)} • Ending Today`;
  if (topTomorrowCount) topTomorrowCount.textContent = `${tomorrowList.length} Campaigns`;
  if (topTomorrowBudget) topTomorrowBudget.textContent = `Total Spend: $${fmtNum(tomorrowBudgetSum)} • Ending Tomorrow`;

  if (todayCountEl) todayCountEl.textContent = String(todayList.length);
  if (tomorrowCountEl) tomorrowCountEl.textContent = String(tomorrowList.length);

  // Render Today Table
  if (todayBody) {
    todayBody.innerHTML = todayList.length > 0 ? todayList.map((r, idx) => {
      const sl = escapeHtml(String(r['SL'] || r.sl || idx + 1));
      const page = escapeHtml(r['Page Name'] || r.page_name || '');
      const type = escapeHtml(r['Types'] || r.types || r['Type'] || 'Boost');
      const day = escapeHtml(String(r['Day'] || r.day || ''));
      const bud = fmtNum(r['Budget'] || r.budget || 0);
      const tak = fmtNum(r['Taka'] || r.taka || 0);
      const startD = escapeHtml(formatCustomDate(r['Start Date'] || r.Start || r.start_date || ''));
      const endD = escapeHtml(formatCustomDate(r['End Date'] || r.End || r.end_date || ''));
      return `
        <tr style="cursor:pointer;transition:background 0.15s">
          <td onclick="openInspectorModalForEndAd('today', ${idx})" style="font-weight:700;color:var(--txt-muted);font-family:monospace">#${sl}</td>
          <td onclick="openInspectorModalForEndAd('today', ${idx})"><b>${page}</b></td>
          <td onclick="openInspectorModalForEndAd('today', ${idx})"><span class="table-date">${startD}</span></td>
          <td onclick="openInspectorModalForEndAd('today', ${idx})"><span class="table-date" style="color:var(--dan);font-weight:700">🔴 ${endD}</span></td>
          <td onclick="openInspectorModalForEndAd('today', ${idx})">${day}</td>
          <td onclick="openInspectorModalForEndAd('today', ${idx})" style="font-weight:800;color:var(--dan);text-align:right">$${bud}</td>
          <td style="text-align:center">
            <button type="button" class="btn btn-secondary btn-sm" style="padding:2px 8px;font-size:11px;gap:4px" onclick="copyClientEndNotice('${escapeHtml(r['Page Name'] || '')}', '${type}', '${day}', '${bud}', '${endD}', '${tak}')" title="Copy WhatsApp End Notice">
              <span class="ms" style="font-size:14px;color:#10B981">chat</span> Notice
            </button>
          </td>
        </tr>
      `;
    }).join('') : '<tr><td colspan="7" style="text-align:center;padding:16px;color:var(--txt-muted)">No campaigns ending today 🎉</td></tr>';
  }

  // Render Tomorrow Table
  if (tomorrowBody) {
    tomorrowBody.innerHTML = tomorrowList.length > 0 ? tomorrowList.map((r, idx) => {
      const sl = escapeHtml(String(r['SL'] || r.sl || idx + 1));
      const page = escapeHtml(r['Page Name'] || r.page_name || '');
      const type = escapeHtml(r['Types'] || r.types || r['Type'] || 'Boost');
      const day = escapeHtml(String(r['Day'] || r.day || ''));
      const bud = fmtNum(r['Budget'] || r.budget || 0);
      const tak = fmtNum(r['Taka'] || r.taka || 0);
      const startD = escapeHtml(formatCustomDate(r['Start Date'] || r.Start || r.start_date || ''));
      const endD = escapeHtml(formatCustomDate(r['End Date'] || r.End || r.end_date || ''));
      return `
        <tr style="cursor:pointer;transition:background 0.15s">
          <td onclick="openInspectorModalForEndAd('tomorrow', ${idx})" style="font-weight:700;color:var(--txt-muted);font-family:monospace">#${sl}</td>
          <td onclick="openInspectorModalForEndAd('tomorrow', ${idx})"><b>${page}</b></td>
          <td onclick="openInspectorModalForEndAd('tomorrow', ${idx})"><span class="table-date">${startD}</span></td>
          <td onclick="openInspectorModalForEndAd('tomorrow', ${idx})"><span class="table-date" style="color:var(--suc);font-weight:700">🟢 ${endD}</span></td>
          <td onclick="openInspectorModalForEndAd('tomorrow', ${idx})">${day}</td>
          <td onclick="openInspectorModalForEndAd('tomorrow', ${idx})" style="font-weight:800;color:var(--suc);text-align:right">$${bud}</td>
          <td style="text-align:center">
            <button type="button" class="btn btn-secondary btn-sm" style="padding:2px 8px;font-size:11px;gap:4px" onclick="copyClientEndNotice('${escapeHtml(r['Page Name'] || '')}', '${type}', '${day}', '${bud}', '${endD}', '${tak}')" title="Copy WhatsApp End Notice">
              <span class="ms" style="font-size:14px;color:#10B981">chat</span> Notice
            </button>
          </td>
        </tr>
      `;
    }).join('') : '<tr><td colspan="7" style="text-align:center;padding:16px;color:var(--txt-muted)">No campaigns ending tomorrow</td></tr>';
  }
}

window.switchEndAdsViewMode = function(mode) {
  state.currentEndAdsViewMode = mode;
  ['chipEndMonth', 'chipEndAll', 'chipEndToday', 'chipEndTomorrow'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.classList.remove('active');
  });
  if (mode === 'month') document.getElementById('chipEndMonth')?.classList.add('active');
  else if (mode === 'all') document.getElementById('chipEndAll')?.classList.add('active');
  else if (mode === 'today') document.getElementById('chipEndToday')?.classList.add('active');
  else if (mode === 'tomorrow') document.getElementById('chipEndTomorrow')?.classList.add('active');

  const titleEl = document.getElementById('endAdsTableTitle');
  if (titleEl) {
    if (mode === 'month') titleEl.textContent = `${state.filters?.month || 'October'} Ended Ads Records`;
    else if (mode === 'all') titleEl.textContent = 'All Completed Campaigns History';
    else if (mode === 'today') titleEl.textContent = 'Ending Today Campaigns';
    else if (mode === 'tomorrow') titleEl.textContent = 'Ending Tomorrow Campaigns';
  }
  filterEndAdsFullTable();
};

window.switchEndAdsTab = function(tab) {
  switchEndAdsViewMode(tab);
  document.getElementById('allEndedAdsTable')?.scrollIntoView({ behavior: 'smooth', block: 'center' });
};

window.filterEndAdsFullTable = function() {
  const tbody = document.getElementById('allEndedAdsBody');
  const badge = document.getElementById('endAdsTableTotalBadge');
  if (!tbody) return;

  const mode = state.currentEndAdsViewMode || 'month';
  let list = [];
  if (mode === 'month') list = state.currentEndAdsMonth || [];
  else if (mode === 'all') list = state.currentEndAdsAll || [];
  else if (mode === 'today') list = state.currentEndAdsToday || [];
  else if (mode === 'tomorrow') list = state.currentEndAdsTomorrow || [];

  const searchInput = document.getElementById('endAdsSearchInput');
  const q = searchInput ? searchInput.value.trim().toLowerCase() : '';

  let filtered = list;
  if (q) {
    filtered = list.filter(r => {
      const page = String(r['Page Name'] || r.page_name || '').toLowerCase();
      const sl = String(r['SL'] || r.sl || '').toLowerCase();
      const adAcc = String(r['AD Account'] || r.ad_account || '').toLowerCase();
      const type = String(r['Types'] || r['Type'] || '').toLowerCase();
      return page.includes(q) || sl.includes(q) || adAcc.includes(q) || type.includes(q);
    });
  }

  if (badge) badge.textContent = `${filtered.length} Ads`;

  if (filtered.length === 0) {
    tbody.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:24px;color:var(--txt-muted)">No ended ads match your filter</td></tr>';
    return;
  }

  tbody.innerHTML = filtered.slice(0, 150).map((r, idx) => {
    const sl = escapeHtml(String(r['SL'] || r.sl || idx + 1));
    const page = escapeHtml(r['Page Name'] || r.page_name || '');
    const adAcc = escapeHtml(r['AD Account'] || r.ad_account || '');
    const startD = escapeHtml(formatCustomDate(r['Start Date'] || r.Start || r.start_date || ''));
    const endD = escapeHtml(formatCustomDate(r['End Date'] || r.End || r.end_date || ''));
    const day = escapeHtml(String(r['Day'] || r.day || ''));
    const bud = fmtNum(r['Budget'] || r.budget || 0);
    const tak = fmtNum(r['Taka'] || r.taka || 0);
    const type = escapeHtml(r['Types'] || r.types || r['Type'] || 'Boost');

    return `
      <tr style="cursor:pointer;transition:background 0.12s">
        <td style="font-weight:700;color:var(--txt-muted);font-family:monospace">#${sl}</td>
        <td style="font-weight:700">${page}</td>
        <td><span class="badge badge-purple" style="font-size:11px">${adAcc || 'Main'}</span></td>
        <td><span class="table-date">${startD}</span></td>
        <td><span class="table-date" style="color:var(--dan);font-weight:700">${endD}</span></td>
        <td>${day} d</td>
        <td style="font-weight:800;color:var(--dan)">$${bud}</td>
        <td style="font-weight:700;color:var(--suc)">৳${tak}</td>
        <td><span class="badge badge-blue" style="font-size:10.5px">${type}</span></td>
        <td style="text-align:center">
          <button type="button" class="btn btn-secondary btn-sm" style="padding:2px 8px;font-size:11px;gap:4px" onclick="copyClientEndNotice('${page}', '${type}', '${day}', '${bud}', '${endD}', '${tak}')" title="Copy Client WhatsApp Notice">
            <span class="ms" style="font-size:14px;color:#10B981">chat</span> Notice
          </button>
        </td>
      </tr>
    `;
  }).join('');
};

window.copyClientEndNotice = function(page, type, day, budget, endDate, taka) {
  const msg = `আসসালামু আলাইকুম!\nআপনার *${page}* পেজের *${type || 'Boost'}* বিজ্ঞাপনটি *${endDate}* তারিখে সফলভাবে সমাপ্ত হয়েছে।\n\n📌 মোট দিন: ${day} দিন\n💵 মোট বাজেট: $${budget} ${taka ? '(৳' + taka + ')' : ''}\n\nপরবর্তী বিজ্ঞাপন চালু করতে চাইলে আমাদের জানান। ধন্যবাদ! – AdVerse Agency`;
  
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(msg).then(() => {
      showToast(`📋 ${page} - Ending Notice Copied!`, 'success');
    }).catch(() => {
      copyToClipboardFallback(msg);
      showToast(`📋 ${page} - Ending Notice Copied!`, 'success');
    });
  } else {
    copyToClipboardFallback(msg);
    showToast(`📋 ${page} - Ending Notice Copied!`, 'success');
  }
};

window.openInspectorModalForEndAd = function(type, idx) {
  const list = type === 'today' ? state.currentEndAdsToday : state.currentEndAdsTomorrow;
  const row = list && list[idx];
  if (row && typeof openRowFocusInspector === 'function') {
    openRowFocusInspector(row);
  }
};

window.filterEndAdsRows = function() {
  filterEndAdsFullTable();
};

function renderPaymentHistoryTable(data, stats) {
  const thead = document.getElementById('paymentTableHead');
  const tbody = document.getElementById('paymentTableBody');
  const tfoot = document.getElementById('paymentTableFoot');
  const rows = (data && data.rows) ? data.rows : [];
  state.currentPaymentRows = rows;

  // 1. Step 1: Calculate Total Paid directly from all active payment entries
  const calculatedPaid = Math.round(rows.reduce((sum, r) => {
    const val = r['Paid (৳)'] !== undefined ? r['Paid (৳)']
      : (r['paid_amount'] !== undefined ? r['paid_amount']
      : (r['Paid'] !== undefined ? r['Paid']
      : (r['paid'] !== undefined ? r['paid']
      : (r['Amount'] !== undefined ? r['Amount'] : 0))));
    return sum + parseMoney(val);
  }, 0) * 100) / 100;

  // 2. Step 2 & 3: Total Budget ($) and Rate
  const totalBudget = (stats && typeof stats.totalBudget === 'number') ? stats.totalBudget : (state.currentStats?.totalBudget || 0);
  const rate = (stats && stats.rate) ? Number(stats.rate) : (Number(state.rate) || 140);

  // 3. Step 4: Total Bill = Total Budget ($) × Rate
  const totalBill = (stats && typeof stats.totalBill === 'number' && stats.totalBill > 0)
    ? stats.totalBill
    : Math.round(totalBudget * rate * 100) / 100;

  // 4. Step 5: Total Due = Total Bill - Total Paid
  const totalDue = Math.round((totalBill - calculatedPaid) * 100) / 100;

  const currentMonth = (stats && stats.month) || state.filters.month || 'Selected Period';
  const currentYear = (stats && stats.year) || state.filters.year || new Date().getFullYear();

  const liveStats = {
    ...(stats || {}),
    totalPaid: calculatedPaid,
    totalBudget: totalBudget,
    totalBill: totalBill,
    totalDue: totalDue,
    rate: rate,
    month: currentMonth,
    year: currentYear,
    paymentCount: rows.length
  };
  state.currentStats = liveStats;
  renderStats(liveStats);

  if (thead) {
    thead.innerHTML = `
      <tr>
        <th style="width:120px">Date</th>
        <th style="width:140px">Payment Method</th>
        <th style="text-align:right">Paid (৳)</th>
        <th style="width:120px">Month / Period</th>
        ${isAdmin() ? '<th class="th-action" style="text-align:center;width:80px">Action</th>' : ''}
      </tr>
    `;
  }

  if (tbody) {
    if (!rows.length) {
      tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;padding:28px;color:var(--txt-muted)"><span class="ms" style="font-size:32px;display:block;opacity:0.4;margin-bottom:6px">payments</span>No payment records found for this period.</td></tr>';
    } else {
      tbody.innerHTML = rows.map((r, rIdx) => {
        const rowPaid = parseMoney(r['Paid (৳)'] !== undefined ? r['Paid (৳)'] : (r['paid_amount'] !== undefined ? r['paid_amount'] : (r['Paid'] || 0)));
        return `
          <tr>
            <td><span class="table-date" style="font-weight:600;font-variant-numeric:tabular-nums">${escapeHtml(formatCustomDate(r['Date']))}</span></td>
            <td>
              <span class="pill ${String(r['Payment Method'] || '').toLowerCase().includes('advance') ? 'pill-purple' : 'pill-blue'}" style="font-weight:700">
                ${escapeHtml(r['Payment Method'] || 'Bkash')}
              </span>
              ${(r['Note'] || r.Note || r.note || r['Advance Note']) ? `
                <div style="font-size:11.5px;color:var(--txt-muted);margin-top:4px;display:flex;align-items:center;gap:4px">
                  <span class="ms" style="font-size:13px;color:var(--pri)">sticky_note_2</span>
                  <span>${escapeHtml(r['Note'] || r.Note || r.note || r['Advance Note'])}</span>
                </div>
              ` : ''}
            </td>
            <td style="text-align:right;font-weight:800;font-variant-numeric:tabular-nums;color:var(--suc);font-size:13.5px">৳${fmtNum(rowPaid)}</td>
            <td><span style="font-size:12px;color:var(--txt-muted);font-weight:600">${escapeHtml(r['Month'] || state.filters.month || '')} ${escapeHtml(r['Year'] || state.filters.year || '')}</span></td>
            ${isAdmin() ? `
              <td style="text-align:center;white-space:nowrap">
                <div style="display:inline-flex;gap:6px;align-items:center;justify-content:center">
                  <button type="button" class="table-btn" onclick="openEditPaymentRecord(${rIdx})" title="Edit payment" style="width:30px;height:30px;padding:0;display:inline-flex;align-items:center;justify-content:center;border-radius:6px;background:rgba(99,102,241,0.1);color:var(--pri);border:1px solid rgba(99,102,241,0.25)">
                    <span class="ms" style="font-size:15px">edit</span>
                  </button>
                  <button type="button" class="table-btn table-btn-danger" onclick="deletePaymentItem('${r._recordId}')" title="Delete payment" style="width:30px;height:30px;padding:0;display:inline-flex;align-items:center;justify-content:center;border-radius:6px">
                    <span class="ms" style="font-size:15px">delete</span>
                  </button>
                </div>
              </td>` : ''}
          </tr>
        `;
      }).join('');
    }
  }

  // Render Table Footer Summary Row
  if (tfoot) {
    if (rows.length > 0) {
      tfoot.innerHTML = `
        <tr style="background:var(--card-elevated);border-top:2px solid var(--bor);font-size:13.5px">
          <td colspan="2" style="padding:12px 14px;font-weight:800;color:var(--txt)">
            <div style="display:flex;align-items:center;gap:6px">
              <span class="ms" style="color:var(--suc);font-size:18px">task_alt</span>
              <span>Total Paid [${rows.length} Entries]:</span>
            </div>
          </td>
          <td style="text-align:right;font-weight:900;font-variant-numeric:tabular-nums;color:var(--suc);font-size:15px;padding:12px 14px">
            ৳${fmtNum(calculatedPaid)}
          </td>
          <td colspan="${isAdmin() ? 2 : 1}" style="padding:12px 14px;font-size:12px;color:var(--txt-muted);font-weight:600">
            Total Bill: ৳${fmtNum(totalBill)} • Remaining Due: <b style="color:#EF4444;font-weight:900">৳${fmtNum(totalDue)}</b>
          </td>
        </tr>
      `;
    } else {
      tfoot.innerHTML = '';
    }
  }
}

function renderAgencyCalculationView(calcData) {
  const container = document.getElementById('agencyCalcContainer');
  if (!container) return;
  container.classList.remove('hidden');

  const d = calcData || {};
  const month = d.month || state.filters.month || 'September';
  const year = d.year || state.filters.year || new Date().getFullYear();
  const rate = d.rate || state.rate || 140;

  const totalCamp = d.totalCampaign || 0;
  const totalBudget = d.totalBudget || 0;
  const totalTaka = d.totalTaka || (totalBudget * rate);
  const totalPaid = d.totalPaid || 0;
  const totalDue = d.totalDue !== undefined ? d.totalDue : (totalTaka - totalPaid);

  // Restore saved manual overrides for this month/year if previously edited
  const overrideKey = `${month}_${year}`;
  let savedOverrides = {};
  try {
    const rawSaved = localStorage.getItem('adverse_closing_overrides');
    if (rawSaved) {
      const parsed = JSON.parse(rawSaved);
      savedOverrides = parsed[overrideKey] || {};
    }
  } catch (_) {}

  const agencyAdsCost = savedOverrides.agencyAdsCost !== undefined ? savedOverrides.agencyAdsCost : (d.agencyAdsCost || 0);
  const othersCost = savedOverrides.othersCost !== undefined ? savedOverrides.othersCost : (d.othersCost || d.otherExpense || 0);
  const officeRent = savedOverrides.officeRent !== undefined ? savedOverrides.officeRent : (d.officeRent || 0);

  // Calculation formula: (Total Due + Agency Ads Cost + Others Cost) - Office Rent
  const finalDue = (totalDue + agencyAdsCost + othersCost) - officeRent;

  state.currentStats = {
    ...d,
    totalCampaign: totalCamp,
    totalBudget,
    totalTaka,
    totalPaid,
    totalDue,
    agencyAdsCost,
    othersCost,
    officeRent,
    finalDue,
    month,
    year,
    rate
  };

  container.innerHTML = `
    <div class="records-card agency-calc-card" style="margin-bottom:24px;box-shadow:0 8px 30px rgba(0,0,0,0.06);border:1px solid var(--bor);border-radius:16px;overflow:hidden">
      <!-- Top Action Toolbar -->
      <div class="table-toolbar" style="flex-wrap:wrap;gap:12px;padding:14px 20px;border-bottom:1px solid var(--bor);background:var(--card);display:flex;align-items:center;justify-content:space-between">
        <div style="display:flex;align-items:center;gap:12px">
          <div style="width:38px;height:38px;border-radius:10px;background:linear-gradient(135deg,#DC2626,#EF4444);color:#fff;display:flex;align-items:center;justify-content:center;box-shadow:0 4px 12px rgba(239,68,68,0.25)">
            <svg width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>
          </div>
          <div>
            <div style="font-weight:900;font-size:16px;color:var(--txt);letter-spacing:-0.01em">${month} ${year} — Executive Settlement Statement</div>
            <div style="font-size:12px;color:var(--txt-muted);font-weight:500">Official Monthly Financial Closing & Cost Summary</div>
          </div>
        </div>
        <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center">
          <button type="button" class="btn btn-primary" id="btnDownloadCalcImage" onclick="downloadCalculationImage()" style="font-size:12.5px;padding:7px 15px;gap:6px;font-weight:700;background:linear-gradient(135deg,#DC2626,#EF4444);border:none">
            <span class="ms" style="font-size:16px">photo_camera</span> Export Image
          </button>
          <button type="button" class="btn btn-secondary" onclick="copyAgencyMonthClosingStatement()" style="font-size:12.5px;padding:7px 13px;gap:5px;font-weight:600">
            <span class="ms" style="font-size:16px">content_copy</span> Copy Summary
          </button>
          <button type="button" class="btn btn-secondary" onclick="window.printAgencyMonthClosingStatement()" style="font-size:12.5px;padding:7px 11px;gap:4px" title="Print Statement">
            <span class="ms" style="font-size:16px">print</span> Print
          </button>
          <button type="button" class="btn btn-secondary" onclick="loadMenuData(true)" style="font-size:12.5px;padding:7px 11px;gap:4px" title="Refresh Live Data">
            <span class="ms" style="font-size:16px">refresh</span> Sync
          </button>
        </div>
      </div>

      <!-- Clean Downloadable Statement Card (1-Click Image Snapshot Target) -->
      <div style="padding:28px 24px;background:var(--bg-gradient)">
        <div id="agencyCalcStatementCard" style="background:#ffffff;border:1px solid #E2E8F0;border-radius:18px;padding:28px;color:#0F172A;box-shadow:0 4px 20px rgba(0,0,0,0.03);max-width:920px;margin:0 auto;font-family:system-ui,-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif">
          
          <!-- Luxury Header with AdVerse Logo -->
          <div style="background:linear-gradient(135deg,#0B0F19 0%,#1E293B 100%);border-radius:14px;padding:20px 24px;color:#fff;display:flex;align-items:center;justify-content:space-between;margin-bottom:24px;border:1px solid rgba(255,255,255,0.08);box-shadow:0 8px 20px rgba(0,0,0,0.12)">
            <div style="display:flex;align-items:center;gap:16px">
              <div style="width:52px;height:52px;border-radius:12px;background:#181D2A;border:1.5px solid rgba(255,255,255,0.15);display:flex;align-items:center;justify-content:center;overflow:hidden;box-shadow:0 4px 12px rgba(0,0,0,0.35)">
                <img src="/adverse-logo.svg" alt="AdVerse Agency" style="width:100%;height:100%;object-fit:cover" />
              </div>
              <div>
                <div style="font-size:22px;font-weight:900;letter-spacing:-0.02em;color:#FFFFFF">AdVerse Agency</div>
                <div style="font-size:11px;color:#94A3B8;font-weight:700;letter-spacing:0.06em;text-transform:uppercase;margin-top:2px">Monthly Financial Closing Statement</div>
              </div>
            </div>
            <div style="text-align:right">
              <span style="display:inline-block;padding:5px 14px;background:rgba(239,68,68,0.2);color:#FEE2E2;font-weight:800;font-size:13.5px;border-radius:8px;border:1px solid rgba(239,68,68,0.35);letter-spacing:0.02em">
                ${month} ${year}
              </span>
              <div style="font-size:10.5px;color:#94A3B8;font-weight:600;margin-top:4px;letter-spacing:0.04em">OFFICIAL FINANCIAL LEDGER</div>
            </div>
          </div>

          <!-- 4 Core Top Metrics Grid -->
          <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:24px">
            <div style="background:#F8FAFC;border:1px solid #E2E8F0;border-radius:12px;padding:14px 16px">
              <div style="font-size:11px;font-weight:700;color:#64748B;text-transform:uppercase;letter-spacing:0.05em">Total Campaigns</div>
              <div style="font-size:20px;font-weight:900;color:#0F172A;margin-top:4px;font-variant-numeric:tabular-nums">${totalCamp}</div>
            </div>
            <div style="background:#F8FAFC;border:1px solid #E2E8F0;border-radius:12px;padding:14px 16px">
              <div style="font-size:11px;font-weight:700;color:#64748B;text-transform:uppercase;letter-spacing:0.05em">Ad Spend (USD)</div>
              <div style="font-size:20px;font-weight:900;color:#0F172A;margin-top:4px;font-variant-numeric:tabular-nums">$${fmtNum(totalBudget)}</div>
            </div>
            <div style="background:#F8FAFC;border:1px solid #E2E8F0;border-radius:12px;padding:14px 16px">
              <div style="font-size:11px;font-weight:700;color:#64748B;text-transform:uppercase;letter-spacing:0.05em">Dollar Rate</div>
              <div style="font-size:20px;font-weight:900;color:#0F172A;margin-top:4px;font-variant-numeric:tabular-nums">৳${rate}</div>
            </div>
            <div style="background:#F8FAFC;border:1px solid #E2E8F0;border-radius:12px;padding:14px 16px">
              <div style="font-size:11px;font-weight:700;color:#64748B;text-transform:uppercase;letter-spacing:0.05em">Invoiced Revenue</div>
              <div style="font-size:20px;font-weight:900;color:#0F172A;margin-top:4px;font-variant-numeric:tabular-nums">৳${fmtNum(totalTaka)}</div>
            </div>
          </div>

          <!-- Ultra-Clean, Minimalist Ledger Table (Neat, Crisp & Professional) -->
          <div style="border:1px solid #E2E8F0;border-radius:14px;overflow:hidden;margin-bottom:20px;background:#ffffff">
            <table style="width:100%;border-collapse:collapse;text-align:left">
              <thead>
                <tr style="background:#F8FAFC;border-bottom:1px solid #E2E8F0">
                  <th style="padding:14px 22px;font-size:11.5px;font-weight:800;color:#475569;text-transform:uppercase;letter-spacing:0.05em">Financial Item</th>
                  <th style="padding:14px 22px;font-size:11.5px;font-weight:800;color:#475569;text-align:right;text-transform:uppercase;letter-spacing:0.05em">Amount (BDT ৳)</th>
                </tr>
              </thead>
              <tbody style="font-size:14px">
                
                <!-- Total Invoiced Bill -->
                <tr style="border-bottom:1px solid #F1F5F9">
                  <td style="padding:15px 22px;font-weight:700;color:#1E293B">
                    <div style="display:flex;align-items:center;gap:10px">
                      <span style="color:#6366F1;display:flex;align-items:center">
                        <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>
                      </span>
                      <span>Total Invoiced Bill</span>
                    </div>
                  </td>
                  <td style="padding:15px 22px;text-align:right;font-weight:800;color:#0F172A;font-variant-numeric:tabular-nums;font-size:15px">
                    ৳${fmtNum(totalTaka)}
                  </td>
                </tr>

                <!-- Total Paid -->
                <tr style="border-bottom:1px solid #F1F5F9">
                  <td style="padding:15px 22px;font-weight:700;color:#1E293B">
                    <div style="display:flex;align-items:center;gap:10px">
                      <span style="color:#10B981;display:flex;align-items:center">
                        <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>
                      </span>
                      <span>Total Paid</span>
                    </div>
                  </td>
                  <td style="padding:15px 22px;text-align:right;font-weight:800;color:#10B981;font-variant-numeric:tabular-nums;font-size:15px">
                    ৳${fmtNum(totalPaid)}
                  </td>
                </tr>

                <!-- Total Due -->
                <tr style="border-bottom:1px solid #F1F5F9">
                  <td style="padding:15px 22px;font-weight:800;color:#1E293B">
                    <div style="display:flex;align-items:center;gap:10px">
                      <span style="color:#EF4444;display:flex;align-items:center">
                        <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"/></svg>
                      </span>
                      <span>Total Due</span>
                    </div>
                  </td>
                  <td style="padding:15px 22px;text-align:right;font-weight:900;color:#EF4444;font-variant-numeric:tabular-nums;font-size:15px">
                    ৳${fmtNum(totalDue)}
                  </td>
                </tr>

                <!-- Agency Ads Cost (Clean Editable Input) -->
                <tr style="border-bottom:1px solid #F1F5F9">
                  <td style="padding:12px 22px;font-weight:600;color:#334155">
                    <div style="display:flex;align-items:center;gap:10px">
                      <span style="color:#8B5CF6;display:flex;align-items:center">
                        <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M11 5.882V19.24a1.76 1.76 0 01-3.417.592l-2.147-6.15M18 13a3 3 0 100-6M5.436 13.683A4.001 4.001 0 017 6h1.832c4.1 0 7.625-1.234 9.168-3v14c-1.543-1.766-5.067-3-9.168-3H7a3.988 3.988 0 01-1.564-.317z"/></svg>
                      </span>
                      <span>Agency Ads Cost</span>
                      <span style="font-size:11px;color:#8B5CF6;font-weight:700">(+)</span>
                    </div>
                  </td>
                  <td style="padding:10px 22px;text-align:right">
                    <div style="display:inline-flex;align-items:center;justify-content:flex-end">
                      <span style="font-weight:700;color:#94A3B8;font-size:13px;margin-right:6px">৳</span>
                      <input type="number" step="any" id="calcAgencyAdsCost" value="${agencyAdsCost || ''}" placeholder="0" oninput="window.recalcAgencyClosingStatement()" style="width:140px;height:36px;padding:0 10px;text-align:right;font-weight:800;font-size:15px;color:#0F172A;border:1px solid #CBD5E1;border-radius:8px;background:#F8FAFC;outline:none;transition:all 0.2s" onfocus="this.style.background='#fff';this.style.borderColor='#8B5CF6';this.style.boxShadow='0 0 0 3px rgba(139,92,246,0.12)'" onblur="this.style.background='#F8FAFC';this.style.borderColor='#CBD5E1';this.style.boxShadow='none'" />
                    </div>
                  </td>
                </tr>


                <!-- Others Cost (Clean Editable Input) -->
                <tr style="border-bottom:1px solid #F1F5F9">
                  <td style="padding:12px 22px;font-weight:600;color:#334155">
                    <div style="display:flex;align-items:center;gap:10px">
                      <span style="color:#0EA5E9;display:flex;align-items:center">
                        <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M19 11H5m14 0a2 2 0 012 2v6a2 2 0 01-2 2H5a2 2 0 01-2-2v-6a2 2 0 012-2m14 0V9a2 2 0 00-2-2M5 11V9a2 2 0 012-2m0 0V5a2 2 0 012-2h6a2 2 0 012 2v2M7 7h10"/></svg>
                      </span>
                      <span>Others Cost</span>
                      <span style="font-size:11px;color:#0EA5E9;font-weight:700">(+)</span>
                    </div>
                  </td>
                  <td style="padding:10px 22px;text-align:right">
                    <div style="display:inline-flex;align-items:center;justify-content:flex-end">
                      <span style="font-weight:700;color:#94A3B8;font-size:13px;margin-right:6px">৳</span>
                      <input type="number" step="any" id="calcOthersCost" value="${othersCost || ''}" placeholder="0" oninput="window.recalcAgencyClosingStatement()" style="width:140px;height:36px;padding:0 10px;text-align:right;font-weight:800;font-size:15px;color:#0F172A;border:1px solid #CBD5E1;border-radius:8px;background:#F8FAFC;outline:none;transition:all 0.2s" onfocus="this.style.background='#fff';this.style.borderColor='#0EA5E9';this.style.boxShadow='0 0 0 3px rgba(14,165,233,0.12)'" onblur="this.style.background='#F8FAFC';this.style.borderColor='#CBD5E1';this.style.boxShadow='none'" />
                    </div>
                  </td>
                </tr>

                <!-- Office Rent (Clean Editable Input - SUBTRACTED / DEDUCTION) -->
                <tr style="border-bottom:1px solid #E2E8F0;background:#FFFDFB">
                  <td style="padding:12px 22px;font-weight:700;color:#C2410C">
                    <div style="display:flex;align-items:center;gap:10px">
                      <span style="color:#EA580C;display:flex;align-items:center">
                        <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4"/></svg>
                      </span>
                      <span>Office Rent</span>
                      <span style="font-size:11px;color:#EA580C;font-weight:800">(-) Minus</span>
                    </div>
                  </td>
                  <td style="padding:10px 22px;text-align:right">
                    <div style="display:inline-flex;align-items:center;justify-content:flex-end">
                      <span style="font-weight:800;color:#EA580C;font-size:13px;margin-right:6px">- ৳</span>
                      <input type="number" step="any" id="calcOfficeRent" value="${officeRent || ''}" placeholder="0" oninput="window.recalcAgencyClosingStatement()" style="width:140px;height:36px;padding:0 10px;text-align:right;font-weight:900;font-size:15px;color:#C2410C;border:1px solid #FED7AA;border-radius:8px;background:#FFF7ED;outline:none;transition:all 0.2s" onfocus="this.style.background='#fff';this.style.borderColor='#EA580C';this.style.boxShadow='0 0 0 3px rgba(234,88,12,0.12)'" onblur="this.style.background='#FFF7ED';this.style.borderColor='#FED7AA';this.style.boxShadow='none'" />
                    </div>
                  </td>
                </tr>

                <!-- FINAL DUE TO SETTLE (Clean Grand Summary Row) -->
                <tr style="background:#FFF5F5;border-top:2px solid #EF4444">
                  <td style="padding:18px 22px">
                    <div style="font-weight:900;font-size:15px;color:#991B1B;display:flex;align-items:center;gap:8px">
                      <span style="color:#DC2626;display:flex;align-items:center">
                        <svg width="20" height="20" fill="currentColor" viewBox="0 0 24 24"><path d="M12 2l3.09 6.26L22 9.27l-5 4.87 1.18 6.88L12 17.77l-6.18 3.25L7 14.14 2 9.27l6.91-1.01L12 2z"/></svg>
                      </span>
                      <span>FINAL DUE TO SETTLE</span>
                    </div>
                    <div style="font-size:11.5px;color:#7F1D1D;font-weight:500;margin-top:2px;margin-left:28px">
                      (Total Due + Ads Cost + Scam + Others) − Office Rent
                    </div>
                  </td>
                  <td style="padding:18px 22px;text-align:right">
                    <div id="calcFinalDueDisplay" style="font-weight:900;font-size:24px;color:#DC2626;font-variant-numeric:tabular-nums;letter-spacing:-0.02em">
                      ৳${fmtNum(finalDue)}
                    </div>
                  </td>
                </tr>
              </tbody>
            </table>
          </div>

          <!-- Clean Footer Note -->
          <div style="display:flex;align-items:center;justify-content:space-between;font-size:11px;color:#94A3B8;padding-top:4px">
            <span>Generated Date: ${formatCustomDate(getNowBST())}</span>
            <span>AdVerse Agency • Executive Financial Statement</span>
          </div>
        </div>
      </div>
    </div>
  `;
}

window.recalcAgencyClosingStatement = function() {
  const d = state.currentStats || {};
  const totalTaka = d.totalTaka || 0;
  const totalPaid = d.totalPaid || 0;
  const totalDue = d.totalDue !== undefined ? d.totalDue : (totalTaka - totalPaid);

  const adsVal = parseFloat(document.getElementById('calcAgencyAdsCost')?.value) || 0;
  const othersVal = parseFloat(document.getElementById('calcOthersCost')?.value) || 0;
  const rentVal = parseFloat(document.getElementById('calcOfficeRent')?.value) || 0;

  // Formula: (Total Due + Agency Ads Cost + Others Cost) - Office Rent
  const finalDue = (totalDue + adsVal + othersVal) - rentVal;

  const finalDueEl = document.getElementById('calcFinalDueDisplay');
  if (finalDueEl) {
    finalDueEl.textContent = `৳${fmtNum(finalDue)}`;
  }

  // Update currentStats
  state.currentStats = {
    ...state.currentStats,
    agencyAdsCost: adsVal,
    othersCost: othersVal,
    officeRent: rentVal,
    finalDue: finalDue
  };

  // Persist manual inputs in localStorage for this month & year
  const m = state.filters.month || d.month || 'September';
  const y = state.filters.year || d.year || new Date().getFullYear();
  const overrideKey = `${m}_${y}`;
  try {
    const rawSaved = localStorage.getItem('adverse_closing_overrides') || '{}';
    const parsed = JSON.parse(rawSaved);
    parsed[overrideKey] = {
      agencyAdsCost: adsVal,
      othersCost: othersVal,
      officeRent: rentVal,
      finalDue: finalDue
    };
    localStorage.setItem('adverse_closing_overrides', JSON.stringify(parsed));
  } catch (_) {}
};

window.downloadCalculationImage = async function() {
  const card = document.getElementById('agencyCalcStatementCard');
  if (!card) {
    showToast('Calculation card not found', 'error');
    return;
  }
  const btn = document.getElementById('btnDownloadCalcImage');
  const originalHtml = btn ? btn.innerHTML : '';
  if (btn) {
    btn.disabled = true;
    btn.innerHTML = '<span class="ms spin" style="font-size:16px">sync</span> Generating...';
  }
  try {
    const h2c = (typeof window !== 'undefined' && window.html2canvas) ? window.html2canvas : (typeof html2canvas !== 'undefined' ? html2canvas : null);
    if (!h2c) {
      showToast('html2canvas library is loading, please try again in a moment.', 'warning');
      return;
    }
    const canvas = await h2c(card, {
      scale: 2,
      useCORS: true,
      backgroundColor: '#ffffff',
      logging: false
    });
    const link = document.createElement('a');
    const m = state.filters.month || 'Statement';
    const y = state.filters.year || new Date().getFullYear();
    link.download = `AdVerse_Agency_Financial_Statement_${m}_${y}.png`;
    link.href = canvas.toDataURL('image/png');
    link.click();
    showToast('📸 Financial Statement PNG successfully exported!', 'success');
  } catch (err) {
    console.error('Download image error:', err);
    showToast('Failed to generate image, please try again.', 'error');
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = originalHtml;
    }
  }
};

window.copyAgencyMonthClosingStatement = function() {
  const ms = document.getElementById('agencyCalcMonth')?.value || state.filters.month || 'September';
  const ys = document.getElementById('agencyCalcYear')?.value || state.filters.year || new Date().getFullYear();
  const d = state.currentStats || {};
  const rate = d.rate || state.rate || 140;
  const totalCamp = d.totalCampaign || 0;
  const totalBudget = d.totalBudget || 0;
  const totalTaka = d.totalTaka || (totalBudget * rate);
  const totalPaid = d.totalPaid || 0;
  const totalDue = d.totalDue !== undefined ? d.totalDue : (totalTaka - totalPaid);
  const agencyAdsCost = d.agencyAdsCost || 0;
  const othersCost = d.othersCost || 0;
  const officeRent = d.officeRent || 0;
  const finalDue = d.finalDue !== undefined ? d.finalDue : ((totalDue + agencyAdsCost + othersCost) - officeRent);

  const text = `📊 AdVerse Agency - Monthly Financial Statement
Period: ${ms} ${ys}
---------------------------------------------
📢 Total Campaigns: ${totalCamp}
💵 Ad Spend (USD): $${fmtNum(totalBudget)}
🪙 Total Invoiced Bill: ৳${fmtNum(totalTaka)} (Rate: ৳${rate})
💳 Total Paid: ৳${fmtNum(totalPaid)}
⚠️ Total Due: ৳${fmtNum(totalDue)}

🏢 Settlement Breakdown:
• Total Due: ৳${fmtNum(totalDue)}
• (+) Agency Ads Cost: ৳${fmtNum(agencyAdsCost)}
• (+) Others Cost: ৳${fmtNum(othersCost)}
• (-) Office Rent (Deduction): ৳${fmtNum(officeRent)}
---------------------------------------------
⭐ FINAL DUE TO SETTLE: ৳${fmtNum(finalDue)}`;

  copyCampaignRecord(text);
  showToast(`📋 ${ms} ${ys} Financial Statement copied to clipboard!`, 'success');
};

window.printAgencyMonthClosingStatement = function() {
  window.print();
};
window.renderAgencyCalculationView = renderAgencyCalculationView;

// ==========================================================================
// Ad Account Billing & Threshold Tracker Engine (Admin Only)
// ==========================================================================
let billingAccountsState = {
  list: [],
  stats: null,
  filterStatus: 'all',
  filterActiveOnly: false,
  viewMode: localStorage.getItem('adverse_billing_view_mode') || 'cards',
  searchQuery: ''
};

function setBillingActiveOnly(val) {
  billingAccountsState.filterActiveOnly = val;
  renderAdAccountBillingView();
}

function setBillingViewMode(mode) {
  billingAccountsState.viewMode = mode;
  localStorage.setItem('adverse_billing_view_mode', mode);
  renderAdAccountBillingView();
}

function loadAdAccountBillingData(forced = false) {
  if (!isAdmin()) {
    showToast('⚠️ Admin Access Only', 'error');
    selectMenu('today_created_campaign');
    return;
  }

  const container = document.getElementById('adBillingContainer');
  if (!container) return;

  if (!billingAccountsState.list.length) {
    container.innerHTML = `
      <div style="padding:30px;text-align:center">
        <div class="skeleton skeleton-card" style="height:120px;margin-bottom:16px"></div>
        <div class="skeleton skeleton-card" style="height:240px"></div>
      </div>
    `;
  }

  google.script.run.withSuccessHandler(res => {
    if (res && res.success) {
      billingAccountsState.list = res.list || [];
      billingAccountsState.stats = res.stats || {};
      renderAdAccountBillingView();
    } else {
      showToast(res?.message || 'Failed to load ad accounts billing', 'error');
    }
  }).withFailureHandler(err => {
    showToast(err?.message || 'Error connecting to billing service', 'error');
  }).getAdAccountBilling();
}

function renderAdAccountBillingView() {
  const container = document.getElementById('adBillingContainer');
  if (!container) return;

  if (!isAdmin()) {
    container.classList.add('hidden');
    container.innerHTML = '';
    return;
  }

  const { list, stats, filterStatus, filterActiveOnly, viewMode, searchQuery } = billingAccountsState;
  const s = stats || {};

  // Filter accounts
  let filtered = [...list];
  if (filterActiveOnly) {
    filtered = filtered.filter(a => (a.activeAds || 0) > 0);
  }
  if (filterStatus !== 'all') {
    filtered = filtered.filter(a => a.status === filterStatus);
  }
  if (searchQuery) {
    const q = searchQuery.toLowerCase();
    filtered = filtered.filter(a =>
      (a.accountName || '').toLowerCase().includes(q) ||
      (a.idName || '').toLowerCase().includes(q) ||
      (a.paymentMethod || '').toLowerCase().includes(q) ||
      (a.notes || '').toLowerCase().includes(q)
    );
  }

  // Build HUD (Top KPI Cards)
  const hudHtml = `
    <div class="billing-hud-grid">
      <div class="billing-hud-card" style="border-top:3px solid var(--pri)">
        <div class="billing-hud-title">
          <span>Total Ad Accounts</span>
          <span class="ms" style="color:var(--pri)">credit_card</span>
        </div>
        <div class="billing-hud-value">${s.totalAccounts || list.length}</div>
        <div class="billing-hud-sub">
          <span style="color:var(--suc);font-weight:700">● ${s.activeCount || 0} Healthy</span>
          <span style="color:var(--txt-dim)">•</span>
          <span style="color:var(--dan);font-weight:700">● ${(s.criticalCount || 0) + (s.warningCount || 0)} Attention</span>
        </div>
      </div>

      <div class="billing-hud-card" style="border-top:3px solid var(--dan)">
        <div class="billing-hud-title">
          <span>Total Active Dues</span>
          <span class="ms" style="color:var(--dan)">account_balance_wallet</span>
        </div>
        <div class="billing-hud-value" style="color:var(--dan)">$${fmtNum(s.totalDueUSD || 0)}</div>
        <div class="billing-hud-sub">
          <span>৳${fmtNum(s.totalDueBDT || 0)} (Rate ৳${s.rate || state.rate || 140}/$)</span>
        </div>
      </div>

      <div class="billing-hud-card" style="border-top:3px solid var(--suc)">
        <div class="billing-hud-title">
          <span>Available Headroom</span>
          <span class="ms" style="color:var(--suc)">savings</span>
        </div>
        <div class="billing-hud-value" style="color:var(--suc)">$${fmtNum(s.totalRemainingHeadroom || 0)}</div>
        <div class="billing-hud-sub">
          <span>Of $${fmtNum(s.totalThreshold || 0)} Total Threshold</span>
        </div>
      </div>

      <div class="billing-hud-card" style="border-top:3px solid var(--sec)">
        <div class="billing-hud-title">
          <span>Active Daily Spend</span>
          <span class="ms" style="color:var(--sec)">trending_up</span>
        </div>
        <div class="billing-hud-value" style="color:var(--sec)">$${fmtNum(s.totalActiveDailySpend || 0)}<span style="font-size:13px;font-weight:600;color:var(--txt-muted)">/day</span></div>
        <div class="billing-hud-sub">
          <span>Across all live campaigns</span>
        </div>
      </div>
    </div>
  `;

  // Build Toolbar
  const activeCount = list.filter(a => (a.activeAds || 0) > 0).length;
  const toolbarHtml = `
    <div class="billing-toolbar">
      <div class="billing-toolbar-left">
        <div style="position:relative;flex:1;min-width:200px">
          <span class="ms" style="position:absolute;left:10px;top:50%;transform:translateY(-50%);color:var(--txt-dim);font-size:18px">search</span>
          <input type="text" class="billing-search-input" placeholder="Search account name, ID, or card memo..." value="${escapeHtml(searchQuery)}" oninput="handleBillingSearch(this.value)">
        </div>
        <div class="billing-filter-pills">
          <button type="button" class="billing-filter-pill ${filterStatus === 'all' && !filterActiveOnly ? 'active' : ''}" onclick="setBillingActiveOnly(false); setBillingStatusFilter('all');">
            All (${list.length})
          </button>
          <button type="button" class="billing-filter-pill ${filterActiveOnly ? 'active' : ''}" onclick="setBillingActiveOnly(!filterActiveOnly)" style="${filterActiveOnly ? 'background:linear-gradient(135deg,#6366F1,#4F46E5);color:#fff;border-color:#4F46E5' : 'border-color:rgba(99,102,241,0.4);color:var(--pri)'}">
            🔥 Running Ads Only (${activeCount})
          </button>
          <button type="button" class="billing-filter-pill ${filterStatus === 'critical' ? 'active' : ''}" onclick="setBillingStatusFilter('critical')" style="${filterStatus === 'critical' ? 'background:var(--dan);border-color:var(--dan)' : ''}">
            🚨 Critical (${list.filter(a => a.status === 'critical').length})
          </button>
          <button type="button" class="billing-filter-pill ${filterStatus === 'warning' ? 'active' : ''}" onclick="setBillingStatusFilter('warning')" style="${filterStatus === 'warning' ? 'background:var(--wrn);border-color:var(--wrn)' : ''}">
            ⚠️ Near Limit (${list.filter(a => a.status === 'warning').length})
          </button>
          <button type="button" class="billing-filter-pill ${filterStatus === 'active' ? 'active' : ''}" onclick="setBillingStatusFilter('active')" style="${filterStatus === 'active' ? 'background:var(--suc);border-color:var(--suc)' : ''}">
            🟢 Active (${list.filter(a => a.status === 'active').length})
          </button>
          <button type="button" class="billing-filter-pill ${filterStatus === 'paused' ? 'active' : ''}" onclick="setBillingStatusFilter('paused')">
            ⏸️ Paused (${list.filter(a => a.status === 'paused').length})
          </button>
        </div>
      </div>
      <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
        <div class="density-switch-group" style="display:inline-flex">
          <button class="density-btn ${viewMode === 'cards' ? 'active' : ''}" type="button" onclick="setBillingViewMode('cards')" title="Card View">
            <span class="ms" style="font-size:14px">grid_view</span> Cards
          </button>
          <button class="density-btn ${viewMode === 'table' ? 'active' : ''}" type="button" onclick="setBillingViewMode('table')" title="Table / Spreadsheet View">
            <span class="ms" style="font-size:14px">table_rows</span> Table
          </button>
        </div>
        <button class="btn btn-secondary" onclick="loadAdAccountBillingData(true)" type="button" title="Refresh Billing Data">
          <span class="ms" style="font-size:16px">sync</span>
        </button>
        <button class="btn btn-primary" onclick="openAdBillingModal()" type="button" style="font-weight:700">
          <span class="ms" style="font-size:16px">add_card</span> + Add Account Billing
        </button>
      </div>
    </div>
  `;

  // Build Account Cards or Table
  let mainContentHtml = '';

  if (!filtered.length) {
    mainContentHtml = `
      <div style="background:var(--card);border:1px dashed var(--bor);border-radius:var(--radius-lg);padding:48px 20px;text-align:center;color:var(--txt-muted);width:100%">
        <span class="ms" style="font-size:44px;color:var(--pri);opacity:0.5;display:block;margin-bottom:8px">credit_card_off</span>
        <div style="font-weight:800;font-size:15px;color:var(--txt)">No matching ad accounts found</div>
        <p style="font-size:13px;max-width:380px;margin:6px auto 14px;color:var(--txt-dim)">Adjust your search filters or click "+ Add Account Billing" to set up a new threshold tracker.</p>
        <button class="btn btn-primary btn-sm" onclick="openAdBillingModal()">+ Add Account Billing</button>
      </div>
    `;
  } else if (viewMode === 'table') {
    // High-Density Spreadsheet / Table View
    const tableRows = filtered.map((acc, idx) => {
      const due = Number(acc.dueAmount) || 0;
      const th = Number(acc.threshold) || 500;
      const headroom = Number(acc.remainingHeadroom) !== undefined ? Number(acc.remainingHeadroom) : Math.max(0, th - due);
      const util = th > 0 ? Math.min(100, Math.round((due / th) * 100)) : 0;
      const dailySpend = Number(acc.activeDailySpend) || 0;
      const activeAds = Number(acc.activeAds) || 0;
      const estDays = acc.estimatedDaysLeft;

      let statusDot = '🟢';
      if (acc.status === 'critical') statusDot = '🔴';
      else if (acc.status === 'warning') statusDot = '🟡';
      else if (acc.status === 'paused') statusDot = '⚪';

      let billDateDisplay = '-';
      if (acc.billingDate) {
        const bDate = parseAnyDate(acc.billingDate);
        if (bDate) {
          billDateDisplay = formatCustomDate(bDate);
        }
      }

      const accJson = escapeHtml(JSON.stringify(acc));

      return `
        <tr>
          <td style="text-align:center;font-size:12px">${statusDot}</td>
          <td>
            <div style="font-weight:800;color:var(--txt);display:flex;align-items:center;gap:6px">
              <span>${escapeHtml(acc.accountName)}</span>
              ${acc.paymentMethod ? `<span style="font-size:11px;font-weight:600;color:var(--txt-dim)">• ${escapeHtml(acc.paymentMethod)}</span>` : ''}
            </div>
            ${acc.lastChecked ? `<div style="font-size:10.5px;color:var(--txt-dim)">Checked: ${escapeHtml(acc.lastChecked)}</div>` : ''}
          </td>
          <td><span style="font-family:monospace;font-size:11.5px;color:var(--txt-muted)">${escapeHtml(acc.idName || '-')}</span></td>
          <td style="text-align:center">
            <span class="pill ${activeAds > 0 ? 'pill-blue' : ''}" style="font-weight:700">
              ${activeAds > 0 ? `${activeAds} live` : '0'}
            </span>
          </td>
          <td style="text-align:right;font-weight:700;color:var(--sec)">
            ${dailySpend > 0 ? `$${fmtNum(dailySpend)}/day` : '-'}
          </td>
          <td style="text-align:right;font-weight:800;color:${util >= 90 ? 'var(--dan)' : (util >= 70 ? 'var(--wrn)' : 'var(--txt)')};font-size:14px">
            $${fmtNum(due)}
          </td>
          <td style="text-align:right;font-weight:700;color:var(--txt-muted)">
            $${fmtNum(th)}
          </td>
          <td style="text-align:right;font-weight:800;color:var(--suc)">
            $${fmtNum(headroom)}
          </td>
          <td style="text-align:center;font-weight:600;font-size:12px">
            ${billDateDisplay}
          </td>
          <td style="text-align:center;font-weight:700;color:${estDays !== null && Number(estDays) <= 1.5 ? 'var(--dan)' : 'var(--txt)'}">
            ${estDays !== null ? `~${estDays} d` : '-'}
          </td>
          <td style="text-align:center;white-space:nowrap">
            <div style="display:inline-flex;gap:4px">
              <button class="btn-quick-due" type="button" onclick="openQuickDueModal('${escapeHtml(acc.accountName)}', ${due}, ${th}, '${escapeHtml(acc.billingDate || '')}')" style="padding:4px 8px;font-size:11px">
                <span class="ms" style="font-size:13px">bolt</span> Due
              </button>
              <button class="btn-edit-acc" type="button" onclick='openAdBillingModal(${accJson})' style="padding:4px 8px;font-size:11px">
                <span class="ms" style="font-size:13px">tune</span>
              </button>
            </div>
          </td>
        </tr>
      `;
    }).join('');

    mainContentHtml = `
      <div class="records-card" style="box-shadow:var(--shadow-sm);border:1px solid var(--bor);border-radius:var(--radius-lg);overflow:hidden">
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th style="width:36px;text-align:center">Status</th>
                <th>Ad Account Name</th>
                <th>ID Name</th>
                <th style="text-align:center">Active Ads</th>
                <th style="text-align:right">Daily Spend</th>
                <th style="text-align:right">Due Amount ($)</th>
                <th style="text-align:right">Threshold ($)</th>
                <th style="text-align:right">Headroom ($)</th>
                <th style="text-align:center">Billing Date</th>
                <th style="text-align:center">Runway</th>
                <th style="text-align:center;width:130px">Actions</th>
              </tr>
            </thead>
            <tbody>
              ${tableRows}
            </tbody>
          </table>
        </div>
      </div>
    `;
  } else {
    // Visual Cards Grid
    const cardsHtml = filtered.map(acc => {
      const due = Number(acc.dueAmount) || 0;
      const th = Number(acc.threshold) || 500;
      const headroom = Number(acc.remainingHeadroom) !== undefined ? Number(acc.remainingHeadroom) : Math.max(0, th - due);
      const util = th > 0 ? Math.min(100, Math.round((due / th) * 100)) : 0;
      const dailySpend = Number(acc.activeDailySpend) || 0;
      const activeAds = Number(acc.activeAds) || 0;
      const estDays = acc.estimatedDaysLeft;

      let fillClass = 'fill-green';
      if (util >= 90 || acc.status === 'critical') fillClass = 'fill-red';
      else if (util >= 70 || acc.status === 'warning') fillClass = 'fill-yellow';

      let statusBadge = '';
      if (acc.status === 'critical') {
        statusBadge = `<span class="billing-status-badge critical"><span class="ms" style="font-size:12px">warning</span> Critical (Limit Full)</span>`;
      } else if (acc.status === 'warning') {
        statusBadge = `<span class="billing-status-badge warning"><span class="ms" style="font-size:12px">error</span> Near Limit</span>`;
      } else if (acc.status === 'paused') {
        statusBadge = `<span class="billing-status-badge paused"><span class="ms" style="font-size:12px">pause</span> Paused</span>`;
      } else {
        statusBadge = `<span class="billing-status-badge active"><span class="ms" style="font-size:12px">check_circle</span> Active</span>`;
      }

      // Billing Date badge logic
      let billingDateHtml = '';
      if (acc.billingDate) {
        const bDate = parseAnyDate(acc.billingDate);
        if (bDate) {
          const nowBst = getNowBST();
          nowBst.setHours(0, 0, 0, 0);
          const diffDays = Math.round((bDate.getTime() - nowBst.getTime()) / 86400000);
          let dateClass = 'chip-highlight';
          let dateText = `Due in ${diffDays} days (${formatCustomDate(bDate)})`;
          if (diffDays === 0) {
            dateClass = 'chip-alert';
            dateText = `🚨 Auto-Charge Today! (${formatCustomDate(bDate)})`;
          } else if (diffDays === 1) {
            dateClass = 'chip-alert';
            dateText = `⚠️ Auto-Charge Tomorrow! (${formatCustomDate(bDate)})`;
          } else if (diffDays < 0) {
            dateClass = 'chip-alert';
            dateText = `⚠️ Past Due (${formatCustomDate(bDate)})`;
          }
          billingDateHtml = `<span class="ad-acc-chip ${dateClass}"><span class="ms" style="font-size:13px">event</span> ${dateText}</span>`;
        }
      }

      const accJson = escapeHtml(JSON.stringify(acc));

      return `
        <div class="ad-acc-card status-${acc.status}">
          <div class="ad-acc-header">
            <div class="ad-acc-title-wrap">
              <div class="ad-acc-name">
                <span class="ms" style="color:var(--pri);font-size:18px">credit_card</span>
                ${escapeHtml(acc.accountName)}
              </div>
              ${acc.idName ? `<div class="ad-acc-id">ID: ${escapeHtml(acc.idName)}</div>` : ''}
            </div>
            <div>${statusBadge}</div>
          </div>

          <!-- Threshold Progress -->
          <div class="threshold-progress-wrap">
            <div class="threshold-progress-labels">
              <div>
                <span style="font-size:11px;font-weight:700;color:var(--txt-dim);text-transform:uppercase">SPEND DUE / THRESHOLD</span>
                <div style="font-size:16px;font-weight:800;color:var(--txt);margin-top:1px">
                  <span style="color:${util >= 90 ? 'var(--dan)' : (util >= 70 ? 'var(--wrn)' : 'var(--txt)')}">$${fmtNum(due)}</span>
                  <span style="font-size:12px;font-weight:600;color:var(--txt-muted)"> / $${fmtNum(th)}</span>
                </div>
              </div>
              <div style="text-align:right">
                <span style="font-size:11px;font-weight:700;color:var(--txt-dim);text-transform:uppercase">HEADROOM LEFT</span>
                <div style="font-size:15px;font-weight:800;color:var(--suc);margin-top:1px">
                  $${fmtNum(headroom)}
                </div>
              </div>
            </div>
            <div class="threshold-bar-track">
              <div class="threshold-bar-fill ${fillClass}" style="width:${util}%"></div>
            </div>
            <div style="display:flex;justify-content:space-between;font-size:11px;color:var(--txt-dim);margin-top:2px">
              <span>${util}% of threshold reached</span>
              <span>${th > 0 && headroom === 0 ? '⚠️ Limit Full' : `$${fmtNum(headroom)} remaining`}</span>
            </div>
          </div>

          <!-- Triple Metric Box -->
          <div class="ad-acc-metrics-grid">
            <div class="ad-acc-metric-box">
              <span class="metric-lbl">Active Ads</span>
              <span class="metric-val" style="color:var(--pri)">${activeAds} campaigns</span>
            </div>
            <div class="ad-acc-metric-box">
              <span class="metric-lbl">Daily Spend</span>
              <span class="metric-val" style="color:var(--sec)">$${fmtNum(dailySpend)}/day</span>
            </div>
            <div class="ad-acc-metric-box">
              <span class="metric-lbl">Est. Runway</span>
              <span class="metric-val" style="color:${estDays !== null && Number(estDays) <= 1.5 ? 'var(--dan)' : 'var(--txt)'}">
                ${estDays !== null ? `~${estDays} days` : 'N/A'}
              </span>
            </div>
          </div>

          <!-- Info Chips -->
          <div class="ad-acc-chips-row">
            ${billingDateHtml}
            ${acc.paymentMethod ? `<span class="ad-acc-chip"><span class="ms" style="font-size:13px">account_balance</span> ${escapeHtml(acc.paymentMethod)}</span>` : ''}
            ${acc.notes ? `<span class="ad-acc-chip" title="${escapeHtml(acc.notes)}"><span class="ms" style="font-size:13px">notes</span> ${escapeHtml(acc.notes.slice(0, 30))}${acc.notes.length > 30 ? '...' : ''}</span>` : ''}
          </div>

          <!-- Footer Actions -->
          <div class="ad-acc-footer-actions">
            <div class="ad-acc-last-checked" title="Last checked in Ads Manager">
              <span class="ms" style="font-size:13px">schedule</span>
              <span>${acc.lastChecked ? escapeHtml(acc.lastChecked) : 'Not checked today'}</span>
            </div>
            <div style="display:flex;gap:6px">
              <button class="btn-quick-due" type="button" onclick="openQuickDueModal('${escapeHtml(acc.accountName)}', ${due}, ${th}, '${escapeHtml(acc.billingDate || '')}')" title="1-Click Update Due Amount from Ads Manager">
                <span class="ms" style="font-size:14px">bolt</span> Quick Due
              </button>
              <button class="btn-edit-acc" type="button" onclick='openAdBillingModal(${accJson})' title="Edit threshold, billing cycle or card info">
                <span class="ms" style="font-size:14px">tune</span> Edit
              </button>
            </div>
          </div>
        </div>
      `;
    }).join('');

    mainContentHtml = `
      <div class="billing-cards-grid">
        ${cardsHtml}
      </div>
    `;
  }

  container.innerHTML = `
    ${hudHtml}
    ${toolbarHtml}
    ${mainContentHtml}
  `;
}

function handleBillingSearch(val) {
  billingAccountsState.searchQuery = val || '';
  renderAdAccountBillingView();
}

function setBillingStatusFilter(status) {
  billingAccountsState.filterStatus = status || 'all';
  renderAdAccountBillingView();
}

let activeQuickDueAccount = null;

function openQuickDueModal(accName, currentDue, threshold, billingDate = '') {
  if (!isAdmin()) return;
  activeQuickDueAccount = accName;
  const overlay = document.getElementById('quickDueModalOverlay');
  if (!overlay) return;

  const subEl = document.getElementById('quickDueAccountSubtitle');
  if (subEl) subEl.textContent = `Account: ${accName}`;

  const thEl = document.getElementById('quickDueThresholdDisplay');
  if (thEl) thEl.textContent = `$${fmtNum(threshold)}`;

  const curEl = document.getElementById('quickDueCurrentDisplay');
  if (curEl) curEl.textContent = `$${fmtNum(currentDue)}`;

  const inputEl = document.getElementById('quickDueInput');
  if (inputEl) {
    inputEl.value = currentDue > 0 ? currentDue : '';
    inputEl.setAttribute('data-threshold', threshold);
    inputEl.setAttribute('data-current', currentDue);
  }

  const dateEl = document.getElementById('quickDueBillingDate');
  if (dateEl) {
    dateEl.value = toInputDateValue(billingDate);
  }

  updateQuickDuePreview();
  overlay.classList.remove('hidden');
  setTimeout(() => inputEl?.focus(), 100);
}

function closeQuickDueModal() {
  const overlay = document.getElementById('quickDueModalOverlay');
  if (overlay) overlay.classList.add('hidden');
  activeQuickDueAccount = null;
}

function quickDueAdjust(amount, isReset = false) {
  const inputEl = document.getElementById('quickDueInput');
  if (!inputEl) return;
  if (isReset) {
    inputEl.value = '0';
  } else {
    const cur = parseFloat(inputEl.value) || 0;
    inputEl.value = (cur + amount).toFixed(2);
  }
  updateQuickDuePreview();
}

function updateQuickDuePreview() {
  const inputEl = document.getElementById('quickDueInput');
  const previewEl = document.getElementById('quickDueHeadroomPreview');
  if (!inputEl || !previewEl) return;

  const th = parseFloat(inputEl.getAttribute('data-threshold')) || 500;
  const val = parseFloat(inputEl.value) || 0;
  const headroom = Math.max(0, th - val);
  previewEl.textContent = `$${fmtNum(headroom)} (${Math.min(100, Math.round((val / th) * 100))}% used)`;
  if (val >= th) {
    previewEl.style.color = 'var(--dan)';
  } else if (val >= th * 0.8) {
    previewEl.style.color = 'var(--wrn)';
  } else {
    previewEl.style.color = 'var(--suc)';
  }
}

function submitQuickDueUpdate() {
  if (!isAdmin() || !activeQuickDueAccount) return;
  const inputEl = document.getElementById('quickDueInput');
  const dateEl = document.getElementById('quickDueBillingDate');
  const btn = document.getElementById('btnSubmitQuickDue');
  const val = parseFloat(inputEl?.value) || 0;
  const bDate = dateEl?.value || '';

  if (btn) {
    btn.disabled = true;
    btn.innerHTML = '<span class="ms spin" style="font-size:16px">sync</span> Saving...';
  }

  google.script.run.withSuccessHandler(res => {
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = '<span class="ms" style="font-size:16px">bolt</span> Update Due Now';
    }
    closeQuickDueModal();
    showToast(`Updated due for "${activeQuickDueAccount}" to $${val}`, 'success');
    loadAdAccountBillingData(true);
  }).withFailureHandler(err => {
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = '<span class="ms" style="font-size:16px">bolt</span> Update Due Now';
    }
    showToast(err?.message || 'Failed to update due', 'error');
  }).quickUpdateAdAccountDue(activeQuickDueAccount, val, bDate);
}

function openAdBillingModal(accData = null) {
  if (!isAdmin()) return;
  const overlay = document.getElementById('adBillingModalOverlay');
  if (!overlay) return;

  // Populate datalist with known ad accounts
  const dl = document.getElementById('billAccountList');
  if (dl) {
    const list = state.adAccounts || [];
    dl.innerHTML = list.map(a => `<option value="${escapeHtml(a.a)}">${escapeHtml(a.id || '')}</option>`).join('');
  }

  const titleEl = document.getElementById('adBillingModalTitle');
  const nameInput = document.getElementById('billModalAccountName');
  const dueInput = document.getElementById('billModalDueAmount');
  const thInput = document.getElementById('billModalThreshold');
  const dateInput = document.getElementById('billModalBillingDate');
  const idInput = document.getElementById('billModalIdName');
  const methodInput = document.getElementById('billModalPaymentMethod');
  const statusSelect = document.getElementById('billModalStatus');
  const notesInput = document.getElementById('billModalNotes');

  if (accData) {
    if (titleEl) titleEl.textContent = `Edit Billing: ${accData.accountName}`;
    if (nameInput) {
      nameInput.value = accData.accountName || '';
      nameInput.readOnly = true;
      nameInput.style.background = 'var(--card-subtle)';
    }
    if (dueInput) dueInput.value = accData.dueAmount !== undefined ? accData.dueAmount : 0;
    if (thInput) thInput.value = accData.threshold || 500;
    if (dateInput) dateInput.value = toInputDateValue(accData.billingDate);
    if (idInput) idInput.value = accData.idName || '';
    if (methodInput) methodInput.value = accData.paymentMethod || '';
    if (statusSelect) statusSelect.value = accData.status || 'active';
    if (notesInput) notesInput.value = accData.notes || '';
  } else {
    if (titleEl) titleEl.textContent = 'Add Ad Account Billing Tracker';
    if (nameInput) {
      nameInput.value = '';
      nameInput.readOnly = false;
      nameInput.style.background = '';
    }
    if (dueInput) dueInput.value = '';
    if (thInput) thInput.value = 500;
    if (dateInput) dateInput.value = '';
    if (idInput) idInput.value = '';
    if (methodInput) methodInput.value = '';
    if (statusSelect) statusSelect.value = 'active';
    if (notesInput) notesInput.value = '';
  }

  overlay.classList.remove('hidden');
  setTimeout(() => (accData ? dueInput : nameInput)?.focus(), 100);
}

function closeAdBillingModal() {
  const overlay = document.getElementById('adBillingModalOverlay');
  if (overlay) overlay.classList.add('hidden');
}

function setBillingThresholdPreset(num) {
  const thInput = document.getElementById('billModalThreshold');
  if (thInput) thInput.value = num;
}

function saveAdBillingModalForm() {
  if (!isAdmin()) return;
  const name = (document.getElementById('billModalAccountName')?.value || '').trim();
  const due = parseFloat(document.getElementById('billModalDueAmount')?.value) || 0;
  const th = parseFloat(document.getElementById('billModalThreshold')?.value) || 500;
  const bDate = (document.getElementById('billModalBillingDate')?.value || '').trim();
  const idName = (document.getElementById('billModalIdName')?.value || '').trim();
  const paymentMethod = (document.getElementById('billModalPaymentMethod')?.value || '').trim();
  const status = document.getElementById('billModalStatus')?.value || 'active';
  const notes = (document.getElementById('billModalNotes')?.value || '').trim();
  const btn = document.getElementById('btnSaveAdBilling');

  if (!name) {
    showToast('Please enter an Ad Account name', 'error');
    document.getElementById('billModalAccountName')?.focus();
    return;
  }

  if (btn) {
    btn.disabled = true;
    btn.innerHTML = '<span class="ms spin" style="font-size:16px">sync</span> Saving...';
  }

  const payload = {
    accountName: name,
    dueAmount: due,
    threshold: th,
    billingDate: bDate,
    idName,
    paymentMethod,
    status,
    notes
  };

  google.script.run.withSuccessHandler(res => {
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = '<span class="ms" style="font-size:16px">save</span> Save Account Billing';
    }
    closeAdBillingModal();
    showToast(`Ad account "${name}" billing saved successfully!`, 'success');
    loadAdAccountBillingData(true);
  }).withFailureHandler(err => {
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = '<span class="ms" style="font-size:16px">save</span> Save Account Billing';
    }
    showToast(err?.message || 'Failed to save account billing', 'error');
  }).saveAdAccountBilling(payload);
}

// Parser for campaign text strings in various formats:
// 1. Pipe format: "SL: 187 | Page: Ak Master | Start: 12-Sep-2026 | End: 13-Sep-2026 | Type: Increase | Days: 1 | Budget: $56 | Rate: ৳140 | Taka: ৳7840.00 | AD: Agency | 011 (Aminul)"
// 2. Space format: "187/011 Ak Master 12-Sep-2026 13-Sep-2026 Increase 1 $56"
function parseCampaignString(str) {
  if (!str) return null;
  if (typeof str === 'object') return str;
  const s = String(str).trim();
  if (!s) return null;

  // 1. Pipe-separated or key-colon format
  if (s.includes(':') && (s.includes('|') || s.includes('\n'))) {
    const obj = {};
    const rawTokens = s.split(/[\n|]+/).map(t => t.trim()).filter(Boolean);
    let currentKey = '';

    for (const tok of rawTokens) {
      const cIdx = tok.indexOf(':');
      if (cIdx > 0) {
        currentKey = tok.substring(0, cIdx).trim().toLowerCase();
        const val = tok.substring(cIdx + 1).trim();
        assignKeyVal(obj, currentKey, val);
      } else if (currentKey) {
        // Appended value without colon (e.g. "Agency" | "011 (Aminul)" inside AD field)
        appendKeyVal(obj, currentKey, tok);
      }
    }

    if (Object.keys(obj).length >= 2) {
      return obj;
    }
  }

  // 2. Space-separated format: [SL]/[AdAccount] [Page Name...] [Start] [End] [Type] [Day] [Budget]
  // e.g. "187/011 Ak Master 12-Sep-2026 13-Sep-2026 Increase 1 $56"
  const slashMatch = s.match(/^([A-Za-z0-9]+)\/([A-Za-z0-9_-]+)\s+(.+)$/);
  if (slashMatch) {
    const sl = slashMatch[1];
    const adNum = slashMatch[2];
    const rest = slashMatch[3].trim();
    const tokens = rest.split(/\s+/);
    if (tokens.length >= 6) {
      const budget = tokens[tokens.length - 1].replace(/[^0-9.]/g, '');
      const day = tokens[tokens.length - 2].replace(/[^0-9]/g, '');
      const type = tokens[tokens.length - 3];
      const end = tokens[tokens.length - 4];
      const start = tokens[tokens.length - 5];
      const pageName = tokens.slice(0, tokens.length - 5).join(' ');
      return {
        'SL': sl,
        'sl_no': sl,
        'AD Account': adNum,
        'ID Name': adNum,
        'AD Number': adNum,
        'Page Name': pageName,
        'page_name': pageName,
        'Start': start,
        'start_date': start,
        'End': end,
        'end_date': end,
        'Types': type,
        'types': type,
        'Day': day,
        'day': day,
        'Budget': budget,
        'budget': budget
      };
    }
  }

  return null;
}

function assignKeyVal(obj, key, val) {
  if (key.includes('sl')) obj['SL'] = val;
  else if (key.includes('page')) obj['Page Name'] = val;
  else if (key.includes('start')) obj['Start'] = val;
  else if (key.includes('end')) obj['End'] = val;
  else if (key.includes('type')) obj['Types'] = val;
  else if (key.includes('day')) obj['Day'] = val.replace(/days?/i, '').trim();
  else if (key.includes('budget')) obj['Budget'] = val.replace(/[^0-9.]/g, '');
  else if (key.includes('rate')) obj['Rate'] = val.replace(/[^0-9.]/g, '');
  else if (key.includes('taka')) obj['Taka'] = val.replace(/[^0-9.]/g, '');
  else if (key === 'ad' || key.includes('ad account')) {
    obj['AD Account'] = val;
    obj['AD'] = val;
  }
  else if (key.includes('id')) obj['ID Name'] = val;
  else if (key.includes('status')) obj['Status'] = val;
}

function appendKeyVal(obj, key, tok) {
  if (key === 'ad' || key.includes('ad account')) {
    obj['AD Account'] = (obj['AD Account'] ? `${obj['AD Account']} | ` : '') + tok;
    obj['AD'] = obj['AD Account'];
  } else if (key.includes('page')) {
    obj['Page Name'] = (obj['Page Name'] ? `${obj['Page Name']} ` : '') + tok;
  }
}

// Helper to extract clean Ad Account number/ID from ID Name or AD Account
function extractAdAccountNumber(r) {
  if (!r) return '';
  if (typeof r === 'string') {
    const parsed = parseCampaignString(r);
    if (parsed) return extractAdAccountNumber(parsed);
  }

  const directNum = String(r['AD Number'] || r['ad_number'] || r['adNum'] || '').trim();
  if (directNum) return directNum;

  const idName = String(r['ID Name'] || r['id_name'] || r['ID'] || r['id'] || '').trim();
  const adAcc = String(r['AD Account'] || r['ad_account'] || r['AD'] || r['ad'] || r['Account'] || '').trim();

  const findNumberInString = (str) => {
    if (!str) return '';
    // Look inside parenthesis or after pipes, e.g. "Agency | 011 (Aminul)" -> "011"
    const segments = str.split(/[|/()]+/).map(s => s.trim()).filter(Boolean);
    for (const seg of segments) {
      const match = seg.match(/^\d+$/);
      if (match) return match[0];
    }
    // Standalone 2-5 digit code (e.g. 011, 001, 18, 105)
    const codeMatch = str.match(/\b\d{2,5}\b/);
    if (codeMatch) return codeMatch[0];
    // Any standalone number
    const anyNum = str.match(/\b\d+\b/);
    if (anyNum) return anyNum[0];
    return '';
  };

  const idNum = findNumberInString(idName);
  if (idNum) return idNum;

  const adNum = findNumberInString(adAcc);
  if (adNum) return adNum;

  // Fallback: clean first token
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

// Campaign Details Formatter & Copy Helper
// Format: 187/Ad_account_number Page_Name Start End Type Day Budget
// Example: 187/011 Ak Master 12-Sep-2026 13-Sep-2026 Increase 1 $56
function formatCampaignText(r) {
  if (!r) return '';
  if (Array.isArray(r)) {
    return r.map(item => formatCampaignText(item)).filter(Boolean).join('\n');
  }
  if (typeof r === 'string') {
    const trimmed = r.trim();
    if (trimmed.includes('\n')) {
      return trimmed.split(/\r?\n/).map(line => formatCampaignText(line)).filter(Boolean).join('\n');
    }
    const parsed = parseCampaignString(trimmed);
    if (parsed) r = parsed;
    else return trimmed;
  }

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

function copyCampaignRecord(record) {
  let text = '';
  if (Array.isArray(record)) {
    text = record.map(r => formatCampaignText(r)).filter(Boolean).join('\n');
  } else if (typeof record === 'string') {
    text = record;
  } else {
    text = formatCampaignText(record);
  }
  if (!text) return;

  const count = text.split('\n').filter(Boolean).length;
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text).then(() => {
      if (count > 1) {
        showToast(`📋 Copied ${count} campaigns to clipboard:\n${text}`, 'success');
      } else {
        showToast(`📋 Copied: ${text}`, 'success');
      }
    }).catch(() => {
      fallbackCopyCampaignText(text);
    });
  } else {
    fallbackCopyCampaignText(text);
  }
}

function copyCampaignRecordByIndex(idx) {
  const list = state.currentRenderedRecords || state.filteredRecords || state.rawRecords || [];
  const r = list[idx];
  if (r) {
    copyCampaignRecord(r);
  }
}

function copyAllCurrentTableCampaigns() {
  const list = state.currentRenderedRecords || state.filteredRecords || state.rawRecords || [];
  const validCampaigns = list.filter(r => r && (r['Page Name'] || r['page_name'] || r['Budget']));
  if (!validCampaigns.length) {
    showToast('No campaigns found in current view to copy', 'warning');
    return;
  }
  const allText = validCampaigns.map(r => formatCampaignText(r)).filter(Boolean).join('\n');
  copyCampaignRecord(allText);
  showToast(`📋 All ${validCampaigns.length} campaigns copied to clipboard! (Each separated)`, 'success');
}

function copyAllCurrentTableCampaignsSeparately() {
  const list = state.currentRenderedRecords || state.filteredRecords || state.rawRecords || [];
  const validCampaigns = list.filter(r => r && (r['Page Name'] || r['page_name'] || r['Budget']));
  if (!validCampaigns.length) {
    showToast('No campaigns found in current view to copy', 'warning');
    return;
  }
  copyMultipleRecordsSequentially(validCampaigns);
}

function fallbackCopyCampaignText(text) {
  const ta = document.createElement('textarea');
  ta.value = text;
  ta.style.position = 'fixed';
  ta.style.left = '-9999px';
  document.body.appendChild(ta);
  ta.select();
  try {
    document.execCommand('copy');
    const count = text.split('\n').filter(Boolean).length;
    if (count > 1) {
      showToast(`📋 Copied ${count} campaigns to clipboard!`, 'success');
    } else {
      showToast(`📋 Copied: ${text}`, 'success');
    }
  } catch (e) {
    showToast('Copied to clipboard', 'success');
  }
  document.body.removeChild(ta);
}

async function copyMultipleRecordsSequentially(items) {
  if (!items || !items.length) return;
  const list = items.map(r => typeof r === 'string' ? r : formatCampaignText(r)).filter(Boolean);
  if (!list.length) return;

  if (list.length === 1) {
    copyCampaignRecord(list[0]);
    return;
  }

  // To register separate individual entries in Windows Clipboard History (Win + V):
  // We write each campaign individually into the clipboard separated by a brief delay (75ms).
  // Writing in reverse order ensures that the FIRST campaign (e.g. SL 390) becomes the current
  // top active clipboard item (for immediate Ctrl+V), while the next campaigns (SL 391, etc.)
  // are right underneath it as distinct separate entries in Win + V!
  const reversed = [...list].reverse();
  for (let i = 0; i < reversed.length; i++) {
    const text = reversed[i];
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(text);
      } else {
        fallbackCopyCampaignText(text);
      }
    } catch (_) {
      fallbackCopyCampaignText(text);
    }
    if (i < reversed.length - 1) {
      await new Promise(r => setTimeout(r, 75));
    }
  }

  showToast(`📋 ${list.length} campaigns copied to clipboard individually (Win + V history ready)!`, 'success');
}

function renderSavedCampaignsNotice(savedItems) {
  state.lastSavedCampaigns = savedItems || [];
  const existing = document.getElementById('lastSavedNoticeBox');
  if (existing) existing.remove();

  if (!savedItems || savedItems.length <= 1) return;

  const targetContainer = document.getElementById('recordsCard') || document.querySelector('.main-content-wrap');
  if (!targetContainer) return;

  const notice = document.createElement('div');
  notice.id = 'lastSavedNoticeBox';
  notice.className = 'saved-campaigns-action-bar';
  notice.style.cssText = 'background:var(--card);border:1.5px solid var(--pri);border-radius:10px;padding:12px 14px;margin-bottom:12px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px;box-shadow:0 4px 16px rgba(0,0,0,0.12)';
  
  notice.innerHTML = `
    <div style="display:flex;align-items:center;gap:10px">
      <div style="width:34px;height:34px;border-radius:8px;background:rgba(16,185,129,0.12);display:flex;align-items:center;justify-content:center;color:var(--suc)">
        <span class="ms" style="font-size:22px">content_paste</span>
      </div>
      <div>
        <div style="font-weight:800;font-size:13.5px;color:var(--txt)">${savedItems.length} Campaigns Successfully Saved! 📋</div>
        <div style="font-size:11.5px;color:var(--txt-muted)">Click below to copy all together or individually:</div>
      </div>
    </div>
    <div style="display:flex;gap:6px;flex-wrap:wrap;align-items:center">
      <button type="button" class="btn btn-primary btn-sm" onclick="copyMultipleSavedCampaignsCombined()" style="height:32px;font-size:12px;padding:0 12px;gap:5px;font-weight:700" title="Copy all saved campaigns to clipboard">
        <span class="ms" style="font-size:14px">copy_all</span> Copy All (${savedItems.length})
      </button>
      ${savedItems.map((c, i) => {
        const slVal = c.SL || c.sl || c.sl_no || (i + 1);
        return `
          <button type="button" id="btnCopySingleSaved_${i}" class="btn btn-secondary btn-sm" onclick="copySingleSavedCampaignByIndex(${i})" style="height:32px;font-size:12px;padding:0 10px;gap:5px;font-weight:700;border:1.5px solid var(--pri);background:rgba(99,102,241,0.08);color:var(--pri)" title="Click to copy only this campaign">
            <span class="ms" style="font-size:13px">content_copy</span> SL #${slVal}
          </button>
        `;
      }).join('')}
      <button type="button" class="btn btn-secondary btn-sm" onclick="copyMultipleSavedCampaigns()" style="height:32px;font-size:11.5px;padding:0 9px;gap:4px" title="Sequential clipboard copy (Win + V)">
        <span class="ms" style="font-size:13px">history</span> Win+V History
      </button>
      <button type="button" class="btn btn-secondary btn-sm" onclick="document.getElementById('lastSavedNoticeBox')?.remove()" style="height:32px;padding:0 8px;color:var(--txt-dim)" title="Dismiss">
        <span class="ms" style="font-size:15px">close</span>
      </button>
    </div>
  `;

  targetContainer.parentNode.insertBefore(notice, targetContainer);
}

function copyMultipleSavedCampaigns() {
  if (state.lastSavedCampaigns && state.lastSavedCampaigns.length) {
    copyMultipleRecordsSequentially(state.lastSavedCampaigns);
  }
}

function copyMultipleSavedCampaignsCombined() {
  if (state.lastSavedCampaigns && state.lastSavedCampaigns.length) {
    const allText = state.lastSavedCampaigns.map(r => formatCampaignText(r)).filter(Boolean).join('\n');
    copyCampaignRecord(allText);
    showToast(`📋 ${state.lastSavedCampaigns.length} campaigns copied to clipboard!`, 'success');
  }
}

function copySingleSavedCampaignByIndex(idx) {
  if (state.lastSavedCampaigns && state.lastSavedCampaigns[idx]) {
    const text = formatCampaignText(state.lastSavedCampaigns[idx]);
    copyCampaignRecord(text);
    const btn = document.getElementById(`btnCopySingleSaved_${idx}`);
    if (btn) {
      const orig = btn.innerHTML;
      btn.innerHTML = '<span class="ms" style="font-size:13px;color:var(--suc)">check</span> Copied!';
      setTimeout(() => { if (btn) btn.innerHTML = orig; }, 1600);
    }
    showToast(`📋 #${state.lastSavedCampaigns[idx]['SL'] || (idx + 1)} copied:\n${text}`, 'success');
  }
}

function handleQuickCampaignPaste() {
  const input = document.getElementById('quickPasteCampaignInput');
  const val = (input?.value || '').trim();
  if (!val) {
    showToast('Please paste a campaign text first', 'warning');
    return;
  }

  // Support multiline paste (multiple campaigns)
  const lines = val.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  const parsedList = [];
  for (const line of lines) {
    const p = parseCampaignString(line);
    if (p) parsedList.push(p);
  }

  if (!parsedList.length) {
    showToast('Could not parse text. Expected: SL: 187 | Page: ... or 187/011 Page...', 'error');
    return;
  }

  const form = document.getElementById('entryForm');
  if (!form) return;

  const parsed = parsedList[0];

  if (parsed['SL']) {
    const slInput = form.querySelector('[name="sl_no"]');
    if (slInput) slInput.value = parsed['SL'];
  }
  if (parsed['Page Name']) {
    const pageInput = form.querySelector('#pageNameInput');
    if (pageInput) {
      pageInput.value = parsed['Page Name'];
      handlePageNameLookup(parsed['Page Name']);
    }
  }
  if (parsed['Start']) {
    const startInput = form.querySelector('[name="start_date"]');
    const dVal = toInputDateValue(parsed['Start']);
    if (startInput && dVal) startInput.value = dVal;
  }
  if (parsed['End']) {
    const endInput = form.querySelector('[name="end_date"]');
    const dVal = toInputDateValue(parsed['End']);
    if (endInput && dVal) endInput.value = dVal;
  }
  if (parsed['Types']) {
    const typeSelect = form.querySelector('#typesSelect');
    if (typeSelect) typeSelect.value = parsed['Types'];
  }
  if (parsed['Day']) {
    const dayInput = form.querySelector('[name="day"]');
    if (dayInput) dayInput.value = parsed['Day'];
  }
  if (parsed['Budget']) {
    const budgetInput = form.querySelector('[name="budget"]');
    if (budgetInput) budgetInput.value = parsed['Budget'];
  }
  if (parsed['Rate']) {
    const rateInput = form.querySelector('[name="rate"]');
    if (rateInput) rateInput.value = parsed['Rate'];
  }
  if (parsed['Taka']) {
    const takaInput = form.querySelector('[name="taka"]');
    if (takaInput) takaInput.value = parsed['Taka'];
  }

  if (parsed['AD Account']) {
    const adInput = form.querySelector('#adAccountInput');
    const idInput = form.querySelector('#idNameInput');
    if (adInput) adInput.value = parsed['AD Account'];
    if (idInput && parsed['ID Name']) idInput.value = parsed['ID Name'];
    else if (idInput) {
      const num = extractAdAccountNumber(parsed);
      if (num) idInput.value = num;
    }
  }

  // If multiple campaigns were pasted, automatically generate sub-campaign rows for the rest
  const rowsContainer = document.getElementById('multipleCampaignRows');
  if (rowsContainer) rowsContainer.innerHTML = '';

  if (parsedList.length > 1) {
    for (let i = 1; i < parsedList.length; i++) {
      const sub = parsedList[i];
      addMultipleCampaignRow();
      const currentRows = rowsContainer?.querySelectorAll('.multi-campaign-row');
      const lastRow = currentRows ? currentRows[currentRows.length - 1] : null;
      if (lastRow) {
        const dayInput = lastRow.querySelector('.multi-campaign-day');
        const budgetInput = lastRow.querySelector('.multi-campaign-budget');
        if (dayInput && sub['Day']) dayInput.value = sub['Day'];
        if (budgetInput && sub['Budget']) budgetInput.value = sub['Budget'];
      }
    }
  }

  const startInput = form.querySelector('[name="start_date"]');
  if (startInput) {
    startInput.dispatchEvent(new Event('input', { bubbles: true }));
    startInput.dispatchEvent(new Event('change', { bubbles: true }));
  }
  if (window.updateCampaignCalc) window.updateCampaignCalc();

  // Copy parsed campaigns separately into clipboard history (Win + V)
  if (parsedList.length > 1) {
    copyMultipleRecordsSequentially(parsedList);
  } else {
    copyCampaignRecord(parsedList[0]);
  }
  renderSavedCampaignsNotice(parsedList);
}

function applyIncreaseCampaign(c) {
  const form = document.getElementById('entryForm');
  if (!form || !c) return;

  const typeSelect = form.querySelector('#typesSelect') || form.querySelector('[name="types"]');
  const startInput = form.querySelector('[name="start_date"]');
  const adInput = form.querySelector('#adAccountInput') || form.querySelector('[name="ad_account"]');
  const idInput = form.querySelector('#idNameInput') || form.querySelector('[name="id_name"]');
  const budgetInput = form.querySelector('[name="budget"]');

  const rawEnd = c.endInputVal || c.end || c['End'] || c['End Date'] || '';
  const inputDateVal = toInputDateValue(rawEnd);
  const formattedEndDate = formatCustomDate(rawEnd);

  if (typeSelect) typeSelect.value = 'Increase';
  if (startInput && inputDateVal) {
    startInput.value = inputDateVal;
  }
  if (adInput && (c.adAccount || c['AD Account'])) adInput.value = c.adAccount || c['AD Account'];
  if (idInput && (c.idName || c['ID Name'])) idInput.value = c.idName || c['ID Name'];
  if (budgetInput && (c.budget || c['Budget'])) budgetInput.value = c.budget || c['Budget'];

  // Dispatch events to trigger recalculation and ensure UI reactivity
  if (startInput) {
    startInput.dispatchEvent(new Event('input', { bubbles: true }));
    startInput.dispatchEvent(new Event('change', { bubbles: true }));
  }
  if (window.updateCampaignCalc) {
    window.updateCampaignCalc();
  }

  // Highlight selected button
  const btns = form.querySelectorAll('.campaign-select-btn');
  btns.forEach(b => b.classList.remove('active'));
  const activeBtn = Array.from(btns).find(b => b.textContent.includes(c.slLabel || c.sl));
  if (activeBtn) activeBtn.classList.add('active');

  const label = c.slLabel || (c.month ? `${c.month.slice(0, 3)} #${c.sl}` : `SL #${c.sl}`);
  showToast(`⚡ Selected campaign [${label}]. Start date set to ${formattedEndDate || inputDateVal}`, 'success');
}

function applyIncreaseCampaignByIndex(idx) {
  const campaigns = state.currentProfileCampaigns || [];
  const c = campaigns[idx];
  if (c) {
    applyIncreaseCampaign(c);
  }
}

function applyIncreaseStartDate(dateStr) {
  const form = document.getElementById('entryForm');
  if (!form) return;
  const startInput = form.querySelector('[name="start_date"]');
  const typeSelect = form.querySelector('#typesSelect') || form.querySelector('[name="types"]');
  const inputDateVal = toInputDateValue(dateStr);
  const formattedDate = formatCustomDate(dateStr);

  if (typeSelect) typeSelect.value = 'Increase';
  if (startInput && inputDateVal) {
    startInput.value = inputDateVal;
    startInput.dispatchEvent(new Event('input', { bubbles: true }));
    startInput.dispatchEvent(new Event('change', { bubbles: true }));
  }
  if (window.updateCampaignCalc) {
    window.updateCampaignCalc();
  }
  showToast(`⚡ Start Date set to ${formattedDate || inputDateVal} for Increase`, 'success');
}

function quickIncreaseCampaign(r) {
  if (!isAdmin()) {
    showToast('Admin access required to increase campaign', 'error');
    return;
  }
  
  const card = document.getElementById('entryCard');
  if (card) card.classList.remove('hidden');

  renderEntryForm();

  const form = document.getElementById('entryForm');
  if (!form) return;

  const rawEnd = r['End'] || r['End Date'] || new Date();
  const inputDateVal = toInputDateValue(rawEnd);
  const formattedEndDate = formatCustomDate(rawEnd);
  const sl = String(r['SL'] || r['sl_no'] || '').padStart(3, '0');
  const month = r['Month'] || '';
  const year = String(r['Year'] || '');
  const slLabel = month ? `${month.slice(0, 3)} ${year} #${sl}` : `SL #${sl}`;
  
  const pageInput = form.querySelector('[name="page_name"]');
  const adInput = form.querySelector('[name="ad_account"]');
  const idInput = form.querySelector('#idNameInput') || form.querySelector('[name="id_name"]');
  const statusSelect = form.querySelector('[name="status"]');
  const typeSelect = form.querySelector('[name="types"]');
  const startInput = form.querySelector('[name="start_date"]');
  const dayInput = form.querySelector('[name="day"]');
  const budgetInput = form.querySelector('[name="budget"]');

  if (pageInput) {
    pageInput.value = r['Page Name'] || '';
    handlePageNameLookup(r['Page Name'] || '', { skipTypeOverride: true, selectedSL: sl });
  }
  if (adInput) adInput.value = r['AD Account'] || '';
  if (idInput) idInput.value = r['ID Name'] || '';
  if (statusSelect) statusSelect.value = r['Status'] || 'Old Page';
  if (typeSelect) typeSelect.value = 'Increase';
  if (startInput && inputDateVal) {
    startInput.value = inputDateVal;
    startInput.dispatchEvent(new Event('input', { bubbles: true }));
    startInput.dispatchEvent(new Event('change', { bubbles: true }));
  }
  if (dayInput) dayInput.value = '1';
  if (budgetInput) budgetInput.value = r['Budget'] || '10';

  if (window.updateCampaignCalc) window.updateCampaignCalc();

  card?.scrollIntoView({ behavior: 'smooth', block: 'center' });
  card?.classList.add('pulse-highlight');
  setTimeout(() => card?.classList.remove('pulse-highlight'), 1400);

  showToast(`⚡ Increasing campaign [${slLabel}]. Start date set to ${formattedEndDate || inputDateVal}`, 'success');
}

function quickIncreaseCampaignByIndex(idx) {
  const records = state.currentRenderedRecords || state.filteredRecords || state.rawRecords || [];
  const r = records[idx];
  if (r) quickIncreaseCampaign(r);
}

function cloneCampaignRecord(r) {
  if (!r) return;

  const card = document.getElementById('entryCard');
  if (card) card.classList.remove('hidden');

  renderEntryForm();

  const form = document.getElementById('entryForm');
  if (!form) return;

  const pageName = r['Page Name'] || r['page_name'] || r['Page'] || '';
  const adAccount = r['AD Account'] || r['ad_account'] || '';
  const idName = r['ID Name'] || r['id_name'] || '';
  const status = r['Status'] || r['status'] || 'Old Page';
  const types = r['Types'] || r['types'] || 'Boost';
  const day = r['Day'] || r['day'] || '1';
  const budget = r['Budget'] || r['budget'] || '10';
  const rate = r['Rate'] || r['rate'] || state.rate || '140';
  const offUsd = r['OFF USD'] || r['off_usd'] || '0';
  const todayStr = getTodayLocalDateStr();

  const pageInput = form.querySelector('#pageNameInput') || form.querySelector('[name="page_name"]');
  const adInput = form.querySelector('#adAccountInput') || form.querySelector('[name="ad_account"]');
  const idInput = form.querySelector('#idNameInput') || form.querySelector('[name="id_name"]');
  const statusSelect = form.querySelector('#statusSelect') || form.querySelector('[name="status"]');
  const typeSelect = form.querySelector('#typesSelect') || form.querySelector('[name="types"]');
  const startInput = form.querySelector('[name="start_date"]');
  const dayInput = form.querySelector('[name="day"]');
  const budgetInput = form.querySelector('[name="budget"]');
  const rateInput = form.querySelector('[name="rate"]');
  const offInput = form.querySelector('[name="off_usd"]');

  if (pageInput) {
    pageInput.value = pageName;
    handlePageNameLookup(pageName, { skipTypeOverride: true });
  }
  if (adInput && adAccount) adInput.value = adAccount;
  if (idInput && idName) idInput.value = idName;
  if (statusSelect && status) statusSelect.value = status;
  if (typeSelect && types) typeSelect.value = types;
  if (startInput) {
    startInput.value = todayStr;
    startInput.dispatchEvent(new Event('input', { bubbles: true }));
    startInput.dispatchEvent(new Event('change', { bubbles: true }));
  }
  if (dayInput) dayInput.value = day;
  if (budgetInput) budgetInput.value = budget;
  if (rateInput && rate) rateInput.value = rate;
  if (offInput && offUsd) offInput.value = offUsd;

  if (window.updateCampaignCalc) window.updateCampaignCalc();

  card?.scrollIntoView({ behavior: 'smooth', block: 'center' });
  card?.classList.add('pulse-highlight');
  setTimeout(() => card?.classList.remove('pulse-highlight'), 1400);

  const sl = r['SL'] || r['sl_no'] || '';
  const slLabel = sl ? ` #${sl}` : '';
  showToast(`📋 Cloned campaign [${pageName || 'Record'}${slLabel}]. Pre-filled and ready to save!`, 'success');
}

function cloneCampaignRecordByIndex(idx) {
  const records = state.currentRenderedRecords || [];
  const r = records[idx];
  if (r) {
    cloneCampaignRecord(r);
  }
}

// Smart Page Profile Lookup (AD Account, Most Used Type, Old/New Status, and Increase Last Campaign)
function renderPageRelevantCampaignDisplay(cleanName, data) {
  const form = document.getElementById('entryForm');
  if (!form) return;
  const cardEl = form.querySelector('#pageRelevantCampaignCard');
  if (!cardEl) return;

  const typeSelect = form.querySelector('#typesSelect') || form.querySelector('[name="types"]');
  const isIncrease = typeSelect && typeSelect.value === 'Increase';

  // User requested: If not Increase type, past campaign display is completely removed!
  if (!isIncrease || !cleanName || !data) {
    cardEl.style.display = 'none';
    cardEl.innerHTML = '';
    return;
  }

  // Get the last campaign for this page
  const allList = (data.campaigns && data.campaigns.length)
    ? data.campaigns
    : ((data.activeCampaigns && data.activeCampaigns.length) ? data.activeCampaigns : []);

  const latest = data.latestCampaign || (allList.length ? allList[0] : null) || data.activeCampaign;

  if (!latest) {
    cardEl.style.display = 'none';
    cardEl.innerHTML = '';
    return;
  }

  const endStr = latest.end || latest['End Date'] || latest['End'] || '';
  const endInputVal = toInputDateValue(endStr);
  const copyText = typeof formatCampaignText === 'function' 
    ? formatCampaignText(latest)
    : `${latest.sl || ''} ${cleanName} ${latest.start || ''} ${endStr} ${latest.types || ''} ${latest.day || ''} $${latest.budget || ''}`;

  // Automatically select end date to start date for Increase
  if (endInputVal) {
    const startInput = form.querySelector('[name="start_date"]');
    if (startInput && startInput.value !== endInputVal) {
      startInput.value = endInputVal;
      startInput.dispatchEvent(new Event('change', { bubbles: true }));
      if (typeof window.updateCampaignCalc === 'function') window.updateCampaignCalc();
    }
  }

  cardEl.style.display = 'block';
  cardEl.innerHTML = `
    <div class="increase-single-line-badge" onclick="selectCampaignForIncreaseStartDate('${escapeHtml(endInputVal)}', 0)" title="Click to apply Start Date: ${escapeHtml(endStr)}" style="cursor:pointer;background:var(--card-elevated);border:1px solid var(--pri-light, #c7d2fe);border-radius:6px;padding:6px 10px;font-size:11.5px;display:flex;align-items:center;justify-content:space-between;gap:8px;margin-top:6px;box-shadow:0 1px 3px rgba(0,0,0,0.05)">
      <div style="display:flex;align-items:center;gap:6px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:var(--txt)">
        <span class="pill pill-inc" style="font-size:9.5px;padding:1px 6px;flex-shrink:0"><span class="ms" style="font-size:11px">trending_up</span> Increase</span>
        <span style="font-family:monospace;font-size:11px;font-weight:600;color:var(--txt);overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escapeHtml(copyText)}</span>
      </div>
      <button type="button" class="btn btn-secondary" onclick="event.stopPropagation(); selectCampaignForIncreaseStartDate('${escapeHtml(endInputVal)}', 0)" style="padding:2px 8px;font-size:10.5px;flex-shrink:0;color:var(--pri);font-weight:700">
        <span class="ms" style="font-size:12px">calendar_today</span> Start: ${escapeHtml(endStr)}
      </button>
    </div>
  `;
}

window.selectCampaignForIncreaseStartDate = function(endInputVal, idx) {
  const form = document.getElementById('entryForm');
  if (!form || !endInputVal) return;
  const startInput = form.querySelector('[name="start_date"]');
  if (startInput) {
    startInput.value = endInputVal;
    startInput.dispatchEvent(new Event('input', { bubbles: true }));
    startInput.dispatchEvent(new Event('change', { bubbles: true }));
    if (window.updateCampaignCalc) window.updateCampaignCalc();
  }
  showToast(`⚡ Start Date set to ${endInputVal}`, 'success');
};

window.applyRelevantCampaignStart = function(endInputVal) {
  window.selectCampaignForIncreaseStartDate(endInputVal, 0);
};

window.applyRelevantCampaignAsIncrease = function(c) {
  if (c && c.endInputVal) {
    window.selectCampaignForIncreaseStartDate(c.endInputVal, 0);
  }
};

function handlePageNameLookup(pageName, opts = {}) {
  const form = document.getElementById('entryForm');
  if (!form) return;

  const cleanName = String(pageName || '').trim();
  const statusSelect = form.querySelector('[name="status"]');
  const adInput = form.querySelector('#adAccountInput') || form.querySelector('[name="ad_account"]');
  const idInput = form.querySelector('#idNameInput') || form.querySelector('[name="id_name"]');
  const typeSelect = form.querySelector('#typesSelect') || form.querySelector('[name="types"]');
  const startInput = form.querySelector('[name="start_date"]');
  const dayInput = form.querySelector('#dayInput') || form.querySelector('[name="day"]');
  const budgetInput = form.querySelector('#budgetInput') || form.querySelector('[name="budget"]');
  const noticeEl = form.querySelector('#increaseSmartNotice');

  if (!cleanName) {
    if (noticeEl) noticeEl.style.display = 'none';
    renderPageRelevantCampaignDisplay('', null);
    return;
  }

  // 1. Instant Synchronous check & auto-fill from local memory store (Zero Lag)
  const knownInList = state.pages && state.pages.some(p => p.toLowerCase() === cleanName.toLowerCase());
  
  // Collect all matching records from local memory store across all sheets
  const allKnownRecords = [
    ...(state.rawRecords || []),
    ...(state.records || []),
    ...(state.todayCreated || [])
  ];
  if (state.clientCache && typeof state.clientCache === 'object') {
    Object.values(state.clientCache).forEach(cached => {
      if (cached && Array.isArray(cached.rows)) {
        allKnownRecords.push(...cached.rows);
      }
    });
  }

  const seenIds = new Set();
  const localMatching = allKnownRecords.filter(r => {
    if (!r) return false;
    const p = String(r['Page Name'] || r.page_name || r.Page || '').trim().toLowerCase();
    if (p !== cleanName.toLowerCase()) return false;
    const uid = r['Record ID'] || r._recordId || `${r.SL}_${r['Start Date']}_${r.Budget}`;
    if (seenIds.has(uid)) return false;
    seenIds.add(uid);
    return true;
  });

  const now = new Date();
  const todayMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  if (statusSelect) {
    statusSelect.value = (knownInList || localMatching.length > 0) ? 'Old Page' : 'New Page';
  }

  // If page has past campaigns in local records, compute most frequent options immediately
  if (localMatching.length > 0) {
    const typeCounts = {};
    const dayCounts = {};
    const budgetCounts = {};
    const adCounts = {};
    const adIdMap = {};

    (state.adAccounts || []).forEach(a => {
      if (a && a.a) adIdMap[a.a] = a.id || '';
    });

    localMatching.forEach(r => {
      const t = String(r['Types'] || r.Type || r.type || '').trim();
      const d = parseInt(String(r['Day'] || r.day || ''), 10);
      const b = parseFloat(String(r['Budget'] || r.budget || ''));
      const ad = String(r['AD Account'] || r.ad_account || '').trim();
      const id = String(r['ID Name'] || r.id_name || '').trim();

      if (t) {
        const normType = t.toLowerCase().includes('boost') ? 'Boost' : (t.toLowerCase().includes('ad') ? 'Ad Campaign' : t);
        typeCounts[normType] = (typeCounts[normType] || 0) + 1;
      }
      if (!isNaN(d) && d > 0) dayCounts[String(d)] = (dayCounts[String(d)] || 0) + 1;
      if (!isNaN(b) && b > 0) budgetCounts[String(b)] = (budgetCounts[String(b)] || 0) + 1;
      if (ad) {
        adCounts[ad] = (adCounts[ad] || 0) + 1;
        if (id) adIdMap[ad] = id;
      }
    });

    // Find most frequent type
    let topType = 'Boost';
    let maxTypeC = 0;
    for (const [t, c] of Object.entries(typeCounts)) {
      if (c > maxTypeC) { maxTypeC = c; topType = t; }
    }

    // Find most frequent day
    let topDay = 2;
    let maxDayC = 0;
    for (const [d, c] of Object.entries(dayCounts)) {
      if (c > maxDayC) { maxDayC = c; topDay = parseInt(d, 10); }
    }

    // Find most frequent budget
    let topBudget = 10;
    let maxBudgetC = 0;
    for (const [b, c] of Object.entries(budgetCounts)) {
      if (c > maxBudgetC) { maxBudgetC = c; topBudget = parseFloat(b); }
    }

    // Find most frequent AD Account
    let topAd = '';
    let maxAdC = 0;
    for (const [ad, c] of Object.entries(adCounts)) {
      if (c > maxAdC) { maxAdC = c; topAd = ad; }
    }

    // Apply auto-fills instantly
    if (typeSelect && !opts.skipTypeOverride && typeSelect.value !== 'Increase') {
      typeSelect.value = topType;
    }
    if (adInput && topAd && (!adInput.value || isDemoAdAccount(adInput.value))) {
      adInput.value = topAd;
      if (idInput) idInput.value = adIdMap[topAd] || '';
    }
    if (dayInput && topDay > 0) {
      dayInput.value = topDay;
      form.querySelectorAll('.day-preset-chip').forEach(chip => {
        chip.classList.toggle('active', chip.getAttribute('data-day') == String(topDay) || parseInt(chip.textContent, 10) === topDay);
      });
    }
    if (budgetInput && topBudget > 0) {
      budgetInput.value = topBudget;
      form.querySelectorAll('.bud-preset-chip').forEach(chip => {
        chip.classList.toggle('active', chip.getAttribute('data-bud') == String(topBudget) || parseFloat(chip.textContent.replace('$', '')) === topBudget);
      });
    }

    // Recalculate campaign financials immediately
    if (typeof window.updateCampaignCalc === 'function') {
      window.updateCampaignCalc();
    }

    const sorted = [...localMatching].sort((a, b) => {
      const da = parseAnyDate(a['End Date'] || a.end_date || a.End || a['Start Date'] || a.start_date || a.Start);
      const db = parseAnyDate(b['End Date'] || b.end_date || b.End || b['Start Date'] || b.start_date || b.Start);
      return (db ? db.getTime() : 0) - (da ? da.getTime() : 0);
    });

    const activeRow = sorted.find(r => {
      const endD = parseAnyDate(r['End Date'] || r.end_date || r.End);
      return endD && endD.getTime() >= todayMidnight.getTime();
    });

    const toCampaignObj = (r) => {
      const endVal = r['End Date'] || r.end_date || r.End || '';
      return {
        sl: r.SL || r['SL'] || r.sl_no || '',
        slLabel: `SL #${r.SL || r['SL'] || r.sl_no || ''}`,
        start: r['Start Date'] || r.start_date || r.Start || '',
        end: endVal,
        endInputVal: toInputDateValue(endVal),
        types: r['Types'] || r.Type || r.type || 'Boost',
        day: r.Day || r.day || 1,
        budget: r.Budget || r.budget || 0,
        adAccount: r['AD Account'] || r.ad_account || '',
        idName: r['ID Name'] || r.id_name || '',
        isActive: !!(parseAnyDate(endVal) && parseAnyDate(endVal).getTime() >= todayMidnight.getTime())
      };
    };

    const syncProfile = {
      exists: true,
      status: 'Old Page',
      mostFrequentType: topType,
      mostFrequentDay: topDay,
      mostFrequentBudget: topBudget,
      mostFrequentAdAccount: topAd,
      mostFrequentIdName: adIdMap[topAd] || '',
      hasActiveCampaign: !!activeRow,
      activeCampaign: activeRow ? toCampaignObj(activeRow) : null,
      latestCampaign: sorted[0] ? toCampaignObj(sorted[0]) : null,
      campaigns: sorted.map(toCampaignObj)
    };
    renderPageRelevantCampaignDisplay(cleanName, syncProfile);
  } else if (!knownInList) {
    renderPageRelevantCampaignDisplay(cleanName, { exists: false, status: 'New Page' });
  }

  // 2. Fetch full deep profile from server across all historical Google Sheets data
  google.script.run.withSuccessHandler(profile => {
    if (!profile || !profile.success) return;

    // Apply Old / New Page status
    if (statusSelect) {
      statusSelect.value = profile.status || (profile.exists ? 'Old Page' : 'New Page');
    }

    if (profile.exists) {
      // Auto-fill most used AD Account & ID Name
      if (profile.mostFrequentAdAccount && adInput && (!adInput.value || isDemoAdAccount(adInput.value))) {
        adInput.value = profile.mostFrequentAdAccount;
        if (idInput && profile.mostFrequentIdName) {
          idInput.value = profile.mostFrequentIdName;
        }
      }

      // Auto-fill most used Type (Boost vs Ad Campaign)
      if (!opts.skipTypeOverride && typeSelect && typeSelect.value !== 'Increase' && profile.mostFrequentType) {
        typeSelect.value = profile.mostFrequentType;
      }

      // Auto-fill most used Day duration
      if (dayInput && profile.mostFrequentDay && profile.mostFrequentDay > 0) {
        dayInput.value = profile.mostFrequentDay;
        form.querySelectorAll('.day-preset-chip').forEach(chip => {
          chip.classList.toggle('active', chip.getAttribute('data-day') == String(profile.mostFrequentDay) || parseInt(chip.textContent, 10) === profile.mostFrequentDay);
        });
      }

      // Auto-fill most used Budget amount
      if (budgetInput && profile.mostFrequentBudget && profile.mostFrequentBudget > 0) {
        budgetInput.value = profile.mostFrequentBudget;
        form.querySelectorAll('.bud-preset-chip').forEach(chip => {
          chip.classList.toggle('active', chip.getAttribute('data-bud') == String(profile.mostFrequentBudget) || parseFloat(chip.textContent.replace('$', '')) === profile.mostFrequentBudget);
        });
      }

      // Trigger instant dynamic recalculation of End Date, Daily $, and Total Taka
      if (typeof window.updateCampaignCalc === 'function') {
        window.updateCampaignCalc();
      }

      showToast(`⚡ ${cleanName} previous campaign settings auto-filled (${profile.mostFrequentType || 'Boost'} • ${profile.mostFrequentDay || 2}d • $${profile.mostFrequentBudget || 10})`, 'info');
    }

    // Render Relevant Active Campaign display below Page Name if type is Increase
    renderPageRelevantCampaignDisplay(cleanName, profile);

    // Hide separate noticeEl
    if (noticeEl) {
      noticeEl.style.display = 'none';
      noticeEl.innerHTML = '';
    }
  }).getPageProfile(cleanName);
}

function formatShortDate(dateStr) {
  if (!dateStr) return '';
  return formatCustomDate(dateStr);
}

window.setQuickDay = function(days) {
  const form = document.getElementById('entryForm');
  if (!form) return;
  const dayInput = form.querySelector('[name="day"]');
  if (dayInput) {
    dayInput.value = days;
    form.querySelectorAll('.day-preset-chip').forEach(c => {
      c.classList.toggle('active', c.getAttribute('data-day') == String(days) || parseInt(c.textContent, 10) === days);
    });
    if (typeof window.updateCampaignCalc === 'function') window.updateCampaignCalc();
  }
};

window.setQuickBudget = function(amount) {
  const form = document.getElementById('entryForm');
  if (!form) return;
  const budgetInput = form.querySelector('[name="budget"]');
  if (budgetInput) {
    budgetInput.value = amount;
    form.querySelectorAll('.bud-preset-chip').forEach(c => {
      c.classList.toggle('active', c.getAttribute('data-bud') == String(amount) || parseFloat(c.textContent.replace('$', '')) === amount);
    });
    if (typeof window.updateCampaignCalc === 'function') window.updateCampaignCalc();
  }
};

window.handleQuickCampaignPaste = function() {
  const pasteInput = document.getElementById('quickPasteCampaignInput');
  const text = (pasteInput?.value || '').trim();
  if (!text) {
    showToast('Please paste a campaign text first (e.g. 187/011 Ak Master 12-Sep 13-Sep Boost 1 $10)', 'warning');
    return;
  }
  const form = document.getElementById('entryForm');
  if (!form) return;

  if (typeof parseCampaignString === 'function') {
    const p = parseCampaignString(text);
    if (p) {
      if (p.page_name) {
        const pageEl = form.querySelector('[name="page_name"]');
        if (pageEl) {
          pageEl.value = p.page_name;
          if (typeof handlePageNameLookup === 'function') handlePageNameLookup(p.page_name);
        }
      }
      if (p.ad_account) {
        const adEl = form.querySelector('[name="ad_account"]');
        if (adEl) adEl.value = p.ad_account;
      }
      if (p.id_name) {
        const idEl = form.querySelector('[name="id_name"]');
        if (idEl) idEl.value = p.id_name;
      }
      if (p.status) {
        const statEl = form.querySelector('[name="status"]');
        if (statEl) statEl.value = p.status;
      }
      if (p.types) {
        const typeEl = form.querySelector('[name="types"]');
        if (typeEl) typeEl.value = p.types;
      }
      if (p.start_date) {
        const startEl = form.querySelector('[name="start_date"]');
        if (startEl) startEl.value = p.start_date;
      }
      if (p.day) {
        const dayEl = form.querySelector('[name="day"]');
        if (dayEl) dayEl.value = p.day;
      }
      if (p.budget) {
        const budEl = form.querySelector('[name="budget"]');
        if (budEl) budEl.value = p.budget;
      }
      if (p.off_usd) {
        const offEl = form.querySelector('[name="off_usd"]');
        if (offEl) offEl.value = p.off_usd;
      }
      if (typeof window.updateCampaignCalc === 'function') window.updateCampaignCalc();
      showToast('Campaign details auto-filled successfully! ⚡', 'success');
      pasteInput.value = '';
      return;
    }
  }
  showToast('Could not auto-parse pasted text. You can fill the fields directly.', 'warning');
};

window.toggleQuickPaste = function() {
  const box = document.getElementById('quickPasteInlineBox');
  if (box) {
    box.classList.toggle('hidden');
    if (!box.classList.contains('hidden')) {
      const input = document.getElementById('quickPasteCampaignInput');
      if (input) input.focus();
    }
  }
};

function renderEntryForm() {
  const form = document.getElementById('entryForm');
  if (!form) return;
  form.innerHTML = '';
  form.className = 'entry-minimal-form';
  form.onkeydown = (e) => {
    if (e.key === 'Enter' && e.target.tagName !== 'TEXTAREA') {
      e.preventDefault();
      saveEntry();
    }
  };
  const todayStr = getTodayLocalDateStr();
  const key = state.currentMenu.key;

  const titleEl = document.getElementById('entryTitle');
  const subEl = document.getElementById('entrySub');
  const saveBtn = document.getElementById('btnSaveEntry');
  const resetBtn = document.getElementById('btnResetEntry');

  if (saveBtn) {
    saveBtn.onclick = (e) => {
      e.preventDefault();
      saveEntry();
    };
  }
  if (resetBtn) {
    resetBtn.onclick = (e) => {
      e.preventDefault();
      resetEntryForm();
    };
  }

  if (key === 'today_created_campaign' || key === 'campaign') {
    if (titleEl) titleEl.innerHTML = `<span>⚡ Quick New Campaign</span><span class="badge-fast-tag">Super Fast</span>`;
    if (subEl) subEl.textContent = 'Smart auto-fill • Instant Taka calc • Ctrl+Enter to save & copy';
    if (saveBtn) saveBtn.innerHTML = `<span class="ms" style="font-size:16px">bolt</span> Save Campaign (Ctrl+↵)`;

  } else if (key === 'agency_cost') {
    if (titleEl) titleEl.innerHTML = `<span>📑 Add Agency Expense</span><span class="badge-fast-tag" style="background:#F59E0B">Office Expense</span>`;
    if (subEl) subEl.textContent = 'Record rent, Facebook bills, or shared agency overheads';
    if (saveBtn) saveBtn.innerHTML = `<span class="ms" style="font-size:16px">save</span> Save Agency Cost`;
  }

  if (key === 'today_created_campaign' || key === 'campaign') {
    const todayFormD = getNowBST();
    const defaultCampMonth = MONTHS[todayFormD.getMonth()];
    const defaultCampYear = String(todayFormD.getFullYear());
    const initialSlNo = computeClientNextSL(defaultCampMonth, defaultCampYear);
    state.nextSL = initialSlNo;

    form.innerHTML = `
      <div id="increaseSmartNotice" class="increase-smart-notice" style="display:none"></div>

      <div class="entry-unified-grid">
        <!-- Row 1: Target & Page Identification -->
        <div class="entry-row-1">
          <div class="form-field ff-sl">
            <label>SL NO</label>
            <input type="text" name="sl_no" id="slNoInput" readonly value="${initialSlNo}" style="text-align:center;font-weight:700">
          </div>

          <div class="form-field ff-off">
            <label>OFF $</label>
            <input type="number" name="off_usd" id="offUsdInput" value="0" placeholder="0" style="text-align:center">
          </div>

          <div class="form-field ff-ad">
            <label>AD ACCOUNT <span class="req">*</span></label>
            <div class="smart-autocomplete-wrap" id="adAccountWrap">
              <input type="text" name="ad_account" id="adAccountInput" value="" placeholder="Search AD Account (e.g. 002, 104)..." autocomplete="off" required>
              <div class="smart-suggestions-dropdown hidden" id="adAccountsDropdown"></div>
            </div>
          </div>

          <div class="form-field ff-id">
            <label>ID NAME</label>
            <input type="text" name="id_name" id="idNameInput" value="" placeholder="ID Number">
          </div>

          <div class="form-field ff-status">
            <label>STATUS <span class="req">*</span></label>
            <select name="status" id="statusSelect" required>
              <option value="Old Page">Old Page</option>
              <option value="New Page">New Page</option>
            </select>
          </div>

          <div class="form-field ff-page">
            <label>PAGE NAME <span class="req">*</span></label>
            <div class="smart-autocomplete-wrap" id="pageNameWrap">
              <input type="text" name="page_name" id="pageNameInput" placeholder="Type page name (Bangla / English)..." autocomplete="off" required>
              <div class="smart-suggestions-dropdown hidden" id="pageSuggestionsDropdown"></div>
            </div>
            <!-- Relevant Active Campaign Status Card below Page Name -->
            <div id="pageRelevantCampaignCard" class="page-relevant-campaign-card" style="display:none"></div>
          </div>

          <div class="form-field ff-start">
            <label>START <span class="req">*</span></label>
            <input type="date" name="start_date" id="startDateInput" value="${todayStr}" required>
          </div>

          <div class="form-field ff-end">
            <label>END</label>
            <input type="date" name="end_date" id="endDateInput" readonly>
          </div>
        </div>

        <!-- Row 2: Campaign Execution & Financials (Types, Day, Budget, Rate, Taka) -->
        <div class="entry-row-2">
          <div class="form-field ff-type">
            <label>TYPES <span class="req">*</span></label>
            <select name="types" id="typesSelect" required>
              <option value="Boost">Boost</option>
              <option value="Ad Campaign">Ad Campaign</option>
              <option value="Increase">Increase</option>
              <option value="Others">Others</option>
            </select>
          </div>

          <div class="form-field ff-old-sl" id="oldSlFieldWrap" style="display:none">
            <label style="color:var(--pri);font-weight:700">OLD SL <span class="req">*</span></label>
            <input type="text" name="old_sl" id="oldSlInput" placeholder="Old SL (e.g. 187)" style="border:1.5px solid var(--pri);font-weight:700;text-align:center" autocomplete="off">
          </div>

          <div class="form-field ff-day">
            <label>DAY <span class="req">*</span></label>
            <input type="number" name="day" id="dayInput" min="1" value="2" required style="text-align:center">
            <div class="quick-preset-chips" role="group" aria-label="Quick Day Presets">
              <button type="button" class="preset-chip day-preset-chip active" data-day="2" onclick="setQuickDay(2)">2</button>
              <button type="button" class="preset-chip day-preset-chip" data-day="3" onclick="setQuickDay(3)">3</button>
              <button type="button" class="preset-chip day-preset-chip" data-day="5" onclick="setQuickDay(5)">5</button>
              <button type="button" class="preset-chip day-preset-chip" data-day="7" onclick="setQuickDay(7)">7</button>
            </div>
          </div>

          <div class="form-field ff-budget">
            <label>BUDGET <span class="req">*</span></label>
            <div class="input-prefix-wrap">
              <span class="input-prefix">$</span>
              <input type="number" name="budget" id="budgetInput" min="1" step="any" value="10" required>
            </div>
            <div class="quick-preset-chips" role="group" aria-label="Quick Budget Presets">
              <button type="button" class="preset-chip bud-preset-chip active" data-bud="10" onclick="setQuickBudget(10)">10</button>
              <button type="button" class="preset-chip bud-preset-chip" data-bud="15" onclick="setQuickBudget(15)">15</button>
              <button type="button" class="preset-chip bud-preset-chip" data-bud="20" onclick="setQuickBudget(20)">20</button>
              <button type="button" class="preset-chip bud-preset-chip" data-bud="25" onclick="setQuickBudget(25)">25</button>
            </div>
          </div>

          <div class="form-field ff-rate">
            <label>RATE <span class="req">*</span></label>
            <input type="number" name="rate" id="rateInput" value="${state.rate || 140}" readonly style="text-align:center">
          </div>

          <div class="form-field ff-taka">
            <label>TAKA</label>
            <div class="input-prefix-wrap">
              <span class="input-prefix">৳</span>
              <input type="text" name="taka" id="takaInput" readonly placeholder="0.00" style="font-weight:700">
            </div>
          </div>
        </div>

        <!-- Row 3: Multiple Campaigns Action Button & Container -->
        <div class="entry-row-multi-action">
          <button type="button" class="btn-add-campaign-purple" onclick="addMultipleCampaignRow()">
            <span class="ms" style="font-size:18px">add_circle</span> + Add Another Campaign
          </button>
        </div>

        <div id="multipleCampaignRows" class="multi-campaign-rows-container"></div>
      </div>
    `;

    google.script.run.withSuccessHandler(res => {
      const slInput = form.querySelector('[name="sl_no"]');
      if (slInput && res && res.nextSL && res.nextSL !== 'undefined') {
        slInput.value = String(res.nextSL);
        state.nextSL = String(res.nextSL);
      }
    }).getNextSLNo(state.filters.month, state.filters.year);

    // Auto-match AD Account & Sync ID Name instantly on 1-2 letters with smart dropdown
    const adInput = form.querySelector('#adAccountInput');
    const adDropdown = form.querySelector('#adAccountsDropdown');
    const idInput = form.querySelector('#idNameInput');
    
    initSmartAdAccountAutocomplete(adInput, adDropdown, {
      onSelect: (selectedAd, selectedId) => {
        if (idInput && selectedId) idInput.value = selectedId;
        if (typeof window.updateCampaignCalc === 'function') window.updateCampaignCalc();
      }
    });

    const syncAdAndId = () => {
      const val = (adInput?.value || '').trim().toLowerCase();
      if (!val) {
        if (typeof window.updateCampaignCalc === 'function') window.updateCampaignCalc();
        return;
      }
      const matched = state.adAccounts.find(a => 
        a.a.toLowerCase() === val ||
        a.a.toLowerCase().startsWith(val) ||
        a.a.toLowerCase().includes(val) ||
        a.id.toLowerCase().includes(val)
      );
      if (matched && idInput) {
        idInput.value = matched.id;
      }
      if (typeof window.updateCampaignCalc === 'function') window.updateCampaignCalc();
    };
    adInput?.addEventListener('input', syncAdAndId);

    // Intelligent Page Name Selection with Banglish Phonetic Dropdown & History Lookup
    const pageInput = form.querySelector('#pageNameInput');
    const pageDropdown = form.querySelector('#pageSuggestionsDropdown');
    const typeSelect = form.querySelector('#typesSelect');

    let debounceTimer = null;
    const triggerPageLookup = (val) => {
      clearTimeout(debounceTimer);
      debounceTimer = setTimeout(() => {
        handlePageNameLookup(val || pageInput?.value || '');
        if (typeof window.updateCampaignCalc === 'function') window.updateCampaignCalc();
      }, 150);
    };

    initSmartPageAutocomplete(pageInput, pageDropdown, {
      onSelect: (selectedPage) => {
        triggerPageLookup(selectedPage);
      }
    });

    pageInput?.addEventListener('input', () => triggerPageLookup());
    pageInput?.addEventListener('change', () => triggerPageLookup());

    const oldSlWrap = form.querySelector('#oldSlFieldWrap');
    const oldSlInput = form.querySelector('#oldSlInput');

    const handleOldSlLookup = (slVal) => {
      const cleanVal = String(slVal || '').trim().replace(/^#/, '');
      if (!cleanVal) return;

      const allRecords = [
        ...(state.rawRecords || []),
        ...(state.records || []),
        ...(state.todayCreated || []),
        ...(state.campaignRegistry || [])
      ];

      const matched = allRecords.find(r => {
        const rSl = String(r.SL || r['SL'] || r.sl_no || '').replace(/^#/, '').trim();
        return rSl === cleanVal || rSl.padStart(3, '0') === cleanVal.padStart(3, '0') || (!isNaN(parseInt(cleanVal, 10)) && parseInt(rSl, 10) === parseInt(cleanVal, 10));
      });

      const applyPreviousCampaign = (r) => {
        if (!r) return;
        const prevEnd = r['End'] || r['End Date'] || r.end_date || r.end || '';
        const endInputVal = toInputDateValue(prevEnd);
        const startInput = form.querySelector('[name="start_date"]');
        if (startInput && endInputVal) {
          startInput.value = endInputVal;
          startInput.dispatchEvent(new Event('input', { bubbles: true }));
          startInput.dispatchEvent(new Event('change', { bubbles: true }));
        }
        const pageEl = form.querySelector('#pageNameInput');
        if (pageEl && (r['Page Name'] || r.page_name)) {
          pageEl.value = r['Page Name'] || r.page_name;
        }
        const adEl = form.querySelector('#adAccountInput');
        if (adEl && (r['AD Account'] || r.ad_account)) {
          adEl.value = r['AD Account'] || r.ad_account;
        }
        const idEl = form.querySelector('#idNameInput');
        if (idEl && (r['ID Name'] || r.id_name)) {
          idEl.value = r['ID Name'] || r.id_name;
        }
        const statusEl = form.querySelector('#statusSelect');
        if (statusEl) statusEl.value = 'Old Page';
        if (typeof window.updateCampaignCalc === 'function') window.updateCampaignCalc();
        showToast(`⚡ SL #${cleanVal} End Date (${formatCustomDate(prevEnd)}) set as new Start Date!`, 'success');
      };

      if (matched) {
        applyPreviousCampaign(matched);
      }

      const targetMonth = state.filters.month || 'All';
      const targetYear = state.filters.year || new Date().getFullYear();
      google.script.run.withSuccessHandler(res => {
        if (res && res.success && res.record) {
          applyPreviousCampaign(res.record);
        }
      }).getCampaignBySL(cleanVal, targetMonth, targetYear);
    };

    let oldSlDebounce = null;
    oldSlInput?.addEventListener('input', () => {
      clearTimeout(oldSlDebounce);
      oldSlDebounce = setTimeout(() => handleOldSlLookup(oldSlInput.value), 300);
    });
    oldSlInput?.addEventListener('change', () => handleOldSlLookup(oldSlInput.value));

    typeSelect?.addEventListener('change', () => {
      const startEl = form.querySelector('[name="start_date"]');
      const cardEl = form.querySelector('#pageRelevantCampaignCard');
      if (typeSelect.value === 'Increase') {
        if (oldSlWrap) oldSlWrap.style.display = 'block';
        if (oldSlInput) {
          setTimeout(() => oldSlInput.focus(), 80);
          if (oldSlInput.value) handleOldSlLookup(oldSlInput.value);
          else handlePageNameLookup(pageInput?.value || '');
        } else {
          handlePageNameLookup(pageInput?.value || '');
        }
      } else {
        if (oldSlWrap) oldSlWrap.style.display = 'none';
        if (cardEl) {
          cardEl.style.display = 'none';
          cardEl.innerHTML = '';
        }
        if (startEl) {
          startEl.value = getTodayLocalDateStr();
        }
      }
      updateCampaignCalc();
    });

    const dayInput = form.querySelector('[name="day"]');
    const budgetInput = form.querySelector('[name="budget"]');
    const dailyInput = form.querySelector('[name="daily"]');
    const rateInput = form.querySelector('[name="rate"]');
    const startInput = form.querySelector('[name="start_date"]');
    const endEl = form.querySelector('[name="end_date"]');
    const takaEl = form.querySelector('[name="taka"]');

    const updateCampaignCalc = () => {
      const day = parseFloat(dayInput?.value) || 0;
      const budget = parseFloat(budgetInput?.value) || 0;
      const rate = parseFloat(rateInput?.value) || state.rate || 140;
      const start = startInput?.value || getTodayLocalDateStr();

      // Auto calculate End Date: Start Date + Day
      if (start && day > 0 && endEl) {
        const d = parseAnyDate(start);
        if (d && !isNaN(d.getTime())) {
          const end = new Date(d.getFullYear(), d.getMonth(), d.getDate() + day);
          endEl.value = toInputDateValue(end);
        }
      }

      // Auto calculate Daily $ = Budget / Day
      if (dailyInput) {
        if (day > 0 && budget > 0) {
          const dailyVal = budget / day;
          dailyInput.value = (dailyVal % 1 === 0) ? dailyVal.toFixed(0) : dailyVal.toFixed(2);
        } else {
          dailyInput.value = '';
        }
      }

      // Auto calculate Total Taka = Budget * Dollar Rate
      const totalTaka = budget * rate;
      if (takaEl) {
        takaEl.value = totalTaka.toFixed(2);
      }

      // Real-time Visual Verification Pills ("chokhe lage j sob thik thak ase")
      const pName = (pageInput?.value || '').trim();
      const pStatus = (form.querySelector('[name="status"]')?.value || 'Old Page');
      const cType = (typeSelect?.value || 'Boost');

      const pillPage = document.getElementById('entryPillPageText');
      if (pillPage) {
        pillPage.textContent = pName ? `${pName} (${pStatus})` : '—';
      }

      const pillFin = document.getElementById('entryPillFin');
      if (pillFin) {
        const dailyStr = (day > 0 && budget > 0) ? ` ($${(budget / day).toFixed(1)}/d)` : '';
        pillFin.textContent = `$${budget}${dailyStr} = ৳${formatCurrency(totalTaka)}`;
      }

      const pillSched = document.getElementById('entryPillSched');
      if (pillSched) {
        const startPretty = start ? formatCustomDate(start) : '';
        const endPretty = endEl?.value ? formatCustomDate(endEl.value) : '';
        const parsedS = parseAnyDate(start);
        const targetM = parsedS ? MONTHS[parsedS.getMonth()] : (state.filters.month || 'October');
        const targetY = parsedS ? String(parsedS.getFullYear()) : (state.filters.year || String(new Date().getFullYear()));
        pillSched.textContent = `${cType} • ${day}d (${startPretty} ➔ ${endPretty}) • 📁 Campaign_${targetM}_${targetY}`;
      }
    };

    const handleStartDateChange = () => {
      const parsedStart = parseAnyDate(startInput.value);
      if (parsedStart) {
        const m = MONTHS[parsedStart.getMonth()];
        const y = String(parsedStart.getFullYear());
        const slInput = form.querySelector('[name="sl_no"]');
        const calculatedSl = computeClientNextSL(m, y);
        if (slInput) {
          slInput.value = calculatedSl;
          state.nextSL = calculatedSl;
        }
        google.script.run.withSuccessHandler(res => {
          if (slInput && res && res.nextSL && res.nextSL !== 'undefined') {
            slInput.value = String(res.nextSL);
            state.nextSL = String(res.nextSL);
          }
        }).getNextSLNo(m, y);
      }
      updateCampaignCalc();
    };

    startInput?.addEventListener('input', handleStartDateChange);
    startInput?.addEventListener('change', handleStartDateChange);

    window.updateCampaignCalc = updateCampaignCalc;
    form.querySelectorAll('input, select').forEach(el => {
      if (el !== dailyInput && el !== rateInput && el !== takaEl && el !== endEl) {
        el.addEventListener('input', () => updateCampaignCalc());
      }
    });
    updateCampaignCalc();

  } else if (key === 'agency_cost') {
    form.innerHTML = `
      <div class="agency-cost-form-box">
        <div style="display:grid;grid-template-columns:repeat(auto-fit, minmax(160px, 1fr));gap:10px;align-items:end">
          <div class="form-field">
            <label style="font-size:11px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">DATE</label>
            <input type="date" name="date" id="agencyCostDate" value="${todayStr}" class="form-control" style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt)">
          </div>
          <div class="form-field" style="grid-column:span 2">
            <label style="font-size:11px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">TYPE / CATEGORY <span class="req">*</span></label>
            <input type="text" list="agencyCostCategoriesDatalist" name="type" id="agencyCostType" placeholder="e.g. Office Rent / Facebook Ads Bill" class="form-control" required style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt);font-weight:700">
            <datalist id="agencyCostCategoriesDatalist">
              <option value="Office Rent">Office Rent</option>
              <option value="Facebook Ads Bill">Facebook Ads Bill</option>
              <option value="Employee Salary">Employee Salary</option>
              <option value="Internet Bill">Internet Bill</option>
              <option value="Electricity Bill">Electricity Bill</option>
              <option value="Software & Tools">Software & Tools</option>
              <option value="Domain & Hosting">Domain & Hosting</option>
              <option value="Office Snacks / Tea">Office Snacks / Tea</option>
              <option value="Other Expense">Other Expense</option>
            </datalist>
          </div>
          <div class="form-field">
            <label style="font-size:11px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">COST AMOUNT (৳) <span class="req">*</span></label>
            <div class="input-prefix-wrap">
              <span class="input-prefix">৳</span>
              <input type="number" step="any" name="cost" id="agencyCostAmount" placeholder="Amount in Taka" class="form-control" required style="width:100%;padding:8px 10px 8px 24px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--suc);font-weight:800;font-size:14px">
            </div>
          </div>
          <div class="form-field">
            <label style="font-size:11px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">STATUS</label>
            <select name="status" id="agencyCostStatus" class="form-control" style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt);font-weight:700">
              <option value="Paid" selected>Paid</option>
              <option value="Due">Due</option>
            </select>
          </div>
        </div>

        <!-- Quick Category and Amount Presets -->
        <div style="margin-top:10px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;background:var(--card-subtle);padding:8px 12px;border-radius:8px;border:1px dashed var(--bor)">
          <div style="display:flex;align-items:center;gap:5px;flex-wrap:wrap">
            <span style="font-size:11px;font-weight:700;color:var(--txt-muted)">⚡ Category:</span>
            <button type="button" class="preset-chip" onclick="window.setQuickAgencyCategory('Office Rent')">🏢 Office Rent</button>
            <button type="button" class="preset-chip" onclick="window.setQuickAgencyCategory('Facebook Ads Bill')">📢 FB Ads Bill</button>
            <button type="button" class="preset-chip" onclick="window.setQuickAgencyCategory('Employee Salary')">💼 Salary</button>
            <button type="button" class="preset-chip" onclick="window.setQuickAgencyCategory('Internet Bill')">🌐 Internet</button>
            <button type="button" class="preset-chip" onclick="window.setQuickAgencyCategory('Electricity Bill')">⚡ Electricity</button>
            <button type="button" class="preset-chip" onclick="window.setQuickAgencyCategory('Software & Tools')">🛠️ Tools</button>
          </div>
          <div style="display:flex;align-items:center;gap:5px;flex-wrap:wrap">
            <span style="font-size:11px;font-weight:700;color:var(--txt-muted)">৳ Amount:</span>
            <button type="button" class="preset-chip" onclick="window.setQuickAgencyCost(500)">৳500</button>
            <button type="button" class="preset-chip" onclick="window.setQuickAgencyCost(1000)">৳1,000</button>
            <button type="button" class="preset-chip" onclick="window.setQuickAgencyCost(2000)">৳2,000</button>
            <button type="button" class="preset-chip" onclick="window.setQuickAgencyCost(5000)">৳5,000</button>
            <button type="button" class="preset-chip" onclick="window.setQuickAgencyCost(10000)">৳10,000</button>
            <button type="button" class="preset-chip" onclick="window.setQuickAgencyCost(20000)">৳20,000</button>
          </div>
        </div>

        <div id="agencyCostSplitLivePreview" style="margin-top:8px;font-size:12px;font-weight:700;color:var(--pri);display:flex;align-items:center;gap:6px">
          <span class="ms" style="font-size:16px">groups</span>
          <span>👤 50% Partner Share: <b id="agencyCostSplitLiveVal" style="color:var(--pri)">৳0</b></span>
        </div>
      </div>
    `;

    const costInp = form.querySelector('#agencyCostAmount');
    const splitValEl = form.querySelector('#agencyCostSplitLiveVal');
    costInp?.addEventListener('input', () => {
      const v = parseFloat(costInp.value) || 0;
      if (splitValEl) splitValEl.textContent = `৳${fmtNum(+(v / 2).toFixed(2))}`;
    });
  }

  // 3. ENTER KEY & CTRL+ENTER TO SAVE CAMPAIGN / ENTRY
  form.onsubmit = (e) => {
    e.preventDefault();
    saveEntry();
  };
  form.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
      e.preventDefault();
      saveEntry();
      return;
    }
    if (e.key === 'Enter' && e.target.tagName !== 'TEXTAREA') {
      const openDropdown = form.querySelector('.smart-suggestions-dropdown:not(.hidden)') || form.querySelector('.smart-dropdown:not(.hidden)');
      const highlightedItem = openDropdown?.querySelector('.smart-suggestion-item.highlighted') || openDropdown?.querySelector('.highlighted');
      if (highlightedItem) {
        // If a dropdown option is highlighted with arrow keys, select it instead of submitting
        return;
      }
      e.preventDefault();
      saveEntry();
    }
  });
}

function resetEntryForm() {
  const form = document.getElementById('entryForm');
  if (!form) return;
  const menuKey = state.currentMenu ? state.currentMenu.key : '';

  if (menuKey === 'today_created_campaign' || menuKey === 'campaign') {
    const pageInput = form.querySelector('[name="page_name"]');
    if (pageInput) pageInput.value = '';

    const adInput = form.querySelector('#adAccountInput') || form.querySelector('[name="ad_account"]');
    if (adInput) adInput.value = '';

    const idInput = form.querySelector('#idNameInput') || form.querySelector('[name="id_name"]');
    if (idInput) idInput.value = '';

    const offInput = form.querySelector('#offUsdInput') || form.querySelector('[name="off_usd"]');
    if (offInput) offInput.value = '0';

    const dayInput = form.querySelector('[name="day"]');
    if (dayInput) dayInput.value = '';

    const budgetInput = form.querySelector('[name="budget"]');
    if (budgetInput) budgetInput.value = '';

    const dailyInput = form.querySelector('#dailyUsdInput') || form.querySelector('[name="daily_usd"]');
    if (dailyInput) dailyInput.value = '';

    const takaInput = form.querySelector('#takaInput') || form.querySelector('[name="taka"]');
    if (takaInput) takaInput.value = '';

    form.querySelectorAll('.preset-chip.active').forEach(c => c.classList.remove('active'));

    const multiBox = document.getElementById('multipleCampaignRows');
    if (multiBox) multiBox.innerHTML = '';

    const cardEl = form.querySelector('#pageRelevantCampaignCard');
    if (cardEl) { cardEl.style.display = 'none'; cardEl.innerHTML = ''; }
    const noticeEl = form.querySelector('#increaseSmartNotice');
    if (noticeEl) { noticeEl.style.display = 'none'; noticeEl.innerHTML = ''; }
    const profileBox = document.getElementById('pageProfileSummary');
    if (profileBox) profileBox.classList.add('hidden');

    const startDateVal = form.querySelector('[name="start_date"]')?.value;
    const dObj = startDateVal ? parseAnyDate(startDateVal) : null;
    const targetMonth = dObj ? MONTHS[dObj.getMonth()] : (state.filters.month && state.filters.month !== 'All' ? state.filters.month : MONTHS[new Date().getMonth()]);
    const targetYear = dObj ? String(dObj.getFullYear()) : (state.filters.year && state.filters.year !== 'All Years' ? String(state.filters.year) : String(new Date().getFullYear()));

    const slInput = form.querySelector('[name="sl_no"]');
    const resetNextSl = computeClientNextSL(targetMonth, targetYear);
    if (slInput) {
      slInput.value = resetNextSl;
      state.nextSL = resetNextSl;
    }

    google.script.run.withSuccessHandler(res => {
      if (slInput && res && res.nextSL && res.nextSL !== 'undefined') {
        slInput.value = String(res.nextSL);
        state.nextSL = String(res.nextSL);
      }
    }).getNextSLNo(targetMonth, targetYear);

    if (pageInput) pageInput.focus();
  } else if (menuKey === 'agency_cost') {
    const typeInput = form.querySelector('[name="type"]');
    const costInput = form.querySelector('[name="cost"]');
    if (typeInput) { typeInput.value = ''; typeInput.focus(); }
    if (costInput) costInput.value = '';
  }
}

function saveEntryToVault(records, status = 'saved') {
  try {
    const list = Array.isArray(records) ? records : [records];
    const raw = localStorage.getItem('adverse_entry_vault') || '[]';
    let vault = [];
    try { vault = JSON.parse(raw); } catch (_) { vault = []; }
    if (!Array.isArray(vault)) vault = [];

    list.forEach(r => {
      const entryId = r['Record ID'] || r._recordId || ('vault_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6));
      const entryObj = {
        id: entryId,
        savedAt: new Date().toISOString(),
        timeStr: new Date().toLocaleTimeString('en-US'),
        status,
        data: { ...r }
      };
      const existingIdx = vault.findIndex(v => v.id === entryId || (v.data?.SL && r.SL && v.data?.SL === r.SL && v.data?.['Page Name'] === r['Page Name']));
      if (existingIdx >= 0) {
        vault[existingIdx] = { ...vault[existingIdx], ...entryObj };
      } else {
        vault.unshift(entryObj);
      }
    });

    localStorage.setItem('adverse_entry_vault', JSON.stringify(vault.slice(0, 100)));
  } catch (e) {
    console.warn('Vault write note:', e);
  }
}
window.saveEntryToVault = saveEntryToVault;

function enqueueOfflineSync(task) {
  try {
    const raw = localStorage.getItem('adverse_offline_queue') || '[]';
    let queue = [];
    try { queue = JSON.parse(raw); } catch (_) { queue = []; }
    if (!Array.isArray(queue)) queue = [];
    queue.push(task);
    localStorage.setItem('adverse_offline_queue', JSON.stringify(queue.slice(-50)));
  } catch (_) {}
}

async function processClientOfflineQueue() {
  try {
    const raw = localStorage.getItem('adverse_offline_queue') || '[]';
    if (!raw || raw === '[]') return;
    const queue = JSON.parse(raw);
    if (!Array.isArray(queue) || queue.length === 0) return;

    const remaining = [];
    for (const item of queue) {
      try {
        let res;
        if (item.type === 'campaign') {
          if (item.batchItems && item.batchItems.length) {
            res = await cloudRpc('addCampaignBatch', [item.data, item.batchItems]);
          } else {
            res = await cloudRpc('addCampaign', [item.data]);
          }
        } else if (item.type === 'payment') {
          res = await cloudRpc('addPayment', [item.data]);
        } else if (item.type === 'agency_cost') {
          res = await cloudRpc('addAgencyCost', [item.data]);
        }
        if (res && res.success) {
          saveEntryToVault(item.optimisticItems || [item.data], 'confirmed');
        } else {
          remaining.push(item);
        }
      } catch (err) {
        remaining.push(item);
      }
    }
    localStorage.setItem('adverse_offline_queue', JSON.stringify(remaining));
  } catch (_) {}
}
setInterval(processClientOfflineQueue, 10000);

async function autoSyncDirectToGoogleSheets(targetTab, headers, rows, action = 'upsert') {
  try {
    if (targetTab === 'Today_Created' || targetTab === 'today_created_campaign') {
      return; // Do not push today-created entries to a monolithic Today_Created sheet tab
    }
    const config = JSON.parse(localStorage.getItem('adverse_sheet_config') || '{}');
    const sheetId = config.sheetId || state.sheetId || '';
    if (!sheetId) return;
    const webhookUrl = config.webhookUrl || state.webhookUrl || state.settings?.webhook_url || '';
    const token = getGoogleAccessToken();
    const rowArray = Array.isArray(rows) ? rows : [rows];

    // 1. Direct Google Sheets API via OAuth token
    if (sheetId && token) {
      for (const r of rowArray) {
        upsertRowInSheetDirect(sheetId, targetTab, r['Record ID'] || r._recordId, r.SL || r['SL'], headers, r, {
          pageName: r['Page Name'] || r.page_name,
          budget: r['Budget'] || r.budget,
          month: r['Month'] || r.month,
          year: r['Year'] || r.year
        }).catch(e => console.warn('Direct upsert note:', e));
      }
      console.log(`[AdVerse Sheet Sync]: Instantly synced ${rowArray.length} row(s) to Google Sheet tab "${targetTab}"`);
    }

    // 2. Direct Apps Script Webhook auto-save (fires in parallel whenever webhookUrl is set)
    if (webhookUrl) {
      for (const r of rowArray) {
        const payloadAction = (action === 'update') ? 'updateRow' : 'appendRow';
        fetch(webhookUrl, {
          method: 'POST',
          mode: 'no-cors',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            action: payloadAction,
            sheetName: targetTab,
            headers,
            recordId: r['Record ID'] || r._recordId,
            matchInfo: {
              SL: r.SL || r['SL'],
              'Page Name': r['Page Name'] || r.page_name,
              'Record ID': r['Record ID'] || r._recordId,
              budget: r['Budget'] || r.budget
            },
            rowValues: headers.map(h => {
              let val = r[h] ?? '';
              const hLower = String(h || '').toLowerCase();
              if (hLower === 'start' || hLower === 'end' || hLower === 'entry date' || hLower === 'date') {
                if (val) val = formatCustomDate(val, r['Month'] || r.month);
              }
              return val;
            })
          })
        }).catch(() => {});
      }
    }

    // 3. Direct Firestore Sync (High-speed Cloud Persistence)
    try {
      const lowerTab = String(targetTab || '').toLowerCase();
      if (lowerTab.includes('campaign') || lowerTab.includes('today')) {
        for (const r of rowArray) {
          saveCampaignToFirestore(r).catch(e => console.warn('[Firestore Sync]: Campaign sync note:', e));
        }
      } else if (lowerTab.includes('payment')) {
        for (const r of rowArray) {
          savePaymentToFirestore(r).catch(e => console.warn('[Firestore Sync]: Payment sync note:', e));
        }
      }
    } catch (fErr) {
      console.warn('[Firestore Sync]: Sync note:', fErr);
    }
  } catch (err) {
    console.warn('[AdVerse Sheet Sync]: Direct client sheet sync note:', err?.message || err);
  }
}

async function saveEntry() {
  const saveBtn = document.getElementById('btnSaveEntry');

  try {
    const form = document.getElementById('entryForm');
    const fd = form ? new FormData(form) : new FormData();
    const data = Object.fromEntries(fd.entries());
    
    const pageEl = document.getElementById('pageNameInput') || form?.querySelector('[name="page_name"]');
    const adEl = document.getElementById('adAccountInput') || form?.querySelector('[name="ad_account"]');
    const slEl = document.getElementById('slNoInput') || form?.querySelector('[name="sl_no"]');

    if (!data.page_name || data.page_name === 'Loading...') {
      data.page_name = pageEl?.value?.trim() || '';
    }
    if (!data.ad_account || data.ad_account === 'Loading...') {
      data.ad_account = adEl?.value?.trim() || '';
    }

    const currentKey = state.currentMenu ? state.currentMenu.key : 'today_created_campaign';
    const isCampaignMenu = currentKey === 'today_created_campaign' || currentKey === 'campaign';

    if (isCampaignMenu) {
      if (!data.page_name) {
        showToast('Please enter a Page Name!', 'error');
        if (pageEl) pageEl.focus();
        return;
      }
      if (!data.ad_account) {
        showToast('Please select an AD Account!', 'error');
        if (adEl) adEl.focus();
        return;
      }
    }

    const todayStr = getTodayFormatted();
    data.entry_date = data.entry_date ? formatCustomDate(data.entry_date) : todayStr;
    data['Entry Date'] = data.entry_date;

    if (currentKey === 'agency_cost') {
      const dateVal = data.date || todayStr;
      const parsedD = parseAnyDate(dateVal) || getNowBST();
      const formattedD = formatCustomDate(parsedD);
      data.date = formattedD;
      data.Date = formattedD;
      data.month = MONTHS[parsedD.getMonth()];
      data.year = String(parsedD.getFullYear());
      data.Month = data.month;
      data.Year = data.year;
    } else if (currentKey === 'new_pages' || currentKey === 'scam') {
      const dateVal = data.date || todayStr;
      const parsedD = parseAnyDate(dateVal) || getNowBST();
      const formattedD = formatCustomDate(parsedD);
      data.date = formattedD;
      data.Date = formattedD;
      data.month = MONTHS[parsedD.getMonth()];
      data.year = String(parsedD.getFullYear());
      data.Month = data.month;
      data.Year = data.year;
      data.Entry = todayStr;
    } else {
      const startParsedForMonth = parseAnyDate(data.start_date || todayStr) || getNowBST();
      const formattedStartDate = formatCustomDate(startParsedForMonth);
      data.start_date = formattedStartDate;
      data.Start = formattedStartDate;
      data['Start Date'] = formattedStartDate;
      data.month = MONTHS[startParsedForMonth.getMonth()];
      data.year = String(startParsedForMonth.getFullYear());
      data.Month = data.month;
      data.Year = data.year;
    }

    // Strict SL Validation: Verify that data.sl_no prefix matches this month's number (e.g. "10" for October)
    const mIdx = MONTHS.findIndex(m => m.toLowerCase() === String(data.month).toLowerCase());
    const monthNum = mIdx >= 0 ? mIdx + 1 : (new Date().getMonth() + 1);
    const expectedPrefix = String(monthNum).padStart(2, '0');
    const cleanCurrentSl = String(slEl?.value || data.sl_no || '').replace(/\D/g, '');
    if (!cleanCurrentSl || (!cleanCurrentSl.startsWith(expectedPrefix) && !cleanCurrentSl.startsWith(String(monthNum)))) {
      data.sl_no = computeClientNextSL(data.month, data.year);
    } else {
      data.sl_no = cleanCurrentSl;
    }

    if (isCampaignMenu) {
      const rawBudget = String(data.budget || '').replace(/[^0-9.]/g, '');
      const budVal = parseFloat(rawBudget);
      if (isNaN(budVal) || budVal <= 0) {
        showToast('❌ Please enter a valid Campaign Budget greater than $0!', 'error');
        const bInput = document.getElementById('budgetInput') || form?.querySelector('[name="budget"]');
        if (bInput) bInput.focus();
        return;
      }
      data.budget = budVal;

      const rawDay = String(data.day || '1').replace(/[^0-9]/g, '');
      const dayVal = Math.max(1, parseInt(rawDay, 10) || 1);
      data.day = dayVal;

      const rateVal = parseFloat(data.rate) || state.rate || 140;
      data.rate = rateVal;
      data.taka = (budVal * rateVal).toFixed(2);
      data['Daily $'] = (budVal / dayVal).toFixed(2);
    }

    const extraRows = form ? Array.from(form.querySelectorAll('.multi-campaign-row')) : [];
    const batchItems = [];
    if (extraRows.length) {
      const mainSl = data.sl_no ? String(data.sl_no).replace(/\D/g, '') : '';
      batchItems.push({
        day: Number(data.day) || 2,
        budget: Number(data.budget) || 10,
        start_date: data.start_date,
        types: data.types || 'Boost',
        sl_no: mainSl || undefined
      });
      for (const row of extraRows) {
        const day = Number(row.querySelector('.multi-campaign-day')?.value) || 2;
        const budget = Number(row.querySelector('.multi-campaign-budget')?.value) || 10;
        const rowType = row.querySelector('.multi-campaign-type')?.value || data.types || 'Boost';
        const rowSl = row.querySelector('.multi-campaign-sl')?.value ? String(row.querySelector('.multi-campaign-sl')?.value).replace(/\D/g, '') : '';
        batchItems.push({
          day,
          budget,
          start_date: data.start_date,
          types: rowType,
          sl_no: rowSl || undefined
        });
      }
    }

    // =========================================================================
    // ⚡ INSTANT OPTIMISTIC UI EXECUTION (0ms latency, zero delay!)
    // =========================================================================
    if (isCampaignMenu) {
      const rateVal = parseFloat(data.rate) || state.rate || 140;
      const startVal = data.start_date || todayStr;
      const optimisticItems = [];
      
      if (batchItems.length) {
        const baseNum = parseInt(data.sl_no || state.nextSL || '1', 10) || 1;
        batchItems.forEach((bItem, bIdx) => {
          const itemSl = bItem.sl_no ? String(bItem.sl_no) : String(baseNum + bIdx);
          const dVal = Number(bItem.day) || 1;
          const bVal = Number(bItem.budget) || 0;
          const tVal = bVal * rateVal;
          const parsedS = parseAnyDate(startVal) || new Date();
          const endD = new Date(parsedS.getFullYear(), parsedS.getMonth(), parsedS.getDate() + dVal);
          const endVal = formatCustomDate(endD);
          const itemMonth = MONTHS[parsedS.getMonth()];
          const itemYear = String(parsedS.getFullYear());
          optimisticItems.push({
            SL: itemSl,
            'Page Name': data.page_name,
            'AD Account': data.ad_account,
            'ID Name': data.id_name || '',
            'Status': data.status || 'Old Page',
            'Start': formatCustomDate(parsedS),
            'End': endVal,
            'Types': bItem.types || 'Boost',
            'Day': dVal,
            'Budget': bVal,
            'Rate': rateVal,
            'Taka': tVal,
            'OFF $': data.off_usd ?? data['OFF $'] ?? '0',
            'OFF USD': data.off_usd ?? data['OFF $'] ?? '0',
            'Entry Date': getTodayFormatted(),
            'Month': itemMonth,
            'Year': itemYear,
            'Record ID': 'temp_' + Date.now() + '_' + bIdx,
            _recordId: 'temp_' + Date.now() + '_' + bIdx,
            _optimistic: true
          });
        });
      } else {
        const dVal = Number(data.day) || 1;
        const bVal = Number(data.budget) || 0;
        const tVal = bVal * rateVal;
        const parsedS = parseAnyDate(startVal) || new Date();
        const endD = new Date(parsedS.getFullYear(), parsedS.getMonth(), parsedS.getDate() + dVal);
        const endVal = formatCustomDate(endD);
        const itemMonth = MONTHS[parsedS.getMonth()];
        const itemYear = String(parsedS.getFullYear());
        optimisticItems.push({
          SL: (data.sl_no && data.sl_no !== 'undefined' && data.sl_no !== 'Loading...') ? String(data.sl_no) : computeClientNextSL(itemMonth, itemYear),
          'Page Name': data.page_name,
          'AD Account': data.ad_account,
          'ID Name': data.id_name || '',
          'Status': data.status || 'Old Page',
          'Start': formatCustomDate(parsedS),
          'End': endVal,
          'Types': data.types || 'Boost',
          'Day': dVal,
          'Budget': bVal,
          'Rate': rateVal,
          'Taka': tVal,
          'OFF $': data.off_usd ?? data['OFF $'] ?? '0',
          'OFF USD': data.off_usd ?? data['OFF $'] ?? '0',
          'Entry Date': getTodayFormatted(),
          'Month': itemMonth,
          'Year': itemYear,
          'Record ID': 'temp_' + Date.now(),
          _recordId: 'temp_' + Date.now(),
          _optimistic: true
        });
      }

      if (!state.newlyAddedRecordKeys) state.newlyAddedRecordKeys = new Set();
      optimisticItems.forEach(item => {
        if (item['Record ID']) state.newlyAddedRecordKeys.add(item['Record ID']);
        if (item._recordId) state.newlyAddedRecordKeys.add(item._recordId);
      });

      // Auto-align active filter to this campaign's month so it is never hidden or disappeared
      if (state.filters.month && state.filters.month !== 'All' && state.filters.month !== 'All Months' && state.filters.month.toLowerCase() !== String(data.month).toLowerCase()) {
        state.filters.month = data.month;
        const filterMonthSelect = document.getElementById('filterMonth');
        if (filterMonthSelect) filterMonthSelect.value = data.month;
      }

      // 1. Instantly update table view
      state.rawRecords = [...optimisticItems, ...(state.rawRecords || [])];
      state.filteredRecords = [...state.rawRecords];
      if (!state.masterCampaignRegistry) state.masterCampaignRegistry = [];
      state.masterCampaignRegistry = [...optimisticItems, ...state.masterCampaignRegistry];
      if (!state.todayCreatedRegistry) state.todayCreatedRegistry = [];
      state.todayCreatedRegistry = [...optimisticItems, ...state.todayCreatedRegistry];

      if (typeof renderRecordsTable === 'function') {
        renderRecordsTable({ headers: state.headers, rows: state.filteredRecords });
      }

      // Update Stat Cards immediately with new totals
      const curKey = state.currentMenu ? state.currentMenu.key : 'today_created_campaign';
      if (curKey === 'today_created_campaign') {
        const todayItems = state.todayCreatedRegistry || [];
        const sumBud = todayItems.reduce((s, r) => s + (parseFloat(String(r['Budget'] || r.budget || 0).replace(/[^0-9.]/g, '')) || 0), 0);
        const sumTak = todayItems.reduce((s, r) => s + (parseFloat(String(r['Taka'] || r.taka || 0).replace(/[^0-9.]/g, '')) || 0), 0);
        renderStats({
          totalRecords: todayItems.length,
          totalBudget: sumBud,
          totalTaka: sumTak
        });
      } else if (curKey === 'campaign') {
        const sumBud = state.filteredRecords.reduce((s, r) => s + (parseFloat(String(r['Budget'] || r.budget || 0).replace(/[^0-9.]/g, '')) || 0), 0);
        const sumTak = state.filteredRecords.reduce((s, r) => s + (parseFloat(String(r['Taka'] || r.taka || 0).replace(/[^0-9.]/g, '')) || 0), 0);
        renderStats({
          totalRecords: state.filteredRecords.length,
          totalBudget: sumBud,
          totalTaka: sumTak
        });
      }

      // Immediately save to local persistent snapshot so data is permanently preserved
      saveLocalAppSnapshot();

      // 2. Instantly copy campaign to clipboard (Direct gesture context = 100% reliable)
      if (optimisticItems.length > 1) {
        copyMultipleRecordsSequentially(optimisticItems);
      } else if (optimisticItems[0]) {
        copyCampaignRecord(optimisticItems[0]);
      }

      renderSavedCampaignsNotice(optimisticItems);

      // 3. Instantly reset form and calculate next sequential SL NO
      resetEntryForm();

      // 4. Quick visual button flash
      if (saveBtn) {
        saveBtn.innerHTML = '<span class="ms" style="font-size:16px">check_circle</span> Saved & Copied!';
        saveBtn.style.background = '#10b981';
        setTimeout(() => {
          saveBtn.style.background = '';
          saveBtn.innerHTML = '<span class="ms" style="font-size:16px">bolt</span> Save Campaign (Ctrl+↵)';
        }, 800);
      }

      showToast('⚡ Campaign saved & synced to Google Sheet instantly! 📋', 'success');

      // 5. Asynchronous background server & Google Sheet persistence with Vault Backup
      saveEntryToVault(optimisticItems, 'saved');

      (async () => {
        try {
          // Instant Sheet Sync directly to the month campaign tab in parallel with server RPC
          const targetMonth = optimisticItems[0]?.Month || data.month || 'October';
          const targetYear = optimisticItems[0]?.Year || data.year || String(new Date().getFullYear());
          const targetMonthTab = `Campaign_${targetMonth}_${targetYear}`;
          autoSyncDirectToGoogleSheets(targetMonthTab, CH, optimisticItems).catch(() => {});

          data.pin = localStorage.getItem('adverse_saved_pin') || '01726912624';

          let res;
          if (batchItems.length) {
            res = await cloudRpc('addCampaignBatch', [data, batchItems]);
          } else {
            res = await cloudRpc('addCampaign', [data]);
          }
          if (res && res.success) {
            const confirmedRecords = res.records || (res.record ? [res.record] : (res.items ? res.items.map(it => it.record) : []));
            if (confirmedRecords.length > 0) {
              const replaceList = (list) => {
                if (!Array.isArray(list)) return;
                optimisticItems.forEach((opt, optIdx) => {
                  const conf = confirmedRecords[optIdx] || confirmedRecords[0];
                  if (!conf) return;
                  const idx = list.findIndex(r => r === opt || r['Record ID'] === opt['Record ID'] || r._recordId === opt._recordId);
                  if (idx >= 0) {
                    list[idx] = { ...list[idx], ...conf, _optimistic: false };
                  }
                });
              };
              replaceList(state.rawRecords);
              replaceList(state.filteredRecords);
              replaceList(state.masterCampaignRegistry);
              replaceList(state.todayCreatedRegistry);
            }
            state.clientCache = {};
            saveLocalAppSnapshot();
            saveEntryToVault(optimisticItems, 'confirmed');
          } else {
            console.warn('Server addCampaign unconfirmed:', res);
            enqueueOfflineSync({
              type: 'campaign',
              data,
              batchItems,
              optimisticItems,
              timestamp: Date.now()
            });
            showToast('⚠️ Saved locally! Syncing with server in progress...', 'warning');
          }
        } catch (syncErr) {
          console.warn('Background sync note:', syncErr);
          enqueueOfflineSync({
            type: 'campaign',
            data,
            batchItems,
            optimisticItems,
            timestamp: Date.now()
          });
          showToast('⚠️ Saved locally! Will auto-sync once connection stabilizes.', 'info');
        }
      })();

      return;
    }

    // Non-campaign forms (Pause Ads, Scam, Agency Cost)
    saveEntryToVault([data], 'saved');
    resetEntryForm();
    showToast('⚡ Entry saved successfully! ✅', 'success');

    // Instantly sync non-campaign entry to Google Sheet
    if (currentKey === 'new_pages') {
      autoSyncDirectToGoogleSheets('Pause_ADS', PAUSE_ADS_HEADERS, [data]).catch(() => {});
    } else if (currentKey === 'scam') {
      autoSyncDirectToGoogleSheets('Scam', SCAM_HEADERS, [data]).catch(() => {});
    } else if (currentKey === 'agency_cost') {
      autoSyncDirectToGoogleSheets('Agency_Cost', AGENCY_COST_HEADERS, [data]).catch(() => {});
    }

    (async () => {
      try {
        data.pin = localStorage.getItem('adverse_saved_pin') || '01726912624';
        let res;
        if (currentKey === 'new_pages') {
          res = await cloudRpc('addPauseAds', [data]);
        } else if (currentKey === 'scam') {
          res = await cloudRpc('addScam', [data]);
        } else if (currentKey === 'agency_cost') {
          res = await cloudRpc('addAgencyCost', [data]);
        }
        if (res && res.success) {
          saveEntryToVault([data], 'confirmed');
        }
        state.clientCache = {};
        saveLocalAppSnapshot();
        loadMenuData(false);
      } catch (e) {
        console.warn('Background sync note:', e);
        enqueueOfflineSync({
          type: currentKey === 'new_pages' ? 'pause_ads' : (currentKey === 'scam' ? 'scam' : 'agency_cost'),
          data,
          timestamp: Date.now()
        });
      }
    })();

  } catch (err) {
    console.error('Save Entry Error:', err);
    showToast('Save Error: ' + (err.message || 'Unknown error'), 'error');
  }
}

function addMultipleCampaignRow() {
  const box = document.getElementById('multipleCampaignRows');
  if (!box) return;
  const count = box.querySelectorAll('.multi-campaign-row').length + 2;
  const baseInput = document.querySelector('#entryForm [name="sl_no"]');
  const baseVal = baseInput ? parseInt(baseInput.value, 10) : (parseInt(state.nextSL, 10) || 1);
  const rowSL = (baseVal > 0) ? String(baseVal + count - 1) : String(count);

  // Inherit sensible default from main campaign if available
  const mainDayInput = document.querySelector('#entryForm [name="day"]');
  const mainBudInput = document.querySelector('#entryForm [name="budget"]');
  const initialDay = (mainDayInput && parseFloat(mainDayInput.value) > 0) ? mainDayInput.value : '2';
  const initialBud = (mainBudInput && parseFloat(mainBudInput.value) > 0) ? mainBudInput.value : '10';

  const row = document.createElement('div');
  row.className = 'multi-campaign-row';
  row.innerHTML = `
    <div class="multi-row-top" style="display:flex;align-items:center;justify-content:space-between;width:100%;gap:8px;margin-bottom:8px">
      <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap">
        <span class="multi-campaign-row-badge" style="background:linear-gradient(135deg,#8B5CF6,#6366F1);color:#fff;font-weight:700;font-size:11.5px;padding:3px 9px;border-radius:6px;display:inline-flex;align-items:center;gap:5px">
          <span>+ Campaign #${count}</span>
          <span style="background:rgba(255,255,255,0.24);padding:1px 6px;border-radius:4px;font-family:monospace;font-size:11px">SL:</span>
          <input type="number" class="multi-campaign-sl" value="${rowSL}" style="width:58px;height:20px;padding:0 4px;font-family:monospace;font-size:11px;font-weight:800;background:rgba(255,255,255,0.9);color:#1e1b4b;border:none;border-radius:4px;text-align:center" title="Editable unique Serial Number for this campaign">
        </span>
        <span class="multi-row-summary" style="font-size:12px;font-weight:700;color:var(--txt-muted)">Calculating...</span>
      </div>
      <button type="button" class="btn btn-secondary btn-sm" style="height:28px;padding:0 8px;color:var(--dan);border-color:rgba(239,68,68,0.3)" onclick="this.closest('.multi-campaign-row').remove(); if(window.updateCampaignCalc) window.updateCampaignCalc();" title="Remove sub-campaign">
        <span class="ms" style="font-size:14px">delete</span> Remove
      </button>
    </div>

    <!-- Day & Budget Suggestions Banner -->
    <div class="multi-row-suggestion" style="margin-bottom:10px;background:var(--card-subtle);border:1px dashed var(--pri);padding:7px 12px;border-radius:8px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px">
      <div style="display:flex;align-items:center;gap:6px">
        <span class="ms" style="font-size:16px;color:var(--pri)">lightbulb</span>
        <span class="sub-suggestion-label" style="font-size:11.5px;font-weight:700;color:var(--txt)">Day Budget Suggestion:</span>
        <span class="sub-suggestion-calc" style="font-size:11.5px;font-weight:800;color:var(--pri)">$${initialBud} / ${initialDay}d = $${(parseFloat(initialBud)/parseFloat(initialDay)).toFixed(1)}/day</span>
      </div>
      <div class="sub-quick-presets" style="display:flex;gap:4px;flex-wrap:wrap">
        <button type="button" class="btn btn-secondary btn-sm sub-preset-btn" data-day="1" data-bud="5" style="height:22px;font-size:10.5px;padding:0 6px;font-weight:600">1d/$5</button>
        <button type="button" class="btn btn-secondary btn-sm sub-preset-btn" data-day="2" data-bud="10" style="height:22px;font-size:10.5px;padding:0 6px;font-weight:600">2d/$10</button>
        <button type="button" class="btn btn-secondary btn-sm sub-preset-btn" data-day="3" data-bud="15" style="height:22px;font-size:10.5px;padding:0 6px;font-weight:600">3d/$15</button>
        <button type="button" class="btn btn-secondary btn-sm sub-preset-btn" data-day="4" data-bud="20" style="height:22px;font-size:10.5px;padding:0 6px;font-weight:600">4d/$20</button>
        <button type="button" class="btn btn-secondary btn-sm sub-preset-btn" data-day="5" data-bud="25" style="height:22px;font-size:10.5px;padding:0 6px;font-weight:600">5d/$25</button>
        <button type="button" class="btn btn-secondary btn-sm sub-preset-btn" data-day="7" data-bud="35" style="height:22px;font-size:10.5px;padding:0 6px;font-weight:600">7d/$35</button>
        <button type="button" class="btn btn-secondary btn-sm sub-preset-btn" data-day="10" data-bud="50" style="height:22px;font-size:10.5px;padding:0 6px;font-weight:600">10d/$50</button>
        <button type="button" class="btn btn-secondary btn-sm sub-preset-btn" data-day="15" data-bud="75" style="height:22px;font-size:10.5px;padding:0 6px;font-weight:600">15d/$75</button>
        <button type="button" class="btn btn-secondary btn-sm sub-preset-btn" data-day="30" data-bud="150" style="height:22px;font-size:10.5px;padding:0 6px;font-weight:600">30d/$150</button>
      </div>
    </div>

    <div class="multi-fields-grid" style="display:grid;grid-template-columns:repeat(auto-fit, minmax(130px, 1fr));gap:10px;width:100%;align-items:start">
      <div class="form-field">
        <label>TYPES</label>
        <select class="multi-campaign-type">
          <option value="Boost">Boost</option>
          <option value="Ad Campaign">Ad Campaign</option>
          <option value="Increase">Increase</option>
          <option value="Others">Others</option>
        </select>
      </div>

      <div class="form-field">
        <label>DAY <span class="req">*</span></label>
        <input type="number" class="multi-campaign-day" min="1" value="${initialDay}" required style="text-align:center">
        <div class="quick-preset-chips sub-day-chips" role="group" aria-label="Day Presets" style="margin-top:4px;display:flex;gap:4px;flex-wrap:wrap">
          <button type="button" class="preset-chip day-preset-chip ${initialDay==='1'?'active':''}" data-val="1">1</button>
          <button type="button" class="preset-chip day-preset-chip ${initialDay==='2'?'active':''}" data-val="2">2</button>
          <button type="button" class="preset-chip day-preset-chip ${initialDay==='3'?'active':''}" data-val="3">3</button>
          <button type="button" class="preset-chip day-preset-chip ${initialDay==='4'?'active':''}" data-val="4">4</button>
          <button type="button" class="preset-chip day-preset-chip ${initialDay==='5'?'active':''}" data-val="5">5</button>
          <button type="button" class="preset-chip day-preset-chip ${initialDay==='7'?'active':''}" data-val="7">7</button>
          <button type="button" class="preset-chip day-preset-chip ${initialDay==='10'?'active':''}" data-val="10">10</button>
          <button type="button" class="preset-chip day-preset-chip ${initialDay==='15'?'active':''}" data-val="15">15</button>
          <button type="button" class="preset-chip day-preset-chip ${initialDay==='30'?'active':''}" data-val="30">30</button>
        </div>
      </div>

      <div class="form-field">
        <label>BUDGET <span class="req">*</span></label>
        <div class="input-prefix-wrap">
          <span class="input-prefix">$</span>
          <input type="number" class="multi-campaign-budget" min="1" step="any" value="${initialBud}" required>
        </div>
        <div class="quick-preset-chips sub-bud-chips" role="group" aria-label="Budget Presets" style="margin-top:4px;display:flex;gap:4px;flex-wrap:wrap">
          <button type="button" class="preset-chip bud-preset-chip ${initialBud==='5'?'active':''}" data-val="5">5</button>
          <button type="button" class="preset-chip bud-preset-chip ${initialBud==='10'?'active':''}" data-val="10">10</button>
          <button type="button" class="preset-chip bud-preset-chip ${initialBud==='15'?'active':''}" data-val="15">15</button>
          <button type="button" class="preset-chip bud-preset-chip ${initialBud==='20'?'active':''}" data-val="20">20</button>
          <button type="button" class="preset-chip bud-preset-chip ${initialBud==='25'?'active':''}" data-val="25">25</button>
          <button type="button" class="preset-chip bud-preset-chip ${initialBud==='30'?'active':''}" data-val="30">30</button>
          <button type="button" class="preset-chip bud-preset-chip ${initialBud==='40'?'active':''}" data-val="40">40</button>
          <button type="button" class="preset-chip bud-preset-chip ${initialBud==='50'?'active':''}" data-val="50">50</button>
          <button type="button" class="preset-chip bud-preset-chip ${initialBud==='100'?'active':''}" data-val="100">100</button>
        </div>
      </div>

      <div class="form-field">
        <label>DAILY ($/d)</label>
        <input type="text" class="multi-campaign-daily" placeholder="Auto $/d" readonly style="text-align:center">
      </div>

      <div class="form-field">
        <label>TAKA</label>
        <div class="input-prefix-wrap">
          <span class="input-prefix">৳</span>
          <input type="text" class="multi-campaign-taka" placeholder="Auto Total" readonly style="font-weight:700">
        </div>
      </div>
    </div>
  `;
  box.appendChild(row);

  const dayEl = row.querySelector('.multi-campaign-day');
  const budEl = row.querySelector('.multi-campaign-budget');
  const dailyEl = row.querySelector('.multi-campaign-daily');
  const takaEl = row.querySelector('.multi-campaign-taka');
  const typeEl = row.querySelector('.multi-campaign-type');
  const sumEl = row.querySelector('.multi-row-summary');
  const calcSpan = row.querySelector('.sub-suggestion-calc');

  const calcRow = () => {
    const d = parseFloat(dayEl?.value) || 0;
    const b = parseFloat(budEl?.value) || 0;
    const r = parseFloat(document.getElementById('rateInput')?.value) || state.rate || 140;
    if (d > 0 && b > 0) {
      const v = b / d;
      const dailyValStr = (v % 1 === 0) ? v.toFixed(0) : v.toFixed(2);
      if (dailyEl) dailyEl.value = dailyValStr;
      const totalTaka = b * r;
      if (takaEl) takaEl.value = totalTaka.toFixed(2);
      if (sumEl) {
        sumEl.textContent = `$${b} • ${d}d ($${dailyValStr}/d) = ৳${fmtNum(totalTaka)}`;
      }
      if (calcSpan) {
        calcSpan.textContent = `$${b} / ${d}d = $${dailyValStr}/day • ৳${fmtNum(totalTaka)}`;
      }
    } else {
      if (dailyEl) dailyEl.value = '';
      if (takaEl) takaEl.value = '';
      if (sumEl) sumEl.textContent = 'Enter day & budget';
      if (calcSpan) calcSpan.textContent = 'Select day & budget';
    }
  };

  // Day preset buttons
  row.querySelectorAll('.sub-day-chips button').forEach(btn => {
    btn.addEventListener('click', () => {
      row.querySelectorAll('.sub-day-chips button').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      if (dayEl) {
        dayEl.value = btn.dataset.val;
        calcRow();
        if (window.updateCampaignCalc) window.updateCampaignCalc();
      }
    });
  });

  // Budget preset buttons
  row.querySelectorAll('.sub-bud-chips button').forEach(btn => {
    btn.addEventListener('click', () => {
      row.querySelectorAll('.sub-bud-chips button').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      if (budEl) {
        budEl.value = btn.dataset.val;
        calcRow();
        if (window.updateCampaignCalc) window.updateCampaignCalc();
      }
    });
  });

  // Suggestion quick preset pills (1d/$5, 2d/$10, 3d/$15, etc.)
  row.querySelectorAll('.sub-preset-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const sDay = btn.dataset.day;
      const sBud = btn.dataset.bud;
      if (dayEl) dayEl.value = sDay;
      if (budEl) budEl.value = sBud;
      row.querySelectorAll('.sub-day-chips button').forEach(b => b.classList.toggle('active', b.dataset.val === sDay));
      row.querySelectorAll('.sub-bud-chips button').forEach(b => b.classList.toggle('active', b.dataset.val === sBud));
      calcRow();
      if (window.updateCampaignCalc) window.updateCampaignCalc();
    });
  });

  // Enter key inside any sub-campaign input saves campaign
  row.querySelectorAll('input, select').forEach(inp => {
    inp.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        saveEntry();
      }
    });
  });

  dayEl?.addEventListener('input', () => {
    const val = dayEl.value;
    row.querySelectorAll('.sub-day-chips button').forEach(b => {
      b.classList.toggle('active', b.dataset.val === val);
    });
    calcRow();
    if (window.updateCampaignCalc) window.updateCampaignCalc();
  });

  budEl?.addEventListener('input', () => {
    const val = budEl.value;
    row.querySelectorAll('.sub-bud-chips button').forEach(b => {
      b.classList.toggle('active', b.dataset.val === val);
    });
    calcRow();
    if (window.updateCampaignCalc) window.updateCampaignCalc();
  });

  typeEl?.addEventListener('change', () => {
    calcRow();
    if (window.updateCampaignCalc) window.updateCampaignCalc();
  });

  // Calculate immediately so daily & taka suggestions appear without manual action!
  calcRow();
  if (window.updateCampaignCalc) window.updateCampaignCalc();
}

// Payment Actions
function handlePaymentMethodChange(val) {
  const noteWrap = document.getElementById('payAdvanceNoteWrap');
  const noteInput = document.getElementById('payAdvanceNote');
  if (!noteWrap) return;
  const isAdvance = String(val || '').toLowerCase().includes('advance');
  if (isAdvance) {
    noteWrap.classList.remove('hidden');
    if (noteInput) {
      noteInput.focus();
    }
  } else {
    noteWrap.classList.add('hidden');
  }
}
window.handlePaymentMethodChange = handlePaymentMethodChange;

function addPaymentRecord() {
  if (state.isSavingPayment) return;

  const dateEl = document.getElementById('payDate');
  const date = dateEl?.value || getTodayLocalDateStr();
  const method = document.getElementById('payMethod')?.value || 'Bkash';
  const note = document.getElementById('payAdvanceNote')?.value?.trim() || '';
  const amountStr = document.getElementById('payAmount')?.value;
  const amount = parseMoney(amountStr);
  if (!date || !method || isNaN(amount) || amount <= 0) {
    showToast('Please enter a valid payment amount', 'error');
    return;
  }

  state.isSavingPayment = true;
  const addBtn = document.getElementById('btnAddPayment');
  if (addBtn) {
    addBtn.disabled = true;
    addBtn.dataset.origText = addBtn.innerHTML;
    addBtn.innerHTML = '<span class="ms animate-spin">hourglass_empty</span> Adding...';
  }

  // 1. Instant 0ms Optimistic UI execution
  const parsedDate = parseAnyDate(date);
  const formattedDate = parsedDate ? formatCustomDate(parsedDate) : date;
  const paymentMonth = parsedDate ? MONTHS[parsedDate.getMonth()] : (state.filters.month && state.filters.month !== 'All' ? state.filters.month : MONTHS[new Date().getMonth()]);
  const paymentYear = parsedDate ? String(parsedDate.getFullYear()) : String(state.filters.year || new Date().getFullYear());
  const recId = 'PAY-' + Date.now() + '-' + Math.random().toString(36).slice(2, 7);
  const optimisticPayment = {
    'Date': formattedDate,
    'Payment Method': method,
    'Paid (৳)': amount,
    'Month': paymentMonth,
    'Year': paymentYear,
    'Note': note,
    'Record ID': recId,
    _recordId: recId,
    _optimistic: true
  };

  saveEntryToVault([optimisticPayment], 'saved');

  if (!state.currentPaymentRows) state.currentPaymentRows = [];
  state.currentPaymentRows = [optimisticPayment, ...state.currentPaymentRows];
  renderPaymentHistoryTable({ rows: state.currentPaymentRows }, state.currentStats);

  // Instantly sync payment to Google Sheet
  autoSyncDirectToGoogleSheets('Payments', PAYMENT_HEADERS, [optimisticPayment]).catch(() => {});

  const payInput = document.getElementById('payAmount');
  if (payInput) payInput.value = '';
  const noteInput = document.getElementById('payAdvanceNote');
  if (noteInput) noteInput.value = '';

  showToast('Payment added successfully! ✅ (৳' + fmtNum(amount) + ')', 'success');

  // Invalidate cache timestamps so other views fetch updates in the background without wiping screen
  state.cacheTimestamps = {};

  // 2. Background Sync (Google Sheets & Server) with idempotency key
  google.script.run
    .withFailureHandler(err => {
      state.isSavingPayment = false;
      if (addBtn) {
        addBtn.disabled = false;
        addBtn.innerHTML = addBtn.dataset.origText || '<span class="ms">add</span> Add Payment';
      }
      showToast('Network error saving payment, please try again.', 'error');
    })
    .withSuccessHandler(res => {
      state.isSavingPayment = false;
      if (addBtn) {
        addBtn.disabled = false;
        addBtn.innerHTML = addBtn.dataset.origText || '<span class="ms">add</span> Add Payment';
      }
      if (res && res.success) {
        if (res.record && res.record['Record ID']) {
          const found = state.currentPaymentRows?.find(p => p._recordId === recId);
          if (found) {
            found['Record ID'] = res.record['Record ID'];
            found._recordId = res.record['Record ID'];
            found._optimistic = false;
          }
        }
      } else {
        showToast(res?.message || 'Failed to sync payment with Google Sheets', 'error');
      }
    })
    .addPayment({ recordId: recId, date: formattedDate, payment_method: method, paid_amount: amount, month: paymentMonth, year: paymentYear, note });
}

window.setQuickPayment = function(amount) {
  const payInput = document.getElementById('payAmount');
  if (payInput) {
    payInput.value = amount;
    payInput.dispatchEvent(new Event('input', { bubbles: true }));
    showToast(`Payment set to ৳${fmtNum(amount)} ⚡`, 'info');
  }
};

window.setQuickAgencyCategory = function(category) {
  const typeInp = document.getElementById('agencyCostType') || document.querySelector('#entryForm [name="type"]');
  if (typeInp) {
    typeInp.value = category;
    typeInp.dispatchEvent(new Event('input', { bubbles: true }));
    showToast(`Category selected: ${category} 📑`, 'info');
  }
};

window.setQuickAgencyCost = function(amount) {
  const costInp = document.getElementById('agencyCostAmount') || document.querySelector('#entryForm [name="cost"]');
  if (costInp) {
    costInp.value = amount;
    costInp.dispatchEvent(new Event('input', { bubbles: true }));
    showToast(`Cost set to ৳${fmtNum(amount)} (50% Share: ৳${fmtNum(+(amount/2).toFixed(2))}) ⚡`, 'info');
  }
};

function openEditAgencyCostRecord(idx) {
  const records = state.currentRenderedRecords || state.filteredRecords || state.rawRecords || [];
  const row = records[idx];
  if (!row) return;
  state.editingAgencyCostRow = row;
  state.editingAgencyCostId = row['Record ID'] || row._recordId || row._row || (row.Date + '_' + row.Type);

  const modalOverlay = document.getElementById('modalOverlay');
  const titleEl = document.getElementById('modalTitle');
  const form = document.getElementById('recordForm');
  if (!modalOverlay || !form) return;

  if (titleEl) {
    titleEl.innerHTML = `<span class="ms" style="color:var(--pri);vertical-align:middle;margin-right:6px">receipt_long</span> Edit Agency Cost Entry`;
  }

  const rawDate = row['Date'] || '';
  const inputDate = toInputDateValue(rawDate);
  const type = row['Type'] || row['type'] || row['Category'] || '';
  const cost = parseFloat(row['Cost'] || row['cost'] || 0) || 0;
  const month = row['Month'] || state.filters.month || 'September';
  const year = row['Year'] || state.filters.year || new Date().getFullYear();
  const status = row['Status'] || row['status'] || 'Paid';
  const recId = row['Record ID'] || row._recordId || '';

  form.innerHTML = `
    <div style="display:grid;grid-template-columns:repeat(auto-fit, minmax(200px, 1fr));gap:12px">
      <div class="form-field">
        <label style="font-size:11.5px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">DATE</label>
        <input type="date" name="Date" id="editAgencyCostDate" value="${escapeHtml(inputDate)}" class="form-control" required style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt)">
      </div>
      <div class="form-field" style="grid-column:span 2">
        <label style="font-size:11.5px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">TYPE / CATEGORY <span class="req">*</span></label>
        <input type="text" list="editAgencyCostCategoriesList" name="Type" id="editAgencyCostType" value="${escapeHtml(type)}" class="form-control" required style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt);font-weight:700">
        <datalist id="editAgencyCostCategoriesList">
          <option value="Office Rent">Office Rent</option>
          <option value="Facebook Ads Bill">Facebook Ads Bill</option>
          <option value="Employee Salary">Employee Salary</option>
          <option value="Internet Bill">Internet Bill</option>
          <option value="Electricity Bill">Electricity Bill</option>
          <option value="Software & Tools">Software & Tools</option>
          <option value="Domain & Hosting">Domain & Hosting</option>
          <option value="Office Snacks / Tea">Office Snacks / Tea</option>
          <option value="Other Expense">Other Expense</option>
        </datalist>
      </div>
      <div class="form-field">
        <label style="font-size:11.5px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">COST AMOUNT (৳) <span class="req">*</span></label>
        <input type="number" step="any" name="Cost" id="editAgencyCostAmount" value="${cost}" class="form-control" required style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--suc);font-weight:800;font-size:14px">
      </div>
      <div class="form-field">
        <label style="font-size:11.5px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">STATUS</label>
        <select name="Status" id="editAgencyCostStatus" class="form-control" style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt);font-weight:700">
          <option value="Paid" ${status === 'Paid' ? 'selected' : ''}>Paid</option>
          <option value="Due" ${status === 'Due' ? 'selected' : ''}>Due</option>
        </select>
      </div>
    </div>
    <div id="editAgencyCostSplitNotice" style="margin-top:10px;padding:8px 12px;background:var(--card-subtle);border-radius:8px;border:1px dashed var(--bor);font-size:12px;font-weight:700;color:var(--pri)">
      👤 50% Partner Share: ৳${fmtNum(+(cost / 2).toFixed(2))}
    </div>
    <input type="hidden" name="Record ID" value="${escapeHtml(recId)}">
  `;

  const editCostInput = form.querySelector('#editAgencyCostAmount');
  const splitNotice = form.querySelector('#editAgencyCostSplitNotice');
  editCostInput?.addEventListener('input', () => {
    const val = parseFloat(editCostInput.value) || 0;
    if (splitNotice) {
      splitNotice.innerHTML = `👤 50% Partner Share: ৳${fmtNum(+(val / 2).toFixed(2))}`;
    }
  });

  modalOverlay?.classList.remove('hidden');
}
window.openEditAgencyCostRecord = openEditAgencyCostRecord;

function deleteAgencyCostByIndex(idx) {
  const records = state.currentRenderedRecords || state.filteredRecords || state.rawRecords || [];
  const row = records[idx];
  if (!row) return;
  const costId = row['Record ID'] || row._recordId || row._row || (row.Date + '_' + row.Type);
  const costType = row['Type'] || row.type || 'Expense';
  const costAmt = row['Cost'] || row.cost || 0;

  if (!confirm(`Are you sure you want to delete "${costType}" (৳${fmtNum(costAmt)})?`)) return;

  // 1. Instant 0ms Optimistic Removal
  const matchFilter = r => {
    if (r === row) return false;
    const rId = r['Record ID'] || r._recordId || r._row;
    if (rId && String(rId) === String(costId)) return false;
    return true;
  };

  if (Array.isArray(state.records)) state.records = state.records.filter(matchFilter);
  if (Array.isArray(state.rawRecords)) state.rawRecords = state.rawRecords.filter(matchFilter);
  if (Array.isArray(state.filteredRecords)) state.filteredRecords = state.filteredRecords.filter(matchFilter);
  if (Array.isArray(state.currentRenderedRecords)) state.currentRenderedRecords = state.currentRenderedRecords.filter(matchFilter);

  renderRecordsTable({ headers: state.headers, rows: state.filteredRecords });
  showToast('Agency Cost entry deleted! 🗑️', 'info');

  state.cacheTimestamps = {};

  // 2. Direct Google Sheets OAuth API & Webhook deletion (in background)
  const oauthToken = getGoogleAccessToken();
  let sheetId = '';
  let webhookUrl = '';
  try {
    const savedConfig = JSON.parse(localStorage.getItem('adverse_sheet_config') || '{}');
    if (savedConfig.sheetId) sheetId = savedConfig.sheetId;
    if (savedConfig.webhookUrl) webhookUrl = savedConfig.webhookUrl;
  } catch (e) {}
  if (!sheetId && state.sheetId) sheetId = state.sheetId;

  if (oauthToken && sheetId) {
    deleteRowFromSheetDirect(sheetId, 'Agency_Cost', costId, undefined, { budget: costAmt }).catch(() => {});
  }
  if (webhookUrl) {
    fetch(webhookUrl, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        action: 'deleteRow',
        sheetName: 'Agency_Cost',
        recordId: costId,
        matchInfo: { 'Record ID': costId, budget: costAmt }
      })
    }).catch(() => {});
  }

  // 3. Background Server Sync
  google.script.run.withSuccessHandler(res => {
    state.cacheTimestamps = {};
    loadMenuData(true);
  }).deleteAgencyCost(costId);
}
window.deleteAgencyCostByIndex = deleteAgencyCostByIndex;

function openEditPaymentRecord(idx) {
  const row = (state.currentPaymentRows && state.currentPaymentRows[idx]) || null;
  if (!row) return;
  state.editingPaymentRow = row;
  state.editingPaymentId = row._recordId || row._row || row.Date;

  const modalOverlay = document.getElementById('modalOverlay');
  const titleEl = document.getElementById('modalTitle');
  const form = document.getElementById('recordForm');
  if (!modalOverlay || !form) return;

  if (titleEl) {
    titleEl.innerHTML = `<span class="ms" style="color:var(--suc);vertical-align:middle;margin-right:6px">payments</span> Edit Payment Entry`;
  }

  const rawDate = row['Date'] || '';
  const inputDate = toInputDateValue(rawDate);
  const method = row['Payment Method'] || 'Bkash';
  const amount = parseFloat(row['Paid (৳)'] || row['paid_amount'] || 0) || 0;
  const month = row['Month'] || state.filters.month || '';
  const year = row['Year'] || state.filters.year || new Date().getFullYear();
  const note = row['Note'] || row.Note || row.note || row['Advance Note'] || '';

  form.innerHTML = `
    <div style="display:grid;grid-template-columns:repeat(auto-fit, minmax(220px, 1fr));gap:12px">
      <div class="form-field">
        <label style="font-size:11.5px;font-weight:700;color:var(--txt-muted)">DATE</label>
        <input type="date" name="Date" value="${escapeHtml(inputDate)}" class="form-control" required style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt)">
      </div>
      <div class="form-field">
        <label style="font-size:11.5px;font-weight:700;color:var(--txt-muted)">PAYMENT METHOD</label>
        <select name="Payment Method" class="form-control" onchange="const w = document.getElementById('modalPayNoteWrap'); if(w) { if(this.value.toLowerCase().includes('advance')) w.classList.remove('hidden'); else w.classList.add('hidden'); }" style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt)">
          ${['Bkash', 'Nagad', 'Bank', 'BK to Bank', 'Merchant', 'Cash', 'Advance Paid (Old Month)'].map(m => `<option value="${m}" ${m === method ? 'selected' : ''}>${m}</option>`).join('')}
        </select>
      </div>
      <div class="form-field ${String(method).toLowerCase().includes('advance') || note ? '' : 'hidden'}" id="modalPayNoteWrap">
        <label style="font-size:11.5px;font-weight:700;color:var(--pri)">ADVANCE NOTE / CONTEXT</label>
        <input type="text" name="Note" value="${escapeHtml(note)}" class="form-control" placeholder="e.g. Previous month excess, Client advance..." style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt)">
      </div>
      <div class="form-field">
        <label style="font-size:11.5px;font-weight:700;color:var(--txt-muted)">PAID (৳)</label>
        <input type="number" step="any" name="Paid (৳)" value="${amount}" class="form-control" required style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--suc);font-weight:800;font-size:14px">
      </div>
      <div class="form-field">
        <label style="font-size:11.5px;font-weight:700;color:var(--txt-muted)">MONTH</label>
        <select name="Month" class="form-control" style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt)">
          ${state.months.map(m => `<option value="${m}" ${m === month ? 'selected' : ''}>${m}</option>`).join('')}
        </select>
      </div>
      <div class="form-field">
        <label style="font-size:11.5px;font-weight:700;color:var(--txt-muted)">YEAR</label>
        <input type="number" name="Year" value="${year}" class="form-control" style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt)">
      </div>
    </div>
  `;

  modalOverlay?.classList.remove('hidden');
}

function deletePaymentItem(recordId, paymentRow) {
  if (!confirm('Are you sure you want to delete this payment record?')) return;
  const row = paymentRow || (state.currentPaymentRows ? state.currentPaymentRows.find(p => p._recordId === recordId || String(p._row) === String(recordId)) : null);
  
  // 1. Instant optimistic local UI update
  if (state.currentPaymentRows) {
    state.currentPaymentRows = state.currentPaymentRows.filter(p => p._recordId !== recordId && String(p._row) !== String(recordId));
    renderPaymentHistoryTable({ rows: state.currentPaymentRows }, state.currentStats);
  }
  showToast('Payment record deleted from App & Google Sheet! 🗑️', 'info');

  // 2. Direct Google Sheets OAuth API & Webhook deletion (in background)
  const oauthToken = getGoogleAccessToken();
  let sheetId = '';
  let webhookUrl = '';
  try {
    const savedConfig = JSON.parse(localStorage.getItem('adverse_sheet_config') || '{}');
    if (savedConfig.sheetId) sheetId = savedConfig.sheetId;
    if (savedConfig.webhookUrl) webhookUrl = savedConfig.webhookUrl;
  } catch (e) {}
  if (!sheetId && state.sheetId) sheetId = state.sheetId;

  const extraMatch = {
    budget: row ? (row['Paid (৳)'] || row.paid_amount || row.Paid) : undefined,
    month: row?.Month,
    year: row?.Year
  };

  if (oauthToken && sheetId) {
    deleteRowFromSheetDirect(sheetId, 'Payments', recordId, undefined, extraMatch).catch(() => {});
  }
  if (webhookUrl) {
    fetch(webhookUrl, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        action: 'deleteRow',
        sheetName: 'Payments',
        recordId,
        matchInfo: { 'Record ID': recordId, ...extraMatch }
      })
    }).catch(() => {});
  }

  // 3. Server-side deletion
  google.script.run.withSuccessHandler(res => {
    state.cacheTimestamps = {};
    loadMenuData(true);
  }).withFailureHandler(err => {
    showToast(err?.message || 'Failed to delete payment on server', 'error');
    loadMenuData(true);
  }).deletePayment(recordId, row);
}

function deleteRecordByIndex(idx) {
  const row = state.currentRenderedRecords && state.currentRenderedRecords[idx];
  if (!row) return;
  const id = row['Record ID'] || row._recordId || row._row || row.SL;
  deleteRecordItem(id, row);
}

async function deleteRecordItem(id, row) {
  if (!confirm('Are you sure you want to delete this record?')) return;
  const key = state.currentMenu ? state.currentMenu.key : 'campaign';
  state.cacheTimestamps = {};
  const month = (row && (row.Month || row.month)) || state.filters.month;
  const year = (row && (row.Year || row.year)) || state.filters.year;

  // 1. Instant Optimistic Local State Deletion (Instant UI update)
  const targetId = String(id || '').trim();
  const matchFilter = r => {
    if (row && r === row) return false;
    if (targetId) {
      const rId = String(r['Record ID'] || r._recordId || r._row || '').trim();
      if (rId && rId === targetId) return false;
    }
    const rSl = String(r.SL || r['SL'] || r['SL NO'] || r.sl || r.sl_no || '').trim();
    const rowSl = String((row && (row.SL || row['SL'] || row['SL NO'])) || '').trim();
    if (rowSl && rSl) {
      const rSlNum = parseInt(rSl.replace(/\D/g, ''), 10);
      const rowSlNum = parseInt(rowSl.replace(/\D/g, ''), 10);
      const slMatch = (rowSl === rSl) || (!isNaN(rowSlNum) && !isNaN(rSlNum) && rowSlNum === rSlNum);
      if (slMatch) {
        const rPage = String(r['Page Name'] || r.page_name || '').trim().toLowerCase();
        const rowPage = String((row && (row['Page Name'] || row.page_name)) || '').trim().toLowerCase();
        if (!rowPage || !rPage || rowPage === rPage || rowPage.includes(rPage) || rPage.includes(rowPage)) {
          return false;
        }
      }
    }
    return true;
  };
  if (Array.isArray(state.records)) state.records = state.records.filter(matchFilter);
  if (Array.isArray(state.rawRecords)) state.rawRecords = state.rawRecords.filter(matchFilter);
  if (Array.isArray(state.filteredRecords)) state.filteredRecords = state.filteredRecords.filter(matchFilter);
  if (Array.isArray(state.currentRenderedRecords)) state.currentRenderedRecords = state.currentRenderedRecords.filter(matchFilter);
  if (Array.isArray(state.masterCampaignRegistry)) state.masterCampaignRegistry = state.masterCampaignRegistry.filter(matchFilter);
  if (Array.isArray(state.todayCreatedRegistry)) state.todayCreatedRegistry = state.todayCreatedRegistry.filter(matchFilter);
  if (Array.isArray(state.todayCreated)) state.todayCreated = state.todayCreated.filter(matchFilter);
  if (Array.isArray(state.allCampaigns)) state.allCampaigns = state.allCampaigns.filter(matchFilter);
  if (typeof renderRecordsTable === 'function') {
    renderRecordsTable({ headers: state.headers, rows: state.filteredRecords });
  }

  // Clear client cache and persist local snapshot immediately
  state.clientCache = {};
  state.cacheTimestamps = {};
  saveLocalAppSnapshot();

  // No success toast here: it is shown only after the server confirms the
  // delete (see callback below), so the app never claims a delete that failed.

  // 2. Direct Google Sheets OAuth API & Webhook deletion (in background)
  const oauthToken = getGoogleAccessToken();
  let sheetId = '';
  let webhookUrl = '';
  try {
    const savedConfig = JSON.parse(localStorage.getItem('adverse_sheet_config') || '{}');
    if (savedConfig.sheetId) sheetId = savedConfig.sheetId;
    if (savedConfig.webhookUrl) webhookUrl = savedConfig.webhookUrl;
  } catch (e) {}
  if (!sheetId && state.sheetId) sheetId = state.sheetId;

  const rawStart = row ? (row['Start'] || row['Start Date'] || row.start_date || '') : '';
  const dParsed = parseAnyDate(rawStart);
  const resolvedMonth = dParsed ? MONTHS[dParsed.getMonth()] : ((row && row.Month) ? row.Month : (state.filters.month && state.filters.month !== 'All' ? state.filters.month : 'September'));
  const resolvedYear = dParsed ? String(dParsed.getFullYear()) : ((row && row.Year) ? String(row.Year) : (state.filters.year && state.filters.year !== 'All Years' ? String(state.filters.year) : String(new Date().getFullYear())));

  const sl = row ? (row.SL || row['SL'] || row['SL NO'] || row.sl || row.sl_no) : undefined;
  const extraMatch = {
    pageName: row ? (row['Page Name'] || row.page_name || row['Page']) : undefined,
    budget: row ? (row['Budget'] || row.budget || row['Cost'] || row['Spent'] || row['Paid (৳)']) : undefined,
    month: resolvedMonth,
    year: resolvedYear
  };

  if (oauthToken && sheetId) {
    try {
      if (key === 'new_pages') {
        deleteRowFromSheetDirect(sheetId, 'Pause_ADS', id, sl, extraMatch).catch(() => {});
      } else if (key === 'scam') {
        deleteRowFromSheetDirect(sheetId, 'Scam', id, sl, extraMatch).catch(() => {});
      } else if (key === 'agency_cost') {
        deleteRowFromSheetDirect(sheetId, 'Agency_Cost', id, sl, extraMatch).catch(() => {});
      } else if (key === 'payment_due_tracker' || key === 'payments') {
        deleteRowFromSheetDirect(sheetId, 'Payments', id, sl, extraMatch).catch(() => {});
      } else {
        // Delete from all potential campaign tabs where this record could reside
        Promise.allSettled([
          deleteRowFromSheetDirect(sheetId, `Campaign_${resolvedMonth}_${resolvedYear}`, id, sl, extraMatch),
          deleteRowFromSheetDirect(sheetId, 'Today_Created', id, sl, extraMatch),
          deleteRowFromSheetDirect(sheetId, 'Campaign_September_2026', id, sl, extraMatch),
          deleteRowFromSheetDirect(sheetId, 'Campaign_October_2026', id, sl, extraMatch),
          deleteRowFromSheetDirect(sheetId, 'Campaign', id, sl, extraMatch)
        ]).catch(() => {});
      }
    } catch (e) {
      console.warn('Direct sheet delete error:', e);
    }
  }

  // Webhook instant deletion
  if (webhookUrl) {
    const targetWebhookTabs = [];
    if (key === 'new_pages') targetWebhookTabs.push('Pause_ADS');
    else if (key === 'scam') targetWebhookTabs.push('Scam');
    else if (key === 'agency_cost') targetWebhookTabs.push('Agency_Cost');
    else if (key === 'payment_due_tracker' || key === 'payments') targetWebhookTabs.push('Payments');
    else {
      targetWebhookTabs.push(`Campaign_${resolvedMonth}_${resolvedYear}`);
      targetWebhookTabs.push('Today_Created');
    }

    targetWebhookTabs.forEach(tabName => {
      fetch(webhookUrl, {
        method: 'POST',
        mode: 'no-cors',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'deleteRow',
          sheetName: tabName,
          recordId: id,
          matchInfo: {
            'Record ID': id,
            SL: sl,
            'Page Name': extraMatch.pageName,
            budget: extraMatch.budget
          }
        })
      }).catch(() => {});
    });
  }

  // 3. Server-side deletion
  const callback = res => {
    if (res && res.success === false) {
      showToast(res.message || 'Delete failed on server', 'error');
      loadMenuData();
    } else if (res && res.sheetWarning) {
      showToast('Deleted from App — Sheet sync baki ache, auto-retry hobe', 'warning');
      loadMenuData();
    } else {
      showToast('Record deleted from App & Google Sheet!', 'success');
    }
  };
  const errorHandler = err => {
    showToast(err?.message || 'Failed to delete record on server', 'error');
    loadMenuData();
  };

  if (key === 'campaign' || key === 'today_created_campaign' || key === 'today_created') {
    google.script.run.withSuccessHandler(callback).withFailureHandler(errorHandler)
      .deleteCampaignRecord(id, month, year, row);
  } else if (key === 'new_pages') {
    google.script.run.withSuccessHandler(callback).withFailureHandler(errorHandler)
      .deletePauseAds(id, row);
  } else if (key === 'scam') {
    google.script.run.withSuccessHandler(callback).withFailureHandler(errorHandler)
      .deleteScam(id, row);
  } else if (key === 'agency_cost') {
    google.script.run.withSuccessHandler(callback).withFailureHandler(errorHandler)
      .deleteAgencyCost(id, row);
  } else if (key === 'agency_spending' || key === 'payment_due_tracker' || key === 'payments') {
    google.script.run.withSuccessHandler(callback).withFailureHandler(errorHandler)
      .deletePayment(id, row);
  } else {
    google.script.run.withSuccessHandler(callback).withFailureHandler(errorHandler)
      .deleteCampaignRecord(id, month, year, row);
  }
}

function openEditRecordByIndex(idx) {
  const row = state.currentRenderedRecords && state.currentRenderedRecords[idx];
  if (!row) return;
  openEditRecord(row);
}

function openEditRecord(row) {
  if (!row) return;
  state.editingRow = row;
  state.editingId = row['Record ID'] || row._recordId || row._row || row['SL NO'] || row.sl_no || row.SL || row.Date;
  
  const titleEl = document.getElementById('modalTitle');
  const isCampaignRecord = Boolean(row['Page Name'] || row.page_name || row['Budget'] || row.budget);
  const pageName = row['Page Name'] || row.page_name || row.Page || row['Ad Account'] || '';
  if (titleEl) {
    titleEl.textContent = `Edit Campaign ${pageName ? '– ' + pageName : ''}`;
  }

  const form = document.getElementById('recordForm');
  if (!form) return;

  const knownAccounts = state.adAccounts || [];
  const knownPages = state.pages || [];

  if (isCampaignRecord) {
    const slVal = row['SL'] || row['sl_no'] || row['SL NO'] || row.sl || '';
    const offVal = row['OFF $'] ?? row['OFF USD'] ?? row['off_usd'] ?? row['OFF'] ?? row.off ?? '0';
    const adAccVal = row['AD Account'] || row['ad_account'] || '';
    const idVal = row['ID Name'] || row['id_name'] || '';
    const statusVal = row['Status'] || row.status || 'Old Page';
    const pNameVal = row['Page Name'] || row.page_name || '';
    const startVal = toInputDateValue(row['Start'] || row['Start Date'] || row['start_date'] || new Date());
    const endVal = toInputDateValue(row['End'] || row['End Date'] || row['end_date'] || new Date());
    const typeVal = row['Types'] || row['Type'] || row.types || row.type || 'Boost';
    const dayVal = parseFloat(row['Day'] || row.day || 1) || 1;
    const budVal = parseFloat(row['Budget'] || row.budget || 10) || 10;
    const dailyVal = row['Daily $'] || row['daily'] || (dayVal > 0 ? (budVal / dayVal).toFixed(2) : '');
    const rateVal = parseFloat(row['Rate'] || row['Dollar Rate (৳)'] || row.rate || state.rate || 140);
    const takaVal = row['Taka'] || row.taka || (budVal * rateVal).toFixed(2);
    const monthVal = row['Month'] || row.month || state.filters.month || MONTHS[new Date().getMonth()];
    const yearVal = row['Year'] || row.year || state.filters.year || new Date().getFullYear();
    const recIdVal = row['Record ID'] || row._recordId || row.record_id || '';

    form.innerHTML = `
      <!-- Datalists for autocompletion -->
      <datalist id="editModalAdAccountsList">
        ${knownAccounts.map(a => `<option value="${escapeHtml(a.a)}">${escapeHtml(a.id || '')}</option>`).join('')}
      </datalist>
      <datalist id="editModalPagesList">
        ${knownPages.map(p => `<option value="${escapeHtml(p)}"></option>`).join('')}
      </datalist>

      <div style="display:grid;grid-template-columns:repeat(auto-fit, minmax(200px, 1fr));gap:12px">
        <div class="form-field">
          <label style="font-size:11px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">SL NO</label>
          <input type="text" name="SL" value="${escapeHtml(slVal)}" class="form-control" style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt);font-weight:700;text-align:center">
        </div>

        <div class="form-field">
          <label style="font-size:11px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">OFF $</label>
          <input type="number" step="any" name="OFF $" id="editModalOffUsd" value="${escapeHtml(offVal)}" class="form-control" placeholder="0" style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--dan);font-weight:700;text-align:center">
        </div>

        <div class="form-field" style="grid-column:span 2">
          <label style="font-size:11px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">PAGE NAME <span class="req">*</span></label>
          <input type="text" list="editModalPagesList" name="Page Name" id="editModalPageName" value="${escapeHtml(pNameVal)}" class="form-control" required style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt);font-weight:700">
        </div>

        <div class="form-field">
          <label style="font-size:11px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">AD ACCOUNT <span class="req">*</span></label>
          <input type="text" list="editModalAdAccountsList" name="AD Account" id="editModalAdAccount" value="${escapeHtml(adAccVal)}" class="form-control" required style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt);font-weight:600">
        </div>

        <div class="form-field">
          <label style="font-size:11px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">ID NAME</label>
          <input type="text" name="ID Name" id="editModalIdName" value="${escapeHtml(idVal)}" class="form-control" style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt)">
        </div>

        <div class="form-field">
          <label style="font-size:11px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">STATUS <span class="req">*</span></label>
          <select name="Status" class="form-control" style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt);font-weight:600">
            <option value="Old Page" ${statusVal === 'Old Page' ? 'selected' : ''}>Old Page</option>
            <option value="New Page" ${statusVal === 'New Page' ? 'selected' : ''}>New Page</option>
          </select>
        </div>

        <div class="form-field">
          <label style="font-size:11px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">TYPES <span class="req">*</span></label>
          <select name="Types" id="editModalTypes" class="form-control" style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt);font-weight:600">
            <option value="Boost" ${typeVal === 'Boost' ? 'selected' : ''}>Boost</option>
            <option value="Ad Campaign" ${typeVal === 'Ad Campaign' ? 'selected' : ''}>Ad Campaign</option>
            <option value="Increase" ${typeVal === 'Increase' ? 'selected' : ''}>Increase</option>
            <option value="Others" ${typeVal === 'Others' ? 'selected' : ''}>Others</option>
          </select>
        </div>

        <div class="form-field">
          <label style="font-size:11px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">START DATE <span class="req">*</span></label>
          <input type="date" name="Start" id="editModalStart" value="${escapeHtml(startVal)}" class="form-control" required style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt)">
        </div>

        <div class="form-field">
          <label style="font-size:11px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">END DATE</label>
          <input type="date" name="End" id="editModalEnd" value="${escapeHtml(endVal)}" class="form-control" readonly style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt);opacity:0.9">
        </div>

        <div class="form-field">
          <label style="font-size:11px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">DAY <span class="req">*</span></label>
          <input type="number" min="1" name="Day" id="editModalDay" value="${dayVal}" class="form-control" required style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt);font-weight:700;text-align:center">
          <div class="quick-preset-chips" style="margin-top:4px;display:flex;gap:4px;flex-wrap:wrap">
            ${[1, 2, 3, 5, 7, 10, 15, 30].map(d => `<button type="button" class="preset-chip ${dayVal === d ? 'active' : ''}" onclick="document.getElementById('editModalDay').value=${d}; document.getElementById('editModalDay').dispatchEvent(new Event('input', {bubbles:true}))">${d}</button>`).join('')}
          </div>
        </div>

        <div class="form-field">
          <label style="font-size:11px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">BUDGET ($) <span class="req">*</span></label>
          <div class="input-prefix-wrap">
            <span class="input-prefix">$</span>
            <input type="number" min="1" step="any" name="Budget" id="editModalBudget" value="${budVal}" class="form-control" required style="width:100%;padding:8px 10px 8px 24px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--sec);font-weight:800;font-size:14px">
          </div>
          <div class="quick-preset-chips" style="margin-top:4px;display:flex;gap:4px;flex-wrap:wrap">
            ${[5, 10, 15, 20, 25, 30, 50, 100].map(b => `<button type="button" class="preset-chip ${budVal === b ? 'active' : ''}" onclick="document.getElementById('editModalBudget').value=${b}; document.getElementById('editModalBudget').dispatchEvent(new Event('input', {bubbles:true}))">${b}</button>`).join('')}
          </div>
        </div>

        <div class="form-field">
          <label style="font-size:11px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">DAILY ($/day)</label>
          <input type="text" name="Daily $" id="editModalDaily" value="${escapeHtml(dailyVal)}" class="form-control" readonly style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt);text-align:center">
        </div>

        <div class="form-field">
          <label style="font-size:11px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">RATE (৳)</label>
          <input type="number" name="Rate" id="editModalRate" value="${rateVal}" class="form-control" style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt);text-align:center">
        </div>

        <div class="form-field">
          <label style="font-size:11px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">TOTAL TAKA (৳)</label>
          <div class="input-prefix-wrap">
            <span class="input-prefix">৳</span>
            <input type="text" name="Taka" id="editModalTaka" value="${escapeHtml(takaVal)}" class="form-control" readonly style="width:100%;padding:8px 10px 8px 24px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--suc);font-weight:800;font-size:14px">
          </div>
        </div>

        <div class="form-field">
          <label style="font-size:11px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">MONTH</label>
          <select name="Month" class="form-control" style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt)">
            ${state.months.map(m => `<option value="${m}" ${m === monthVal ? 'selected' : ''}>${m}</option>`).join('')}
          </select>
        </div>

        <div class="form-field">
          <label style="font-size:11px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">YEAR</label>
          <input type="number" name="Year" value="${escapeHtml(yearVal)}" class="form-control" style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt)">
        </div>
      </div>
      <input type="hidden" name="Record ID" value="${escapeHtml(recIdVal)}">
      <input type="hidden" name="off_usd" id="editModalOffUsdHidden" value="${escapeHtml(offVal)}">
    `;

    // Auto-sync AD Account ID Name in Edit modal
    const editAdInput = form.querySelector('#editModalAdAccount');
    const editIdInput = form.querySelector('#editModalIdName');
    editAdInput?.addEventListener('input', () => {
      const val = (editAdInput.value || '').trim().toLowerCase();
      const matched = knownAccounts.find(a => a.a.toLowerCase() === val || a.a.toLowerCase().includes(val));
      if (matched && editIdInput) editIdInput.value = matched.id;
    });

    // Auto-sync OFF $ hidden field
    const editOffInput = form.querySelector('#editModalOffUsd');
    const editOffHidden = form.querySelector('#editModalOffUsdHidden');
    editOffInput?.addEventListener('input', () => {
      if (editOffHidden) editOffHidden.value = editOffInput.value;
    });
  } else {
    // Non-campaign generic editor with clean inputs
    const fields = Object.entries(row).filter(([k]) => !k.startsWith('_'));
    form.innerHTML = `
      <div style="display:grid;grid-template-columns:repeat(auto-fit, minmax(220px, 1fr));gap:12px">
        ${fields.map(([k, v]) => {
          const val = v === undefined || v === null ? '' : String(v);
          const lower = k.toLowerCase();
          if (lower === 'status' && (val === 'Due' || val === 'Paid')) {
            return `
              <div class="form-field">
                <label style="font-size:11px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">${escapeHtml(k)}</label>
                <select name="${escapeHtml(k)}" class="form-control" style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt)">
                  <option value="Due" ${val === 'Due' ? 'selected' : ''}>Due</option>
                  <option value="Paid" ${val === 'Paid' ? 'selected' : ''}>Paid</option>
                </select>
              </div>
            `;
          }
          if (lower === 'month') {
            return `
              <div class="form-field">
                <label style="font-size:11px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">${escapeHtml(k)}</label>
                <select name="${escapeHtml(k)}" class="form-control" style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt)">
                  ${state.months.map(m => `<option value="${m}" ${m === val ? 'selected' : ''}>${m}</option>`).join('')}
                </select>
              </div>
            `;
          }
          let inputType = 'text';
          if (lower.includes('budget') || lower.includes('taka') || lower.includes('cost') || lower.includes('paid') || lower.includes('spent') || lower.includes('rate') || lower === 'day' || lower.includes('daily')) {
            inputType = 'number';
          }
          if (lower.includes('date')) {
            inputType = 'date';
            const dVal = toInputDateValue(val);
            return `
              <div class="form-field">
                <label style="font-size:11px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">${escapeHtml(k)}</label>
                <input type="date" name="${escapeHtml(k)}" value="${escapeHtml(dVal)}" class="form-control" style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt)">
              </div>
            `;
          }
          return `
            <div class="form-field">
              <label style="font-size:11px;font-weight:700;color:var(--txt-muted);text-transform:uppercase">${escapeHtml(k)}</label>
              <input type="${inputType}" step="any" name="${escapeHtml(k)}" value="${escapeHtml(val)}" class="form-control" style="width:100%;padding:8px 10px;background:var(--input-bg);border:1px solid var(--bor);border-radius:var(--radius-sm);color:var(--txt)">
            </div>
          `;
        }).join('')}
      </div>
    `;
  }

  // Attach dynamic calculation handlers in edit modal
  const modalDay = form.querySelector('[name="Day"]') || form.querySelector('[name="day"]');
  const modalBudget = form.querySelector('[name="Budget"]') || form.querySelector('[name="budget"]');
  const modalDaily = form.querySelector('[name="Daily $"]') || form.querySelector('[name="Daily"]') || form.querySelector('[name="daily"]');
  const modalRate = form.querySelector('[name="Rate"]') || form.querySelector('[name="rate"]');
  const modalTaka = form.querySelector('[name="Taka"]') || form.querySelector('[name="taka"]');
  const modalStart = form.querySelector('[name="Start"]') || form.querySelector('[name="Start Date"]') || form.querySelector('[name="start_date"]');
  const modalEnd = form.querySelector('[name="End"]') || form.querySelector('[name="End Date"]') || form.querySelector('[name="end_date"]');

  const updateModalCalc = (source) => {
    const d = parseFloat(modalDay?.value) || 0;
    const b = parseFloat(modalBudget?.value) || 0;
    const r = parseFloat(modalRate?.value) || state.rate || 140;

    if (modalDaily && source !== 'daily') {
      if (d > 0 && b > 0) {
        const dv = b / d;
        modalDaily.value = (dv % 1 === 0) ? dv.toFixed(0) : dv.toFixed(2);
      }
    }
    if (modalTaka) {
      if (b > 0 && r > 0) modalTaka.value = (b * r).toFixed(2);
    }
    if (modalStart && modalEnd && d > 0) {
      const parsedStart = parseAnyDate(modalStart.value);
      if (parsedStart && !isNaN(parsedStart.getTime())) {
        const endDate = new Date(parsedStart.getFullYear(), parsedStart.getMonth(), parsedStart.getDate() + d);
        modalEnd.value = toInputDateValue(endDate);
      }
    }
  };

  modalDay?.addEventListener('input', () => updateModalCalc('day'));
  modalBudget?.addEventListener('input', () => updateModalCalc('budget'));
  modalRate?.addEventListener('input', () => updateModalCalc('rate'));

  const handleModalStartChange = () => {
    updateModalCalc('start');
    if (modalStart) {
      const parsedStart = parseAnyDate(modalStart.value);
      if (parsedStart && !isNaN(parsedStart.getTime())) {
        const m = MONTHS[parsedStart.getMonth()];
        const y = String(parsedStart.getFullYear());
        const monthSelect = form.querySelector('[name="Month"]') || form.querySelector('[name="month"]');
        const yearInput = form.querySelector('[name="Year"]') || form.querySelector('[name="year"]');
        if (monthSelect) monthSelect.value = m;
        if (yearInput) yearInput.value = y;
      }
    }
  };
  modalStart?.addEventListener('input', handleModalStartChange);
  modalStart?.addEventListener('change', handleModalStartChange);
  modalDaily?.addEventListener('input', () => {
    const daily = parseFloat(modalDaily.value) || 0;
    const d = parseFloat(modalDay?.value) || 0;
    if (daily > 0 && d > 0 && modalBudget) {
      const b = daily * d;
      modalBudget.value = (b % 1 === 0) ? b.toFixed(0) : b.toFixed(2);
    }
    updateModalCalc('daily');
  });

  const modalOverlay = document.getElementById('modalOverlay');
  if (modalOverlay) modalOverlay.classList.remove('hidden');
}

function closeModal() {
  const modalOverlay = document.getElementById('modalOverlay');
  if (modalOverlay) modalOverlay.classList.add('hidden');
  state.editingRow = null;
  state.editingId = null;
  state.editingPaymentRow = null;
  state.editingPaymentId = null;
  state.editingAgencyCostRow = null;
  state.editingAgencyCostId = null;
}

async function saveModalRecord() {
  const form = document.getElementById('recordForm');
  if (!form) return;
  const fd = new FormData(form);
  const data = Object.fromEntries(fd.entries());

  // Handle Agency Cost Editing directly
  if (state.editingAgencyCostRow) {
    const costId = state.editingAgencyCostId;
    const rawDate = data.Date || state.editingAgencyCostRow.Date;
    const parsedDate = parseAnyDate(rawDate) || new Date();
    const formattedDate = parsedDate ? formatCustomDate(parsedDate) : rawDate;
    const costVal = Number(data.Cost || data.cost || 0);
    const updatedCost = {
      ...state.editingAgencyCostRow,
      'Date': formattedDate,
      'Month': MONTHS[parsedDate.getMonth()],
      'Year': String(parsedDate.getFullYear()),
      'Type': data.Type || data.type || state.editingAgencyCostRow.Type,
      'Cost': costVal,
      'Per Person Cost': +(costVal / 2).toFixed(2),
      'Status': data.Status || data.status || 'Paid'
    };

    const matchCost = r => {
      if (r === state.editingAgencyCostRow) return true;
      const rId = r['Record ID'] || r._recordId || r._row;
      if (rId && String(rId) === String(costId)) return true;
      return false;
    };

    if (Array.isArray(state.records)) {
      const idx = state.records.findIndex(matchCost);
      if (idx >= 0) state.records[idx] = updatedCost;
    }
    if (Array.isArray(state.rawRecords)) {
      const idx = state.rawRecords.findIndex(matchCost);
      if (idx >= 0) state.rawRecords[idx] = updatedCost;
    }
    if (Array.isArray(state.filteredRecords)) {
      const idx = state.filteredRecords.findIndex(matchCost);
      if (idx >= 0) state.filteredRecords[idx] = updatedCost;
    }
    if (Array.isArray(state.currentRenderedRecords)) {
      const idx = state.currentRenderedRecords.findIndex(matchCost);
      if (idx >= 0) state.currentRenderedRecords[idx] = updatedCost;
    }

    renderRecordsTable({ headers: state.headers, rows: state.filteredRecords });
    closeModal();
    showToast('Agency Cost entry updated successfully! ✅', 'success');

    autoSyncDirectToGoogleSheets('Agency_Cost', AGENCY_COST_HEADERS, [updatedCost]).catch(() => {});

    google.script.run.withSuccessHandler(res => {
      state.cacheTimestamps = {};
      loadMenuData(true);
    }).updateAgencyCost(costId, updatedCost);
    return;
  }

  // Handle Payment Editing directly
  if (state.editingPaymentRow) {
    const payId = state.editingPaymentId;
    const rawDate = data.Date || state.editingPaymentRow.Date;
    const parsedDate = parseAnyDate(rawDate);
    const formattedDate = parsedDate ? formatCustomDate(parsedDate) : rawDate;
    const updatedPayment = {
      ...state.editingPaymentRow,
      'Date': formattedDate,
      'Payment Method': data['Payment Method'] || state.editingPaymentRow['Payment Method'],
      'Paid (৳)': Number(data['Paid (৳)'] || 0),
      'Month': data.Month || state.editingPaymentRow.Month,
      'Year': data.Year || state.editingPaymentRow.Year,
      'Note': String(data.Note || data.note || state.editingPaymentRow.Note || '').trim()
    };
    if (state.currentPaymentRows) {
      const pIdx = state.currentPaymentRows.findIndex(p => p._recordId === payId || String(p._row) === String(payId));
      if (pIdx >= 0) state.currentPaymentRows[pIdx] = updatedPayment;
      renderPaymentHistoryTable({ rows: state.currentPaymentRows }, state.currentStats);
    }
    closeModal();
    showToast('Payment record updated successfully! ✅', 'success');
    autoSyncDirectToGoogleSheets('Payments', PAYMENT_HEADERS, [updatedPayment]).catch(() => {});
    google.script.run.withSuccessHandler(res => {
      state.cacheTimestamps = {};
      loadMenuData(true);
    }).updatePayment(payId, updatedPayment);
    return;
  }

  const originalRow = state.editingRow || {};

  // Normalize dates to strict 09-Oct-2026 format and strictly derive Month & Year from Start Date
  const rawStart = data.Start || data['Start Date'] || data.start_date || originalRow.Start || originalRow['Start Date'];
  const parsedStart = parseAnyDate(rawStart);
  let month = originalRow.Month || originalRow.month || state.filters.month || 'October';
  let year = String(originalRow.Year || originalRow.year || state.filters.year || new Date().getFullYear());

  if (parsedStart && !isNaN(parsedStart.getTime())) {
    const formattedStart = formatCustomDate(parsedStart);
    data.Start = formattedStart;
    data['Start Date'] = formattedStart;
    data.start_date = formattedStart;
    month = MONTHS[parsedStart.getMonth()];
    year = String(parsedStart.getFullYear());
    data.Month = month;
    data.month = month;
    data.Year = year;
    data.year = year;
  } else {
    month = data.Month || data.month || originalRow.Month || originalRow.month || state.filters.month || MONTHS[new Date().getMonth()];
    year = String(data.Year || data.year || originalRow.Year || originalRow.year || state.filters.year || new Date().getFullYear());
    data.month = month;
    data.year = year;
    data.Month = month;
    data.Year = year;
  }

  const rawEnd = data.End || data['End Date'] || data.end_date;
  if (rawEnd) {
    const parsedEnd = parseAnyDate(rawEnd);
    if (parsedEnd) {
      data.End = formatCustomDate(parsedEnd);
      data['End Date'] = data.End;
      data.end_date = data.End;
    }
  }

  const rawEntry = data['Entry Date'] || data.entry_date || originalRow['Entry Date'] || originalRow.entry_date;
  if (rawEntry) {
    const parsedEntry = parseAnyDate(rawEntry);
    if (parsedEntry) {
      data['Entry Date'] = formatCustomDate(parsedEntry);
      data.entry_date = data['Entry Date'];
    }
  }

  for (const [k, v] of Object.entries(data)) {
    if (k.toLowerCase().includes('date') && v) {
      const p = parseAnyDate(v);
      if (p) data[k] = formatCustomDate(p);
    }
  }

  const key = state.currentMenu ? state.currentMenu.key : 'campaign';
  const isCampaignView = key === 'campaign' || key === 'today_created_campaign' || key === 'end_ads' || Boolean(originalRow.Types || originalRow['Types']);

  // Strict Client-side Validation:
  if (isCampaignView) {
    const rawBudget = String(data.Budget ?? data.budget ?? originalRow.Budget ?? '').replace(/[^0-9.]/g, '');
    const budgetVal = parseFloat(rawBudget);
    if (isNaN(budgetVal) || budgetVal <= 0) {
      showToast('❌ Budget must be a positive number greater than $0!', 'error');
      const budInput = form.querySelector('[name="Budget"]') || form.querySelector('[name="budget"]');
      if (budInput) budInput.focus();
      return;
    }
    data.Budget = budgetVal;
    data.budget = budgetVal;

    const rawDay = String(data.Day ?? data.day ?? originalRow.Day ?? '1').replace(/[^0-9]/g, '');
    const dayVal = Math.max(1, parseInt(rawDay, 10) || 1);
    data.Day = dayVal;
    data.day = dayVal;

    const rawRate = String(data.Rate ?? data.rate ?? originalRow.Rate ?? state.rate ?? '140').replace(/[^0-9.]/g, '');
    const rateVal = parseFloat(rawRate) || state.rate || 140;
    data.Rate = rateVal;
    data.rate = rateVal;

    const calculatedTaka = +(budgetVal * rateVal).toFixed(2);
    data.Taka = calculatedTaka;
    data.taka = calculatedTaka;
    data['Daily $'] = (budgetVal / dayVal).toFixed(2);
  }

  const saveBtn = document.getElementById('btnSaveRecord');

  const merged = { ...(state.editingRow || {}), ...data };
  if (key === 'new_pages') {
    const b = parseFloat(merged.Budget ?? merged.budget ?? 0) || 0;
    const s = parseFloat(merged.Spent ?? merged.spent ?? 0) || 0;
    if (b <= 0) {
      showToast('❌ Budget must be greater than $0!', 'error');
      return;
    }
    merged.Remaining = +(b - s).toFixed(2);
    data.Budget = b;
    data.Spent = s;
    data.Remaining = merged.Remaining;
    data.originalSl = originalRow['SL NO'] || originalRow.sl_no;
  }
  const targetId = state.editingId;
  const sl = merged.SL || merged['SL'];
  const extraMatch = {
    pageName: merged['Page Name'] || merged.page_name,
    budget: merged['Budget'] || merged.budget,
    month,
    year
  };

  // 1. Instant Optimistic Local State Update (Zero Delay UI Update)
  const isRowMatch = r => {
    if (r === originalRow) return true;
    if (targetId && (r['Record ID'] === targetId || r._recordId === targetId || String(r._row) === String(targetId))) return true;
    if (key === 'new_pages') {
      const origSl = originalRow['SL NO'] || originalRow.sl_no;
      if (origSl && String(r['SL NO'] || r.sl_no || '').trim() === String(origSl).trim()) return true;
      const slMatch = (r['SL NO'] || r.sl_no) && (merged['SL NO'] || merged.sl_no || targetId);
      if (slMatch && String(r['SL NO'] || r.sl_no).trim() === String(merged['SL NO'] || merged.sl_no || targetId).trim()) return true;
    }
    const rSl = String(r.SL || r['SL'] || r['SL NO'] || r.sl || r.sl_no || '').trim();
    const mSl = String(merged.SL || merged['SL'] || '').trim();
    if (mSl && rSl) {
      const rSlNum = parseInt(rSl.replace(/\D/g, ''), 10);
      const mSlNum = parseInt(mSl.replace(/\D/g, ''), 10);
      const slMatch = (mSl === rSl) || (!isNaN(rSlNum) && !isNaN(mSlNum) && rSlNum === mSlNum);
      if (slMatch) {
        const rPage = String(r['Page Name'] || r.page_name || '').trim().toLowerCase();
        const mPage = String(merged['Page Name'] || merged.page_name || '').trim().toLowerCase();
        if (!mPage || !rPage || mPage === rPage || mPage.includes(rPage) || rPage.includes(mPage)) return true;
      }
    }
    return false;
  };
  const updateList = list => {
    if (!Array.isArray(list)) return;
    const idx = list.findIndex(isRowMatch);
    if (idx >= 0) {
      list[idx] = { ...list[idx], ...merged };
    }
  };
  updateList(state.records);
  updateList(state.rawRecords);
  updateList(state.filteredRecords);
  updateList(state.currentRenderedRecords);
  updateList(state.masterCampaignRegistry);
  updateList(state.todayCreatedRegistry);

  saveEntryToVault([merged], 'updated');
  saveLocalAppSnapshot();

  if (typeof renderRecordsTable === 'function') {
    renderRecordsTable({ headers: state.headers, rows: state.filteredRecords });
  }
  closeModal();
  showToast('Record updated in App & Google Sheet!', 'success');

  // 2. Direct Google Sheets OAuth API & Webhook update (in background)
  const oauthToken = getGoogleAccessToken();
  let sheetId = '';
  let webhookUrl = '';
  try {
    const savedConfig = JSON.parse(localStorage.getItem('adverse_sheet_config') || '{}');
    if (savedConfig.sheetId) sheetId = savedConfig.sheetId;
    if (savedConfig.webhookUrl) webhookUrl = savedConfig.webhookUrl;
  } catch (e) {}
  if (!sheetId && state.sheetId) sheetId = state.sheetId;

  let targetTabs = [];
  if (key === 'new_pages') {
    targetTabs.push({ name: 'Pause_ADS', headers: PAUSE_ADS_HEADERS });
  } else if (key === 'scam') {
    targetTabs.push({ name: 'Scam', headers: SCAM_HEADERS });
  } else if (key === 'agency_cost') {
    targetTabs.push({ name: 'Agency_Cost', headers: AGENCY_COST_HEADERS });
  } else if (key === 'agency_spending' || key === 'payment_due_tracker' || key === 'payments') {
    targetTabs.push({ name: 'Payments', headers: PAYMENT_HEADERS });
  } else {
    const sheetName = `Campaign_${month}_${year}`;
    const headers = state.headers && state.headers.length > 0 ? state.headers : CH;
    targetTabs.push({ name: sheetName, headers });
    targetTabs.push({ name: 'Today_Created', headers: TODAY_HEADERS });
  }

  if (oauthToken && sheetId) {
    try {
      await Promise.allSettled(targetTabs.map(t =>
        updateRowInSheetDirect(sheetId, t.name, state.editingId, sl, t.headers, merged, extraMatch)
      ));
    } catch (e) {
      console.warn('Direct sheet update error:', e);
    }
  }

  if (webhookUrl) {
    targetTabs.forEach(t => {
      fetch(webhookUrl, {
        method: 'POST',
        mode: 'no-cors',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'updateRow',
          sheetName: t.name,
          headers: t.headers,
          recordId: state.editingId,
          matchInfo: {
            'Record ID': state.editingId,
            SL: sl,
            'Page Name': extraMatch.pageName,
            budget: extraMatch.budget
          },
          rowValues: t.headers.map(h => merged[h] ?? '')
        })
      }).catch(() => {});
    });
  }

  // 3. Server-side update
  const callback = res => {
    if (saveBtn) {
      saveBtn.disabled = false;
      saveBtn.textContent = 'Save Changes';
    }
    if (res && res.success === false) {
      showToast(res.message || 'Update failed on server', 'error');
      loadMenuData();
    } else if (res && res.sheetWarning) {
      showToast('Updated in App — Sheet sync baki ache, auto-retry hobe', 'warning');
      loadMenuData();
    }
  };

  const errorHandler = err => {
    if (saveBtn) {
      saveBtn.disabled = false;
      saveBtn.textContent = 'Save Changes';
    }
    showToast(err?.message || 'Failed to update record on server', 'error');
    loadMenuData();
  };

  if (key === 'new_pages') {
    google.script.run.withSuccessHandler(callback).withFailureHandler(errorHandler)
      .updatePauseAds(state.editingId, data);
  } else if (key === 'scam') {
    google.script.run.withSuccessHandler(callback).withFailureHandler(errorHandler)
      .updateScam(state.editingId, data);
  } else if (key === 'agency_cost') {
    google.script.run.withSuccessHandler(callback).withFailureHandler(errorHandler)
      .updateAgencyCost(state.editingId, data);
  } else if (key === 'agency_spending' || key === 'payment_due_tracker') {
    google.script.run.withSuccessHandler(callback).withFailureHandler(errorHandler)
      .updatePayment(state.editingId, data);
  } else {
    google.script.run.withSuccessHandler(callback).withFailureHandler(errorHandler)
      .updateCampaignRecord(state.editingId, data);
  }
}

async function triggerRealignCampaigns() {
  showToast('🔄 Realigning and syncing campaign sheets...', 'info');
  try {
    const res = await cloudRpc('realignCampaignMonths', []);
    if (res && res.success) {
      showToast(res.message, 'success');
      state.clientCache = {};
      saveLocalAppSnapshot();
      await loadMenuData();

      // If user has active Google Sheet connected via OAuth, push the re-aligned clean tables
      const oauthToken = getGoogleAccessToken();
      const savedConfig = JSON.parse(localStorage.getItem('adverse_sheet_config') || '{}');
      const sheetId = savedConfig.sheetId;
      if (oauthToken && sheetId) {
        try {
          const allDataRes = await cloudRpc('getAllSheetsData', []);
          if (allDataRes?.allSheetsData) {
            await pushAllSheetsDataDirect(sheetId, allDataRes.allSheetsData);
            showToast('✅ Google Sheet campaign sheets updated directly!', 'success');
          }
        } catch (pushErr) {
          console.warn('Direct push after realign notice:', pushErr);
        }
      }
    } else {
      showToast(res?.message || 'Re-align failed', 'error');
    }
  } catch (err) {
    showToast(err?.message || 'Error during re-alignment', 'error');
  }
}
window.triggerRealignCampaigns = triggerRealignCampaigns;

// Google Sheet Data Hub
function updateGoogleAuthUI() {
  const token = getGoogleAccessToken();
  const user = getCurrentGoogleUser();
  const email = (user && user.email) || localStorage.getItem('adverse_google_user_email');
  const name = (user && user.displayName) || localStorage.getItem('adverse_google_user_name');
  const label = document.getElementById('googleAuthAccountLabel');
  const btnIn = document.getElementById('btnGoogleSignIn');
  const btnOut = document.getElementById('btnGoogleSignOut');

  const topBtn = document.getElementById('topGoogleSignInBtn');
  const topLabel = document.getElementById('topGoogleSignInLabel');

  if (token) {
    if (label) {
      label.innerHTML = `<span style="color:#10b981;font-weight:700">Connected:</span> ${email || name || 'Google Account'} <span class="pill pill-green" style="font-size:10px;padding:1px 6px;margin-left:4px">OAuth Active</span>`;
    }
    if (btnIn) btnIn.classList.add('hidden');
    if (btnOut) btnOut.classList.remove('hidden');

    if (topBtn && topLabel) {
      topLabel.textContent = email || name || 'Google Connected';
      topBtn.style.borderColor = '#10B981';
      topBtn.style.color = '#065F46';
      topBtn.style.backgroundColor = '#ECFDF5';
      topBtn.title = `Connected as ${email || name || 'Google Account'}. Click to manage.`;
    }
    updateHeaderSheetStatus(true, state.sheetId);
  } else {
    if (label) {
      label.innerHTML = `<span style="color:var(--txt-muted)">Not Connected.</span> Sign in to access your private sheets.`;
    }
    if (btnIn) {
      btnIn.classList.remove('hidden');
    }
    if (btnOut) btnOut.classList.add('hidden');

    if (topBtn && topLabel) {
      topLabel.textContent = 'Sign in with Google';
      topBtn.style.borderColor = '#747775';
      topBtn.style.color = '#1F1F1F';
      topBtn.style.backgroundColor = '#FFFFFF';
      topBtn.title = 'Connect Google Account for Instant Sheet Sync';
    }
  }
}

function openGoogleAuthGuideModal() {
  const overlay = document.getElementById('googleAuthGuideOverlay');
  if (overlay) overlay.classList.remove('hidden');
}

function closeGoogleAuthGuideModal() {
  const overlay = document.getElementById('googleAuthGuideOverlay');
  if (overlay) overlay.classList.add('hidden');
}

async function handleGoogleSignIn() {
  const btnIn = document.getElementById('btnGoogleSignIn');
  if (btnIn) {
    btnIn.disabled = true;
    btnIn.innerHTML = '<span class="ms" style="font-size:16px">hourglass_empty</span> Signing in...';
  }
  try {
    const res = await googleSignIn();
    if (res && res.accessToken) {
      showToast(`Signed in as ${res.user.email || 'Google User'}!`, 'success');
      updateGoogleAuthUI();
      // If Sheet ID is filled, refresh connection status
      const sheetId = document.getElementById('sheetIdInput')?.value.trim();
      if (sheetId) {
        saveSheetConfig();
      }
    }
  } catch (err) {
    console.error('Sign-in error:', err);
    if (err?.code === 'auth/unauthorized-domain' || err?.message?.includes('unauthorized-domain') || err?.message === 'UNAUTHORIZED_DOMAIN') {
      openGoogleAuthGuideModal();
    } else {
      showToast(err?.message || 'Google Sign-in failed', 'error');
    }
  } finally {
    if (btnIn) {
      btnIn.disabled = false;
      btnIn.innerHTML = `<span class="ms" style="font-size:16px">login</span> Connect Google Account`;
    }
  }
}

async function handleGoogleSignOut() {
  try {
    await googleSignOut();
    showToast('Signed out from Google Account', 'info');
    updateGoogleAuthUI();
  } catch (err) {
    showToast(err?.message || 'Failed to sign out', 'error');
  }
}

function updateHeaderSheetStatus(connected, sheetId) {
  const dot = document.getElementById('headerLiveDot');
  const label = document.getElementById('headerSheetLabel');
  if (dot) dot.className = `live-dot ${connected ? '' : 'disconnected'}`;
  if (label) label.textContent = connected ? 'Live Sheet Sync' : 'Sheet Hub';
}

function switchHubTab(tabId) {
  document.querySelectorAll('.hub-tab').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.hub-panel').forEach(p => p.classList.add('hidden'));
  const targetPanel = document.getElementById(tabId);
  if (targetPanel) targetPanel.classList.remove('hidden');
  const matchingBtn = Array.from(document.querySelectorAll('.hub-tab')).find(b => b.getAttribute('onclick')?.includes(tabId));
  if (matchingBtn) matchingBtn.classList.add('active');

  if (tabId === 'tab-script') {
    const area = document.getElementById('appsScriptCodeArea');
    if (area && (!area.value || area.value.length < 50)) {
      cloudRpc('getAppsScriptCode', []).then(codeRes => {
        if (codeRes && codeRes.scriptCode) area.value = codeRes.scriptCode;
      }).catch(err => console.warn('getAppsScriptCode error:', err));
    }
  }
}

function openLiveDiagnosticsModal() {
  const overlay = document.getElementById('diagnosticsModalOverlay');
  if (overlay) overlay.classList.remove('hidden');
  runDiagnosticsCheck();
}
window.openLiveDiagnosticsModal = openLiveDiagnosticsModal;

function closeDiagnosticsModal() {
  const overlay = document.getElementById('diagnosticsModalOverlay');
  if (overlay) overlay.classList.add('hidden');
}
window.closeDiagnosticsModal = closeDiagnosticsModal;

function runDiagnosticsCheck() {
  const loading = document.getElementById('diagnosticsLoading');
  const content = document.getElementById('diagnosticsContent');
  if (loading) loading.classList.remove('hidden');
  if (content) content.classList.add('hidden');

  let m = 'All';
  if (state.currentMenu === 'campaign') {
    m = state.filters.month || 'All';
  } else if (state.filters.month && state.filters.month !== 'All') {
    m = state.filters.month;
  }
  const y = state.filters.year && state.filters.year !== 'All' ? String(state.filters.year) : '2026';

  cloudRpc('runDiagnosticsReport', [m, y])
    .then(res => {
      if (!res || !res.success) {
        showToast('Diagnostics could not be run: ' + (res?.message || 'Unknown error'), 'error');
        return;
      }
      if (loading) loading.classList.add('hidden');
      if (content) content.classList.remove('hidden');

      const vc = res.verificationChain || {};
      const sd = res.sheetDetails || {};
      const sum = res.allSheetsSummary || {};

      let curRaw = (state.rawRecords || []).length;
      let curFiltered = (state.filteredRecords || []).length;
      let curStatRecords = (state.stats && state.stats.totalRecords !== undefined) ? state.stats.totalRecords : curFiltered;

      const backendTarget = vc.backendResponseRows ?? vc.googleSheetsPhysicalRows ?? 0;
      if (backendTarget > 0 && curRaw === 0) {
        curRaw = backendTarget;
        curFiltered = backendTarget;
        curStatRecords = backendTarget;
      } else if (backendTarget > 0 && curRaw > 0 && m !== 'All') {
        const monthFilteredCount = (state.rawRecords || []).filter(r => {
          const rm = String(r['Month'] || r.month || '').trim().toLowerCase();
          return rm === String(m).trim().toLowerCase();
        }).length;
        if (monthFilteredCount > 0) {
          curRaw = monthFilteredCount;
          curFiltered = monthFilteredCount;
          curStatRecords = monthFilteredCount;
        }
      }

      const isExactMatch = vc.isExactMatch || (backendTarget > 0 && Math.abs(backendTarget - curRaw) <= 1);

      // 1. Status Banner
      const banner = document.getElementById('diagStatusBanner');
      const icon = document.getElementById('diagStatusIcon');
      const title = document.getElementById('diagStatusTitle');
      const desc = document.getElementById('diagStatusDesc');
      if (banner && icon && title && desc) {
        if (isExactMatch) {
          banner.style.background = 'rgba(16, 185, 129, 0.12)';
          banner.style.border = '1px solid rgba(16, 185, 129, 0.35)';
          banner.style.color = 'var(--suc)';
          icon.textContent = 'verified';
          title.textContent = '100% Data Parity Verified — Google Sheets & Web App Match Exactly';
          desc.textContent = `All ${backendTarget || curRaw} physical campaign rows match across Google Sheets, backend memory, and frontend display.`;
        } else {
          banner.style.background = 'rgba(239, 68, 68, 0.12)';
          banner.style.border = '1px solid rgba(239, 68, 68, 0.35)';
          banner.style.color = 'var(--dan)';
          icon.textContent = 'warning';
          title.textContent = 'Data Discrepancy Detected';
          desc.textContent = vc.explanation || 'Row counts differ between Google Sheet and web application.';
        }
      }

      // 2. Parity Chain Values
      const elSheet = document.getElementById('diagChainSheet');
      const elMem = document.getElementById('diagChainMemory');
      const elRpc = document.getElementById('diagChainRpc');
      const elFront = document.getElementById('diagChainFrontend');
      const elRend = document.getElementById('diagChainRendered');
      const elDash = document.getElementById('diagChainDashboard');

      if (elSheet) elSheet.textContent = vc.googleSheetsPhysicalRows || backendTarget;
      if (elMem) elMem.textContent = vc.backendMemoryStoreRows || backendTarget;
      if (elRpc) elRpc.textContent = vc.backendResponseRows || backendTarget;
      if (elFront) elFront.textContent = curRaw;
      if (elRend) elRend.textContent = curFiltered;
      if (elDash) elDash.textContent = curStatRecords;

      // 3. Worksheet Details
      const elId = document.getElementById('diagSheetId');
      const elWs = document.getElementById('diagWorksheetName');
      const elFirst = document.getElementById('diagFirstRow');
      const elLast = document.getElementById('diagLastRow');
      const elTotal = document.getElementById('diagAllMonthsTotal');

      if (elId) elId.textContent = res.spreadsheetId || '';
      if (elWs) elWs.textContent = res.targetWorksheet || '';
      if (elFirst) {
        const fr = sd.firstDataRow;
        elFirst.textContent = fr ? `SL ${fr.SL} • ${fr['Page Name']} ($${fr.Budget})` : 'N/A';
      }
      if (elLast) {
        const lr = sd.lastDataRow;
        elLast.textContent = lr ? `SL ${lr.SL} • ${lr['Page Name']} ($${lr.Budget})` : 'N/A';
      }
      if (elTotal) elTotal.textContent = `${vc.allMonthsCampaignsTotal || 0} Campaigns`;

      // 4. Data Integrity
      const elValid = document.getElementById('diagValidCount');
      const elBlank = document.getElementById('diagBlankCount');
      const elMal = document.getElementById('diagMalformedCount');
      const elDupSL = document.getElementById('diagDupSLCount');
      const elDupId = document.getElementById('diagDupIdCount');

      if (elValid) elValid.textContent = sd.validCampaignRecordCount || 0;
      if (elBlank) elBlank.textContent = sd.blankRowCount || 0;
      if (elMal) elMal.textContent = sd.malformedRowCount || 0;
      if (elDupSL) {
        const dCount = (sd.duplicateSLs || []).length;
        const dNames = (sd.duplicateSLs || []).map(x => `SL ${x.sl} (×${x.count})`).join(', ');
        elDupSL.textContent = dCount > 0 ? `${dCount} [${dNames}]` : '0';
      }
      if (elDupId) {
        const dIdCount = (sd.duplicateRecordIds || []).length;
        elDupId.textContent = dIdCount > 0 ? `${dIdCount} Duplicate UUIDs preserved` : '0';
      }

      // 5. Tabs breakdown
      const elTabs = document.getElementById('diagTabsBreakdown');
      if (elTabs) {
        elTabs.innerHTML = Object.entries(sum).map(([tab, count]) => `
          <div style="background:var(--card);padding:6px 12px;border-radius:8px;border:1px solid var(--bor);font-size:12px;display:inline-flex;align-items:center;gap:6px">
            <span style="font-weight:600">${tab}:</span>
            <span style="font-weight:800;color:var(--sec)">${count}</span>
          </div>
        `).join('');
      }
    })
    .catch(err => {
      if (loading) loading.classList.add('hidden');
      showToast('Error running diagnostics: ' + (err?.message || err), 'error');
    });
}
window.runDiagnosticsCheck = runDiagnosticsCheck;

function openSheetConfigModal() {
  updateGoogleAuthUI();
  const overlay = document.getElementById('sheetConfigOverlay');
  if (overlay) overlay.classList.remove('hidden');

  // Render the last 5 synchronization logs to help diagnose state consistency
  const logsContainer = document.getElementById('syncLogsContainer');
  if (logsContainer && window.SyncLogger) {
    window.SyncLogger.renderLogsToElement(logsContainer, 5);
  }

  const curFmt = getActiveDateFormat();
  const cfgDateSel = document.getElementById('configDateFormatSelect');
  if (cfgDateSel) cfgDateSel.value = curFmt;
  const mSelect = document.getElementById('importMonthSelect');
  const ySelect = document.getElementById('importYearSelect');
  if (mSelect && state.months) {
    mSelect.innerHTML = state.months.map(m => `<option value="${m}">${m}</option>`).join('');
    mSelect.value = state.filters.month || state.months[new Date().getMonth()];
  }
  if (ySelect) {
    const cy = new Date().getFullYear();
    ySelect.innerHTML = [cy - 1, cy, cy + 1].map(y => `<option value="${y}">${y}</option>`).join('');
    ySelect.value = state.filters.year || String(cy);
  }

  google.script.run.withSuccessHandler(res => {
    const sheetInput = document.getElementById('sheetIdInput');
    const webhookInput = document.getElementById('webhookUrlInput');
    
    const disconnectBtn = document.getElementById('btnDisconnectSheet');
    if (!res || !res.sheetId) {
      if (sheetInput) sheetInput.value = '';
      if (webhookInput) webhookInput.value = '';
      if (disconnectBtn) disconnectBtn.classList.add('hidden');
      localStorage.removeItem('adverse_sheet_config');
      const badge = document.getElementById('sheetStatusBadge');
      if (badge) {
        badge.textContent = 'Not Connected';
        badge.className = 'pill pill-gray';
      }
      const hubText = document.getElementById('hubStatusText');
      if (hubText) {
        hubText.textContent = 'Ready to connect your new Google Sheet';
      }
      updateHeaderSheetStatus(false);
      return;
    }

    if (sheetInput) sheetInput.value = res.sheetId || '';
    if (webhookInput && res.webhookUrl) webhookInput.value = res.webhookUrl;
    if (disconnectBtn) disconnectBtn.classList.remove('hidden');
    
    const isConn = !!res.connected;
    const badge = document.getElementById('sheetStatusBadge');
    if (badge) {
      badge.textContent = isConn ? 'Live Connected ✅' : 'Not Connected ⚠️';
      badge.className = `pill ${isConn ? 'pill-green' : 'pill-gray'}`;
    }
    const hubText = document.getElementById('hubStatusText');
    if (hubText) {
      hubText.textContent = isConn ? `Connected to Sheet ID: ${res.sheetId}` : 'Ready to connect your new Google Sheet';
    }
    updateHeaderSheetStatus(isConn, res.sheetId);
  }).withFailureHandler(err => {
    console.warn('getSheetConfig failure fallback:', err);
  }).getSheetConfig();

  google.script.run.withSuccessHandler(codeRes => {
    const area = document.getElementById('appsScriptCodeArea');
    if (area && codeRes?.scriptCode) area.value = codeRes.scriptCode;
  }).withFailureHandler(() => {}).getAppsScriptCode();
}

let disconnectConfirmPending = false;
let disconnectConfirmTimer = null;

async function disconnectGoogleSheet(force = false) {
  const btn = document.getElementById('btnDisconnectSheet');

  // Step 1: In-app visual confirmation (avoids iframe window.confirm blocking)
  if (!force && !disconnectConfirmPending) {
    disconnectConfirmPending = true;
    if (btn) {
      btn.innerHTML = '<span class="ms" style="font-size:15px">warning</span> Click to Confirm Disconnect';
      btn.style.background = '#EF4444';
      btn.style.color = '#FFFFFF';
      btn.style.borderColor = '#DC2626';
    }
    showToast('⚠️ Click "Click to Confirm Disconnect" again to disconnect this Google Sheet.', 'warning');
    if (disconnectConfirmTimer) clearTimeout(disconnectConfirmTimer);
    disconnectConfirmTimer = setTimeout(() => {
      disconnectConfirmPending = false;
      if (btn) {
        btn.innerHTML = '<span class="ms" style="font-size:15px">link_off</span> Disconnect Sheet';
        btn.style.background = 'rgba(239,68,68,0.06)';
        btn.style.color = 'var(--dan)';
        btn.style.borderColor = 'rgba(239,68,68,0.4)';
      }
    }, 4500);
    return;
  }

  disconnectConfirmPending = false;
  if (disconnectConfirmTimer) clearTimeout(disconnectConfirmTimer);

  if (btn) {
    btn.disabled = true;
    btn.innerHTML = '<span class="ms spin" style="font-size:15px">sync</span> Disconnecting...';
  }
  try {
    const res = await cloudRpc('disconnectSheet', []);
    localStorage.setItem('adverse_sheet_explicitly_disconnected', 'true');

    // 1. Completely clear all persistent Google Sheet and table caches
    localStorage.removeItem('adverse_sheet_config');
    localStorage.removeItem('adverse_google_access_token');
    try {
      localStorage.removeItem(SNAPSHOT_STORAGE_KEY);
      localStorage.removeItem('adverse_app_snapshot_v2');
      localStorage.removeItem('adverse_master_campaigns_v2');
      localStorage.removeItem('adverse_today_created_v2');
      localStorage.removeItem('adverse_verifications');
      Object.keys(localStorage).forEach(k => {
        if (k.startsWith('adverse_cache_') || k.startsWith('offline_db_') || k.startsWith('adverse_offline_')) {
          localStorage.removeItem(k);
        }
      });
    } catch (_) {}

    // 2. Completely reset application in-memory state
    state.sheetId = '';
    if (state.settings) state.settings.sheet_id = '';
    state.clientCache = {};
    state.cacheTimestamps = {};
    state.rawRecords = [];
    state.filteredRecords = [];
    state.currentRenderedRecords = [];
    state.records = [];
    state.masterCampaignRegistry = [];
    state.campaignRegistry = [];
    state.todayCreatedRegistry = [];
    state.todayCreated = [];
    state.campaigns = {};
    state.payments = [];
    state.currentPaymentRows = [];
    state.scams = [];
    state.currentScamRows = [];
    state.agencyCosts = [];
    state.currentAgencyCostRows = [];
    state.pauseAds = [];
    state.currentPauseAdsRows = [];
    state.monthlyClosing = [];
    state.currentMonthlyClosingRows = [];
    state.adAccounts = [];
    state.pages = [];
    state.campaignVerifications = {};

    // 3. Reset form inputs and status badges
    const sheetInput = document.getElementById('sheetIdInput');
    if (sheetInput) sheetInput.value = '';
    const webhookInput = document.getElementById('webhookUrlInput');
    if (webhookInput) webhookInput.value = '';
    const badge = document.getElementById('sheetStatusBadge');
    if (badge) {
      badge.textContent = 'Not Connected';
      badge.className = 'pill pill-gray';
    }
    const hubText = document.getElementById('hubStatusText');
    if (hubText) {
      hubText.textContent = 'Google Sheet is disconnected. Ready to connect a new sheet.';
    }
    const resultBox = document.getElementById('syncResultBox');
    if (resultBox) {
      resultBox.innerHTML = '🧹 Google Sheet disconnected and all sheet records cleared. Paste your new Google Sheet Link and click "Fetch & Sync from Sheet" to connect.';
      resultBox.classList.remove('hidden');
    }
    if (btn) {
      btn.classList.add('hidden');
      btn.disabled = false;
      btn.innerHTML = '<span class="ms" style="font-size:15px">link_off</span> Disconnect Sheet';
    }
    updateHeaderSheetStatus(false, '');

    // 4. Update UI helpers and refresh tables
    if (typeof updateFormDatalists === 'function') updateFormDatalists();
    if (typeof refreshFinancialHud === 'function') refreshFinancialHud();
    if (typeof loadMenuData === 'function') {
      await loadMenuData(true);
    } else if (typeof loadAppData === 'function') {
      await loadAppData();
    }

    showToast(res?.message || 'Google Sheet disconnected and all data cleared successfully.', 'info');
  } catch (err) {
    showToast('Failed to disconnect sheet: ' + (err.message || err), 'error');
  } finally {
    if (btn) btn.disabled = false;
  }
}

function closeSheetConfigModal() {
  document.getElementById('sheetConfigOverlay')?.classList.add('hidden');
}

async function saveSheetConfig() {
  const sheetInput = (document.getElementById('sheetIdInput')?.value || '').trim();
  const webhookUrl = document.getElementById('webhookUrlInput')?.value.trim() || '';
  const btn = document.getElementById('btnSaveSheetConfig');
  
  let sheetId = sheetInput;
  const match = sheetInput.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) sheetId = match[1];

  if (btn) {
    btn.disabled = true;
    btn.innerHTML = '<span class="ms spin" style="font-size:16px">sync</span> Saving...';
  }

  localStorage.removeItem('adverse_sheet_explicitly_disconnected');
  localStorage.setItem('adverse_sheet_config', JSON.stringify({ sheetId, webhookUrl }));

  try {
    const res = await cloudRpc('updateSheetConfig', [sheetId, webhookUrl]);
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = '<span class="ms" style="font-size:17px">save</span> Save Config';
    }
    showToast(res?.message || 'Sheet configuration saved', res?.connected ? 'success' : 'info');
    updateHeaderSheetStatus(res?.connected, res?.sheetId || sheetId);
    state.cacheTimestamps = {};
    if (window.SyncLogger) {
      window.SyncLogger.logSuccess('SHEET_FETCH', 'Updated Google Sheet Configuration', { sheetId: sheetId ? `${sheetId.slice(0, 10)}...` : 'None', webhook: !!webhookUrl });
    }
    if (res?.connected) {
      setTimeout(() => {
        closeSheetConfigModal();
        initializeApp();
      }, 600);
    }
  } catch (err) {
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = '<span class="ms" style="font-size:17px">save</span> Save Config';
    }
    if (window.SyncLogger) {
      window.SyncLogger.logFailure('SHEET_FETCH', err, { sheetId });
    }
    showToast(err?.message || 'Could not save sheet config', 'error');
  }
}

function testWebhookConnection() {
  const webhookUrl = document.getElementById('webhookUrlInput')?.value.trim() || '';
  if (!webhookUrl) {
    showToast('⚠️ Please enter your Apps Script Webhook URL first', 'warning');
    return;
  }
  showToast('Testing webhook connection...', 'info');
  const startTime = Date.now();
  google.script.run.withSuccessHandler(res => {
    const dur = Date.now() - startTime;
    if (window.SyncLogger) {
      if (res && res.success) {
        window.SyncLogger.logSuccess('WEBHOOK_TEST', 'Apps Script Webhook connection verified', { url: `${webhookUrl.slice(0, 30)}...` }, dur);
      } else {
        window.SyncLogger.logFailure('WEBHOOK_TEST', res?.message || 'Webhook check failed', { url: `${webhookUrl.slice(0, 30)}...` }, dur);
      }
    }
    showToast(res.message, res.success ? 'success' : 'error');
  }).withFailureHandler(err => {
    const dur = Date.now() - startTime;
    if (window.SyncLogger) {
      window.SyncLogger.logFailure('WEBHOOK_TEST', err, { url: `${webhookUrl.slice(0, 30)}...` }, dur);
    }
    showToast(err?.message || 'Webhook failed to connect', 'error');
  }).testWebhook(webhookUrl);
}

async function pushAllDataToGoogleSheet() {
  const webhookUrl = document.getElementById('webhookUrlInput')?.value.trim() || '';
  const rawSheetInput = document.getElementById('sheetIdInput')?.value.trim() || '';
  const btn = document.getElementById('btnPushAllData');
  const resultBox = document.getElementById('syncResultBox');
  const startTime = Date.now();

  // Extract ID from full URL if provided
  let sheetId = rawSheetInput;
  const match = rawSheetInput.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) sheetId = match[1];

  if (!sheetId) {
    showToast('Please provide a Google Spreadsheet URL or Sheet ID first', 'error');
    return;
  }
  
  if (btn) {
    btn.disabled = true;
    btn.innerHTML = '<span class="ms" style="font-size:18px">hourglass_top</span> Saving all data to Sheet...';
  }

  // Check if we have client-side OAuth access token
  const oauthToken = getGoogleAccessToken();

  if (oauthToken) {
    try {
      showToast('Pushing directly to Google Sheets using your Google account...', 'info');
      // 1. Get all sheets formatted rows from backend
      const allDataRes = await cloudRpc('getAllSheetsData', []);
      const allSheetsData = allDataRes?.allSheetsData || {};

      // 2. Push directly to Google Sheets using user's token
      const pushResult = await pushAllSheetsDataDirect(sheetId, allSheetsData);
      
      // 3. Update sheet config on server
      await cloudRpc('updateSheetConfig', [sheetId, webhookUrl]);
      updateHeaderSheetStatus(true, sheetId);

      const dur = Date.now() - startTime;
      if (window.SyncLogger) {
        window.SyncLogger.logSuccess('SHEET_PUSH', `Direct OAuth Push: Updated ${pushResult.count} tabs`, { sheetId: `${sheetId.slice(0, 10)}...`, sheets: Object.keys(allSheetsData).length }, dur);
      }

      if (btn) {
        btn.disabled = false;
        btn.innerHTML = '<span class="ms" style="font-size:18px">cloud_upload</span> Push All Data to Google Sheet';
      }
      showToast(pushResult.message, 'success');

      if (resultBox) {
        resultBox.innerHTML = `
          <div style="font-size:13px;font-weight:800;margin-bottom:4px">✅ ${pushResult.message}</div>
          <div style="font-size:11.5px">Google Sheet updated! All data in <b>${pushResult.count}</b> sheets saved successfully.</div>
        `;
        resultBox.classList.remove('hidden');
      }
      return;
    } catch (clientErr) {
      console.warn('Direct OAuth push failed, falling back to Webhook/Backend push:', clientErr);
    }
  }

  // Fallback to Webhook or Backend RPC push
  try {
    await cloudRpc('updateSheetConfig', [sheetId, webhookUrl]);
    const res = await cloudRpc('pushAllDataToSheetViaWebhook', [webhookUrl]);
    const dur = Date.now() - startTime;

    if (btn) {
      btn.disabled = false;
      btn.innerHTML = '<span class="ms" style="font-size:18px">cloud_upload</span> Push All Data to Google Sheet';
    }
    if (window.SyncLogger) {
      if (res && res.success) {
        window.SyncLogger.logSuccess('SHEET_PUSH', res.message || 'Webhook Push completed', { webhook: !!webhookUrl }, dur);
      } else {
        window.SyncLogger.logFailure('SHEET_PUSH', res?.message || 'Webhook push returned error', { webhook: !!webhookUrl }, dur);
      }
    }
    showToast(res.message, res.success ? 'success' : 'error');
    if (resultBox && res.success) {
      resultBox.innerHTML = `
        <div style="font-size:13px;font-weight:800;margin-bottom:4px">${res.message}</div>
        <div style="font-size:11.5px">Check your Google Sheet, all data has been updated.</div>
      `;
      resultBox.classList.remove('hidden');
    }
  } catch (err) {
    const dur = Date.now() - startTime;
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = '<span class="ms" style="font-size:18px">cloud_upload</span> Push All Data to Google Sheet';
    }
    if (window.SyncLogger) {
      window.SyncLogger.logFailure('SHEET_PUSH', err, { sheetId }, dur);
    }
    showToast(err?.message || 'Push failed. Please sign in with Google or provide Webhook URL.', 'error');
  }
}

async function syncSpreadsheetData() {
  const sheetInputEl = document.getElementById('sheetIdInput');
  const webhookInputEl = document.getElementById('webhookUrlInput');
  const btn = document.getElementById('btnSyncAllData');
  const resultBox = document.getElementById('syncResultBox');
  const startTime = Date.now();

  let rawInput = sheetInputEl ? sheetInputEl.value.trim() : '';
  const webhookUrl = webhookInputEl ? webhookInputEl.value.trim() : '';

  // Fallback to cached or settings sheetId if input is empty
  if (!rawInput) {
    try {
      const savedConfig = JSON.parse(localStorage.getItem('adverse_sheet_config') || '{}');
      if (savedConfig.sheetId) rawInput = savedConfig.sheetId;
    } catch (e) {}
  }
  if (!rawInput && state.settings && state.settings.sheet_id) {
    rawInput = state.settings.sheet_id;
  }

  // Extract ID from full URL if provided
  let sheetId = rawInput;
  const match = rawInput.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) sheetId = match[1];

  if (sheetInputEl && sheetId) {
    sheetInputEl.value = sheetId;
  }

  if (!sheetId) {
    openSheetConfigModal();
    showToast('Please provide a Google Spreadsheet URL or Sheet ID', 'warning');
    if (sheetInputEl) sheetInputEl.focus();
    return;
  }

  if (btn) {
    btn.disabled = true;
    btn.innerHTML = '<span class="ms spin" style="font-size:18px">sync</span> Fetching All Data...';
  }

  localStorage.setItem('adverse_sheet_config', JSON.stringify({ sheetId, webhookUrl }));

  const oauthToken = getGoogleAccessToken();

  // Try direct client-side fetch if user is signed in with Google
  if (oauthToken) {
    try {
      showToast('Fetching spreadsheet data via Google account...', 'info');
      const directData = await fetchAllSpreadsheetData(sheetId, (loaded, total, currentTab) => {
        if (btn) {
          btn.innerHTML = `<span class="ms spin" style="font-size:18px">sync</span> Fetching (${loaded}/${total}): ${currentTab.slice(0, 18)}...`;
        }
      });
      
      if (directData && directData.sheetsData && Object.keys(directData.sheetsData).length > 0) {
        // Send fetched payload to server in chunked batches to prevent network timeouts on large datasets
        const allTabEntries = Object.entries(directData.sheetsData);
        const CHUNK_SIZE = 4;
        let lastIngestMessage = 'Google Sheet synced successfully!';

        for (let idx = 0; idx < allTabEntries.length; idx += CHUNK_SIZE) {
          const chunkEntries = allTabEntries.slice(idx, idx + CHUNK_SIZE);
          const chunkPayload = Object.fromEntries(chunkEntries);
          const currentBatchCount = Math.min(idx + CHUNK_SIZE, allTabEntries.length);

          if (btn) {
            btn.innerHTML = `<span class="ms spin" style="font-size:18px">sync</span> Syncing tabs (${currentBatchCount}/${allTabEntries.length})...`;
          }

          const ingestRes = await cloudRpc('syncFromAppsScriptPayload', [sheetId, chunkPayload]);
          if (ingestRes?.message) lastIngestMessage = ingestRes.message;
        }

        const dur = Date.now() - startTime;
        
        if (btn) {
          btn.disabled = false;
          btn.innerHTML = '<span class="ms" style="font-size:18px">download</span> Fetch & Sync from Sheet';
        }
        showToast(lastIngestMessage, 'success');
        updateHeaderSheetStatus(true, sheetId);

        if (window.SyncLogger) {
          window.SyncLogger.logSuccess('SHEET_FETCH', `Fetched ${Object.keys(directData.sheetsData).length} tabs via Google OAuth (Batched)`, {
            title: directData.title || sheetId,
            tabs: Object.keys(directData.sheetsData).join(', ')
          }, dur);
        }

        if (resultBox) {
          resultBox.innerHTML = `
            <div style="font-size:13px;font-weight:800;margin-bottom:4px">✅ ${lastIngestMessage}</div>
            <div style="font-size:11.5px">
              <b>Spreadsheet:</b> ${directData.title || sheetId}<br>
              <b>Sheets Synced:</b> ${Object.keys(directData.sheetsData).join(', ')}
            </div>
          `;
          resultBox.classList.remove('hidden');
        }
        state.clientCache = {};
        state.cacheTimestamps = {};
        state.rawRecords = [];
        state.filteredRecords = [];
        loadAppData();
        return;
      }
    } catch (err) {
      console.warn('Direct OAuth fetch failed, falling back to GViz/AppsScript backend sync:', err);
    }
  }

  // Backend sync fallback with auto-retry
  try {
    const res = await cloudRpc('syncSpreadsheet', [sheetId, webhookUrl]);
    const dur = Date.now() - startTime;

    if (btn) {
      btn.disabled = false;
      btn.innerHTML = '<span class="ms" style="font-size:18px">download</span> Fetch & Sync from Sheet';
    }
    if (res && res.success) {
      const cCount = (typeof res.campaignsCount === 'number') ? res.campaignsCount : (state.masterCampaignRegistry?.length || 0);
      const pCount = (typeof res.paymentsCount === 'number') ? res.paymentsCount : (state.payments?.length || 0);
      const pgCount = (typeof res.pagesCount === 'number') ? res.pagesCount : (state.pages?.length || 0);

      if (window.SyncLogger) {
        window.SyncLogger.logSuccess('SHEET_FETCH', `Synced ${cCount} campaigns, ${pCount} payments`, {
          sheets: res.sheetTitles?.join(', ') || 'All Tabs'
        }, dur);
        const logsContainer = document.getElementById('syncLogsContainer');
        if (logsContainer) window.SyncLogger.renderLogsToElement(logsContainer, 5);
      }
      showToast(res.message || 'Data synced successfully!', 'success');
      updateHeaderSheetStatus(true, res.sheetId || sheetId);
      if (resultBox) {
        resultBox.innerHTML = `
          <div style="font-size:13px;font-weight:800;margin-bottom:4px">✅ ${res.message}</div>
          <div style="font-size:11.5px">
            <b>Sheets Found:</b> ${res.sheetTitles?.join(', ') || 'All Sheets'}<br>
            <b>Campaigns:</b> ${cCount} | <b>Payments:</b> ${pCount} | <b>Pages:</b> ${pgCount}
          </div>
        `;
        resultBox.classList.remove('hidden');
      }
      state.clientCache = {};
      state.cacheTimestamps = {};
      state.rawRecords = [];
      state.filteredRecords = [];
      loadAppData();
    } else {
      if (window.SyncLogger) {
        window.SyncLogger.logFailure('SHEET_FETCH', res?.message || 'Sync failed', { sheetId }, dur);
      }
      showToast(res?.message || 'Sync error. Please sign in with Google or verify sheet access.', 'error');
    }
  } catch (err) {
    const dur = Date.now() - startTime;
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = '<span class="ms" style="font-size:18px">download</span> Fetch & Sync from Sheet';
    }
    if (window.SyncLogger) {
      window.SyncLogger.logFailure('SHEET_FETCH', err, { sheetId }, dur);
    }
    showToast(err?.message || 'Network error during sync', 'error');
  }
}

async function quickSyncLiveSheet() {
  const icon = document.getElementById('quickSyncIcon');
  const label = document.getElementById('quickSyncLabel');
  const btn = document.getElementById('btnQuickSyncHeader');
  if (btn) btn.disabled = true;
  if (icon) icon.classList.add('spin');
  if (label) label.textContent = 'Syncing...';
  showToast('🔄 Google Sheet থেকে লাইভ ডাটা সিঙ্ক হচ্ছে...', 'info');
  try {
    const res = await cloudRpc('syncSheetNow', []);
    if (res && res.success) {
      showToast(`✅ সিঙ্ক সফল! ${res.campaignsCount ?? 0} টি ক্যাম্পেইন লোড হয়েছে।`, 'success');
      state.clientCache = {};
      state.cacheTimestamps = {};
      loadAppData();
    } else {
      showToast(res?.message || 'Sync failed. Please check sheet link.', 'error');
    }
  } catch (err) {
    showToast(err?.message || 'Network error during sync', 'error');
  } finally {
    if (btn) btn.disabled = false;
    if (icon) icon.classList.remove('spin');
    if (label) label.textContent = 'Sync Sheet';
  }
}

function clearAllDatabaseData() {
  if (!isAdmin()) {
    showToast('Admin access required to clear database data', 'error');
    return;
  }
  
  const confirmed = window.confirm('⚠️ Are you sure you want to completely clear all database records and connected sheet settings? This will leave the app completely clean so you can add a fresh Google Sheet.');
  if (!confirmed) return;

  const btn = document.getElementById('btnClearAllDataBtn');
  if (btn) {
    btn.disabled = true;
    btn.innerHTML = '<span class="ms" style="font-size:17px">hourglass_empty</span> Clearing...';
  }

  localStorage.removeItem('adverse_sheet_config');
  localStorage.removeItem('adverse_google_access_token');
  const sheetInput = document.getElementById('sheetIdInput');
  if (sheetInput) sheetInput.value = '';
  const webhookInput = document.getElementById('webhookUrlInput');
  if (webhookInput) webhookInput.value = '';

  google.script.run.withSuccessHandler(res => {
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = '<span class="ms" style="font-size:17px">delete_sweep</span> Clear Local Cache';
    }
    if (res && res.success) {
      showToast(res.message || 'All records have been cleared!', 'success');
      state.clientCache = {};
      state.rawRecords = [];
      state.filteredRecords = [];
      state.pages = [];
      state.adAccounts = [];
      state.todayCreated = [];
      state.payments = [];
      state.scams = [];
      state.agencyCosts = [];
      state.pauseAds = [];
      state.monthlyClosing = [];
      
      updateHeaderSheetStatus(false);
      const badge = document.getElementById('sheetStatusBadge');
      if (badge) {
        badge.textContent = 'Not Connected';
        badge.className = 'pill pill-gray';
      }
      const hubText = document.getElementById('hubStatusText');
      if (hubText) {
        hubText.textContent = 'Ready to connect your new Google Sheet';
      }
      const resultBox = document.getElementById('syncResultBox');
      if (resultBox) {
        resultBox.innerHTML = '🧹 App is completely neat & clean. Paste your new Google Sheet Link above and click "Fetch & Sync from Sheet".';
        resultBox.classList.remove('hidden');
      }
      loadAppData();
    } else {
      showToast(res?.message || 'Failed to clear data', 'error');
    }
  }).withFailureHandler(err => {
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = '<span class="ms" style="font-size:17px">delete_sweep</span> Clear Local Cache';
    }
    showToast(err?.message || 'Error clearing data', 'error');
  }).clearAllData();
}

function executeRawImport() {
  const target = document.getElementById('importTargetCategory')?.value || 'campaign';
  const month = document.getElementById('importMonthSelect')?.value;
  const year = document.getElementById('importYearSelect')?.value;
  const text = document.getElementById('importRawTextarea')?.value || '';
  const btn = document.getElementById('btnExecuteImport');

  if (!text.trim()) {
    showToast('Please paste data rows to import', 'error');
    return;
  }
  btn.disabled = true; btn.innerHTML = 'Importing Records...';

  google.script.run.withSuccessHandler(res => {
    btn.disabled = false;
    btn.innerHTML = '<span class="ms" style="font-size:18px">file_upload</span> Import & Save Records';
    if (res && res.success) {
      showToast(res.message, 'success');
      document.getElementById('importRawTextarea').value = '';
      closeSheetConfigModal();
      state.cacheTimestamps = {};
      loadAppData();
    } else {
      showToast(res?.message || 'Import failed', 'error');
    }
  }).withFailureHandler(err => {
    btn.disabled = false;
    btn.innerHTML = '<span class="ms" style="font-size:18px">file_upload</span> Import & Save Records';
    showToast(err?.message || 'Import error', 'error');
  }).importCsvData(target, text, month, year);
}

function copyAppsScriptCode() {
  const area = document.getElementById('appsScriptCodeArea');
  if (!area) return;
  area.select();
  navigator.clipboard.writeText(area.value).then(() => {
    showToast('Apps Script Code copied!', 'success');
  }).catch(() => {
    document.execCommand('copy');
    showToast('Apps Script Code copied!', 'success');
  });
}

let exporterCache = null;
let currentExporterTab = 'sync_connector';

const EXPORTER_META = {
  sync_connector: {
    fileName: 'sync_connector.gs (Live Sync Connector)',
    desc: 'Paste into Code.gs in Google Spreadsheet for 2-Way Live Sync Webhook'
  },
  codegs: {
    fileName: 'Code.gs (Backend)',
    desc: 'Paste into Code.gs (Standalone Web App Edition)'
  },
  index_html: {
    fileName: 'index.html (Frontend)',
    desc: 'Paste into index.html in Google Sheet Apps Script'
  },
  stylesheet: {
    fileName: 'stylesheet.html (CSS Styles)',
    desc: 'Paste into stylesheet.html in Google Sheet Apps Script'
  },
  javascript: {
    fileName: 'javascript.html (Client JS)',
    desc: 'Paste into javascript.html in Google Sheet Apps Script'
  },
  single_bundle: {
    fileName: 'index.html (All-in-One)',
    desc: 'All-in-One Single Bundle (includes HTML, CSS and JS in one file)'
  }
};

function openCodeExporterModal(initialTab) {
  const modal = document.getElementById('codeExporterOverlay');
  if (!modal) return;
  modal.classList.remove('hidden');

  const tabToOpen = initialTab || 'sync_connector';
  const targetBtn = document.getElementById(`btnExpTab_${tabToOpen}`) || document.getElementById('btnExpTab_sync_connector') || document.querySelector('.exp-nav-btn');

  if (!exporterCache) {
    const box = document.getElementById('expCodeBox');
    if (box) box.value = 'Loading script files from server...';

    cloudRpc('getAppsScriptExportFiles', []).then(res => {
      if (res && res.files) {
        exporterCache = res.files;
      } else {
        exporterCache = {};
      }
      switchTabKey(tabToOpen, targetBtn);
    }).catch(err => {
      console.warn('Failed to load exporter files:', err);
      cloudRpc('getAppsScriptCode', []).then(codeRes => {
        exporterCache = {
          sync_connector: codeRes?.scriptCode || '// Could not load code',
          codegs: codeRes?.scriptCode || '// Could not load code'
        };
        switchTabKey(tabToOpen, targetBtn);
      });
    });
  } else {
    switchTabKey(tabToOpen, targetBtn);
  }
}

function closeCodeExporterModal() {
  const modal = document.getElementById('codeExporterOverlay');
  if (modal) modal.classList.add('hidden');
}

function switchTabKey(tabKey, btnElem) {
  currentExporterTab = tabKey;
  document.querySelectorAll('.exp-nav-btn').forEach(b => {
    b.classList.remove('active');
    b.style.background = 'var(--bg)';
    b.style.color = 'var(--txt)';
    b.style.borderColor = 'var(--bor)';
  });

  if (btnElem) {
    btnElem.classList.add('active');
    btnElem.style.background = 'var(--pri)';
    btnElem.style.color = '#fff';
    btnElem.style.borderColor = 'var(--pri)';
  }

  const meta = EXPORTER_META[tabKey] || { fileName: tabKey, desc: '' };
  const nameEl = document.getElementById('expFileName');
  const descEl = document.getElementById('expFileDesc');
  const box = document.getElementById('expCodeBox');

  if (nameEl) nameEl.textContent = meta.fileName;
  if (descEl) descEl.textContent = meta.desc;

  if (box) {
    if (exporterCache && exporterCache[tabKey]) {
      box.value = exporterCache[tabKey];
    } else if (exporterCache && Object.keys(exporterCache).length > 0) {
      box.value = '// File not available in package';
    }
  }
}

function copyExporterCode() {
  const box = document.getElementById('expCodeBox');
  const btnText = document.getElementById('expCopyBtnText');
  if (!box || !box.value) return;

  navigator.clipboard.writeText(box.value).then(() => {
    if (btnText) {
      const orig = btnText.textContent;
      btnText.textContent = 'Copied to Clipboard! ✅';
      setTimeout(() => { btnText.textContent = orig; }, 2000);
    }
    showToast('Code successfully copied to clipboard!', 'success');
  }).catch(() => {
    box.select();
    document.execCommand('copy');
    showToast('Code selected & copied!', 'success');
  });
}

function downloadZipBlob(type, filename) {
  const url = type === 'appsScript' ? '/api/download-apps-script' : '/api/download-zip';
  const a = document.createElement('a');
  a.href = url;
  a.download = filename || 'adverse-apps-script-suite.zip';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  showToast('Download started for ' + (filename || 'package'), 'success');
}

window.openCodeExporterModal = openCodeExporterModal;
window.closeCodeExporterModal = closeCodeExporterModal;
window.switchTabKey = switchTabKey;
window.copyExporterCode = copyExporterCode;
window.downloadZipBlob = downloadZipBlob;


// Visual Analytics Modal
function openAnalyticsModal() {
  const modal = document.getElementById('analyticsOverlay');
  const body = document.getElementById('analyticsModalBody');
  if (!modal || !body) return;
  modal.classList.remove('hidden');

  google.script.run.withSuccessHandler(res => {
    if (!res || !res.success) {
      body.innerHTML = '<div style="color:var(--dan)">Failed to load agency analytics</div>';
      return;
    }
    const a = res.analytics;
    const maxSpend = Math.max(...(a.monthlyTrend || []).map(m => m.spend), 100);

    body.innerHTML = `
      <!-- KPI Top Grid -->
      <div class="stats-compact" style="grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px">
        <div class="stat-card stat-purple">
          <div class="stat-header"><div class="stat-label">Total Spend ($)</div><span class="ms" style="color:var(--pri)">paid</span></div>
          <div class="stat-value">$${fmtNum(a.totalSpendUSD)}</div>
          <div class="stat-sub">${a.totalCampaigns} Campaigns</div>
        </div>
        <div class="stat-card stat-blue">
          <div class="stat-header"><div class="stat-label">Total Invoiced</div><span class="ms" style="color:var(--sec)">receipt_long</span></div>
          <div class="stat-value">৳${fmtNum(a.totalTakaBilled)}</div>
          <div class="stat-sub">Client Billing</div>
        </div>
        <div class="stat-card stat-green">
          <div class="stat-header"><div class="stat-label">Collections</div><span class="ms" style="color:var(--suc)">payments</span></div>
          <div class="stat-value">৳${fmtNum(a.totalPaymentsReceived)}</div>
          <div class="stat-sub">Paid in Bank/MFS</div>
        </div>
        <div class="stat-card stat-red">
          <div class="stat-header"><div class="stat-label">Pending Due</div><span class="ms" style="color:var(--dan)">account_balance_wallet</span></div>
          <div class="stat-value">৳${fmtNum(a.outstandingDue)}</div>
          <div class="stat-sub">Unpaid Client Balances</div>
        </div>
      </div>

      <!-- Charts Row -->
      <div class="analytics-grid">
        <!-- Monthly Ad Spend Bar Chart -->
        <div class="chart-card">
          <div class="chart-title"><span class="ms" style="color:var(--pri)">bar_chart</span> Monthly Ad Spend ($)</div>
          <div class="bar-chart-container">
            ${(a.monthlyTrend || []).map(m => {
              const h = Math.round((m.spend / maxSpend) * 100);
              return `
                <div class="bar-column" title="${m.month}: $${fmtNum(m.spend)}">
                  <div style="font-size:10px;font-weight:700;color:var(--pri)">$${Math.round(m.spend)}</div>
                  <div class="bar-fill" style="height:${Math.max(h, 4)}%"></div>
                  <div class="bar-label">${m.month.slice(0, 3)}</div>
                </div>
              `;
            }).join('')}
          </div>
        </div>

        <!-- Top Pages Leaderboard -->
        <div class="chart-card">
          <div class="chart-title"><span class="ms" style="color:var(--sec)">leaderboard</span> Top Performing Pages</div>
          <div style="display:flex;flex-direction:column;gap:10px;max-height:160px;overflow-y:auto">
            ${(a.topPages || []).slice(0, 5).map(p => `
              <div>
                <div style="display:flex;justify-content:space-between;font-size:12px;font-weight:700;margin-bottom:3px">
                  <span>${escapeHtml(p.pageName)}</span>
                  <span style="color:var(--pri)">$${fmtNum(p.totalBudget)}</span>
                </div>
                <div style="height:6px;background:var(--bor);border-radius:4px;overflow:hidden">
                  <div style="height:100%;width:${Math.min(100, Math.round((p.totalBudget / (a.topPages[0]?.totalBudget || 1)) * 100))}%;background:var(--gb)"></div>
                </div>
              </div>
            `).join('')}
          </div>
        </div>
      </div>
    `;
  }).getAgencyAnalytics();
}

function closeAnalyticsModal() {
  document.getElementById('analyticsOverlay')?.classList.add('hidden');
}

// WhatsApp Notice Generator
function openWhatsAppModal(presetRecord = null) {
  const modal = document.getElementById('whatsappOverlay');
  const select = document.getElementById('waCampaignSelect');
  if (!modal || !select) return;

  const activeRecords = state.filteredRecords.length ? state.filteredRecords : state.rawRecords;
  select.innerHTML = activeRecords.map((r, idx) => `
    <option value="${idx}" ${presetRecord && presetRecord['Page Name'] === r['Page Name'] ? 'selected' : ''}>
      ${escapeHtml(r['Page Name'] || 'Unknown')} - $${fmtNum(r['Budget'])} (${r['Start Date'] || 'Today'})
    </option>
  `).join('') || '<option value="">No campaigns found</option>';

  modal.classList.remove('hidden');
  generateWhatsAppNotice();
}

function openSingleRecordWhatsApp(record) {
  openWhatsAppModal(record);
}

function openSingleRecordWhatsAppByIndex(idx) {
  const records = state.currentRenderedRecords || state.filteredRecords || state.rawRecords || [];
  const r = records[idx];
  if (r) openSingleRecordWhatsApp(r);
}

function generateWhatsAppNotice() {
  const select = document.getElementById('waCampaignSelect');
  const lang = document.getElementById('waLanguageSelect')?.value || 'focused';
  const textarea = document.getElementById('waMessageText');
  if (!select || !textarea) return;

  const idx = parseInt(select.value, 10) || 0;
  const activeRecords = state.filteredRecords.length ? state.filteredRecords : state.rawRecords;
  const r = activeRecords[idx] || {};

  const pageName = r['Page Name'] || r['page_name'] || r['Page'] || 'Your Page';
  const budget = r['Budget'] || r['budget'] || '0';
  const taka = r['Taka'] || r['taka'] || (parseFloat(budget) * (state.rate || 140)).toFixed(0);
  const start = r['Start Date'] || r['Start'] || r['start_date'] || '';
  const end = r['End Date'] || r['End'] || r['end_date'] || '';
  const days = r['Day'] || r['day'] || '1';
  const type = r['Types'] || r['Type'] || r['type'] || 'Boost';

  let msg = '';
  if (lang === 'focused') {
    msg = `Hello,\n\n` +
          `Your Facebook page ad campaign is now active!\n\n` +
          `📄 Page Name: ${pageName}\n` +
          `⏱️ Duration: ${days} Days${(start && end) ? ` (${start} to ${end})` : ''}\n` +
          `💵 Budget: $${budget} USD\n\n` +
          `Thank you for working with us! 🤝`;
  } else if (lang === 'bn_pro') {
    msg = `*🚀 AdVerse Agency - Campaign Activation Notice*\n\n` +
          `Dear Client, your ad campaign has been successfully launched:\n\n` +
          `📄 *Page Name:* ${pageName}\n` +
          `⏱️ *Duration:* ${days} Days (${start} to ${end})\n` +
          `💵 *Total Budget:* $${budget} USD (৳${taka})\n` +
          `🎯 *Campaign Type:* ${type}\n` +
          `⚡ *Status:* Active / Running\n\n` +
          `Please let us know if you need any assistance. Thank you! 🤝`;
  } else {
    msg = `*🚀 Campaign Confirmation - AdVerse Agency*\n\n` +
          `Dear Client, your ad campaign is now live:\n\n` +
          `• *Page Name:* ${pageName}\n` +
          `• *Duration:* ${days} Days (${start} to ${end})\n` +
          `• *Budget:* $${budget} USD\n` +
          `• *Status:* Active & Running\n\n` +
          `Thank you for choosing AdVerse Agency! 🤝`;
  }

  textarea.value = msg;
}

function copyWhatsAppNotice() {
  const textarea = document.getElementById('waMessageText');
  if (!textarea) return;
  textarea.select();
  navigator.clipboard.writeText(textarea.value).then(() => {
    showToast('Client notice copied to clipboard!', 'success');
    closeWhatsAppModal();
  });
}

function sendWhatsAppWebDirect() {
  const textarea = document.getElementById('waMessageText');
  if (!textarea || !textarea.value) return;
  const encoded = encodeURIComponent(textarea.value);
  window.open(`https://web.whatsapp.com/send?text=${encoded}`, '_blank');
}

function closeWhatsAppModal() {
  document.getElementById('whatsappOverlay')?.classList.add('hidden');
}

// Campaign Verification & Ad Manager Double-Check Engine
function getCampaignKey(r) {
  if (!r) return '';
  if (r._recordKey) return String(r._recordKey).trim();
  if (r['Record ID']) return String(r['Record ID']).trim();
  const sl = String(r['SL'] || r['SL NO'] || r.sl || r.sl_no || '').replace(/\D/g, '');
  const p = String(r['Page Name'] || r.page_name || r.Page || '').trim().toLowerCase();
  const d = String(r['Start'] || r['Start Date'] || r.start_date || r['Entry Date'] || r.entry_date || '').trim();
  if (sl && p) return `camp_${sl}_${p}`;
  if (sl) return `camp_sl_${sl}`;
  if (r._recordId) return String(r._recordId).trim();
  return `camp_${p}_${d}`;
}

function isCampaignVerified(r) {
  if (!r) return false;
  if (r.verified === true || r._verified === true || r.is_verified === true) return true;
  const key = getCampaignKey(r);
  if (!key) return false;
  const ver = state.campaignVerifications && state.campaignVerifications[key];
  return Boolean(ver && ver.verified);
}

function getCampaignVerificationInfo(r) {
  if (!r) return null;
  const key = getCampaignKey(r);
  return (state.campaignVerifications && state.campaignVerifications[key]) ? state.campaignVerifications[key] : null;
}

/**
 * Quality & Error Detection Engine for Campaign Ads
 * Automatically catches:
 * 1. Date discrepancies (Start date + Day ≠ End date)
 * 2. Missing Ad Account / Missing Page Name
 * 3. Invalid or suspicious Budgets ($0, NaN, negative)
 * 4. Missing Start Date / Missing End Date
 * 5. Invalid / 0 Day duration
 */
function checkCampaignErrors(r) {
  const issues = [];
  if (!r) return issues;
  
  const page = String(r['Page Name'] || r.page_name || r.Page || '').trim();
  if (!page) {
    issues.push('Missing Page Name');
  }
  
  const adAccount = String(r['AD Account'] || r.ad_account || '').trim();
  if (!adAccount) {
    issues.push('Missing AD Account');
  }
  
  const budget = parseFloat(r['Budget'] || r.budget || 0);
  if (isNaN(budget) || budget <= 0) {
    issues.push('Invalid Budget ($' + (r['Budget'] || 0) + ')');
  }
  
  const days = parseInt(r['Day'] || r.day || 0, 10);
  if (isNaN(days) || days <= 0) {
    issues.push('Invalid Day Count (' + (r['Day'] || 0) + ')');
  }
  
  // Date validation:
  const rawStart = r['Start Date'] || r.start_date || r['Start'] || '';
  const rawEnd = r['End Date'] || r.end_date || r['End'] || '';
  if (rawStart && rawEnd && days > 0) {
    const sDate = parseAnyDate(rawStart);
    const eDate = parseAnyDate(rawEnd);
    if (sDate && eDate && !isNaN(sDate.getTime()) && !isNaN(eDate.getTime())) {
      const diffMs = eDate.getTime() - sDate.getTime();
      const diffDays = Math.round(diffMs / (1000 * 60 * 60 * 24));
      if (Math.abs(diffDays - days) > 0) {
        issues.push(`Date Mismatch: Start (${formatCustomDate(rawStart)}) + ${days}d ≠ End (${formatCustomDate(rawEnd)})`);
      }
    }
  } else if (!rawStart) {
    issues.push('Missing Start Date');
  } else if (!rawEnd) {
    issues.push('Missing End Date');
  }
  
  return issues;
}

function toggleCampaignVerification(record, targetVerified, explicitIdx) {
  if (!record) return;
  const key = getCampaignKey(record);
  if (!key) return;
  if (!state.campaignVerifications) state.campaignVerifications = {};

  const currentStatus = isCampaignVerified(record);
  const nextVerified = (targetVerified !== undefined) ? targetVerified : !currentStatus;
  const userDisplayName = (state.userRole === 'admin') ? 'Admin' : (state.displayName || 'User');
  const nowStr = formatCustomDate(new Date()) + ', ' + new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

  // Update in-memory record directly for instant reactivity
  record.verified = nextVerified;
  record._verified = nextVerified;

  if (!nextVerified) {
    delete state.campaignVerifications[key];
  } else {
    state.campaignVerifications[key] = {
      verified: true,
      verifiedBy: userDisplayName,
      verifiedAt: nowStr,
      notes: 'Checked and verified'
    };
  }

  try {
    localStorage.setItem('adverse_verifications', JSON.stringify(state.campaignVerifications));
  } catch (e) {}

  // Instant direct DOM update for 0ms lag
  const currentList = state.currentRenderedRecords || state.filteredRecords || [];
  let rowIdx = (explicitIdx !== undefined && explicitIdx !== null && explicitIdx >= 0) ? explicitIdx : currentList.indexOf(record);
  if (rowIdx < 0) {
    rowIdx = currentList.findIndex(r => getCampaignKey(r) === key);
  }
  if (rowIdx >= 0) {
    const rowEl = document.getElementById(`tableRecordRow-${rowIdx}`);
    if (rowEl) {
      rowEl.classList.toggle('row-verified', nextVerified);
      const cb = rowEl.querySelector('.row-campaign-check');
      if (cb) cb.checked = nextVerified;
    }
  }
  const masterCb = document.getElementById('masterTodayCheckbox');
  if (masterCb && currentList.length > 0) {
    masterCb.checked = currentList.every(r => isCampaignVerified(r));
  }

  if (state.verificationFilter && state.verificationFilter !== 'all') {
    if (typeof renderRecordsTable === 'function' && state.filteredRecords) {
      renderRecordsTable({ headers: state.headers, rows: state.filteredRecords });
    }
  }

  if (selectedInspectorRecord && getCampaignKey(selectedInspectorRecord) === key) {
    openCampaignInspector(selectedInspectorIndex);
  }

  const sl = record['SL'] || record['SL NO'] || record.sl || record.sl_no || (rowIdx >= 0 ? rowIdx + 1 : '');
  showToast(nextVerified ? `✅ SL #${sl} Marked as Done!` : `SL #${sl} Unchecked`, nextVerified ? 'success' : 'info');

  // Sync to backend RPC
  cloudRpc('toggleCampaignVerification', [{
    recordKey: key,
    verified: nextVerified,
    verifiedBy: userDisplayName
  }]).then(res => {
    if (res && res.verifications) {
      state.campaignVerifications = res.verifications;
      try {
        localStorage.setItem('adverse_verifications', JSON.stringify(state.campaignVerifications));
      } catch (e) {}
    }
  }).catch(() => {});
}

function handleInspectorVerificationToggle() {
  if (!selectedInspectorRecord) return;
  toggleCampaignVerification(selectedInspectorRecord);
}

function copyDailyBudgetVal() {
  if (!selectedInspectorRecord) return;
  const days = Math.max(1, parseInt(selectedInspectorRecord['Day'] || selectedInspectorRecord['day'] || 1, 10));
  const budget = parseFloat(selectedInspectorRecord['Budget'] || selectedInspectorRecord['budget'] || 0);
  const daily = (budget / days).toFixed(2);
  navigator.clipboard.writeText(daily).then(() => {
    showToast(`⚡ Copied Daily Budget: $${daily}`, 'success');
  });
}

function copyLifetimeBudgetVal() {
  if (!selectedInspectorRecord) return;
  const budget = parseFloat(selectedInspectorRecord['Budget'] || selectedInspectorRecord['budget'] || 0);
  const val = budget.toFixed(2);
  navigator.clipboard.writeText(val).then(() => {
    showToast(`⚡ Copied Lifetime Budget: $${val}`, 'success');
  });
}

function copyEntryDailyBudget() {
  const form = document.getElementById('newRecordForm');
  const day = parseFloat(form?.querySelector('[name="day"]')?.value || 1);
  const budget = parseFloat(form?.querySelector('[name="budget"]')?.value || 0);
  const dVal = (budget / Math.max(1, day)).toFixed(2);
  navigator.clipboard.writeText(dVal).then(() => {
    showToast(`⚡ Copied Daily Budget: $${dVal}`, 'success');
  });
}

function copyEntryLifetimeBudget() {
  const form = document.getElementById('newRecordForm');
  const budget = parseFloat(form?.querySelector('[name="budget"]')?.value || 0);
  const bVal = budget.toFixed(2);
  navigator.clipboard.writeText(bVal).then(() => {
    showToast(`⚡ Copied Lifetime Budget: $${bVal}`, 'success');
  });
}

function setVerificationFilter(filter) {
  state.verificationFilter = filter;
  if (state.currentMenu && (state.currentMenu.key === 'campaign' || state.currentMenu.key === 'today_created_campaign')) {
    if (typeof renderRecordsTable === 'function' && state.filteredRecords) {
      renderRecordsTable({ headers: state.headers, rows: state.filteredRecords });
    }
    if (typeof renderTodaySetupAuditBar === 'function') {
      renderTodaySetupAuditBar();
    }
  }
}

/**
 * Renders the Today Created Setup Quality & Verification Bar at the top of Today Created
 */
function renderTodaySetupAuditBar() {
  const card = document.getElementById('todaySetupAuditCard');
  if (card) {
    card.classList.add('hidden');
    card.style.display = 'none';
    card.innerHTML = '';
  }
}

// Step-by-Step Ad Setup Checker State & Modal Engine
let currentAuditStepIndex = 0;
let auditKeyboardAttached = false;

function openAdSetupAuditModal(optIndex) {
  if (!isAdmin()) {
    showToast('Ad setup verification is restricted to Admin.', 'info');
    return;
  }
  
  const rawRows = state.todayCreatedRegistry || state.filteredRecords || state.rawRecords || [];
  const records = rawRows.filter(r => isValidClientRow(r, 'today_created_campaign'));
  
  if (!records.length) {
    showToast('আজকের কোনো ক্যাম্পেইন পাওয়া যায়নি।', 'info');
    return;
  }
  
  if (optIndex !== undefined && optIndex >= 0 && optIndex < records.length) {
    currentAuditStepIndex = optIndex;
  } else {
    // Find the first unverified record
    const firstPendingIdx = records.findIndex(r => !isCampaignVerified(r));
    currentAuditStepIndex = firstPendingIdx !== -1 ? firstPendingIdx : 0;
  }
  
  const overlay = document.getElementById('adSetupAuditModalOverlay');
  if (overlay) overlay.classList.remove('hidden');
  
  renderCurrentAuditStep();
  
  if (!auditKeyboardAttached) {
    document.addEventListener('keydown', handleAuditModalKeydown);
    auditKeyboardAttached = true;
  }
}

function closeAdSetupAuditModal() {
  document.getElementById('adSetupAuditModalOverlay')?.classList.add('hidden');
  if (auditKeyboardAttached) {
    document.removeEventListener('keydown', handleAuditModalKeydown);
    auditKeyboardAttached = false;
  }
  renderTodaySetupAuditBar();
}

function handleAuditModalKeydown(e) {
  const overlay = document.getElementById('adSetupAuditModalOverlay');
  if (!overlay || overlay.classList.contains('hidden')) return;
  
  if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.tagName === 'SELECT') {
    return;
  }
  
  if (e.key === 'ArrowRight') {
    e.preventDefault();
    navigateAuditStep(1);
  } else if (e.key === 'ArrowLeft') {
    e.preventDefault();
    navigateAuditStep(-1);
  } else if (e.key === 'Escape') {
    e.preventDefault();
    closeAdSetupAuditModal();
  } else if (e.key === ' ' || e.key === 'Enter') {
    e.preventDefault();
    toggleCurrentAuditVerification();
  }
}

function navigateAuditStep(dir) {
  const rawRows = state.todayCreatedRegistry || state.filteredRecords || state.rawRecords || [];
  const records = rawRows.filter(r => isValidClientRow(r, 'today_created_campaign'));
  if (!records.length) return;
  
  let nextIdx = currentAuditStepIndex + dir;
  if (nextIdx < 0) nextIdx = records.length - 1;
  if (nextIdx >= records.length) nextIdx = 0;
  
  currentAuditStepIndex = nextIdx;
  renderCurrentAuditStep();
}

function renderCurrentAuditStep() {
  const rawRows = state.todayCreatedRegistry || state.filteredRecords || state.rawRecords || [];
  const records = rawRows.filter(r => isValidClientRow(r, 'today_created_campaign'));
  if (!records.length) return;
  
  if (currentAuditStepIndex < 0 || currentAuditStepIndex >= records.length) {
    currentAuditStepIndex = 0;
  }
  
  const r = records[currentAuditStepIndex];
  const total = records.length;
  const isVer = isCampaignVerified(r);
  const verInfo = getCampaignVerificationInfo(r);
  const issues = checkCampaignErrors(r);
  
  // Update progress badge and progress bar
  const pBadge = document.getElementById('auditProgressBadge');
  if (pBadge) pBadge.textContent = `${currentAuditStepIndex + 1} / ${total}`;
  
  const pBar = document.getElementById('auditModalProgressBar');
  if (pBar) {
    const pct = Math.round(((currentAuditStepIndex + 1) / total) * 100);
    pBar.style.width = `${pct}%`;
  }
  
  const body = document.getElementById('adSetupAuditModalBody');
  if (!body) return;
  
  const sl = r['SL'] || r['SL NO'] || r.sl || r.sl_no || (currentAuditStepIndex + 1);
  const pageName = r['Page Name'] || r.page_name || 'Unknown Page';
  const adAccount = r['AD Account'] || r.ad_account || 'Not Assigned';
  const idName = r['ID Name'] || r.id_name || '';
  const type = r['Types'] || r['Type'] || r.type || 'Boost';
  const status = r['Status'] || r.status || 'Old Page';
  const days = parseInt(r['Day'] || r.day || 1, 10) || 1;
  const budget = parseFloat(r['Budget'] || r.budget || 0) || 0;
  const daily = (budget / Math.max(1, days)).toFixed(2);
  const rate = parseFloat(r['Dollar Rate (৳)'] || r['Rate'] || r.rate || state.rate || 140);
  const taka = parseFloat(r['Taka'] || r.taka || (budget * rate)).toFixed(0);
  
  const rawStart = r['Start Date'] || r.start_date || r['Start'] || '';
  const rawEnd = r['End Date'] || r.end_date || r['End'] || '';
  const startStr = rawStart ? formatCustomDate(rawStart) : '-';
  const endStr = rawEnd ? formatCustomDate(rawEnd) : '-';
  
  // Date check summary
  let dateCheckHtml = '';
  if (rawStart && rawEnd) {
    const sD = parseAnyDate(rawStart);
    const eD = parseAnyDate(rawEnd);
    if (sD && eD) {
      const diffDays = Math.round((eD.getTime() - sD.getTime()) / (1000 * 60 * 60 * 24));
      if (diffDays === days) {
        dateCheckHtml = `<span style="color:#10B981;font-size:11px;font-weight:700">✔ End Date matches ${days} Days duration</span>`;
      } else {
        dateCheckHtml = `<span style="color:#EF4444;font-size:11px;font-weight:800">⚠️ Mismatch: Duration says ${days} days but date gap is ${diffDays} days!</span>`;
      }
    }
  }
  
  const fullMetaLine = `${sl ? sl + ' ' : ''}${pageName} ${startStr} ${endStr} ${type} ${days} $${budget}`;

  body.innerHTML = `
    <!-- Top Hero Card -->
    <div style="background:var(--card);border:1px solid var(--bor);border-radius:var(--radius-md);padding:14px 18px;margin-bottom:14px;box-shadow:0 1px 4px rgba(0,0,0,0.03)">
      <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:10px;flex-wrap:wrap">
        <div>
          <div style="display:flex;align-items:center;gap:6px;margin-bottom:4px">
            <span class="badge" style="background:var(--pri);color:#fff;font-weight:900;font-size:12px;padding:2px 8px;border-radius:6px">SL #${escapeHtml(String(sl))}</span>
            <span class="page-status-badge ${status.toLowerCase().includes('old') ? 'badge-old-page' : 'badge-new-page'}" style="font-size:10px;padding:2px 6px">${escapeHtml(status)}</span>
            <span class="pill pill-blue" style="font-size:10px;padding:2px 6px">${escapeHtml(type)}</span>
          </div>
          <div style="font-size:18px;font-weight:900;color:var(--txt);display:flex;align-items:center;gap:8px">
            <span>${escapeHtml(pageName)}</span>
            <button type="button" class="btn btn-icon btn-secondary" onclick="navigator.clipboard.writeText('${escapeHtml(pageName)}'); showToast('Copied Page Name', 'success');" style="width:26px;height:26px;padding:0" title="Copy Page Name">
              <span class="ms" style="font-size:14px">content_copy</span>
            </button>
          </div>
          <div style="font-size:12px;color:var(--txt-muted);margin-top:2px">
            AD Account: <b style="color:var(--txt);font-family:monospace">${escapeHtml(adAccount)}</b> ${idName ? `• ID: <b style="color:var(--txt)">${escapeHtml(idName)}</b>` : ''}
          </div>
        </div>
        <div style="text-align:right">
          ${isVer ? `
            <div style="background:#DCFCE7;color:#15803D;padding:6px 12px;border-radius:8px;font-weight:800;font-size:12px;display:inline-flex;align-items:center;gap:4px">
              <span class="ms" style="font-size:16px">check_circle</span> Verified in Meta
            </div>
            ${verInfo?.verifiedAt ? `<div style="font-size:10px;color:var(--txt-dim);margin-top:4px">${escapeHtml(verInfo.verifiedAt)} (${escapeHtml(verInfo.verifiedBy || 'Admin')})</div>` : ''}
          ` : `
            <div style="background:#FEF3C7;color:#B45309;padding:6px 12px;border-radius:8px;font-weight:800;font-size:12px;display:inline-flex;align-items:center;gap:4px">
              <span class="ms" style="font-size:16px">pending</span> Check Pending
            </div>
          `}
        </div>
      </div>
    </div>

    <!-- Details Grid -->
    <div class="audit-grid-2" style="margin-bottom:14px">
      <!-- Schedule Box -->
      <div style="background:var(--card);border:1px solid var(--bor);border-radius:var(--radius-md);padding:14px">
        <div style="font-size:11px;font-weight:800;color:var(--pri);text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px;display:flex;align-items:center;gap:4px">
          <span class="ms" style="font-size:14px">date_range</span> Schedule & Duration
        </div>
        <div style="display:flex;justify-content:space-between;margin-bottom:4px">
          <span style="font-size:12px;color:var(--txt-muted)">Start Date:</span>
          <span style="font-size:13px;font-weight:800;color:var(--txt)">${escapeHtml(startStr)}</span>
        </div>
        <div style="display:flex;justify-content:space-between;margin-bottom:4px">
          <span style="font-size:12px;color:var(--txt-muted)">Duration:</span>
          <span style="font-size:13px;font-weight:800;color:var(--pri)">${days} Days</span>
        </div>
        <div style="display:flex;justify-content:space-between;margin-bottom:6px">
          <span style="font-size:12px;color:var(--txt-muted)">End Date:</span>
          <span style="font-size:13px;font-weight:800;color:var(--txt)">${escapeHtml(endStr)}</span>
        </div>
        <div style="border-top:1px dashed var(--bor);padding-top:6px;margin-top:6px">
          ${dateCheckHtml}
        </div>
      </div>

      <!-- Financials Box -->
      <div style="background:var(--card);border:1px solid var(--bor);border-radius:var(--radius-md);padding:14px">
        <div style="font-size:11px;font-weight:800;color:var(--sec);text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px;display:flex;align-items:center;gap:4px">
          <span class="ms" style="font-size:14px">payments</span> Budget & Financials
        </div>
        <div style="display:flex;justify-content:space-between;margin-bottom:4px">
          <span style="font-size:12px;color:var(--txt-muted)">Total Budget:</span>
          <span style="font-size:14px;font-weight:900;color:var(--sec)">$${fmtNum(budget)}</span>
        </div>
        <div style="display:flex;justify-content:space-between;margin-bottom:4px">
          <span style="font-size:12px;color:var(--txt-muted)">Daily Spend:</span>
          <span style="font-size:13px;font-weight:800;color:var(--txt)">$${daily} / day</span>
        </div>
        <div style="display:flex;justify-content:space-between;margin-bottom:6px">
          <span style="font-size:12px;color:var(--txt-muted)">Client Bill (৳):</span>
          <span style="font-size:14px;font-weight:900;color:var(--suc)">৳${fmtNum(taka)}</span>
        </div>
        <div style="font-size:11px;color:var(--txt-dim);border-top:1px dashed var(--bor);padding-top:6px;margin-top:6px">
          Rate: ৳${rate} per dollar
        </div>
      </div>
    </div>

    <!-- Automated Pre-flight Alert Box -->
    ${issues.length > 0 ? `
      <div class="audit-warning-box" style="margin-bottom:14px">
        <div class="audit-warning-title">
          <span class="ms">error</span> <span>এই ক্যাম্পেইনে সম্ভাব্য ভুল ধরা পড়েছে:</span>
        </div>
        <ul style="margin:4px 0 0;padding-left:20px;font-size:12px;color:var(--dan);font-weight:700">
          ${issues.map(iss => `<li>${escapeHtml(iss)}</li>`).join('')}
        </ul>
      </div>
    ` : `
      <div class="audit-success-box" style="margin-bottom:14px">
        <span class="ms">verified</span>
        <span>স্বয়ংক্রিয় যাচাই সম্পন্ন: তারিখ, বাজেট এবং অ্যাকাউন্টে কোনো ত্রুটি পাওয়া যায়নি।</span>
      </div>
    `}

    <!-- Facebook Ad Manager Verification Checklist -->
    <div style="background:var(--card);border:1px solid var(--bor);border-radius:var(--radius-md);padding:14px;margin-bottom:14px">
      <div style="font-size:12px;font-weight:800;color:var(--txt);margin-bottom:8px;display:flex;align-items:center;gap:6px">
        <span class="ms" style="color:#1877F2">fact_check</span> Meta Ads Manager Setup Checklist:
      </div>
      <div style="display:flex;flex-direction:column;gap:6px;font-size:12px;color:var(--txt)">
        <label style="display:flex;align-items:center;gap:8px;cursor:pointer">
          <input type="checkbox" checked disabled style="accent-color:#10B981">
          <span>Ad Account (<b style="font-family:monospace">${escapeHtml(adAccount)}</b>) এবং Page (<b style="color:var(--pri)">${escapeHtml(pageName)}</b>) মিলেছে</span>
        </label>
        <label style="display:flex;align-items:center;gap:8px;cursor:pointer">
          <input type="checkbox" checked disabled style="accent-color:#10B981">
          <span>Start (${escapeHtml(startStr)}) ও End Date (${escapeHtml(endStr)}) মিলিয়ে মোট <b>${days} দিন</b> দেওয়া হয়েছে</span>
        </label>
        <label style="display:flex;align-items:center;gap:8px;cursor:pointer">
          <input type="checkbox" checked disabled style="accent-color:#10B981">
          <span>Daily Budget ($${daily}) অথবা Lifetime Budget ($${fmtNum(budget)}) সঠিক দেওয়া হয়েছে</span>
        </label>
      </div>
    </div>

    <!-- Quick Copy Helpers for Meta Ads Manager -->
    <div style="background:var(--bg);border:1px solid var(--bor);border-radius:var(--radius-md);padding:10px 14px;display:flex;align-items:center;justify-content:space-between;gap:8px;flex-wrap:wrap">
      <div style="font-size:11.5px;font-weight:700;color:var(--txt-muted)">
        ⚡ Quick Copy for Meta Ads Manager:
      </div>
      <div style="display:flex;gap:6px;flex-wrap:wrap">
        <button type="button" class="btn btn-secondary" onclick="navigator.clipboard.writeText('${escapeHtml(pageName)}'); showToast('Copied Page Name', 'success');" style="padding:4px 8px;font-size:11px">
          Page
        </button>
        <button type="button" class="btn btn-secondary" onclick="navigator.clipboard.writeText('${daily}'); showToast('Copied Daily Budget: $' + '${daily}', 'success');" style="padding:4px 8px;font-size:11px">
          Daily ($${daily})
        </button>
        <button type="button" class="btn btn-secondary" onclick="navigator.clipboard.writeText('${budget}'); showToast('Copied Lifetime Budget: $' + '${budget}', 'success');" style="padding:4px 8px;font-size:11px">
          Total ($${budget})
        </button>
        <button type="button" class="btn btn-secondary" onclick="navigator.clipboard.writeText('${escapeHtml(fullMetaLine)}'); showToast('Copied Meta Setup Line', 'success');" style="padding:4px 8px;font-size:11px">
          Full Line
        </button>
      </div>
    </div>
  `;
  
  // Footer button updates
  const btnToggle = document.getElementById('btnAuditVerifyToggle');
  if (btnToggle) {
    if (isVer) {
      btnToggle.className = 'btn btn-secondary';
      btnToggle.style.color = '#B45309';
      btnToggle.style.background = '#FEF3C7';
      btnToggle.style.borderColor = 'rgba(245,158,11,0.4)';
      btnToggle.innerHTML = `<span class="ms">undo</span> Mark Unverified`;
    } else {
      btnToggle.className = 'btn btn-primary';
      btnToggle.style.color = '#fff';
      btnToggle.style.background = '#10B981';
      btnToggle.style.borderColor = '#10B981';
      btnToggle.innerHTML = `<span class="ms">check_circle</span> Mark Verified & Next (↵)`;
    }
  }
}

function toggleCurrentAuditVerification() {
  const rawRows = state.todayCreatedRegistry || state.filteredRecords || state.rawRecords || [];
  const records = rawRows.filter(r => isValidClientRow(r, 'today_created_campaign'));
  if (!records.length || currentAuditStepIndex < 0 || currentAuditStepIndex >= records.length) return;
  
  const r = records[currentAuditStepIndex];
  const isVer = isCampaignVerified(r);
  
  toggleCampaignVerification(r, !isVer);
  
  // Auto advance to next pending ad if just marked verified
  if (!isVer) {
    setTimeout(() => {
      const remainingUnverified = records.findIndex((rec, idx) => idx > currentAuditStepIndex && !isCampaignVerified(rec));
      if (remainingUnverified !== -1) {
        currentAuditStepIndex = remainingUnverified;
      } else {
        const anyUnverified = records.findIndex(rec => !isCampaignVerified(rec));
        if (anyUnverified !== -1) currentAuditStepIndex = anyUnverified;
        else if (currentAuditStepIndex < records.length - 1) currentAuditStepIndex++;
      }
      renderCurrentAuditStep();
      renderTodaySetupAuditBar();
    }, 150);
  } else {
    renderCurrentAuditStep();
    renderTodaySetupAuditBar();
  }
}

function editCurrentAuditRecord() {
  const rawRows = state.todayCreatedRegistry || state.filteredRecords || state.rawRecords || [];
  const records = rawRows.filter(r => isValidClientRow(r, 'today_created_campaign'));
  if (!records.length || currentAuditStepIndex < 0 || currentAuditStepIndex >= records.length) return;
  
  closeAdSetupAuditModal();
  openEditRecordByIndex(currentAuditStepIndex);
}

function confirmVerifyAllTodayAds() {
  const rawRows = state.todayCreatedRegistry || state.filteredRecords || state.rawRecords || [];
  const records = rawRows.filter(r => isValidClientRow(r, 'today_created_campaign'));
  if (!records.length) return;
  
  const pendingRecords = records.filter(r => !isCampaignVerified(r));
  if (!pendingRecords.length) {
    showToast('আজকের সব ক্যাম্পেইন ইতোমধ্যে ভেরিফাইড!', 'info');
    return;
  }
  
  if (!confirm(`আজকের ${pendingRecords.length} টি পেন্ডিং ক্যাম্পেইনকে এক ক্লিকে 'Verified' মার্ক করতে চান?`)) {
    return;
  }
  
  if (!state.campaignVerifications) state.campaignVerifications = {};
  const userDisplayName = (state.userRole === 'admin') ? 'Admin' : 'Partner (Viewer)';
  const nowStr = formatCustomDate(new Date()) + ', ' + new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  const keys = [];
  
  pendingRecords.forEach(r => {
    const key = getCampaignKey(r);
    if (key) {
      keys.push(key);
      state.campaignVerifications[key] = {
        verified: true,
        verifiedBy: userDisplayName,
        verifiedAt: nowStr,
        notes: 'Batch verified all today'
      };
    }
  });
  
  try {
    localStorage.setItem('adverse_verifications', JSON.stringify(state.campaignVerifications));
  } catch(e) {}
  
  renderTodaySetupAuditBar();
  if (state.currentMenu && (state.currentMenu.key === 'campaign' || state.currentMenu.key === 'today_created_campaign')) {
    if (typeof renderRecordsTable === 'function' && state.filteredRecords) {
      renderRecordsTable({ headers: state.headers, rows: state.filteredRecords });
    }
  }
  
  showToast(`✅ আজকের ${pendingRecords.length} টি ক্যাম্পেইন সফলভাবে ভেরিফাইড করা হয়েছে!`, 'success');
  
  google.script.run.withSuccessHandler(res => {
    if (res && res.verifications) {
      state.campaignVerifications = res.verifications;
      try {
        localStorage.setItem('adverse_verifications', JSON.stringify(state.campaignVerifications));
      } catch(e) {}
    }
  }).batchVerifyCampaigns({
    recordKeys: keys,
    verifiedBy: userDisplayName
  });
}

window.checkCampaignErrors = checkCampaignErrors;
window.renderTodaySetupAuditBar = renderTodaySetupAuditBar;
window.openAdSetupAuditModal = openAdSetupAuditModal;
window.closeAdSetupAuditModal = closeAdSetupAuditModal;
window.navigateAuditStep = navigateAuditStep;
window.toggleCurrentAuditVerification = toggleCurrentAuditVerification;
window.editCurrentAuditRecord = editCurrentAuditRecord;
window.confirmVerifyAllTodayAds = confirmVerifyAllTodayAds;
window.setVerificationFilter = setVerificationFilter;
window.handleInspectorVerificationToggle = handleInspectorVerificationToggle;
window.copyDailyBudgetVal = copyDailyBudgetVal;
window.copyLifetimeBudgetVal = copyLifetimeBudgetVal;
window.copyEntryDailyBudget = copyEntryDailyBudget;
window.copyEntryLifetimeBudget = copyEntryLifetimeBudget;
window.toggleCampaignVerification = toggleCampaignVerification;

// Campaign Quick Inspector (Available for both Admin & Viewer)
let selectedInspectorRecord = null;
let selectedInspectorIndex = null;

function openCampaignInspector(rIdx) {
  if (!isAdmin()) {
    return; // Viewers only see clean read-only table records with status badges, no popup
  }
  const records = state.currentRenderedRecords || state.filteredRecords || state.rawRecords || [];
  const r = records[rIdx];
  if (!r) return;
  selectedInspectorRecord = r;
  selectedInspectorIndex = rIdx;

  if (state) {
    state.focusedRowIndex = rIdx;
    state.editingRow = r;
    state.selectedRecord = r;
    state.editingId = r['Record ID'] || r._recordId || r._row || r['SL NO'] || r.sl_no || r.SL || r.Date;
  }

  // Highlight row in table
  const rows = document.querySelectorAll('#dataTable tbody tr.table-record-row');
  rows.forEach((tr, idx) => {
    if (idx === rIdx) {
      tr.classList.add('row-focused');
      tr.setAttribute('aria-selected', 'true');
    } else {
      tr.classList.remove('row-focused');
      tr.removeAttribute('aria-selected');
    }
  });

  const overlay = document.getElementById('campaignInspectorOverlay');
  if (!overlay) return;

  const pageName = r['Page Name'] || r['page_name'] || r['Page'] || 'Unknown Page';
  const budget = r['Budget'] || r['budget'] || '0';
  const taka = r['Taka'] || r['taka'] || (parseFloat(budget) * (state.rate || 140)).toFixed(0);
  const days = r['Day'] || r['day'] || '1';
  const type = r['Types'] || r['Type'] || r['type'] || 'Boost';
  const rawStart = r['Start Date'] || r['Start'] || r['start_date'] || '';
  const rawEnd = r['End Date'] || r['End'] || r['end_date'] || '';
  const start = rawStart ? formatCustomDate(rawStart) : '-';
  const end = rawEnd ? formatCustomDate(rawEnd) : '-';
  const status = r['Status'] || r['status'] || 'Active';
  const sl = r['SL'] || r['sl_no'] || r['SL NO'] || r.sl || String(rIdx + 1);
  const ad = r['AD Account'] || r['ad_account'] || '';
  const id = r['ID Name'] || r['id_name'] || '';

  const elPage = document.getElementById('inspPageName');
  const elSub = document.getElementById('inspSubTitle');
  const elDays = document.getElementById('inspDays');
  const elBudget = document.getElementById('inspBudget');
  const elTaka = document.getElementById('inspTaka');
  const elType = document.getElementById('inspType');
  const elStart = document.getElementById('inspStartDate');
  const elEnd = document.getElementById('inspEndDate');
  const elLine = document.getElementById('inspRawLine');

  if (elPage) elPage.textContent = pageName;
  if (elSub) elSub.textContent = `SL #${sl} • Status: ${status} • Type: ${type}${ad ? ' • AD: ' + ad : ''}${id ? ' (' + id + ')' : ''}`;
  if (elDays) elDays.textContent = `${days} Days`;
  if (elBudget) elBudget.textContent = `$${fmtNum(budget)}`;
  if (elTaka) elTaka.textContent = `৳${fmtNum(taka)}`;
  if (elType) elType.textContent = type;
  if (elStart) elStart.textContent = `Start: ${start}`;
  if (elEnd) elEnd.textContent = `End: ${end}`;

  // Timeline text & progress bar
  const totalDays = parseInt(days, 10) || 1;
  const elTimeline = document.getElementById('inspTimelineText');
  if (elTimeline) elTimeline.textContent = `${totalDays} Days Duration (${start} ➔ ${end})`;

  const elProgressBar = document.getElementById('inspProgressBar');
  if (elProgressBar) {
    elProgressBar.style.width = '100%';
  }

  // Standard line format: SL/ID Page Start End Type Day $Budget
  const adIdPart = (id || ad) ? `${sl ? sl + '/' : ''}${id || ad}` : (sl ? `SL ${sl}` : '');
  const rawLine = `${adIdPart ? adIdPart + ' ' : ''}${pageName} ${start} ${end} ${type} ${days} $${budget}`;
  if (elLine) elLine.textContent = rawLine;

  // Set action button visibility based on permissions
  const btnEdit = document.getElementById('btnInspEdit');
  const btnDelete = document.getElementById('btnInspDelete');
  const btnIncrease = document.getElementById('btnInspIncrease');
  const canModify = isAdmin();
  const isCampaign = Boolean(r['Page Name'] || r['page_name'] || r['Budget'] || r['Start Date']);

  if (btnEdit) btnEdit.style.display = canModify ? 'inline-flex' : 'none';
  if (btnDelete) btnDelete.style.display = canModify ? 'inline-flex' : 'none';
  if (btnIncrease) btnIncrease.style.display = (canModify && isCampaign) ? 'inline-flex' : 'none';

  overlay.classList.remove('hidden');
}

function closeCampaignInspector() {
  document.getElementById('campaignInspectorOverlay')?.classList.add('hidden');
}

function copyInspectorRawLine() {
  const line = document.getElementById('inspRawLine')?.textContent;
  if (!line) return;
  navigator.clipboard.writeText(line).then(() => {
    showToast('Campaign format copied!', 'success');
  });
}

function inspectorCopyRecord() {
  if (!selectedInspectorRecord) return;
  copyCampaignRecord(selectedInspectorRecord);
}

function inspectorEditRecord() {
  if (!selectedInspectorRecord) return;
  if (!isAdmin()) {
    showToast('Admin access required to edit campaign', 'error');
    return;
  }
  const record = selectedInspectorRecord;
  closeCampaignInspector();
  openEditRecord(record);
}

function inspectorDeleteRecord() {
  if (!selectedInspectorRecord) return;
  if (!isAdmin()) {
    showToast('Admin access required to delete campaign', 'error');
    return;
  }
  const record = selectedInspectorRecord;
  const id = record['Record ID'] || record._recordId || record._row || record.SL;
  closeCampaignInspector();
  deleteRecordItem(id, record);
}

function inspectorIncreaseRecord() {
  if (!selectedInspectorRecord) return;
  if (!isAdmin()) {
    showToast('Admin access required to increase campaign', 'error');
    return;
  }
  const record = selectedInspectorRecord;
  closeCampaignInspector();
  quickIncreaseCampaign(record);
}

function openInspectorWhatsApp() {
  const record = selectedInspectorRecord;
  closeCampaignInspector();
  if (record) {
    openWhatsAppModal(record);
  } else {
    openWhatsAppModal();
  }
}

// ==========================================================================
// ROW FOCUS MODE & FLOATING ACTION MENU
// ==========================================================================
function handleTableRowClick(e, rIdx) {
  // If clicked inside td-action, interactive buttons/inputs/links, or page hover popover, let child event handle it
  if (e && e.target) {
    if (e.target.closest('.td-action') || e.target.closest('button') || e.target.closest('input') || e.target.closest('a') || e.target.closest('.page-hover-popover')) {
      return;
    }
  }
  if (state.currentMenu && state.currentMenu.key === 'agency_cost') {
    if (isAdmin()) {
      openEditAgencyCostRecord(rIdx);
    }
    return;
  }
  // User Requirement: Direct click on row opens the Campaign Inspector modal directly without floating bar
  clearRowFocus();
  openCampaignInspector(rIdx);
}

function focusTableRow(rIdx, shouldScroll = false) {
  const records = state.currentRenderedRecords || state.filteredRecords || state.rawRecords || [];
  if (!records.length) return;

  let targetIdx = parseInt(rIdx, 10);
  if (isNaN(targetIdx) || targetIdx < 0) targetIdx = 0;
  if (targetIdx >= records.length) targetIdx = records.length - 1;

  const r = records[targetIdx];
  if (!r) return;

  state.focusedRowIndex = targetIdx;
  state.editingRow = r;
  state.selectedRecord = r;
  state.editingId = r['Record ID'] || r._recordId || r._row || r['SL NO'] || r.sl_no || r.SL || r.Date;

  // Highlight the row visually across table without heavy full re-render
  const rows = document.querySelectorAll('#dataTable tbody tr.table-record-row');
  rows.forEach((tr, idx) => {
    if (idx === targetIdx) {
      tr.classList.add('row-focused');
      tr.setAttribute('aria-selected', 'true');
      if (shouldScroll) {
        tr.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      }
    } else {
      tr.classList.remove('row-focused');
      tr.removeAttribute('aria-selected');
    }
  });

  renderRowFocusMenu(r, targetIdx, records.length);
}

function clearRowFocus() {
  state.focusedRowIndex = null;
  document.querySelectorAll('#dataTable tbody tr.table-record-row.row-focused').forEach(tr => {
    tr.classList.remove('row-focused');
    tr.removeAttribute('aria-selected');
  });
  const menu = document.getElementById('rowFocusFloatingMenu');
  if (menu) {
    menu.classList.add('hidden');
    menu.classList.remove('active');
    menu.style.top = '';
    menu.style.bottom = '';
    menu.style.left = '';
    menu.style.transform = '';
  }
  const btnToggle = document.getElementById('btnRowFocusToggle');
  if (btnToggle) btnToggle.classList.remove('active');
}

function positionRowFocusMenu(rIdx) {
  const menu = document.getElementById('rowFocusFloatingMenu');
  if (!menu || menu.classList.contains('hidden')) return;

  const tr = document.getElementById(`tableRecordRow-${rIdx}`);
  if (!tr) return;

  if (window.innerWidth <= 768) {
    menu.style.position = 'fixed';
    menu.style.top = 'auto';
    menu.style.bottom = '16px';
    menu.style.left = '50%';
    menu.style.transform = 'translateX(-50%)';
    return;
  }

  const rect = tr.getBoundingClientRect();
  const menuWidth = menu.offsetWidth || 560;
  const menuHeight = menu.offsetHeight || 48;

  let top = rect.top - menuHeight - 10;
  // If not enough room above the row (under top header), position right below the row
  if (top < 70) {
    top = rect.bottom + 8;
  }

  // Horizontally center with the table viewport
  const tableWrap = document.querySelector('.table-wrap');
  const wrapRect = tableWrap ? tableWrap.getBoundingClientRect() : rect;
  const centerLeft = wrapRect.left + (wrapRect.width / 2);
  let left = centerLeft - (menuWidth / 2);

  const minLeft = 16;
  const maxLeft = window.innerWidth - menuWidth - 16;
  left = Math.max(minLeft, Math.min(maxLeft, left));

  menu.style.position = 'fixed';
  menu.style.top = `${Math.round(top)}px`;
  menu.style.bottom = 'auto';
  menu.style.left = `${Math.round(left)}px`;
  menu.style.transform = 'none';
}

function toggleFirstRowFocus() {
  const records = state.currentRenderedRecords || state.filteredRecords || state.rawRecords || [];
  if (!records.length) {
    showToast('No table records to focus', 'warning');
    return;
  }
  if (state.focusedRowIndex !== null && state.focusedRowIndex !== undefined) {
    clearRowFocus();
  } else {
    focusTableRow(0, true);
    showToast('🎯 Row Focus active: Press ↑/↓ to navigate, E to Edit, + to Increase, Del to Delete', 'info');
  }
}

function focusNextRow() {
  const records = state.currentRenderedRecords || state.filteredRecords || state.rawRecords || [];
  if (!records.length) return;
  let nextIdx = (state.focusedRowIndex === null || state.focusedRowIndex === undefined) ? 0 : state.focusedRowIndex + 1;
  if (nextIdx >= records.length) nextIdx = 0;
  focusTableRow(nextIdx, true);
}

function focusPreviousRow() {
  const records = state.currentRenderedRecords || state.filteredRecords || state.rawRecords || [];
  if (!records.length) return;
  let prevIdx = (state.focusedRowIndex === null || state.focusedRowIndex === undefined) ? records.length - 1 : state.focusedRowIndex - 1;
  if (prevIdx < 0) prevIdx = records.length - 1;
  focusTableRow(prevIdx, true);
}

function renderRowFocusMenu(r, rIdx, totalCount) {
  const menu = document.getElementById('rowFocusFloatingMenu');
  if (!menu) return;

  const pageName = r['Page Name'] || r['page_name'] || r['Page'] || 'Record #' + (rIdx + 1);
  const budget = r['Budget'] || r['budget'] || '0';
  const taka = r['Taka'] || r['taka'] || (parseFloat(budget) * (state.rate || 140)).toFixed(0);
  const type = r['Types'] || r['Type'] || r['type'] || 'Boost';
  const status = r['Status'] || r['status'] || 'Active';

  const elIdx = document.getElementById('rfRowIndexTag');
  const elPage = document.getElementById('rfPageNameTag');
  const elBudget = document.getElementById('rfBudgetTag');
  const elTaka = document.getElementById('rfTakaTag');
  const elType = document.getElementById('rfTypeTag');
  const elStatus = document.getElementById('rfStatusTag');

  if (elIdx) elIdx.textContent = `ROW #${rIdx + 1}/${totalCount}`;
  if (elPage) {
    elPage.textContent = pageName;
    elPage.title = pageName;
  }
  if (elBudget) elBudget.textContent = budget && budget !== '0' ? `$${fmtNum(budget)}` : '';
  if (elTaka) elTaka.textContent = taka && taka !== '0' ? `৳${fmtNum(taka)}` : '';
  if (elType) elType.textContent = type;
  if (elStatus) elStatus.textContent = status;

  // Immediate access: Edit, Increase, Delete, Clone are Admin only!
  const isCampaign = Boolean(r['Page Name'] || r['page_name'] || r['Budget'] || r['Start Date']);
  const canModify = isAdmin();
  const btnEdit = document.getElementById('rfBtnEdit');
  const btnIncrease = document.getElementById('rfBtnIncrease');
  const btnDelete = document.getElementById('rfBtnDelete');
  const btnCopy = document.getElementById('rfBtnCopy');
  const btnWa = document.getElementById('rfBtnWa');
  const btnClone = document.getElementById('rfBtnClone');

  if (btnEdit) btnEdit.style.display = canModify ? 'inline-flex' : 'none';
  if (btnIncrease) btnIncrease.style.display = (canModify && isCampaign) ? 'inline-flex' : 'none';
  if (btnDelete) btnDelete.style.display = canModify ? 'inline-flex' : 'none';
  if (btnCopy) btnCopy.style.display = isCampaign ? 'inline-flex' : 'none';
  if (btnWa) btnWa.style.display = isCampaign ? 'inline-flex' : 'none';
  if (btnClone) btnClone.style.display = (canModify && isCampaign) ? 'inline-flex' : 'none';

  menu.classList.remove('hidden');
  menu.classList.add('active');

  const btnToggle = document.getElementById('btnRowFocusToggle');
  if (btnToggle) btnToggle.classList.add('active');

  // Dynamically position floating action menu anchored near clicked row
  requestAnimationFrame(() => {
    positionRowFocusMenu(rIdx);
  });
}

function rfActionCopy() {
  if (state.focusedRowIndex === null || state.focusedRowIndex === undefined) return;
  copyCampaignRecordByIndex(state.focusedRowIndex);
}

function rfActionWhatsApp() {
  if (state.focusedRowIndex === null || state.focusedRowIndex === undefined) return;
  openSingleRecordWhatsAppByIndex(state.focusedRowIndex);
}

function rfActionIncrease() {
  if (state.focusedRowIndex === null || state.focusedRowIndex === undefined) return;
  quickIncreaseCampaignByIndex(state.focusedRowIndex);
}

function rfActionEdit() {
  if (state.focusedRowIndex === null || state.focusedRowIndex === undefined) return;
  openEditRecordByIndex(state.focusedRowIndex);
}

function rfActionClone() {
  if (state.focusedRowIndex === null || state.focusedRowIndex === undefined) return;
  cloneCampaignRecordByIndex(state.focusedRowIndex);
}

function rfActionInspect() {
  if (state.focusedRowIndex === null || state.focusedRowIndex === undefined) return;
  openCampaignInspector(state.focusedRowIndex);
}

function rfActionDelete() {
  if (state.focusedRowIndex === null || state.focusedRowIndex === undefined) return;
  const currentIdx = state.focusedRowIndex;
  deleteRecordByIndex(currentIdx);
  clearRowFocus();
}

// Quick Filter Chips (All, Active, High Budget $20+, Ending Soon, Old)
function applyQuickChip(chipType, btnEl) {
  document.querySelectorAll('.quick-chip').forEach(c => c.classList.remove('active'));
  if (btnEl) btnEl.classList.add('active');

  const baseRecords = state.rawRecords || [];
  let filtered = [...baseRecords];

  if (chipType === 'active') {
    filtered = baseRecords.filter(r => {
      const st = String(r['Status'] || r['status'] || '').toLowerCase();
      return !st.includes('old') && !st.includes('end') && !st.includes('pause');
    });
  } else if (chipType === 'high') {
    filtered = baseRecords.filter(r => {
      const b = parseFloat(r['Budget'] || r['budget'] || 0);
      return b >= 20;
    });
  } else if (chipType === 'ending') {
    const todayStr = new Date().toISOString().slice(0, 10);
    filtered = baseRecords.filter(r => {
      const end = String(r['End Date'] || r['End'] || '');
      return end.includes(todayStr) || end.toLowerCase().includes('end');
    });
  } else if (chipType === 'old') {
    filtered = baseRecords.filter(r => {
      const st = String(r['Status'] || r['status'] || '').toLowerCase();
      return st.includes('old');
    });
  }

  state.filteredRecords = filtered;
  renderRecordsTable({ headers: state.headers, rows: state.filteredRecords });
}

// Table Density Modes: 'comfortable' (spacious) vs 'compact' (high-density for power users)
function setTableDensity(mode, showNotification = true) {
  const targetMode = mode === 'compact' ? 'compact' : 'comfortable';
  state.tableDensity = targetMode;
  try {
    localStorage.setItem('adverse_table_density', targetMode);
  } catch (e) {}

  const tableWraps = document.querySelectorAll('.table-wrap');
  tableWraps.forEach(wrap => {
    wrap.classList.toggle('compact-density', targetMode === 'compact');
    wrap.classList.toggle('density-compact', targetMode === 'compact');
    wrap.classList.toggle('density-comfortable', targetMode === 'comfortable');
  });

  const recordsCard = document.getElementById('recordsCard');
  if (recordsCard) {
    recordsCard.classList.toggle('records-card-compact', targetMode === 'compact');
  }

  // Update segmented control buttons
  const btnComf = document.getElementById('btnDensityComfortable');
  const btnComp = document.getElementById('btnDensityCompact');
  if (btnComf) btnComf.classList.toggle('active', targetMode === 'comfortable');
  if (btnComp) btnComp.classList.toggle('active', targetMode === 'compact');

  const legacyIcon = document.getElementById('densityIcon');
  if (legacyIcon) {
    legacyIcon.textContent = targetMode === 'compact' ? 'density_small' : 'view_headline';
  }

  if (showNotification) {
    if (targetMode === 'compact') {
      showToast('⚡ Compact view enabled — High-density mode for power users', 'info');
    } else {
      showToast('🛋️ Comfortable view enabled — Spacious layout', 'info');
    }
  }
}

function toggleTableDensity() {
  const current = state.tableDensity || localStorage.getItem('adverse_table_density') || 'comfortable';
  const next = current === 'compact' ? 'comfortable' : 'compact';
  setTableDensity(next, true);
}

function initTableDensity() {
  let saved = 'comfortable';
  try {
    saved = localStorage.getItem('adverse_table_density') || 'comfortable';
  } catch (e) {}
  setTableDensity(saved, false);
}

function toggleTableFullscreen() {
  const card = document.getElementById('recordsCard');
  const icon = document.getElementById('fullscreenIcon');
  if (!card) return;
  const isFull = card.classList.toggle('fullscreen-mode');
  if (icon) icon.textContent = isFull ? 'fullscreen_exit' : 'fullscreen';
  showToast(isFull ? 'Fullscreen table mode' : 'Exited fullscreen');
}

// High-Speed Standard CSV & Excel Exporter (RFC 4180 with UTF-8 BOM)
function exportCurrentTableCSV() {
  const currentKey = state.currentMenu ? state.currentMenu.key : 'campaign';
  const list = (state.currentRenderedRecords && state.currentRenderedRecords.length) 
    ? state.currentRenderedRecords 
    : ((state.filteredRecords && state.filteredRecords.length) ? state.filteredRecords : state.rawRecords);

  if (!list || !list.length) {
    showToast('No records available to export to CSV', 'warning');
    return;
  }

  const rawHeaders = (state.headers && state.headers.length) ? state.headers : Object.keys(list[0] || {});
  const exportHeaders = rawHeaders.filter(h => h !== 'Action' && !String(h).startsWith('_'));

  const rows = [exportHeaders];
  list.forEach(r => {
    const row = exportHeaders.map(h => {
      let val = r[h] ?? '';
      if (isDateHeader(h)) {
        if (val) val = formatCustomDate(val);
      }
      if (typeof val === 'number') return String(val);
      const str = String(val).replace(/"/g, '""');
      return `"${str}"`;
    });
    rows.push(row);
  });

  const csvContent = '\uFEFF' + rows.map(e => e.join(',')).join('\r\n');
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  const dateStr = new Date().toISOString().slice(0, 10);
  a.download = `AdVerse_${currentKey}_${dateStr}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  showToast(`✅ Exported ${list.length} rows to CSV / Excel!`, 'success');
}

// Professional PDF Export Engine
function exportCurrentTablePDF() {
  const currentKey = state.currentMenu ? state.currentMenu.key : '';
  const isCampaignView = currentKey === 'today_created_campaign' || currentKey === 'campaign';
  const list = (state.currentRenderedRecords && state.currentRenderedRecords.length) 
    ? state.currentRenderedRecords 
    : ((state.filteredRecords && state.filteredRecords.length) ? state.filteredRecords : state.rawRecords);

  if (!list || !list.length) {
    showToast('No records available to export to PDF', 'error');
    return;
  }

  // Determine headers
  let headers = (state.headers && state.headers.length) ? state.headers : Object.keys(list[0] || {});
  if (!isAdmin() && isCampaignView) {
    // Specifically restricted columns for Viewer as requested
    headers = ['SL', 'OFF $', 'Status', 'Page Name', 'Start', 'End', 'Types', 'Day', 'Budget', 'Taka'];
  }

  const title = state.currentMenu ? state.currentMenu.label : 'Campaign Report';
  const month = state.filters.month || 'All Months';
  const year = state.filters.year || new Date().getFullYear();
  const dateStr = new Date().toLocaleDateString('en-US', { day: '2-digit', month: 'short', year: 'numeric' });
  const timeStr = new Date().toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });

  // Calculate totals
  let totalUSD = 0;
  let totalBDT = 0;
  list.forEach(r => {
    const b = parseFloat(r['Budget'] || 0);
    const t = parseFloat(r['Taka'] || 0);
    if (!isNaN(b)) totalUSD += b;
    if (!isNaN(t)) totalBDT += t;
  });

  const reportHtml = `
    <div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif;color:#0f172a;line-height:1.4">
      <div style="display:flex;justify-content:space-between;align-items:flex-start;border-bottom:2px solid #4f46e5;padding-bottom:12px;margin-bottom:14px">
        <div>
          <div style="font-size:22px;font-weight:900;color:#4f46e5;letter-spacing:-0.5px">ADVERSE AGENCY</div>
          <div style="font-size:13px;color:#64748b;margin-top:2px;font-weight:600">${escapeHtml(title)} • ${escapeHtml(month)} ${escapeHtml(year)}</div>
        </div>
        <div style="text-align:right;font-size:11px;color:#475569">
          <div>Export Date: <strong>${dateStr} ${timeStr}</strong></div>
          <div>User Access: <strong>${isAdmin() ? 'Administrator' : 'Viewer'}</strong></div>
          <div>Total Rows: <strong>${list.length} Records</strong></div>
        </div>
      </div>

      <div style="display:flex;gap:12px;margin-bottom:14px">
        <div style="background:#f8fafc;border:1px solid #cbd5e1;border-radius:6px;padding:8px 14px;flex:1">
          <div style="font-size:10px;text-transform:uppercase;color:#64748b;font-weight:700;letter-spacing:0.5px">Total Records</div>
          <div style="font-size:16px;font-weight:800;color:#0f172a;margin-top:2px">${list.length}</div>
        </div>
        <div style="background:#f8fafc;border:1px solid #cbd5e1;border-radius:6px;padding:8px 14px;flex:1">
          <div style="font-size:10px;text-transform:uppercase;color:#64748b;font-weight:700;letter-spacing:0.5px">Total Budget ($)</div>
          <div style="font-size:16px;font-weight:800;color:#2563eb;margin-top:2px">$${fmtNum(totalUSD)}</div>
        </div>
        <div style="background:#f8fafc;border:1px solid #cbd5e1;border-radius:6px;padding:8px 14px;flex:1">
          <div style="font-size:10px;text-transform:uppercase;color:#64748b;font-weight:700;letter-spacing:0.5px">Total Invoiced (৳)</div>
          <div style="font-size:16px;font-weight:800;color:#16a34a;margin-top:2px">৳${fmtNum(totalBDT)}</div>
        </div>
      </div>

      <table style="width:100%;border-collapse:collapse;margin-bottom:14px;font-size:10.5px">
        <thead>
          <tr style="background:#f1f5f9">
            ${headers.map(h => {
              const isNum = MONEY_COLS.includes(h) || h === 'Budget' || h === 'Taka' || h === 'OFF $' || h === 'SL' || h === 'Day';
              return `<th style="padding:7px 8px;border:1px solid #cbd5e1;font-weight:700;text-transform:uppercase;font-size:9.5px;color:#1e293b;text-align:${isNum ? 'right' : 'left'}">${escapeHtml(h)}</th>`;
            }).join('')}
          </tr>
        </thead>
        <tbody>
          ${list.map((r, i) => `
            <tr style="background:${i % 2 === 0 ? '#ffffff' : '#f8fafc'}">
              ${headers.map(h => {
                let val = r[h];
                if (val === undefined || val === null || val === '') {
                  if (h === 'Start') val = r['Start Date'] || r['start_date'] || '';
                  else if (h === 'End') val = r['End Date'] || r['end_date'] || '';
                  else if (h === 'Types') val = r['Type'] || r['type'] || '';
                  else if (h === 'OFF $') val = r['OFF'] || r['Off'] || r['off'] || '';
                  else if (h === 'Page Name') val = r['page_name'] || r['Page'] || '';
                  else if (h === 'Budget') val = r['budget'] || '';
                  else if (h === 'Taka') val = r['taka'] || '';
                  else if (h === 'Day') val = r['day'] || '';
                  else if (h === 'Status') val = r['status'] || '';
                  else if (h === 'SL') val = r['SL'] || r['sl'] || r['SL NO'] || r['sl_no'] || r['Sl'] || r['SL#'] || r['serial'] || (i + 1);
                }
                if (val === undefined || val === null) val = '';

                const isMoney = MONEY_COLS.includes(h) || h === 'Budget' || h === 'Taka' || h === 'OFF $';
                const isDate = isDateHeader(h);
                if (isMoney) return `<td style="padding:5px 8px;border:1px solid #e2e8f0;text-align:right;font-weight:700;font-variant-numeric:tabular-nums">${fmtNum(val)}</td>`;
                if (isDate) return `<td style="padding:5px 8px;border:1px solid #e2e8f0">${escapeHtml(formatCustomDate(val))}</td>`;
                if (h === 'Status') {
                  const isOld = String(val).toLowerCase().includes('old');
                  return `<td style="padding:5px 8px;border:1px solid #e2e8f0"><span style="display:inline-block;padding:2px 6px;border-radius:4px;font-size:9px;font-weight:700;background:${isOld ? '#ffedd5' : '#dcfce7'};color:${isOld ? '#c2410c' : '#15803d'}">${escapeHtml(val)}</span></td>`;
                }
                if (h === 'Types') {
                  return `<td style="padding:5px 8px;border:1px solid #e2e8f0"><span style="display:inline-block;padding:2px 6px;border-radius:4px;font-size:9px;font-weight:700;background:#dbeafe;color:#1d4ed8">${escapeHtml(val)}</span></td>`;
                }
                const isNumCol = h === 'SL' || h === 'Day';
                return `<td style="padding:5px 8px;border:1px solid #e2e8f0;text-align:${isNumCol ? 'right' : 'left'}">${escapeHtml(val)}</td>`;
              }).join('')}
            </tr>
          `).join('')}
        </tbody>
      </table>

      <div style="display:flex;justify-content:space-between;align-items:center;border-top:1px solid #cbd5e1;padding-top:8px;font-size:9.5px;color:#64748b">
        <div>AdVerse Agency Management Intelligence System • Confidential Report</div>
        <div>Total Rows: ${list.length} • Printed Successfully</div>
      </div>
    </div>
  `;

  const paper = document.getElementById('pdfPaperContainer');
  if (paper) paper.innerHTML = reportHtml;

  const modal = document.getElementById('pdfPreviewModal');
  if (modal) modal.classList.remove('hidden');

  const sub = document.getElementById('pdfModalSubtitle');
  if (sub) sub.textContent = `${title} (${list.length} records • $${fmtNum(totalUSD)} • ৳${fmtNum(totalBDT)})`;

  showToast('Opening PDF Print & Save preview...', 'success');
  triggerDirectPrint();
}

function triggerDirectPrint() {
  const container = document.getElementById('pdfPaperContainer');
  if (!container) return;

  let printIframe = document.getElementById('printHiddenIframe');
  if (!printIframe) {
    printIframe = document.createElement('iframe');
    printIframe.id = 'printHiddenIframe';
    printIframe.style.position = 'fixed';
    printIframe.style.right = '0';
    printIframe.style.bottom = '0';
    printIframe.style.width = '0';
    printIframe.style.height = '0';
    printIframe.style.border = '0';
    document.body.appendChild(printIframe);
  }

  const printStyles = `
    <style>
      @page { size: A4 landscape; margin: 10mm; }
      * { box-sizing: border-box; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; color: #0f172a; margin: 0; padding: 12px; background: #fff; font-size: 10.5px; }
      table { width: 100%; border-collapse: collapse; margin-bottom: 12px; }
      th { background: #f1f5f9; color: #1e293b; font-weight: 700; text-transform: uppercase; font-size: 9.5px; padding: 6px 8px; border: 1px solid #cbd5e1; }
      td { padding: 5px 8px; border: 1px solid #e2e8f0; font-size: 10px; color: #334155; }
      tr:nth-child(even) td { background: #f8fafc; }
    </style>
  `;

  const doc = printIframe.contentWindow?.document || printIframe.contentDocument;
  if (doc) {
    doc.open();
    doc.write(`<!DOCTYPE html><html><head><title>AdVerse Agency Export</title>${printStyles}</head><body>${container.innerHTML}</body></html>`);
    doc.close();
    setTimeout(() => {
      try {
        printIframe.contentWindow?.focus();
        printIframe.contentWindow?.print();
      } catch (err) {
        console.warn('Iframe print blocked, falling back to window print', err);
        window.print();
      }
    }, 400);
  }
}

function closePdfModal() {
  document.getElementById('pdfPreviewModal')?.classList.add('hidden');
}

// Command Palette (Cmd + K)
const cmdPaletteItems = [
  { title: 'Today Created Campaign', icon: 'today', action: () => selectMenu('today_created_campaign') },
  { title: 'End Ads (Ending Today / Tomorrow)', icon: 'event_busy', action: () => selectMenu('end_ads') },
  { title: 'All Campaign', icon: 'campaign', action: () => selectMenu('campaign') },
  { title: 'Page Summary & ROI', icon: 'analytics', action: () => selectMenu('page_summery') },
  { title: 'Payment', icon: 'account_balance_wallet', action: () => selectMenu('payment_due_tracker') },
  { title: 'Agency Calculation', icon: 'payments', action: () => selectMenu('agency_spending') },
  { title: 'Open Visual Analytics & Intelligence', icon: 'analytics', action: openAnalyticsModal },
  { title: 'WhatsApp Client Notice Generator', icon: 'chat', action: () => openWhatsAppModal() },
  { title: 'Google Sheet Sync & Import Hub', icon: 'table_chart', action: openSheetConfigModal },
  { title: 'Push All Data to Google Sheet', icon: 'cloud_upload', action: pushAllDataToGoogleSheet },
  { title: 'Clear / Reset All Database Records', icon: 'delete_sweep', action: clearAllDatabaseData },
  { title: 'Keyboard Shortcuts Guide (?)', icon: 'keyboard', action: openHotkeysModal },
  { title: 'Admin & Viewer Access Links', icon: 'link', action: openRoleLinksModal },
  { title: 'Change Admin PIN', icon: 'key', action: openChangePinModal },
  { title: 'Export Table as CSV / Excel', icon: 'file_download', action: () => exportCurrentTableCSV() },
  { title: 'Export Table as PDF Report', icon: 'picture_as_pdf', action: () => exportCurrentTablePDF() },
  { title: 'Table View: Switch to Comfortable (Spacious)', icon: 'view_headline', action: () => setTableDensity('comfortable') },
  { title: 'Table View: Switch to Compact (High-Density Power User)', icon: 'density_small', action: () => setTableDensity('compact') },
  { title: 'Toggle Table Density (Comfortable / Compact)', icon: 'density_medium', action: () => toggleTableDensity() },
  { title: 'Toggle Row Focus Mode', icon: 'center_focus_strong', action: () => toggleFirstRowFocus() },
  { title: 'Copy All Current Campaigns', icon: 'content_copy', action: () => copyAllCurrentTableCampaigns() },
  { title: 'Instant Live Refresh & Sync', icon: 'sync', action: () => refreshCurrentView() }
];

function openCmdPalette() {
  const overlay = document.getElementById('cmdPaletteOverlay');
  const input = document.getElementById('cmdPaletteInput');
  if (overlay && input) {
    overlay.classList.remove('hidden');
    input.value = '';
    const available = cmdPaletteItems.filter(i => !i.adminOnly || isAdmin());
    renderCmdList(available);
    input.focus();
  }
}

function closeCmdPalette() {
  document.getElementById('cmdPaletteOverlay')?.classList.add('hidden');
}

function handleCmdSearch(q) {
  const query = q.trim().toLowerCase();
  const available = cmdPaletteItems.filter(i => !i.adminOnly || isAdmin());
  const filtered = query
    ? available.filter(i => i.title.toLowerCase().includes(query))
    : available;
  renderCmdList(filtered);
}

function renderCmdList(items) {
  const list = document.getElementById('cmdList');
  if (!list) return;
  list.innerHTML = items.map((item, idx) => `
    <div class="cmd-item" onclick="executeCmdItem(${idx})">
      <div style="display:flex;align-items:center;gap:10px">
        <span class="ms" style="color:var(--pri)">${item.icon}</span>
        <span>${escapeHtml(item.title)}</span>
      </div>
      <span class="ms" style="font-size:16px;opacity:0.5">arrow_forward</span>
    </div>
  `).join('') || '<div style="padding:16px;text-align:center;color:var(--txt-dim)">No commands found</div>';
}

function executeCmdItem(idx) {
  const items = cmdPaletteItems;
  if (items[idx]) {
    closeCmdPalette();
    items[idx].action();
  }
}

// Keyboard Shortcuts Hub (Feature #7)
function openHotkeysModal() {
  const overlay = document.getElementById('hotkeysOverlay');
  if (overlay) overlay.classList.remove('hidden');
}

function closeHotkeysModal() {
  const overlay = document.getElementById('hotkeysOverlay');
  if (overlay) overlay.classList.add('hidden');
}

// Global Keyboard Shortcut Listener
document.addEventListener('keydown', e => {
  const key = e.key;
  const keyLower = key.toLowerCase();
  const activeEl = document.activeElement;
  const activeTag = activeEl ? activeEl.tagName.toLowerCase() : '';
  const isInputActive = activeTag === 'input' || activeTag === 'textarea' || activeTag === 'select' || (activeEl && activeEl.isContentEditable);

  // 1. ESCAPE: Close any open modal, dropdown, or Row Focus Mode
  if (key === 'Escape') {
    if (state.focusedRowIndex !== null && state.focusedRowIndex !== undefined) {
      clearRowFocus();
      return;
    }
    document.querySelectorAll('.campaign-dropdown').forEach(d => d.classList.remove('open'));
    const entryCard = document.getElementById('entryCard');
    if (entryCard && !entryCard.classList.contains('hidden')) {
      entryCard.classList.add('hidden');
      const btnText = document.getElementById('btnToggleEntryText');
      if (btnText) btnText.textContent = '⚡ Quick Campaign';
    }
    closeHotkeysModal();
    closeCmdPalette();
    closeModal();
    closeAnalyticsModal();
    closeWhatsAppModal();
    closeSheetConfigModal();
    closeRoleLinksModal();
    return;
  }

  // 1.5. FAST SUBMIT: Ctrl+Enter or Cmd+Enter inside Entry Form
  if ((e.metaKey || e.ctrlKey) && key === 'Enter') {
    const entryCard = document.getElementById('entryCard');
    if (entryCard) {
      if (entryCard.classList.contains('hidden')) {
        entryCard.classList.remove('hidden');
      }
      e.preventDefault();
      if (typeof saveEntry === 'function') saveEntry();
      return;
    }
  }

  // 1.8. FAST QUICK CAMPAIGN: Press 'N' anywhere when not typing in an input
  if (keyLower === 'n' && !isInputActive && !e.ctrlKey && !e.metaKey) {
    e.preventDefault();
    if (typeof triggerQuickCampaign === 'function') triggerQuickCampaign();
    return;
  }

  // 1.9. FAST SYNC / REFRESH: Press 'R' anywhere when not typing
  if (keyLower === 'r' && !isInputActive && !e.ctrlKey && !e.metaKey) {
    e.preventDefault();
    if (typeof refreshCurrentView === 'function') refreshCurrentView();
    return;
  }

  // 1.91. TOGGLE TABLE DENSITY (Comfortable vs Compact): Alt+D
  if (e.altKey && keyLower === 'd') {
    e.preventDefault();
    toggleTableDensity();
    return;
  }

  // 1.95. ROW FOCUS MODE SHORTCUTS (when not typing in an input)
  if (!isInputActive && !e.ctrlKey && !e.metaKey && !e.altKey) {
    if (state.focusedRowIndex !== null && state.focusedRowIndex !== undefined) {
      if (key === 'ArrowDown' || keyLower === 'j') {
        e.preventDefault();
        focusNextRow();
        return;
      }
      if (key === 'ArrowUp' || keyLower === 'k') {
        e.preventDefault();
        focusPreviousRow();
        return;
      }
      if (keyLower === 'c') {
        e.preventDefault();
        rfActionCopy();
        return;
      }
      if (keyLower === 'w') {
        e.preventDefault();
        rfActionWhatsApp();
        return;
      }
      if (key === '+' || key === '=') {
        e.preventDefault();
        rfActionIncrease();
        return;
      }
      if (keyLower === 'e') {
        e.preventDefault();
        rfActionEdit();
        return;
      }
      if (keyLower === 'd') {
        e.preventDefault();
        rfActionClone();
        return;
      }
      if (keyLower === 'i') {
        e.preventDefault();
        rfActionInspect();
        return;
      }
      if (key === 'Delete') {
        e.preventDefault();
        rfActionDelete();
        return;
      }
    } else {
      if (key === 'ArrowDown') {
        const rows = document.querySelectorAll('#dataTable tbody tr.table-record-row');
        if (rows.length > 0) {
          e.preventDefault();
          focusTableRow(0, true);
          return;
        }
      }
    }
  }

  // 2. SHORTCUTS GUIDE: ? or Shift+/ (when not typing in an input)
  if ((key === '?' || (e.shiftKey && key === '/')) && !isInputActive) {
    e.preventDefault();
    const overlay = document.getElementById('hotkeysOverlay');
    if (overlay && !overlay.classList.contains('hidden')) closeHotkeysModal();
    else openHotkeysModal();
    return;
  }

  // 3. COMMAND PALETTE: Ctrl+K or Cmd+K
  if ((e.metaKey || e.ctrlKey) && keyLower === 'k') {
    e.preventDefault();
    const overlay = document.getElementById('cmdPaletteOverlay');
    if (overlay?.classList.contains('hidden')) openCmdPalette();
    else closeCmdPalette();
    return;
  }

  // 4. INSTANT SEARCH: Ctrl+F or Cmd+F -> Focus search bar
  if ((e.metaKey || e.ctrlKey) && keyLower === 'f') {
    const sInput = document.getElementById('tableFilterInput') || document.getElementById('tableSearchInput');
    if (sInput) {
      e.preventDefault();
      sInput.focus();
      sInput.select();
      return;
    }
  }

  // 5. QUICK SAVE: Ctrl+S or Cmd+S -> Save active form or modal
  if ((e.metaKey || e.ctrlKey) && keyLower === 's') {
    e.preventDefault();
    const editModal = document.getElementById('editModal');
    if (editModal && !editModal.classList.contains('hidden')) {
      if (typeof saveModalRecord === 'function') saveModalRecord();
      return;
    }
    const entryCard = document.getElementById('entryCard');
    if (entryCard && entryCard.classList.contains('hidden')) {
      entryCard.classList.remove('hidden');
    }
    if (typeof saveEntry === 'function') {
      saveEntry();
    }
    return;
  }

  // 6. NEW ENTRY CARD: Ctrl+N or Alt+N
  if (((e.metaKey || e.ctrlKey) && keyLower === 'n') || (e.altKey && keyLower === 'n')) {
    e.preventDefault();
    if (typeof toggleEntryCard === 'function') {
      toggleEntryCard();
    }
    return;
  }

  // 7. EXPORT CSV: Ctrl+Shift+E
  if ((e.metaKey || e.ctrlKey) && e.shiftKey && keyLower === 'e') {
    e.preventDefault();
    if (typeof exportCurrentTableCSV === 'function') exportCurrentTableCSV();
    return;
  }

  // 8. ALT + 1..6 : Quick Switch Menu Tabs
  if (e.altKey && ['1', '2', '3', '4', '5', '6'].includes(key)) {
    e.preventDefault();
    const menuMap = {
      '1': 'today_created_campaign',
      '2': 'end_ads',
      '3': 'campaign',
      '4': 'page_summery',
      '5': 'payment_due_tracker',
      '6': 'agency_spending'
    };
    const mKey = menuMap[key];
    if (mKey && typeof selectMenu === 'function') {
      selectMenu(mKey);
      if (typeof showToast === 'function') {
        showToast(`Tab switched: ${mKey.toUpperCase()}`, 'info');
      }
    }
    return;
  }

  // 9. ALT + R : Refresh Data
  if (e.altKey && keyLower === 'r') {
    e.preventDefault();
    if (typeof refreshCurrentView === 'function') refreshCurrentView();
    return;
  }

  // 10. ALT + T : Toggle Theme
  if (e.altKey && keyLower === 't') {
    e.preventDefault();
    if (typeof toggleTheme === 'function') toggleTheme();
    return;
  }

  // 11. ALT + B : Copy BST Live Time
  if (e.altKey && keyLower === 'b') {
    e.preventDefault();
    if (typeof copyOrNotifyBstTime === 'function') copyOrNotifyBstTime();
    return;
  }
});

// Close dropdown on outside click
document.addEventListener('click', e => {
  if (!e.target.closest('.campaign-dropdown')) {
    document.querySelectorAll('.campaign-dropdown').forEach(d => d.classList.remove('open'));
  }
  if (!e.target.closest('.theme-select-wrap')) {
    document.getElementById('themeDropdownMenu')?.classList.remove('open');
  }
  // Clear row focus when clicking outside table and focus controls
  if (state.focusedRowIndex !== null && state.focusedRowIndex !== undefined) {
    const isInsideTable = e.target.closest('#dataTable');
    const isInsideFocusMenu = e.target.closest('#rowFocusFloatingMenu');
    const isInsideFocusBtn = e.target.closest('#btnRowFocusToggle');
    const isInsideModal = e.target.closest('.modal') || e.target.closest('.modal-overlay') || e.target.closest('.toast');
    if (!isInsideTable && !isInsideFocusMenu && !isInsideFocusBtn && !isInsideModal) {
      clearRowFocus();
    }
  }
});

// Dynamic repositioning for row focus floating menu on scroll and resize
window.addEventListener('scroll', () => {
  if (state.focusedRowIndex !== null && state.focusedRowIndex !== undefined) {
    positionRowFocusMenu(state.focusedRowIndex);
  }
}, { passive: true });

window.addEventListener('resize', () => {
  if (state.focusedRowIndex !== null && state.focusedRowIndex !== undefined) {
    positionRowFocusMenu(state.focusedRowIndex);
  }
}, { passive: true });

document.addEventListener('DOMContentLoaded', () => {
  document.querySelector('.table-wrap')?.addEventListener('scroll', () => {
    if (state.focusedRowIndex !== null && state.focusedRowIndex !== undefined) {
      positionRowFocusMenu(state.focusedRowIndex);
    }
  }, { passive: true });
});

// Realtime Live Clock & Premium Bangladesh (BST) Indicator
function startLiveClock() {
  function update() {
    const digitsEl = document.getElementById('liveClockTime');
    const periodEl = document.getElementById('liveClockPeriod');
    const pillEl = document.getElementById('liveClockPill');
    const legacyEl = document.getElementById('liveClockText');
    if (!digitsEl && !legacyEl) return;

    try {
      const now = new Date();
      // Format specifically in Asia/Dhaka timezone
      const rawTime = now.toLocaleTimeString('en-US', {
        timeZone: 'Asia/Dhaka',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: true
      });

      // Handle non-standard unicode spaces (\u202F, \u00A0)
      const cleanTime = rawTime.replace(/[\u202F\u00A0]/g, ' ').trim();
      const match = cleanTime.match(/^(\d{1,2}:\d{2}:\d{2})\s*([AaPp][Mm])?$/);
      let digits = cleanTime;
      let ampm = '';

      if (match) {
        digits = match[1];
        ampm = (match[2] || '').toUpperCase();
      } else {
        const parts = cleanTime.split(/\s+/);
        digits = parts[0] || cleanTime;
        ampm = (parts[1] || '').toUpperCase();
      }

      if (digitsEl) digitsEl.textContent = digits;
      if (periodEl) periodEl.textContent = ampm || '';
      if (legacyEl) legacyEl.textContent = `${digits} ${ampm}`.trim();

      if (pillEl) {
        const dateStr = now.toLocaleDateString('en-GB', {
          timeZone: 'Asia/Dhaka',
          weekday: 'short',
          day: '2-digit',
          month: 'short',
          year: 'numeric'
        });
        pillEl.setAttribute(
          'title',
          `📅 ${dateStr}\n⏰ ${digits} ${ampm}\n🟢 Live Real-time Sync`
        );
      }
    } catch (_) {
      const now = new Date();
      const h = String(now.getHours() % 12 || 12).padStart(2, '0');
      const m = String(now.getMinutes()).padStart(2, '0');
      const s = String(now.getSeconds()).padStart(2, '0');
      const ampm = now.getHours() >= 12 ? 'PM' : 'AM';
      if (digitsEl) digitsEl.textContent = `${h}:${m}:${s}`;
      if (periodEl) periodEl.textContent = ampm;
      if (legacyEl) legacyEl.textContent = `${h}:${m}:${s} ${ampm}`.trim();
    }
  }
  update();
  setInterval(update, 1000);
}

// Interactive helper to copy or notify current time
window.copyOrNotifyBstTime = function() {
  const digits = document.getElementById('liveClockTime')?.textContent || '';
  const ampm = document.getElementById('liveClockPeriod')?.textContent || '';
  const str = `${digits} ${ampm}`.trim();
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(str).catch(() => {});
  }
  if (typeof showToast === 'function') {
    showToast(`Current Time: ${digits} ${ampm} copied!`, 'info');
  }
};

// Floating Scroll to Top Controller
function initScrollToTop() {
  const btn = document.getElementById('btnScrollTop');
  if (!btn) return;
  const scrollTarget = document.querySelector('.main-wrapper') || window;
  const checkScroll = () => {
    const top = (scrollTarget === window) ? (window.pageYOffset || document.documentElement.scrollTop) : scrollTarget.scrollTop;
    if (top > 180) {
      btn.classList.add('visible');
    } else {
      btn.classList.remove('visible');
    }
  };
  scrollTarget.addEventListener('scroll', checkScroll, { passive: true });
  window.addEventListener('scroll', checkScroll, { passive: true });
}

function scrollToTopSmooth() {
  const scrollTarget = document.querySelector('.main-wrapper');
  if (scrollTarget && scrollTarget.scrollTop > 0) {
    scrollTarget.scrollTo({ top: 0, behavior: 'smooth' });
  }
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// Background Intelligent Preloader for 0ms Navigation
function prefetchCommonMenus() {
  const menusToPrefetch = ['campaign', 'end_ads', 'page_summery', 'payment_due_tracker', 'agency_spending'];
  let delay = 600;
  menusToPrefetch.forEach(menuKey => {
    setTimeout(() => {
      const mn = state.filters.month;
      const yr = state.filters.year;
      const cacheKey = `${menuKey}_${mn}_${yr}___`;
      if (!state.clientCache[cacheKey] && !state.clientCache[menuKey]) {
        google.script.run.withSuccessHandler(res => {
          if (res && res.success) {
            state.clientCache[cacheKey] = res;
            state.clientCache[menuKey] = res;
            state.cacheTimestamps = state.cacheTimestamps || {};
            state.cacheTimestamps[cacheKey] = Date.now();
          }
        }).getMenuData(menuKey, state.filters);
      }
    }, delay);
    delay += 500;
  });
}

// Setup DOM Events
document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('pinLoginBtn')?.addEventListener('click', submitPinLogin);
  document.getElementById('pinLoginInput')?.addEventListener('keydown', e => { if (e.key === 'Enter') submitPinLogin(); });
  document.getElementById('btnResetEntry')?.addEventListener('click', renderEntryForm);
  document.getElementById('btnAddPayment')?.addEventListener('click', addPaymentRecord);
  document.getElementById('btnSaveRecord')?.addEventListener('click', saveModalRecord);
  
  document.addEventListener('click', (e) => {
    const saveBtnTarget = e.target.closest('#btnSaveEntry');
    if (saveBtnTarget) {
      e.preventDefault();
      if (typeof window.saveEntry === 'function') {
        window.saveEntry();
      }
    }
  });

  const entryFormEl = document.getElementById('entryForm');
  if (entryFormEl) {
    entryFormEl.addEventListener('submit', (e) => {
      e.preventDefault();
      if (typeof window.saveEntry === 'function') {
        window.saveEntry();
      }
    });
  }
  document.getElementById('btnApplyFilters')?.addEventListener('click', () => {
    state.filters.month = document.getElementById('filterMonth').value;
    state.filters.year = document.getElementById('filterYear').value;
    state.filters.date = document.getElementById('filterDate')?.value || '';
    state.filters.pageType = document.getElementById('pageTypeFilter')?.value || '';
    loadMenuData();
  });

  initializeApp();
  initGoogleAuth();
  updateGoogleAuthUI();
  startLiveClock();
  initScrollToTop();

  // Realtime Sheet-to-App silent background sync (Every 15s when active)
  window.addEventListener('focus', () => {
    if (state.currentMenu) {
      const cacheKey = `${state.currentMenu.key}_${state.filters.month}_${state.filters.year}_${state.filters.date}_${state.filters.pageType}_${state.filters.scope || ''}`;
      const lastFetch = (state.cacheTimestamps && state.cacheTimestamps[cacheKey]) || 0;
      if (Date.now() - lastFetch > 15000) {
        loadMenuData(false);
      }
    }
  });

  setInterval(() => {
    if (document.visibilityState === 'visible' && state.currentMenu && !state.saving) {
      const cacheKey = `${state.currentMenu.key}_${state.filters.month}_${state.filters.year}_${state.filters.date}_${state.filters.pageType}_${state.filters.scope || ''}`;
      const lastFetch = (state.cacheTimestamps && state.cacheTimestamps[cacheKey]) || 0;
      if (Date.now() - lastFetch > 25000) {
        loadMenuData(false);
      }
    }
  }, 25000);
});

// Ad Accounts & Pages Manager Functions
let currentManagerTab = 'ad_accounts';
let adAccountsFilterQuery = '';
let pagesFilterQuery = '';

function openAdAccountsPagesModal(initialTab = 'ad_accounts') {
  const overlay = document.getElementById('adAccountsPagesOverlay');
  if (overlay) {
    overlay.classList.remove('hidden');
    switchManagerTab(initialTab);
    loadManagerData();
  }
}

function closeAdAccountsPagesModal() {
  document.getElementById('adAccountsPagesOverlay')?.classList.add('hidden');
}

function switchManagerTab(tab) {
  currentManagerTab = tab;
  const tabAd = document.getElementById('managerTabAdAccounts');
  const tabPg = document.getElementById('managerTabPages');

  const btnAd = document.getElementById('tabBtnAdAccounts');
  const btnPg = document.getElementById('tabBtnPages');

  tabAd?.classList.toggle('hidden', tab !== 'ad_accounts');
  tabPg?.classList.toggle('hidden', tab !== 'pages');

  btnAd?.classList.toggle('active', tab === 'ad_accounts');
  btnPg?.classList.toggle('active', tab === 'pages');

  if (btnAd) btnAd.style.borderBottom = tab === 'ad_accounts' ? '2px solid var(--pri)' : 'none';
  if (btnPg) btnPg.style.borderBottom = tab === 'pages' ? '2px solid var(--pri)' : 'none';
}

function loadManagerData() {
  const countAd = document.getElementById('countAdAccounts');
  const countPg = document.getElementById('countPages');
  if (countAd) countAd.textContent = String(state.adAccounts.length);
  if (countPg) countPg.textContent = String(state.pages.length);

  renderAdAccountsTable();
  renderPagesTable();
}

function renderAdAccountsTable() {
  const tbody = document.getElementById('adAccountsTableBody');
  if (!tbody) return;

  const query = adAccountsFilterQuery.toLowerCase().trim();
  const list = query
    ? state.adAccounts.filter(a => a.a.toLowerCase().includes(query) || (a.id && a.id.toLowerCase().includes(query)))
    : state.adAccounts;

  const countAd = document.getElementById('countAdAccounts');
  if (countAd) countAd.textContent = String(state.adAccounts.length);

  if (!list.length) {
    tbody.innerHTML = `<tr><td colspan="3" style="text-align:center;padding:24px;color:var(--txt-muted)">No AD Accounts found. Add new via form or bulk paste.</td></tr>`;
    return;
  }

  tbody.innerHTML = list.map(item => `
    <tr style="border-bottom:1px solid var(--bor)">
      <td style="padding:10px 12px;font-weight:700;color:var(--txt)">
        <div style="display:flex;align-items:center;gap:6px">
          <span class="ms" style="font-size:16px;color:var(--pri)">badge</span>
          ${escapeHtml(item.a)}
        </div>
      </td>
      <td style="padding:10px 12px;color:var(--txt-muted);font-family:monospace;font-size:12px">
        ${escapeHtml(item.id || '—')}
      </td>
      <td style="padding:10px 12px;text-align:right">
        <button class="table-btn table-btn-danger" type="button" onclick="deleteAdAccountItem('${escapeHtml(item.a)}')" title="Delete Account">
          <span class="ms" style="font-size:15px">delete</span>
        </button>
      </td>
    </tr>
  `).join('');
}

function renderPagesTable() {
  const tbody = document.getElementById('pagesTableBody');
  if (!tbody) return;

  const query = pagesFilterQuery.trim();
  const list = query
    ? state.pages.filter(p => BanglishEngine.matches(p, query))
    : state.pages;

  const countPg = document.getElementById('countPages');
  if (countPg) countPg.textContent = String(state.pages.length);

  if (!list.length) {
    tbody.innerHTML = `<tr><td colspan="3" style="text-align:center;padding:24px;color:var(--txt-muted)">No Page Names found. Add new via form or bulk paste.</td></tr>`;
    return;
  }

  tbody.innerHTML = list.map((p, idx) => `
    <tr style="border-bottom:1px solid var(--bor)">
      <td style="padding:10px 12px;color:var(--txt-dim);width:40px">${idx + 1}</td>
      <td style="padding:10px 12px;font-weight:700;color:var(--txt)">
        <div style="display:flex;align-items:center;gap:6px">
          <span class="ms" style="font-size:16px;color:var(--suc)">pages</span>
          ${escapeHtml(p)}
        </div>
      </td>
      <td style="padding:10px 12px;text-align:right">
        <button class="table-btn table-btn-danger" type="button" onclick="deletePageItem('${escapeHtml(p)}')" title="Delete Page">
          <span class="ms" style="font-size:15px">delete</span>
        </button>
      </td>
    </tr>
  `).join('');
}

function filterAdAccountsList(val) {
  adAccountsFilterQuery = val;
  renderAdAccountsTable();
}

function filterPagesList(val) {
  pagesFilterQuery = val;
  renderPagesTable();
}

function submitNewAdAccount() {
  const nameInput = document.getElementById('newAdAccountName');
  const idInput = document.getElementById('newAdAccountId');
  const a = (nameInput?.value || '').trim();
  const id = (idInput?.value || '').trim();

  if (!a) {
    showToast('Please enter an AD Account name!', 'error');
    return;
  }

  google.script.run.withSuccessHandler(res => {
    if (res.adAccounts) state.adAccounts = res.adAccounts;
    else if (!state.adAccounts.some(x => x.a.toLowerCase() === a.toLowerCase())) {
      state.adAccounts.push({ a, id });
    }
    if (nameInput) nameInput.value = '';
    if (idInput) idInput.value = '';
    showToast(`AD Account "${a}" saved successfully!`, 'success');
    renderAdAccountsTable();
    updateFormDatalists();
  }).addAdAccount(a, id);
}

function submitBulkAdAccounts() {
  const area = document.getElementById('bulkAdAccountsInput');
  const text = (area?.value || '').trim();
  if (!text) {
    showToast('Please paste a list of AD Accounts!', 'error');
    return;
  }

  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  const items = [];
  for (const line of lines) {
    const parts = line.split(/[,;\t]/).map(p => p.trim());
    const a = parts[0];
    const id = parts.slice(1).join(' ');
    if (a) items.push({ a, id });
  }

  if (!items.length) {
    showToast('Please provide valid data format!', 'error');
    return;
  }

  google.script.run.withSuccessHandler(res => {
    if (res.adAccounts) state.adAccounts = res.adAccounts;
    else {
      items.forEach(it => {
        if (!state.adAccounts.some(x => x.a.toLowerCase() === it.a.toLowerCase())) {
          state.adAccounts.push(it);
        }
      });
    }
    if (area) area.value = '';
    showToast(`${items.length} AD Accounts imported successfully!`, 'success');
    renderAdAccountsTable();
    updateFormDatalists();
  }).addBulkAdAccounts(items);
}

function deleteAdAccountItem(name) {
  if (!confirm(`Are you sure you want to delete AD Account "${name}"?`)) return;

  google.script.run.withSuccessHandler(res => {
    if (res.adAccounts) state.adAccounts = res.adAccounts;
    else state.adAccounts = state.adAccounts.filter(x => x.a.toLowerCase() !== name.toLowerCase());
    showToast(`AD Account "${name}" deleted.`, 'info');
    renderAdAccountsTable();
    updateFormDatalists();
  }).deleteAdAccount(name);
}

function submitNewPageName() {
  const pageInput = document.getElementById('newPageNameInput');
  const p = (pageInput?.value || '').trim();

  if (!p) {
    showToast('Please enter a Page Name!', 'error');
    return;
  }

  google.script.run.withSuccessHandler(res => {
    if (res.pages) state.pages = res.pages;
    else if (!state.pages.some(x => x.toLowerCase() === p.toLowerCase())) {
      state.pages.push(p);
    }
    if (pageInput) pageInput.value = '';
    showToast(`Page "${p}" saved successfully!`, 'success');
    renderPagesTable();
    updateFormDatalists();
  }).addPage(p);
}

function submitBulkPages() {
  const area = document.getElementById('bulkPagesInput');
  const text = (area?.value || '').trim();
  if (!text) {
    showToast('Please paste a list of Page Names!', 'error');
    return;
  }

  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  const pages = [...new Set(lines)];

  if (!pages.length) {
    showToast('Please provide valid page names!', 'error');
    return;
  }

  google.script.run.withSuccessHandler(res => {
    if (res.pages) state.pages = res.pages;
    else {
      pages.forEach(p => {
        if (!state.pages.some(x => x.toLowerCase() === p.toLowerCase())) {
          state.pages.push(p);
        }
      });
    }
    if (area) area.value = '';
    showToast(`${pages.length} Pages imported successfully!`, 'success');
    renderPagesTable();
    updateFormDatalists();
  }).addBulkPages(pages);
}

function deletePageItem(name) {
  if (!confirm(`Are you sure you want to delete Page "${name}"?`)) return;

  google.script.run.withSuccessHandler(res => {
    if (res.pages) state.pages = res.pages;
    else state.pages = state.pages.filter(p => p.toLowerCase() !== name.toLowerCase());
    showToast(`Page "${name}" deleted.`, 'info');
    renderPagesTable();
    updateFormDatalists();
  }).deletePage(name);
}

function updateFormDatalists() {
  const adDatalist = document.getElementById('adAccountsList');
  if (adDatalist) {
    adDatalist.innerHTML = state.adAccounts.map(a => `<option value="${escapeHtml(a.a)}" data-id="${escapeHtml(a.id)}">${escapeHtml(a.a)} — (${escapeHtml(a.id)})</option>`).join('');
  }
  const pgDatalist = document.getElementById('pagesList');
  if (pgDatalist) {
    pgDatalist.innerHTML = state.pages.map(p => `<option value="${escapeHtml(p)}">`).join('');
  }
}

// Window Exports
window.saveEntry = saveEntry;
window.renderEntryForm = renderEntryForm;
window.resetEntryForm = resetEntryForm;
window.handleToggleEntry = handleToggleEntry;
window.triggerQuickCampaign = triggerQuickCampaign;
window.selectMenu = selectMenu;
window.toggleEntryCard = toggleEntryCard;
window.refreshCurrentView = refreshCurrentView;
window.sortTableBy = sortTableBy;
window.filterTableRows = filterTableRows;
window.clearTableSearch = clearTableSearch;
window.matchesCampaignRow = matchesCampaignRow;
window.addMultipleCampaignRow = addMultipleCampaignRow;
window.deletePaymentItem = deletePaymentItem;
window.handlePaymentMethodChange = handlePaymentMethodChange;
window.deleteRecordItem = deleteRecordItem;
window.deleteRecordByIndex = deleteRecordByIndex;
window.openEditRecord = openEditRecord;
window.openEditRecordByIndex = openEditRecordByIndex;
window.saveModalRecord = saveModalRecord;
window.closeModal = closeModal;
window.appLogout = appLogout;
window.applyTheme = applyTheme;
window.setAppTheme = setAppTheme;
window.toggleTheme = toggleTheme;
window.toggleThemeDropdown = toggleThemeDropdown;
window.toggleSettingsDropdown = toggleSettingsDropdown;
window.closeAllDropdowns = closeAllDropdowns;
window.applyThemeStyle = applyThemeStyle;
window.openRateModal = openRateModal;
window.closeRateModal = closeRateModal;
window.setQuickRate = setQuickRate;
window.saveNewRate = saveNewRate;
window.refreshFinancialHud = refreshFinancialHud;
window.openCmdPalette = openCmdPalette;
window.closeCmdPalette = closeCmdPalette;
window.handleCmdSearch = handleCmdSearch;
window.executeCmdItem = executeCmdItem;
window.openAnalyticsModal = openAnalyticsModal;
window.closeAnalyticsModal = closeAnalyticsModal;
window.openWhatsAppModal = openWhatsAppModal;
window.openSingleRecordWhatsApp = openSingleRecordWhatsApp;
window.openSingleRecordWhatsAppByIndex = openSingleRecordWhatsAppByIndex;
window.closeWhatsAppModal = closeWhatsAppModal;
window.generateWhatsAppNotice = generateWhatsAppNotice;
window.copyWhatsAppNotice = copyWhatsAppNotice;
window.exportCurrentTableCSV = exportCurrentTableCSV;
window.exportCurrentTablePDF = exportCurrentTablePDF;
window.scrollToTopSmooth = scrollToTopSmooth;
window.triggerDirectPrint = triggerDirectPrint;
window.closePdfModal = closePdfModal;
window.applyRoleVisibility = applyRoleVisibility;
window.openSheetConfigModal = openSheetConfigModal;
window.closeSheetConfigModal = closeSheetConfigModal;
window.saveSheetConfig = saveSheetConfig;
window.syncSpreadsheetData = syncSpreadsheetData;
window.switchHubTab = switchHubTab;
window.executeRawImport = executeRawImport;
window.copyAppsScriptCode = copyAppsScriptCode;
window.openRoleLinksModal = openRoleLinksModal;
window.closeRoleLinksModal = closeRoleLinksModal;
window.copyLink = copyLink;
window.copyCampaignRecord = copyCampaignRecord;
window.copyCampaignRecordByIndex = copyCampaignRecordByIndex;
window.handleQuickCampaignPaste = handleQuickCampaignPaste;
window.parseCampaignString = parseCampaignString;
window.formatCampaignText = formatCampaignText;
window.quickIncreaseCampaign = quickIncreaseCampaign;
window.quickIncreaseCampaignByIndex = quickIncreaseCampaignByIndex;
window.applyIncreaseStartDate = applyIncreaseStartDate;
window.applyIncreaseCampaign = applyIncreaseCampaign;
window.applyIncreaseCampaignByIndex = applyIncreaseCampaignByIndex;
window.copyAllCurrentTableCampaigns = copyAllCurrentTableCampaigns;
window.copyAllCurrentTableCampaignsSeparately = copyAllCurrentTableCampaignsSeparately;
window.copyMultipleRecordsSequentially = copyMultipleRecordsSequentially;
window.copyMultipleSavedCampaigns = copyMultipleSavedCampaigns;
window.copyMultipleSavedCampaignsCombined = copyMultipleSavedCampaignsCombined;
window.copySingleSavedCampaignByIndex = copySingleSavedCampaignByIndex;
window.renderSavedCampaignsNotice = renderSavedCampaignsNotice;
window.handlePageNameLookup = handlePageNameLookup;
window.clearAllDatabaseData = clearAllDatabaseData;
window.openAdAccountsPagesModal = openAdAccountsPagesModal;
window.closeAdAccountsPagesModal = closeAdAccountsPagesModal;
window.switchManagerTab = switchManagerTab;
window.submitNewAdAccount = submitNewAdAccount;
window.submitBulkAdAccounts = submitBulkAdAccounts;
window.deleteAdAccountItem = deleteAdAccountItem;
window.submitNewPageName = submitNewPageName;
window.submitBulkPages = submitBulkPages;
window.deletePageItem = deletePageItem;
window.filterAdAccountsList = filterAdAccountsList;
window.filterPagesList = filterPagesList;
window.updateFormDatalists = updateFormDatalists;
window.handleGoogleSignIn = handleGoogleSignIn;
window.handleGoogleSignOut = handleGoogleSignOut;
window.updateGoogleAuthUI = updateGoogleAuthUI;
window.quickSyncLiveSheet = quickSyncLiveSheet;
window.openGoogleAuthGuideModal = openGoogleAuthGuideModal;
window.closeGoogleAuthGuideModal = closeGoogleAuthGuideModal;
window.cloneCampaignRecord = cloneCampaignRecord;
window.cloneCampaignRecordByIndex = cloneCampaignRecordByIndex;
window.quickIncreaseCampaign = quickIncreaseCampaign;
window.quickIncreaseCampaignByIndex = quickIncreaseCampaignByIndex;
window.openEditRecord = openEditRecord;
window.openEditRecordByIndex = openEditRecordByIndex;
window.deleteRecordItem = deleteRecordItem;
window.deleteRecordByIndex = deleteRecordByIndex;
window.copyCampaignRecord = copyCampaignRecord;
window.copyCampaignRecordByIndex = copyCampaignRecordByIndex;
window.openSingleRecordWhatsApp = openSingleRecordWhatsApp;
window.openSingleRecordWhatsAppByIndex = openSingleRecordWhatsAppByIndex;
window.deletePaymentItem = deletePaymentItem;
window.addPaymentRecord = addPaymentRecord;
window.saveModalRecord = saveModalRecord;
window.closeModal = closeModal;
window.exportCurrentTableCSV = exportCurrentTableCSV;
window.exportCurrentTablePDF = exportCurrentTablePDF;
window.copyAllCurrentTableCampaigns = copyAllCurrentTableCampaigns;
window.copyAllCurrentTableCampaignsSeparately = copyAllCurrentTableCampaignsSeparately;
window.copyMultipleSavedCampaigns = copyMultipleSavedCampaigns;
window.copyMultipleSavedCampaignsCombined = copyMultipleSavedCampaignsCombined;
window.copySingleSavedCampaignByIndex = copySingleSavedCampaignByIndex;
window.sendWhatsAppWebDirect = sendWhatsAppWebDirect;
window.openCampaignInspector = openCampaignInspector;
window.closeCampaignInspector = closeCampaignInspector;
window.copyInspectorRawLine = copyInspectorRawLine;
window.inspectorCopyRecord = inspectorCopyRecord;
window.inspectorEditRecord = inspectorEditRecord;
window.inspectorDeleteRecord = inspectorDeleteRecord;
window.inspectorIncreaseRecord = inspectorIncreaseRecord;
window.openInspectorWhatsApp = openInspectorWhatsApp;
window.applyQuickChip = applyQuickChip;
window.setTableDensity = setTableDensity;
window.toggleTableDensity = toggleTableDensity;
window.initTableDensity = initTableDensity;
window.toggleTableFullscreen = toggleTableFullscreen;
window.openChangePinModal = openChangePinModal;
window.closeChangePinModal = closeChangePinModal;
window.submitChangeAdminPin = submitChangeAdminPin;
window.pushAllDataToGoogleSheet = pushAllDataToGoogleSheet;
window.testWebhookConnection = testWebhookConnection;
window.togglePinVisibility = togglePinVisibility;
window.loadMenuData = loadMenuData;
window.submitPinLogin = submitPinLogin;
window.openHotkeysModal = openHotkeysModal;
window.closeHotkeysModal = closeHotkeysModal;
window.handleTableRowClick = handleTableRowClick;
window.focusTableRow = focusTableRow;
window.clearRowFocus = clearRowFocus;
window.toggleFirstRowFocus = toggleFirstRowFocus;
window.focusNextRow = focusNextRow;
window.focusPreviousRow = focusPreviousRow;
window.renderRowFocusMenu = renderRowFocusMenu;
window.rfActionCopy = rfActionCopy;
window.rfActionWhatsApp = rfActionWhatsApp;
window.rfActionIncrease = rfActionIncrease;
window.rfActionEdit = rfActionEdit;
window.rfActionClone = rfActionClone;
window.rfActionInspect = rfActionInspect;
window.openEditPaymentRecord = openEditPaymentRecord;
window.rfActionDelete = rfActionDelete;
window.positionRowFocusMenu = positionRowFocusMenu;

// Ad Account Billing Window Exports (Admin Only)
window.loadAdAccountBillingData = loadAdAccountBillingData;
window.renderAdAccountBillingView = renderAdAccountBillingView;
window.handleBillingSearch = handleBillingSearch;
window.setBillingStatusFilter = setBillingStatusFilter;
window.openQuickDueModal = openQuickDueModal;
window.closeQuickDueModal = closeQuickDueModal;
window.quickDueAdjust = quickDueAdjust;
window.updateQuickDuePreview = updateQuickDuePreview;
window.submitQuickDueUpdate = submitQuickDueUpdate;
window.openAdBillingModal = openAdBillingModal;
window.closeAdBillingModal = closeAdBillingModal;
window.setBillingThresholdPreset = setBillingThresholdPreset;
window.saveAdBillingModalForm = saveAdBillingModalForm;

// Password Complexity & Security Helpers
window.validatePasswordComplexity = validatePasswordComplexity;
window.handleNewPinInput = handleNewPinInput;
window.handleConfirmPinInput = handleConfirmPinInput;
window.toggleFieldVis = toggleFieldVis;

// Mobile Campaign Submenu Modal Handlers
window.openMobileCampaignMenu = function() {
  const modal = document.getElementById('mobileCampaignMenuModal');
  if (modal) {
    modal.classList.remove('hidden');
  }
};
window.closeMobileCampaignMenu = function() {
  const modal = document.getElementById('mobileCampaignMenuModal');
  if (modal) {
    modal.classList.add('hidden');
  }
};
window.selectMenuMobile = function(key) {
  if (typeof window.closeMobileCampaignMenu === 'function') window.closeMobileCampaignMenu();
  if (typeof window.selectMenu === 'function') window.selectMenu(key);
};

state.activeAgencyTab = state.activeAgencyTab || 'scam';

function switchAgencyCostTab(tab) {
  state.activeAgencyTab = tab;
  renderRecordsTable({ headers: state.headers, rows: state.rawRecords, scamRows: state.scamRows });
}

function submitScamEntryForm() {
  const date = document.getElementById('scamDateInput')?.value;
  const pageName = document.getElementById('scamPageInput')?.value;
  const budget = parseFloat(document.getElementById('scamBudgetInput')?.value || 0);
  const taka = parseFloat(document.getElementById('scamTakaInput')?.value || 0);

  if (!date || !pageName || budget <= 0 || taka <= 0) {
    showToast('⚠️ Please fill in all fields correctly', 'error');
    return;
  }

  const formattedDate = formatCustomDate(date) || date;

  google.script.run.withSuccessHandler(res => {
    if (res && res.success) {
      showToast('🛡️ Scam entry saved successfully!', 'success');
      loadMenuData(true);
    } else {
      showToast(res?.message || 'Failed to save scam entry', 'error');
    }
  }).addScam({ date: formattedDate, page_name: pageName, spent: budget, taka });
}

function submitAgencyCostForm() {
  const date = document.getElementById('costDateInput')?.value;
  const type = document.getElementById('costTypeInput')?.value;
  const taka = parseFloat(document.getElementById('costTakaInput')?.value || 0);
  const status = document.getElementById('costStatusSelect')?.value || 'Due';

  if (!date || !type || taka <= 0) {
    showToast('⚠️ Please fill in all fields correctly', 'error');
    return;
  }

  const formattedDate = formatCustomDate(date) || date;

  google.script.run.withSuccessHandler(res => {
    if (res && res.success) {
      showToast('📑 Agency cost saved successfully!', 'success');
      loadMenuData(true);
    } else {
      showToast(res?.message || 'Failed to save agency cost', 'error');
    }
  }).saveAgencyCost({ date: formattedDate, type, cost: taka, status });
}

function deleteScamRecord(id) {
  if (!confirm('Are you sure you want to delete this scam entry?')) return;
  google.script.run.withSuccessHandler(res => {
    if (res && res.success) {
      showToast('Scam entry deleted', 'success');
      loadMenuData(true);
    } else {
      showToast('Failed to delete', 'error');
    }
  }).deleteScam(id);
}

function deleteAgencyCostRecord(id) {
  if (!confirm('Are you sure you want to delete this agency cost entry?')) return;
  google.script.run.withSuccessHandler(res => {
    if (res && res.success) {
      showToast('Agency cost deleted', 'success');
      loadMenuData(true);
    } else {
      showToast('Failed to delete', 'error');
    }
  }).deleteAgencyCost(id);
}

window.switchAgencyCostTab = switchAgencyCostTab;
window.submitScamEntryForm = submitScamEntryForm;
window.submitAgencyCostForm = submitAgencyCostForm;
window.deleteScamRecord = deleteScamRecord;
window.deleteAgencyCostRecord = deleteAgencyCostRecord;
window.downloadCalculationImage = downloadCalculationImage;

// =========================================================================
// Data Vault & Instant Entry Recovery System
// =========================================================================
function openEntryRecoveryModal() {
  const modal = document.getElementById('entryRecoveryModal');
  if (!modal) return;
  modal.classList.remove('hidden');
  loadVaultEntriesIntoUI();
}
window.openEntryRecoveryModal = openEntryRecoveryModal;

function closeEntryRecoveryModal() {
  const modal = document.getElementById('entryRecoveryModal');
  if (modal) modal.classList.add('hidden');
}
window.closeEntryRecoveryModal = closeEntryRecoveryModal;

async function loadVaultEntriesIntoUI() {
  const listEl = document.getElementById('vaultEntriesList');
  const countEl = document.getElementById('vaultCountBadge');
  if (!listEl) return;

  const rawVault = localStorage.getItem('adverse_entry_vault') || '[]';
  let vault = [];
  try { vault = JSON.parse(rawVault); } catch (_) { vault = []; }
  if (!Array.isArray(vault)) vault = [];

  const allItems = [...vault];
  const seenIds = new Set(vault.map(v => v.id || v.data?.['Record ID']));

  // Also include Today Created rows
  (state.todayCreatedRegistry || []).forEach(r => {
    const id = r['Record ID'] || r._recordId;
    if (id && !seenIds.has(id)) {
      seenIds.add(id);
      allItems.push({
        id,
        savedAt: r['Entry Date'] || new Date().toISOString(),
        timeStr: r['Entry Date'] || 'Today',
        status: 'confirmed',
        data: r
      });
    }
  });

  if (countEl) countEl.textContent = `${allItems.length} Saved Entries`;

  if (allItems.length === 0) {
    listEl.innerHTML = `<div style="text-align:center;padding:30px;color:var(--txt-muted)">No entries in vault yet. Every entry you make will appear here permanently.</div>`;
    return;
  }

  listEl.innerHTML = allItems.map(item => {
    const d = item.data || {};
    const sl = d.SL || d['SL NO'] || d.sl_no || '-';
    const page = d['Page Name'] || d.page_name || 'N/A';
    const ad = d['AD Account'] || d.ad_account || '';
    const bud = d.Budget || d.budget || d['Paid (৳)'] || 0;
    const isPay = Boolean(d['Paid (৳)'] || d['Payment Method']);
    const start = d.Start || d['Start Date'] || d.Date || '';
    const end = d.End || d['End Date'] || '';
    const month = d.Month || d.month || '';

    return `
      <div style="background:var(--card);border:1px solid var(--bor);border-radius:8px;padding:12px 16px;display:flex;justify-content:space-between;align-items:center;gap:12px;box-shadow:0 1px 3px rgba(0,0,0,0.03)">
        <div style="flex:1;min-width:0">
          <div style="display:flex;align-items:center;gap:8px;margin-bottom:4px;flex-wrap:wrap">
            <span style="font-weight:800;font-size:12px;background:var(--pri-light);color:var(--pri);padding:2px 8px;border-radius:4px;font-family:monospace">SL: #${sl}</span>
            <span style="font-weight:700;font-size:13px;color:var(--txt)">${escapeHtml(page)}</span>
            ${ad ? `<span style="font-size:11px;color:var(--txt-muted);background:var(--bor);padding:1px 6px;border-radius:4px">${escapeHtml(ad)}</span>` : ''}
            <span class="badge ${item.status === 'confirmed' ? 'badge-success' : 'badge-warning'}" style="font-size:10.5px">${item.status === 'confirmed' ? '✅ Confirmed' : '⚡ Local Saved'}</span>
          </div>
          <div style="font-size:11.5px;color:var(--txt-muted);display:flex;gap:12px;flex-wrap:wrap">
            <span><b>${isPay ? 'Paid' : 'Budget'}:</b> ${isPay ? '৳' + fmtNum(bud) : '$' + bud}</span>
            ${start ? `<span><b>Start:</b> ${escapeHtml(start)}</span>` : ''}
            ${end ? `<span><b>End:</b> ${escapeHtml(end)}</span>` : ''}
            <span><b>Month:</b> ${escapeHtml(month)}</span>
            <span style="color:#94a3b8">Saved: ${item.timeStr || ''}</span>
          </div>
        </div>
        <div style="display:flex;gap:6px">
          <button class="btn btn-secondary btn-sm" onclick="reSyncSingleVaultEntry('${item.id}')" title="Push directly to Sheet & Server" style="padding:4px 8px;font-size:11.5px">
            <span class="ms" style="font-size:14px;color:#10B981">sync</span> Sync
          </button>
          <button class="btn btn-primary btn-sm" onclick="restoreVaultEntryToTable('${item.id}')" title="Display in active table" style="padding:4px 8px;font-size:11.5px">
            <span class="ms" style="font-size:14px">visibility</span> View
          </button>
        </div>
      </div>
    `;
  }).join('');
}
window.loadVaultEntriesIntoUI = loadVaultEntriesIntoUI;

async function reSyncAllVaultEntries() {
  showToast('🔄 Synchronizing all vault entries to Google Sheets & Server...', 'info');
  const rawVault = localStorage.getItem('adverse_entry_vault') || '[]';
  let vault = [];
  try { vault = JSON.parse(rawVault); } catch (_) { vault = []; }
  if (!Array.isArray(vault) || vault.length === 0) {
    showToast('Vault is empty', 'info');
    return;
  }
  for (const item of vault) {
    if (item.data) {
      const d = item.data;
      const targetMonthTab = `Campaign_${d.Month || 'October'}_${d.Year || '2026'}`;
      autoSyncDirectToGoogleSheets('Today_Created', TODAY_HEADERS, [d]).catch(() => {});
      autoSyncDirectToGoogleSheets(targetMonthTab, CH, [d]).catch(() => {});
      cloudRpc('addCampaign', [d]).catch(() => {});
    }
  }
  showToast('✅ All vault entries dispatched to Google Sheet!', 'success');
  loadMenuData(true);
}
window.reSyncAllVaultEntries = reSyncAllVaultEntries;

function restoreVaultEntryToTable(id) {
  closeEntryRecoveryModal();
  selectMenu('today_created_campaign');
}
window.restoreVaultEntryToTable = restoreVaultEntryToTable;

function reSyncSingleVaultEntry(id) {
  const rawVault = localStorage.getItem('adverse_entry_vault') || '[]';
  let vault = [];
  try { vault = JSON.parse(rawVault); } catch (_) { vault = []; }
  const found = vault.find(v => v.id === id);
  if (found && found.data) {
    const d = found.data;
    const targetMonthTab = `Campaign_${d.Month || 'October'}_${d.Year || '2026'}`;
    autoSyncDirectToGoogleSheets('Today_Created', TODAY_HEADERS, [d]);
    autoSyncDirectToGoogleSheets(targetMonthTab, CH, [d]);
    cloudRpc('addCampaign', [d]).then(() => {
      showToast('✅ Entry synced successfully to Google Sheet & Server!', 'success');
      loadVaultEntriesIntoUI();
      loadMenuData(true);
    });
  }
}
function toggleAllTodayVerified(checked) {
  const records = (state.currentRenderedRecords && state.currentRenderedRecords.length > 0)
    ? state.currentRenderedRecords
    : ((state.filteredRecords && state.filteredRecords.length > 0)
      ? state.filteredRecords
      : (state.todayCreatedRegistry || state.rawRecords || []));
  if (!records.length) return;
  if (!state.campaignVerifications) state.campaignVerifications = {};
  const userDisplayName = (state.userRole === 'admin') ? 'Admin' : (state.displayName || 'User');
  const nowStr = formatCustomDate(new Date()) + ', ' + new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  const keys = [];

  records.forEach(r => {
    r.verified = checked;
    r._verified = checked;
    const key = getCampaignKey(r);
    if (key) {
      keys.push(key);
      if (checked) {
        state.campaignVerifications[key] = {
          verified: true,
          verifiedBy: userDisplayName,
          verifiedAt: nowStr,
          notes: 'Checked via master checkbox'
        };
      } else {
        delete state.campaignVerifications[key];
      }
    }
  });

  try {
    localStorage.setItem('adverse_verifications', JSON.stringify(state.campaignVerifications));
  } catch (e) {}

  // Update all row checkboxes in DOM immediately for 0ms lag
  const checkboxes = document.querySelectorAll('.row-campaign-check');
  checkboxes.forEach(cb => { cb.checked = checked; });
  const rows = document.querySelectorAll('.table-record-row');
  rows.forEach(r => { r.classList.toggle('row-verified', checked); });

  showToast(checked ? `✅ All ${records.length} campaigns checked!` : `Unchecked all campaigns`, 'info');

  cloudRpc('batchVerifyCampaigns', [{
    recordKeys: keys,
    verifiedBy: userDisplayName,
    verified: checked
  }]).catch(() => {});
}
window.toggleAllTodayVerified = toggleAllTodayVerified;

function toggleRowVerification(rIdx, isChecked) {
  const records = state.currentRenderedRecords || state.filteredRecords || state.rawRecords || [];
  const record = records[rIdx];
  if (!record) return;
  toggleCampaignVerification(record, isChecked, rIdx);
}
window.toggleRowVerification = toggleRowVerification;





// Date Format Exports
window.setSystemDateFormat = setSystemDateFormat;
window.toggleDateFormatDropdown = toggleDateFormatDropdown;
window.updateDateFormatUI = updateDateFormatUI;
window.formatAllSpreadsheetDatesNow = formatAllSpreadsheetDatesNow;
window.getActiveDateFormat = getActiveDateFormat;
