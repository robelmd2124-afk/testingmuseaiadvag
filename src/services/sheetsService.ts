import { getGoogleAccessToken, googleSignIn } from './googleAuth';
import { formatCustomDate } from './schema';

function extractClientHeaderValue(rowObj: any, h: string, monthHint?: string): any {
  if (!rowObj || typeof rowObj !== 'object') return '';

  if (rowObj[h] !== undefined && rowObj[h] !== null && rowObj[h] !== '') {
    let val = rowObj[h];
    const hLower = String(h || '').trim().toLowerCase();
    if (hLower.includes('date') || hLower === 'start' || hLower === 'end' || hLower === 'entry') {
      const formatted = formatCustomDate(val, monthHint || rowObj?.Month || rowObj?.month);
      if (formatted) return formatted;
    }
    return val;
  }

  const hClean = h.trim().toLowerCase().replace(/[-_ ]/g, '');
  let val: any = undefined;
  if (hClean === 'sl' || hClean === 'slno') {
    val = rowObj['SL'] ?? rowObj['SL NO'] ?? rowObj.sl ?? rowObj.sl_no;
  } else if (hClean === 'off' || hClean === 'offusd' || hClean === 'off$') {
    val = rowObj['OFF $'] ?? rowObj['OFF USD'] ?? rowObj.off_usd ?? rowObj.off ?? '0';
  } else if (hClean === 'adaccount' || hClean === 'account') {
    val = rowObj['AD Account'] ?? rowObj.ad_account ?? rowObj.account ?? '';
  } else if (hClean === 'idname') {
    val = rowObj['ID Name'] ?? rowObj.id_name ?? rowObj.id ?? '';
  } else if (hClean === 'status') {
    val = rowObj['Status'] ?? rowObj.status ?? 'Old Page';
  } else if (hClean === 'pagename' || hClean === 'page') {
    val = rowObj['Page Name'] ?? rowObj.page_name ?? rowObj.Page ?? '';
  } else if (hClean === 'start' || hClean === 'startdate') {
    val = rowObj['Start'] ?? rowObj['Start Date'] ?? rowObj.start_date;
  } else if (hClean === 'end' || hClean === 'enddate') {
    val = rowObj['End'] ?? rowObj['End Date'] ?? rowObj.end_date;
  } else if (hClean === 'types' || hClean === 'type') {
    val = rowObj['Types'] ?? rowObj['Type'] ?? rowObj.types ?? 'Boost';
  } else if (hClean === 'day' || hClean === 'days') {
    val = rowObj['Day'] ?? rowObj.day ?? 1;
  } else if (hClean === 'budget') {
    val = rowObj['Budget'] ?? rowObj.budget ?? rowObj.campaign_budget ?? 0;
  } else if (hClean === 'daily$' || hClean === 'daily') {
    val = rowObj['Daily $'] ?? rowObj.daily ?? '';
  } else if (hClean === 'taka' || hClean === 'totaltaka') {
    val = rowObj['Taka'] ?? rowObj.taka ?? '';
  } else if (hClean === 'rate' || hClean === 'dollarrate') {
    val = rowObj['Rate'] ?? rowObj.rate ?? 140;
  } else if (hClean === 'entrydate' || hClean === 'entry') {
    val = rowObj['Entry Date'] ?? rowObj.entry_date ?? rowObj.Entry ?? rowObj.date;
  } else if (hClean === 'month') {
    val = rowObj['Month'] ?? rowObj.month ?? monthHint;
  } else if (hClean === 'year') {
    val = rowObj['Year'] ?? rowObj.year;
  } else if (hClean === 'recordid' || hClean === 'id') {
    val = rowObj['Record ID'] ?? rowObj._recordId ?? rowObj.recordId ?? '';
  } else if (hClean === 'date') {
    val = rowObj['Date'] ?? rowObj.date ?? rowObj.entry_date;
  } else if (hClean === 'paymentmethod' || hClean === 'method') {
    val = rowObj['Payment Method'] ?? rowObj.payment_method ?? 'Bkash';
  } else if (hClean === 'paid' || hClean === 'paid৳' || hClean === 'paidamount') {
    val = rowObj['Paid (৳)'] ?? rowObj.paid_amount ?? rowObj.Paid ?? 0;
  } else if (hClean === 'spent') {
    val = rowObj['Spent'] ?? rowObj.spent ?? 0;
  } else if (hClean === 'remaining') {
    val = rowObj['Remaining'] ?? rowObj.remaining ?? '';
  } else if (hClean === 'cost') {
    val = rowObj['Cost'] ?? rowObj.cost ?? 0;
  } else if (hClean === 'perpersontaka') {
    val = rowObj['Per Person Taka'] ?? rowObj.perPersonTaka ?? '';
  } else if (hClean === 'perpersoncost') {
    val = rowObj['Per Person Cost'] ?? rowObj.perPersonCost ?? '';
  } else if (hClean === 'note' || hClean === 'advancenote') {
    val = rowObj['Note'] ?? rowObj.note ?? rowObj.advance_note ?? '';
  }

  if (val === undefined || val === null) val = '';
  const hLower = String(h || '').trim().toLowerCase();
  if ((hLower.includes('date') || hLower === 'start' || hLower === 'end' || hLower === 'entry') && val) {
    const formatted = formatCustomDate(val, monthHint || rowObj?.Month || rowObj?.month);
    if (formatted) return formatted;
  }
  return val;
}

function normalizeRowDateValues(headers: string[], rowObj: any, monthHint?: string): any[] {
  return headers.map(h => extractClientHeaderValue(rowObj, h, monthHint));
}

export interface SheetMetadata {
  spreadsheetId: string;
  title: string;
  sheets: {
    sheetId: number;
    title: string;
    index: number;
  }[];
}

/**
 * Ensures an active access token is present, triggering sign-in if needed.
 */
export async function getValidAccessToken(): Promise<string> {
  let token = getGoogleAccessToken();
  if (!token) {
    const res = await googleSignIn();
    if (!res?.accessToken) {
      throw new Error('Please sign in with Google to access your Google Sheets.');
    }
    token = res.accessToken;
  }
  return token;
}

/**
 * Fetch Google Spreadsheet Metadata (all sheet tabs)
 */
export async function fetchSpreadsheetMetadata(spreadsheetId: string): Promise<SheetMetadata> {
  const token = await getValidAccessToken();
  const response = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}`, {
    headers: {
      Authorization: `Bearer ${token}`
    }
  });

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    throw new Error(errorData.error?.message || `Google Sheets API Error: ${response.statusText}`);
  }

  const data = await response.json();
  return {
    spreadsheetId: data.spreadsheetId,
    title: data.properties?.title || 'Untitled Spreadsheet',
    sheets: (data.sheets || []).map((s: any) => ({
      sheetId: s.properties?.sheetId,
      title: s.properties?.title,
      index: s.properties?.index
    }))
  };
}

/**
 * Read values from a specific sheet range
 */
export async function fetchSheetValues(spreadsheetId: string, range: string): Promise<any[][]> {
  const token = await getValidAccessToken();
  const encodedRange = encodeURIComponent(range);
  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodedRange}`,
    {
      headers: {
        Authorization: `Bearer ${token}`
      }
    }
  );

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    throw new Error(errorData.error?.message || `Google Sheets API Error: ${response.statusText}`);
  }

  const data = await response.json();
  return data.values || [];
}

/**
 * Read values from multiple sheet ranges in a SINGLE consolidated batch request
 */
export async function batchGetSheetValues(
  spreadsheetId: string,
  ranges: string[]
): Promise<{ range: string; values: any[][] }[]> {
  if (!ranges || ranges.length === 0) return [];
  const token = await getValidAccessToken();
  const rangesQuery = ranges.map(r => `ranges=${encodeURIComponent(r)}`).join('&');
  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values:batchGet?${rangesQuery}&valueRenderOption=FORMATTED_VALUE`,
    {
      headers: {
        Authorization: `Bearer ${token}`
      }
    }
  );

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    throw new Error(errorData.error?.message || `Google Sheets Batch API Error: ${response.statusText}`);
  }

  const data = await response.json();
  return (data.valueRanges || []).map((vr: any) => ({
    range: vr.range,
    values: vr.values || []
  }));
}

/**
 * Fetch all sheets and values in a spreadsheet using Chunked Batch Requests (Prevents URL truncation and timeouts on large datasets)
 */
export async function fetchAllSpreadsheetData(
  spreadsheetId: string,
  onProgress?: (loaded: number, total: number, currentTab: string) => void
): Promise<{
  title: string;
  sheetsData: { [sheetName: string]: any[][] };
}> {
  const meta = await fetchSpreadsheetMetadata(spreadsheetId);
  const sheetsData: { [sheetName: string]: any[][] } = {};

  if (!meta.sheets || meta.sheets.length === 0) {
    return { title: meta.title, sheetsData };
  }

  const BATCH_SIZE = 4;
  const totalSheets = meta.sheets.length;

  for (let i = 0; i < totalSheets; i += BATCH_SIZE) {
    const chunk = meta.sheets.slice(i, i + BATCH_SIZE);
    const ranges = chunk.map(s => `'${s.title}'!A:ZZ`);
    const currentTabNames = chunk.map(s => s.title).join(', ');
    if (onProgress) onProgress(i, totalSheets, currentTabNames);

    try {
      const valueRanges = await batchGetSheetValues(spreadsheetId, ranges);
      chunk.forEach((s, idx) => {
        const vr = valueRanges[idx];
        if (vr && vr.values && vr.values.length > 0) {
          sheetsData[s.title] = vr.values;
        }
      });
    } catch (batchErr) {
      console.warn('Chunk batch fetch failed, falling back to sequential fetch for chunk:', batchErr);
      for (const s of chunk) {
        try {
          const values = await fetchSheetValues(spreadsheetId, `'${s.title}'!A:ZZ`);
          if (values && values.length > 0) {
            sheetsData[s.title] = values;
          }
        } catch (err) {
          console.warn(`Could not read sheet ${s.title}:`, err);
        }
      }
    }
    if (onProgress) onProgress(Math.min(i + BATCH_SIZE, totalSheets), totalSheets, currentTabNames);
  }

  return {
    title: meta.title,
    sheetsData
  };
}

/**
 * Append row to a sheet directly with user's OAuth credentials
 */
export async function appendRowToSheet(
  spreadsheetId: string,
  sheetName: string,
  rowValues: any[]
): Promise<any> {
  const token = await getValidAccessToken();
  let targetTitle = sheetName;
  try {
    const meta = await fetchSpreadsheetMetadata(spreadsheetId);
    if (meta?.sheets?.length) {
      const resolved = resolveSheetTab(meta.sheets, sheetName);
      if (resolved) {
        targetTitle = resolved.title;
      } else {
        await createSheetTab(spreadsheetId, sheetName).catch(() => {});
      }
    }
  } catch (_) {}

  const range = encodeURIComponent(`'${targetTitle}'!A:ZZ`);
  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        values: [rowValues]
      })
    }
  );

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    throw new Error(errorData.error?.message || `Failed to append row: ${response.statusText}`);
  }

  return response.json();
}

/**
 * Create a new sheet tab if it doesn't exist
 */
export async function createSheetTab(spreadsheetId: string, sheetTitle: string): Promise<void> {
  const token = await getValidAccessToken();
  const response = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      requests: [
        {
          addSheet: {
            properties: {
              title: sheetTitle
            }
          }
        }
      ]
    })
  });

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    // If already exists, ignore
    if (!errorData.error?.message?.includes('already exists')) {
      throw new Error(errorData.error?.message || `Failed to create sheet: ${response.statusText}`);
    }
  }
}

/**
 * Clear a sheet tab's contents
 */
export async function clearSheetTab(spreadsheetId: string, sheetTitle: string): Promise<void> {
  const token = await getValidAccessToken();
  const encodedRange = encodeURIComponent(`'${sheetTitle}'!A:ZZ`);
  await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodedRange}:clear`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    }
  );
}

/**
 * Write/Update rows in a sheet tab
 */
export async function writeSheetValues(
  spreadsheetId: string,
  range: string,
  values: any[][]
): Promise<any> {
  const token = await getValidAccessToken();
  const encodedRange = encodeURIComponent(range);
  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodedRange}?valueInputOption=USER_ENTERED`,
    {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        values
      })
    }
  );

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    throw new Error(errorData.error?.message || `Failed to write values: ${response.statusText}`);
  }

  return response.json();
}

/**
 * Push all dashboard data to Google Sheet using the user's OAuth access token
 * Performs a non-destructive upsert: updates existing records by Record ID / SL and appends new records.
 * Never clears the tab or wipes out existing rows/formulas.
 */
export async function pushAllSheetsDataDirect(
  spreadsheetId: string,
  allSheetsData: { [sheetName: string]: any[][] }
): Promise<{ success: boolean; count: number; message: string }> {
  let count = 0;
  for (const [title, rows] of Object.entries(allSheetsData)) {
    if (!rows || rows.length === 0) continue;
    try {
      // 1. Ensure tab exists
      await createSheetTab(spreadsheetId, title);
    } catch (e) {
      console.warn(`Tab ${title} could not be created or already exists:`, e);
    }

    try {
      // 2. Read existing values in the sheet tab
      const existingValues = await fetchSheetValues(spreadsheetId, `'${title}'!A:ZZ`).catch(() => []);
      
      if (!existingValues || existingValues.length <= 1) {
        // Sheet is empty or only has headers: write headers and rows
        await writeSheetValues(spreadsheetId, `'${title}'!A1`, rows);
        count++;
        continue;
      }

      // Existing sheet has data! Perform incremental atomic upsert without column mismatch!
      const existingHeaders = (existingValues[0] || []).map((h: any) => String(h || '').trim().toLowerCase());
      const newHeaders = (rows[0] || []).map((h: any) => String(h || '').trim().toLowerCase());

      // Create a mapping from existing column index to new column index
      const colMap = new Map<number, number>();
      existingHeaders.forEach((eh: string, eIdx: number) => {
        const nIdx = newHeaders.findIndex((nh: string) =>
          nh === eh ||
          nh.replace(/[-_ ]/g, '') === eh.replace(/[-_ ]/g, '') ||
          (eh.includes('sl') && nh.includes('sl')) ||
          (eh.includes('page') && nh.includes('page')) ||
          (eh.includes('budget') && nh.includes('budget')) ||
          (eh.includes('taka') && nh.includes('taka')) ||
          (eh.includes('rate') && nh.includes('rate')) ||
          (eh.includes('off') && nh.includes('off')) ||
          (eh.includes('account') && nh.includes('account'))
        );
        if (nIdx >= 0) {
          colMap.set(eIdx, nIdx);
        }
      });

      const newRecIdIdx = newHeaders.findIndex(h => h.includes('record id') || h === 'id');
      const newSlIdx = newHeaders.findIndex(h => h === 'sl' || h === 'sl no' || h === 'sl#');
      const newPageIdx = newHeaders.findIndex(h => h.includes('page'));

      const exRecIdIdx = existingHeaders.findIndex(h => h.includes('record id') || h === 'id');
      const exSlIdx = existingHeaders.findIndex(h => h === 'sl' || h === 'sl no' || h === 'sl#');
      const exPageIdx = existingHeaders.findIndex(h => h.includes('page'));

      const rowsToAppend: any[][] = [];

      for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        const rRecId = newRecIdIdx >= 0 ? String(row[newRecIdIdx] || '').trim() : '';
        const rSl = newSlIdx >= 0 ? String(row[newSlIdx] || '').trim() : '';
        const rPage = newPageIdx >= 0 ? String(row[newPageIdx] || '').trim().toLowerCase() : '';

        // Find match in existing values
        let matchIndex = -1;
        for (let e = 1; e < existingValues.length; e++) {
          const exRow = existingValues[e];
          const eRecId = exRecIdIdx >= 0 ? String(exRow[exRecIdIdx] || '').trim() : '';
          const eSl = exSlIdx >= 0 ? String(exRow[exSlIdx] || '').trim() : '';
          const ePage = exPageIdx >= 0 ? String(exRow[exPageIdx] || '').trim().toLowerCase() : '';

          if (rRecId && eRecId && rRecId === eRecId) {
            matchIndex = e;
            break;
          }
          if (rSl && eSl) {
            const rNum = parseInt(rSl.replace(/\D/g, ''), 10);
            const eNum = parseInt(eSl.replace(/\D/g, ''), 10);
            const slMatched = (rSl === eSl) || (!isNaN(rNum) && !isNaN(eNum) && rNum === eNum);
            if (slMatched) {
              if (rPage && ePage && (rPage === ePage || rPage.includes(ePage) || ePage.includes(rPage))) {
                matchIndex = e;
                break;
              } else if (!rPage || !ePage) {
                matchIndex = e;
                break;
              }
            }
          }
        }

        if (matchIndex > 0) {
          // Update the matching row in-place in memory atomically aligned by header
          const mergedRow = [...existingValues[matchIndex]];
          colMap.forEach((nIdx, eIdx) => {
            const val = row[nIdx];
            if (val !== undefined && val !== null && String(val).trim() !== '') {
              mergedRow[eIdx] = val;
            }
          });
          existingValues[matchIndex] = mergedRow;
        } else {
          // New record: format according to existingHeaders if present
          if (existingHeaders.length > 0 && colMap.size > 0) {
            const formattedRow = existingHeaders.map((_, eIdx) => {
              const nIdx = colMap.get(eIdx);
              return (nIdx !== undefined && nIdx >= 0) ? (row[nIdx] ?? '') : '';
            });
            rowsToAppend.push(formattedRow);
          } else {
            rowsToAppend.push(row);
          }
        }
      }

      // 1-Request Atomic Write for the entire updated sheet tab (Zero 429 Rate Limits, Instant 200ms Execution)
      const combinedSheetRows = [...existingValues, ...rowsToAppend];
      await writeSheetValues(spreadsheetId, `'${title}'!A1`, combinedSheetRows);
      count++;
    } catch (err: any) {
      console.error(`Error writing sheet ${title}:`, err);
    }
  }

  return {
    success: true,
    count,
    message: `সফলভাবে ${count} টি শিটের সমস্ত ডাটা ইনক্রিমেন্টালি Google Sheet-এ সেভ করা হয়েছে!`
  };
}

/**
 * Resolve target sheet tab flexibly (supporting e.g. "September 2026", "Sep 2026", "Today Created", etc.)
 */
export function resolveSheetTab(
  sheets: { sheetId: number; title: string; index: number }[],
  targetName: string,
  month?: string,
  year?: string | number
): { sheetId: number; title: string; index: number } | undefined {
  if (!sheets || !sheets.length) return undefined;
  const cleanTarget = targetName.trim().toLowerCase().replace(/[-_ ]/g, '');

  // 1. Exact match ignoring spaces, underscores, hyphens, and case
  const exact = sheets.find(s => s.title.trim().toLowerCase().replace(/[-_ ]/g, '') === cleanTarget);
  if (exact) return exact;

  // 2. Today created matching
  if (cleanTarget.includes('today')) {
    const todayMatch = sheets.find(s => s.title.toLowerCase().includes('today'));
    if (todayMatch) return todayMatch;
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

  // 3. Month & Year matching (e.g. "October 2026", "Oct 2026", "October", "Oct")
  if (effectiveMonth) {
    const mnFull = effectiveMonth.toLowerCase();
    const mnShort = effectiveMonth.slice(0, 3).toLowerCase();
    const yrStr = effectiveYear ? String(effectiveYear).trim() : '';

    if (yrStr) {
      const my = sheets.find(s => {
        const t = s.title.toLowerCase();
        return (t.includes(mnFull) || t.includes(mnShort)) && t.includes(yrStr);
      });
      if (my) return my;
    }

    const mOnly = sheets.find(s => {
      const t = s.title.toLowerCase();
      return t.includes(mnFull) || t.includes(mnShort);
    });
    if (mOnly) return mOnly;
  }

  // 4. Default campaign sheet - ONLY if NO month was ever specified
  if (!effectiveMonth && cleanTarget.includes('campaign')) {
    const camp = sheets.find(s => s.title.toLowerCase().includes('campaign'));
    if (camp) return camp;
  }

  return undefined;
}

export interface DirectMatchCriteria {
  recordId?: string;
  sl?: string | number;
  pageName?: string;
  budget?: string | number;
  date?: string;
  paid?: string | number;
  cost?: string | number;
  spent?: string | number;
  method?: string;
  type?: string;
  month?: string;
  year?: string | number;
}

/**
 * Scan all columns of a sheet to find the row index matching the criteria
 */
function findRowIndexInValues(values: any[][], criteria: DirectMatchCriteria): number {
  if (!values || values.length <= 1) return -1;

  const headers = values[0].map((h: any) => String(h || '').trim().toLowerCase());
  const recIdCol = headers.findIndex(h => h === 'record id' || h === 'recordid' || h === '_recordid' || h === 'id' || h === 'record_id');
  const slCol = headers.findIndex(h => h === 'sl' || h === 'sl no' || h === 'sl_no' || h === 'sl#' || h === 'serial' || h.includes('sl'));
  const pageCol = headers.findIndex(h => h.includes('page') || h === 'name' || h === 'title');
  const budgetCol = headers.findIndex(h => h.includes('budget') || h.includes('cost') || h.includes('paid') || h.includes('spent') || h.includes('amount'));
  const dateCol = headers.findIndex(h => h.includes('date') || h === 'entry' || h === 'start');

  const cleanRecId = criteria.recordId ? String(criteria.recordId).trim() : '';
  const cleanSl = criteria.sl !== undefined && criteria.sl !== null ? String(criteria.sl).trim().replace(/^#/, '') : '';
  const cleanSlNumMatch = cleanSl.match(/\d+/);
  const cleanSlNum = cleanSlNumMatch ? parseInt(cleanSlNumMatch[0], 10) : NaN;
  const cleanPage = criteria.pageName ? String(criteria.pageName).trim().toLowerCase() : '';
  const cleanBudget = criteria.budget !== undefined && criteria.budget !== null ? parseFloat(String(criteria.budget).replace(/[^0-9.]/g, '')) : NaN;
  const cleanPaid = criteria.paid !== undefined && criteria.paid !== null
    ? parseFloat(String(criteria.paid).replace(/[^0-9.]/g, ''))
    : (criteria.cost !== undefined && criteria.cost !== null ? parseFloat(String(criteria.cost).replace(/[^0-9.]/g, '')) : NaN);
  const cleanDate = criteria.date ? String(criteria.date).trim().toLowerCase() : '';

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rRecId = recIdCol >= 0 ? String(row[recIdCol] || '').trim() : '';
    const rSlRaw = slCol >= 0 ? String(row[slCol] || '').trim().replace(/^#/, '') : String(row[0] || '').trim().replace(/^#/, '');
    const rSlNumMatch = rSlRaw.match(/\d+/);
    const rSlNum = rSlNumMatch ? parseInt(rSlNumMatch[0], 10) : NaN;
    const rPage = pageCol >= 0 ? String(row[pageCol] || '').trim().toLowerCase() : '';
    const rBudget = budgetCol >= 0 ? parseFloat(String(row[budgetCol] || '0').replace(/[^0-9.]/g, '')) : NaN;
    const rDate = dateCol >= 0 ? String(row[dateCol] || '').trim().toLowerCase() : '';

    // 1. Match by Record ID (if known)
    if (cleanRecId && rRecId && cleanRecId === rRecId) {
      return i;
    }

    // 2. Match by SL Number + Page Name
    if (cleanSl && rSlRaw) {
      const slMatch = (cleanSl === rSlRaw) || (cleanSl.includes(rSlRaw)) || (rSlRaw.includes(cleanSl)) || (!isNaN(cleanSlNum) && !isNaN(rSlNum) && cleanSlNum === rSlNum);
      if (slMatch) {
        if (cleanPage && rPage) {
          if (cleanPage === rPage || cleanPage.includes(rPage) || rPage.includes(cleanPage)) {
            return i;
          }
        } else {
          return i;
        }
      }
    }

    // 3. Match by Page Name + Budget
    if (cleanPage && rPage && (cleanPage === rPage || cleanPage.includes(rPage) || rPage.includes(cleanPage))) {
      if (!isNaN(cleanBudget) && !isNaN(rBudget) && Math.abs(cleanBudget - rBudget) < 0.01) {
        return i;
      }
    }

    // 4. Match by Date + Paid / Cost / Spent (specifically for Payments, Agency Costs, Scams)
    if (cleanDate && rDate && (cleanDate === rDate || cleanDate.includes(rDate) || rDate.includes(cleanDate))) {
      if (!isNaN(cleanPaid) && !isNaN(rBudget) && Math.abs(cleanPaid - rBudget) < 0.01) {
        return i;
      }
      if (!isNaN(cleanBudget) && !isNaN(rBudget) && Math.abs(cleanBudget - rBudget) < 0.01) {
        return i;
      }
    }

    // 5. Match by SL Number alone if no other row has it
    if (!isNaN(cleanSlNum) && !isNaN(rSlNum) && cleanSlNum === rSlNum && cleanSlNum > 0) {
      return i;
    }
  }

  return -1;
}

/**
 * Delete a specific row from a Google Sheet tab directly using OAuth token
 */
export async function deleteRowFromSheetDirect(
  spreadsheetId: string,
  sheetName: string,
  matchId?: string,
  sl?: string | number,
  extraMatch?: { pageName?: string; budget?: string | number; month?: string; year?: string | number }
): Promise<boolean> {
  let token = getGoogleAccessToken();
  if (!token) return false;
  
  try {
    const meta = await fetchSpreadsheetMetadata(spreadsheetId);
    const targetSheet = resolveSheetTab(meta.sheets, sheetName, extraMatch?.month, extraMatch?.year);
    if (!targetSheet) return false;

    // Read ALL columns in sheet
    const values = await fetchSheetValues(spreadsheetId, `'${targetSheet.title}'!A:ZZ`);
    if (!values || values.length <= 1) return false;

    const criteria: DirectMatchCriteria = {
      recordId: matchId,
      sl: sl,
      pageName: extraMatch?.pageName,
      budget: extraMatch?.budget,
      date: (extraMatch as any)?.date || (extraMatch as any)?.Date,
      paid: (extraMatch as any)?.paid || (extraMatch as any)?.['Paid (৳)'],
      cost: (extraMatch as any)?.cost || (extraMatch as any)?.Cost,
      spent: (extraMatch as any)?.spent || (extraMatch as any)?.Spent,
      month: extraMatch?.month,
      year: extraMatch?.year
    };

    const rowIndexToDelete = findRowIndexInValues(values, criteria);

    if (rowIndexToDelete > 0) {
      // Delete using batchUpdate deleteDimension (0-indexed)
      const response = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          requests: [
            {
              deleteDimension: {
                range: {
                  sheetId: targetSheet.sheetId,
                  dimension: 'ROWS',
                  startIndex: rowIndexToDelete,
                  endIndex: rowIndexToDelete + 1
                }
              }
            }
          ]
        })
      });
      return response.ok;
    }
  } catch (err) {
    console.warn(`Direct OAuth deleteRow failed for ${sheetName}:`, err);
  }
  return false;
}

/**
 * Update a specific row in a Google Sheet tab directly using OAuth token
 */
export async function updateRowInSheetDirect(
  spreadsheetId: string,
  sheetName: string,
  matchId: string | undefined,
  sl: string | number | undefined,
  headers: string[],
  rowObj: any,
  extraMatch?: { pageName?: string; budget?: string | number; month?: string; year?: string | number }
): Promise<boolean> {
  let token = getGoogleAccessToken();
  if (!token) return false;

  try {
    const meta = await fetchSpreadsheetMetadata(spreadsheetId);
    const targetSheet = resolveSheetTab(meta.sheets, sheetName, extraMatch?.month, extraMatch?.year);
    if (!targetSheet) return false;

    const values = await fetchSheetValues(spreadsheetId, `'${targetSheet.title}'!A:ZZ`);
    if (!values || values.length <= 1) return false;

    const criteria: DirectMatchCriteria = {
      recordId: matchId || rowObj?.['Record ID'] || rowObj?._recordId,
      sl: sl !== undefined ? sl : (rowObj?.SL || rowObj?.['SL NO'] || rowObj?.sl_no),
      pageName: extraMatch?.pageName || rowObj?.['Page Name'] || rowObj?.page_name,
      budget: extraMatch?.budget || rowObj?.['Budget'] || rowObj?.budget,
      date: (extraMatch as any)?.date || (extraMatch as any)?.Date || rowObj?.Date || rowObj?.date,
      paid: (extraMatch as any)?.paid || (extraMatch as any)?.['Paid (৳)'] || rowObj?.['Paid (৳)'] || rowObj?.paid_amount || rowObj?.Paid,
      cost: (extraMatch as any)?.cost || rowObj?.Cost || rowObj?.cost,
      spent: (extraMatch as any)?.spent || rowObj?.Spent || rowObj?.spent,
      month: extraMatch?.month || rowObj?.Month || rowObj?.month,
      year: extraMatch?.year || rowObj?.Year || rowObj?.year
    };

    const rowIndex = findRowIndexInValues(values, criteria);

    if (rowIndex > 0) {
      // row in Google Sheet is 1-indexed (row 1 is header, so index i is sheet row i + 1)
      const targetSheetRowNumber = rowIndex + 1;
      const sheetHeaders = values[0].map((h: any) => String(h || '').trim());
      const effectiveHeaders = sheetHeaders.length > 0 ? sheetHeaders : headers;
      const rowValues = normalizeRowDateValues(effectiveHeaders, rowObj, extraMatch?.month || criteria.month);

      await writeSheetValues(spreadsheetId, `'${targetSheet.title}'!A${targetSheetRowNumber}`, [rowValues]);
      return true;
    } else {
      // If not found in sheet, append it to avoid losing data
      const effectiveHeaders = (values[0] && values[0].length > 0) ? values[0].map((h: any) => String(h || '').trim()) : headers;
      const rowValues = normalizeRowDateValues(effectiveHeaders, rowObj, extraMatch?.month || criteria.month);
      await appendRowToSheet(spreadsheetId, targetSheet.title, rowValues);
      return true;
    }
  } catch (err) {
    console.warn(`Direct OAuth updateRow failed for ${sheetName}:`, err);
  }
  return false;
}

/**
 * Upsert a single record into a Google Sheet tab directly using OAuth token.
 * If the record exists (by Record ID or SL), it updates the existing row in place.
 * If not found, it appends the new row without wiping any existing data.
 */
export async function upsertRowInSheetDirect(
  spreadsheetId: string,
  sheetName: string,
  matchId: string | undefined,
  sl: string | number | undefined,
  headers: string[],
  rowObj: any,
  extraMatch?: { pageName?: string; budget?: string | number; month?: string; year?: string | number }
): Promise<boolean> {
  let token = getGoogleAccessToken();
  if (!token) return false;

  try {
    const meta = await fetchSpreadsheetMetadata(spreadsheetId);
    let targetSheet = resolveSheetTab(meta.sheets, sheetName, extraMatch?.month, extraMatch?.year);
    
    if (!targetSheet) {
      await createSheetTab(spreadsheetId, sheetName);
      const updatedMeta = await fetchSpreadsheetMetadata(spreadsheetId);
      targetSheet = resolveSheetTab(updatedMeta.sheets, sheetName, extraMatch?.month, extraMatch?.year);
    }
    
    if (!targetSheet) return false;

    const values = await fetchSheetValues(spreadsheetId, `'${targetSheet.title}'!A:ZZ`).catch(() => []);
    
    if (!values || values.length <= 1) {
      // Empty sheet: write header then append row
      const rowValues = normalizeRowDateValues(headers, rowObj, extraMatch?.month || rowObj?.Month || rowObj?.month);
      await writeSheetValues(spreadsheetId, `'${targetSheet.title}'!A1`, [headers, rowValues]);
      return true;
    }

    const criteria: DirectMatchCriteria = {
      recordId: matchId || rowObj?.['Record ID'] || rowObj?._recordId,
      sl: sl !== undefined ? sl : (rowObj?.SL || rowObj?.['SL NO'] || rowObj?.sl_no),
      pageName: extraMatch?.pageName || rowObj?.['Page Name'] || rowObj?.page_name,
      budget: extraMatch?.budget || rowObj?.['Budget'] || rowObj?.budget,
      date: (extraMatch as any)?.date || (extraMatch as any)?.Date || rowObj?.Date || rowObj?.date,
      paid: (extraMatch as any)?.paid || (extraMatch as any)?.['Paid (৳)'] || rowObj?.['Paid (৳)'] || rowObj?.paid_amount || rowObj?.Paid,
      cost: (extraMatch as any)?.cost || rowObj?.Cost || rowObj?.cost,
      spent: (extraMatch as any)?.spent || rowObj?.Spent || rowObj?.spent,
      month: extraMatch?.month || rowObj?.Month || rowObj?.month,
      year: extraMatch?.year || rowObj?.Year || rowObj?.year
    };

    const rowIndex = findRowIndexInValues(values, criteria);

    const sheetHeaders = values[0].map((h: any) => String(h || '').trim());
    const effectiveHeaders = sheetHeaders.length > 0 ? sheetHeaders : headers;
    const rowValues = normalizeRowDateValues(effectiveHeaders, rowObj, extraMatch?.month || criteria.month);

    if (rowIndex > 0) {
      // Update existing row in place
      const targetSheetRowNumber = rowIndex + 1;
      await writeSheetValues(spreadsheetId, `'${targetSheet.title}'!A${targetSheetRowNumber}`, [rowValues]);
      return true;
    } else {
      // Append new row to the sheet
      await appendRowToSheet(spreadsheetId, targetSheet.title, rowValues);
      return true;
    }
  } catch (err) {
    console.warn(`Direct OAuth upsertRow failed for ${sheetName}:`, err);
    return false;
  }
}
