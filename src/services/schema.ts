export const MONTHS=['January','February','March','April','May','June','July','August','September','October','November','December'];
export const CH=['SL','OFF $','AD Account','ID Name','Status','Page Name','Start','End','Types','Day','Budget','Daily $','Taka','Rate','Entry Date','Month','Year','Record ID'];
export const PAYMENT_HEADERS=['Date','Payment Method','Paid (৳)','Month','Year','Note'];
export const TODAY_HEADERS=['SL','OFF $','AD Account','ID Name','Status','Page Name','Start','End','Types','Day','Budget','Daily $','Taka','Rate','Entry Date','Month','Year','Record ID'];
export const PAGE_SUMMARY_HEADERS=['Page Name','Total Ads','Total Budget $'];
export const AGENCY_SPENDING_HEADERS=['Month','Year','Total Spend $'];
export const MONTHLY_CLOSING_HEADERS=['Month','Year','Status','Closed At','Closed By','Total Campaign','Total Budget $','Total Campaign Taka','Total Paid (৳)','Total Due (৳)','Agency Cost (৳)','Rate'];
export const AGENCY_COST_HEADERS=['Date','Month','Year','Type','Cost','Per Person Cost','Status','Record ID'];
export const END_ADS_HEADERS=['Page Name','Start Date','End Date','Day','Budget','Type'];
export const AD_ACCOUNT_BILLING_HEADERS=['Ad Account','ID Name','Due ($)','Threshold ($)','Remaining ($)','Billing Date','Active Spend ($)','Active Ads','Card / Method','Status','Last Checked'];
export const MENU=[
  {key:'today_created_campaign',label:'Today Created',custom:1},
  {key:'end_ads',label:'End Ads',custom:9},
  {key:'campaign',label:'All Campaign',custom:4},
  {key:'page_summery',label:'Page Summary',custom:3},
  {key:'payment_due_tracker',label:'Payment',custom:2},
  {key:'agency_spending',label:'Agency Calculation',custom:5}
];

export const MONTH_NAMES_SHORT = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
export const MONTH_NAMES_FULL = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

export const ALL_DATE_HEADERS = [
  'Start', 'End', 'Start Date', 'End Date',
  'Entry Date', 'Entry', 'entry_date',
  'Date', 'date', 'Payment Date',
  'Billing Date', 'billing_date',
  'Closed At', 'closed_at',
  'Last Checked'
];

export function isDateHeader(headerName: string): boolean {
  if (!headerName) return false;
  const lower = String(headerName).trim().toLowerCase();
  return (
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

export function normalizeRowDates(row: any, monthHint?: string): any {
  if (!row || typeof row !== 'object') return row;
  const out = { ...row };
  for (const k of Object.keys(out)) {
    if (isDateHeader(k)) {
      const val = out[k];
      if (val !== undefined && val !== null && val !== '') {
        const formatted = formatCustomDate(val, monthHint || out.Month || out.month);
        if (formatted) out[k] = formatted;
      }
    }
  }
  return out;
}

export function parseAnyDate(str: any, monthHint?: string | number): Date | null {
  if (!str && str !== 0) return null;
  if (str instanceof Date) return isNaN(str.getTime()) ? null : str;

  // Handle Excel / Google Sheets numerical date serials (e.g. 45560 = ~2024, 46259 = ~2026)
  if (typeof str === 'number' && str > 30000 && str < 70000) {
    const utcDays = Math.round(str - 25569);
    // Add 12 hours (noon UTC) so local timezone offset never changes the calendar day
    const d = new Date((utcDays * 86400 + 43200) * 1000);
    return isNaN(d.getTime()) ? null : d;
  }

  let s = String(str).trim();
  if (!s) return null;
  const lowerCheck = s.toLowerCase();
  if (lowerCheck === 'undefined' || lowerCheck === 'null' || lowerCheck === 'nan' || lowerCheck === 'invalid date' || lowerCheck === 'loading...') {
    return null;
  }

  // Convert Bengali numerals (০-৯) to Western numerals (0-9)
  const bnToEn: { [key: string]: string } = { '০': '0', '১': '1', '২': '2', '৩': '3', '৪': '4', '৫': '5', '৬': '6', '৭': '7', '৮': '8', '৯': '9' };
  s = s.replace(/[০-৯]/g, d => bnToEn[d] || d);

  // Convert Bengali month names to English
  const bnMonthMap: { [key: string]: string } = {
    'জানুয়ারি': 'Jan', 'জানুয়ারী': 'Jan', 'ফেব্রুয়ারি': 'Feb', 'ফেব্রুয়ারী': 'Feb',
    'মার্চ': 'Mar', 'এপ্রিল': 'Apr', 'মে': 'May', 'জুন': 'Jun',
    'জুলাই': 'Jul', 'আগস্ট': 'Aug', 'সেপ্টেম্বর': 'Sep',
    'অক্টোবর': 'Oct', 'নভেম্বর': 'Nov', 'ডিসেম্বর': 'Dec'
  };
  for (const [bnM, enM] of Object.entries(bnMonthMap)) {
    if (s.includes(bnM)) {
      s = s.replace(bnM, enM);
      break;
    }
  }

  // If numeric string in serial range
  if (/^\d{5}(\.\d+)?$/.test(s)) {
    const num = parseFloat(s);
    if (num > 30000 && num < 70000) {
      const utcDays = Math.round(num - 25569);
      const d = new Date((utcDays * 86400 + 43200) * 1000);
      if (!isNaN(d.getTime())) return d;
    }
  }

  // Match GViz Date constructor: Date(2026, 7, 29) - note month is 0-indexed in GViz
  const matchGviz = s.match(/Date\((\d{4}),\s*(\d{1,2}),\s*(\d{1,2})/i);
  if (matchGviz) {
    return new Date(parseInt(matchGviz[1], 10), parseInt(matchGviz[2], 10), parseInt(matchGviz[3], 10));
  }

  // Clean off timestamp if present: e.g. "25-Aug-2026 12:00:00" -> "25-Aug-2026"
  const cleanDateStr = s.replace(/T\d{2}:\d{2}.*$/, '').replace(/\s+\d{1,2}:\d{2}(:\d{2})?.*$/, '').trim();

  // Match 12-Sep-2026 or 12-September-2026 or 12/Sep/2026 or 12.Sep.2026 or 12-Sep-26
  const matchDMonY = cleanDateStr.match(/^(\d{1,2})[-/ .]([A-Za-z]+)[-/ .](\d{2,4})/);
  if (matchDMonY) {
    const day = parseInt(matchDMonY[1], 10);
    const monStr = matchDMonY[2].toLowerCase();
    const yrStr = matchDMonY[3];
    const yr = yrStr.length === 2 ? 2000 + parseInt(yrStr, 10) : parseInt(yrStr, 10);
    
    let monIdx = MONTH_NAMES_SHORT.findIndex(m => m.toLowerCase() === monStr.slice(0, 3));
    if (monIdx === -1) {
      monIdx = MONTH_NAMES_FULL.findIndex(m => m.toLowerCase() === monStr);
    }
    if (monIdx !== -1 && day >= 1 && day <= 31) {
      return new Date(yr, monIdx, day, 12, 0, 0);
    }
  }

  // Match Sep 25, 2026 or September 25 2026 or Sep-25-2026
  const matchMonDY = cleanDateStr.match(/^([A-Za-z]+)[-/ .](\d{1,2}),?[-/ .](\d{2,4})/);
  if (matchMonDY) {
    const monStr = matchMonDY[1].toLowerCase();
    const day = parseInt(matchMonDY[2], 10);
    const yrStr = matchMonDY[3];
    const yr = yrStr.length === 2 ? 2000 + parseInt(yrStr, 10) : parseInt(yrStr, 10);

    let monIdx = MONTH_NAMES_SHORT.findIndex(m => m.toLowerCase() === monStr.slice(0, 3));
    if (monIdx === -1) {
      monIdx = MONTH_NAMES_FULL.findIndex(m => m.toLowerCase() === monStr);
    }
    if (monIdx !== -1 && day >= 1 && day <= 31) {
      return new Date(yr, monIdx, day, 12, 0, 0);
    }
  }

  // Match YYYY-MM-DD or YYYY/MM/DD or YYYY.MM.DD
  const matchYMD = cleanDateStr.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (matchYMD) {
    return new Date(parseInt(matchYMD[1], 10), parseInt(matchYMD[2], 10) - 1, parseInt(matchYMD[3], 10), 12, 0, 0);
  }

  // Match DD-MM-YYYY or DD/MM/YYYY or MM/DD/YYYY
  const matchDMY = cleanDateStr.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})/);
  if (matchDMY) {
    const p1 = parseInt(matchDMY[1], 10);
    const p2 = parseInt(matchDMY[2], 10);
    let yr = parseInt(matchDMY[3], 10);
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
      // If p1 matches current month (e.g. 10 for October) and p2 is a valid day, treat as MM-DD-YYYY
      const curMonth = (new Date()).getMonth() + 1;
      if (p1 === curMonth && p2 !== curMonth && p2 <= 31) {
        return new Date(yr, p1 - 1, p2, 12, 0, 0);
      }
      // Standard Bangladesh dash/dot format (DD-MM-YYYY)
      return new Date(yr, p2 - 1, p1, 12, 0, 0);
    }
  }

  const standard = new Date(cleanDateStr);
  return isNaN(standard.getTime()) ? null : standard;
}

export const SUPPORTED_DATE_FORMATS = [
  { id: 'DD-Mon-YY', label: '25-Oct-26 (DD-Mon-YY)', example: '25-Oct-26' },
  { id: 'DD-Mon-YYYY', label: '12-Oct-2026 (DD-Mon-YYYY)', example: '12-Oct-2026' },
  { id: 'DD/MM/YYYY', label: '12/10/2026 (DD/MM/YYYY)', example: '12/10/2026' },
  { id: 'YYYY-MM-DD', label: '2026-10-12 (YYYY-MM-DD)', example: '2026-10-12' },
  { id: 'DD-MM-YYYY', label: '12-10-2026 (DD-MM-YYYY)', example: '12-10-2026' },
  { id: 'D-Mon-YYYY', label: '1-Oct-2026 (D-Mon-YYYY)', example: '1-Oct-2026' }
];

let globalActiveDateFormat = 'DD-Mon-YY';

export function setActiveDateFormat(fmt: string) {
  if (fmt && typeof fmt === 'string') {
    globalActiveDateFormat = fmt.trim();
  }
}

export function getActiveDateFormat(): string {
  return globalActiveDateFormat;
}

export function formatCustomDate(dateInput: any, monthHint?: string | number, formatOverride?: string): string {
  if (!dateInput && dateInput !== 0) return '';
  const targetFormat = formatOverride || globalActiveDateFormat || 'DD-Mon-YY';

  let d: Date | null = null;
  if (dateInput instanceof Date) {
    if (isNaN(dateInput.getTime())) return '';
    d = dateInput;
  } else {
    const str = String(dateInput).trim();
    if (!str) return '';
    const lower = str.toLowerCase();
    if (lower === 'undefined' || lower === 'null' || lower === 'nan' || lower === 'invalid date' || lower === 'loading...') return '';
    d = parseAnyDate(str, monthHint);
  }

  if (!d || isNaN(d.getTime())) return '';
  const dayPadded = String(d.getDate()).padStart(2, '0');
  const daySingle = String(d.getDate());
  const monIdx = d.getMonth();
  const monShort = MONTH_NAMES_SHORT[monIdx] || 'Jan';
  const monNum = String(monIdx + 1).padStart(2, '0');
  const yr = String(d.getFullYear());

  switch (targetFormat) {
    case 'DD/MM/YYYY':
      return `${dayPadded}/${monNum}/${yr}`;
    case 'YYYY-MM-DD':
      return `${yr}-${monNum}-${dayPadded}`;
    case 'DD-MM-YYYY':
      return `${dayPadded}-${monNum}-${yr}`;
    case 'D-Mon-YYYY':
      return `${daySingle}-${monShort}-${yr}`;
    case 'DD-Mon-YY':
      return `${dayPadded}-${monShort}-${yr.slice(-2)}`;
    case 'DD-Mon-YYYY':
    default:
      return `${dayPadded}-${monShort}-${yr}`;
  }
}

export function isSameDay(d1Input: any, d2Input: any): boolean {
  const d1 = parseAnyDate(d1Input);
  const d2 = parseAnyDate(d2Input);
  if (!d1 || !d2) return false;
  return d1.getFullYear() === d2.getFullYear() &&
         d1.getMonth() === d2.getMonth() &&
         d1.getDate() === d2.getDate();
}

export function getNowBST(): Date {
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
    const parts: { [key: string]: string } = {};
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

export function getTodayFormatted(): string {
  return formatCustomDate(getNowBST());
}

export function toInputDateValue(dateInput: any): string {
  const d = parseAnyDate(dateInput) || getNowBST();
  const yr = d.getFullYear();
  const mo = String(d.getMonth() + 1).padStart(2, '0');
  const da = String(d.getDate()).padStart(2, '0');
  return `${yr}-${mo}-${da}`;
}

export function calculateEndDate(startDateInput: any, days: number): string {
  const d = parseAnyDate(startDateInput);
  if (!d || isNaN(d.getTime())) return formatCustomDate(startDateInput);
  const numDays = Math.max(0, Number(days) || 0);
  const end = new Date(d.getFullYear(), d.getMonth(), d.getDate() + numDays);
  return formatCustomDate(end);
}

export function formatMonthSL(slInput: any, monthName?: string, year?: string | number): string {
  const raw = String(slInput || '').trim().replace(/^#/, '');
  if (!raw) return '';

  const targetMonth = monthName || 'October';
  const mIdx = MONTHS.findIndex(m => m.toLowerCase() === String(targetMonth).toLowerCase());
  const monthNum = mIdx >= 0 ? mIdx + 1 : 10;
  const monthPrefix = String(monthNum).padStart(2, '0');

  if (raw.length >= 4 && raw.startsWith(monthPrefix)) {
    return raw;
  }

  const numOnly = parseInt(raw.replace(/\D/g, ''), 10);
  if (isNaN(numOnly)) return raw;

  if (numOnly >= 1000 && String(numOnly).startsWith(monthPrefix)) {
    return String(numOnly);
  }

  // If numOnly was 4+ digits from another month (e.g. 1045 but target is September '09' or vice versa), extract sequence
  if (String(numOnly).length >= 4) {
    const seq = parseInt(String(numOnly).slice(2), 10);
    if (!isNaN(seq) && seq > 0) {
      return `${monthPrefix}${String(seq).padStart(2, '0')}`;
    }
  }

  if (numOnly < 1000) {
    return `${monthPrefix}${String(numOnly).padStart(2, '0')}`;
  }

  return `${monthPrefix}${String(numOnly).padStart(2, '0')}`;
}

export function normalizeCampaign(d: any, defaultMonth?: string, defaultYear?: string | number, sl?: number): any {
  const x: any = {};
  CH.forEach(h => {
    x[h] = d[h] ?? d[h.replace(/ /g, '_').toLowerCase()] ?? '';
  });
  const nowBst = getNowBST();
  const targetMonth = d['Month'] || d.month || defaultMonth || MONTHS[nowBst.getMonth()];
  const targetYear = d['Year'] || d.year || defaultYear || nowBst.getFullYear();

  const candidateSl = (sl !== undefined && sl !== null) ? sl : (x['SL'] || d['SL'] || d.sl || d.sl_no || '');
  x['SL'] = formatMonthSL(candidateSl, targetMonth, targetYear);
  x.SL = x['SL'];

  // 1. Entry Date: When the ad entry was created in the app (Today Created ad)
  const monthHint = defaultMonth || d['Month'] || d.month || MONTHS[nowBst.getMonth()];
  const rawEntry = d['Entry Date'] || d.entry_date || d['Date'] || d.date || '';
  const parsedEntry = parseAnyDate(rawEntry, monthHint);
  x['Entry Date'] = parsedEntry ? formatCustomDate(parsedEntry) : formatCustomDate(nowBst);

  // 2. Start Date: Month-wise campaign schedule start date
  const rawStart = d.start_date || d.Start || d['Start Date'] || d['start'] || '';
  const parsedStart = parseAnyDate(rawStart, monthHint);
  if (parsedStart) {
    x['Start'] = formatCustomDate(parsedStart);
  } else if (rawStart) {
    x['Start'] = formatCustomDate(rawStart, monthHint);
  } else {
    x['Start'] = x['Entry Date'] || formatCustomDate(nowBst);
  }

  // 3. Month & Year: Strictly derived from Start Date (Month wise all campaign)
  const finalStartParsed = parseAnyDate(x['Start'], monthHint);
  if (finalStartParsed) {
    x['Month'] = MONTHS[finalStartParsed.getMonth()];
    x['Year'] = String(finalStartParsed.getFullYear());
  } else if (defaultMonth && defaultMonth !== 'All' && defaultMonth !== 'All Months') {
    x['Month'] = defaultMonth;
    x['Year'] = String(defaultYear || nowBst.getFullYear());
  } else {
    x['Month'] = d['Month'] || d.month || MONTHS[nowBst.getMonth()];
    x['Year'] = String(d['Year'] || d.year || nowBst.getFullYear());
  }

  // Extract OFF $ robustly
  const rawOff = d['OFF $'] !== undefined && d['OFF $'] !== null && d['OFF $'] !== '' ? d['OFF $']
    : (d['OFF USD'] !== undefined && d['OFF USD'] !== null && d['OFF USD'] !== '' ? d['OFF USD']
    : (d.off_usd !== undefined && d.off_usd !== null && d.off_usd !== '' ? d.off_usd
    : (d.off !== undefined && d.off !== null && d.off !== '' ? d.off
    : (d.OFF !== undefined && d.OFF !== null && d.OFF !== '' ? d.OFF : (x['OFF $'] || '')))));
  x['OFF $'] = (rawOff !== '' && rawOff !== undefined && rawOff !== null) ? String(rawOff).trim() : '';
  x['OFF USD'] = x['OFF $'];

  // Strict Budget Validation: Critical fields like Budget and Taka must never become 0 or empty
  let rawBud = String(d.budget ?? d.Budget ?? x['Budget'] ?? '').replace(/[^0-9.]/g, '');
  let parsedBud = parseFloat(rawBud);
  if (isNaN(parsedBud) || parsedBud <= 0) {
    // If updating a record and previous valid budget exists, preserve it
    const prevBud = parseFloat(String(d.previousBudget || d._prevBudget || x['Budget'] || '').replace(/[^0-9.]/g, ''));
    parsedBud = (!isNaN(prevBud) && prevBud > 0) ? prevBud : 10;
  }
  x['Budget'] = parsedBud;

  const rawDay = String(d.day ?? d.Day ?? x['Day'] ?? '1').replace(/[^0-9]/g, '');
  x['Day'] = Math.max(1, parseInt(rawDay, 10) || 1);

  const rawRate = String(d.rate ?? d.Rate ?? d['Dollar Rate (৳)'] ?? x['Rate'] ?? '140').replace(/[^0-9.]/g, '');
  const parsedRate = parseFloat(rawRate);
  x['Rate'] = (!isNaN(parsedRate) && parsedRate > 0) ? parsedRate : 140;

  // Strict Taka Validation: Always compute from Budget * Rate if empty, 0, or invalid
  const rawTaka = String(d.taka ?? d.Taka ?? d['Total Taka'] ?? '').replace(/[^0-9.]/g, '');
  const parsedTaka = parseFloat(rawTaka);
  const calculatedTaka = Number(x['Budget']) * Number(x['Rate']);
  x['Taka'] = (!isNaN(parsedTaka) && parsedTaka > 0)
    ? parsedTaka.toFixed(2)
    : (calculatedTaka > 0 ? calculatedTaka.toFixed(2) : (Number(x['Budget']) * 140).toFixed(2));

  x['Types'] = d.types || d.Types || x['Types'] || 'Boost';
  x['Status'] = d.status || d.Status || x['Status'] || 'Old Page';
  x['AD Account'] = d.ad_account || d['AD Account'] || x['AD Account'] || '';
  x['ID Name'] = d.id_name || d['ID Name'] || x['ID Name'] || '';
  x['Page Name'] = d.page_name || d['Page Name'] || x['Page Name'] || '';
  x['Daily $'] = x['Day'] > 0 ? (Number(x['Budget']) / Number(x['Day'])).toFixed(2) : Number(x['Budget']).toFixed(2);

  // 4. End Date: Start Date + Day
  x['End'] = calculateEndDate(x['Start'], x['Day']);
  x['Record ID'] = d['Record ID'] || d._recordId || crypto.randomUUID();
  if (d._row !== undefined) {
    x._row = d._row;
  }

  return x;
}
