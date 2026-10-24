/**
 * AdVerse Agency Google Sheet 2-Way Sync Connector (v2.0)
 * 
 * SETUP INSTRUCTIONS:
 * 1. In your Google Spreadsheet, click: Extensions > Apps Script
 * 2. Delete any existing code and paste this entire code.
 * 3. Click Save (Disk icon 💾).
 * 4. Deploy > New deployment > Select type: Web App > Execute as: Me > Who has access: Anyone > Deploy
 * 5. Copy the Web App URL and paste it into the AdVerse Dashboard's Webhook URL box.
 * 6. Reload the spreadsheet to see the menu: "🚀 AdVerse Dashboard"
 */

const ADVERSE_APP_URL = "https://ais-dev-nwgnfwsedvngbm6nzushq5-212874512310.asia-southeast1.run.app/api/rpc";
const ADMIN_PIN = "01726912624";

function syncToAdverseApp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const allData = {};
  
  sheets.forEach(function(s) {
    const name = s.getName();
    const values = s.getDataRange().getValues();
    if (values && values.length > 0) {
      allData[name] = values;
    }
  });

  const payload = {
    method: "syncFromAppsScriptPayload",
    args: [ss.getId(), allData],
    pin: ADMIN_PIN
  };
  
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const res = UrlFetchApp.fetch(ADVERSE_APP_URL, options);
    var json = null;
    try {
      json = JSON.parse(res.getContentText());
    } catch (parseErr) {
      SpreadsheetApp.getUi().alert("⚠️ Server responded with non-JSON output. Please verify app status.");
      return;
    }

    if (json && json.ok) {
      const msg = (json.result && json.result.message) ? json.result.message : "Synchronization completed successfully!";
      SpreadsheetApp.getUi().alert("✅ " + msg);
    } else {
      const err = (json && json.error) ? json.error : "Unknown server response";
      SpreadsheetApp.getUi().alert("❌ Sync Error: " + err);
    }
  } catch (err) {
    SpreadsheetApp.getUi().alert("❌ Connection Error: " + err.message);
  }
}

function pullAllAppDataToSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const payload = {
    method: "getAllSheetsData",
    args: [],
    pin: ADMIN_PIN
  };
  
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const res = UrlFetchApp.fetch(ADVERSE_APP_URL, options);
    const json = JSON.parse(res.getContentText());
    if (json && json.ok && json.result && json.result.allSheetsData) {
      const allSheets = json.result.allSheetsData;
      let count = 0;
      for (const sheetName in allSheets) {
        let sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
          sheet = ss.insertSheet(sheetName);
        }
        const rows = allSheets[sheetName];
        if (rows && rows.length > 0) {
          sheet.clear();
          sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
          count++;
        }
      }
      SpreadsheetApp.getUi().alert("✅ Success! Updated " + count + " sheets directly from AdVerse App.");
    } else {
      SpreadsheetApp.getUi().alert("❌ Pull Failed: " + (json.error || "No data received"));
    }
  } catch (err) {
    SpreadsheetApp.getUi().alert("❌ Network Error: " + err.message);
  }
}

/**
 * Webhook Post Entry Point: Receives live sync requests from web app (appendRow, updateRow, deleteRow)
 */
// ---------------------------------------------------------------------------
// Anti-double-write guard (added 2026-10-24): one manual Save used to reach
// the sheet through parallel paths (browser-direct + server), each scanning
// an empty sheet and appending the same SL twice (e.g. Fairy Style SL 10507:
// a temp_ row + a UUID row, identical except Daily $). Two protections:
//  1. A script lock serializes concurrent webhook writes, so scan-then-write
//     is atomic across executions.
//  2. A strong duplicate key (SL + Page + Start + Budget + AD Account) catches
//     the same entry arriving under a DIFFERENT Record ID. All five parts must
//     match, so genuinely distinct same-SL campaigns (e.g. two real Amar Sunnah
//     rows sharing SL 10491 with different budgets/accounts, or two different
//     pages sharing SL 10496) are NEVER merged.
// ---------------------------------------------------------------------------
function normalizeDateKey(v) {
  if (v === undefined || v === null || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    if (isNaN(v.getTime())) return '';
    return v.getFullYear() + '-' + ('0' + (v.getMonth() + 1)).slice(-2) + '-' + ('0' + v.getDate()).slice(-2);
  }
  var s = String(v).trim();
  if (!s) return '';
  var m = s.match(/^(\d{1,2})[-/ .]([A-Za-z]{3,9}|\d{1,2})[-/ .](\d{2,4})$/);
  if (m) {
    var months = { jan: '01', feb: '02', mar: '03', apr: '04', may: '05', jun: '06', jul: '07', aug: '08', sep: '09', oct: '10', nov: '11', dec: '12' };
    var mon = m[2];
    var mm = /^[0-9]+$/.test(mon) ? ('0' + parseInt(mon, 10)).slice(-2) : (months[mon.slice(0, 3).toLowerCase()] || '');
    if (!mm) return '';
    var yy = m[3].length === 2 ? '20' + m[3] : m[3];
    return yy + '-' + mm + '-' + ('0' + parseInt(m[1], 10)).slice(-2);
  }
  var iso = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (iso) return iso[1] + '-' + ('0' + parseInt(iso[2], 10)).slice(-2) + '-' + ('0' + parseInt(iso[3], 10)).slice(-2);
  return s;
}

function campaignRowStrongKey(headersLower, rowVals) {
  function col(name) {
    var i = headersLower.indexOf(name);
    return i === -1 ? '' : String(rowVals[i] || '').trim();
  }
  var page = '';
  for (var i = 0; i < headersLower.length; i++) {
    if (headersLower[i].indexOf('page') !== -1) { page = String(rowVals[i] || '').trim().toLowerCase(); break; }
  }
  var slRaw = col('sl') || col('sl no') || col('sl_no') || String(rowVals[0] || '');
  var sl = slRaw.replace(/\D/g, '');
  var start = normalizeDateKey(col('start') || col('start date'));
  var budgetRaw = col('budget');
  var budgetNum = budgetRaw === '' ? '' : String(parseFloat(String(budgetRaw).replace(/,/g, '')));
  if (budgetNum === 'NaN') budgetNum = String(budgetRaw).trim();
  var acct = col('ad account');
  if (!sl || !page || !start || !budgetNum || !acct) return '';
  return sl + '|' + page + '|' + start + '|' + budgetNum + '|' + acct.trim();
}

function doPost(e) {
  var lock = null;
  var lockHeld = false;
  try {
    var raw = e && e.postData && e.postData.contents ? e.postData.contents : "{}";
    var body = JSON.parse(raw);
    // Serialize concurrent webhook writes (see anti-double-write guard above).
    var WRITE_ACTIONS = ['appendRow', 'upsertRow', 'updateRow', 'deleteRow', 'batch', 'pushAllData'];
    if (body && WRITE_ACTIONS.indexOf(body.action) !== -1) {
      lock = LockService.getScriptLock();
      if (lock.tryLock(15000)) {
        lockHeld = true;
      } else {
        return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'lock_timeout_retry' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: "No active spreadsheet found." }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    function getFlexibleSheet(name) {
      if (!name) return null;
      var direct = ss.getSheetByName(name);
      if (direct) return direct;
      var clean = String(name).toLowerCase().replace(/[-_ ]/g, "");
      var all = ss.getSheets();
      for (var i = 0; i < all.length; i++) {
        var sTitle = all[i].getName();
        var sClean = sTitle.toLowerCase().replace(/[-_ ]/g, "");
        if (sClean === clean) return all[i];
        if (clean.indexOf("today") !== -1 && sClean.indexOf("today") !== -1) return all[i];
      }
      return null;
    }

    // 0. Batch Processing (Ultra fast bulk sync)
    if (body.action === "batch" && Array.isArray(body.batch)) {
      var batch = body.batch;
      for (var b = 0; b < batch.length; b++) {
        var item = batch[b];
        var s = getFlexibleSheet(item.sheetName) || ss.getSheetByName(item.sheetName);
        if (!s) {
          s = ss.insertSheet(item.sheetName);
          if (item.headers && item.headers.length > 0) s.appendRow(item.headers);
        }
        if (item.rowValues && item.rowValues.length > 0) {
          s.appendRow(item.rowValues);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true, count: batch.length }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 1. Append / Upsert Row (with automatic deduplication)
    if (body.action === "appendRow" || body.action === "upsertRow") {
      var sheet = getFlexibleSheet(body.sheetName) || ss.getSheetByName(body.sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(body.sheetName);
        if (body.headers && body.headers.length > 0) {
          sheet.appendRow(body.headers);
        }
      }
      if (body.rowValues && body.rowValues.length > 0) {
        var lastRow = sheet.getLastRow();
        var existingIdx = -1;
        if (lastRow > 1 && (body.recordId || body.matchInfo)) {
          var data = sheet.getDataRange().getValues();
          var headers = data[0].map(function(h) { return String(h || "").trim().toLowerCase(); });
          var recCol = headers.indexOf("record id");
          if (recCol === -1) recCol = headers.indexOf("_recordid");
          if (recCol === -1) recCol = headers.indexOf("recordid");
          var slCol = headers.indexOf("sl");
          if (slCol === -1) slCol = headers.indexOf("sl no");
          if (slCol === -1) slCol = headers.indexOf("sl_no");
          var pageCol = -1;
          for (var c = 0; c < headers.length; c++) {
            if (headers[c].indexOf("page") !== -1) { pageCol = c; break; }
          }
          var targetId = String(body.recordId || "").trim();
          var m = body.matchInfo || {};

          for (var i = 1; i < data.length; i++) {
            var rowVals = data[i];
            var rRecId = recCol >= 0 ? String(rowVals[recCol] || "").trim() : "";
            var rSl = slCol >= 0 ? String(rowVals[slCol] || "").trim().replace(/^#/, "") : String(rowVals[0] || "").trim().replace(/^#/, "");
            var rPage = pageCol >= 0 ? String(rowVals[pageCol] || "").trim().toLowerCase() : "";
            var rNum = parseInt(rSl.replace(/\D/g, ""), 10);
            var mNum = m.SL || m["SL NO"] ? parseInt(String(m.SL || m["SL NO"]).replace(/\D/g, ""), 10) : NaN;
            var targetNum = targetId ? parseInt(targetId.replace(/\D/g, ""), 10) : NaN;

            if (targetId && rRecId && targetId === rRecId) { existingIdx = i + 1; break; }
            if (targetId && rSl && targetId === rSl) { existingIdx = i + 1; break; }
            if (!isNaN(targetNum) && !isNaN(rNum) && targetNum === rNum && targetNum > 0) {
              if (m["Page Name"] && rPage) {
                var mp = String(m["Page Name"]).trim().toLowerCase();
                if (mp === rPage || rPage.indexOf(mp) !== -1 || mp.indexOf(rPage) !== -1) { existingIdx = i + 1; break; }
              } else { existingIdx = i + 1; break; }
            }
            if (m.SL && ((!isNaN(mNum) && !isNaN(rNum) && mNum === rNum) || String(m.SL).trim() === rSl)) {
              if (m["Page Name"] && rPage) {
                var mp = String(m["Page Name"]).trim().toLowerCase();
                if (mp === rPage || rPage.indexOf(mp) !== -1 || mp.indexOf(rPage) !== -1) { existingIdx = i + 1; break; }
              } else { existingIdx = i + 1; break; }
            }
            if (m["Record ID"] && rRecId && String(m["Record ID"]).trim() === rRecId) { existingIdx = i + 1; break; }
          }
        }

        // Strong duplicate key: same entry under a DIFFERENT Record ID
        // (the classic double-write signature) must update, never append.
        if (existingIdx <= 1 && body.rowValues && body.headers) {
          var inHeaders = body.headers.map(function(h) { return String(h || '').trim().toLowerCase(); });
          var inKey = campaignRowStrongKey(inHeaders, body.rowValues);
          if (inKey) {
            var fresh = sheet.getDataRange().getValues();
            var freshHeaders = fresh[0].map(function(h) { return String(h || '').trim().toLowerCase(); });
            for (var k = 1; k < fresh.length; k++) {
              if (campaignRowStrongKey(freshHeaders, fresh[k]) === inKey) { existingIdx = k + 1; break; }
            }
          }
        }

        if (existingIdx > 1) {
          sheet.getRange(existingIdx, 1, 1, body.rowValues.length).setValues([body.rowValues]);
        } else {
          sheet.appendRow(body.rowValues);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true, action: "appendRow" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 2. Update Row
    if (body.action === "updateRow") {
      var sheet = getFlexibleSheet(body.sheetName);
      if (sheet && body.rowValues && body.rowValues.length > 0) {
        var lastRow = sheet.getLastRow();
        var updated = false;
        if (lastRow > 1 && (body.recordId || body.matchInfo)) {
          var data = sheet.getDataRange().getValues();
          var headers = data[0].map(function(h) { return String(h || "").trim().toLowerCase(); });
          var recCol = headers.indexOf("record id");
          if (recCol === -1) recCol = headers.indexOf("_recordid");
          if (recCol === -1) recCol = headers.indexOf("recordid");
          var slCol = headers.indexOf("sl");
          if (slCol === -1) slCol = headers.indexOf("sl no");
          var pageCol = -1;
          for (var c = 0; c < headers.length; c++) {
            if (headers[c].indexOf("page") !== -1) { pageCol = c; break; }
          }
          var targetId = String(body.recordId || "").trim();
          var m = body.matchInfo || {};
          // Strong duplicate key (same entry, different Record ID).
          var updKey = '';
          if (body.rowValues && body.headers) {
            updKey = campaignRowStrongKey(body.headers.map(function(h) { return String(h || '').trim().toLowerCase(); }), body.rowValues);
          }

          for (var i = 1; i < data.length; i++) {
            var rowVals = data[i];
            var match = false;
            var rRecId = recCol >= 0 ? String(rowVals[recCol] || "").trim() : "";
            var rSl = slCol >= 0 ? String(rowVals[slCol] || "").trim() : String(rowVals[0] || "").trim();
            var rPage = pageCol >= 0 ? String(rowVals[pageCol] || "").trim().toLowerCase() : "";
            var rNum = parseInt(rSl.replace(/\D/g, ""), 10);
            var mNum = m.SL ? parseInt(String(m.SL).replace(/\D/g, ""), 10) : NaN;

            if (targetId && rRecId && targetId === rRecId) match = true;
            else if (targetId && rSl && targetId === rSl) match = true;
            else if (m.SL && ((!isNaN(mNum) && !isNaN(rNum) && mNum === rNum) || String(m.SL).trim() === rSl)) {
              if (m["Page Name"] && rPage) {
                var mp = String(m["Page Name"]).trim().toLowerCase();
                if (mp === rPage || rPage.indexOf(mp) !== -1 || mp.indexOf(rPage) !== -1) match = true;
              } else { match = true; }
            }
            if (!match && updKey && campaignRowStrongKey(headers, rowVals) === updKey) match = true;

            if (match) {
              sheet.getRange(i + 1, 1, 1, body.rowValues.length).setValues([body.rowValues]);
              updated = true;
              break;
            }
          }
        }
        if (!updated && (body.createIfMissing !== false)) {
          sheet.appendRow(body.rowValues);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true, action: "updateRow" })).setMimeType(ContentService.MimeType.JSON);
    }

    // 3. Delete Row
    if (body.action === "deleteRow") {
      var sheet = getFlexibleSheet(body.sheetName);
      if (sheet && (body.recordId || body.matchInfo)) {
        var lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          var data = sheet.getDataRange().getValues();
          var headers = data[0].map(function(h) { return String(h || "").trim().toLowerCase(); });
          var recCol = headers.indexOf("record id");
          if (recCol === -1) recCol = headers.indexOf("_recordid");
          var slCol = headers.indexOf("sl");
          if (slCol === -1) slCol = headers.indexOf("sl no");
          var pageCol = -1;
          for (var c = 0; c < headers.length; c++) {
            if (headers[c].indexOf("page") !== -1) { pageCol = c; break; }
          }
          var targetId = String(body.recordId || "").trim();
          var m = body.matchInfo || {};

          for (var i = 1; i < data.length; i++) {
            var rowVals = data[i];
            if (!rowVals || !Array.isArray(rowVals) || rowVals.length === 0) continue;
            var match = false;
            var rRecId = (recCol >= 0 && recCol < rowVals.length && rowVals[recCol] != null) ? String(rowVals[recCol]).trim() : "";
            var rSl = (slCol >= 0 && slCol < rowVals.length && rowVals[slCol] != null) 
              ? String(rowVals[slCol]).trim().replace(/^#/, "") 
              : (rowVals[0] != null ? String(rowVals[0]).trim().replace(/^#/, "") : "");
            var rPage = (pageCol >= 0 && pageCol < rowVals.length && rowVals[pageCol] != null) ? String(rowVals[pageCol]).trim().toLowerCase() : "";
            var rNum = parseInt(rSl.replace(/\D/g, ""), 10);
            var mNum = m.SL || m["SL NO"] ? parseInt(String(m.SL || m["SL NO"]).replace(/\D/g, ""), 10) : NaN;
            var targetNum = targetId ? parseInt(targetId.replace(/\D/g, ""), 10) : NaN;

            if (targetId && rRecId && targetId === rRecId) match = true;
            else if (targetId && rSl && targetId === rSl) match = true;
            else if (!isNaN(targetNum) && !isNaN(rNum) && targetNum === rNum && targetNum > 0) {
              if (m["Page Name"] && rPage) {
                var mp = String(m["Page Name"]).trim().toLowerCase();
                if (mp === rPage || rPage.indexOf(mp) !== -1 || mp.indexOf(rPage) !== -1) match = true;
              } else { match = true; }
            }
            else if (m.SL && ((!isNaN(mNum) && !isNaN(rNum) && mNum === rNum) || String(m.SL).trim() === rSl)) {
              if (m["Page Name"] && rPage) {
                var mp = String(m["Page Name"]).trim().toLowerCase();
                if (mp === rPage || rPage.indexOf(mp) !== -1 || mp.indexOf(rPage) !== -1) match = true;
              } else { match = true; }
            } else if (m["Record ID"] && rRecId && String(m["Record ID"]).trim() === rRecId) {
              match = true;
            }

            if (match) {
              sheet.deleteRow(i + 1);
              break;
            }
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true, action: "deleteRow" })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // 4. Push All Data (Full Sync) - with rectangular 2D array padding to prevent setValues errors
    if (body.action === "pushAllData" && body.allSheetsData) {
      var allData = body.allSheetsData;
      for (var sName in allData) {
        var s = ss.getSheetByName(sName) || ss.insertSheet(sName);
        var rows = allData[sName];
        if (rows && rows.length > 0) {
          s.clear();
          var maxCols = 0;
          for (var r = 0; r < rows.length; r++) {
            if (rows[r] && rows[r].length > maxCols) maxCols = rows[r].length;
          }
          if (maxCols > 0) {
            var paddedRows = rows.map(function(rRow) {
              var arr = Array.isArray(rRow) ? rRow.slice() : [];
              while (arr.length < maxCols) arr.push("");
              return arr;
            });
            s.getRange(1, 1, paddedRows.length, maxCols).setValues(paddedRows);
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true, action: "pushAllData" })).setMimeType(ContentService.MimeType.JSON);
    }

    // 5. Ping
    if (body.action === "ping") {
      return ContentService.createTextOutput(JSON.stringify({ success: true, message: "pong" })).setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    try { if (lockHeld && lock) lock.releaseLock(); } catch (_) {}
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: "ok", app: "AdVerse Sync Connector Live" })).setMimeType(ContentService.MimeType.JSON);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🚀 AdVerse Dashboard")
    .addItem("📅 1-Click: Format All Dates to DD-Mon-YYYY (12-Oct-2026)", "formatAllDatesToDDMonYYYY")
    .addItem("📅 1-Click: Format All Dates to DD/MM/YYYY (12/10/2026)", "formatAllDatesToDDMMYYYY")
    .addItem("📅 1-Click: Format All Dates to YYYY-MM-DD (2026-10-12)", "formatAllDatesToYYYYMMDD")
    .addItem("📅 1-Click: Format All Dates to DD-MM-YYYY (12-10-2026)", "formatAllDatesToDDMMYYYY_Dash")
    .addSeparator()
    .addItem("📥 Pull All App Data to Sheet", "pullAllAppDataToSheet")
    .addItem("🔄 Sync this Sheet to AdVerse App", "syncToAdverseApp")
    .addToUi();
}

function formatAllDatesToDDMonYYYY() { return formatAllDatesInEntireSpreadsheet('DD-Mon-YYYY'); }
function formatAllDatesToDDMMYYYY() { return formatAllDatesInEntireSpreadsheet('DD/MM/YYYY'); }
function formatAllDatesToYYYYMMDD() { return formatAllDatesInEntireSpreadsheet('YYYY-MM-DD'); }
function formatAllDatesToDDMMYYYY_Dash() { return formatAllDatesInEntireSpreadsheet('DD-MM-YYYY'); }

/**
 * 📅 MULTI-FORMAT UTILITY: Format All Dates in Entire Spreadsheet to user's desired format
 * Supported formats: 'DD-Mon-YYYY' (default), 'DD/MM/YYYY', 'YYYY-MM-DD', 'DD-MM-YYYY', 'D-Mon-YYYY'
 */
function formatAllDatesInEntireSpreadsheet(chosenFormat) {
  var targetFormat = chosenFormat || 'DD-Mon-YYYY';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var shortMonths = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var fullMonths = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  
  var totalSheetsUpdated = 0;
  var totalCellsFormatted = 0;
  var summaryLines = [];

  function formatDateWithPattern(dayNum, mIndex, yearStr) {
    var dPad = ('0' + parseInt(dayNum, 10)).slice(-2);
    var dSingle = String(parseInt(dayNum, 10));
    var mShort = shortMonths[mIndex] || 'Jan';
    var mNum = ('0' + (mIndex + 1)).slice(-2);
    var yStr = String(yearStr);
    switch (targetFormat) {
      case 'DD/MM/YYYY': return dPad + '/' + mNum + '/' + yStr;
      case 'YYYY-MM-DD': return yStr + '-' + mNum + '-' + dPad;
      case 'DD-MM-YYYY': return dPad + '-' + mNum + '-' + yStr;
      case 'D-Mon-YYYY': return dSingle + '-' + mShort + '-' + yStr;
      case 'DD-Mon-YYYY':
      default: return dPad + '-' + mShort + '-' + yStr;
    }
  }

  function parseToStandardDate(val, monthHint) {
    if (val === undefined || val === null || val === '') return '';
    
    // 1. If already a Date object
    if (Object.prototype.toString.call(val) === '[object Date]') {
      if (isNaN(val.getTime())) return '';
      return formatDateWithPattern(val.getDate(), val.getMonth(), val.getFullYear());
    }

    // 2. If number (Google Sheets Date serial e.g. 45560 = ~2024, 46259 = ~2026)
    if (typeof val === 'number' && val > 30000 && val < 70000) {
      var utcDays = Math.round(val - 25569);
      var d = new Date((utcDays * 86400 + 43200) * 1000);
      if (!isNaN(d.getTime())) {
        return formatDateWithPattern(d.getDate(), d.getMonth(), d.getFullYear());
      }
    }

    var str = String(val).trim();
    if (!str) return '';
    var lower = str.toLowerCase();
    if (lower === 'undefined' || lower === 'null' || lower === 'nan' || lower === 'invalid date' || lower === 'loading...') return '';

    // Convert Bengali digits ০-৯ to 0-9
    var bnToEn = { '০':'0', '১':'1', '২':'2', '৩':'3', '৪':'4', '৫':'5', '৬':'6', '৭':'7', '৮':'8', '৯':'9' };
    str = str.replace(/[০-৯]/g, function(d) { return bnToEn[d] || d; });

    // Convert Bengali month names to English
    var bnMonthMap = {
      'জানুয়ারি': 'Jan', 'জানুয়ারী': 'Jan', 'ফেব্রুয়ারি': 'Feb', 'ফেব্রুয়ারী': 'Feb',
      'মার্চ': 'Mar', 'এপ্রিল': 'Apr', 'মে': 'May', 'জুন': 'Jun',
      'জুলাই': 'Jul', 'আগস্ট': 'Aug', 'সেপ্টেম্বর': 'Sep',
      'অক্টোবর': 'Oct', 'নভেম্বর': 'Nov', 'ডিসেম্বর': 'Dec'
    };
    for (var bnM in bnMonthMap) {
      if (str.indexOf(bnM) !== -1) {
        str = str.replace(bnM, bnMonthMap[bnM]);
        break;
      }
    }

    // If already in DD-Mon-YYYY (e.g. 11-Oct-2026, 9-Oct-26, 25-Aug-2026)
    var exactMatch = str.match(/^(\d{1,2})[-/ .]([A-Za-z]{3,9})[-/ .](\d{2,4})$/);
    if (exactMatch) {
      var day = ('0' + parseInt(exactMatch[1], 10)).slice(-2);
      var monPart = exactMatch[2].slice(0, 3).toLowerCase();
      var mIdx = -1;
      for (var m = 0; m < shortMonths.length; m++) {
        if (shortMonths[m].toLowerCase() === monPart) { mIdx = m; break; }
      }
      var mon = mIdx >= 0 ? shortMonths[mIdx] : (exactMatch[2].charAt(0).toUpperCase() + exactMatch[2].slice(1, 3).toLowerCase());
      var yr = exactMatch[3];
      if (yr.length === 2) yr = '20' + yr;
      return day + '-' + mon + '-' + yr;
    }

    // Clean off timestamp if present (e.g. 2026-10-11 12:00:00 or 11/10/2026 00:00:00)
    var cleanStr = str.replace(/T\d{2}:\d{2}.*$/, '').replace(/\s+\d{1,2}:\d{2}(:\d{2})?.*$/, '').trim();

    // YYYY-MM-DD
    var isoMatch = cleanStr.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
    if (isoMatch) {
      var yr = isoMatch[1];
      var mNum = parseInt(isoMatch[2], 10) - 1;
      var day = ('0' + parseInt(isoMatch[3], 10)).slice(-2);
      if (mNum >= 0 && mNum < 12) {
        return day + '-' + shortMonths[mNum] + '-' + yr;
      }
    }

    // Mon-DD-YYYY (e.g. Oct 11, 2026 or Oct-11-2026)
    var monDYMatch = cleanStr.match(/^([A-Za-z]{3,9})[-/ .](\d{1,2}),?[-/ .](\d{2,4})$/);
    if (monDYMatch) {
      var monPart = monDYMatch[1].slice(0, 3).toLowerCase();
      var mIdx = -1;
      for (var m = 0; m < shortMonths.length; m++) {
        if (shortMonths[m].toLowerCase() === monPart) { mIdx = m; break; }
      }
      if (mIdx !== -1) {
        var day = ('0' + parseInt(monDYMatch[2], 10)).slice(-2);
        var yr = monDYMatch[3];
        if (yr.length === 2) yr = '20' + yr;
        return day + '-' + shortMonths[mIdx] + '-' + yr;
      }
    }

    // DD/MM/YYYY or MM/DD/YYYY
    var dmyMatch = cleanStr.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
    if (dmyMatch) {
      var p1 = parseInt(dmyMatch[1], 10);
      var p2 = parseInt(dmyMatch[2], 10);
      var yr = dmyMatch[3];
      if (yr.length === 2) yr = '20' + yr;
      var day = p1;
      var mIdx = p2 - 1;

      // Disambiguate using monthHint (e.g. from sheet name "Campaign_October_2026")
      if (monthHint) {
        var hintLower = String(monthHint).toLowerCase();
        var hintIdx = -1;
        for (var m = 0; m < fullMonths.length; m++) {
          if (fullMonths[m].toLowerCase() === hintLower || shortMonths[m].toLowerCase() === hintLower.slice(0, 3)) {
            hintIdx = m;
            break;
          }
        }
        if (hintIdx !== -1) {
          if (p1 === hintIdx + 1 && p2 <= 31) {
            day = p2;
            mIdx = hintIdx;
          } else if (p2 === hintIdx + 1 && p1 <= 31) {
            day = p1;
            mIdx = hintIdx;
          }
        }
      } else {
        if (p1 > 12 && p2 <= 12) {
          day = p1;
          mIdx = p2 - 1;
        } else if (p1 <= 12 && p2 > 12) {
          day = p2;
          mIdx = p1 - 1;
        } else if (cleanStr.indexOf('/') !== -1) {
          day = p2;
          mIdx = p1 - 1;
        }
      }

      if (mIdx >= 0 && mIdx < 12 && day >= 1 && day <= 31) {
        return formatDateWithPattern(day, mIdx, yr);
      }
    }

    // Fallback native Date parse
    var generic = new Date(cleanStr);
    if (!isNaN(generic.getTime())) {
      return formatDateWithPattern(generic.getDate(), generic.getMonth(), generic.getFullYear());
    }

    return str;
  }

  function isDateColumnHeader(header) {
    if (!header) return false;
    var lower = String(header).trim().toLowerCase();
    return (
      lower === 'start' ||
      lower === 'end' ||
      lower === 'start date' ||
      lower === 'end date' ||
      lower === 'entry date' ||
      lower === 'entry' ||
      lower === 'date' ||
      lower === 'billing date' ||
      lower === 'closed at' ||
      lower === 'payment date' ||
      lower.slice(-5) === ' date' ||
      lower.slice(0, 5) === 'date ' ||
      lower === 'date/time'
    );
  }

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    if (lastRow < 2 || lastCol < 1) continue;

    var monthHint = '';
    for (var m = 0; m < fullMonths.length; m++) {
      if (sheetName.toLowerCase().indexOf(fullMonths[m].toLowerCase()) !== -1) {
        monthHint = fullMonths[m];
        break;
      }
    }

    var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var dateColIndices = [];

    for (var c = 0; c < headerRow.length; c++) {
      if (isDateColumnHeader(headerRow[c])) {
        dateColIndices.push(c + 1);
      }
    }

    if (dateColIndices.length === 0) continue;

    var sheetCellsChanged = 0;

    for (var dIdx = 0; dIdx < dateColIndices.length; dIdx++) {
      var colNum = dateColIndices[dIdx];
      var range = sheet.getRange(2, colNum, lastRow - 1, 1);
      var colValues = range.getValues();
      var changed = false;

      for (var r = 0; r < colValues.length; r++) {
        var rawVal = colValues[r][0];
        if (rawVal !== undefined && rawVal !== null && rawVal !== '') {
          var formatted = parseToStandardDate(rawVal, monthHint);
          if (formatted && formatted !== String(rawVal)) {
            colValues[r][0] = formatted;
            sheetCellsChanged++;
            changed = true;
          }
        }
      }

      if (changed) {
        range.setNumberFormat('@');
        range.setValues(colValues);
      }
    }

    if (sheetCellsChanged > 0) {
      totalSheetsUpdated++;
      totalCellsFormatted += sheetCellsChanged;
      summaryLines.push(sheetName + ': ' + sheetCellsChanged + ' dates');
    }
  }

  var msg = '🎉 Complete! Formatted ' + totalCellsFormatted + ' date cells across ' + totalSheetsUpdated + ' sheets to standard DD-Mon-YYYY (e.g. 11-Oct-2026)!';
  if (summaryLines.length > 0) {
    msg += '\n\nDetails:\n' + summaryLines.join('\n');
  }
  
  Logger.log(msg);
  try {
    SpreadsheetApp.getUi().alert('📅 AdVerse Date Normalizer', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (uiErr) {}
  return { success: true, totalCells: totalCellsFormatted, totalSheets: totalSheetsUpdated, details: summaryLines };
}
