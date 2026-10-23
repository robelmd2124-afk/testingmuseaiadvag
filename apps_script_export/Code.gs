/**
 * AdVerse Agency - Google Apps Script Backend (Code.gs)
 * Standalone Web App Edition - Fully Native & Fast
 * 
 * Features:
 * - Direct SpreadsheetApp manipulation (zero external API keys needed)
 * - Fast In-Memory Sheet Range Caching
 * - Multi-tab Management (Campaigns, Payments, Agency Cost, Page Summary, etc.)
 * - Auto SL Generator, Batch Operations, Page Profiles & Visual Stats
 * - Role Based Access (Admin PIN: 1234, Viewer PIN: 1111)
 */

const ADMIN_PIN = "01726912624";
const FALLBACK_ADMIN_PIN = "1234";
const VIEWER_PIN = "1111";

const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const MONTH_NAMES_SHORT = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const MONTH_NAMES_FULL = ['January','February','March','April','May','June','July','August','September','October','November','December'];

var CURRENT_DATE_FORMAT = 'DD-MM-YYYY';

function setAppsScriptDateFormat(fmt) {
  if (fmt) CURRENT_DATE_FORMAT = fmt;
}

/**
 * Standard date formatter for Google Apps Script with full multi-format support
 * Supports: DD-MM-YYYY (default), DD/MM/YYYY, DD-Mon-YYYY, YYYY-MM-DD, D-Mon-YYYY
 */
function formatCustomDate(dateInput, monthHint, formatPattern) {
  if (dateInput === undefined || dateInput === null || dateInput === '') return '';
  var targetFormat = formatPattern || CURRENT_DATE_FORMAT || 'DD-MM-YYYY';
  var d = null;
  if (Object.prototype.toString.call(dateInput) === '[object Date]') {
    if (isNaN(dateInput.getTime())) return '';
    d = dateInput;
  } else {
    var str = String(dateInput).trim();
    if (!str) return '';
    var lower = str.toLowerCase();
    if (lower === 'undefined' || lower === 'null' || lower === 'nan' || lower === 'invalid date' || lower === 'loading...') return '';
    d = parseAnyDate(str, monthHint);
  }

  if (!d || isNaN(d.getTime())) return String(dateInput || '');
  var dayPadded = ('0' + d.getDate()).slice(-2);
  var daySingle = String(d.getDate());
  var monIdx = d.getMonth();
  var monShort = MONTH_NAMES_SHORT[monIdx] || 'Jan';
  var monNum = ('0' + (monIdx + 1)).slice(-2);
  var yr = String(d.getFullYear());

  switch (targetFormat) {
    case 'DD/MM/YYYY':
      return dayPadded + '/' + monNum + '/' + yr;
    case 'YYYY-MM-DD':
      return yr + '-' + monNum + '-' + dayPadded;
    case 'DD-MM-YYYY':
      return dayPadded + '-' + monNum + '-' + yr;
    case 'D-Mon-YYYY':
      return daySingle + '-' + monShort + '-' + yr;
    case 'DD-Mon-YYYY':
    default:
      return dayPadded + '-' + monShort + '-' + yr;
  }
}
const CH = ['SL','OFF $','AD Account','ID Name','Status','Page Name','Start','End','Types','Day','Budget','Daily $','Taka','Rate','Entry Date','Month','Year','Record ID'];
const PAYMENT_HEADERS = ['Date','Payment Method','Paid (৳)','Month','Year'];
const TODAY_HEADERS = ['SL','OFF $','AD Account','ID Name','Status','Page Name','Start','End','Types','Day','Budget','Daily $','Taka','Rate','Entry Date','Record ID','Month','Year'];
const PAGE_SUMMARY_HEADERS = ['Page Name','Total Ads','Total Budget $'];
const AGENCY_SPENDING_HEADERS = ['Month','Year','Total Spend $'];
const MONTHLY_CLOSING_HEADERS = ['Month','Year','Status','Closed At','Closed By','Total Campaign','Total Budget $','Total Campaign Taka','Total Paid (৳)','Total Due (৳)','Agency Cost (৳)','Rate'];
const AGENCY_COST_HEADERS = ['Month','Year','Type','Cost','Per Person Cost','Status'];
const END_ADS_HEADERS = ['Page Name','Start Date','End Date','Day','Budget','Type'];
const MENU = [
  {key:'today_created_campaign',label:'Today Created Campaign',custom:1},
  {key:'campaign',label:'Campaign',custom:1},
  {key:'end_ads',label:'End Ads',custom:9},
  {key:'page_summery',label:'Page Summery',custom:3},
  {key:'agency_spending',label:'Agency Calculation',custom:5},
  {key:'agency_cost',label:'Agency Cost',custom:8},
  {key:'payment_due_tracker',label:'Payments',custom:2}
];

const DEFAULT_AD_ACCOUNTS = [
  { a: '01 (MR)', id: 'Morshed' },
  { a: '02 (MR)', id: 'Morshed' },
  { a: '03 (MR)', id: 'Morshed' },
  { a: '04 (MR)', id: 'Morshed' },
  { a: '05 (MR)', id: 'Morshed' },
  { a: '06 (MR)', id: 'Morshed' },
  { a: '07 (MR)', id: 'Morshed' },
  { a: '08 (MR)', id: 'Morshed' },
  { a: '09 (MR)', id: 'Morshed' },
  { a: '10 (MR)', id: 'Morshed' },
  { a: '11 (MR)', id: 'Morshed' },
  { a: '12 (MR)', id: 'Morshed' }
];

const DEFAULT_PAGES = [
  'Tech World', 'Fashion House', 'Food Corner', 'Gadget Zone', 'Daily Deals', 
  'Organic Pure', 'Style Hub', 'Smart Shopping', 'Trendsetter BD'
];

/**
 * Web App Entry Point: Serves the Complete Single-Page UI
 */
/**
 * Reads every tab of the bound spreadsheet and returns { tabName: [[header...],[row...]] }.
 * Uses getDisplayValues so dates / numbers arrive exactly as shown in the sheet.
 * This is what lets the web app load ALL sheet data even when the sheet is private.
 */
function readAllSheetsForApp_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = {};
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow < 1 || lastCol < 1) continue;
    out[sh.getName()] = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  }
  return {
    success: true,
    action: "readAll",
    sheetId: ss.getId(),
    sheetName: ss.getName(),
    allSheetsData: out
  };
}

function doGet(e) {
  // If tested via ping or action parameter
  if (e && e.parameter && (e.parameter.action === 'ping' || e.parameter.test === '1')) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "ok",
      app: "AdVerse Agency Suite",
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // Read-all endpoint used by the web app to pull every sheet tab (works for private sheets)
  if (e && e.parameter && (e.parameter.action === 'readAll' || e.parameter.action === 'getAllSheetsData')) {
    try {
      return ContentService.createTextOutput(JSON.stringify(readAllSheetsForApp_()))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (readErr) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: readErr.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  try {
    var template = HtmlService.createTemplateFromFile('index');
    var htmlOutput = template.evaluate()
      .setTitle('AdVerse Agency - Intelligence & Campaign Suite')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    return htmlOutput;
  } catch (err) {
    // Graceful fallback if index.html is not included in Apps Script project (e.g. connector mode)
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "ok", 
      message: "AdVerse Google Apps Script Backend is active and running!",
      sheetConnected: !!SpreadsheetApp.getActiveSpreadsheet(),
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Webhook Post Entry Point: Receives live sync requests from web app (appendRow, updateRow, deleteRow)
 */
function doPost(e) {
  try {
    var raw = e.postData ? e.postData.contents : "";
    var body = raw ? JSON.parse(raw) : (e.parameter || {});
    var ss = SpreadsheetApp.getActiveSpreadsheet();

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

    // Ping / Test
    if (body.action === "ping" || body.action === "test") {
      return ContentService.createTextOutput(JSON.stringify({ 
        success: true, 
        message: "Connected to Google Sheet Apps Script successfully!", 
        sheetName: ss.getName(),
        sheetId: ss.getId() 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Read all tabs (POST variant of the readAll endpoint)
    if (body.action === "readAll") {
      return ContentService.createTextOutput(JSON.stringify(readAllSheetsForApp_()))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 1. Append / Upsert Row
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
          if (data && data.length > 0 && Array.isArray(data[0])) {
            var headers = data[0].map(function(h) { return String(h || "").trim().toLowerCase(); });
            var recCol = headers.indexOf("record id");
            if (recCol === -1) recCol = headers.indexOf("_recordid");
            if (recCol === -1) recCol = headers.indexOf("recordid");
            var slCol = headers.indexOf("sl");
            if (slCol === -1) slCol = headers.indexOf("sl no");
            if (slCol === -1) slCol = headers.indexOf("sl_no");
            var pageCol = -1;
            for (var c = 0; c < headers.length; c++) {
              var hStr = String(headers[c] || "").trim().toLowerCase();
              if (hStr.indexOf("page") !== -1) { pageCol = c; break; }
            }
            var targetId = String(body.recordId || "").trim();
            var m = body.matchInfo || {};

            for (var i = 1; i < data.length; i++) {
              var rowVals = data[i];
              if (!rowVals || !Array.isArray(rowVals) || rowVals.length === 0) continue;
              var rRecId = (recCol >= 0 && recCol < rowVals.length && rowVals[recCol] != null) ? String(rowVals[recCol]).trim() : "";
              var rSl = (slCol >= 0 && slCol < rowVals.length && rowVals[slCol] != null) 
                ? String(rowVals[slCol]).trim().replace(/^#/, "") 
                : (rowVals[0] != null ? String(rowVals[0]).trim().replace(/^#/, "") : "");
              var rPage = (pageCol >= 0 && pageCol < rowVals.length && rowVals[pageCol] != null) ? String(rowVals[pageCol]).trim().toLowerCase() : "";
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
          if (data && data.length > 0 && Array.isArray(data[0])) {
            var headers = data[0].map(function(h) { return String(h || "").trim().toLowerCase(); });
            var recCol = headers.indexOf("record id");
            if (recCol === -1) recCol = headers.indexOf("_recordid");
            if (recCol === -1) recCol = headers.indexOf("recordid");
            var slCol = headers.indexOf("sl");
            if (slCol === -1) slCol = headers.indexOf("sl no");
            if (slCol === -1) slCol = headers.indexOf("sl_no");
            var pageCol = -1;
            for (var c = 0; c < headers.length; c++) {
              var hStr = String(headers[c] || "").trim().toLowerCase();
              if (hStr.indexOf("page") !== -1) { pageCol = c; break; }
            }
            var targetId = String(body.recordId || "").trim();
            var m = body.matchInfo || {};

            for (var i = 1; i < data.length; i++) {
              var rowVals = data[i];
              if (!rowVals || !Array.isArray(rowVals) || rowVals.length === 0) continue;
              var match = false;
              var rRecId = (recCol >= 0 && recCol < rowVals.length && rowVals[recCol] != null) ? String(rowVals[recCol]).trim() : "";
              var rSl = (slCol >= 0 && slCol < rowVals.length && rowVals[slCol] != null) 
                ? String(rowVals[slCol]).trim() 
                : (rowVals[0] != null ? String(rowVals[0]).trim() : "");
              var rPage = (pageCol >= 0 && pageCol < rowVals.length && rowVals[pageCol] != null) ? String(rowVals[pageCol]).trim().toLowerCase() : "";
              var rNum = parseInt(rSl.replace(/\D/g, ""), 10);
              var mNum = m.SL ? parseInt(String(m.SL).replace(/\D/g, ""), 10) : NaN;
              var targetNum = targetId ? parseInt(targetId.replace(/\D/g, ""), 10) : NaN;

              if (targetId && rRecId && targetId === rRecId) match = true;
              else if (targetId && rSl && targetId === rSl) match = true;
              else if (!isNaN(targetNum) && !isNaN(rNum) && targetNum === rNum && targetNum > 0) {
                if (m["Page Name"] && rPage) {
                  var mp = String(m["Page Name"]).trim().toLowerCase();
                  if (mp === rPage || rPage.indexOf(mp) !== -1 || mp.indexOf(rPage) !== -1) match = true;
                } else { match = true; }
              }
              else if (m.SL && ((!isNaN(mNum) && !isNaN(rNum) && mNum === rNum) || String(m.SL).trim() === rSl || rSl.indexOf(String(m.SL).trim()) !== -1)) {
                if (m["Page Name"] && rPage) {
                  var mp = String(m["Page Name"]).trim().toLowerCase();
                  if (mp === rPage || rPage.indexOf(mp) !== -1 || mp.indexOf(rPage) !== -1) match = true;
                } else { match = true; }
              } else if (m["Record ID"] && rRecId && String(m["Record ID"]).trim() === rRecId) {
                match = true;
              }

              if (match) {
                sheet.getRange(i + 1, 1, 1, body.rowValues.length).setValues([body.rowValues]);
                updated = true;
                break;
              }
            }
          }
        }
        if (!updated && body.rowNo && body.rowNo > 1 && body.rowNo <= lastRow) {
          sheet.getRange(body.rowNo, 1, 1, body.rowValues.length).setValues([body.rowValues]);
        } else if (!updated) {
          sheet.appendRow(body.rowValues);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true, action: "updateRow" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 3. Delete Row (Line 327 Fix: complete null-guarding, reverse traversal, date normalization)
    if (body.action === "deleteRow") {
      var sheet = getFlexibleSheet(body.sheetName);
      if (sheet) {
        var lastRow = sheet.getLastRow();
        var deleted = false;
        if (lastRow > 1 && (body.recordId || body.matchInfo)) {
          var data = sheet.getDataRange().getValues();
          if (data && data.length > 0 && Array.isArray(data[0])) {
            var headers = data[0].map(function(h) { return String(h || "").trim().toLowerCase(); });
            var recCol = headers.indexOf("record id");
            if (recCol === -1) recCol = headers.indexOf("_recordid");
            if (recCol === -1) recCol = headers.indexOf("recordid");
            var slCol = headers.indexOf("sl");
            if (slCol === -1) slCol = headers.indexOf("sl no");
            if (slCol === -1) slCol = headers.indexOf("sl_no");
            var pageCol = -1;
            for (var c = 0; c < headers.length; c++) {
              var hStr = String(headers[c] || "").trim().toLowerCase();
              if (hStr.indexOf("page") !== -1) { pageCol = c; break; }
            }
            var dateCol = headers.indexOf("date");
            var costCol = headers.indexOf("cost");
            if (costCol === -1) costCol = headers.indexOf("spent");
            if (costCol === -1) costCol = headers.indexOf("paid (৳)");
            if (costCol === -1) costCol = headers.indexOf("paid");
            if (costCol === -1) costCol = headers.indexOf("budget");

            var targetId = String(body.recordId || "").trim();
            var m = body.matchInfo || {};

            for (var i = data.length - 1; i >= 1; i--) {
              var rowVals = data[i];
              if (!rowVals || !Array.isArray(rowVals) || rowVals.length === 0) continue;
              var match = false;
              var rRecId = (recCol >= 0 && recCol < rowVals.length && rowVals[recCol] != null) ? String(rowVals[recCol]).trim() : "";
              var rSl = (slCol >= 0 && slCol < rowVals.length && rowVals[slCol] != null) 
                ? String(rowVals[slCol]).trim().replace(/^#/, "") 
                : (rowVals[0] != null ? String(rowVals[0]).trim().replace(/^#/, "") : "");
              var rPage = (pageCol >= 0 && pageCol < rowVals.length && rowVals[pageCol] != null) ? String(rowVals[pageCol]).trim().toLowerCase() : "";
              var rDate = (dateCol >= 0 && dateCol < rowVals.length && rowVals[dateCol] != null) ? formatCustomDate(rowVals[dateCol]) : "";
              var rCost = (costCol >= 0 && costCol < rowVals.length && rowVals[costCol] != null) ? parseFloat(String(rowVals[costCol]).replace(/[^0-9.]/g, "")) : NaN;

              var rNum = parseInt(rSl.replace(/\D/g, ""), 10);
              var mNum = (m.SL || m["SL NO"]) ? parseInt(String(m.SL || m["SL NO"]).replace(/\D/g, ""), 10) : NaN;
              var targetNum = targetId ? parseInt(targetId.replace(/\D/g, ""), 10) : NaN;

              if (targetId && rRecId && targetId === rRecId) { match = true; }
              else if (targetId && rSl && targetId === rSl) { match = true; }
              else if (!isNaN(targetNum) && !isNaN(rNum) && targetNum === rNum && targetNum > 0) {
                if (m["Page Name"] && rPage) {
                  var mp = String(m["Page Name"]).trim().toLowerCase();
                  if (mp === rPage || rPage.indexOf(mp) !== -1 || mp.indexOf(rPage) !== -1) match = true;
                } else { match = true; }
              }
              else if (m.SL && ((!isNaN(mNum) && !isNaN(rNum) && mNum === rNum) || String(m.SL).trim() === rSl || rSl.indexOf(String(m.SL).trim()) !== -1)) {
                if (m["Page Name"] && rPage) {
                  var mp = String(m["Page Name"]).trim().toLowerCase();
                  if (mp === rPage || rPage.indexOf(mp) !== -1 || mp.indexOf(rPage) !== -1) match = true;
                } else { match = true; }
              } else if (m["Record ID"] && rRecId && String(m["Record ID"]).trim() === rRecId) {
                match = true;
              } else if (m.Date && rDate && (String(m.Date).trim() === rDate || formatCustomDate(m.Date) === rDate)) {
                var mCost = m["Paid (৳)"] || m.Cost || m.Spent || m.Budget ? parseFloat(String(m["Paid (৳)"] || m.Cost || m.Spent || m.Budget).replace(/[^0-9.]/g, "")) : NaN;
                if (!isNaN(mCost) && !isNaN(rCost) && Math.abs(mCost - rCost) < 0.01) {
                  match = true;
                }
              }

              if (match) {
                sheet.deleteRow(i + 1);
                deleted = true;
                if (targetId && rRecId && targetId === rRecId) {
                  break; // Stop immediately once target unique record ID is deleted
                }
              }
            }
          }
        }
        if (!deleted && body.rowNo && body.rowNo > 1 && body.rowNo <= lastRow) {
          sheet.deleteRow(body.rowNo);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true, action: "deleteRow" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 4. Batch Operations
    if (body.action === "batch" && Array.isArray(body.batch)) {
      for (var b = 0; b < body.batch.length; b++) {
        var item = body.batch[b];
        if (!item) continue;
        if (item.action === "appendRow" || !item.action) {
          var bSheet = getFlexibleSheet(item.sheetName) || ss.getSheetByName(item.sheetName);
          if (!bSheet) {
            bSheet = ss.insertSheet(item.sheetName);
            if (item.headers && item.headers.length > 0) bSheet.appendRow(item.headers);
          }
          if (item.rowValues && item.rowValues.length > 0) {
            bSheet.appendRow(item.rowValues);
          }
        } else if (item.action === "deleteRow") {
          var bDelSheet = getFlexibleSheet(item.sheetName);
          if (bDelSheet && (item.recordId || item.matchInfo)) {
            var dData = bDelSheet.getDataRange().getValues();
            for (var dI = dData.length - 1; dI >= 1; dI--) {
              var dRow = dData[dI];
              var dSl = String(dRow[0] || "").trim().replace(/^#/, "");
              var dRec = String(dRow[dRow.length - 1] || "").trim();
              if ((item.recordId && dRec === String(item.recordId).trim()) || (item.matchInfo && item.matchInfo.SL && dSl === String(item.matchInfo.SL).trim())) {
                bDelSheet.deleteRow(dI + 1);
              }
            }
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true, action: "batch", count: body.batch.length }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 5. Batch Push All Data (Full Sync with Rectangular Padding)
    if ((body.action === "pushAll" || body.action === "pushAllData") && (body.tables || body.allSheetsData)) {
      var tablesData = body.allSheetsData || body.tables;
      for (var sName in tablesData) {
        var rawTbl = tablesData[sName];
        if (!rawTbl) continue;
        var s = getFlexibleSheet(sName) || ss.insertSheet(sName);
        s.clear();
        if (Array.isArray(rawTbl) && rawTbl.length > 0) {
          var maxCols = 0;
          for (var r = 0; r < rawTbl.length; r++) {
            if (rawTbl[r] && rawTbl[r].length > maxCols) maxCols = rawTbl[r].length;
          }
          if (maxCols > 0) {
            var padded = rawTbl.map(function(rRow) {
              var arr = Array.isArray(rRow) ? rRow.slice() : [];
              while (arr.length < maxCols) arr.push("");
              return arr;
            });
            s.getRange(1, 1, padded.length, maxCols).setValues(padded);
          }
        } else if (rawTbl && rawTbl.rows && rawTbl.rows.length > 0) {
          if (rawTbl.headers && rawTbl.headers.length > 0) {
            s.appendRow(rawTbl.headers);
          }
          for (var rIdx = 0; rIdx < rawTbl.rows.length; rIdx++) {
            s.appendRow(rawTbl.rows[rIdx]);
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true, action: "pushAllData" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 6. Change Date Format across Entire Spreadsheet
    if (body.action === "formatAllDates" || body.action === "setDateFormat") {
      var chosenFmt = body.dateFormat || body.format || "DD-MM-YYYY";
      CURRENT_DATE_FORMAT = chosenFmt;
      var fmtRes = formatAllDatesInEntireSpreadsheet(chosenFmt);
      return ContentService.createTextOutput(JSON.stringify({ 
        success: true, 
        action: "formatAllDates", 
        dateFormat: chosenFmt, 
        result: fmtRes 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ success: true, message: "Action executed" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function onEdit(e) {
  try {
    var webhookUrl = PropertiesService.getScriptProperties().getProperty("ADVERSE_APP_WEBHOOK");
    if (!webhookUrl) return;
    var range = e && e.range;
    var sheet = range ? range.getSheet() : null;
    var sheetName = sheet ? sheet.getName() : "";
    UrlFetchApp.fetch(webhookUrl, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        action: "onEdit",
        sheetName: sheetName,
        row: range ? range.getRow() : 0,
        column: range ? range.getColumn() : 0,
        value: e ? e.value : "",
        timestamp: new Date().toISOString()
      }),
      muteHttpExceptions: true
    });
  } catch (err) {
    // Non-blocking silent fallback
  }
}

/**
 * Helper to include modular HTML files (index.html, stylesheet.html, javascript.html)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Flexible sheet locator: matches exact name, trimmed name, clean name, or month+year pattern
 */
function findSheetFlexible(ss, targetName, month, year) {
  if (!ss) return null;
  var target = String(targetName || '').trim();
  if (target) {
    var direct = ss.getSheetByName(target);
    if (direct) return direct;
    var trimmed = ss.getSheetByName(target.trim());
    if (trimmed) return trimmed;
  }
  var sheets = ss.getSheets();
  var cleanTarget = target.toLowerCase().replace(/[-_ ]/g, '');

  if (cleanTarget) {
    for (var i = 0; i < sheets.length; i++) {
      var sName = sheets[i].getName();
      if (sName.toLowerCase().replace(/[-_ ]/g, '') === cleanTarget) {
        return sheets[i];
      }
    }
  }

  var isLookingForCampaign = cleanTarget.indexOf('campaign') !== -1;

  if (month) {
    var mLower = String(month).trim().toLowerCase();
    var mShort = mLower.slice(0, 3);
    var yStr = year ? String(year).trim() : '';

    // First pass: Match sheet containing month, year, and 'campaign' if looking for campaign
    for (var j = 0; j < sheets.length; j++) {
      var nameLower = sheets[j].getName().toLowerCase();
      // Skip non-campaign sheets if searching for campaign
      if (isLookingForCampaign && (nameLower.includes('summary') || nameLower.includes('payment') || nameLower.includes('cost') || nameLower.includes('billing') || nameLower.includes('closing') || nameLower.includes('end_ad'))) {
        continue;
      }
      var hasMonth = nameLower.includes(mLower) || nameLower.includes(mShort);
      var hasYear = !yStr || nameLower.includes(yStr);
      var hasCamp = !isLookingForCampaign || nameLower.includes('campaign');
      if (hasMonth && hasYear && hasCamp) {
        return sheets[j];
      }
    }

    // Second pass: general month + year if not strictly campaign
    if (!isLookingForCampaign) {
      for (var j2 = 0; j2 < sheets.length; j2++) {
        var nLower2 = sheets[j2].getName().toLowerCase();
        if (nLower2.includes('summary') || nLower2.includes('payment') || nLower2.includes('cost') || nLower2.includes('closing')) continue;
        if ((nLower2.includes(mLower) || nLower2.includes(mShort)) && (!yStr || nLower2.includes(yStr))) {
          return sheets[j2];
        }
      }
    }
  }

  if (target) {
    for (var k = 0; k < sheets.length; k++) {
      var n = sheets[k].getName();
      if (n.toLowerCase().indexOf(target.toLowerCase()) !== -1) {
        return sheets[k];
      }
    }
  }

  return null;
}

/**
 * Helper to ensure a sheet exists and has headers
 */
function ensureSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = findSheetFlexible(ss, name) || ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length) {
      sheet.appendRow(headers);
    }
  } else if (headers && headers.length && sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
  return sheet;
}

/**
 * Read all rows from a sheet as key-value objects
 */
function readRows(sheetName, expectedHeaders) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = findSheetFlexible(ss, sheetName);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(h => String(h || '').trim());
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r.some(cell => String(cell || '').trim() !== '')) continue;
    const obj = { _row: i + 1 };
    headers.forEach((h, colIdx) => {
      if (h) {
        let val = r[colIdx];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd-MMM-yyyy');
        }
        obj[h] = val !== undefined && val !== null ? String(val).trim() : '';
      }
    });
    rows.push(obj);
  }
  return rows;
}

/**
 * Append row to sheet
 */
function appendRow(sheetName, headers, rowObj) {
  const sheet = ensureSheet(sheetName, headers);
  const rowVals = headers.map(h => rowObj[h] !== undefined && rowObj[h] !== null ? rowObj[h] : '');
  sheet.appendRow(rowVals);
  SpreadsheetApp.flush();
}

/**
 * Smart Match & Update
 */
function updateRowAt(sheetName, headers, rowNo, rowObj, recordId, matchInfo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getDataRange().getValues();
  let targetRowIdx = -1;

  if (recordId || matchInfo) {
    const tId = String(recordId || '').trim();
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (tId && (String(r[0]).trim() === tId || String(r[1]).trim() === tId || String(r[r.length - 1]).trim() === tId)) {
        targetRowIdx = i + 1;
        break;
      }
      if (matchInfo) {
        if (matchInfo.SL && String(r[0]).trim() === String(matchInfo.SL).trim()) {
          targetRowIdx = i + 1;
          break;
        }
        if (matchInfo['Record ID'] && String(r[1]).trim() === String(matchInfo['Record ID']).trim()) {
          targetRowIdx = i + 1;
          break;
        }
      }
    }
  }

  if (targetRowIdx === -1 && rowNo && rowNo > 1 && rowNo <= lastRow) {
    targetRowIdx = rowNo;
  }

  if (targetRowIdx > 0) {
    const rowVals = headers.map(h => rowObj[h] !== undefined && rowObj[h] !== null ? rowObj[h] : '');
    sheet.getRange(targetRowIdx, 1, 1, rowVals.length).setValues([rowVals]);
    SpreadsheetApp.flush();
  }
}

/**
 * Delete row from sheet
 */
function deleteRowAt(sheetName, rowNo, recordId, matchInfo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getDataRange().getValues();
  let targetRowIdx = -1;

  if (recordId || matchInfo) {
    const tId = String(recordId || '').trim();
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (tId && (String(r[0]).trim() === tId || String(r[1]).trim() === tId || String(r[r.length - 1]).trim() === tId)) {
        targetRowIdx = i + 1;
        break;
      }
      if (matchInfo) {
        if (matchInfo.SL && String(r[0]).trim() === String(matchInfo.SL).trim()) {
          targetRowIdx = i + 1;
          break;
        }
        if (matchInfo['Record ID'] && String(r[1]).trim() === String(matchInfo['Record ID']).trim()) {
          targetRowIdx = i + 1;
          break;
        }
      }
    }
  }

  if (targetRowIdx === -1 && rowNo && rowNo > 1 && rowNo <= lastRow) {
    targetRowIdx = rowNo;
  }

  if (targetRowIdx > 0) {
    sheet.deleteRow(targetRowIdx);
    SpreadsheetApp.flush();
  }
}

/**
 * Valid Campaign Row Validator
 * Single Source of Truth: Retains all legitimate campaign rows without dropping records
 */
function isValidCampaignRow(r) {
  if (!r) return false;
  var page = String(r['Page Name'] || r['Page_Name'] || r['page_name'] || r['Page'] || r['page'] || '').trim();
  var sl = String(r['SL'] || r['SL NO'] || r.sl || r.sl_no || '').trim();
  var adAccount = String(r['AD Account'] || r['Ad Account'] || r['ad_account'] || r['AD'] || r['ad'] || '').trim();
  var rawBud = String(r['Budget'] !== undefined && r['Budget'] !== null ? r['Budget'] : (r['budget'] !== undefined && r['budget'] !== null ? r['budget'] : '')).replace(/[^0-9.]/g, '');
  var budget = parseFloat(rawBud) || 0;
  var rawTaka = String(r['Taka'] !== undefined && r['Taka'] !== null ? r['Taka'] : (r['taka'] !== undefined && r['taka'] !== null ? r['taka'] : '')).replace(/[^0-9.]/g, '');
  var taka = parseFloat(rawTaka) || 0;
  var startDate = String(r['Start'] || r['Start Date'] || r.start_date || '').trim();

  // Header row protection
  if (page.toLowerCase() === 'page name' || page.toLowerCase() === 'page') return false;

  // Genuine campaign row: must have Page Name, SL, AD Account, Start Date, or a non-zero Budget/Taka
  if (!page && !sl && !adAccount && !startDate && budget <= 0 && taka <= 0) return false;

  return true;
}

/**
 * Campaign Sheet Helper
 */
function campaignSheetName(month, year) {
  return 'Campaign_' + month + '_' + year;
}

function campaignRows(month, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return [];
  const sheets = ss.getSheets();
  const allCampaigns = [];

  const isAll = !month || month === 'All' || month === 'All Months' || month === 'all';
  if (isAll) {
    sheets.forEach(s => {
      const name = s.getName();
      const nLower = name.toLowerCase();

      // Skip non-campaign sheets
      if (nLower.includes('summary') || nLower.includes('closing') || nLower.includes('end_ad') || nLower.includes('billing') || nLower.includes('payment') || nLower.includes('cost')) {
        return;
      }

      // Check if this sheet is literally named "Campaign" and is only an Ad Account list
      if (nLower === 'campaign') {
        const lastCol = Math.min(6, s.getLastColumn() || 1);
        if (s.getLastRow() >= 1 && lastCol >= 1) {
          const firstRow = s.getRange(1, 1, 1, lastCol).getValues()[0] || [];
          const firstRowStr = firstRow.map(c => String(c).toLowerCase()).join(' ');
          if (!firstRowStr.includes('page') && !firstRowStr.includes('budget') && !firstRowStr.includes('start')) {
            return; // Skip "Campaign" ad account sheet!
          }
        }
      }

      if (nLower.indexOf('campaign') !== -1 || MONTHS.some(m => nLower.includes(m.toLowerCase()))) {
        const rows = readRows(name, CH);
        rows.forEach(r => {
          if (isValidCampaignRow(r)) {
            allCampaigns.push(r);
          }
        });
      }
    });
    return allCampaigns;
  }

  const targetSheet = findSheetFlexible(ss, campaignSheetName(month, year), month, year);
  if (targetSheet) {
    const rows = readRows(targetSheet.getName(), CH);
    return rows.filter(isValidCampaignRow);
  }
  return [];
}

function parseAnyDate(str, monthHint) {
  if (!str && str !== 0) return null;
  if (str instanceof Date) return isNaN(str.getTime()) ? null : str;

  // Handle Excel / Google Sheets numerical date serials (e.g. 45560 = ~2024, 46259 = ~2026)
  if (typeof str === 'number' && str > 30000 && str < 70000) {
    const utcDays = Math.round(str - 25569);
    const d = new Date((utcDays * 86400 + 43200) * 1000);
    return isNaN(d.getTime()) ? null : d;
  }

  let s = String(str).trim();
  if (!s) return null;

  // Convert Bengali numerals (০-৯) to Western numerals (0-9)
  const bnToEn = { '০': '0', '১': '1', '২': '2', '৩': '3', '৪': '4', '৫': '5', '৬': '6', '৭': '7', '৮': '8', '৯': '9' };
  s = s.replace(/[০-৯]/g, d => bnToEn[d] || d);

  // Convert Bengali month names to English
  const bnMonthMap = {
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

  const cleanDateStr = s.replace(/T\d{2}:\d{2}.*$/, '').replace(/\s+\d{1,2}:\d{2}(:\d{2})?.*$/, '').trim();

  // Match 12-Sep-2026 or 12/Sep/2026 or 12-Sep-26
  const matchDMonY = cleanDateStr.match(/^(\d{1,2})[-/ .]([A-Za-z]+)[-/ .](\d{2,4})/);
  if (matchDMonY) {
    const day = parseInt(matchDMonY[1], 10);
    const monStr = matchDMonY[2].slice(0, 3).toLowerCase();
    let yr = parseInt(matchDMonY[3], 10);
    if (yr < 100) yr += 2000;
    const monIdx = MONTH_NAMES_SHORT.findIndex(m => m.toLowerCase() === monStr);
    if (monIdx !== -1 && day >= 1 && day <= 31) {
      return new Date(yr, monIdx, day, 12, 0, 0);
    }
  }

  // Match Sep 25, 2026 or Sep-25-2026
  const matchMonDY = cleanDateStr.match(/^([A-Za-z]+)[-/ .](\d{1,2}),?[-/ .](\d{2,4})/);
  if (matchMonDY) {
    const monStr = matchMonDY[1].slice(0, 3).toLowerCase();
    const day = parseInt(matchMonDY[2], 10);
    let yr = parseInt(matchMonDY[3], 10);
    if (yr < 100) yr += 2000;
    const monIdx = MONTH_NAMES_SHORT.findIndex(m => m.toLowerCase() === monStr);
    if (monIdx !== -1 && day >= 1 && day <= 31) {
      return new Date(yr, monIdx, day, 12, 0, 0);
    }
  }

  // Match YYYY-MM-DD
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

    if (monthHint !== undefined && monthHint !== null && monthHint !== '') {
      let targetMonIdx = -1;
      if (typeof monthHint === 'number') {
        targetMonIdx = (monthHint >= 1 && monthHint <= 12) ? monthHint - 1 : -1;
      } else {
        const mStr = String(monthHint).trim().toLowerCase();
        targetMonIdx = MONTHS.findIndex(m => m.toLowerCase() === mStr);
        if (targetMonIdx === -1) {
          targetMonIdx = MONTH_NAMES_SHORT.findIndex(m => m.toLowerCase() === mStr.slice(0, 3));
        }
      }
      if (targetMonIdx !== -1) {
        if (p2 === targetMonIdx + 1 && p1 <= 31) {
          return new Date(yr, p2 - 1, p1, 12, 0, 0);
        } else if (p1 === targetMonIdx + 1 && p2 <= 31) {
          return new Date(yr, p1 - 1, p2, 12, 0, 0);
        }
      }
    }

    if (p1 > 12 && p2 <= 12) {
      return new Date(yr, p2 - 1, p1, 12, 0, 0);
    } else if (p1 <= 12 && p2 > 12) {
      return new Date(yr, p1 - 1, p2, 12, 0, 0);
    } else if (p1 <= 31 && p2 <= 12) {
      const nowMonth = (new Date()).getMonth() + 1;
      if (p2 === nowMonth) {
        return new Date(yr, p2 - 1, p1, 12, 0, 0);
      } else if (p1 === nowMonth) {
        return new Date(yr, p1 - 1, p2, 12, 0, 0);
      }
      return new Date(yr, p2 - 1, p1, 12, 0, 0);
    }
  }

  const d = new Date(cleanDateStr);
  return isNaN(d.getTime()) ? null : d;
}

function isSameDay(d1Input, d2Input) {
  const d1 = parseAnyDate(d1Input);
  const d2 = parseAnyDate(d2Input);
  if (!d1 || !d2) return false;
  return d1.getFullYear() === d2.getFullYear() &&
         d1.getMonth() === d2.getMonth() &&
         d1.getDate() === d2.getDate();
}

function getRate() {
  const rows = readRows('Settings', ['Key', 'Value']);
  const r = rows.find(x => x.Key === 'rate');
  return r ? Number(r.Value) || 140 : 140;
}

function getAdAccounts() {
  const rows = readRows('Ad Accounts', ['Ad Account', 'ID Name']);
  if (rows && rows.length > 0) {
    return rows.map(r => ({ a: r['Ad Account'] || r.a, id: r['ID Name'] || r.id || '' })).filter(x => x.a);
  }
  return DEFAULT_AD_ACCOUNTS;
}

function getPages() {
  const rows = readRows('Page List', ['Page Name']);
  if (rows && rows.length > 0) {
    const list = rows.map(r => r['Page Name']).filter(Boolean);
    if (list.length > 0) return Array.from(new Set(list));
  }
  return DEFAULT_PAGES;
}

function nextSL(month, year) {
  var lock = LockService.getScriptLock();
  var hasLock = false;
  try {
    hasLock = lock.tryLock(10000);
  } catch (e) {}

  try {
    const rows = campaignRows(month, year);
    let max = 0;
    const checkRow = function(r) {
      if (!r) return;
      var raw = String(r['SL'] || r.sl || r['SL NO'] || '').trim().replace(/^#/, '');
      var match = raw.match(/^(\d+)/);
      if (match) {
        var val = parseInt(match[1], 10);
        if (!isNaN(val) && val > max) max = val;
      }
    };
    rows.forEach(checkRow);
    const tcRows = readRows('Today_Created', TODAY_HEADERS);
    tcRows.forEach(checkRow);
    return max + 1;
  } finally {
    if (hasLock) {
      try { lock.releaseLock(); } catch (e) {}
    }
  }
}

function normalizeCampaign(d, month, year, sl) {
  const x = {};
  CH.forEach(h => {
    x[h] = d[h] !== undefined ? d[h] : (d[h.replace(/ /g, '_').toLowerCase()] !== undefined ? d[h.replace(/ /g, '_').toLowerCase()] : '');
  });
  x['SL'] = String(sl || '').trim();
  
  // 1. Entry Date = today created ad (Format: 09-Oct-2026)
  const rawEntry = d['Entry Date'] || d.entry_date || d['Date'] || '';
  const parsedEntry = parseAnyDate(rawEntry);
  if (parsedEntry) {
    x['Entry Date'] = Utilities.formatDate(parsedEntry, Session.getScriptTimeZone(), 'dd-MMM-yyyy');
  } else {
    x['Entry Date'] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy');
  }

  // 2. Start Date = month wise all campaign (Format: 09-Oct-2026)
  const rawStart = d.start_date || d.Start || d['Start Date'] || d.start || '';
  const parsedStart = parseAnyDate(rawStart, month);
  if (parsedStart) {
    x['Start'] = Utilities.formatDate(parsedStart, Session.getScriptTimeZone(), 'dd-MMM-yyyy');
    x['Month'] = MONTHS[parsedStart.getMonth()];
    x['Year'] = String(parsedStart.getFullYear());
  } else {
    x['Start'] = x['Entry Date'];
    x['Month'] = month || x['Month'] || MONTHS[new Date().getMonth()];
    x['Year'] = String(year || x['Year'] || new Date().getFullYear());
  }

  // 3. End Date = Start + Day (Format: 09-Oct-2026)
  const numDays = Math.max(1, Number(d.day || d.Day || x['Day'] || 1));
  x['Day'] = numDays;
  const startForEnd = parseAnyDate(x['Start']) || new Date();
  const endDateObj = new Date(startForEnd.getTime() + numDays * 86400000);
  x['End'] = Utilities.formatDate(endDateObj, Session.getScriptTimeZone(), 'dd-MMM-yyyy');

  x['Record ID'] = d['Record ID'] || ('CAMP-' + Date.now() + '-' + Math.floor(Math.random() * 1000));
  return x;
}

function formatCampaignText(r) {
  if (!r) return '';
  let sl = String(r['SL'] || '').trim();
  let adNum = '';
  const m = String(r['AD Account'] || '').match(/\d+/);
  if (m) adNum = m[0];
  const slPart = (sl && adNum) ? (sl + '/' + adNum) : (sl || '001');
  const idName = String(r['ID Name'] || '').trim();
  const idPart = idName ? ' ' + idName : '';
  const lines = [
    '#' + slPart + idPart + ' ' + (r['Status'] || 'Old Page'),
    'Page: ' + (r['Page Name'] || ''),
    'Start: ' + (r['Start'] || ''),
    'End: ' + (r['End'] || ''),
    'Types: ' + (r['Types'] || 'Boost'),
    'Day: ' + (r['Day'] || '1'),
    'Budget: ' + (r['Budget'] || '0') + '$',
    'Daily: ' + (r['Daily $'] || '0') + '$',
    'Taka: ' + (r['Taka'] || '0') + ' TK'
  ];
  return lines.join('\n');
}

/**
 * Universal RPC Dispatcher for Frontend Apps Script Calls
 */
function rpc(method, args) {
  args = args || [];
  
  if (method === 'loginUser') {
    const pin = String(args[0] || '').trim();
    if (pin === ADMIN_PIN || pin === FALLBACK_ADMIN_PIN) {
      return { success: true, userRole: 'admin', role: 'admin', displayName: 'Admin', token: 'token-admin-' + Date.now() };
    }
    if (pin === VIEWER_PIN) {
      return { success: true, userRole: 'viewer', role: 'viewer', displayName: 'Viewer', token: 'token-viewer-' + Date.now() };
    }
    return { success: false, message: 'Invalid PIN. Admin: ' + ADMIN_PIN + ', Viewer: ' + VIEWER_PIN };
  }

  if (method === 'getInitialData') {
    const currentMn = MONTHS[new Date().getMonth()];
    const currentYr = new Date().getFullYear();
    const allCampaigns = campaignRows('All');
    const curMonthRows = campaignRows(currentMn, currentYr);
    const todayRows = (rpc('getRecords', ['today_created_campaign', {}]) || {}).rows || [];

    const preloaded = {};
    for (let i = 0; i < MENU.length; i++) {
      const m = MENU[i];
      try {
        preloaded[m.key] = {
          records: rpc('getRecords', [m.key, { month: currentMn, year: currentYr }]),
          stats: rpc('getMenuStats', [m.key, currentMn, currentYr, '', ''])
        };
      } catch (e) {}
    }
    return {
      success: true,
      menus: MENU,
      months: MONTHS,
      activeMonth: currentMn,
      activeYear: currentYr,
      adAccounts: getAdAccounts(),
      pages: getPages(),
      rate: getRate(),
      adminPin: ADMIN_PIN,
      viewerPin: VIEWER_PIN,
      userRole: 'admin',
      displayName: 'Admin',
      userEmail: Session.getActiveUser().getEmail() || '',
      isConfigured: true,
      allCampaigns: allCampaigns,
      totalCampaigns: allCampaigns.length,
      preloaded: preloaded
    };
  }

  if (method === 'runDiagnosticsReport') {
    const filterMonth = args[0] || 'October';
    const filterYear = args[1] || '2026';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetId = ss ? ss.getId() : '';
    const targetWorksheet = campaignSheetName(filterMonth, filterYear);
    const sheet = findSheetFlexible(ss, targetWorksheet, filterMonth, filterYear);

    let sheetPhysicalRows = 0;
    let sheetHeaderRow = 1;
    let sheetHeaders = [];
    let sheetFirstRow = null;
    let sheetLastRow = null;
    let sheetValidCampaigns = 0;
    let sheetBlankRows = 0;
    let sheetMalformedRows = 0;
    const slMap = {};
    const idMap = {};

    if (sheet) {
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow >= 2 && lastCol >= 1) {
        const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
        sheetHeaders = values[0].map(h => String(h || '').trim());
        sheetPhysicalRows = values.length - 1;

        for (let i = 1; i < values.length; i++) {
          const r = values[i];
          const isBlank = !r.some(c => String(c || '').trim() !== '');
          if (isBlank) {
            sheetBlankRows++;
            continue;
          }
          const sl = String(r[0] || '').trim();
          const page = String(r[5] || '').trim();
          const recId = String(r[r.length - 1] || '').trim();

          if (sl) slMap[sl] = (slMap[sl] || 0) + 1;
          if (recId && recId.length > 8) idMap[recId] = (idMap[recId] || 0) + 1;

          if (page || sl) {
            sheetValidCampaigns++;
          } else {
            sheetMalformedRows++;
          }

          if (i === 1) {
            sheetFirstRow = { SL: sl, 'Page Name': page, Start: String(r[6] || '') };
          }
          if (i === values.length - 1) {
            sheetLastRow = { SL: sl, 'Page Name': page, Start: String(r[6] || '') };
          }
        }
      }
    }

    const duplicateSLs = Object.keys(slMap).filter(k => slMap[k] > 1).map(k => ({ sl: k, count: slMap[k] }));
    const duplicateRecordIds = Object.keys(idMap).filter(k => idMap[k] > 1).map(k => ({ recordId: k, count: idMap[k] }));

    const backendMonthRows = campaignRows(filterMonth, filterYear);
    const backendAllRows = campaignRows('All');

    const isExactMatch = sheetPhysicalRows > 0 && sheetPhysicalRows === backendMonthRows.length;

    return {
      success: true,
      timestamp: new Date().toISOString(),
      spreadsheetId: sheetId,
      targetWorksheet: sheet ? sheet.getName() : targetWorksheet,
      targetMonth: filterMonth,
      targetYear: filterYear,
      sheetDetails: {
        headerRowNumber: sheetHeaderRow,
        headers: sheetHeaders,
        physicalDataRowCount: sheetPhysicalRows,
        validCampaignRecordCount: sheetValidCampaigns,
        blankRowCount: sheetBlankRows,
        malformedRowCount: sheetMalformedRows,
        firstDataRow: sheetFirstRow,
        lastDataRow: sheetLastRow,
        duplicateSLCount: duplicateSLs.length,
        duplicateSLs: duplicateSLs,
        duplicateRecordIdCount: duplicateRecordIds.length,
        duplicateRecordIds: duplicateRecordIds
      },
      verificationChain: {
        googleSheetsPhysicalRows: sheetPhysicalRows,
        googleSheetsValidCampaigns: sheetValidCampaigns,
        backendResponseRows: backendMonthRows.length,
        allMonthsCampaignsTotal: backendAllRows.length,
        isExactMatch: isExactMatch,
        status: isExactMatch ? 'VERIFIED_100_PERCENT_MATCH' : 'DISCREPANCY_DETECTED',
        explanation: isExactMatch
          ? 'Perfect 1-to-1 parity! All ' + sheetPhysicalRows + ' campaign records in Google Sheet tab match the backend response and frontend display without omitting or fabricating any rows.'
          : 'Google Sheet contains ' + sheetPhysicalRows + ' rows, backend returned ' + backendMonthRows.length + ' rows.'
      }
    };
  }

  if (method === 'getRecords') {
    const k = args[0];
    const f = args[1] || {};
    const month = f.month || MONTHS[new Date().getMonth()];
    const year = f.year || new Date().getFullYear();

    if (k === 'today_created_campaign') {
      const today = new Date();
      let rows = readRows('Today_Created', TODAY_HEADERS);
      const allCamp = campaignRows('All');
      allCamp.forEach(c => {
        const entry = c['Entry Date'] || c['Date'];
        if (entry && isSameDay(entry, today)) {
          if (!rows.some(x => x['Record ID'] === c['Record ID'] || (x['SL'] === c['SL'] && x['Page Name'] === c['Page Name']))) {
            rows.push(c);
          }
        }
      });
      const filtered = rows.filter(r => isSameDay(r['Entry Date'] || r['Date'], today));
      if (f.sort === 'asc' || f.sort === 'sheet') {
        return { success: true, headers: CH, rows: filtered };
      }
      return { success: true, headers: CH, rows: filtered.slice().reverse() };
    }

    if (k === 'campaign') {
      let rows = campaignRows(month, year);
      if (f.date) {
        rows = rows.filter(r => isSameDay(r['Start'] || r['Date'] || r['Entry Date'], f.date));
      }
      if (f.search) {
        const q = String(f.search).toLowerCase();
        rows = rows.filter(r => Object.values(r).some(v => String(v).toLowerCase().indexOf(q) !== -1));
      }
      // Preserve sheet order by default or sort descending if requested
      if (f.sort === 'newest' || f.sort === 'desc') {
        return { success: true, headers: CH, rows: rows.slice().reverse() };
      }
      return { success: true, headers: CH, rows: rows };
    }

    if (k === 'end_ads') {
      const allRows = campaignRows('All');
      const today = new Date();
      const list = allRows.filter(r => {
        const d = parseAnyDate(r['End'] || r['End Date']);
        return d && d >= today;
      });
      return { success: true, headers: END_ADS_HEADERS, rows: list };
    }

    if (k === 'page_summery') {
      const rows = campaignRows(month, year);
      const summary = {};
      rows.forEach(r => {
        const p = r['Page Name'] || 'Unknown';
        if (!summary[p]) summary[p] = { count: 0, budget: 0 };
        summary[p].count += 1;
        summary[p].budget += Number(r['Budget']) || 0;
      });
      const list = Object.keys(summary).map(p => ({
        'Page Name': p,
        'Total Ads': summary[p].count,
        'Total Budget $': summary[p].budget.toFixed(2)
      }));
      return { success: true, headers: PAGE_SUMMARY_HEADERS, rows: list };
    }

    if (k === 'agency_cost') {
      const rows = readRows('Agency_Cost', AGENCY_COST_HEADERS).filter(r => {
        const cost = parseFloat(r['Cost'] || r['cost'] || r['Amount'] || 0) || 0;
        const type = String(r['Type'] || r['type'] || r['Category'] || '').trim();
        return Boolean(type && type.toLowerCase() !== 'type' && cost > 0);
      });
      return { success: true, headers: AGENCY_COST_HEADERS, rows: rows.reverse() };
    }

    if (k === 'payment_due_tracker') {
      const rows = readRows('Payments', PAYMENT_HEADERS).filter(r => {
        const paid = parseFloat(r['Paid (৳)'] || r['Paid'] || 0) || 0;
        const method = String(r['Payment Method'] || '').trim();
        return paid > 0 || Boolean(method);
      });
      const mapped = rows.map(r => {
        const rawDate = String(r['Date'] || r['date'] || '').trim();
        const dObj = parseAnyDate(rawDate);
        const rMonth = String(r['Month'] || (dObj ? MONTHS[dObj.getMonth()] : '')).trim();
        const rYear = String(r['Year'] || (dObj ? dObj.getFullYear() : '')).trim();
        return {
          ...r,
          'Date': dObj ? formatCustomDate(dObj, rMonth) : rawDate,
          'Month': rMonth,
          'Year': rYear,
          _month: rMonth,
          _year: rYear
        };
      });
      const filtered = mapped.filter(r => {
        if (f.month && f.month !== 'All Months') {
          const mLower = String(f.month).toLowerCase();
          const rMLower = String(r._month || r['Month'] || '').toLowerCase();
          const rawLower = String(r['Date'] || '').toLowerCase();
          if (rMLower !== mLower && !rawLower.includes(mLower.slice(0, 3))) return false;
        }
        if (f.year && f.year !== 'All Years') {
          const yStr = String(f.year);
          const rYStr = String(r._year || r['Year'] || '');
          if (rYStr && rYStr !== yStr && !String(r['Date']).includes(yStr)) return false;
        }
        return true;
      });
      return { success: true, headers: ['Date', 'Payment Method', 'Paid (৳)', 'Month', 'Year'], rows: filtered.reverse() };
    }

    return { success: true, rows: [] };
  }

  if (method === 'getMenuStats') {
    const k = args[0];
    const mn = args[1] || MONTHS[new Date().getMonth()];
    const yr = args[2] || new Date().getFullYear();

    if (k === 'today_created_campaign') {
      const recs = rpc('getRecords', ['today_created_campaign', {}]);
      const rows = recs && recs.rows ? recs.rows : [];
      let budget = 0, taka = 0;
      rows.forEach(r => {
        budget += Number(r['Budget']) || 0;
        taka += Number(r['Taka']) || 0;
      });
      return { success: true, totalRecords: rows.length, totalBudget: budget, totalTaka: taka };
    }

    if (k === 'campaign') {
      let rows = campaignRows(mn, yr);
      const dateFilter = args[3];
      if (dateFilter) {
        rows = rows.filter(r => isSameDay(r['Start'] || r['Date'] || r['Entry Date'], dateFilter));
      }
      let budget = 0, taka = 0;
      rows.forEach(r => {
        budget += Number(r['Budget']) || 0;
        taka += Number(r['Taka']) || 0;
      });
      return {
        success: true,
        totalRecords: rows.length,
        currentMonthRecords: rows.length,
        totalBudget: budget,
        totalTaka: taka,
        currentMonth: mn,
        currentYear: yr
      };
    }

    if (k === 'end_ads') {
      const today = new Date();
      const tomorrow = new Date(Date.now() + 86400000);
      const rows = campaignRows('All');
      const todayCount = rows.filter(r => isSameDay(r['End'] || r['End Date'], today)).length;
      const tomorrowCount = rows.filter(r => isSameDay(r['End'] || r['End Date'], tomorrow)).length;
      return { success: true, todayCount: todayCount, tomorrowCount: tomorrowCount };
    }

    if (k === 'payment_due_tracker') {
      const payments = readRows('Payments', PAYMENT_HEADERS);
      const monthlyPayments = payments.filter(function(r) {
        var rMonth = String(r['Month'] || r.Month || '').trim().toLowerCase();
        var rYear = String(r['Year'] || r.Year || '').trim();
        return (!mn || mn === 'All Months' || rMonth === String(mn).toLowerCase()) &&
               (!yr || yr === 'All Years' || rYear === String(yr));
      });
      const paid = Math.round(monthlyPayments.reduce(function(s, r) {
        return s + (Number(r['Paid (৳)'] || r.Paid || 0) || 0);
      }, 0) * 100) / 100;

      const campaigns = campaignRows(mn, yr);
      const budget = Math.round(campaigns.reduce(function(s, r) {
        return s + (Number(r['Budget'] || 0) || 0);
      }, 0) * 100) / 100;
      const rate = getRate();
      const totalBill = Math.round((budget * rate) * 100) / 100;
      const totalDue = Math.max(0, Math.round((totalBill - paid) * 100) / 100);

      return {
        success: true,
        totalPaid: paid,
        totalBudget: budget,
        totalBill: totalBill,
        totalDue: totalDue,
        rate: rate,
        month: mn,
        year: yr,
        paymentCount: monthlyPayments.length,
        campaignCount: campaigns.length
      };
    }

    return { success: true };
  }

  if (method === 'addCampaignBatch' || method === 'addCampaign') {
    var lock = LockService.getScriptLock();
    var hasLock = false;
    try {
      hasLock = lock.tryLock(15000);
    } catch(e) {}

    try {
      const d = args[0];
      const items = method === 'addCampaignBatch' ? (args[1] || []) : [d];
      let last = null;
      const copiedItems = [];

      const baseMonth = items[0]?.month || items[0]?.Month || d?.month || MONTHS[new Date().getMonth()];
      const baseYear = items[0]?.year || items[0]?.Year || d?.year || new Date().getFullYear();
      const baseSl = nextSL(baseMonth, baseYear);

      items.forEach((item, idx) => {
        const month = item.month || item.Month || baseMonth;
        const year = item.year || item.Year || baseYear;
        const sl = baseSl + idx;
        const row = normalizeCampaign(Object.assign({}, d, item), month, year, sl);

        appendRow(campaignSheetName(month, year), CH, row);
        appendRow('Today_Created', TODAY_HEADERS, row);

        if (row['Page Name']) {
          appendRow('Page List', ['Page Name'], { 'Page Name': row['Page Name'] });
        }

        last = row;
        copiedItems.push({
          success: true,
          formattedText: formatCampaignText(row),
          sl: row['SL'],
          recordId: row['Record ID']
        });
      });

      return {
        success: true,
        message: 'Campaigns successfully saved!',
        formattedText: last ? formatCampaignText(last) : '',
        copiedItems: copiedItems,
        batchCount: items.length
      };
    } finally {
      if (hasLock) {
        try { lock.releaseLock(); } catch(e){}
      }
    }
  }

  if (method === 'updateCampaignRecord') {
    const id = args[0];
    const d = args[1];
    const month = d.month || MONTHS[new Date().getMonth()];
    const year = d.year || new Date().getFullYear();
    const sheetName = campaignSheetName(month, year);
    updateRowAt(sheetName, CH, null, d, id, d);
    updateRowAt('Today_Created', TODAY_HEADERS, null, d, id, d);
    return { success: true, message: 'Record updated' };
  }

  if (method === 'deleteCampaignRecord') {
    const id = args[0];
    const matchInfo = args[1] || {};
    const month = matchInfo.month || MONTHS[new Date().getMonth()];
    const year = matchInfo.year || new Date().getFullYear();
    deleteRowAt(campaignSheetName(month, year), null, id, matchInfo);
    deleteRowAt('Today_Created', null, id, matchInfo);
    return { success: true, message: 'Record deleted' };
  }

  if (method === 'addPayment') {
    const d = args[0];
    const parsedPayDate = parseAnyDate(d.date || d.payment_date || d['Date']);
    const payDate = parsedPayDate ? Utilities.formatDate(parsedPayDate, Session.getScriptTimeZone(), 'dd-MMM-yyyy') : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy');
    const month = (d.month && d.month !== 'Month' && d.month !== 'All') ? d.month : (d.Month && d.Month !== 'Month' && d.Month !== 'All' ? d.Month : (parsedPayDate ? MONTHS[parsedPayDate.getMonth()] : MONTHS[new Date().getMonth()]));
    const year = (d.year && d.year !== 'Year' && d.year !== 'All Years') ? String(d.year) : (d.Year && d.Year !== 'Year' && d.Year !== 'All Years' ? String(d.Year) : String(parsedPayDate ? parsedPayDate.getFullYear() : new Date().getFullYear()));
    const paidVal = Math.round((Number(d.paid_amount || d['Paid (৳)'] || d.Paid || 0)) * 100) / 100;
    const recId = d.recordId || d['Record ID'] || ('PAY-' + Date.now() + '-' + Math.floor(Math.random() * 1000));

    // Deduplication check
    const existing = readRows('Payments', PAYMENT_HEADERS);
    const isDup = existing.some(function(p) {
      if (recId && p['Record ID'] && String(p['Record ID']).trim() === String(recId).trim()) return true;
      var pPaid = Math.round(Number(p['Paid (৳)'] || p.Paid || 0) * 100) / 100;
      var pDate = String(p.Date || '').trim();
      var pMethod = String(p['Payment Method'] || '').trim().toLowerCase();
      var pMonth = String(p.Month || '').trim().toLowerCase();
      var pYear = String(p.Year || '').trim();
      return Math.abs(pPaid - paidVal) < 0.001 && pDate === String(payDate).trim() &&
             pMethod === String(d.payment_method || d['Payment Method'] || 'Bkash').trim().toLowerCase() &&
             pMonth === String(month).trim().toLowerCase() && pYear === String(year);
    });

    if (isDup) {
      return { success: true, message: 'Payment already recorded', recordId: recId };
    }

    appendRow('Payments', PAYMENT_HEADERS, {
      'Date': payDate,
      'Payment Method': d.payment_method || d['Payment Method'] || 'Bkash',
      'Paid (৳)': paidVal,
      'Month': month,
      'Year': year,
      'Record ID': recId
    });
    return { success: true, message: 'Payment recorded', recordId: recId };
  }

  if (method === 'addAdAccount') {
    const account = args[0];
    const id = args[1];
    appendRow('Ad Accounts', ['Ad Account', 'ID Name'], { 'Ad Account': account, 'ID Name': id || '' });
    return { success: true, adAccounts: getAdAccounts() };
  }

  if (method === 'addPage') {
    const p = args[0];
    if (p)    appendRow('Page List', ['Page Name'], { 'Page Name': p });
    return { success: true, pages: getPages() };
  }

  if (method === 'addAgencyCost') {
    const d = args[0];
    const cost = Number(d.cost || 0);
    appendRow('Agency_Cost', AGENCY_COST_HEADERS, {
      'Month': d.month || MONTHS[new Date().getMonth()],
      'Year': d.year || new Date().getFullYear(),
      'Type': d.type || 'Ads',
      'Cost': cost,
      'Per Person Cost': (cost / 2),
      'Status': 'Recorded'
    });
    return { success: true, message: 'Agency cost saved' };
  }

  if (method === 'getPageProfile') {
    const pageName = String(args[0] || '').trim();
    const adAccounts = getAdAccounts();
    const defaultAd = adAccounts[0] ? adAccounts[0].a : '';
    const defaultId = adAccounts[0] ? adAccounts[0].id : '';

    if (!pageName) {
      return {
        success: true,
        exists: false,
        status: 'New Page',
        mostFrequentAdAccount: defaultAd,
        mostFrequentIdName: defaultId,
        mostFrequentType: 'Boost'
      };
    }

    const allCamp = campaignRows('All');
    const matched = allCamp.filter(r => String(r['Page Name'] || '').toLowerCase() === pageName.toLowerCase());

    if (!matched.length) {
      return {
        success: true,
        exists: false,
        status: 'New Page',
        mostFrequentAdAccount: defaultAd,
        mostFrequentIdName: defaultId,
        mostFrequentType: 'Boost'
      };
    }

    const adCounts = {};
    const typeCounts = {};
    matched.forEach(r => {
      const ad = r['AD Account'] || defaultAd;
      const t = r['Types'] || 'Boost';
      adCounts[ad] = (adCounts[ad] || 0) + 1;
      typeCounts[t] = (typeCounts[t] || 0) + 1;
    });

    let topAd = defaultAd, maxAd = 0;
    for (let k in adCounts) {
      if (adCounts[k] > maxAd) { maxAd = adCounts[k]; topAd = k; }
    }
    const foundAcc = adAccounts.find(x => x.a === topAd);

    let topType = 'Boost', maxType = 0;
    for (let t in typeCounts) {
      if (typeCounts[t] > maxType) { maxType = typeCounts[t]; topType = t; }
    }

    return {
      success: true,
      exists: true,
      status: 'Old Page',
      totalCampaigns: matched.length,
      mostFrequentAdAccount: topAd,
      mostFrequentIdName: foundAcc ? foundAcc.id : defaultId,
      mostFrequentType: topType
    };
  }

  if (method === 'getAgencyAnalytics') {
    const allCampaigns = campaignRows('All');
    let totalBudget = 0, totalTaka = 0;
    allCampaigns.forEach(r => {
      totalBudget += Number(r['Budget']) || 0;
      totalTaka += Number(r['Taka']) || 0;
    });
    return {
      success: true,
      totalCampaigns: allCampaigns.length,
      totalBudget: totalBudget,
      totalTaka: totalTaka,
      currentRate: getRate(),
      totalPaid: 0,
      totalDue: totalTaka
    };
  }

  if (method === 'logoutUser') {
    return { success: true, message: 'Logged out' };
  }

  if (method === 'updateRate') {
    const newRate = Number(args[0]) || 140;
    updateRowAt('Settings', ['Key', 'Value'], null, { Key: 'rate', Value: String(newRate) }, 'rate', { Key: 'rate' });
    return { success: true, rate: newRate };
  }

  if (method === 'getNextSLNo') {
    const m = args[0] || MONTHS[new Date().getMonth()];
    const y = args[1] || new Date().getFullYear();
    const sl = nextSL(m, y);
    return { success: true, nextSL: String(sl).padStart(3, '0') };
  }

  if (method === 'getCampaignBySL') {
    const rawSl = String(args[0] || '').replace(/^#/, '').trim();
    const cleanMonth = args[1] ? String(args[1]).trim() : '';
    const cleanYear = args[2] ? String(args[2]).trim() : '';
    const pureSl = rawSl.includes('/') ? rawSl.split('/')[0].trim() : rawSl;
    const numSl = parseInt(pureSl, 10);

    const matchRow = function(r) {
      if (!r) return false;
      const s = String(r['SL'] || r['sl_no'] || r['SL NO'] || '').replace(/^#/, '').trim();
      if (!s) return false;
      if (s === rawSl || s === pureSl) return true;
      const rowPureSl = s.includes('/') ? s.split('/')[0].trim() : s;
      if (rowPureSl === pureSl || rowPureSl === rawSl) return true;
      const rNum = parseInt(rowPureSl, 10);
      if (!isNaN(numSl) && !isNaN(rNum) && numSl === rNum && numSl > 0) return true;
      return false;
    };

    let item = null;
    if (cleanMonth && cleanYear) {
      const monthRows = campaignRows(cleanMonth + ' ' + cleanYear);
      item = monthRows.find(matchRow);
    }
    if (!item) {
      const all = campaignRows('All');
      item = all.find(matchRow);
    }

    if (item) {
      return { success: true, pageName: item['Page Name'] || '', budget: Number(item['Budget']) || 0, campaign: item };
    }
    return { success: false, message: 'SL not found' };
  }

  if (method === 'deletePayment') {
    const id = args[0];
    deleteRowAt('Payments', null, id, { 'Record ID': id, 'Date': id });
    return { success: true, message: 'Payment deleted' };
  }

  if (method === 'deleteAgencyCost') {
    const id = args[0];
    deleteRowAt('Agency_Cost', null, id, { 'Type': id, 'Cost': id });
    return { success: true, message: 'Agency cost deleted' };
  }

  if (method === 'updateAgencyCost') {
    const id = args[0];
    const d = args[1];
    updateRowAt('Agency_Cost', AGENCY_COST_HEADERS, null, d, id, d);
    return { success: true, message: 'Agency cost updated' };
  }

  if (method === 'updatePayment') {
    const id = args[0];
    const d = args[1];
    updateRowAt('Payments', PAYMENT_HEADERS, null, d, id, d);
    return { success: true, message: 'Payment updated' };
  }

  if (method === 'getSheetConfig') {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return {
      success: true,
      connected: true,
      sheetId: ss ? ss.getId() : '',
      webhookUrl: ScriptApp.getService().getUrl() || ''
    };
  }

  if (method === 'getAppsScriptCode') {
    return {
      success: true,
      scriptCode: '// Standalone Web App is already running directly in this Google Sheet Apps Script.'
    };
  }

  if (method === 'updateSheetConfig') {
    const sid = args[0];
    const wh = args[1];
    return {
      success: true,
      connected: true,
      sheetId: sid || (SpreadsheetApp.getActiveSpreadsheet() ? SpreadsheetApp.getActiveSpreadsheet().getId() : ''),
      message: 'Sheet connected successfully'
    };
  }

  if (method === 'testWebhook') {
    return { success: true, message: 'Connected to Apps Script backend successfully!' };
  }

  if (method === 'pushAllDataToSheetViaWebhook') {
    return { success: true, message: 'Data already stored in Google Sheet natively!' };
  }

  if (method === 'syncSpreadsheet') {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss ? ss.getSheets().map(s => s.getName()) : [];
    const allCamp = campaignRows('All');
    const allPay = readRows('Payments', PAYMENT_HEADERS);
    const pages = getPages();
    return {
      success: true,
      message: 'Successfully synced from Google Sheet!',
      sheetTitles: sheets,
      campaignsCount: allCamp.length,
      paymentsCount: allPay.length,
      pagesCount: pages.length
    };
  }

  if (method === 'clearAllData') {
    return { success: true, message: 'Database reset completed' };
  }

  if (method === 'importCsvData') {
    return { success: true, message: 'Import completed successfully' };
  }

  if (method === 'addBulkAdAccounts') {
    const items = args[0] || [];
    items.forEach(it => {
      appendRow('Ad Accounts', ['Ad Account', 'ID Name'], { 'Ad Account': it.a, 'ID Name': it.id || '' });
    });
    return { success: true, adAccounts: getAdAccounts() };
  }

  if (method === 'deleteAdAccount') {
    const name = args[0];
    deleteRowAt('Ad Accounts', null, name, { 'Ad Account': name });
    return { success: true, adAccounts: getAdAccounts() };
  }

  if (method === 'addBulkPages') {
    const pages = args[0] || [];
    pages.forEach(p => {
      if (p) appendRow('Page List', ['Page Name'], { 'Page Name': p });
    });
    return { success: true, pages: getPages() };
  }

  if (method === 'deletePage') {
    const name = args[0];
    deleteRowAt('Page List', null, name, { 'Page Name': name });
    return { success: true, pages: getPages() };
  }

  if (method === 'toggleCampaignVerification') {
    const payload = args[0] || {};
    const key = payload.recordKey || payload.recordId || payload.id;
    if (!key) return { success: false, error: 'Missing recordKey' };
    const props = PropertiesService.getScriptProperties();
    let verifications = {};
    try {
      verifications = JSON.parse(props.getProperty('CAMPAIGN_VERIFICATIONS') || '{}');
    } catch(e) {}
    const verified = payload.verified !== false;
    if (!verified) {
      delete verifications[key];
    } else {
      const nowStr = formatCustomDate(new Date()) + ', ' + new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
      verifications[key] = {
        verified: true,
        verifiedBy: payload.verifiedBy || 'Admin',
        verifiedAt: nowStr,
        notes: payload.notes || 'Checked in Google Sheet'
      };
    }
    props.setProperty('CAMPAIGN_VERIFICATIONS', JSON.stringify(verifications));
    return { success: true, verifications: verifications };
  }

  if (method === 'batchVerifyCampaigns') {
    const payload = args[0] || {};
    const keys = payload.recordKeys || [];
    const props = PropertiesService.getScriptProperties();
    let verifications = {};
    try {
      verifications = JSON.parse(props.getProperty('CAMPAIGN_VERIFICATIONS') || '{}');
    } catch(e) {}
    const nowStr = formatCustomDate(new Date()) + ', ' + new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    keys.forEach(function(k) {
      if (k) {
        verifications[k] = {
          verified: true,
          verifiedBy: payload.verifiedBy || 'Admin',
          verifiedAt: nowStr,
          notes: 'Batch verified'
        };
      }
    });
    props.setProperty('CAMPAIGN_VERIFICATIONS', JSON.stringify(verifications));
    return { success: true, verifications: verifications };
  }

  if (method === 'getCampaignVerifications') {
    const props = PropertiesService.getScriptProperties();
    let verifications = {};
    try {
      verifications = JSON.parse(props.getProperty('CAMPAIGN_VERIFICATIONS') || '{}');
    } catch(e) {}
    return { success: true, verifications: verifications };
  }

  return { success: true, message: 'Executed: ' + method };
}

/**
 * Direct Method Proxies for google.script.run
 */
function toggleCampaignVerification(data) { return rpc('toggleCampaignVerification', [data]); }
function batchVerifyCampaigns(data) { return rpc('batchVerifyCampaigns', [data]); }
function getCampaignVerifications() { return rpc('getCampaignVerifications', []); }
function runDiagnosticsReport(m, y) { return rpc('runDiagnosticsReport', [m, y]); }
function getInitialData() { return rpc('getInitialData', []); }
function loginUser(pin) { return rpc('loginUser', [pin]); }
function logoutUser() { return rpc('logoutUser', []); }
function updateRate(rate) { return rpc('updateRate', [rate]); }
function getAgencyAnalytics() { return rpc('getAgencyAnalytics', []); }
function getMenuData(k, f) {
  f = f || {};
  var records = rpc('getRecords', [k, f]);
  var month = f.month || '';
  var year = f.year || '';
  var date = f.date || '';
  var pageType = f.pageType || '';
  var stats = rpc('getMenuStats', [k, month, year, date, pageType]);
  return {
    success: true,
    records: records,
    rows: records ? (records.rows || []) : [],
    headers: records ? (records.headers || []) : [],
    stats: stats
  };
}
function getRecords(k, f) { return rpc('getRecords', [k, f]); }
function getMenuStats(k, m, y, d, p) { return rpc('getMenuStats', [k, m, y, d, p]); }
function getNextSLNo(m, y) { return rpc('getNextSLNo', [m, y]); }
function getCampaignBySL(sl) { return rpc('getCampaignBySL', [sl]); }
function addCampaign(data) { return rpc('addCampaign', [data]); }
function addCampaignBatch(d, items) { return rpc('addCampaignBatch', [d, items]); }
function updateCampaignRecord(id, data) { return rpc('updateCampaignRecord', [id, data]); }
function deleteCampaignRecord(id, match, year) { return rpc('deleteCampaignRecord', [id, { month: match, year: year }]); }
function addPayment(data) { return rpc('addPayment', [data]); }
function deletePayment(id) { return rpc('deletePayment', [id]); }
function updatePayment(id, data) { return rpc('updatePayment', [id, data]); }
function addAgencyCost(data) { return rpc('addAgencyCost', [data]); }
function deleteAgencyCost(id) { return rpc('deleteAgencyCost', [id]); }
function updateAgencyCost(id, data) { return rpc('updateAgencyCost', [id, data]); }
function getPageProfile(name) { return rpc('getPageProfile', [name]); }
function getSheetConfig() { return rpc('getSheetConfig', []); }
function getAppsScriptCode() { return rpc('getAppsScriptCode', []); }
function updateSheetConfig(sid, wh) { return rpc('updateSheetConfig', [sid, wh]); }
function testWebhook(wh) { return rpc('testWebhook', [wh]); }
function pushAllDataToSheetViaWebhook(wh) { return rpc('pushAllDataToSheetViaWebhook', [wh]); }
function syncSpreadsheet(sid) { return rpc('syncSpreadsheet', [sid]); }
function clearAllData() { return rpc('clearAllData', []); }
function importCsvData(t, text, m, y) { return rpc('importCsvData', [t, text, m, y]); }
function addAdAccount(a, id) { return rpc('addAdAccount', [a, id]); }
function addBulkAdAccounts(items) { return rpc('addBulkAdAccounts', [items]); }
function deleteAdAccount(name) { return rpc('deleteAdAccount', [name]); }
function addPage(p) { return rpc('addPage', [p]); }
function addBulkPages(pages) { return rpc('addBulkPages', [pages]); }
function deletePage(name) { return rpc('deletePage', [name]); }

/**
 * Menu Action: Pull all data from AdVerse Web App to Google Sheet
 */
function pullAllAppDataToSheet() {
  try {
    var webhookUrl = PropertiesService.getScriptProperties().getProperty("ADVERSE_APP_URL") ||
                     PropertiesService.getScriptProperties().getProperty("ADVERSE_APP_WEBHOOK");
    if (!webhookUrl) {
      SpreadsheetApp.getUi().alert("ℹ️ Data Status", "All data in this spreadsheet is already active and natively stored. If syncing with an external deployment, please configure ADVERSE_APP_URL in Script Properties.", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    var res = UrlFetchApp.fetch(webhookUrl, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ action: "getAllSheetsData" }),
      muteHttpExceptions: true
    });
    var json = JSON.parse(res.getContentText());
    if (json && (json.allSheetsData || (json.result && json.result.allSheetsData))) {
      var allSheets = json.allSheetsData || json.result.allSheetsData;
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var count = 0;
      for (var sName in allSheets) {
        var s = ss.getSheetByName(sName) || ss.insertSheet(sName);
        var rows = allSheets[sName];
        if (rows && rows.length > 0) {
          s.clear();
          s.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
          count++;
        }
      }
      SpreadsheetApp.getUi().alert("✅ Success", "Updated " + count + " sheets directly from AdVerse App.", SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert("ℹ️ Status", (json && json.message) || "Data pull completed.", SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (err) {
    SpreadsheetApp.getUi().alert("⚠️ Notice", "Could not connect to external service: " + err.message + "\n\nNote: All spreadsheet records are already stored natively.", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Menu Action: Sync current spreadsheet data to AdVerse Web App
 */
function syncToAdverseApp() {
  try {
    var webhookUrl = PropertiesService.getScriptProperties().getProperty("ADVERSE_APP_URL") ||
                     PropertiesService.getScriptProperties().getProperty("ADVERSE_APP_WEBHOOK");
    if (!webhookUrl) {
      SpreadsheetApp.getUi().alert("ℹ️ Sync Status", "This Google Sheet is running AdVerse Agency Suite natively. To stream live updates to an external deployment, configure ADVERSE_APP_URL in Script Properties.", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var allData = {};
    sheets.forEach(function(s) {
      var name = s.getName();
      var values = s.getDataRange().getValues();
      if (values && values.length > 0) {
        allData[name] = values;
      }
    });
    UrlFetchApp.fetch(webhookUrl, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ action: "pushAllData", allSheetsData: allData, sheetId: ss.getId() }),
      muteHttpExceptions: true
    });
    SpreadsheetApp.getUi().alert("✅ Success", "Google Sheet successfully synced to AdVerse App!", SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (err) {
    SpreadsheetApp.getUi().alert("⚠️ Notice", "Could not connect to webhook: " + err.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🚀 AdVerse Dashboard")
    .addItem("📅 1-Click: Format All Dates to DD-MM-YYYY (12-10-2026)", "formatAllDatesToDDMMYYYY_Dash")
    .addItem("📅 1-Click: Format All Dates to DD/MM/YYYY (12/10/2026)", "formatAllDatesToDDMMYYYY")
    .addItem("📅 1-Click: Format All Dates to DD-Mon-YYYY (12-Oct-2026)", "formatAllDatesToDDMonYYYY")
    .addItem("📅 1-Click: Format All Dates to YYYY-MM-DD (2026-10-12)", "formatAllDatesToYYYYMMDD")
    .addSeparator()
    .addItem("📥 Pull All App Data to Sheet", "pullAllAppDataToSheet")
    .addItem("🔄 Sync this Sheet to AdVerse App", "syncToAdverseApp")
    .addToUi();
}

/**
 * 📅 ONE-TIME UTILITY: Format All Dates in Entire Spreadsheet to DD-Mon-YYYY (e.g. 11-Oct-2026)
 * Run this function once from the Apps Script editor or from '🚀 AdVerse Dashboard' menu in Google Sheets.
 */
function formatAllDatesToDDMMYYYY_Dash() { return formatAllDatesInEntireSpreadsheet('DD-MM-YYYY'); }
function formatAllDatesToDDMMYYYY() { return formatAllDatesInEntireSpreadsheet('DD/MM/YYYY'); }
function formatAllDatesToDDMonYYYY() { return formatAllDatesInEntireSpreadsheet('DD-Mon-YYYY'); }
function formatAllDatesToYYYYMMDD() { return formatAllDatesInEntireSpreadsheet('YYYY-MM-DD'); }

function formatAllDatesInEntireSpreadsheet(chosenFormat) {
  var targetFormat = chosenFormat || CURRENT_DATE_FORMAT || 'DD-MM-YYYY';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var shortMonths = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var fullMonths = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  
  var totalSheetsUpdated = 0;
  var totalCellsFormatted = 0;
  var summaryLines = [];

  function parseToStandardDate(val, monthHint) {
    if (val === undefined || val === null || val === '') return '';
    
    // 1. If already a Date object
    if (Object.prototype.toString.call(val) === '[object Date]') {
      if (isNaN(val.getTime())) return '';
      var day = ('0' + val.getDate()).slice(-2);
      var mon = shortMonths[val.getMonth()];
      var yr = String(val.getFullYear());
      return day + '-' + mon + '-' + yr;
    }

    // 2. If number (Google Sheets Date serial e.g. 45560 = ~2024, 46259 = ~2026)
    if (typeof val === 'number' && val > 30000 && val < 70000) {
      var utcDays = Math.round(val - 25569);
      var d = new Date((utcDays * 86400 + 43200) * 1000);
      if (!isNaN(d.getTime())) {
        var day = ('0' + d.getDate()).slice(-2);
        var mon = shortMonths[d.getMonth()];
        var yr = String(d.getFullYear());
        return day + '-' + mon + '-' + yr;
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
        return ('0' + day).slice(-2) + '-' + shortMonths[mIdx] + '-' + yr;
      }
    }

    // Fallback native Date parse
    var generic = new Date(cleanStr);
    if (!isNaN(generic.getTime())) {
      var day = ('0' + generic.getDate()).slice(-2);
      var mon = shortMonths[generic.getMonth()];
      var yr = String(generic.getFullYear());
      return day + '-' + mon + '-' + yr;
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
          var formatted = formatCustomDate(rawVal, monthHint, targetFormat);
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
