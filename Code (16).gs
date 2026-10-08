/*******************************************************
 * JIWIN APPARELS — FABRIC ERP (v3 full rebuild)
 * Style Tracker is the single master. Process rates per JO drive
 * which processes are allowed at Process DC. PO print/PDF is gated
 * on approval. Bill Passing -> Bill Approval -> Weight/Loss check
 * -> Debit Note on excess loss. JO Details shows a live timeline.
 *******************************************************/

const SS = SpreadsheetApp.getActiveSpreadsheet();

const SHEETS = {
  STYLE_TRACKER: 'Style Tracker',
  ACCESSORIES_BOM: 'Accessories BOM',
  ACCESSORIES_PO: 'Accessories PO',
  ACCESSORIES_RECEIPT: 'Accessories Receipt',
  ACCESSORIES_MANUAL_STOCK: 'Accessories Manual Stock',
  PROCESS_RATES: 'JO Process Rates',
  PROCESS_MASTER: 'Process Master',
  SUPPLIERS: 'Supplier Master',
  YARN_PO: 'Yarn PO',
  YARN_INWARD: 'Yarn Inward',
  PROCESS_DC: 'Process DC',
  JO_TRANSFER: 'JO Transfer',
  BILL_PASSING: 'Bill Passing',
  WEIGHT_LOSS: 'Weight Loss',
  DEBIT_NOTE: 'Debit Note',
  JO_STATUS: 'JO Status',
  ACCESSORIES_BILL: 'Accessories Bill',
  YARN_BILL: 'Yarn Bill Passing',
  CMT_BILL: 'CMT Bill Passing',
  CMT_RATES: 'CMT Rates',
  GENERAL_BILL: 'General Bill Passing'
};

const HEADERS = {
  'Style Tracker': ['JO Number', 'Style Number', 'Buyer Name', 'Style Description', 'Part', 'Color',
    'Order Quantity', 'Excess %', 'No Of Parts', 'Excess Quantity',
    'Size 1', 'Qty 1', 'Size 2', 'Qty 2', 'Size 3', 'Qty 3', 'Size 4', 'Qty 4', 'Size 5', 'Qty 5',
    'Size 6', 'Qty 6', 'Size 7', 'Qty 7', 'Size 8', 'Qty 8', 'Size 9', 'Qty 9', 'Size 10', 'Qty 10',
    'Sizewise Total', 'Size Check',
    'CAD Weight (g)', 'Loss %', 'Fabric Required (kg)', 'Counts', 'Yarn Type', 'Content', 'Has Lycra', 'Lycra %',
    'Main Fiber Reqd (kg)', 'Lycra Reqd (kg)',
    'Fabric type', 'GSM',
    'Knitting Dia', 'Finishing dia', 'Loop length', 'Gauge', 'Yarn color', 'Yarn rate',
    'Yarn dyeing rate', 'Yarn dyeing loss %', 'Cone winding rate', 'Cone winding loss %',
    'Knitting rate', 'Knitting loss %', 'Heatsetting rate', 'Heatsetting loss %',
    'Dyeing rate', 'Dyeing loss %', 'Tubular compacting rate', 'Tubular compacting loss %',
    'Stener rate', 'Stenter loss %', 'Compacting rate', 'Compacting loss %',
    'Rasing /Peach fnish', 'Raising loss %', 'Bio washing rate', 'Bio washing loss %',
    'Jiwin Inward loss %',
    'Order/OCN No', 'Buyer Order No', 'Despatch Date', 'Season', 'Merchandiser', 'Target Date'],
  'Accessories BOM': ['JO Number', 'Style Number', 'Buyer', 'Color', 'Trims / Accessories', 'Size',
    'Qty Needed', 'Per How Many Garments', 'Rate (Rs/unit)', 'I/H Needed Date'],
  'Accessories PO': ['PO No', 'PO Date', 'JO Number', 'Style Number', 'Trims / Accessories', 'Size',
    'Required Qty', 'Qty Ordered', 'Rate (Rs/unit)', 'PO Value', 'Supplier', 'GST No', 'GST %',
    'Status', 'Approved By', 'Approval Date', 'Rejection Reason', 'Purchase Type', 'Quotation URL', 'General Description', 'PO Group'],
  'Accessories Receipt': ['Receipt No', 'Date', 'Accessories PO No', 'JO Number', 'Trims / Accessories',
    'Supplier', 'Ordered Qty', 'Received Qty', 'Invoice No', 'Remarks'],
  /* A simple manual ledger, separate from the PO-based Accessories Receipt above — for stock that's
     physically in the store but was never tracked through a PO (an opening balance, old stock being
     entered for the first time). "Type" is either Receipt (adds to stock) or Issue (uses it up against
     a JO). Balance per item = every Receipt qty minus every Issue qty. */
  'Accessories Manual Stock': ['Entry No', 'Date', 'Type', 'Item', 'JO Number', 'Qty', 'Rate', 'Remarks', 'Entered By'],
  'JO Process Rates': ['JO Number', 'Process Name', 'Rate (Rs/kg)'],
  'Process Master': ['Process Name'],
  'Supplier Master': ['Supplier Name', 'Address', 'GST', 'Active', 'Mill Name', 'Delivery Address', 'WhatsApp', 'Contact Number', 'Process', 'Email'],
  'Yarn PO': ['PO No', 'PO Date', 'PO Type', 'JO Number', 'Style Number', 'Buyer', 'Color', 'Counts',
    'Yarn Color', 'Rate (Rs/kg)', 'Qty KG', 'PO Value', 'Supplier', 'Delivery Address', 'GST No',
    'Status', 'Approved By', 'Approval Date', 'Rejection Reason', 'Bag Weight (KG)', 'No of Bags',
    'Content', 'GSM', 'GST %', 'Knitting Unit Phone', 'Prepared By', 'Payment Terms', 'Delivery Date',
    'Knitting Unit Name', 'Knitting Unit GSTIN', 'Supplier Address', 'Knitting Dia'],
  'Yarn Inward': ['Inward No', 'Date', 'Yarn PO No', 'JO Number', 'Mill Name', 'Mill DC/Invoice No',
    'Received Bags', 'Received KG', 'Remarks', 'Counts', 'Yarn Type', 'Content', 'GSM', 'Knitting Unit'],
  'Process DC': ['DC No', 'Date', 'JO Number', 'Style Number', 'Color', 'From Process', 'From Supplier',
    'To Process', 'To Supplier', 'Lot No', 'Previous DC No', 'Previous DC Qty KG', 'Available KG',
    'DC Issue Qty KG', 'Rolls', 'Balance KG', 'Remarks', 'Status', 'Approved By', 'Approval Date',
    'Counts', 'Content', 'GSM', 'Loss KG', 'Loss %', 'Vehicle No', 'Balance Status', 'Balance Note',
    'Supplier DC No'],
  'JO Transfer': ['Transfer No', 'Date', 'Source JO', 'Destination JO', 'Qty KG', 'Remarks', 'Status', 'Counts', 'Stock Type', 'Content', 'GSM'],
  'Bill Passing': ['Bill No', 'Date', 'Supplier', 'JO Number', 'DC Numbers', 'DC Qty KG',
    'Applicable Rate', 'Expected Amount', 'Bill Amount', 'Variance', 'Remarks', 'Status',
    'Approved By', 'Approval Date', 'Rejection Reason', 'Bill Photo URL',
    'GST %', 'GST Amount', 'Total Incl GST', 'Bill Stated Total', 'GST Match', 'Supplier Invoice No'],
  'Weight Loss': ['WL No', 'Date', 'JO Number', 'DC No', 'Bill No', 'Previous Qty KG', 'Available Qty KG',
    'Loss KG', 'Loss %', 'Allowed Loss %', 'Result', 'Debit Note No'],
  'Debit Note': ['Debit Note No', 'Date', 'JO Number', 'Supplier', 'DC No', 'Bill No', 'Excess Loss KG',
    'Rate', 'Amount', 'Status'],
  'JO Status': ['JO Number', 'Status', 'Closed Date', 'Closed By', 'Remarks'],
  'Accessories Bill': ['Bill No', 'Date', 'Supplier', 'JO Number', 'Accessories PO No', 'Trims / Accessories',
    'Ordered Qty', 'Applicable Rate', 'Expected Amount', 'Bill Amount', 'Variance',
    'GST %', 'GST Amount', 'Total Incl GST', 'Bill Stated Total', 'GST Match',
    'Remarks', 'Status', 'Approved By', 'Approval Date', 'Rejection Reason', 'Supplier Invoice No'],
  'Yarn Bill Passing': ['Bill No', 'Date', 'Supplier', 'JO Number', 'PO Numbers', 'PO Qty KG',
    'Applicable Rate', 'Expected Amount', 'Bill Amount', 'Variance', 'Remarks', 'Status',
    'Approved By', 'Approval Date', 'Rejection Reason', 'Bill Photo URL',
    'GST %', 'GST Amount', 'Total Incl GST', 'Bill Stated Total', 'GST Match', 'Supplier Invoice No'],
  'CMT Bill Passing': ['Bill No', 'Date', 'Department', 'Contractor', 'JO Number', 'Style Number', 'Color',
    'Qty', 'Rate Per Pc', 'Amount', 'Remarks', 'Status', 'Approved By', 'Approval Date', 'Rejection Reason',
    'Bill Photo URL', 'Bill Group'],
  'CMT Rates': ['JO Number', 'Department', 'Costing Rate Per Pc', 'Rate Per Pc', 'Updated By', 'Updated At'],
  'General Bill Passing': ['Bill No', 'Date', 'Category', 'Paid To', 'JO Number', 'Description', 'Amount',
    'GST %', 'GST Amount', 'Total Incl GST', 'Remarks', 'Status', 'Approved By', 'Approval Date',
    'Rejection Reason', 'Bill Photo URL']
};

const DEFAULT_PROCESSES = ['Knitting', 'OW Heatsetting', 'Dyeing (Bio)', 'Dyeing (Non Bio)', 'Stenter', 'Compacting', 'Other'];

// The two special "process" labels outside Style Tracker's rated processes:
// Yarn Store = where a chain starts (yarn drawn from stock, issued to Knitting).
// Jiwin Inward = where a chain ends (finished fabric returned to Jiwin, no rate needed).
const YARN_STORE_LABEL = 'Yarn Store';
const CORA_STORE_LABEL = 'Cora Fabric Store';
const DYED_STORE_LABEL = 'Dyed Fabric Store';
const JIWIN_INWARD_LABEL = 'Jiwin Inward';

// Where a Process DC chain can originate: which stock type it draws from,
// whether that stock is color-specific, and whether GSM identifies it
// (GSM applies once knitted — Cora/Dyed Fabric — not to raw Yarn).
const CHAIN_ORIGINS = [
  { sourceType: 'YARN', stockType: 'Yarn', fromProcess: YARN_STORE_LABEL, colorSpecific: false, useGsm: false },
  { sourceType: 'CORA', stockType: 'Cora Fabric', fromProcess: CORA_STORE_LABEL, colorSpecific: false, useGsm: true },
  { sourceType: 'DYED', stockType: 'Dyed Fabric', fromProcess: DYED_STORE_LABEL, colorSpecific: true, useGsm: true }
];

/*************** SETUP ***************/
function ensureAllSheets() {
  Object.keys(HEADERS).forEach(function (name) {
    var sh = SS.getSheetByName(name);
    if (!sh) sh = SS.insertSheet(name);
    if (sh.getMaxColumns() < HEADERS[name].length) sh.insertColumnsAfter(sh.getMaxColumns(), HEADERS[name].length - sh.getMaxColumns());
    var headerRange = sh.getRange(1, 1, 1, HEADERS[name].length);
    var current = headerRange.getValues()[0];
    if (sh.getLastRow() === 0 || current.join('') === '') {
      headerRange.setValues([HEADERS[name]]);
      headerRange.setFontWeight('bold').setBackground('#1B2A4A').setFontColor('#FFFFFF');
      sh.setFrozenRows(1);
    } else {
      // Label any header cell left blank (columns added to the code later), never rename existing ones.
      HEADERS[name].forEach(function (h, i) {
        if (String(current[i] || '').trim() === '') {
          sh.getRange(1, i + 1).setValue(h).setFontWeight('bold').setBackground('#1B2A4A').setFontColor('#FFFFFF');
        }
      });
    }
  });
  var pm = SS.getSheetByName(SHEETS.PROCESS_MASTER);
  if (pm.getLastRow() < 2) {
    DEFAULT_PROCESSES.forEach(function (p) { pm.appendRow([p]); });
  }
}

function onOpen() {
  ensureAllSheets();
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Jiwin ERP').addItem('Open Management Portal', 'showPortalDialog')
    .addItem('Setup / Repair Sheets', 'ensureAllSheets').addToUi();
  ui.createMenu('Jiwin OIH').addItem('1. Migrate OIH data (one time)', 'migrateOihData')
    .addItem('2. Set management PIN', 'setManagementPin')
    .addItem('3. Authorize Drive & Mail', 'authorizeServices')
    .addItem('4. Daily reminders ON (9 AM)', 'remindersOn')
    .addItem('5. Daily reminders OFF', 'remindersOff')
    .addItem('6. Send reminders now (test)', 'sendRemindersNow')
    .addItem('7. Set Anthropic API key (for Import PO)', 'setAnthropicApiKey')
    .addToUi();
}

function doGet(e) {
  repairForeignSheets_();   // e.g. a template "Yarn PO" tab with different columns
  ensureAllSheets();
  setupCoreSheets_(); // CONFIG + TECHPACKS only. ORDERS / PRICING / IMAGES / REQS are merged into Style Tracker and no longer created.
  var out = HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Jiwin Fabric ERP')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover')
    .addMetaTag('mobile-web-app-capable', 'yes')
    .addMetaTag('apple-mobile-web-app-capable', 'yes');
  var icon = PropertiesService.getScriptProperties().getProperty('APP_ICON_URL');
  if (icon) { try { out.setFaviconUrl(icon); } catch (err) { } }
  return out;
}

/* Run ONCE from the Apps Script editor: saves the Jiwin logo to Drive and uses it as the app icon
   (browser tab, desktop shortcut, phone home-screen icon). */
function setAppIcon() {
  var b64 = JIWIN_LOGO_PNG.split(',')[1];
  var file = DriveApp.createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/png', 'Jiwin_App_Icon.png'));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var url = 'https://drive.google.com/uc?export=view&id=' + file.getId() + '#.png';
  PropertiesService.getScriptProperties().setProperty('APP_ICON_URL', url);
  Logger.log('App icon set: ' + url);
  return url;
}

function showPortalDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Index').setWidth(1300).setHeight(860);
  SpreadsheetApp.getUi().showModalDialog(html, 'Jiwin Apparels - Fabric ERP');
}

function getArchiveFolder() {
  const folderName = 'JIWIN_ERP_PDFs';
  const folders = DriveApp.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
}

function getBillPhotoFolder() {
  const folderName = 'JIWIN_Bill_Photos';
  const folders = DriveApp.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
}

/**
 * Attaches a photo of the physical bill/invoice to a Bill Passing record.
 * Expects the client to have already compressed the image (kept "low MB")
 * before sending — this just stores whatever base64 it's given.
 */
function uploadBillPhoto(billNo, fileName, base64Data, mimeType) {
  try {
    if (!billNo) return { success: false, error: 'Bill No is required.' };
    var rowIdx = findRowIndex(SHEETS.BILL_PASSING, 'Bill No', billNo);
    if (rowIdx === -1) return { success: false, error: 'Bill ' + billNo + ' not found.' };

    var folder = getBillPhotoFolder();
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType || 'image/jpeg', fileName || (billNo + '.jpg'));
    var file = folder.createFile(blob);
    var url = file.getUrl();
    setCell(SHEETS.BILL_PASSING, rowIdx, 'Bill Photo URL', url);
    return { success: true, url: url, message: 'Bill photo attached to ' + billNo + '.' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/*************** HELPERS ***************/
function formatDateVal(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(val).split('T')[0];
}

/* Repeated sheetToObjects() calls for the same sheet inside ONE read-only report (stock lists, pending-yarn
   reports...) are served from memory instead of going back to Sheets every time. This is switched on only
   by withSheetCache_() around a function that never writes — everything else (approvals, saves, receipts)
   always reads the live sheet, so nothing that edits a sheet can ever read back a stale copy of it. Only
   the raw values are kept; every call still gets its own fresh objects. */
var _sheetValuesCache_ = {};
var _sheetCacheOn_ = 0;
function withSheetCache_(fn) {
  _sheetCacheOn_++;
  try { return fn(); }
  finally { _sheetCacheOn_--; if (_sheetCacheOn_ <= 0) { _sheetCacheOn_ = 0; _sheetValuesCache_ = {}; } }
}
function sheetToObjects(sheetName) {
  var headers = HEADERS[sheetName];
  var raw = _sheetCacheOn_ > 0 ? _sheetValuesCache_[sheetName] : null;
  if (!raw) {
    var sh = SS.getSheetByName(sheetName);
    var lastRow = sh.getLastRow();
    raw = lastRow < 2 ? [] : sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
    if (_sheetCacheOn_ > 0) _sheetValuesCache_[sheetName] = raw;
  }
  return raw.map(function (row, i) {
    var obj = { _row: i + 2 };
    headers.forEach(function (h, j) {
      obj[h] = (row[j] instanceof Date) ? formatDateVal(row[j]) : row[j];
    });
    return obj;
  });
}

/**
 * Reads a sheet by its ACTUAL row-1 header labels rather than the static
 * HEADERS array — so columns the user adds directly in the sheet (new
 * processes, extra fields) show up automatically without a code change,
 * as long as they're appended after the existing columns (never inserted
 * mid-sheet, which would silently misalign every column after it).
 */
/* A picture pasted INTO a Google Sheets cell comes back from getValues() as a CellImage object (its text form is "CellImage").
   Keep it out of anything sent to the browser; the app reads the real picture separately (stCellImageData_). */
function isCellImage_(v) { return v !== null && typeof v === 'object' && !(v instanceof Date) && typeof v.getContentUrl === 'function'; }
function stPlainCell_(v) { return isCellImage_(v) ? 'CellImage' : v; }

function sheetToObjectsDynamic(sheetName) {
  var sh = SS.getSheetByName(sheetName);
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  var headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h || '').trim(); });
  var data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return data.map(function (row, i) {
    var obj = { _row: i + 2 };
    headerRow.forEach(function (h, j) {
      if (!h) return;
      obj[h] = (row[j] instanceof Date) ? formatDateVal(row[j]) : stPlainCell_(row[j]);
    });
    return obj;
  });
}

/**
 * Process rate/loss/DC-code maps, starting from the known legacy overrides
 * (columns whose names don't follow the plain "<Process> rate" pattern —
 * Stenter/Rasing.../etc) and auto-discovering any further "<Name> rate"
 * (+ optional "<Name> loss %") column pair the user has appended directly
 * to Style Tracker. Newly discovered processes are appended to the end of
 * the sequence, as their columns are appended to the end of the sheet.
 */
/* One-time tidy-up: moves the "Process 6".."Process 10" rate/loss % column pairs (wherever they
   currently sit) to right after the Raising / Bio Washing columns, so every process column lives
   together in one place instead of being scattered near the order/despatch columns. Purely a column
   REORDER — no data is changed, each column keeps its header and every row's values move with it.
   Safe to run more than once; does nothing if the columns are already in place or don't exist. */
function reorderProcessColumns() {
  var sh = SS.getSheetByName(SHEETS.STYLE_TRACKER);
  var lastCol = sh.getLastColumn();
  var head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h || '').trim(); });

  var anchor = head.indexOf('Bio washing loss %');
  if (anchor < 0) anchor = head.indexOf('Raising loss %');
  if (anchor < 0) return { success: false, error: 'Could not find a "Raising loss %" or "Bio washing loss %" column to anchor on.' };

  var moveNames = [];
  for (var n = 6; n <= 10; n++) { moveNames.push('Process ' + n + ' rate'); moveNames.push('Process ' + n + ' loss %'); }
  var moveCols = moveNames.map(function (nm) { return head.indexOf(nm); }).filter(function (i) { return i >= 0; });
  if (!moveCols.length) return { success: false, error: 'No "Process 6".."Process 10" columns found.' };
  moveCols.sort(function (a, b) { return a - b; });

  // Already contiguous and already sitting right after the anchor — nothing to do.
  var alreadyContiguous = moveCols.every(function (c, i) { return i === 0 || c === moveCols[i - 1] + 1; });
  if (alreadyContiguous && moveCols[0] === anchor + 1) return { success: true, message: 'Already in place — nothing to move.' };

  // moveColumns needs one contiguous block; if these columns are scattered, bring them together
  // first (right after the anchor), then they're already in final position.
  var destination1based = anchor + 2;   // Sheet.moveColumns destination is 1-based and is WHERE THE BLOCK ENDS UP STARTING
  if (alreadyContiguous) {
    var startCol1based = moveCols[0] + 1;
    sh.moveColumns(sh.getRange(1, startCol1based, sh.getMaxRows(), moveCols.length), destination1based);
  } else {
    // Move one at a time, in order, each right after the anchor — re-reading positions as columns shift.
    moveNames.forEach(function (nm) {
      var curHead = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h || '').trim(); });
      var idx = curHead.indexOf(nm);
      var anchorNow = curHead.indexOf('Bio washing loss %'); if (anchorNow < 0) anchorNow = curHead.indexOf('Raising loss %');
      if (idx < 0 || anchorNow < 0) return;
      sh.moveColumns(sh.getRange(1, idx + 1, sh.getMaxRows(), 1), anchorNow + 2);
    });
  }
  return { success: true, message: 'Moved Process 6–10 to sit right after ' + head[anchor] + '.' };
}

function getProcessMaps() {
  var rateMap = {}, lossMap = {}, codeMap = {}, sequence = [];
  Object.keys(RATE_COLUMN_MAP).forEach(function (p) { rateMap[p] = RATE_COLUMN_MAP[p]; sequence.push(p); });
  Object.keys(LOSS_COLUMN_MAP).forEach(function (p) { lossMap[p] = LOSS_COLUMN_MAP[p]; });
  Object.keys(PROCESS_CODE_MAP).forEach(function (p) { codeMap[p] = PROCESS_CODE_MAP[p]; });
  // Move known processes into the canonical order first, then append anything new after.
  sequence = PROCESS_SEQUENCE.slice();

  var sh = SS.getSheetByName(SHEETS.STYLE_TRACKER);
  var lastCol = sh.getLastColumn();
  var headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h || '').trim(); });
  var knownRateCols = {};
  Object.keys(rateMap).forEach(function (p) { knownRateCols[rateMap[p].toLowerCase()] = true; });

  // Keep a short all-caps word as typed (an acronym like "AOP", "OE"); title-case ordinary words.
  var titleCase_ = function (base) {
    return base.replace(/\S+/g, function (w) {
      if (w.length > 1 && w === w.toUpperCase()) return w;
      return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
    });
  };
  var usedCols = {}; Object.keys(rateMap).forEach(function (p) { usedCols[rateMap[p].toLowerCase()] = true; });
  Object.keys(lossMap).forEach(function (p) { usedCols[lossMap[p].toLowerCase()] = true; });
  var claimProcess_ = function (processName, rateCol, lossCol) {
    if (rateMap[processName]) return; // already discovered (shouldn't happen, but stay safe)
    rateMap[processName] = rateCol; usedCols[rateCol.toLowerCase()] = true;
    if (lossCol) { lossMap[processName] = lossCol; usedCols[lossCol.toLowerCase()] = true; }
    if (!codeMap[processName]) {
      var initials = processName.split(/\s+/).map(function (w) { return w.charAt(0).toUpperCase(); }).join('').slice(0, 3) || 'PX';
      codeMap[processName] = initials;
    }
    sequence.push(processName);
  };

  // Pass 1 — a column named "<Name> rate" (the normal, expected shape).
  headerRow.forEach(function (col) {
    if (!col) return;
    var m = /^(.*?)\s+rate$/i.exec(col);
    if (!m) return;
    if (knownRateCols[col.toLowerCase()]) return;
    var base = m[1].trim();
    var processName = titleCase_(base);
    var lossCol = headerRow.filter(function (c) { return c.toLowerCase() === (base + ' loss %').toLowerCase(); })[0];
    claimProcess_(processName, col, lossCol);
  });

  // Pass 2 — a column that's been renamed to something that doesn't end in " rate" (e.g. "Raising/Peach
  // Finish" retitled to just "Washing") is still found this way: whatever sits immediately before its
  // matching "<Name> loss %" column is that process's rate, whatever it's actually called.
  // A BARE "Loss %" with nothing before "loss %" in its own name isn't a process's loss column at all
  // (it's the generic per-garment consumption loss next to CAD Weight) — skip it rather than naming
  // the "process" after whatever happens to sit in the column before it.
  headerRow.forEach(function (col, i) {
    if (!col || !/loss\s*%$/i.test(col) || usedCols[col.toLowerCase()]) return;
    var ownName = col.replace(/\s*loss\s*%$/i, '').trim();
    if (!ownName) return;
    var rateCol = headerRow[i - 1];
    if (!rateCol || usedCols[rateCol.toLowerCase()]) return;
    claimProcess_(titleCase_(ownName), rateCol, col);
  });

  return { rateMap: rateMap, lossMap: lossMap, codeMap: codeMap, sequence: sequence };
}

function nextId(prefix, sheetName, colIndex) {
  var sh = SS.getSheetByName(sheetName);
  var lastRow = sh.getLastRow();
  var max = 0;
  if (lastRow >= 2) {
    var vals = sh.getRange(2, colIndex, lastRow - 1, 1).getValues();
    vals.forEach(function (r) {
      var m = String(r[0]).match(/(\d+)$/);
      if (m) max = Math.max(max, parseInt(m[1], 10));
    });
  }
  return prefix + Utilities.formatString('%04d', max + 1);
}

/**
 * Like nextId, but scoped to rows that actually start with this exact
 * prefix (e.g. "K-"), so each process keeps its own counter — K-0001,
 * D-0001, ST-0001 — instead of one shared global sequence.
 */
function nextIdForPrefix(sheetName, colIndex, prefix) {
  var sh = SS.getSheetByName(sheetName);
  var lastRow = sh.getLastRow();
  var max = 0;
  var marker = prefix + '-';
  if (lastRow >= 2) {
    var vals = sh.getRange(2, colIndex, lastRow - 1, 1).getValues();
    vals.forEach(function (r) {
      var val = String(r[0]);
      if (val.indexOf(marker) !== 0) return;
      var m = val.match(/(\d+)$/);
      if (m) max = Math.max(max, parseInt(m[1], 10));
    });
  }
  return marker + Utilities.formatString('%04d', max + 1);
}

function findRowIndex(sheetName, idColHeader, idValue) {
  var rows = sheetToObjects(sheetName);
  var row = rows.find(function (r) { return String(r[idColHeader]).trim() === String(idValue).trim(); });
  return row ? row._row : -1;
}

function setCell(sheetName, rowIndex, header, value) {
  var col = HEADERS[sheetName].indexOf(header) + 1;
  SS.getSheetByName(sheetName).getRange(rowIndex, col).setValue(value);
  delete _sheetValuesCache_[sheetName];   // this sheet changed — the next sheetToObjects() call for it must re-read, not use a stale cache
}

/*************** PORTAL BOOTSTRAP ***************/
function getPortalInitialData() {
  ensureAllSheets();
  return {
    suppliers: getSuppliers(),
    styleTracker: getStyleTracker(),
    processMaster: sheetToObjects(SHEETS.PROCESS_MASTER).map(function (r) { return r['Process Name']; }),
    yarnPOs: sheetToObjects(SHEETS.YARN_PO),
    processDCs: sheetToObjects(SHEETS.PROCESS_DC),
    nextPONumber: nextId('YPO-', SHEETS.YARN_PO, 1),
    nextInwardNumber: nextId('YINW-', SHEETS.YARN_INWARD, 1),
    nextDCNumber: nextId('PDC-', SHEETS.PROCESS_DC, 1)
  };
}

/*************** 2. SUPPLIER MASTER ***************/
function getSuppliers() { return sheetToObjects(SHEETS.SUPPLIERS); }

/**
 * Suppliers capable of one process/type — matched against Supplier Master's
 * "Process" field, which can hold several comma-separated capabilities
 * (e.g. "Knitting, Heatsetting"). Feeds the process-filtered supplier
 * dropdowns in Yarn PO (matched against PO Type: Yarn / Cora Fabric / Dyed
 * Fabric) and Process DC (matched against To Process).
 */
var YARN_PO_TOLERANCE_KG = 20;   // max kg a Yarn PO may go above the balance
var SAMPLE_TRIAL_MIN_KG = 30;    // a sample JO's computed requirement is usually far below what a mill
                                  // will actually run as a trial lot — always allow ordering up to this much
                                  // for Yarn / Cora Fabric / Dyed Fabric POs against a SAMPLE order.
var SAMPLE_WT_COL = 'Sample Weight (kg)';   // weight-only sample line inside a normal JO (no pcs / sizes)
function isSampleWtRow_(r) { return Number(r[SAMPLE_WT_COL]) > 0; }
function isSampleJO_(jo) {
  if (/^S\s*-/i.test(jo) || /^SAMPLE/i.test(jo)) return true;
  // A "sample JO" = every real (piece-based) line is a SAMPLE. Weight-only sample lines added to a normal
  // order JO do NOT turn the whole JO into a sample.
  var bulk = getStyleTracker().filter(function (r) { return String(r['JO Number']).trim() === String(jo).trim() && !isSampleWtRow_(r); });
  return bulk.length > 0 && bulk.every(function (r) { return String(r['Order Type']).trim().toUpperCase() === 'SAMPLE'; });
}

function getSuppliersForProcess(processName) {
  var all = getSuppliers().map(function (s) { return s['Supplier Name']; });
  if (!processName) return all;
  var matched = getSuppliers().filter(function (s) {
    // "Yarn" matches "Yarn", "YARN PURCHASE", "Yarn Supplier"; "Knitting" matches "Knitting, Heatsetting"
    var want = String(processName).trim().toLowerCase().split(/\s+/)[0];
    var list = String(s['Process'] || '').split(',').map(function (p) { return p.trim().toLowerCase(); });
    return list.some(function (p) { return p === String(processName).trim().toLowerCase() || (want && p.indexOf(want) !== -1); });
  }).map(function (s) { return s['Supplier Name']; });
  // Fall back to the full supplier list when nobody's been tagged with this
  // process yet — an empty dropdown is worse than an unfiltered one.
  return matched.length ? matched : all;
}

function addSupplier(d) {
  var name = String(d.name || '').trim();
  var address = String(d.address || '').trim();
  var gst = String(d.gst || '').trim();
  var contactNumber = String(d.contactNumber || '').trim();
  var missing = [];
  if (!name) missing.push('Name');
  if (!address) missing.push('Address');
  if (!gst) missing.push('GST');
  if (!contactNumber) missing.push('Contact Number');
  if (missing.length) {
    return { success: false, error: 'Missing required field(s): ' + missing.join(', ') + '.' };
  }
  if (!/^\d{10}$/.test(contactNumber.replace(/\D/g, '')) || contactNumber.replace(/\D/g, '').length !== 10) {
    return { success: false, error: 'Contact Number should be a 10-digit phone number.' };
  }
  if (gst.length !== 15) {
    return { success: false, error: 'GST number should be 15 characters (e.g. 33AAAAA0000A1Z1). Got "' + gst + '".' };
  }
  var existing = getSuppliers();
  var gstDupe = existing.find(function (s) { return String(s['GST'] || '').trim().toUpperCase() === gst.toUpperCase(); });
  if (gstDupe) {
    return { success: false, error: 'This GST number is already used by "' + gstDupe['Supplier Name'] + '".' };
  }
  var addressDupe = existing.find(function (s) { return String(s['Address'] || '').trim().toLowerCase() === address.toLowerCase(); });
  if (addressDupe) {
    return { success: false, error: 'This Address is already used by "' + addressDupe['Supplier Name'] + '".' };
  }
  var contactDupe = existing.find(function (s) { return String(s['Contact Number'] || '').replace(/\D/g, '') === contactNumber.replace(/\D/g, ''); });
  if (contactDupe) {
    return { success: false, error: 'This Contact Number is already used by "' + contactDupe['Supplier Name'] + '".' };
  }
  var emailCheck = parseEmails_(d.email);
  if (emailCheck.bad.length) return { success: false, error: 'This email address does not look right: ' + emailCheck.bad.join(', ') + '.' };
  appendRowSafe_(SHEETS.SUPPLIERS, [
    name, address, gst, d.active || 'Active', d.millName || '', d.deliveryAddress || '',
    d.whatsapp || '', contactNumber, d.process || '', emailCheck.ok.join(', ')
  ]);
  return { success: true, message: name + ' saved.', suppliers: getSuppliers() };
}

/*************** 1. STYLE TRACKER ***************/
function getStyleTracker() { return sheetToObjectsDynamic(SHEETS.STYLE_TRACKER); }

function getStyleTrackerJOs() {
  var sh = SS.getSheetByName(SHEETS.STYLE_TRACKER);
  if (!sh || sh.getLastRow() < 2) return [];
  var values = sh.getDataRange().getValues();
  var headers = values.shift().map(function(h){ return String(h).trim(); });
  var idx = headers.indexOf('JO Number');
  if (idx < 0) return [];
  var seen = {};
  var out = [];
  values.forEach(function(row){
    var jo = String(row[idx] == null ? '' : row[idx]).trim();
    if (jo && !seen[jo]) {
      seen[jo] = true;
      out.push(jo);
    }
  });
  return out;
}

function getNextJONumber() { return nextId('JO-', SHEETS.STYLE_TRACKER, 1); }

// NOTE: Style Tracker rows are entered directly in the Google Sheet, not
// through this web app. addStyleTrackerRow / saveStyleWithSizes were removed
// because they wrote the old single-Size-column layout, which no longer
// matches the sheet (Size 1..10 / Qty 1..10). The portal only reads.

/**
 * Total Cut Qty for a JO (Order Qty + Excess Qty, once per color — never
 * multiplied across Parts) — both the JO-wide total and a per-size
 * breakdown, for Accessories BOM's consumption math.
 */
function getTotalCutQtyForJO(joNumber, color, style) {
  // colour / style matched loosely ("Pink" finds "PINK", "SPG019" finds "SPG019M1");
  // a line with no size breakdown uses its Order Quantity
  var rows = getStyleTracker().filter(function (r) {
    if (String(r['JO Number']).trim() !== String(joNumber).trim()) return false;
    if (color && stLike_(r['Color'], color) !== 2 && stNorm_(r['Color']) !== stNorm_(color)) return false;
    if (style && stLike_(r['Style Number'], style) === 0) return false;
    return true;
  });
  var seen = {}, total = 0, bySize = {};
  rows.forEach(function (r) {
    var key = stNorm_(r['Style Number']) + '|' + stNorm_(r['Color'] || '-');
    if (seen[key]) return;
    seen[key] = true;
    var excessPct = Number(r['Excess %']) || 0, any = false;
    for (var i = 1; i <= 10; i++) {
      var label = String(r['Size ' + i] || '').trim();
      if (!label) continue;
      any = true;
      var cutQty = Math.round((Number(r['Qty ' + i]) || 0) * (1 + excessPct / 100));
      total += cutQty;
      bySize[label] = (bySize[label] || 0) + cutQty;
    }
    if (!any) total += Math.round((Number(r['Order Quantity']) || 0) * (1 + excessPct / 100));
  });
  return { total: total, bySize: bySize };
}

/** Size-by-size cut qty (Order+Excess) for one JO+Color — the preview table for size-based accessories. */
function getSizeBreakdownForJOColor(joNumber, color) {
  var bySize = getTotalCutQtyForJO(joNumber, color).bySize;
  var rows = getStyleTracker().filter(function (r) {
    return String(r['JO Number']).trim() === String(joNumber).trim() && String(r['Color'] || '').trim() === String(color).trim();
  });
  var order = [];
  if (rows.length) {
    for (var i = 1; i <= 10; i++) {
      var label = rows[0]['Size ' + i];
      if (label) order.push(label);
    }
  }
  return order.map(function (label) { return { size: label, cutQty: bySize[label] || 0 }; });
}

/**
 * Accessories BOM for a JO — Required Quantity and Amount are always
 * computed here, never typed in: Required Qty = CEIL(Total Cut Qty / Per
 * How Many Garments) x Qty Needed, rounded up since you can't buy a
 * fraction of a cone/roll. Blank Size uses the whole JO's cut qty; a
 * filled Size uses just that size's own cut qty (for accessories like a
 * size-printed wash care label that varies by size).
 */
function getAccessoriesBOM(joNumber) {
  var rows = sheetToObjects(SHEETS.ACCESSORIES_BOM);
  if (joNumber) {
    rows = rows.filter(function (r) { return String(r['JO Number']).trim() === String(joNumber).trim(); });
  }
  if (!rows.length) return rows;
  var cutQtyCache = {}, extras = accBomExtrasByRow_();
  return rows.map(function (r) {
    var jo = r['JO Number'];
    var color = r['Color'] || '';
    var style = r['Style Number'] || '';
    var cacheKey = jo + '||' + color + '||' + style;
    if (!cutQtyCache[cacheKey]) cutQtyCache[cacheKey] = getTotalCutQtyForJO(jo, color, style);
    var cutQty = r['Size'] ? (cutQtyCache[cacheKey].bySize.hasOwnProperty(r['Size']) ? cutQtyCache[cacheKey].bySize[r['Size']] : cutQtyCache[cacheKey].total) : cutQtyCache[cacheKey].total;
    var perGarments = Number(r['Per How Many Garments']) || 1;
    var qtyNeeded = Number(r['Qty Needed']) || 0;
    var autoQty = Math.ceil(cutQty / perGarments) * qtyNeeded;
    var ex = extras[r._row] || {};
    var planned = Number(ex.planned) > 0 ? Math.min(Number(ex.planned), autoQty) : autoQty;   // never more than the order needs
    var extraQty = Number(ex.extra) || 0, extraOk = ex.status === 'Approved' ? extraQty : 0;   // extra only after Management approval
    var requiredQty = planned + extraOk;
    var rate = Number(r['Rate (Rs/unit)']) || 0;
    return Object.assign({}, r, {
      'Auto Required': autoQty, 'Planned Qty': planned, 'Extra Qty': extraQty, 'Extra Status': ex.status || '', 'Extra Reason': ex.reason || '',
      'Required Quantity': requiredQty,
      'Amount': requiredQty * rate
    });
  });
}

/**
 * Adds one Accessories BOM entry from the app (Wash Care, Fusing Sticker,
 * Main Label, Gum Tape, etc.) — Style Number/Buyer are pulled from Style
 * Tracker via the JO, never re-typed. Required Quantity/Amount are computed
 * on read (getAccessoriesBOM), not stored, so they always reflect the JO's
 * current cut quantity.
 */
/**
 * Adds Accessories BOM entries from the app — Style Number/Buyer are pulled
 * from Style Tracker via the JO, never re-typed. When sizeBased is true, one
 * row is created per size of the chosen Color (each with Qty Needed = 1 —
 * the per-size cut qty itself is what varies), so a single form submission
 * covers the whole size run; otherwise one row is created with the manually
 * entered Quantity Needed and no Size. Required Quantity/Amount are always
 * computed on read (getAccessoriesBOM), never stored.
 */
function saveAccessoryBOM(d) {
  var jo = getJOInfo(d.joNumber);
  if (!jo) return { success: false, error: 'JO "' + d.joNumber + '" not found in Style Tracker.' };
  if (!d.name) return { success: false, error: 'Trims / Accessories name is required.' };
  if (!d.color) return { success: false, error: 'Color is required.' };

  var sheet = SS.getSheetByName(SHEETS.ACCESSORIES_BOM);
  var ihDate = d.ihNeededDate || '';
  var dupe = sheetToObjects(SHEETS.ACCESSORIES_BOM).filter(function (r) {
    return String(r['JO Number']).trim() === String(d.joNumber).trim() &&
      String(r['Color'] || '').trim().toUpperCase() === String(d.color || '').trim().toUpperCase() &&
      String(r['Trims / Accessories'] || '').trim().toUpperCase() === String(d.name || '').trim().toUpperCase();
  })[0];
  if (dupe) return { success: false, error: d.name + ' is already planned for ' + d.joNumber + ' / ' + d.color + '. Change that line instead of adding it twice.' };

  if (d.sizeBased) {
    var sizeRows;
    try { sizeRows = JSON.parse(d.sizeRows || '[]'); } catch (e) { sizeRows = []; }
    if (!sizeRows.length) return { success: false, error: 'No sizes found for ' + d.joNumber + ' / ' + d.color + '.' };
    sizeRows.forEach(function (s) {
      sheet.appendRow([
        d.joNumber, jo['Style Number'] || '', jo['Buyer'] || '', d.color, d.name, s.size,
        1, Number(s.perGarments) || 1, Number(s.rate) || 0, ihDate
      ]);
    });
    return { success: true, message: d.name + ' added for all ' + sizeRows.length + ' sizes of ' + d.joNumber + ' / ' + d.color + '.' };
  }

  var qtyNeeded = Number(d.qtyNeeded) || 0;
  if (qtyNeeded <= 0) return { success: false, error: 'Quantity must be greater than 0.' };
  if (qtyNeeded > 50 && !d.confirmBigQty) return { success: false, error: 'Qty per garment ' + qtyNeeded + ' is too high - it is the number on ONE garment (usually 1). Required Qty is worked out from the order.' };
  appendRowSafe_(SHEETS.ACCESSORIES_BOM, [
    d.joNumber, jo['Style Number'] || '', jo['Buyer'] || '', d.color, d.name, '',
    qtyNeeded, Number(d.perGarments) || 1, Number(d.rate) || 0, ihDate
  ]);
  return { success: true, message: d.name + ' added to Accessories BOM for ' + d.joNumber + ' / ' + d.color + '.' };
}

/** Total accessories budget for a JO, and per-piece (divided by that JO's Total Cut Qty). */
function getAccessoriesBudgetForJO(joNumber) {
  var rows = getAccessoriesBOM(joNumber);
  var totalBudget = rows.reduce(function (s, r) { return s + (Number(r['Amount']) || 0); }, 0);
  var cutQty = getTotalCutQtyForJO(joNumber).total;
  return {
    joNumber: joNumber, itemCount: rows.length, totalCutQty: cutQty,
    budgetPerJO: totalBudget, budgetPerPiece: cutQty > 0 ? (totalBudget / cutQty) : 0
  };
}

/*************** ACCESSORIES PO + RECEIPT ***************/
/**
 * Qty already placed on Accessories PO for one JO+Trim(+Size), excluding
 * Rejected — Pending and Approved both hold against the planned Required
 * Qty, same pattern as Yarn PO's usage check.
 */
function getAccessoryPOUsage(joNumber, trimName, size) {
  return sheetToObjects(SHEETS.ACCESSORIES_PO).filter(function (r) {
    if (String(r['JO Number']).trim() !== String(joNumber).trim()) return false;
    if (r['Status'] === 'Rejected') return false;
    if (String(r['Trims / Accessories'] || '').trim() !== String(trimName).trim()) return false;
    if (String(r['Size'] || '').trim() !== String(size || '').trim()) return false;
    return true;
  }).reduce(function (s, r) { return s + (Number(r['Qty Ordered']) || 0); }, 0);
}

/** Planned Accessories BOM lines for a JO, each with how much is already on PO and the remaining balance — feeds the Accessories PO picker. */
function getAccessoryItemsForJO(joNumber) {
  return getAccessoriesBOM(joNumber).map(function (r) {
    var requiredQty = Number(r['Required Quantity']) || 0;
    var already = getAccessoryPOUsage(joNumber, r['Trims / Accessories'], r['Size']);
    return {
      name: r['Trims / Accessories'], size: r['Size'] || '', color: r['Color'] || '', style: r['Style Number'] || '',
      requiredQty: requiredQty, alreadyOrdered: already, balance: Math.max(0, requiredQty - already),
      rate: Number(r['Rate (Rs/unit)']) || 0
    };
  });
}

function saveAccessoryPO(d) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var jo = getJOInfo(d.joNumber);
    if (!jo) return { success: false, error: 'JO "' + d.joNumber + '" not found in Style Tracker.' };
    var qty = Number(d.qtyOrdered) || 0;
    if (qty <= 0) return { success: false, error: 'Qty Ordered must be greater than 0.' };
    var requiredQty = Number(d.requiredQty) || 0;
    var already = getAccessoryPOUsage(d.joNumber, d.trimName, d.size);
    var balance = requiredQty - already;
    if (qty > balance + 0.5) {
      return {
        success: false,
        error: 'Qty (' + qty + ') exceeds the planned requirement for ' + d.trimName + (d.size ? ' / Size ' + d.size : '') +
          '. Required: ' + requiredQty + ', already on PO: ' + already + ', balance: ' + Math.max(0, balance) + '.'
      };
    }
    var rate = Number(d.rate) || 0;
    var poNo = nextId('APO-', SHEETS.ACCESSORIES_PO, 1);
    appendRowSafe_(SHEETS.ACCESSORIES_PO, [
      poNo, d.poDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      d.joNumber, jo['Style Number'] || '', d.trimName, d.size || '', requiredQty, qty, rate, rate * qty,
      d.supplier || '', d.gstNo || '', Number(d.gstPct) || 0, 'Pending Approval', '', '', ''
    ]);
    return { success: true, poNo: poNo, message: 'Accessories PO ' + poNo + ' submitted — awaiting approval.' };
  } catch (err) {
    return { success: false, error: err.message };
  } finally { lock.releaseLock(); }
}

function approveAccessoryPO(poNo, decision, approver, reason) {
  var rowIdx = findRowIndex(SHEETS.ACCESSORIES_PO, 'PO No', poNo);
  if (rowIdx === -1) return { success: false, error: 'Accessories PO not found: ' + poNo };
  var status = decision === 'APPROVE' ? 'Approved' : 'Rejected';
  setCell(SHEETS.ACCESSORIES_PO, rowIdx, 'Status', status);
  setCell(SHEETS.ACCESSORIES_PO, rowIdx, 'Approved By', approver || '');
  setCell(SHEETS.ACCESSORIES_PO, rowIdx, 'Approval Date', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  setCell(SHEETS.ACCESSORIES_PO, rowIdx, 'Rejection Reason', decision === 'REJECT' ? (reason || '') : '');
  // multi-line PO: every line of the same PO goes with it
  var grp = apoGroupRows_(poNo), extra = 0;
  grp.forEach(function (g) {
    if (g.no === poNo) return;
    setCell(SHEETS.ACCESSORIES_PO, g.row, 'Status', status);
    setCell(SHEETS.ACCESSORIES_PO, g.row, 'Approved By', approver || '');
    setCell(SHEETS.ACCESSORIES_PO, g.row, 'Approval Date', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
    setCell(SHEETS.ACCESSORIES_PO, g.row, 'Rejection Reason', decision === 'REJECT' ? (reason || '') : '');
    extra++;
  });
  var mailNote = (status === 'Approved') ? autoEmailAccessoryPo_(poNo) : '';
  return { success: true, message: 'Accessories PO ' + poNo + (extra ? ' (' + (extra + 1) + ' lines)' : '') + ' ' + status + '.' + (mailNote ? ' ' + mailNote : '') };
}

/** Approved Accessories POs, newest first, one row per PO (a multi-line PO is one row) - feeds the "Approved POs - PDF" list on the Approvals screen. */
function getApprovedAccessoryPOs(limit) {
  var sh = SS.getSheetByName(SHEETS.ACCESSORIES_PO);
  if (!sh || sh.getLastRow() < 2) return [];
  var head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(function (x) { return String(x).trim(); });
  var ix = function (n) { return head.indexOf(n); };
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getValues();
  var tz = Session.getScriptTimeZone(), gi = ix('PO Group');
  var fmt = function (x) { return x instanceof Date ? Utilities.formatDate(x, tz, 'yyyy-MM-dd') : String(x || ''); };
  var map = {}, order = [];
  v.forEach(function (r) {
    if (String(r[ix('Status')]) !== 'Approved') return;
    var key = (gi > -1 && String(r[gi] || '').trim()) || String(r[0]);
    if (!map[key]) { map[key] = { poNo: String(r[0]), poDate: fmt(r[ix('PO Date')]), supplier: String(r[ix('Supplier')] || ''), jos: [], lines: 0, value: 0, approvedBy: String(r[ix('Approved By')] || ''), approvalDate: fmt(r[ix('Approval Date')]) }; order.push(key); }
    var m = map[key], jo = String(r[ix('JO Number')] || '');
    if (jo && m.jos.indexOf(jo) < 0) m.jos.push(jo);
    m.lines++; m.value += Number(r[ix('PO Value')]) || 0;
  });
  return order.reverse().slice(0, Number(limit) || 30).map(function (k) { var m = map[k]; m.value = Math.round(m.value * 100) / 100; return m; });
}

function getAccessoryPOsForJO(joNumber) {
  return sheetToObjects(SHEETS.ACCESSORIES_PO)
    .filter(function (r) { return String(r['JO Number']).trim() === String(joNumber).trim(); })
    .reverse();
}

function saveAccessoryReceipt(d) {
  var po = sheetToObjects(SHEETS.ACCESSORIES_PO).find(function (r) { return String(r['PO No']).trim() === String(d.accessoryPoNo).trim(); });
  if (po && po['Status'] !== 'Approved') {
    return { success: false, error: 'Accessories PO ' + d.accessoryPoNo + ' is not Approved (status: ' + po['Status'] + '). Cannot receive against it.' };
  }
  var receiptNo = nextId('ARCT-', SHEETS.ACCESSORIES_RECEIPT, 1);
  appendRowSafe_(SHEETS.ACCESSORIES_RECEIPT, [
    receiptNo, d.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    d.accessoryPoNo, po ? po['JO Number'] : (d.joNumber || ''), po ? po['Trims / Accessories'] : (d.trimName || ''),
    po ? po['Supplier'] : (d.supplier || ''), po ? Number(po['Qty Ordered']) || 0 : 0,
    Number(d.receivedQty) || 0, d.invoiceNo || '', d.remarks || ''
  ]);
  return { success: true, message: 'Receipt ' + receiptNo + ' recorded.', receiptNo: receiptNo };
}

/*************** ACCESSORIES MANUAL STOCK — a simple Receipt/Issue ledger for stock never tracked
 * through a PO (old stock being entered for the first time, so the store's real balance is known).
 * Deliberately separate from the PO-based Accessories Receipt/Bill flow above — nothing here touches
 * a PO, a supplier bill, or JO budget locking. ***************/
function saveManualAccessoryReceipt(d) {
  var item = String(d.item || '').trim();
  var qty = Number(d.qty) || 0;
  var rate = Number(d.rate) || 0;
  if (!item) return { success: false, error: 'Enter the item name.' };
  if (!(qty > 0)) return { success: false, error: 'Enter a quantity greater than zero.' };
  if (!(rate >= 0)) return { success: false, error: 'Enter a valid rate.' };
  var entryNo = nextId('AMR-', SHEETS.ACCESSORIES_MANUAL_STOCK, 1);
  appendRowSafe_(SHEETS.ACCESSORIES_MANUAL_STOCK, [
    entryNo, d.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    'Receipt', item, d.joNumber || '', qty, rate, d.remarks || '', d.enteredBy || ''
  ]);
  return { success: true, message: entryNo + ': added ' + qty + ' of "' + item + '" to stock.' };
}
function saveManualAccessoryIssue(d) {
  var item = String(d.item || '').trim();
  var qty = Number(d.qty) || 0;
  var jo = String(d.joNumber || '').trim();
  if (!item) return { success: false, error: 'Enter the item name.' };
  if (!(qty > 0)) return { success: false, error: 'Enter a quantity greater than zero.' };
  if (!jo) return { success: false, error: 'Enter the JO Number this is being issued to.' };
  var bal = getAccessoryManualStockBalance().find(function (r) { return r.item.toUpperCase() === item.toUpperCase(); });
  var available = bal ? bal.balance : 0;
  if (qty > available + 0.001) {
    return { success: false, error: 'Only ' + available + ' of "' + item + '" is available in stock — cannot issue ' + qty + '.' };
  }
  var entryNo = nextId('AMI-', SHEETS.ACCESSORIES_MANUAL_STOCK, 1);
  appendRowSafe_(SHEETS.ACCESSORIES_MANUAL_STOCK, [
    entryNo, d.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    'Issue', item, jo, qty, '', d.remarks || '', d.enteredBy || ''
  ]);
  return { success: true, message: entryNo + ': issued ' + qty + ' of "' + item + '" to ' + jo + '.' };
}
/** Current store balance per item: every Receipt qty minus every Issue qty, with a weighted average
    rate (from Receipts only) so the balance can also show an approximate stock value. */
function getAccessoryManualStockBalance() {
  var rows = sheetToObjects(SHEETS.ACCESSORIES_MANUAL_STOCK);
  var byItem = {}, order = [];
  rows.forEach(function (r) {
    var item = String(r['Item'] || '').trim();
    if (!item) return;
    if (!byItem[item]) { byItem[item] = { item: item, qtyIn: 0, qtyOut: 0, rateTotal: 0, rateQty: 0 }; order.push(item); }
    var qty = Number(r['Qty']) || 0;
    if (r['Type'] === 'Receipt') {
      byItem[item].qtyIn += qty;
      var rate = Number(r['Rate']) || 0;
      if (rate > 0) { byItem[item].rateTotal += rate * qty; byItem[item].rateQty += qty; }
    } else if (r['Type'] === 'Issue') {
      byItem[item].qtyOut += qty;
    }
  });
  return order.map(function (item) {
    var b = byItem[item];
    var avgRate = b.rateQty > 0 ? (b.rateTotal / b.rateQty) : 0;
    var balance = Math.round((b.qtyIn - b.qtyOut) * 1000) / 1000;
    return { item: item, qtyIn: b.qtyIn, qtyOut: b.qtyOut, balance: balance, avgRate: Math.round(avgRate * 100) / 100, value: Math.round(balance * avgRate * 100) / 100 };
  }).sort(function (a, b) { return a.item.localeCompare(b.item); });
}
/** Full ledger (both Receipt and Issue rows), newest first, for the activity register under the balance table. */
function getAccessoryManualStockLedger() {
  return sheetToObjects(SHEETS.ACCESSORIES_MANUAL_STOCK).reverse();
}
/** Distinct item names already used (manual stock + BOM/PO/Receipt) for the autocomplete suggestion list. */
function getAccessoryItemSuggestions() {
  var names = {};
  sheetToObjects(SHEETS.ACCESSORIES_MANUAL_STOCK).forEach(function (r) { if (r['Item']) names[String(r['Item']).trim()] = true; });
  sheetToObjects(SHEETS.ACCESSORIES_BOM).forEach(function (r) { if (r['Trims / Accessories']) names[String(r['Trims / Accessories']).trim()] = true; });
  return Object.keys(names).sort();
}

function getAccessoryReceiptsForJO(joNumber) {
  return sheetToObjects(SHEETS.ACCESSORIES_RECEIPT)
    .filter(function (r) { return String(r['JO Number']).trim() === String(joNumber).trim(); })
    .reverse();
}

/*************** ACCESSORIES BILL PASSING ***************/
/** Suppliers with an Approved, not-yet-billed Accessories PO. */
function getSuppliersWithBillableAccessoryPOs() {
  var billed = {};
  sheetToObjects(SHEETS.ACCESSORIES_BILL).forEach(function (b) {
    if (b['Status'] !== 'Rejected') billed[b['Accessories PO No']] = true;
  });
  var suppliers = {};
  sheetToObjects(SHEETS.ACCESSORIES_PO).forEach(function (r) {
    if (r['Status'] !== 'Approved') return;
    if (billed[r['PO No']]) return;
    var s = String(r['Supplier'] || '').trim();
    if (s) suppliers[s] = true;
  });
  return Object.keys(suppliers);
}

/** Approved, not-yet-billed Accessories POs for one supplier. */
function getBillableAccessoryPOsForSupplier(supplier) {
  var billed = {};
  sheetToObjects(SHEETS.ACCESSORIES_BILL).forEach(function (b) {
    if (b['Status'] !== 'Rejected') billed[b['Accessories PO No']] = true;
  });
  return sheetToObjects(SHEETS.ACCESSORIES_PO).filter(function (r) {
    return r['Status'] === 'Approved' && String(r['Supplier'] || '').trim() === String(supplier).trim() && !billed[r['PO No']];
  });
}

function saveAccessoryBillPassing(d) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var po = sheetToObjects(SHEETS.ACCESSORIES_PO).find(function (r) { return r['PO No'] === d.accessoryPoNo; });
    if (!po) return { success: false, error: 'Accessories PO "' + d.accessoryPoNo + '" not found.' };
    if (po['Status'] !== 'Approved') return { success: false, error: 'Accessories PO ' + d.accessoryPoNo + ' is "' + po['Status'] + '" — only an Approved PO can be billed.' };
    if (String(po['Supplier'] || '').trim() !== String(d.supplier).trim()) {
      return { success: false, error: 'Accessories PO ' + d.accessoryPoNo + ' was placed with ' + po['Supplier'] + ', not ' + d.supplier + '.' };
    }
    var alreadyBilled = sheetToObjects(SHEETS.ACCESSORIES_BILL).some(function (b) {
      return b['Status'] !== 'Rejected' && b['Accessories PO No'] === d.accessoryPoNo;
    });
    if (alreadyBilled) return { success: false, error: 'Accessories PO ' + d.accessoryPoNo + ' has already been billed (or is pending billing).' };

    var orderedQty = Number(po['Qty Ordered']) || 0;
    var rate = Number(po['Rate (Rs/unit)']) || 0;
    var expectedAmount = orderedQty * rate;
    var billAmount = Number(d.billAmount) || 0;
    if (billAmount > expectedAmount + 0.01) {
      return {
        success: false,
        error: 'Bill Amount (₹' + billAmount.toFixed(2) + ') cannot exceed Qty × Rate (₹' + expectedAmount.toFixed(2) + ' = ' + orderedQty + ' × ₹' + rate.toFixed(2) + ').'
      };
    }
    var variance = billAmount - expectedAmount;

    var gstPct = Number(d.gstPct) || 0;
    var gstAmount = billAmount * gstPct / 100;
    var totalInclGst = billAmount + gstAmount;
    var billStatedTotal = d.billStatedTotal !== undefined && d.billStatedTotal !== '' ? Number(d.billStatedTotal) : null;
    var gstMatch = billStatedTotal !== null ? (Math.abs(totalInclGst - billStatedTotal) <= 0.5 ? 'Match' : 'Mismatch') : '';

    var billNo = nextId('ABILL-', SHEETS.ACCESSORIES_BILL, 1);
    // The supplier's own Invoice No is normally already captured on the
    // Accessories Receipt when the goods came in — reuse it here rather
    // than re-typing it, falling back to whatever was typed on this form.
    var receipt = sheetToObjects(SHEETS.ACCESSORIES_RECEIPT).find(function (r) { return r['Accessories PO No'] === d.accessoryPoNo; });
    var invoiceNo = d.supplierInvoiceNo || (receipt ? receipt['Invoice No'] : '') || '';
    appendRowSafe_(SHEETS.ACCESSORIES_BILL, [
      billNo, d.billDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      d.supplier, po['JO Number'], d.accessoryPoNo, po['Trims / Accessories'], orderedQty, rate, expectedAmount, billAmount, variance,
      gstPct, gstAmount, totalInclGst, billStatedTotal === null ? '' : billStatedTotal, gstMatch,
      d.remarks || '', 'Pending Approval', '', '', '', invoiceNo
    ]);
    var gstNote = billStatedTotal !== null
      ? (gstMatch === 'Match' ? ' GST check: matches the bill.' : ' GST check: MISMATCH vs bill\'s stated total ₹' + billStatedTotal.toFixed(2) + '.')
      : '';
    return {
      success: true, billNo: billNo, expectedAmount: expectedAmount, variance: variance,
      message: 'Accessories Bill ' + billNo + ' submitted for approval. Expected ₹' + expectedAmount.toFixed(2) + ', variance ₹' + variance.toFixed(2) + '.' + gstNote
    };
  } catch (err) {
    return { success: false, error: err.message };
  } finally { lock.releaseLock(); }
}

function approveAccessoryBillPassing(billNo, decision, approver, reason) {
  var rowIdx = findRowIndex(SHEETS.ACCESSORIES_BILL, 'Bill No', billNo);
  if (rowIdx === -1) return { success: false, error: 'Accessories Bill not found: ' + billNo };
  var status = decision === 'APPROVE' ? 'Approved' : 'Rejected';
  setCell(SHEETS.ACCESSORIES_BILL, rowIdx, 'Status', status);
  setCell(SHEETS.ACCESSORIES_BILL, rowIdx, 'Approved By', approver || '');
  setCell(SHEETS.ACCESSORIES_BILL, rowIdx, 'Approval Date', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  setCell(SHEETS.ACCESSORIES_BILL, rowIdx, 'Rejection Reason', decision === 'REJECT' ? (reason || '') : '');
  return { success: true, message: 'Accessories Bill ' + billNo + ' ' + status + '.' };
}

function getAccessoryBillPassingQueue() { return sheetToObjects(SHEETS.ACCESSORIES_BILL); }

/**
 * "Do you have this accessory in stock elsewhere for this JO?" grid — a
 * plain list of every trim already received for the JO, from Accessories
 * Receipt, for a quick check before raising a new PO.
 */
function getAccessoryStockForJO(joNumber) {
  return sheetToObjects(SHEETS.ACCESSORIES_RECEIPT)
    .filter(function (r) { return String(r['JO Number']).trim() === String(joNumber).trim(); });
}

/**
 * Distinct Counts within a JO, each with its own required fabric qty (kg)
 * summed across every Part/Color row that uses that Count, plus the fabric
 * spec of the first matching row (type, yarn color, rate, Yarn Type/Content).
 * Lets Yarn PO be raised per-Count when a JO mixes counts across parts.
 */
/**
 * Distinct fabric "specs" within a JO (+ Color filter, for color-specific PO
 * types). A spec is Counts + Content + GSM together — not Counts alone, since
 * a JO's Body and Rib rows can share the same Counts (e.g. "30s") while being
 * physically different fabrics (different GSM, sometimes different Content).
 * Grouping by Counts only would silently blend their required qty and let a
 * PO for one part's fabric double as cover for the other's.
 */
function getCountsForJO(joNumber, color, useGsm) {
  var rows = getStyleTracker().filter(function (r) {
    if (String(r['JO Number']).trim() !== String(joNumber).trim()) return false;
    if (color && String(r['Color'] || '').trim() !== String(color).trim()) return false;
    return true;
  });
  // Lycra is its own YARN: for raw-yarn specs (useGsm false) each row's lycra kg becomes a separate
  // "LYCRA" line (bought on its own Yarn PO, stocked and issued to knitting on its own), and the main
  // count only needs Main Fiber kg. Knitted-fabric specs (Cora / Dyed fabric) keep the full fabric kg.
  var lyStock = rows.length ? dcxLycraStockSpec_(rows) : { counts: 'LYCRA', content: '' };
  var order = [];
  var map = {};
  var add = function (key, base, kg, r) {
    if (!map[key]) { order.push(key); base.requiredKg = 0; base.sampleKg = 0; base.parts = []; base.colors = []; map[key] = base; }
    // A grouped fabric spec may appear on several Style Tracker rows. If its
    // first row has no knitting dia, take the first nonblank dia from another
    // row belonging to the same exact spec.
    if (!String(map[key].knittingDia || '').trim() && String(base.knittingDia || '').trim()) {
      map[key].knittingDia = String(base.knittingDia).trim();
    }
    map[key].requiredKg += kg;
    if (isSampleWtRow_(r)) map[key].sampleKg += kg;
    if (map[key].parts.indexOf(r['Part']) === -1) map[key].parts.push(r['Part'] || '');
    var col = r['Color'] || '';
    if (col && map[key].colors.indexOf(col) === -1) map[key].colors.push(col);
  };
  rows.forEach(function (r) {
    var counts = String(r['Counts'] || '').trim();
    if (!counts) return;
    var content = String(r['Content'] || '').trim();
    var gsm = String(r['GSM'] || '').trim();
    var yarnColor = String(r['Yarn color'] || '').trim();
    var kg = Number(r['Fabric Required (kg)']) || 0;
    // GSM is a knitted-fabric property, not a yarn property — only Cora / Dyed Fabric specs split by it.
    // Yarn Color is always split (Greige vs Melange can share Counts + Content).
    var ly = useGsm ? { pure: false, mainKg: kg, lycraKg: 0 } : dcxLycra_(r);
    var lyKey = 'LYCRA||' + lyStock.counts + '||' + lyStock.content;
    var lyBase = function () {
      return { counts: lyStock.counts, content: lyStock.content, gsm: '', fabricType: 'Lycra yarn', yarnColor: '',
        rate: ly.pure ? (Number(r['Yarn rate']) || 0) : (Number(r['Lycra rate']) || 0), yarnType: 'Lycra / Spandex', isLycra: true };
    };
    if (ly.pure) { add(lyKey, lyBase(), kg, r); return; }
    var key = counts + '||' + content + (useGsm ? '||' + gsm : '') + '||' + yarnColor;
    add(key, { counts: counts, content: content, gsm: useGsm ? gsm : '', knittingDia: String(r['Knitting Dia'] || '').trim(), fabricType: r['Fabric type'] || '', yarnColor: yarnColor,
      rate: Number(r['Yarn rate']) || 0, yarnType: r['Yarn Type'] || '' }, ly.mainKg, r);
    if (ly.lycraKg > 0) add(lyKey, lyBase(), ly.lycraKg, r);
  });
  return order.map(function (k) { map[k].requiredKg = Math.round(map[k].requiredKg * 1000) / 1000; map[k].sampleKg = Math.round(map[k].sampleKg * 1000) / 1000; return map[k]; });
}

/** Distinct colors for a JO, in first-seen order — for color-specific PO types. */
function getColorsForJO(joNumber) {
  var rows = getStyleTracker().filter(function (r) { return String(r['JO Number']).trim() === String(joNumber).trim(); });
  var order = [];
  rows.forEach(function (r) {
    var c = String(r['Color'] || '').trim();
    if (c && order.indexOf(c) === -1) order.push(c);
  });
  return order;
}

/**
 * Kg already placed on Yarn PO for a JO (+ Counts, + Color if given), excluding
 * Rejected POs — Pending and Approved both hold against the allowance so two
 * people can't both order the full required qty before either is approved.
 * Color is only meaningful for Dyed Fabric POs (grey Yarn/Cora Fabric POs are
 * shared across every colorway of the JO, so they leave Color blank).
 */
/**
 * Net kg moved OUT of a JO for one Counts(+Content+GSM) via JO Transfer
 * (Yarn stock type only) minus whatever moved IN — positive means the JO
 * physically has less yarn on hand than its own Style Tracker requirement
 * accounts for, so that much MORE needs to be purchased to make up for it;
 * negative means incoming transfers already cover part of the requirement.
 */
function getJOTransferNetOut(joNumber, counts, content, gsm) {
  var transfers = sheetToObjects(SHEETS.JO_TRANSFER).filter(function (r) {
    if (r['Stock Type'] !== 'Yarn') return false;
    if (String(r['Counts'] || '').trim() !== String(counts || '').trim()) return false;
    if (content && String(r['Content'] || '').trim() !== String(content).trim()) return false;
    if (gsm && String(r['GSM'] || '').trim() !== String(gsm).trim()) return false;
    if (r['Status'] === 'Rejected') return false;
    return true;
  });
  var out = transfers.filter(function (r) { return String(r['Source JO']).trim() === String(joNumber).trim(); })
    .reduce(function (s, r) { return s + (Number(r['Qty KG']) || 0); }, 0);
  var inn = transfers.filter(function (r) { return String(r['Destination JO']).trim() === String(joNumber).trim(); })
    .reduce(function (s, r) { return s + (Number(r['Qty KG']) || 0); }, 0);
  return out - inn;
}

/**
 * Effective quantity already covering a JO+Counts(+Color+Content+GSM+YarnColor)'s
 * requirement — Yarn PO qty placed, adjusted for any net JO Transfer in/out
 * of that same spec (yarn transferred away no longer counts as covering this
 * JO's own need; yarn transferred in does).
 */
function getYarnBalanceInfo(joNumber, counts, color, content, gsm, yarnColor) {
  return { used: getYarnPOUsage(joNumber, counts, color, content, gsm, yarnColor),
    extraApproved: getYarnExtraApproved({ joNumber: joNumber, counts: counts, content: content, gsm: gsm, color: color, yarnColor: yarnColor }),
    isSample: isSampleJO_(joNumber) };
}
function getYarnPOUsage(joNumber, counts, color, content, gsm, yarnColor) {
  var poQty = sheetToObjects(SHEETS.YARN_PO)
    .filter(function (r) {
      if (String(r['JO Number']).trim() !== String(joNumber).trim()) return false;
      if (r['Status'] === 'Rejected') return false;
      if (counts && String(r['Counts'] || '').trim() !== String(counts).trim()) return false;
      if (color && String(r['Color'] || '').trim() !== String(color).trim()) return false;
      if (content && String(r['Content'] || '').trim() !== String(content).trim()) return false;
      if (gsm && String(r['GSM'] || '').trim() !== String(gsm).trim()) return false;
      if (yarnColor && String(r['Yarn Color'] || '').trim() !== String(yarnColor).trim()) return false;
      return true;
    })
    .reduce(function (s, r) { return s + (Number(r['Qty KG']) || 0); }, 0);
  var netTransferOut = getJOTransferNetOut(joNumber, counts, content, gsm);
  return poQty - netTransferOut;
}

function getJOInfo(joNumber) {
  var rows = getStyleTracker().filter(function (r) { return String(r['JO Number']).trim() === String(joNumber).trim(); });
  if (!rows.length) return null;
  var first = rows[0];
  var totalFabricKg = rows.reduce(function (s, r) { return s + Number(r['Fabric Required (kg)'] || 0); }, 0);
  var accessories = getAccessoriesBOM(joNumber);
  // Aliases for older frontend code that still reads 'Buyer' / 'Yarn Rate' / 'Order Qty' / 'Excess Qty'.
  return Object.assign({}, first, {
    Buyer: first['Buyer Name'],
    'Yarn Rate': first['Yarn rate'],
    'Order Qty': first['Order Quantity'],
    'Excess Qty': first['Excess Quantity'],
    totalFabricRequiredKg: totalFabricKg, parts: rows, accessories: accessories
  });
}

function getStyleTrackerPdfData(joNumber) {
  var rows = getStyleTracker().filter(function (r) {
    return String(r['JO Number']).trim() === String(joNumber == null ? '' : joNumber).trim();
  });
  if (!rows.length) return null;

  var wanted = String(joNumber).trim();
  var first = rows[0];

  // Size names: union of Size 1..10 across all rows, in first-seen order.
  var sizeNames = [];
  rows.forEach(function (r) {
    for (var i = 1; i <= 10; i++) {
      var lbl = String(r['Size ' + i] || '').trim();
      if (lbl && sizeNames.indexOf(lbl) === -1) sizeNames.push(lbl);
    }
  });
  if (!sizeNames.length) sizeNames = ['M', 'L', 'XL', '2XL', '3XL'];

  // Color rows: one per unique Color. The same size curve is repeated across
  // every Part for a given color (Body/Rib/Folding all ship the same pieces),
  // so take the size quantities from the first row seen for that color rather
  // than summing across parts, which would multiply the true order quantity.
  var colorOrder = [];
  var colorMap = {};
  rows.forEach(function (r) {
    var color = String(r['Color'] || '').trim() || '-';
    if (colorMap[color]) return;
    colorOrder.push(color);
    var sizes = {};
    for (var i = 1; i <= 10; i++) {
      var lbl = String(r['Size ' + i] || '').trim();
      if (lbl) sizes[lbl] = Number(r['Qty ' + i]) || 0;
    }
    colorMap[color] = sizes;
  });

  var colorRows = colorOrder.map(function (color) {
    var sizes = colorMap[color];
    var sizeCells = sizeNames.map(function (n) { return Number(sizes[n]) || 0; });
    var total = sizeCells.reduce(function (a, b) { return a + b; }, 0);
    return { color: color, sizeCells: sizeCells, total: total };
  });

  var sizeTotals = sizeNames.map(function (n, idx) {
    return colorRows.reduce(function (s, c) { return s + c.sizeCells[idx]; }, 0);
  });
  var totalQty = sizeTotals.reduce(function (a, b) { return a + b; }, 0);
  var sizePercents = sizeTotals.map(function (x) { return totalQty ? Math.round(x * 100 / totalQty) : 0; });

  // Fabric / consumption: each Style Tracker row is one Part+Color with its
  // own fabric spec and its own already-computed Fabric Required (kg).
  var fabrics = rows.map(function (r) {
    return {
      part: r['Part'] || '',
      color: r['Color'] || '',
      fabricType: r['Fabric type'] || '',
      counts: r['Counts'] || '',
      yarnColor: r['Yarn color'] || '',
      yarnType: r['Yarn Type'] || '',
      content: r['Content'] || '',
      gsm: r['GSM'] || '',
      cadWeight: Number(r['CAD Weight (g)']) || 0,
      lossPct: Number(r['Loss %']) || 0,
      rate: Number(r['Yarn rate']) || 0,
      required: Number(r['Fabric Required (kg)']) || 0
    };
  });

  var accessories = getAccessoriesBOM(wanted).map(function (a) {
    return {
      name: a['Trims / Accessories'] || '',
      color: a['Color'] || '',
      size: a['Size'] || '',
      qtyNeeded: Number(a['Qty Needed'] || 0),
      perGarments: Number(a['Per How Many Garments'] || 1),
      requiredQty: Number(a['Required Quantity'] || 0),
      rate: Number(a['Rate (Rs/unit)'] || 0),
      amount: Number(a['Amount'] || 0),
      ihNeededDate: a['I/H Needed Date'] || ''
    };
  });

  var totalFabricRequiredKg = rows.reduce(function (s, r) { return s + Number(r['Fabric Required (kg)'] || 0); }, 0);

  // Aliased header fields so the existing order-sheet renderer (which expects
  // 'Buyer' / 'Order Qty' / 'Excess Qty') keeps working unchanged.
  var jo = Object.assign({}, first, {
    Buyer: first['Buyer Name'],
    'Order Qty': first['Order Quantity'],
    'Excess Qty': first['Excess Quantity']
  });

  return {
    joNumber: wanted,
    jo: jo,
    quantity: {
      sizeNames: sizeNames,
      rows: colorRows,
      sizeTotals: sizeTotals,
      sizePercents: sizePercents,
      totalQty: totalQty
    },
    sizes: colorRows,
    fabrics: fabrics,
    accessories: accessories,
    totalFabricRequiredKg: Number(totalFabricRequiredKg.toFixed(2))
  };
}

/*************** 3. PROCESS RATES (columns on Style Tracker, per JO) ***************/
function getProcessMaster() { return sheetToObjects(SHEETS.PROCESS_MASTER).map(function (r) { return r['Process Name']; }); }

// Maps a display process name to the actual Style Tracker rate column that
// holds it, and to a short DC-number prefix for that process.
var RATE_COLUMN_MAP = {
  'Cone Winding': 'Cone winding rate',
  'Yarn Dyeing': 'Yarn dyeing rate',
  'Knitting': 'Knitting rate',
  'Heatsetting': 'Heatsetting rate',
  'Dyeing': 'Dyeing rate',
  'Tubular Compacting': 'Tubular compacting rate',
  'Stenter': 'Stener rate',
  'Compacting': 'Compacting rate',
  'Raising/Peach Finish': 'Rasing /Peach fnish',
  'Bio Washing': 'Bio washing rate'
};

var PROCESS_CODE_MAP = {
  'Cone Winding': 'CW',
  'Yarn Dyeing': 'YD',
  'Knitting': 'K',
  'Heatsetting': 'H',
  'Dyeing': 'D',
  'Tubular Compacting': 'TC',
  'Stenter': 'ST',
  'Compacting': 'C',
  'Raising/Peach Finish': 'RP',
  'Bio Washing': 'BW'
};

// Each process's own Loss % column on Style Tracker, for the Fabric Program's
// forward stage-by-stage calculation (yarn qty -> each stage's output qty).
var LOSS_COLUMN_MAP = {
  'Cone Winding': 'Cone winding loss %',
  'Yarn Dyeing': 'Yarn dyeing loss %',
  'Knitting': 'Knitting loss %',
  'Heatsetting': 'Heatsetting loss %',
  'Dyeing': 'Dyeing loss %',
  'Tubular Compacting': 'Tubular compacting loss %',
  'Stenter': 'Stenter loss %',
  'Compacting': 'Compacting loss %',
  'Bio Washing': 'Bio washing loss %',
  'Raising/Peach Finish': 'Raising loss %'
};

// Default real-world process order for the Fabric Program's forward walk.
// A JO only walks through whichever of these it has a rate set for
// (via getAllowedProcessesForJO) — this just fixes the sequence.
var PROCESS_SEQUENCE = ['Yarn Dyeing', 'Cone Winding', 'Knitting', 'Heatsetting', 'Dyeing',
  'Tubular Compacting', 'Stenter', 'Compacting', 'Bio Washing', 'Raising/Peach Finish'];

function getProcessRates(joNumber) {
  var jo = getJOInfo(joNumber);
  if (!jo) return [];
  var maps = getProcessMaps();
  return maps.sequence.map(function (process) {
    return { 'JO Number': joNumber, 'Process Name': process, 'Rate (Rs/kg)': Number(jo[maps.rateMap[process]]) || 0 };
  });
}

function getAllowedProcessesForJO(joNumber) {
  var jo = getJOInfo(joNumber);
  if (!jo) return [];
  var maps = getProcessMaps();
  return maps.sequence
    .map(function (process) { return { process: process, rate: Number(jo[maps.rateMap[process]]) || 0 }; })
    .filter(function (x) { return x.rate > 0; });
}

function getRateForJOProcess(joNumber, processName) {
  var jo = getJOInfo(joNumber);
  var maps = getProcessMaps();
  var col = maps.rateMap[processName];
  if (!jo || !col) return 0;
  return Number(jo[col]) || 0;
}

/*************** 5. YARN PO (Yarn / Cora Fabric / Dyed Fabric) ***************/

/* Most recent approved/billed rate paid for this exact Counts + Content (+ Yarn Colour), from any JO —
   a sensible starting rate for Auto Yarn PO so the only things a person must type are the mill, the
   knitting unit, and the bag weight. Returns 0 if nothing on file yet for this spec. */
function getLastYarnRate_(counts, content, yarnColor) {
  // "30's", "30s" and "30S" are the same yarn, and stray double spaces in a content ("60% C  40% P") must not
  // make an old PO look like a different one — compare them the way a person would.
  var n = function (v) { return String(v || '').trim().toLowerCase().replace(/\s+/g, ' '); };
  var c = function (v) { return n(normalizeCounts_(String(v || ''))); };
  var want = { counts: c(counts), content: n(content), yc: n(yarnColor) };
  var rows = sheetToObjects(SHEETS.YARN_PO).filter(function (r) {
    return String(r['PO Type'] || 'Yarn').trim() === 'Yarn' &&
      c(r['Counts']) === want.counts &&
      n(r['Content']) === want.content &&
      (!want.yc || n(r['Yarn Color']) === want.yc) &&
      (Number(r['Rate (Rs/kg)']) || 0) > 0 && (r['Status'] === 'Approved' || r['Status'] === 'Billed');
  });
  if (!rows.length) return 0;
  rows.sort(function (a, b) { return new Date(b['PO Date']) - new Date(a['PO Date']); });
  return Number(rows[0]['Rate (Rs/kg)']) || 0;
}

/* The lines an Auto Yarn PO would raise for a locked JO: one per Counts/Content/Yarn Colour still pending,
   whole bags, rate from the last PO paid for that spec. Shared by the preview and the real thing so what
   is shown is exactly what gets created. Returns {error} or {lines, skipped}. */
function autoYarnPoBuild_(jo, bagWeight) {
  return withSheetCache_(function () {
  jo = String(jo || '').trim();
  if (!jo) return { error: 'Pick a JO.' };
  if (!getJoDataLock(jo)) return { error: jo + ' is not locked yet. Lock the Order Sheet + Fabric Program first — Auto Yarn PO only works from a finalized JO.' };
  bagWeight = Number(bagWeight) || 0;
  if (bagWeight <= 0) return { error: 'Enter a bag weight greater than zero.' };

  var jInfo = getJOInfo(jo) || {};
  var specs = getCountsForJO(jo, '', false);   // Yarn PO: useGsm = false (GSM is a knitted-fabric property)
  var lines = [], skipped = [];
  specs.forEach(function (spec) {
    var yarnColor = spec.yarnColor || '';
    var already = getYarnPOUsage(jo, spec.counts, '', spec.content, '', yarnColor);
    var extra = getYarnExtraApproved({ joNumber: jo, counts: spec.counts, content: spec.content, gsm: '', color: '', yarnColor: yarnColor });
    var balance = Math.max(0, (Number(spec.requiredKg) || 0) + extra - already);
    if (isSampleJO_(jo)) balance = Math.max(balance, Math.max(0, SAMPLE_TRIAL_MIN_KG - already));
    if (balance <= 0.01) return;   // this spec is already fully covered
    var bags = Math.ceil((balance - 0.000001) / bagWeight);
    var qty = bags * bagWeight;
    var rate = getLastYarnRate_(spec.counts, spec.content, yarnColor);
    if (!(rate > 0)) { skipped.push(spec.counts + (yarnColor ? ' / ' + yarnColor : '') + (spec.content ? ' / ' + spec.content : '')); return; }
    lines.push({
      joNumber: jo, counts: spec.counts, content: spec.content, gsm: '', yarnColor: yarnColor,
      qtyKg: qty, rate: rate, styleNumber: jInfo['Style Number'] || '', buyer: jInfo['Buyer Name'] || '',
      color: '', bagWeight: bagWeight, noOfBags: bags
    });
  });
  return { lines: lines, skipped: skipped };
  });
}

/* Locked JOs that still have yarn to order — what the Auto Yarn PO picker on the Yarn PO page lists. */
function getAutoYarnPoJos() {
  return withSheetCache_(function () {
  var pending = {};
  getJosNeedingYarnPO().forEach(function (j) { pending[j] = true; });
  var sh = SS.getSheetByName(JO_DATA_LOCK_SHEET), out = [];
  if (sh && sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) {
      var jo = String(r[0] || '').trim();
      if (jo && pending[jo] && out.indexOf(jo) < 0) out.push(jo);
    });
  }
  return out;
  });
}

/* What Auto Yarn PO would raise, without raising it — shown on the Yarn PO page before anything is created. */
function previewAutoYarnPO(jo, bagWeight) {
  var b = autoYarnPoBuild_(jo, bagWeight);
  if (b.error) return { success: false, error: b.error };
  var totalKg = 0, totalAmount = 0;
  var lines = b.lines.map(function (l) {
    totalKg += l.qtyKg; totalAmount += l.qtyKg * l.rate;
    return { counts: normalizeCounts_(l.counts), content: l.content, yarnColor: l.yarnColor, qtyKg: l.qtyKg, bags: l.noOfBags, rate: l.rate, amount: Math.round(l.qtyKg * l.rate * 100) / 100 };
  });
  if (!lines.length && !b.skipped.length) return { success: false, error: String(jo).trim() + ' has no pending yarn requirement — nothing to raise.' };
  return { success: true, lines: lines, skipped: b.skipped, totalKg: Math.round(totalKg * 100) / 100, totalAmount: Math.round(totalAmount * 100) / 100 };
}

/* Auto Yarn PO — only available once a JO is locked (Order Sheet + Fabric Program final), since that's
   what makes its yarn requirement trustworthy enough to raise a PO from automatically. A spec with no
   rate on file yet is skipped rather than guessed — reported back so it can be added by hand. The person
   only ever types the mill, the knitting unit, and the bag weight. */
function autoCreateYarnPO(jo, millName, knittingUnitName, bagWeight) {
  millName = String(millName || '').trim();
  knittingUnitName = String(knittingUnitName || '').trim();
  if (!String(jo || '').trim()) return { success: false, error: 'Pick a JO.' };
  if (!millName) return { success: false, error: 'Enter the Mill (supplier) name.' };
  if (!knittingUnitName) return { success: false, error: 'Enter the Knitting Unit name.' };
  var b = autoYarnPoBuild_(jo, bagWeight);
  if (b.error) return { success: false, error: b.error };
  bagWeight = Number(bagWeight) || 0;
  var lines = b.lines, skipped = b.skipped;

  if (!lines.length) {
    return {
      success: false,
      error: skipped.length
        ? ('Nothing could be auto-filled — every pending spec has no past rate on file yet: ' + skipped.join('; ') + '. Raise those lines manually once, and Auto Yarn PO will find them next time.')
        : (String(jo).trim() + ' has no pending yarn requirement — nothing to raise.')
    };
  }
  try {
    // The same details the manual PO form fills in by itself once the mill / knitting unit is chosen, taken
    // from Supplier Master — otherwise the printed PO would have no vendor GSTIN, address or delivery point.
    var masters = getSuppliers();
    var byName = function (name) {
      var w = String(name || '').trim().toLowerCase();
      return masters.find(function (m) { return String(m['Supplier Name'] || '').trim().toLowerCase() === w; }) || {};
    };
    var mill = byName(millName), knit = byName(knittingUnitName);
    var res = saveYarnPOMultiLines_({
      poType: 'Yarn', supplier: millName, supplierAddress: mill['Address'] || '', gstNo: mill['GST'] || '',
      knittingUnitName: knittingUnitName, deliveryAddress: knit['Address'] || '',
      knittingUnitPhone: knit['Contact Number'] || '', knittingUnitGstin: knit['GST'] || '',
      bagWeight: bagWeight,
      poDate: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'), preparedBy: 'Auto Yarn PO'
    }, lines);
    if (res.success && skipped.length) res.message += ' ' + skipped.length + ' spec(s) skipped (no past rate on file): ' + skipped.join('; ') + '.';
    return res;
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function saveYarnPOMultiLines_(d, lines) {
  if (!lines || !lines.length) throw new Error('Add at least one PO line.');
  var poType = d.poType || 'Yarn';
  var useGsm = poType !== 'Yarn';
  var isColorSpecific = poType === 'Dyed Fabric' && d.color;
  var colorForCheck = isColorSpecific ? d.color : '';
  var prepared = [];
  var seenLineSpecs = {};
  lines.forEach(function (line, i) {
    var jo = String(line.joNumber || '').trim();
    var counts = String(line.counts || '').trim();
    var content = String(line.content || '').trim();
    var gsm = useGsm ? String(line.gsm || '').trim() : '';
    var yarnColor = String(line.yarnColor || '').trim();
    var lineKey = [jo,counts,content,gsm,yarnColor,colorForCheck].join('|').toLowerCase();
    if (seenLineSpecs[lineKey]) throw new Error('Duplicate JO / count specification on lines ' + seenLineSpecs[lineKey] + ' and ' + (i+1) + '. Keep one line per JO and count specification.');
    seenLineSpecs[lineKey] = i+1;
    var qty = Number(line.qtyKg) || 0;
    var rate = Number(line.rate) || 0;
    if (!jo || !counts || qty <= 0 || rate <= 0) throw new Error('Line ' + (i + 1) + ': JO, count, positive quantity and rate are required.');
    var specs = getCountsForJO(jo, colorForCheck, useGsm);
    var counted = specs.find(function (c) {
      return String(c.counts || '').trim() === counts &&
        String(c.content || '').trim() === content &&
        String(c.gsm || '').trim() === gsm &&
        String(c.yarnColor || '').trim() === yarnColor;
    });
    if (!counted) throw new Error('Line ' + (i + 1) + ': specification not found for ' + jo + ' / ' + counts + ' / ' + gsm + ' GSM.');
    var already = getYarnPOUsage(jo, counts, colorForCheck, content, gsm, yarnColor);
    var extra = getYarnExtraApproved({joNumber:jo,counts:counts,content:content,gsm:gsm,color:colorForCheck,yarnColor:yarnColor});
    var balance = Math.max(0, (Number(counted.requiredKg)||0) + extra - already);
    if (isSampleJO_(jo)) balance = Math.max(balance, Math.max(0, SAMPLE_TRIAL_MIN_KG - already));
    if (balance <= 0.01) throw new Error('Line ' + (i + 1) + ': ' + jo + ' / ' + counts + ' is already fully covered.');
    var isManualFabric = poType === 'Cora Fabric' || poType === 'Dyed Fabric';
    var bagWeight = Number(line.bagWeight || d.bagWeight) || 0;
    if (bagWeight <= 0) throw new Error('Line ' + (i + 1) + ': roll weight must be greater than zero.');
    var bags = 0;
    if (isManualFabric) {
      // Cora/Dyed Fabric: roll weight and PO quantity are manual. Do not round
      // quantity to bags and do not permit any quantity above the exact pending balance.
      if (qty > balance + 0.001) {
        throw new Error('Line ' + (i + 1) + ': quantity ' + qty.toFixed(2) + ' kg exceeds pending balance ' + balance.toFixed(2) + ' kg.');
      }
    } else {
      bags = Number(line.noOfBags) || Math.ceil((qty - 0.000001) / bagWeight);
      var wholeBagQty = bags * bagWeight;
      if (bags <= 0 || Math.abs(wholeBagQty - qty) > 0.02) {
        throw new Error('Line ' + (i + 1) + ': PO quantity must be in full bags (' + bagWeight + ' kg per bag).');
      }
      // Permit the next whole-bag quantity needed to cover the pending balance, even if
      // the bag rounding exceeds the generic kg tolerance.
      var maxWholeBagQty = Math.ceil((balance - 0.000001) / bagWeight) * bagWeight;
      if (qty > maxWholeBagQty + 0.02) {
        throw new Error('Line ' + (i + 1) + ': quantity ' + qty.toFixed(2) + ' kg exceeds the full-bag quantity needed for pending balance ' + balance.toFixed(2) + ' kg (' + maxWholeBagQty.toFixed(2) + ' kg).');
      }
    }
    var knittingDia = String(line.knittingDia || '').trim();
    if (poType === 'Cora Fabric' && !knittingDia) {
      var diaRows = getStyleTracker().filter(function (r) {
        return String(r['JO Number']||'').trim()===jo &&
          String(r['Counts']||'').trim()===counts &&
          String(r['Content']||'').trim()===content &&
          String(r['GSM']||'').trim()===gsm &&
          String(r['Knitting Dia']||'').trim()!=='';
      });
      if (diaRows.length) knittingDia=String(diaRows[0]['Knitting Dia']).trim();
    }
    prepared.push({jo:jo,counts:counts,content:content,gsm:gsm,yarnColor:yarnColor,qty:qty,rate:rate,
      styleNumber:line.styleNumber||'',buyer:line.buyer||'',color:line.color||d.color||'',
      knittingDia:knittingDia,bagWeight:bagWeight,bags:bags});
  });
  var poNo = nextId('YPO-', SHEETS.YARN_PO, 1);
  var poDate=d.poDate || Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
  prepared.forEach(function (line) {
    appendRowSafe_(SHEETS.YARN_PO, [
      poNo, poDate, poType, line.jo, line.styleNumber, line.buyer, line.color, line.counts, line.yarnColor,
      line.rate, line.qty, line.rate*line.qty, d.supplier, d.deliveryAddress, d.gstNo,
      'Pending Approval', '', '', '', line.bagWeight, line.bags, line.content, line.gsm,
      Number(d.gstPct)||0, d.knittingUnitPhone||'', d.preparedBy||'',
      d.paymentTerms||'', d.deliveryDate||'', d.knittingUnitName||'', d.knittingUnitGstin||'',
      d.supplierAddress||'', line.knittingDia
    ]);
  });
  return {success:true,poNumber:poNo,message:'PO '+poNo+' submitted with '+prepared.length+' JO/count lines — one approval and one combined PDF.'};
}

function saveYarnPO(d) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var _lineItems = [];
    try { _lineItems = JSON.parse(String(d.lineItemsJson || '[]')); } catch (_) { _lineItems = []; }
    if (_lineItems.length) return saveYarnPOMultiLines_(d, _lineItems);
    var qty = Number(d.qtyKg) || 0;
    var rate = Number(d.rate) || 0;
    var isColorSpecific = (d.poType === 'Dyed Fabric') && d.color;
    var colorForCheck = isColorSpecific ? d.color : '';
    var useGsm = d.poType !== 'Yarn';
    var content = String(d.content || '').trim();
    var gsm = useGsm ? String(d.gsm || '').trim() : '';
    // Resolve knitting dia on the server from the JO's Style Tracker as well,
    // so the CORA PO still gets the dia even if the browser submits it blank.
    var knittingDia = String(d.knittingDia || '').trim();
    if (d.poType === 'Cora Fabric' && !knittingDia && d.joNumber && d.counts) {
      var diaRows = getStyleTracker().filter(function (r) {
        return String(r['JO Number'] || '').trim() === String(d.joNumber).trim() &&
          String(r['Counts'] || '').trim() === String(d.counts).trim() &&
          String(r['Content'] || '').trim() === content &&
          String(r['GSM'] || '').trim() === gsm &&
          String(r['Knitting Dia'] || '').trim() !== '';
      });
      if (diaRows.length) knittingDia = String(diaRows[0]['Knitting Dia']).trim();
    }

    // Multi-count PO: validate the combined quantity against each selected count's remaining balance.
    var multiSpecs = [];
    try { multiSpecs = JSON.parse(String(d.multiSpecsJson || '[]')); } catch (_) { multiSpecs = []; }
    if (d.joNumber && multiSpecs.length > 1) {
      var combinedBalance = 0;
      var coveredSpecs = [];
      multiSpecs.forEach(function (spec) {
        var specCounts = String(spec.counts || '').trim();
        var specContent = String(spec.content || '').trim();
        var specGsm = useGsm ? String(spec.gsm || '').trim() : '';
        var specYarnColor = String(spec.yarnColor || '').trim();
        var specRows = getCountsForJO(d.joNumber, colorForCheck, useGsm);
        var countedSpec = specRows.find(function (c) {
          return c.counts === specCounts && c.content === specContent && c.gsm === specGsm && c.yarnColor === specYarnColor;
        });
        if (!countedSpec) throw new Error('Selected count specification not found: ' + specCounts + ' / ' + specContent + ' / ' + specGsm);
        var specUsed = getYarnPOUsage(d.joNumber, specCounts, colorForCheck, specContent, specGsm, specYarnColor);
        var specExtra = getYarnExtraApproved({ joNumber: d.joNumber, counts: specCounts, content: specContent, gsm: specGsm, color: colorForCheck, yarnColor: specYarnColor });
        var specBalance = Math.max(0, countedSpec.requiredKg + specExtra - specUsed);
        if (isSampleJO_(d.joNumber)) specBalance = Math.max(specBalance, Math.max(0, SAMPLE_TRIAL_MIN_KG - specUsed));
        if (specBalance <= 0.01) coveredSpecs.push(specCounts);
        combinedBalance += specBalance;
      });
      if (combinedBalance <= 0.01) return { success: false, error: 'All selected counts are already fully covered: ' + coveredSpecs.join(', ') };
      if (qty > combinedBalance + YARN_PO_TOLERANCE_KG + 0.001) {
        return { success: false, error: 'Combined PO quantity (' + qty.toFixed(2) + ' kg) exceeds the remaining balance for selected counts (' + combinedBalance.toFixed(2) + ' kg) plus allowed tolerance of ' + YARN_PO_TOLERANCE_KG + ' kg.' };
      }
    }

    if (d.joNumber && d.counts && multiSpecs.length <= 1) {
      var yarnColorForCheck = String(d.yarnColor || '').trim();
      var counted = getCountsForJO(d.joNumber, colorForCheck, useGsm).find(function (c) {
        return c.counts === d.counts && c.content === content && c.gsm === gsm && c.yarnColor === yarnColorForCheck;
      });
      if (counted) {
        var already = getYarnPOUsage(d.joNumber, d.counts, colorForCheck, content, gsm, yarnColorForCheck);
        var extraApproved = getYarnExtraApproved({ joNumber: d.joNumber, counts: d.counts, content: content, gsm: gsm, color: colorForCheck, yarnColor: yarnColorForCheck });
        var balance = counted.requiredKg + extraApproved - already;
        var isSample = isSampleJO_(d.joNumber);
        var sampleRoom = isSample ? Math.max(0, SAMPLE_TRIAL_MIN_KG - already) : 0;
        // A sample JO's bulk-calculated requirement is usually tiny, but mills still need a real
        // trial lot — so a sample order is never blocked below SAMPLE_TRIAL_MIN_KG total.
        var effectiveBalance = Math.max(balance, sampleRoom);
        // Once the requirement is already fully covered by prior POs, no
        // further quantity is allowed at all — no bag-rounding tolerance
        // applies here, since there's nothing left to round up to.
        if (effectiveBalance <= 0.01) {
          return {
            success: false,
            error: d.joNumber + ' / Counts ' + d.counts + (gsm ? ' / GSM ' + gsm : '') + (content ? ' / ' + content : '') +
              (colorForCheck ? ' / Color ' + colorForCheck : '') + (yarnColorForCheck ? ' / Yarn Color ' + yarnColorForCheck : '') +
              ' is already fully covered (required: ' + counted.requiredKg.toFixed(2) + ' kg' + (extraApproved ? ' + approved extra ' + extraApproved.toFixed(2) + ' kg' : '') + ', already on PO: ' + already.toFixed(2) +
              ' kg). No further quantity can be ordered for this spec — use \'Request Extra\' if more is genuinely needed.'
          };
        }
        // Yarn ships in whole bags, so qty (bags x bag weight) will routinely
        // land a bit over the exact fractional balance — allow up to one
        // extra bag's worth of rounding rather than hard-blocking every order.
        // Whole-bag rounding may go over the balance, but never by more than 20 kg.
        var tolerance = (d.poType === 'Cora Fabric' || d.poType === 'Dyed Fabric') ? 0.001 : (YARN_PO_TOLERANCE_KG + 0.001);
        if (qty > effectiveBalance + tolerance) {
          return {
            success: false,
            error: 'Qty (' + qty + ' kg) exceeds the required qty for ' + d.joNumber + ' / Counts ' + d.counts +
              (gsm ? ' / GSM ' + gsm : '') + (content ? ' / ' + content : '') +
              (colorForCheck ? ' / Color ' + colorForCheck : '') + (yarnColorForCheck ? ' / Yarn Color ' + yarnColorForCheck : '') +
              ' by more than ' + (d.poType === 'Cora Fabric' || d.poType === 'Dyed Fabric' ? 'the exact pending balance' : (YARN_PO_TOLERANCE_KG + ' kg')) + '. Required: ' + counted.requiredKg.toFixed(2) + ' kg' + (extraApproved ? ' + approved extra ' + extraApproved.toFixed(2) + ' kg' : '') +
              (isSample ? ' (sample order — trial lots up to ' + SAMPLE_TRIAL_MIN_KG + ' kg are always allowed)' : '') +
              ', already on PO: ' + already.toFixed(2) + ' kg, balance: ' + Math.max(0, effectiveBalance).toFixed(2) + ' kg. Use \'Request Extra\' on the Yarn PO tab if more is genuinely needed.'
          };
        }
      }
    }

    var poNo = nextId('YPO-', SHEETS.YARN_PO, 1);
    var bagWeight = Number(d.bagWeight) || 0;
    var noOfBags = Number(d.noOfBags) || (bagWeight > 0 ? Math.round(qty / bagWeight) : 0);
    appendRowSafe_(SHEETS.YARN_PO, [
      poNo, d.poDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      d.poType || 'Yarn', d.joNumber, d.styleNumber, d.buyer, d.color, d.counts, d.yarnColor,
      rate, qty, rate * qty, d.supplier, d.deliveryAddress, d.gstNo,
      'Pending Approval', '', '', '', bagWeight, noOfBags, content, gsm,
      Number(d.gstPct) || 0, d.knittingUnitPhone || '', d.preparedBy || '',
      d.paymentTerms || '', d.deliveryDate || '',
      d.knittingUnitName || '', d.knittingUnitGstin || '', d.supplierAddress || '', knittingDia
    ]);
    return { success: true, poNumber: poNo, message: 'PO ' + poNo + ' submitted — awaiting approval. It cannot be printed until approved.' };
  } catch (err) {
    return { success: false, error: err.message };
  } finally { lock.releaseLock(); }
}

/*************** EMAIL THE PO TO THE SUPPLIER ***************
 * An approved PO goes to the supplier's email (from Supplier Master) with the PO PDF attached, the moment it
 * is approved. Approval never depends on it: if there is no address, the quota is used up, or Google has not
 * been given mail permission yet, the PO is still approved and the reason is shown / logged. Every attempt is
 * written to the PO EMAIL LOG sheet, and a PO is only auto-mailed once (use the Email button to send again). ***************/
var AUTO_EMAIL_PO = true;      // false = never mail automatically (the Email button on the Yarn PO register still works)
var PO_EMAIL_CC = '';          // optional: copy every PO email to this address (several: comma separated)
var PO_EMAIL_LOG_SHEET = 'PO EMAIL LOG';

function parseEmails_(raw) {
  var ok = [], bad = [];
  String(raw || '').split(/[,;\s]+/).forEach(function (e) {
    e = e.trim(); if (!e) return;
    if (/^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]{2,}$/.test(e)) { if (ok.indexOf(e.toLowerCase()) < 0) ok.push(e.toLowerCase()); } else bad.push(e);
  });
  return { ok: ok, bad: bad };
}
function poEmailLogSheet_() {
  var sh = SS.getSheetByName(PO_EMAIL_LOG_SHEET);
  if (!sh) {
    sh = SS.insertSheet(PO_EMAIL_LOG_SHEET);
    sh.getRange(1, 1, 1, 7).setValues([['PO No', 'Sent At', 'To', 'Result', 'Detail', 'How', 'Sent By']])
      .setFontWeight('bold').setBackground('#0B6E74').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }
  return sh;
}
function logPoEmail_(poNo, to, result, detail, how) {
  try { poEmailLogSheet_().appendRow([poNo, new Date(), to || '', result, detail || '', how || '', Session.getEffectiveUser().getEmail() || '']); } catch (e) { }
}
/* Latest email attempt per PO: { 'YPO-0027': { at, to, result } } — shown beside each PO on the Yarn PO register. */
function getPoEmailStatus() {
  var sh = SS.getSheetByName(PO_EMAIL_LOG_SHEET), out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  var tz = Session.getScriptTimeZone();
  sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues().forEach(function (r) {
    out[String(r[0]).trim()] = { at: r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'dd-MM HH:mm') : String(r[1]), to: String(r[2] || ''), result: String(r[3] || ''), detail: String(r[4] || '') };
  });
  return out;
}
function poAlreadyMailed_(poNo) {
  var sh = SS.getSheetByName(PO_EMAIL_LOG_SHEET);
  if (!sh || sh.getLastRow() < 2) return false;
  return sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().some(function (r) { return String(r[0]).trim() === String(poNo).trim() && r[3] === 'Sent'; });
}
/* The part every PO email shares: find the supplier's address in Supplier Master, check it and the daily
   limit, turn the document into a PDF, send it, and log what happened. o = { poNo (log key), supplier, html,
   fileName, subject, body, label ('PO' / 'Accessories PO'), how }.  Returns { success, message | error }. */
function mailPoDocument_(o) {
  var master = getSuppliers().filter(function (m) { return String(m['Supplier Name'] || '').trim().toLowerCase() === String(o.supplier || '').trim().toLowerCase(); })[0] || {};
  var em = parseEmails_(master['Email']);
  if (em.bad.length) { logPoEmail_(o.poNo, master['Email'], 'Failed', 'Invalid address on file: ' + em.bad.join(', '), o.how); return { success: false, error: 'The email on file for ' + o.supplier + ' does not look right (' + em.bad.join(', ') + '). Correct it in Masters → Suppliers.' }; }
  if (!em.ok.length) { logPoEmail_(o.poNo, '', 'No email', 'No email address on file for ' + o.supplier, o.how); return { success: false, error: 'No email address on file for ' + o.supplier + '. Add it in Masters → Suppliers, then use the Email button.' }; }
  if (MailApp.getRemainingDailyQuota() < 1) { logPoEmail_(o.poNo, em.ok.join(', '), 'Failed', 'Daily email limit reached', o.how); return { success: false, error: 'Google\'s daily email limit has been reached — try again tomorrow, or download the PDF and send it yourself.' }; }
  var pdf = Utilities.newBlob(o.html, 'text/html', o.fileName).getAs('application/pdf');
  pdf.setName(o.fileName);
  var opts = { to: em.ok.join(','), subject: o.subject, body: o.body, attachments: [pdf], name: 'Jiwin Apparels' };
  var cc = parseEmails_(PO_EMAIL_CC).ok; if (cc.length) opts.cc = cc.join(',');
  MailApp.sendEmail(opts);
  logPoEmail_(o.poNo, em.ok.join(', '), 'Sent', o.fileName, o.how);
  return { success: true, message: (o.label || 'PO') + ' ' + o.poNo + ' emailed to ' + o.supplier + ' (' + em.ok.join(', ') + ') with the PDF attached.' };
}
function poMailError_(poNo, e, how) {
  var msg = e && e.message ? e.message : String(e);
  logPoEmail_(poNo, '', 'Failed', msg, how);
  return { success: false, error: /permission|authori[sz]/i.test(msg)
    ? 'Google has not allowed this app to send mail yet. In the Apps Script editor run the function authorizeEmail once, accept the permission, then try again.'
    : 'Could not email PO ' + poNo + ': ' + msg };
}
var PO_MAIL_SIGNATURE_ = '\n\nRegards,\nJiwin Apparels\n10/227, SF No 4/470, Ganapathypalayam Road, Tirupur\nPhone: 9092998674';
var PO_MAIL_ASK_ = 'Kindly acknowledge this order and confirm the delivery date. Please quote the PO number and JO number(s) in your delivery challan and invoice.';
var PO_MAIL_DMY_ = function (d) { return String(d || '').slice(0, 10).split('-').reverse().join('-'); };

/* Yarn / Cora Fabric / Dyed Fabric PO. manual = true when a person pressed Email (sends again even if it was
   mailed before). Returns { success, message | error }. */
function emailPoToSupplier(poNo, manual) {
  poNo = String(poNo || '').trim();
  var how = manual ? 'Manual' : 'Auto';
  try {
    var lines = sheetToObjects(SHEETS.YARN_PO).filter(function (r) { return String(r['PO No']).trim() === poNo && r['Status'] !== 'Rejected'; });
    if (!lines.length) return { success: false, error: 'PO ' + poNo + ' not found.' };
    var html = buildPrintablePoHtml_(poNo);          // throws unless the PO is approved
    var typeName = String(lines[0]['PO Type'] || 'Yarn');
    var totalKg = lines.reduce(function (a, r) { return a + (Number(r['Qty KG']) || 0); }, 0);
    var jos = []; lines.forEach(function (r) { var j = String(r['JO Number'] || '').trim(); if (j && jos.indexOf(j) < 0) jos.push(j); });
    var date = PO_MAIL_DMY_(lines[0]['PO Date']);
    return mailPoDocument_({
      poNo: poNo, supplier: String(lines[0]['Supplier'] || '').trim(), html: html, how: how, label: 'PO',
      fileName: typeName.replace(/\s+/g, '_') + '_PO_' + poNo.replace(/[^\w-]+/g, '_') + '.pdf',
      subject: 'Purchase Order ' + poNo + ' — Jiwin Apparels',
      body: 'Dear Sir / Madam,\n\nPlease find attached our Purchase Order ' + poNo + (date ? ' dated ' + date : '') + ' for ' +
        totalKg.toLocaleString('en-IN', { maximumFractionDigits: 2 }) + ' kg (' + typeName + ')' + (jos.length ? ' against ' + jos.join(', ') : '') + '.\n\n' + PO_MAIL_ASK_ + PO_MAIL_SIGNATURE_
    });
  } catch (e) { return poMailError_(poNo, e, how); }
}
/* Used by approvePO: one short sentence about what happened with the email, or '' when nothing is to be said. */
function autoEmailPo_(poNo) {
  if (!AUTO_EMAIL_PO) return '';
  try {
    if (poAlreadyMailed_(poNo)) return '';
    var r = emailPoToSupplier(poNo, false);
    return r.success ? 'Emailed to the supplier.' : '⚠ Not emailed: ' + r.error;
  } catch (e) { return '⚠ Not emailed: ' + (e && e.message ? e.message : e); }
}

/* Accessories PO (JO-linked and General Purchase alike). A multi-line Accessories PO is several PO numbers in
   one group; the email, its log entry and the "already mailed" check all go by the FIRST number of the group,
   which is also the number shown in the lists. Any number from the group can be passed in. */
function accessoryPoForMail_(poNo) {
  var all = sheetToObjects(SHEETS.ACCESSORIES_PO);
  var apo = all.filter(function (r) { return String(r['PO No']).trim() === String(poNo).trim(); })[0];
  if (!apo) throw new Error('Accessories PO ' + poNo + ' not found.');
  var nos = apoGroupRows_(poNo).map(function (g) { return g.no; });
  var lines = nos.length > 1 ? all.filter(function (r) { return nos.indexOf(String(r['PO No'])) > -1; }) : [apo];
  var bad = lines.filter(function (r) { return r['Status'] !== 'Approved'; })[0];
  if (bad) throw new Error('Accessories PO ' + bad['PO No'] + ' is "' + bad['Status'] + '" — only an Approved PO can be mailed.');
  return { leader: String(lines[0]['PO No']), lines: lines, html: lines.length > 1 ? generateAccessoryPoMultiHtml_(lines) : generateAccessoryPoHtml(lines[0]) };
}
function emailAccessoryPoToSupplier(poNo, manual) {
  var how = manual ? 'Manual' : 'Auto', key = String(poNo || '').trim();
  try {
    var g = accessoryPoForMail_(poNo); key = g.leader;
    var items = g.lines.map(function (r) {
      var name = String(r['Trims / Accessories'] || r['General Description'] || '').trim() || 'item';
      var q = Number(r['Qty Ordered']) || 0;
      return name + (q ? ' (' + q.toLocaleString('en-IN', { maximumFractionDigits: 2 }) + ')' : '');
    });
    var shown = items.slice(0, 8).join(', ') + (items.length > 8 ? ' and ' + (items.length - 8) + ' more' : '');
    var jos = []; g.lines.forEach(function (r) { var j = String(r['JO Number'] || '').trim(); if (j && j.toUpperCase() !== 'GENERAL' && jos.indexOf(j) < 0) jos.push(j); });
    var date = PO_MAIL_DMY_(g.lines[0]['PO Date']);
    return mailPoDocument_({
      poNo: key, supplier: String(g.lines[0]['Supplier'] || '').trim(), html: g.html, how: how, label: 'Accessories PO',
      fileName: 'Accessories_PO_' + key.replace(/[^\w-]+/g, '_') + '.pdf',
      subject: 'Purchase Order ' + key + ' — Jiwin Apparels',
      body: 'Dear Sir / Madam,\n\nPlease find attached our Accessories Purchase Order ' + key + (date ? ' dated ' + date : '') + (jos.length ? ' against ' + jos.join(', ') : '') + '.\n\n' +
        'Items: ' + shown + '.\n\n' + PO_MAIL_ASK_ + PO_MAIL_SIGNATURE_
    });
  } catch (e) { return poMailError_(key, e, how); }
}
function autoEmailAccessoryPo_(poNo) {
  if (!AUTO_EMAIL_PO) return '';
  try {
    var key = accessoryPoForMail_(poNo).leader;
    if (poAlreadyMailed_(key)) return '';
    var r = emailAccessoryPoToSupplier(poNo, false);
    return r.success ? 'Emailed to the supplier.' : '⚠ Not emailed: ' + r.error;
  } catch (e) { return '⚠ Not emailed: ' + (e && e.message ? e.message : e); }
}

/* Run this ONCE from the Apps Script editor after deploying: it asks Google for permission to send mail. */
function authorizeEmail() {
  Logger.log('Email permission OK. Emails left today: ' + MailApp.getRemainingDailyQuota());
  return MailApp.getRemainingDailyQuota();
}
/* Add / change a supplier's email from the Suppliers list without opening the sheet. */
function setSupplierEmail(name, email) {
  var w = String(name || '').trim().toLowerCase();
  var rows = sheetToObjects(SHEETS.SUPPLIERS).filter(function (r) { return String(r['Supplier Name'] || '').trim().toLowerCase() === w; });
  if (!rows.length) return { success: false, error: 'Supplier "' + name + '" not found.' };
  var chk = parseEmails_(email);
  if (chk.bad.length) return { success: false, error: 'This email address does not look right: ' + chk.bad.join(', ') + '.' };
  setCell(SHEETS.SUPPLIERS, rows[0]._row, 'Email', chk.ok.join(', '));
  return { success: true, message: chk.ok.length ? name + ': email saved (' + chk.ok.join(', ') + ').' : name + ': email removed.', suppliers: getSuppliers() };
}

function approvePO(poNo, decision, approver, reason) {
  var all = sheetToObjects(SHEETS.YARN_PO);
  var matches = all.filter(function(r){return String(r['PO No']).trim()===String(poNo).trim();});
  if (!matches.length) return { success: false, error: 'PO not found: ' + poNo };
  var status = decision === 'APPROVE' ? 'Approved' : 'Rejected';
  var sh=SS.getSheetByName(SHEETS.YARN_PO);
  var headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var col={};headers.forEach(function(h,i){col[String(h).trim()]=i+1;});
  var now=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
  matches.forEach(function(r){
    var row=findRowIndex(SHEETS.YARN_PO,'PO No',poNo);
    // Find the physical row for this PO line by matching JO, count, quantity and rate.
    var vals=sh.getDataRange().getValues();
    for(var i=1;i<vals.length;i++){
      if(String(vals[i][col['PO No']-1]).trim()===String(poNo).trim() &&
         String(vals[i][col['JO Number']-1]).trim()===String(r['JO Number']||'').trim() &&
         String(vals[i][col['Counts']-1]).trim()===String(r['Counts']||'').trim() &&
         String(vals[i][col['Content']-1]||'').trim()===String(r['Content']||'').trim() &&
         String(vals[i][col['GSM']-1]||'').trim()===String(r['GSM']||'').trim() &&
         String(vals[i][col['Yarn Color']-1]||'').trim()===String(r['Yarn Color']||'').trim() &&
         Number(vals[i][col['Qty KG']-1])===Number(r['Qty KG']) &&
         Number(vals[i][col['Rate (Rs/kg)']-1])===Number(r['Rate (Rs/kg)']) &&
         String(vals[i][col['Status']-1])==='Pending Approval'){
        sh.getRange(i+1,col['Status']).setValue(status);
        sh.getRange(i+1,col['Approved By']).setValue(approver||'');
        sh.getRange(i+1,col['Approval Date']).setValue(now);
        sh.getRange(i+1,col['Rejection Reason']).setValue(decision==='REJECT'?(reason||''):'');
        break;
      }
    }
  });
  var mailNote = (status === 'Approved') ? autoEmailPo_(poNo) : '';
  return { success: true, message: 'PO ' + poNo + ' ('+matches.length+' line(s)) ' + status + '.' + (mailNote ? ' ' + mailNote : '') };
}

/*************** 6. YARN RECEIPT / INWARD ***************/
/**
 * Everything Yarn Receipt needs to autofill for a chosen PO: Mill Name (the
 * PO's Supplier), Counts, and — matched against Style Tracker via JO+Counts —
 * Yarn Type (Combed/Semi Combed) and Content (e.g. 100% Cotton).
 */
function getYarnPODetails(poNumber, joNumber, counts) {
  var po = sheetToObjects(SHEETS.YARN_PO).find(function (r) {
    return String(r['PO No']).trim() === String(poNumber).trim() &&
      (!joNumber || String(r['JO Number']||'').trim()===String(joNumber).trim()) &&
      (!counts || String(r['Counts']||'').trim()===String(counts).trim());
  });
  if (!po) return null;
  var content = po['Content'] || '';
  var gsm = po['GSM'] || '';
  var useGsm = po['PO Type'] !== 'Yarn';
  var colorForLookup = po['PO Type'] === 'Dyed Fabric' ? po['Color'] : '';
  var poYarnColor = String(po['Yarn Color'] || '').trim();
  var counted = getCountsForJO(po['JO Number'], colorForLookup, useGsm).find(function (c) {
    return c.counts === String(po['Counts'] || '').trim() && c.content === String(content).trim() && c.gsm === String(gsm).trim() && c.yarnColor === poYarnColor;
  });
  return {
    poNo: po['PO No'],
    joNumber: po['JO Number'],
    millName: po['Supplier'],
    counts: po['Counts'],
    yarnColor: po['Yarn Color'],
    qtyKg: Number(po['Qty KG']) || 0,
    status: po['Status'],
    fabricType: counted ? counted.fabricType : '',
    yarnType: counted ? counted.yarnType : '',
    content: content,
    gsm: gsm
  };
}

function getOpenYarnPOsForInward(joNumber) {
  var pos = sheetToObjects(SHEETS.YARN_PO);
  var inwards = sheetToObjects(SHEETS.YARN_INWARD);
  function norm(v) { return String(v == null ? '' : v).trim().toLowerCase(); }
  function lineKey(r) {
    return [r['Yarn PO No'] || r['PO No'] || r['Yarn PO'] || '', r['JO Number'] || '', r['Counts'] || '', r['Content'] || '', r['GSM'] || '']
      .map(norm).join('|');
  }
  var receivedByLine = {};
  inwards.forEach(function (r) {
    var key = lineKey(r);
    if (!key || key === '||||') return;
    receivedByLine[key] = (receivedByLine[key] || 0) + (Number(r['Received KG']) || 0);
  });
  return pos.filter(function (r) {
    if (String(r['JO Number'] || '').trim() !== String(joNumber || '').trim()) return false;
    if (String(r['Status'] || '').trim() !== 'Approved') return false;
    var ordered = Number(r['Qty KG']) || 0;
    var received = receivedByLine[lineKey(r)] || 0;
    return ordered - received > 0.009;
  }).map(function (r) {
    var ordered = Number(r['Qty KG']) || 0;
    var received = receivedByLine[lineKey(r)] || 0;
    var remaining = Math.max(0, ordered - received);
    var out = {};
    Object.keys(r).forEach(function (k) { out[k] = r[k]; });
    out['Received KG'] = Math.round(received * 1000) / 1000;
    out['Remaining KG'] = Math.round(remaining * 1000) / 1000;
    return out;
  });
}

/*************** AUTO YARN INWARD ***************
 * The mill ships the whole PO straight to the knitting unit named on the PO, so the inward does not have to be
 * typed line by line: for the chosen JO it receives each approved PO line's full remaining quantity, with the
 * knitting unit from the PO recorded as where it was delivered. Each line goes through saveYarnReceipt, so every
 * normal check applies (approved PO only, never more than the PO balance) and the typed-weight inward screen
 * keeps working exactly as before for part deliveries. line (optional) = { poNo, counts, content, gsm } limits it
 * to one PO line. ***************/
function planAutoYarnInward_(joNumber, line) {
  var norm = function (v) { return String(v == null ? '' : v).trim().toLowerCase(); };
  return getOpenYarnPOsForInward(joNumber).filter(function (r) {
    if (!line || !line.poNo) return true;
    return norm(r['PO No']) === norm(line.poNo) && (!line.counts || norm(r['Counts']) === norm(line.counts)) &&
      (!line.content || norm(r['Content']) === norm(line.content)) && (!line.gsm || norm(r['GSM']) === norm(line.gsm));
  }).map(function (r) {
    var kg = Number(r['Remaining KG']) || 0, bw = Number(r['Bag Weight (KG)']) || 0;
    return { poNo: r['PO No'], poType: r['PO Type'] || 'Yarn', supplier: r['Supplier'] || '', counts: r['Counts'] || '', content: r['Content'] || '', gsm: r['GSM'] || '',
      remainingKg: kg, bags: bw > 0 ? Math.round(kg / bw) : 0, knittingUnit: String(r['Knitting Unit Name'] || '').trim(), deliveryAddress: r['Delivery Address'] || '' };
  });
}
function previewAutoYarnInward(joNumber, line) {
  joNumber = String(joNumber || '').trim();
  if (!joNumber) return { success: false, error: 'Select the JO first.' };
  var lines = planAutoYarnInward_(joNumber, line);
  if (!lines.length) return { success: false, error: 'No approved PO with a balance left to receive for ' + joNumber + '.' };
  return { success: true, lines: lines, totalKg: Math.round(lines.reduce(function (a, l) { return a + l.remainingKg; }, 0) * 1000) / 1000 };
}
function autoYarnInward(joNumber, line, opts) {
  opts = opts || {};
  joNumber = String(joNumber || '').trim();
  var plan = previewAutoYarnInward(joNumber, line);
  if (!plan.success) return plan;
  var done = [], failed = [], totalKg = 0, units = [];
  plan.lines.forEach(function (l) {
    var r = saveYarnReceipt({
      yarnPoNo: l.poNo, joNumber: joNumber, poLineCounts: l.counts, poLineContent: l.content, poLineGsm: l.gsm,
      receivedKg: l.remainingKg, receivedBags: l.bags, date: opts.date || '', millDcInvoiceNo: opts.invoiceNo || '',
      whatsappTo: opts.whatsappTo || '', knittingUnit: l.knittingUnit,
      remarks: ['Auto inward — full PO quantity', opts.remarks || ''].filter(Boolean).join(' | ')
    });
    if (r.success) { done.push(r.inwardNo); totalKg += l.remainingKg; if (l.knittingUnit && units.indexOf(l.knittingUnit) < 0) units.push(l.knittingUnit); }
    else failed.push(l.poNo + ' ' + l.counts + ': ' + r.error);
  });
  if (!done.length) return { success: false, error: 'Nothing was recorded. ' + failed.join(' | ') };
  return {
    success: true, inwardNos: done, totalKg: Math.round(totalKg * 1000) / 1000, knittingUnits: units, failed: failed,
    message: 'Auto inward done for ' + joNumber + ': ' + done.length + ' line(s), ' + (Math.round(totalKg * 1000) / 1000).toLocaleString('en-IN') + ' kg received' +
      (units.length ? ', delivered to ' + units.join(', ') : '') + '. Inward no: ' + done.join(', ') + '.' + (failed.length ? ' ⚠ Not recorded: ' + failed.join(' | ') : '')
  };
}

function saveYarnReceipt(d) {
  return once_(d && d.submitToken, function () { return saveYarnReceipt_(d); });
}
function saveYarnReceipt_(d) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var poNo = String(d.yarnPoNo || '').trim();
    if (!poNo) return { success: false, error: 'Select an approved PO line.' };
    var pos = sheetToObjects(SHEETS.YARN_PO);
    var po = pos.find(function (r) {
      return String(r['PO No'] || '').trim() === poNo &&
        (!d.joNumber || String(r['JO Number'] || '').trim() === String(d.joNumber).trim()) &&
        (!d.poLineCounts || String(r['Counts'] || '').trim() === String(d.poLineCounts).trim()) &&
        (!d.poLineContent || String(r['Content'] || '').trim() === String(d.poLineContent).trim()) &&
        (!d.poLineGsm || String(r['GSM'] || '').trim() === String(d.poLineGsm).trim());
    });
    if (!po) return { success: false, error: 'Selected PO line was not found. Refresh the list and try again.' };
    if (String(po['Status'] || '').trim() !== 'Approved') {
      return { success: false, error: 'PO ' + poNo + ' is not Approved. Cannot receive against it.' };
    }

    var orderedKg = Number(po['Qty KG']) || 0;
    var inwardRows = sheetToObjects(SHEETS.YARN_INWARD);
    function normInw(v) { return String(v == null ? '' : v).trim().toLowerCase(); }
    function inwardLineKey(r) {
      return [r['Yarn PO No'] || r['PO No'] || r['Yarn PO'] || '', r['JO Number'] || '', r['Counts'] || '', r['Content'] || '', r['GSM'] || '']
        .map(normInw).join('|');
    }
    var targetLineKey = inwardLineKey({
      'Yarn PO No': poNo, 'JO Number': po['JO Number'] || d.joNumber || '',
      'Counts': po['Counts'] || d.poLineCounts || '', 'Content': po['Content'] || d.poLineContent || '',
      'GSM': po['GSM'] || d.poLineGsm || ''
    });
    var alreadyReceivedKg = inwardRows.reduce(function (sum, r) {
      return sum + (inwardLineKey(r) === targetLineKey ? (Number(r['Received KG']) || 0) : 0);
    }, 0);
    var remainingKg = Math.max(0, orderedKg - alreadyReceivedKg);
    var receivedKg = Number(d.receivedKg) || 0;
    if (remainingKg <= 0.009) {
      return { success: false, error: 'This PO is already fully inwarded. It has been removed from the inward list.' };
    }
    if (receivedKg <= 0) return { success: false, error: 'Received KG must be greater than zero.' };
    if (receivedKg - remainingKg > 0.009) {
      return { success: false, error: 'Received quantity exceeds the remaining PO balance of ' + remainingKg.toFixed(3) + ' KG. Please correct the quantity.' };
    }

    // All identifying details come from the approved PO / Style Tracker source.
    var details = getYarnPODetails(poNo, d.joNumber, d.poLineCounts);
    var millName = details ? details.millName : (po['Supplier'] || '');
    var counts = details ? details.counts : (po['Counts'] || '');
    var yarnType = details ? details.yarnType : '';
    var content = details ? details.content : (po['Content'] || '');
    var gsm = details ? details.gsm : (po['GSM'] || '');

    var knittingUnit = String(d.knittingUnit || po['Knitting Unit Name'] || '').trim();
    var inwNo = nextId('YINW-', SHEETS.YARN_INWARD, 1);
    appendRowSafe_(SHEETS.YARN_INWARD, [
      inwNo, d.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      poNo, po['JO Number'] || d.joNumber || '', millName, d.millDcInvoiceNo || '',
      Number(d.receivedBags) || 0, receivedKg, d.remarks || '',
      counts, yarnType, content, gsm, knittingUnit
    ]);

    var msgText = 'Jiwin Apparels - Yarn Receipt Confirmation\n' +
      'Inward No: ' + inwNo + '\nPO No: ' + poNo + '\nJO No: ' + (po['JO Number'] || d.joNumber || '-') +
      '\nMill: ' + (millName || '-') + '\nCounts: ' + (counts || '-') + (yarnType ? ' (' + yarnType + ')' : '') +
      '\nReceived: ' + receivedKg + ' KG (' + (d.receivedBags || 0) + ' bags)' + (knittingUnit ? '\nDelivered to: ' + knittingUnit : '') +
      '\nInvoice/DC: ' + (d.millDcInvoiceNo || '-');
    var waNumber = String(d.whatsappTo || '').replace(/[^0-9]/g, '');
    if (waNumber && !waNumber.startsWith('91') && waNumber.length === 10) waNumber = '91' + waNumber;
    var waLink = waNumber ? 'https://wa.me/' + waNumber + '?text=' + encodeURIComponent(msgText) : '';
    var balanceAfter = Math.max(0, remainingKg - receivedKg);
    return { success: true, message: 'Receipt ' + inwNo + ' recorded. Remaining PO balance: ' + balanceAfter.toFixed(3) + ' KG.', whatsappLink: waLink, inwardNo: inwNo, remainingKg: balanceAfter, fullyInwarded: balanceAfter <= 0.009 };
  } catch (e) {
    return { success: false, error: e && e.message ? e.message : String(e) };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function getYarnInward() { return sheetToObjects(SHEETS.YARN_INWARD); }

/*************** 7. PROCESS DC (process control from Style Tracker rates) ***************/
/* A save sent twice must be saved once. The screen gives every form a token that stays the same until that form
   has really been saved; a second request carrying the same token (a repeat click, a retry after a slow
   response, a double tap) is answered with the first result instead of being saved again. A failed save does
   not use the token up, so correcting the form and pressing the button again works normally. */
function once_(token, fn) {
  token = String(token || '').trim();
  if (!token) return fn();
  var cache = CacheService.getScriptCache(), key = 'once_' + token.slice(0, 80);
  var prev = cache.get(key);
  if (prev) {
    if (prev === 'pending') return { success: false, duplicate: true, error: 'This is already being saved — check the register in a moment. The repeat click was ignored.' };
    try { var p = JSON.parse(prev); p.duplicate = true; p.message = 'Already saved — ' + (p.message || '') + ' The repeat click was ignored, nothing was saved twice.'; return p; } catch (e) { }
  }
  cache.put(key, 'pending', 120);
  var res;
  try { res = fn(); } catch (e) { cache.remove(key); throw e; }
  if (res && res.success) {
    var keep = JSON.stringify(res);
    cache.put(key, keep.length < 90000 ? keep : JSON.stringify({ success: true, message: res.message || '' }), 900);
  } else cache.remove(key);
  return res;
}
function saveProcessDC(d) {
  return once_(d && d.submitToken, function () { return saveProcessDC_(d); });
}
function saveProcessDC_(d) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var isFinal = d.toProcess === JIWIN_INWARD_LABEL;
    if (!isFinal) {
      var allowed = getAllowedProcessesForJO(d.joNumber).map(function (x) { return x.process; });
      if (allowed.indexOf(d.toProcess) === -1) {
        return {
          success: false,
          error: '"' + d.toProcess + '" has no rate set for JO ' + d.joNumber + ' in Style Tracker. Add a process rate first, or choose "' + JIWIN_INWARD_LABEL + '" if this spec\'s processing is complete.'
        };
      }
    }

    var origin = CHAIN_ORIGINS.filter(function (o) { return o.sourceType === d.sourceType; })[0];
    var availableKg = (d.sourceType === 'DC') ? getDCChainBalance(d.sourceDcNo)
      : (origin ? getStockBalance(d.joNumber, d.counts, origin.stockType, origin.colorSpecific ? d.color : '', d.content, origin.useGsm ? d.gsm : '').available : 0);

    // A Process DC can only draw a previous Process DC belonging to the same JO.
    // Cross-JO material movement must be done through JO Transfer.
    if (d.sourceType === 'DC' && d.sourceDcNo) {
      var selectedSourceRow = sheetToObjects(SHEETS.PROCESS_DC).filter(function (r) { return r['DC No'] === d.sourceDcNo; })[0];
      if (!selectedSourceRow) return { success: false, error: 'Selected source DC was not found. Refresh the Process DC source list.' };
      if (String(selectedSourceRow['JO Number'] || '').trim() !== String(d.joNumber || '').trim()) {
        return { success: false, error: 'Source DC belongs to ' + String(selectedSourceRow['JO Number'] || '').trim() + ', not ' + String(d.joNumber || '').trim() + '. Use JO Transfer for cross-JO movement.' };
      }
      if (String(d.knittingDia || '').trim() && String(selectedSourceRow['Knitting Dia'] || '').trim() && String(d.knittingDia || '').trim().toUpperCase() !== String(selectedSourceRow['Knitting Dia'] || '').trim().toUpperCase()) {
        return { success: false, error: 'Source DC Dia ' + String(selectedSourceRow['Knitting Dia'] || '').trim() + ' does not match selected knitting Dia ' + String(d.knittingDia || '').trim() + '.' };
      }
    }

    // Goods already at Jiwin Inward going out again: user must say Process or Re-process,
    // and the DC waits for Management approval (not auto-approved).
    var fromJiwin = false;
    if (d.sourceType === 'DC' && d.sourceDcNo) {
      var srcRowObj = sheetToObjects(SHEETS.PROCESS_DC).filter(function (r) { return r['DC No'] === d.sourceDcNo; })[0];
      fromJiwin = !!(srcRowObj && srcRowObj['To Process'] === JIWIN_INWARD_LABEL);
    }
    var procType = String(d.processType || '').trim();
    if (fromJiwin) {
      if (procType !== 'Process' && procType !== 'Re-process') return { success: false, error: 'These goods are already at Jiwin Inward. Choose "Process" or "Re-process" for this DC.' };
      if (procType === 'Re-process' && !String(d.reprocessReason || '').trim()) return { success: false, error: 'Enter the reason for re-process (e.g. shade variation, GSM low).' };
    }

    var dcQtyKg = Number(d.dcQtyKg) || 0;
    if (dcQtyKg <= 0) return { success: false, error: 'DC Quantity must be greater than 0.' };

    // Fabric type + locked quantity: a DC must name the Style Tracker fabric it carries, and can never
    // send more of that fabric to this process than the JO needs (minus what was already sent).
    var fabricType = dcxS_(d.fabricType);
    var knittingDia = dcxS_(d.knittingDia);
    var isReprocessDc = fromJiwin && procType === 'Re-process';
    if (dcJoFabricTypes_(d.joNumber).length && !isReprocessDc) {
      if (!fabricType) return { success: false, error: 'Select the Fabric Type for this DC.' };
      var fabBal = getDcFabricBalance(d.joNumber, fabricType, d.color || '', d.toProcess, knittingDia);
      if (!fabBal.found) return { success: false, error: 'Fabric type \"' + fabricType + '\" is not in Style Tracker for ' + d.joNumber + (d.color ? ' / ' + d.color : '') + '.' };
      if (dcQtyKg > fabBal.balanceKg + 0.01) {
        return {
          success: false,
          error: 'Fabric ' + fabBal.fabricType + (knittingDia ? ' / Dia ' + knittingDia : '') + (d.color ? ' / ' + d.color : '') + ': ' + dcQtyKg + ' KG is more than the locked quantity. Required ' +
            fabBal.requiredKg.toFixed(2) + ' KG, already sent to ' + d.toProcess + ' ' + fabBal.issuedKg.toFixed(2) + ' KG, balance ' + Math.max(0, fabBal.balanceKg).toFixed(2) + ' KG.'
        };
      }
    }
    if (dcQtyKg > availableKg + 0.01) {
      return {
        success: false,
        error: 'DC Issue Qty (' + dcQtyKg + ' KG) exceeds what is available from this source (' + availableKg.toFixed(2) + ' KG). Refresh the source list — someone may have used part of it already.'
      };
    }

    // Issuing part of a source is normal (several Knitting DCs from one Yarn Store lot):
    // the rest stays available for the next DC. It is only LOSS when the user closes a
    // processed lot ("last DC from this source") - never for Yarn / Cora / Dyed store stock.
    var remainKg = Math.max(0, availableKg - dcQtyKg);
    var closeLot = d.sourceType === 'DC' && (d.closeLot === true || d.closeLot === 'on' || d.closeLot === 'true');
    var lossKg = closeLot ? remainKg : 0;
    var lossPct = closeLot && availableKg > 0 ? (lossKg / availableKg * 100) : 0;
    var fromProcess = d.sourceType === 'DC' ? (d.fromProcess || '') : (origin ? origin.fromProcess : '');
    var fromSupplier = d.sourceType === 'DC' ? (d.fromSupplier || '') : '';
    var previousDcNo = d.sourceType === 'DC' ? (d.sourceDcNo || '') : '';

    var dcCode = getProcessMaps().codeMap[d.toProcess] || (isFinal ? 'JI' : 'DC');
    var dcNo = d._dcNo || nextIdForPrefix(SHEETS.PROCESS_DC, 1, dcCode);
    var dcRowNo = appendRowSafe_(SHEETS.PROCESS_DC, [
      dcNo, d.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      d.joNumber, d.styleNo || '', d.color || '', fromProcess, fromSupplier, d.toProcess, d.toSupplier || '',
      d.lotNo || '', previousDcNo, availableKg, availableKg, dcQtyKg, Number(d.rolls) || 0,
      remainKg, (fromJiwin ? '[' + procType.toUpperCase() + (d.reprocessReason ? ': ' + String(d.reprocessReason).trim() : '') + '] ' : '') + (d.remarks || ''),
      fromJiwin ? 'Pending Approval' : 'Approved', fromJiwin ? '' : 'Auto', fromJiwin ? '' : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      d.counts || '', d.content || '', d.gsm || '', lossKg, lossPct, d.vehicleNo || '', '', '', d.supplierDcNo || ''
    ]);
    if (fromJiwin) {
      var hpt = dcxEnsureHeaders_();
      hpt.sheet.getRange(dcRowNo, hpt.col['Process Type'] + 1).setValue(procType);
    }
    if (fabricType) dcFabStamp_(dcRowNo, d.joNumber, fabricType, d.color || '', knittingDia);
    var photoUrl = d.photoUrl || (d.photoB64 ? dcPhotoSave_(d.photoB64, d.photoMime, d.photoName, dcNo) : '');
    if (photoUrl) {
      var hp = dcxEnsureHeaders_();
      hp.sheet.getRange(dcRowNo, hp.col['Supplier DC Photo'] + 1).setValue(photoUrl);
    }
    if (d._dcGroup && d._dcGroup !== '__FIRST__') {   // one printed DC with several fabric / colour lines
      var hx = dcxEnsureHeaders_();
      hx.sheet.getRange(dcRowNo, hx.col['DC Group'] + 1).setValue(d._dcGroup);
    }
    if (closeLot) {
      var srcRow = findRowIndex(SHEETS.PROCESS_DC, 'DC No', d.sourceDcNo);
      if (srcRow > 0) {
        setCell(SHEETS.PROCESS_DC, srcRow, 'Balance Status', 'Closed');
        setCell(SHEETS.PROCESS_DC, srcRow, 'Balance Note', 'Lot closed on ' + dcNo + ' - ' + lossKg.toFixed(2) + ' kg process loss');
      }
    }
    var lossNote = closeLot && lossKg > 0.01 ? (' Lot closed - process loss ' + lossKg.toFixed(2) + ' kg (' + lossPct.toFixed(1) + '%).')
      : (remainKg > 0.01 ? (' Balance ' + remainKg.toFixed(2) + ' kg is still available for the next DC.') : '');
    if (fromJiwin) return { success: true, dcNo: dcNo, pending: true, message: procType + ' DC ' + dcNo + ' sent to Management for approval (goods from Jiwin Inward).' };
    return { success: true, dcNo: dcNo, message: 'Process DC ' + dcNo + ' recorded and approved.' + lossNote };
  } catch (err) {
    return { success: false, error: err.message };
  } finally { lock.releaseLock(); }
}

function approveDC(dcNo, decision, approver, reason) {
  var rowIdx = findRowIndex(SHEETS.PROCESS_DC, 'DC No', dcNo);
  if (rowIdx === -1) return { success: false, error: 'DC not found: ' + dcNo };
  var status = decision === 'APPROVE' ? 'Approved' : 'Rejected';
  setCell(SHEETS.PROCESS_DC, rowIdx, 'Status', status);
  setCell(SHEETS.PROCESS_DC, rowIdx, 'Approved By', approver || '');
  setCell(SHEETS.PROCESS_DC, rowIdx, 'Approval Date', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  return { success: true, message: 'DC ' + dcNo + ' ' + status + (reason ? ' (' + reason + ')' : '') + '.' };
}

function getProcessDCs() {
  var rows = sheetToObjects(SHEETS.PROCESS_DC);
  try {
    var m = {};
    var ext = dcxReadDc_();
    ext.forEach(function (r) { if (r['Supplier DC Photo']) m[r['DC No']] = r['Supplier DC Photo']; });
    // The fabric belongs to the ROW. Looking it up by DC number put the LAST line's fabric on every line of a
    // multi-line DC that shares one number (S/J + Rib + Neck Tape all showed as "NECK TAPE").
    var sameRows = ext.length === rows.length;
    var perDc = {};
    ext.forEach(function (r) { if (r['Fabric type']) perDc[r['DC No']] = r['Fabric type']; });
    rows.forEach(function (r, i) {
      r['Supplier DC Photo'] = m[r['DC No']] || '';
      r['Fabric type'] = (sameRows ? ext[i]['Fabric type'] : perDc[r['DC No']]) || '';
    });
  } catch (e) { }
  return rows;
}

/*************** SUPPLIER DC PHOTO ***************/
function dcPhotoFolder_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('DC_PHOTO_FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { } }
  var f = DriveApp.createFolder('JIWIN_Supplier_DC_Photos');
  p.setProperty('DC_PHOTO_FOLDER', f.getId());
  return f;
}
function dcPhotoSave_(b64, mime, name, ref) {
  if (!b64) return '';
  if (b64.length > 8000000) throw new Error('Supplier DC photo is too large (max 5 MB).');
  var file = dcPhotoFolder_().createFile(Utilities.newBlob(Utilities.base64Decode(b64), mime || 'image/jpeg',
    String(ref || 'DC').replace(/[^\w-]/g, '_') + '_' + String(name || 'supplier_dc.jpg').replace(/[\\/:*?"<>|]+/g, '_')));
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { }
  return 'https://drive.google.com/file/d/' + file.getId() + '/view';
}
/* add / replace the supplier's DC photo later (all rows of that DC) */
function addDcPhoto(dcNo, b64, mime, name) {
  try {
    var url = dcPhotoSave_(b64, mime, name, dcNo);
    var h = dcxEnsureHeaders_(), all = dcxReadDc_(), row = all.filter(function (r) { return r['DC No'] === dcNo; })[0];
    if (!row) return { success: false, error: 'DC ' + dcNo + ' not found.' };
    var grp = String(row['DC Group'] || '').trim() || dcNo, n = 0;
    all.forEach(function (r) {
      if (r['DC No'] === dcNo || (String(r['DC Group'] || '').trim() || r['DC No']) === grp) { h.sheet.getRange(r._row, h.col['Supplier DC Photo'] + 1).setValue(url); n++; }
    });
    return { success: true, url: url, message: 'Supplier DC photo saved on ' + grp + (n > 1 ? ' (' + n + ' lines)' : '') + '.' };
  } catch (err) { return { success: false, error: err.message }; }
}

/*************** 8b. KNITTING DC + DYEING DC (from Style Tracker) ***************
 * Knitting DC : Yarn Store -> Knitting. Fabric lines come from Style Tracker
 *   (all fabrics). Lycra is issued as its own yarn line. Qty per fabric line
 *   <= Style Tracker balance, and yarn <= Yarn Store stock (getStockBalance).
 * Dyeing DC   : Knitting -> Dyeing, colour-wise from Style Tracker. Grey is
 *   drawn from Approved Knitting DCs' un-forwarded balance (Previous DC No),
 *   so stock, chain, bill passing and costing keep working as before.
 * One printed DC = one "DC Group" (e.g. K-0012); each fabric/yarn line is its
 *   own row (K-0012, K-0012-B, K-0012-C ...) so the DC chain stays per lot.
 *******************************************************************************/
var DCX_EXTRA_HEADERS = ['DC Group', 'Fabric type', 'Design', 'Gauge', 'Loop length', 'Knitting Dia',
  'Finishing dia', 'Line Key', 'Grey Key', 'Yarn Kind', 'Bags', 'Cones', 'Supplier DC Photo', 'Process Type'];
var DCX_CHECK_LYCRA_STOCK = false;  // true = lycra must be received on a Yarn PO before knitting (not bought separately yet)
var DCX_TOLERANCE_KG = 0;           // generic process tolerance
var DCX_BODY_EXTRA_KG = 10;          // up to 10 kg extra may be added to a BODY fabric line per colour + knitting dia
var DCX_FROM = {
  name: 'JIWIN APPARELS',
  address: 'D.NO.10/227, SF No.4/470, Ganapathypalayam Road, Kunnangalpalaiyam, Veerapandi Post, 641605-TIRUPUR, TAMILNADU, INDIA',
  gst: '33BOWPS1618P2Z2', pan: 'BOWPS1618P'
};

function dcxS_(v) { return String(v == null ? '' : v).trim(); }
function dcxN_(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  var n = parseFloat(String(v).replace(/[,%\s]/g, ''));
  return isFinite(n) ? n : 0;
}
function dcxR3_(n) { return Math.round((Number(n) || 0) * 1000) / 1000; }
function dcxToday_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function dcxDmy_(v) {
  var s = dcxS_(v), m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
  return m ? m[3] + '/' + m[2] + '/' + m[1] : s;
}
function dcxE_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

/* Process DC with every column (base HEADERS + extras) */
function dcxReadDc_() {
  var sh = SS.getSheetByName(SHEETS.PROCESS_DC);
  var lastRow = sh.getLastRow(), lastCol = Math.max(sh.getLastColumn(), 1);
  if (lastRow < 2) return [];
  var base = HEADERS[SHEETS.PROCESS_DC];
  var head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h, i) { return i < base.length ? base[i] : dcxS_(h); });
  return sh.getRange(2, 1, lastRow - 1, lastCol).getValues().map(function (r, i) {
    var o = { _row: i + 2 };
    head.forEach(function (h, j) { if (h) o[h] = (r[j] instanceof Date) ? formatDateVal(r[j]) : r[j]; });
    return o;
  });
}

/* Adds the extra DC columns after the standard ones (never moves existing columns) */
function dcxEnsureHeaders_() {
  var sh = SS.getSheetByName(SHEETS.PROCESS_DC);
  var base = HEADERS[SHEETS.PROCESS_DC];
  var lastCol = Math.max(sh.getLastColumn(), base.length);
  if (sh.getMaxColumns() < lastCol) sh.insertColumnsAfter(sh.getMaxColumns(), lastCol - sh.getMaxColumns());
  var head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(dcxS_);
  var col = {}, changed = false;
  base.forEach(function (h, i) { col[h] = i; });
  DCX_EXTRA_HEADERS.forEach(function (h) {
    var idx = -1;
    for (var j = base.length; j < head.length; j++) if (head[j] === h) { idx = j; break; }
    if (idx < 0) { head.push(h); idx = head.length - 1; changed = true; }
    col[h] = idx;
  });
  if (sh.getMaxColumns() < head.length) sh.insertColumnsAfter(sh.getMaxColumns(), head.length - sh.getMaxColumns());
  if (changed) {
    sh.getRange(1, base.length + 1, 1, head.length - base.length).setValues([head.slice(base.length)])
      .setFontWeight('bold').setBackground('#1B2A4A').setFontColor('#FFFFFF');
  }
  return { sheet: sh, width: head.length, col: col };
}

function dcxWriteRows_(objs) {
  var h = dcxEnsureHeaders_();
  var rows = objs.map(function (o) {
    var r = [];
    for (var i = 0; i < h.width; i++) r.push('');
    Object.keys(o).forEach(function (k) { if (h.col.hasOwnProperty(k)) r[h.col[k]] = o[k]; });
    return r;
  });
  h.sheet.getRange(h.sheet.getLastRow() + 1, 1, rows.length, h.width).setValues(rows);
  SpreadsheetApp.flush();
}

/* K-0012, K-0012-B, K-0012-C ... (letters so nextIdForPrefix keeps counting from 0012) */
function dcxRowId_(base, i) {
  if (i === 0) return base;
  var s = '', n = i + 1; // 1 -> B
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return base + '-' + s;
}

function dcxSpec_(r) {
  return {
    fabric: dcxS_(r['Fabric type']), counts: dcxS_(r['Counts']), content: dcxS_(r['Content']),
    design: (dcxS_(r['Design']) || 'SOLID').toUpperCase(), gsm: dcxS_(r['GSM']),
    gg: dcxS_(r['Gauge']), ll: dcxS_(r['Loop length']),
    kDia: dcxS_(r['Knitting Dia']), fDia: dcxS_(r['Finishing dia']),
    yarnColor: dcxS_(r['Yarn color']), yarnType: dcxS_(r['Yarn Type'])
  };
}
function dcxGreyKey_(s) {
  return [s.fabric, s.counts, s.content, s.design, s.gsm, s.gg, s.ll, s.kDia, s.fDia].join('|').toUpperCase();
}
function dcxIsLycra_(s) { return /^(LYCRA|SPANDEX|ELASTANE)\b/i.test(dcxS_(s)); }

/* Main yarn kg + lycra kg of one Style Tracker row */
function dcxLycra_(r) {
  var req = dcxN_(r['Fabric Required (kg)']);
  var counts = dcxS_(r['Counts']);
  var parts = counts.split(/\s*[,+\/]\s*/).map(dcxS_).filter(Boolean);
  var mainParts = parts.filter(function (p) { return !dcxIsLycra_(p); });
  if ((parts.length && !mainParts.length) || (/^(LYCRA|SPANDEX|ELASTANE)$/i.test(dcxS_(r['Fabric type'])) && !mainParts.length)) {
    return { pure: true, mainKg: 0, lycraKg: req };
  }
  var mCol = dcxS_(r['Main Fiber Reqd (kg)']), lCol = dcxS_(r['Lycra Reqd (kg)']);
  var hasLy = /^(Y|YES|TRUE)$/i.test(dcxS_(r['Has Lycra']));
  var pct = dcxN_(r['Lycra %']);
  if (pct > 0 && pct < 1) pct = pct * 100;
  var ly = 0;
  if (mCol !== '' || lCol !== '') ly = dcxN_(lCol);
  else if (hasLy && pct > 0) ly = req * pct / 100;
  ly = Math.min(ly, req);
  return { pure: false, mainKg: req - ly, lycraKg: ly };
}

function dcxJoRows_(jo) {
  return getStyleTracker().filter(function (r) { return dcxS_(r['JO Number']) === dcxS_(jo); });
}

/* Lycra yarn spec as bought on Yarn PO: the JO's lycra-only row if any, else "LYCRA" */
function dcxLycraStockSpec_(stRows) {
  var p = stRows.filter(function (r) { return dcxLycra_(r).pure; })[0];
  return p ? { counts: dcxS_(p['Counts']), content: dcxS_(p['Content']) } : { counts: 'LYCRA', content: '' };
}

/* Knitting lines = Style Tracker rows merged by grey spec (all colours) */
function dcxKnitLines_(stRows) {
  var order = [], map = {};
  stRows.forEach(function (r) {
    var li = dcxLycra_(r);
    if (li.pure) return;
    var s = dcxSpec_(r), key = dcxGreyKey_(s);
    if (!map[key]) { s.key = key; s.parts = []; s.colors = []; s.required = 0; s.mainKg = 0; s.lycraKg = 0; map[key] = s; order.push(key); }
    var L = map[key], part = dcxS_(r['Part']), color = dcxS_(r['Color']);
    if (part && L.parts.indexOf(part) < 0) L.parts.push(part);
    if (color && L.colors.indexOf(color) < 0) L.colors.push(color);
    L.required += dcxN_(r['Fabric Required (kg)']);
    L.mainKg += li.mainKg; L.lycraKg += li.lycraKg;
  });
  return order.map(function (k) {
    var L = map[k], t = L.mainKg + L.lycraKg;
    L.lycraShare = t ? L.lycraKg / t : 0;
    return L;
  });
}

function dcxProcessAllowed_(jo, proc) {
  return getAllowedProcessesForJO(jo).map(function (x) { return x.process; }).indexOf(proc) > -1;
}

/* ---------------- KNITTING DC ---------------- */
function getKnitDcForm(joNumber) {
  var jo = dcxS_(joNumber);
  var st = dcxJoRows_(jo);
  if (!st.length) throw new Error('JO "' + jo + '" not found in Style Tracker.');
  var lines = dcxKnitLines_(st);
  var dcRows = dcxReadDc_();
  var knitted = {};
  dcRows.forEach(function (r) {
    if (dcxS_(r['JO Number']) !== jo || r['To Process'] !== 'Knitting' || r['Status'] === 'Rejected') return;
    var k = dcxS_(r['Line Key']).toUpperCase();
    if (k) knitted[k] = (knitted[k] || 0) + dcxN_(r['DC Issue Qty KG']);
  });
  var ly = dcxLycraStockSpec_(st);
  var stock = {};
  var addStock = function (counts, content, label) {
    var k = counts + '||' + content;
    if (!stock[k]) stock[k] = { label: label, available: dcxR3_(getStockBalance(jo, counts, 'Yarn', '', content, '').available) };
    return k;
  };
  var out = lines.map(function (L) {
    var done = knitted[L.key] || 0;
    var mainKey = addStock(L.counts, L.content, L.counts + (L.content ? ' (' + L.content + ')' : ''));
    var lyKey = L.lycraKg > 0 ? addStock(ly.counts, ly.content, ly.counts + (ly.content ? ' (' + ly.content + ')' : '')) : '';
    return {
      key: L.key, fabric: L.fabric, parts: L.parts.join(', '), colors: L.colors.join(', '), counts: L.counts,
      content: L.content, design: L.design, gsm: L.gsm, gg: L.gg, ll: L.ll, kDia: L.kDia, fDia: L.fDia,
      required: dcxR3_(L.required), knitted: dcxR3_(done), balance: dcxR3_(Math.max(0, L.required - done)),
      lycraShare: L.lycraShare, lycraPct: Math.round(L.lycraShare * 1000) / 10, mainKey: mainKey, lyKey: lyKey
    };
  });
  return {
    jo: jo, styleNo: dcxS_(st[0]['Style Number']), buyer: dcxS_(st[0]['Buyer Name']),
    allowed: dcxProcessAllowed_(jo, 'Knitting'), lines: out, stock: stock, lyCheck: DCX_CHECK_LYCRA_STOCK,
    suppliers: getSuppliersForProcess('Knitting')
  };
}

/* d = { joNumber, date, toSupplier, supplierDcNo, vehicleNo, remarks, lines:[{key, qty, bags, cones}] } */
function saveKnitDc(d) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var jo = dcxS_(d.joNumber), sup = dcxS_(d.toSupplier);
    if (!jo) return { success: false, error: 'Select a JO.' };
    if (!sup) return { success: false, error: 'Select the Knitting unit (To Supplier).' };
    if (!dcxProcessAllowed_(jo, 'Knitting')) {
      return { success: false, error: '"Knitting" has no rate set for JO ' + jo + ' in Style Tracker. Add the Knitting rate first.' };
    }
    var f = getKnitDcForm(jo), byKey = {};
    f.lines.forEach(function (L) { byKey[L.key] = L; });
    var st = dcxJoRows_(jo), ly = dcxLycraStockSpec_(st), spec = {};
    dcxKnitLines_(st).forEach(function (L) { spec[L.key] = L; });

    var lines = (d.lines || []).filter(function (x) { return dcxN_(x.qty) > 0; });
    if (!lines.length) return { success: false, error: 'Enter Issue Kg for at least one fabric.' };

    var tol = DCX_TOLERANCE_KG + 0.0005, use = {}, stockLeft = {};
    Object.keys(f.stock).forEach(function (k) { stockLeft[k] = f.stock[k].available; });
    var plan = [];
    for (var i = 0; i < lines.length; i++) {
      var x = lines[i], L = byKey[dcxS_(x.key).toUpperCase()];
      if (!L) return { success: false, error: 'Fabric not found in Style Tracker for ' + jo + '. Reload the JO.' };
      var q = dcxR3_(dcxN_(x.qty));
      if (q > L.balance + tol) {
        return { success: false, error: 'DC not saved. ' + L.fabric + ' ' + L.kDia + '/' + L.fDia + ': ' + q.toFixed(3) +
          ' kg is more than the Style Tracker balance ' + L.balance.toFixed(3) + ' kg.' };
      }
      var lyQ = dcxR3_(q * L.lycraShare), mainQ = dcxR3_(q - lyQ);
      use[L.mainKey] = (use[L.mainKey] || 0) + mainQ;
      if (lyQ > 0 && DCX_CHECK_LYCRA_STOCK) use[L.lyKey] = (use[L.lyKey] || 0) + lyQ;
      plan.push({ L: L, x: x, mainQ: mainQ, lyQ: lyQ });
    }
    var bad = Object.keys(use).filter(function (k) { return use[k] > (f.stock[k] ? f.stock[k].available : 0) + tol; })[0];
    if (bad) {
      return { success: false, error: 'DC not saved. Yarn ' + f.stock[bad].label + ': need ' + use[bad].toFixed(3) +
        ' kg but only ' + f.stock[bad].available.toFixed(3) + ' kg is in Yarn Store.' +
        (bad === (ly.counts + '||' + ly.content) ? ' (Lycra must be received on a Yarn PO first.)' : '') };
    }

    var code = getProcessMaps().codeMap['Knitting'] || 'K';
    var base = nextIdForPrefix(SHEETS.PROCESS_DC, 1, code);
    var date = d.date || dcxToday_(), today = dcxToday_(), rows = [];
    var addRow = function (L, kind, counts, content, qty, bags, cones) {
      var k = counts + '||' + content;
      var avail = stockLeft.hasOwnProperty(k) ? stockLeft[k] : dcxR3_(getStockBalance(jo, counts, 'Yarn', '', content, '').available);
      stockLeft[k] = dcxR3_(avail - qty);
      rows.push({
        'DC No': dcxRowId_(base, rows.length), 'Date': date, 'JO Number': jo, 'Style Number': f.styleNo, 'Color': '',
        'From Process': YARN_STORE_LABEL, 'From Supplier': '', 'To Process': 'Knitting', 'To Supplier': sup,
        'Lot No': '', 'Previous DC No': '', 'Previous DC Qty KG': avail, 'Available KG': avail, 'DC Issue Qty KG': qty,
        'Rolls': 0, 'Balance KG': dcxR3_(avail - qty), 'Remarks': dcxS_(d.remarks), 'Status': 'Approved',
        'Approved By': 'Auto', 'Approval Date': today, 'Counts': counts, 'Content': content, 'GSM': L.gsm,
        'Loss KG': 0, 'Loss %': 0, 'Vehicle No': dcxS_(d.vehicleNo), 'Supplier DC No': dcxS_(d.supplierDcNo),
        'DC Group': base, 'Fabric type': L.fabric, 'Design': L.design, 'Gauge': L.gg, 'Loop length': L.ll,
        'Knitting Dia': L.kDia, 'Finishing dia': L.fDia, 'Line Key': L.key, 'Grey Key': L.key,
        'Yarn Kind': kind, 'Bags': bags, 'Cones': cones
      });
    };
    plan.forEach(function (p) {
      var sp = spec[p.L.key];
      addRow(sp, 'Main', sp.counts, sp.content, p.mainQ, dcxN_(p.x.bags), dcxN_(p.x.cones));
      if (p.lyQ > 0) addRow(sp, 'Lycra', ly.counts, ly.content, p.lyQ, 0, 0);
    });
    dcxWriteRows_(rows);
    var total = plan.reduce(function (s, p) { return s + p.mainQ + p.lyQ; }, 0);
    return { success: true, dcNo: base, message: 'Knitting DC ' + base + ' created for ' + total.toFixed(3) + ' kg.' };
  } catch (err) {
    return { success: false, error: err.message };
  } finally { lock.releaseLock(); }
}

/* ---------------- DYEING DC ---------------- */
function dcxDyeContext_(jo) {
  var st = dcxJoRows_(jo);
  if (!st.length) throw new Error('JO "' + jo + '" not found in Style Tracker.');
  var order = [], map = {};
  st.forEach(function (r) {
    if (dcxLycra_(r).pure) return;
    var s = dcxSpec_(r), g = dcxGreyKey_(s), color = dcxS_(r['Color']) || '-';
    var key = (color + '#' + g).toUpperCase();
    if (!map[key]) { s.key = key; s.greyKey = g; s.color = color; s.parts = []; s.required = 0; map[key] = s; order.push(key); }
    var part = dcxS_(r['Part']);
    if (part && map[key].parts.indexOf(part) < 0) map[key].parts.push(part);
    map[key].required += dcxN_(r['Fabric Required (kg)']);
  });
  var lines = order.map(function (k) { return map[k]; });

  var dc = dcxReadDc_(), fwd = {}, dyed = {};
  dc.forEach(function (r) {
    if (r['Status'] === 'Rejected') return;
    var p = dcxS_(r['Previous DC No']);
    if (p) fwd[p] = (fwd[p] || 0) + dcxN_(r['DC Issue Qty KG']);
    if (dcxS_(r['JO Number']) === jo && r['To Process'] === 'Dyeing') {
      var k = dcxS_(r['Line Key']).toUpperCase();
      if (k) dyed[k] = (dyed[k] || 0) + dcxN_(r['DC Issue Qty KG']);
    }
  });
  // Grey sources per grey key: Approved Knitting DC rows with balance.
  // Old Knitting DCs (no Line Key) are matched by Counts + Content + GSM.
  var sources = {}, bal = {};
  var knit = dc.filter(function (r) {
    return dcxS_(r['JO Number']) === jo && r['To Process'] === 'Knitting' && (r['Status'] === 'Approved' || r['Status'] === 'Billed');
  });
  knit.forEach(function (r) { bal[r['DC No']] = dcxR3_(dcxN_(r['DC Issue Qty KG']) - (fwd[r['DC No']] || 0)); });
  lines.forEach(function (L) {
    if (sources[L.greyKey]) return;
    sources[L.greyKey] = knit.filter(function (r) {
      if (bal[r['DC No']] <= 0.0005) return false;
      var lk = dcxS_(r['Line Key']).toUpperCase();
      if (lk) return lk === L.greyKey;
      return dcxS_(r['Counts']) === L.counts && dcxS_(r['Content']) === L.content &&
        (!dcxS_(r['GSM']) || dcxS_(r['GSM']) === L.gsm);
    }).map(function (r) { return { dcNo: r['DC No'], supplier: dcxS_(r['To Supplier']), issue: dcxN_(r['DC Issue Qty KG']) }; });
  });
  return { st: st, lines: lines, dyed: dyed, sources: sources, bal: bal };
}

function getDyeDcForm(joNumber) {
  var jo = dcxS_(joNumber), c = dcxDyeContext_(jo), grey = {};
  c.lines.forEach(function (L) {
    if (grey[L.greyKey]) return;
    var avail = c.sources[L.greyKey].reduce(function (s, x) { return s + c.bal[x.dcNo]; }, 0);
    grey[L.greyKey] = {
      label: L.fabric + ' ' + L.kDia + '/' + L.fDia + (L.gsm ? ' ' + L.gsm + ' GSM' : ''),
      available: dcxR3_(avail), from: c.sources[L.greyKey].map(function (x) { return x.dcNo; }).join(', ')
    };
  });
  return {
    jo: jo, styleNo: dcxS_(c.st[0]['Style Number']), buyer: dcxS_(c.st[0]['Buyer Name']),
    allowed: dcxProcessAllowed_(jo, 'Dyeing'), grey: grey, suppliers: getSuppliersForProcess('Dyeing'),
    lines: c.lines.map(function (L) {
      var done = c.dyed[L.key] || 0;
      return {
        key: L.key, greyKey: L.greyKey, color: L.color, fabric: L.fabric, parts: L.parts.join(', '),
        design: L.design, counts: L.counts, content: L.content, gsm: L.gsm, gg: L.gg, ll: L.ll,
        kDia: L.kDia, fDia: L.fDia, required: dcxR3_(L.required), dyed: dcxR3_(done),
        balance: dcxR3_(Math.max(0, L.required - done))
      };
    })
  };
}

/* d = { joNumber, date, toSupplier, supplierDcNo, vehicleNo, remarks, lines:[{key, qty, rolls}] } */
function saveDyeDc(d) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var jo = dcxS_(d.joNumber), sup = dcxS_(d.toSupplier);
    if (!jo) return { success: false, error: 'Select a JO.' };
    if (!sup) return { success: false, error: 'Select the Dyeing unit (To Supplier).' };
    if (!dcxProcessAllowed_(jo, 'Dyeing')) {
      return { success: false, error: '"Dyeing" has no rate set for JO ' + jo + ' in Style Tracker. Add the Dyeing rate first.' };
    }
    var c = dcxDyeContext_(jo), byKey = {};
    c.lines.forEach(function (L) { byKey[L.key] = L; });
    var lines = (d.lines || []).filter(function (x) { return dcxN_(x.qty) > 0; });
    if (!lines.length) return { success: false, error: 'Enter Issue Kg for at least one colour.' };

    var tol = DCX_TOLERANCE_KG + 0.0005, bal = c.bal, allocs = [];
    for (var i = 0; i < lines.length; i++) {
      var x = lines[i], L = byKey[dcxS_(x.key).toUpperCase()];
      if (!L) return { success: false, error: 'Colour / fabric not found in Style Tracker for ' + jo + '. Reload the JO.' };
      var q = dcxR3_(dcxN_(x.qty)), done = c.dyed[L.key] || 0, lineBal = L.required - done;
      if (q > lineBal + tol) {
        return { success: false, error: 'DC not saved. ' + L.color + ' ' + L.fabric + ' ' + L.kDia + '/' + L.fDia + ': ' +
          q.toFixed(3) + ' kg is more than the colour balance ' + Math.max(0, lineBal).toFixed(3) + ' kg.' };
      }
      var need = q, srcs = c.sources[L.greyKey] || [], first = true;
      for (var j = 0; j < srcs.length && need > 0.0005; j++) {
        var s = srcs[j], have = bal[s.dcNo] || 0;
        if (have <= 0.0005) continue;
        var take = dcxR3_(Math.min(have, need));
        allocs.push({ L: L, src: s, before: have, qty: take, rolls: first ? dcxN_(x.rolls) : 0 });
        bal[s.dcNo] = dcxR3_(have - take); need = dcxR3_(need - take); first = false;
      }
      if (need > tol) {
        return { success: false, error: 'DC not saved. Grey ' + L.fabric + ' ' + L.kDia + '/' + L.fDia + ': ' + q.toFixed(3) +
          ' kg asked but only ' + (q - need).toFixed(3) + ' kg of knitted grey is available (Knitting DCs not yet sent on).' };
      }
    }

    var code = getProcessMaps().codeMap['Dyeing'] || 'D';
    var base = nextIdForPrefix(SHEETS.PROCESS_DC, 1, code);
    var date = d.date || dcxToday_(), today = dcxToday_(), styleNo = dcxS_(c.st[0]['Style Number']);
    var rows = allocs.map(function (a, i) {
      var L = a.L;
      return {
        'DC No': dcxRowId_(base, i), 'Date': date, 'JO Number': jo, 'Style Number': styleNo, 'Color': L.color,
        'From Process': 'Knitting', 'From Supplier': a.src.supplier, 'To Process': 'Dyeing', 'To Supplier': sup,
        'Lot No': '', 'Previous DC No': a.src.dcNo, 'Previous DC Qty KG': a.src.issue, 'Available KG': a.before,
        'DC Issue Qty KG': a.qty, 'Rolls': a.rolls, 'Balance KG': dcxR3_(a.before - a.qty), 'Remarks': dcxS_(d.remarks),
        'Status': 'Approved', 'Approved By': 'Auto', 'Approval Date': today, 'Counts': L.counts, 'Content': L.content,
        'GSM': L.gsm, 'Loss KG': 0, 'Loss %': 0, 'Vehicle No': dcxS_(d.vehicleNo), 'Supplier DC No': dcxS_(d.supplierDcNo),
        'DC Group': base, 'Fabric type': L.fabric, 'Design': L.design, 'Gauge': L.gg, 'Loop length': L.ll,
        'Knitting Dia': L.kDia, 'Finishing dia': L.fDia, 'Line Key': L.key, 'Grey Key': L.greyKey,
        'Yarn Kind': '', 'Bags': '', 'Cones': ''
      };
    });
    dcxWriteRows_(rows);
    var total = allocs.reduce(function (s, a) { return s + a.qty; }, 0);
    return { success: true, dcNo: base, message: 'Dyeing DC ' + base + ' created for ' + total.toFixed(3) + ' kg.' };
  } catch (err) {
    return { success: false, error: err.message };
  } finally { lock.releaseLock(); }
}

/* ---------------- DC PRINT (all Process DCs) ---------------- */
function generateDcHtml(dc) {
  try {
    return dcxDcHtml_(dc['DC No']);
  } catch (e) {
    console.log('New DC layout failed, using classic: ' + e.message);
    return generateDcHtmlClassic_(dc);
  }
}

function dcxQr_(text) {
  try {
    var r = UrlFetchApp.fetch('https://api.qrserver.com/v1/create-qr-code/?size=120x120&margin=0&data=' +
      encodeURIComponent(text), { muteHttpExceptions: true });
    if (r.getResponseCode() === 200) return 'data:image/png;base64,' + Utilities.base64Encode(r.getContent());
  } catch (e) { }
  return '';
}

function dcxDcHtml_(dcNo) {
  var all = dcxReadDc_();
  var row = all.filter(function (r) { return r['DC No'] === dcNo; })[0];
  if (!row) throw new Error('DC ' + dcNo + ' not found');
  var group = dcxS_(row['DC Group']) || row['DC No'];
  var rows = all.filter(function (r) { return (dcxS_(r['DC Group']) || r['DC No']) === group && r['Status'] !== 'Rejected'; });
  if (!rows.length) rows = [row];
  var h0 = rows[0], jo = dcxS_(h0['JO Number']);
  var st = dcxJoRows_(jo), first = st[0] || {};

  var specOf = function (r) {
    if (dcxS_(r['Fabric type'])) {
      return { fabric: dcxS_(r['Fabric type']), content: dcxS_(r['Content']), design: dcxS_(r['Design']) || 'SOLID',
        gsm: dcxS_(r['GSM']), gg: dcxS_(r['Gauge']), ll: dcxS_(r['Loop length']), kDia: dcxS_(r['Knitting Dia']),
        fDia: dcxS_(r['Finishing dia']), counts: dcxS_(r['Counts']), yarnColor: '', yarnType: '' };
    }
    var m = st.filter(function (s) {
      return dcxS_(s['Counts']) === dcxS_(r['Counts']) && dcxS_(s['Content']) === dcxS_(r['Content']) &&
        (!dcxS_(r['GSM']) || dcxS_(s['GSM']) === dcxS_(r['GSM']));
    })[0] || first;
    var s = dcxSpec_(m);
    s.counts = dcxS_(r['Counts']); s.content = dcxS_(r['Content']) || s.content; s.gsm = dcxS_(r['GSM']) || s.gsm;
    return s;
  };
  var fabDesc = function (s) { return s.fabric + (s.content ? '( ' + s.content.toUpperCase() + ' )' : ''); };

  var supRow = getSuppliers().filter(function (s) { return dcxS_(s['Supplier Name']) === dcxS_(h0['To Supplier']); })[0] || {};
  var toGst = dcxS_(supRow['GST']);
  var to = { name: dcxS_(h0['To Supplier']) || JIWIN_INWARD_LABEL, address: dcxS_(supRow['Address']), gst: toGst,
    pan: toGst.length >= 12 ? toGst.substr(2, 10) : '' };
  var info = {
    dcNo: group, date: dcxDmy_(h0['Date']), buyer: dcxS_(first['Buyer Name']), jo: jo,
    order: dcxS_(first['Order/OCN No']) || dcxS_(first['Buyer Order No']), merch: dcxS_(first['Merchandiser']),
    supDc: dcxS_(h0['Supplier DC No']), vehicle: dcxS_(h0['Vehicle No']) || '-', remarks: dcxS_(h0['Remarks']),
    approvedBy: dcxS_(h0['Approved By']), qr: dcxQr_(group)
  };
  var toProc = dcxS_(h0['To Process']);
  var isKnit = toProc === 'Knitting' && dcxS_(h0['From Process']) === YARN_STORE_LABEL;
  return isKnit ? dcxKnitHtml_(rows, info, to, specOf, fabDesc) : dcxProcessHtml_(rows, info, to, specOf, fabDesc, toProc);
}

function dcxCss_() {
  return '<style>@page{size:A4 landscape;margin:8mm}' +
    'body{font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#000;margin:0;padding:6px}' +
    'table{border-collapse:collapse;width:100%}' +
    '.box{border:1px solid #000}.box td{vertical-align:top}' +
    '.g th,.g td{border:1px solid #000;padding:4px 5px;font-size:11px}.g th{font-size:12px;background:#fff}' +
    '.c{text-align:center}.r{text-align:right}.b{font-weight:bold}' +
    '.ttl{text-align:center;font-weight:bold;font-size:16px;padding:4px;border-bottom:1px solid #000}' +
    '.pty td{padding:5px 10px;width:50%}.pty .a{padding-left:16px;line-height:1.35}' +
    'table.inf{width:auto}.inf td{padding:3px 8px;font-size:12px;white-space:nowrap}' +
    '</style>';
}

function dcxParties_(to) {
  var f = DCX_FROM;
  var cell = function (lbl, p) {
    return '<td><b>' + lbl + ' : ' + dcxE_(p.name) + '</b><div class="a" style="margin-top:8px;min-height:40px">' + dcxE_(p.address) +
      '</div><div class="a" style="margin-top:10px">GST NO : ' + dcxE_(p.gst) + ' / PAN NO : ' + dcxE_(p.pan) + '</div></td>';
  };
  return '<table class="pty" style="border-bottom:1px solid #000"><tr>' + cell('From', f).replace('<td>', '<td style="border-right:1px solid #000">') +
    cell('To', to) + '</tr></table>';
}

function dcxFooter_(info) {
  return '<table style="border-top:1px solid #000"><tr><td style="padding:6px 12px;width:55%"><b>INVOICE TO</b> : ' + DCX_FROM.name +
    '</td><td style="padding:6px 12px"><b>FOLLOWER</b> : </td></tr>' +
    '<tr><td style="padding:6px 12px"><b>VEHICLE NO</b> : ' + dcxE_(info.vehicle) + ' &nbsp;&nbsp;&nbsp;&nbsp; <b>REMARKS</b> : ' +
    dcxE_(info.remarks) + '</td><td></td></tr></table>' +
    '<table style="border-top:1px solid #000"><tr><td colspan="4" class="r" style="padding:8px 40px 18px">For ' + DCX_FROM.name + '</td></tr>' +
    '<tr><td class="c b" style="padding:4px 0 8px">Prepared By</td><td class="c b">Verified By</td><td class="c b">Received By</td>' +
    '<td class="c b">Approved By' + (info.approvedBy && info.approvedBy !== 'Auto' ? '<br><span style="font-weight:normal">' + dcxE_(info.approvedBy) + '</span>' : '') + '</td></tr></table>';
}

function dcxQrCell_(info) {
  return '<td style="width:95px;text-align:right;vertical-align:middle">' +
    (info.qr ? '<img src="' + info.qr + '" style="width:80px;height:80px">' : '') + '</td>';
}

/* YARN DELIVERY TO KNITTING */
function dcxKnitHtml_(rows, info, to, specOf, fabDesc) {
  var fo = [], fm = {}, yo = [], ym = {};
  rows.forEach(function (r) {
    var s = specOf(r), q = dcxN_(r['DC Issue Qty KG']);
    var fk = dcxS_(r['Line Key']) || dcxGreyKey_(s);
    if (!fm[fk]) { fm[fk] = { d: fabDesc(s), design: s.design, gsm: s.gsm, gg: s.gg, ll: s.ll, kDia: s.kDia, fDia: s.fDia, qty: 0 }; fo.push(fk); }
    fm[fk].qty += q;
    var kind = dcxS_(r['Yarn Kind']) || 'Main';
    var yk = dcxS_(r['Counts']) + '|' + dcxS_(r['Content']) + '|' + kind;
    if (!ym[yk]) {
      var yc = kind === 'Lycra' ? '' : (s.yarnColor || 'GREY');
      ym[yk] = { d: dcxS_(r['Counts']) + (dcxS_(r['Content']) ? ' (' + dcxS_(r['Content']) + ')' : ''), color: yc, cones: 0, bags: 0, qty: 0 };
      yo.push(yk);
    }
    ym[yk].cones += dcxN_(r['Cones']); ym[yk].bags += dcxN_(r['Bags']); ym[yk].qty += q;
  });
  var h = '<!DOCTYPE html><html><head><meta charset="utf-8"><title>' + dcxE_(info.dcNo) + '</title>' + dcxCss_() + '</head><body><div class="box">';
  h += '<div class="ttl">YARN DELIVERY TO KNITTING</div>' + dcxParties_(to);
  h += '<table style="border-bottom:1px solid #000"><tr><td style="padding:14px 6px 6px"><table class="inf">' +
    '<tr><td class="b">DC NO</td><td>: ' + dcxE_(info.dcNo) + '</td><td class="b">DATE</td><td>: ' + dcxE_(info.date) +
    '</td><td class="b">BUYER</td><td>: ' + dcxE_(info.buyer) + '</td><td class="b">JOB NO</td><td>: ' + dcxE_(info.jo) + '</td></tr>' +
    '<tr><td class="b">PROCESS</td><td>: KNITTING</td><td class="b">' + (info.supDc ? 'SUPPLIER DC' : '') + '</td><td>' +
    (info.supDc ? ': ' + dcxE_(info.supDc) : '') + '</td><td></td><td></td><td class="b">ORDER</td><td>: ' + dcxE_(info.order) + '</td></tr>' +
    '</table></td>' + dcxQrCell_(info) + '</tr></table>';
  h += '<table class="g"><tr><th>SNO</th><th>FABRIC</th><th>COLORNAME</th><th>DESIGN</th><th>GSM</th><th>GG</th><th>LL</th>' +
    '<th>KDIA</th><th>FDIA</th><th>UOM</th><th>Prog Qty</th></tr>';
  var ft = 0;
  fo.forEach(function (k, i) {
    var f = fm[k]; ft += f.qty;
    h += '<tr><td class="c">' + (i + 1) + '</td><td>' + dcxE_(f.d) + '</td><td>GREY</td><td class="r">' + dcxE_(f.design) +
      '</td><td class="c">' + dcxE_(f.gsm) + '</td><td class="c">' + dcxE_(f.gg) + '</td><td class="c">' + dcxE_(f.ll) +
      '</td><td class="c">' + dcxE_(f.kDia) + '</td><td class="c">' + dcxE_(f.fDia) + '</td><td class="c">KGS</td><td class="r">' +
      f.qty.toFixed(3) + '</td></tr>';
  });
  h += '<tr><td colspan="9" class="c b">TOTAL</td><td></td><td class="r b">' + ft.toFixed(3) + '</td></tr>';
  h += '<tr><th>SNO</th><th colspan="3">Yarn Description</th><th colspan="2">Color</th><th>Uom</th><th>Cones</th><th>Bags</th><th colspan="2">Iss Qty</th></tr>';
  var yt = 0, yb = 0, yc = 0;
  yo.forEach(function (k, i) {
    var y = ym[k]; yt += y.qty; yb += y.bags; yc += y.cones;
    h += '<tr><td class="c">' + (i + 1) + '</td><td colspan="3">' + dcxE_(y.d) + '</td><td colspan="2">' + dcxE_(y.color) +
      '</td><td class="c">KGS</td><td class="c">' + y.cones + '</td><td class="c">' + y.bags + '</td><td colspan="2" class="r">' +
      y.qty.toFixed(3) + '</td></tr>';
  });
  h += '<tr><td></td><td colspan="3" class="c b">TOTAL</td><td colspan="2"></td><td></td><td class="c b">' + yc + '</td><td class="c b">' +
    yb + '</td><td colspan="2" class="r b">' + yt.toFixed(3) + '</td></tr></table>';
  h += '<div style="height:40px"></div>' + dcxFooter_(info) + '</div></body></html>';
  return h;
}

/* FABRIC DELIVERY TO PROCESS (Dyeing and every other process) */
function dcxProcessHtml_(rows, info, to, specOf, fabDesc, toProc) {
  var po = [], pm = {}, go = [], gm = {};
  var fromProc = dcxS_(rows[0]['From Process']);
  rows.forEach(function (r) {
    var s = specOf(r), q = dcxN_(r['DC Issue Qty KG']), color = dcxS_(r['Color']);
    var pk = dcxS_(r['Line Key']) || (color + '#' + dcxGreyKey_(s));
    if (!pm[pk]) { pm[pk] = { d: fabDesc(s), design: s.design, color: color || '-', fDia: s.fDia, kDia: s.kDia, gsm: s.gsm, gg: s.gg, ll: s.ll, rolls: 0, qty: 0 }; po.push(pk); }
    pm[pk].rolls += dcxN_(r['Rolls']); pm[pk].qty += q;
    var gk = dcxS_(r['Grey Key']) || dcxGreyKey_(s);
    if (fromProc !== 'Knitting') gk += '#' + color;
    if (!gm[gk]) {
      gm[gk] = { d: [fabDesc(s), s.design, s.gsm, s.gg, s.ll, s.kDia, s.fDia].join(' / '), counts: s.counts,
        color: fromProc === 'Knitting' ? 'GREY' : (color || '-'), rolls: 0, qty: 0 };
      go.push(gk);
    }
    gm[gk].rolls += dcxN_(r['Rolls']); gm[gk].qty += q;
  });
  var procLabel = toProc === JIWIN_INWARD_LABEL ? 'JIWIN INWARD' : 'FABRIC ' + toProc.toUpperCase();
  var h = '<!DOCTYPE html><html><head><meta charset="utf-8"><title>' + dcxE_(info.dcNo) + '</title>' + dcxCss_() + '</head><body><div class="box">';
  h += '<div class="ttl">FABRIC DELIVERY TO PROCESS</div>' + dcxParties_(to);
  h += '<table style="border-bottom:1px solid #000"><tr><td style="padding:6px"><table class="inf">' +
    '<tr><td class="b">DC NO :</td><td>' + dcxE_(info.dcNo) + '</td><td style="width:60px"></td><td class="b">DATE :</td><td>' + dcxE_(info.date) + '</td></tr>' +
    '<tr><td class="b">BUYER :</td><td>' + dcxE_(info.buyer) + '</td><td></td><td class="b">JOB NO :</td><td>' + dcxE_(info.jo) + '</td></tr>' +
    '<tr><td class="b">PROCESS :</td><td>' + dcxE_(procLabel) + '</td><td></td><td class="b">ORDER NO</td><td>' + dcxE_(info.order) + '</td></tr>' +
    '<tr><td class="b">ASSTMERCHANDISER :</td><td>' + dcxE_(info.merch) + '</td><td></td><td class="b">' + (info.supDc ? 'SUPPLIER DC :' : '') +
    '</td><td>' + dcxE_(info.supDc) + '</td></tr></table></td>' + dcxQrCell_(info) + '</tr></table>';
  h += '<table class="g"><tr><th>SNO</th><th>Fabric</th><th>Design</th><th>Color</th><th>FDIA</th><th>KDIA</th><th>GSM</th>' +
    '<th>GG</th><th>LL</th><th>UOM</th><th>Roll</th><th>Prog Qty</th></tr>';
  var pt = 0, pr = 0;
  po.forEach(function (k, i) {
    var p = pm[k]; pt += p.qty; pr += p.rolls;
    h += '<tr><td class="c">' + (i + 1) + '</td><td>' + dcxE_(p.d) + '</td><td class="c">' + dcxE_(p.design) + '</td><td>' +
      dcxE_(p.color) + '</td><td class="c">' + dcxE_(p.fDia) + '</td><td class="c">' + dcxE_(p.kDia) + '</td><td class="c">' +
      dcxE_(p.gsm) + '</td><td class="c">' + dcxE_(p.gg) + '</td><td class="c">' + dcxE_(p.ll) + '</td><td class="c">KGS</td><td class="c">' +
      p.rolls + '</td><td class="r">' + p.qty.toFixed(3) + '</td></tr>';
  });
  h += '<tr><td colspan="9" class="c b">TOTAL</td><td></td><td class="c b">' + pr + '</td><td class="r b">' + pt.toFixed(3) + '</td></tr></table>';
  h += '<div class="b" style="padding:8px 10px;font-size:12px">' + dcxE_((fromProc || 'SOURCE').toUpperCase()) +
    ' DELIVERY DETAILS -(HSN CODE - 998821)</div>';
  h += '<table class="g"><tr><th>SNO</th><th>Fabric / Design / Gsm / GG / LL / KDia / FDia</th><th>Counts</th><th>Color</th>' +
    '<th>Uom</th><th>Rolls</th><th>Iss Qty</th></tr>';
  var gt = 0, gr = 0;
  go.forEach(function (k, i) {
    var g = gm[k]; gt += g.qty; gr += g.rolls;
    h += '<tr><td class="c">' + (i + 1) + '</td><td>' + dcxE_(g.d) + '</td><td>' + dcxE_(g.counts) + '</td><td>' + dcxE_(g.color) +
      '</td><td class="c">KGS</td><td class="c">' + g.rolls + '</td><td class="r">' + g.qty.toFixed(3) + '</td></tr>';
  });
  h += '<tr><td></td><td class="c b">TOTAL</td><td></td><td></td><td></td><td class="c b">' + gr + '</td><td class="r b">' + gt.toFixed(3) + '</td></tr></table>';
  h += '<div style="height:90px"></div>' + dcxFooter_(info) + '</div></body></html>';
  return h;
}

/**
 * For the Search & PDF picker: list of Yarn POs / Process DCs belonging to
 * one JO, newest first, so the person can pick from a short list instead of
 * typing the exact PO/DC number.
 */
function getYarnPOsForJO(joNumber) {
  return sheetToObjects(SHEETS.YARN_PO)
    .filter(function (r) { return String(r['JO Number']).trim() === String(joNumber).trim(); })
    .reverse();
}

function getProcessDCsForJO(joNumber) {
  return sheetToObjects(SHEETS.PROCESS_DC)
    .filter(function (r) { return String(r['JO Number']).trim() === String(joNumber).trim(); })
    .reverse();
}

function getYarnInwardsForJO(joNumber) {
  return sheetToObjects(SHEETS.YARN_INWARD)
    .filter(function (r) { return String(r['JO Number']).trim() === String(joNumber).trim(); })
    .reverse();
}

/*************** 9/10. JO DETAILS + TIMELINE ***************/
/** Current Open/Closed status for a JO — one row per JO in the JO Status sheet, absent = Open. */
function getJOStatus(joNumber) {
  var row = sheetToObjects(SHEETS.JO_STATUS).find(function (r) { return String(r['JO Number']).trim() === String(joNumber).trim(); });
  return row || { 'JO Number': joNumber, Status: 'Open', 'Closed Date': '', 'Closed By': '', Remarks: '' };
}

/** Marks a JO Closed — doesn't touch Style Tracker (kept read-only) or block any further action, just records status for display. */
function closeJO(joNumber, closedBy, remarks) {
  if (!joNumber) return { success: false, error: 'JO Number is required.' };
  var rowIdx = findRowIndex(SHEETS.JO_STATUS, 'JO Number', joNumber);
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (rowIdx === -1) {
    appendRowSafe_(SHEETS.JO_STATUS, [joNumber, 'Closed', today, closedBy || '', remarks || '']);
  } else {
    setCell(SHEETS.JO_STATUS, rowIdx, 'Status', 'Closed');
    setCell(SHEETS.JO_STATUS, rowIdx, 'Closed Date', today);
    setCell(SHEETS.JO_STATUS, rowIdx, 'Closed By', closedBy || '');
    setCell(SHEETS.JO_STATUS, rowIdx, 'Remarks', remarks || '');
  }
  return { success: true, message: joNumber + ' marked Closed.' };
}

function reopenJO(joNumber) {
  if (!joNumber) return { success: false, error: 'JO Number is required.' };
  var rowIdx = findRowIndex(SHEETS.JO_STATUS, 'JO Number', joNumber);
  if (rowIdx === -1) return { success: true, message: joNumber + ' is already Open.' };
  setCell(SHEETS.JO_STATUS, rowIdx, 'Status', 'Open');
  setCell(SHEETS.JO_STATUS, rowIdx, 'Closed Date', '');
  setCell(SHEETS.JO_STATUS, rowIdx, 'Closed By', '');
  return { success: true, message: joNumber + ' reopened.' };
}

function getJODetails(joNumber) {
  var jo = getJOInfo(joNumber);
  var timeline = [];

  getStyleTracker().filter(function (r) { return r['JO Number'] === joNumber; }).forEach(function (r) {
    timeline.push({ date: '', stage: 'JO Created', supplier: '—', status: 'Created', reference: joNumber });
  });
  sheetToObjects(SHEETS.YARN_PO).filter(function (r) { return r['JO Number'] === joNumber; }).forEach(function (r) {
    timeline.push({ date: r['PO Date'], stage: 'Yarn PO (' + r['PO Type'] + ')', supplier: r['Supplier'], status: r['Status'], reference: r['PO No'] });
  });
  sheetToObjects(SHEETS.YARN_INWARD).filter(function (r) { return r['JO Number'] === joNumber; }).forEach(function (r) {
    timeline.push({ date: r['Date'], stage: 'Yarn Receipt', supplier: r['Mill Name'], status: 'Received', reference: r['Inward No'] });
  });
  sheetToObjects(SHEETS.PROCESS_DC).filter(function (r) { return r['JO Number'] === joNumber; }).forEach(function (r) {
    timeline.push({ date: r['Date'], stage: r['To Process'], supplier: r['To Supplier'], status: r['Status'], reference: r['DC No'] });
  });
  sheetToObjects(SHEETS.JO_TRANSFER).filter(function (r) { return r['Source JO'] === joNumber || r['Destination JO'] === joNumber; }).forEach(function (r) {
    timeline.push({ date: r['Date'], stage: 'JO Transfer', supplier: r['Source JO'] + ' → ' + r['Destination JO'], status: r['Status'], reference: r['Transfer No'] });
  });
  sheetToObjects(SHEETS.BILL_PASSING).filter(function (r) { return r['JO Number'] === joNumber; }).forEach(function (r) {
    timeline.push({ date: r['Date'], stage: 'Bill Passing', supplier: r['Supplier'], status: r['Status'], reference: r['Bill No'] });
  });

  timeline.sort(function (a, b) { return String(a.date).localeCompare(String(b.date)); });

  var lastDc = sheetToObjects(SHEETS.PROCESS_DC).filter(function (r) { return r['JO Number'] === joNumber; }).pop();

  return {
    jo: jo,
    jobStatus: getJOStatus(joNumber),
    currentProcess: lastDc ? lastDc['To Process'] : '—',
    currentSupplier: lastDc ? lastDc['To Supplier'] : '—',
    lastUpdated: lastDc ? lastDc['Date'] : '—',
    timeline: timeline
  };
}

/**
 * Budget vs Actual for one JO, across yarn/fabric procurement, process
 * (job-work) costs, and accessories. "Actual" only counts what the system
 * can see (Approved Yarn POs, Approved Bill Passing), so anything paid
 * outside the app won't show here.
 */
function getBudgetSummary(joNumber) {
  var styleRows = getStyleTracker().filter(function (r) { return String(r['JO Number']).trim() === String(joNumber).trim(); });
  if (!styleRows.length) return null;

  var totalRequiredKg = styleRows.reduce(function (s, r) { return s + (Number(r['Fabric Required (kg)']) || 0); }, 0);

  // Yarn / fabric procurement budget: each row's own required kg x its own yarn rate.
  var yarnBudget = styleRows.reduce(function (s, r) {
    return s + (Number(r['Fabric Required (kg)']) || 0) * (Number(r['Yarn rate']) || 0);
  }, 0);

  var yarnPOs = sheetToObjects(SHEETS.YARN_PO).filter(function (r) { return String(r['JO Number']).trim() === String(joNumber).trim(); });
  var yarnActual = yarnPOs.filter(function (r) { return r['Status'] === 'Approved'; })
    .reduce(function (s, r) { return s + (Number(r['PO Value']) || 0); }, 0);
  var yarnPending = yarnPOs.filter(function (r) { return r['Status'] === 'Pending Approval'; })
    .reduce(function (s, r) { return s + (Number(r['PO Value']) || 0); }, 0);

  // Process (job-work) budget: each active rate x the JO's total required kg,
  // as a planning proxy (assumes the full order passes through that stage).
  var processes = getAllowedProcessesForJO(joNumber); // [{process, rate}], rate > 0 only
  var dcRows = sheetToObjects(SHEETS.PROCESS_DC).filter(function (r) { return String(r['JO Number']).trim() === String(joNumber).trim(); });
  var dcFromProcess = {};
  dcRows.forEach(function (r) { dcFromProcess[r['DC No']] = r['From Process']; });

  var billRows = sheetToObjects(SHEETS.BILL_PASSING).filter(function (r) { return String(r['JO Number']).trim() === String(joNumber).trim(); });
  var processLines = processes.map(function (p) {
    var budget = totalRequiredKg * p.rate;
    var actual = 0, pending = 0;
    billRows.forEach(function (bill) {
      var firstDc = String(bill['DC Numbers'] || '').split(',')[0].trim();
      if (dcFromProcess[firstDc] !== p.process) return;
      if (bill['Status'] === 'Approved') actual += Number(bill['Bill Amount']) || 0;
      if (bill['Status'] === 'Pending Approval') pending += Number(bill['Bill Amount']) || 0;
    });
    return { process: p.process, rate: p.rate, budget: budget, actual: actual, pending: pending, variance: budget - actual };
  });
  var processBudgetTotal = processLines.reduce(function (s, p) { return s + p.budget; }, 0);
  var processActualTotal = processLines.reduce(function (s, p) { return s + p.actual; }, 0);
  var processPendingTotal = processLines.reduce(function (s, p) { return s + p.pending; }, 0);

  var debitNotes = sheetToObjects(SHEETS.DEBIT_NOTE).filter(function (r) { return String(r['JO Number']).trim() === String(joNumber).trim(); });
  var debitNoteTotal = debitNotes.reduce(function (s, r) { return s + (Number(r['Amount']) || 0); }, 0);

  // Accessories: budgeted only where a Rate (Rs/unit) has been entered; actual
  // spend on accessories isn't tracked anywhere in the system yet.
  var accRows = getAccessoriesBOM(joNumber);
  var accessoriesBudget = accRows.reduce(function (s, r) { return s + (Number(r['Amount']) || 0); }, 0);
  var accessoriesRated = accRows.filter(function (r) { return Number(r['Rate (Rs/unit)']) > 0; }).length;

  var budgetTotal = yarnBudget + processBudgetTotal + accessoriesBudget;
  var actualTotal = yarnActual + processActualTotal; // accessories excluded — not tracked

  return {
    joNumber: joNumber,
    totalRequiredKg: totalRequiredKg,
    yarn: { budget: yarnBudget, actual: yarnActual, pending: yarnPending, variance: yarnBudget - yarnActual },
    processes: processLines,
    processTotal: { budget: processBudgetTotal, actual: processActualTotal, pending: processPendingTotal, variance: processBudgetTotal - processActualTotal },
    debitNoteTotal: debitNoteTotal,
    accessories: { budget: accessoriesBudget, ratedCount: accessoriesRated, totalCount: accRows.length },
    grand: { budget: budgetTotal, actual: actualTotal, variance: budgetTotal - actualTotal }
  };
}

/*************** 12. JO TRANSFER + STOCK (Yarn / Cora Fabric / Dyed Fabric) ***************/
/**
 * Balance for one JO + Counts (+ Color, for Dyed Fabric which is color-specific)
 * within one stock type. "Received" comes from Yarn Inward joined back to the
 * Yarn PO that placed it (so Yarn/Cora Fabric/Dyed Fabric receipts don't mix).
 * Transfers move stock of one type between JOs; Yarn/Cora Fabric transfers are
 * Counts-scoped (grey, shared across colors); Dyed Fabric transfers, if used,
 * are also Counts-scoped only — color isn't tracked on the transfer record,
 * since dyed fabric is rarely moved between JOs.
 */
/**
 * Weighted-average landed rate (₹/kg) for received stock of one spec, from
 * the Yarn PO(s) that placed it. Used as the starting value for a fresh lot
 * (before any processing cost is added on top of it).
 */
function getWeightedStockRate(joNumber, counts, stockType, color, content, gsm) {
  var poMeta = {};
  sheetToObjects(SHEETS.YARN_PO).forEach(function (p) {
    poMeta[p['PO No']] = { type: p['PO Type'], counts: p['Counts'], color: p['Color'], content: p['Content'], gsm: p['GSM'], rate: Number(p['Rate (Rs/kg)']) || 0 };
  });
  var totalKg = 0, totalValue = 0;
  sheetToObjects(SHEETS.YARN_INWARD).forEach(function (r) {
    if (String(r['JO Number']).trim() !== String(joNumber).trim()) return;
    var meta = poMeta[r['Yarn PO No']];
    if (!meta || meta.type !== stockType) return;
    if (String(meta.counts || '').trim() !== String(counts).trim()) return;
    if (color && String(meta.color || '').trim() !== String(color).trim()) return;
    if (content && String(meta.content || '').trim() !== String(content).trim()) return;
    if (gsm && String(meta.gsm || '').trim() !== String(gsm).trim()) return;
    var kg = Number(r['Received KG']) || 0;
    totalKg += kg;
    totalValue += kg * meta.rate;
  });
  return totalKg > 0 ? totalValue / totalKg : 0;
}

/**
 * Accumulated ₹ value of one Approved DC's output, walking back through the
 * chain to where the lot started (Yarn/Cora Fabric/Dyed Fabric stock). Every
 * rated process a lot passes through (Knitting, Dyeing, Stenter, ...) adds
 * its own job-work cost — rate × that stage's own Issue Qty — on top of the
 * value it inherited, so the loss at each stage stays sunk into the
 * surviving fabric's ₹/kg rather than vanishing from the books. Memoised
 * per call since the same upstream DC can feed several downstream ones.
 */
function computeDCValue(dcNo, memo) {
  memo = memo || {};
  if (memo[dcNo]) return memo[dcNo];
  var dc = sheetToObjects(SHEETS.PROCESS_DC).find(function (r) { return r['DC No'] === dcNo; });
  if (!dc) return null;
  var issueQty = Number(dc['DC Issue Qty KG']) || 0;
  var availableQty = Number(dc['Available KG']) || 0;
  var sourceValue;
  var origin = CHAIN_ORIGINS.filter(function (o) { return o.fromProcess === dc['From Process']; })[0];

  if (origin || !dc['Previous DC No']) {
    var stockType = origin ? origin.stockType : 'Yarn';
    var color = (origin && origin.colorSpecific) ? dc['Color'] : '';
    var gsm = (origin && origin.useGsm) ? dc['GSM'] : '';
    var rate = getWeightedStockRate(dc['JO Number'], dc['Counts'], stockType, color, dc['Content'], gsm);
    sourceValue = rate * availableQty;
  } else {
    var prev = computeDCValue(dc['Previous DC No'], memo);
    sourceValue = prev ? prev.valuePerKg * availableQty : 0;
  }

  var stageRate = origin ? 0 : getRateForJOProcess(dc['JO Number'], dc['From Process']);
  var stageCost = stageRate * issueQty;
  var totalValue = sourceValue + stageCost;
  var result = {
    dcNo: dcNo, issueQty: issueQty, sourceValue: sourceValue, stageCost: stageCost,
    totalValue: totalValue, valuePerKg: issueQty > 0 ? totalValue / issueQty : 0
  };
  memo[dcNo] = result;
  return result;
}

/** Finished fabric sitting at Jiwin — every Approved "-> Jiwin Inward" DC, with its accumulated Qty + Value. */
function getJiwinInwardStock() {
  var memo = {};
  return sheetToObjects(SHEETS.PROCESS_DC).filter(function (r) {
    return r['To Process'] === JIWIN_INWARD_LABEL && (r['Status'] === 'Approved' || r['Status'] === 'Billed');
  }).map(function (dc) {
    var v = computeDCValue(dc['DC No'], memo);
    // Some of this lot may since have been drawn back out for a reprocess
    // (another DC with this one as its Previous DC No) — show only what's
    // actually still sitting here, not the original full receipt.
    var remainingKg = getDCChainBalance(dc['DC No']);
    return {
      dcNo: dc['DC No'], date: dc['Date'], joNumber: dc['JO Number'], styleNumber: dc['Style Number'],
      color: dc['Color'] || '', counts: dc['Counts'] || '', content: dc['Content'] || '', gsm: dc['GSM'] || '',
      fromProcess: dc['From Process'], qtyKg: remainingKg, value: v.valuePerKg * remainingKg, valuePerKg: v.valuePerKg
    };
  }).filter(function (row) { return row.qtyKg > 0.01; });
}

function getStockBalance(joNumber, counts, stockType, color, content, gsm) {
  var poMeta = {};
  sheetToObjects(SHEETS.YARN_PO).forEach(function (p) {
    poMeta[p['PO No']] = { type: p['PO Type'], counts: p['Counts'], color: p['Color'], content: p['Content'], gsm: p['GSM'] };
  });
  var received = sheetToObjects(SHEETS.YARN_INWARD).filter(function (r) {
    if (String(r['JO Number']).trim() !== String(joNumber).trim()) return false;
    var meta = poMeta[r['Yarn PO No']];
    if (!meta || meta.type !== stockType) return false;
    if (String(meta.counts || '').trim() !== String(counts).trim()) return false;
    if (color && String(meta.color || '').trim() !== String(color).trim()) return false;
    if (content && String(meta.content || '').trim() !== String(content).trim()) return false;
    if (gsm && String(meta.gsm || '').trim() !== String(gsm).trim()) return false;
    return true;
  }).reduce(function (s, r) { return s + (Number(r['Received KG']) || 0); }, 0);

  var transfers = sheetToObjects(SHEETS.JO_TRANSFER).filter(function (r) {
    if (String(r['Counts'] || '').trim() !== String(counts).trim()) return false;
    if (String(r['Stock Type'] || 'Yarn') !== stockType) return false;
    if (content && String(r['Content'] || '').trim() !== String(content).trim()) return false;
    if (gsm && String(r['GSM'] || '').trim() !== String(gsm).trim()) return false;
    return true;
  });
  var transferredIn = transfers.filter(function (r) { return String(r['Destination JO']).trim() === String(joNumber).trim(); })
    .reduce(function (s, r) { return s + (Number(r['Qty KG']) || 0); }, 0);
  var transferredOut = transfers.filter(function (r) { return String(r['Source JO']).trim() === String(joNumber).trim(); })
    .reduce(function (s, r) { return s + (Number(r['Qty KG']) || 0); }, 0);
  // Yarn already issued into production (Process DC, Yarn Store -> Knitting) is
  // no longer "available" stock — it's WIP at the knitter.
  var issuedToProduction = getStockIssuedToProduction(joNumber, counts, stockType, color, content, gsm);
  return {
    received: received, transferredIn: transferredIn, transferredOut: transferredOut,
    issuedToProduction: issuedToProduction,
    available: received + transferredIn - transferredOut - issuedToProduction
  };
}

/** Kg of a Counts/Content already drawn from Yarn stock into production via Process DC. */
function getYarnIssuedToKnitting(joNumber, counts, content) {
  return getStockIssuedToProduction(joNumber, counts, 'Yarn', '', content, '');
}

/** Kg of a stock spec already drawn from its origin (Yarn Store / Cora Fabric Store / Dyed Fabric Store) into the DC chain. */
function getStockIssuedToProduction(joNumber, counts, stockType, color, content, gsm) {
  var origin = CHAIN_ORIGINS.filter(function (o) { return o.stockType === stockType; })[0];
  if (!origin) return 0;
  return sheetToObjects(SHEETS.PROCESS_DC).filter(function (r) {
    if (String(r['JO Number']).trim() !== String(joNumber).trim()) return false;
    if (r['From Process'] !== origin.fromProcess) return false;
    if (String(r['Counts'] || '').trim() !== String(counts).trim()) return false;
    if (content && String(r['Content'] || '').trim() !== String(content).trim()) return false;
    if (origin.useGsm && gsm && String(r['GSM'] || '').trim() !== String(gsm).trim()) return false;
    if (origin.colorSpecific && color && String(r['Color'] || '').trim() !== String(color).trim()) return false;
    return r['Status'] !== 'Rejected';
  }).reduce(function (s, r) { return s + (Number(r['DC Issue Qty KG']) || 0); }, 0);
}

/** Remaining un-forwarded balance of one Process DC (its output minus whatever later DCs already drew from it). Only Approved DCs count as valid inputs to the next stage. */
function getDCChainBalance(dcNo) {
  var dc = sheetToObjects(SHEETS.PROCESS_DC).find(function (r) { return r['DC No'] === dcNo; });
  if (!dc || (dc['Status'] !== 'Approved' && dc['Status'] !== 'Billed')) return 0;
  if (dc['Balance Status'] === 'Closed' || dc['Balance Status'] === 'Debited') return 0;
  var issued = Number(dc['DC Issue Qty KG']) || 0;
  var forwarded = sheetToObjects(SHEETS.PROCESS_DC).filter(function (r) {
    return r['Previous DC No'] === dcNo && r['Status'] !== 'Rejected';
  }).reduce(function (s, r) { return s + (Number(r['DC Issue Qty KG']) || 0); }, 0);
  return issued - forwarded;
}

/**
 * Everything a new Process DC can draw from for a JO: Yarn Store balances
 * (yarn received but not yet issued to Knitting) and the un-forwarded output
 * of every Approved DC in the chain (e.g. Knitting's output waiting to go to
 * Dyeing). Each entry carries the Counts/Content/GSM/Color it represents, so
 * the next DC inherits the right fabric spec automatically.
 */
function getProcessChainSources(joNumber) {
  var sources = [];
  var jo = getJOInfo(joNumber);
  var styleNumber = jo ? jo['Style Number'] : '';
  var dcFabOf = {};
  try { dcxReadDc_().forEach(function (r) { if (dcxS_(r['Fabric type'])) dcFabOf[r['DC No']] = dcxS_(r['Fabric type']); }); } catch (e) { }
  var stRowsForJo = dcJoRowsAll_(joNumber);

  CHAIN_ORIGINS.forEach(function (origin) {
    if (origin.colorSpecific) {
      getColorsForJO(joNumber).forEach(function (color) {
        getCountsForJO(joNumber, color, origin.useGsm).forEach(function (meta) {
          var bal = getStockBalance(joNumber, meta.counts, origin.stockType, color, meta.content, meta.gsm);
          if (bal.available > 0.01) {
            sources.push({
              sourceType: origin.sourceType, sourceDcNo: '', fromProcess: origin.fromProcess, fromSupplier: '',
              counts: meta.counts, content: meta.content, gsm: meta.gsm, color: color, yarnColor: meta.yarnColor,
              styleNumber: styleNumber, availableKg: bal.available
            });
          }
        });
      });
    } else {
      getCountsForJO(joNumber, '', origin.useGsm).forEach(function (meta) {
        var bal = getStockBalance(joNumber, meta.counts, origin.stockType, '', meta.content, meta.gsm);
        if (bal.available > 0.01) {
          sources.push({
            sourceType: origin.sourceType, sourceDcNo: '', fromProcess: origin.fromProcess, fromSupplier: '',
            counts: meta.counts, content: meta.content, gsm: meta.gsm, color: '', yarnColor: meta.yarnColor,
            styleNumber: styleNumber, availableKg: bal.available
          });
        }
      });
    }
  });

  // Every Approved DC's un-forwarded balance is a valid source for further
  // processing — including ones already at Jiwin Inward. This lets goods
  // that reached "finished" stock still go back out for a genuine reprocess
  // (a second Bio Washing pass, a redo, etc.) without inventing a separate
  // workflow — it's just another DC drawing from that lot's remaining balance.
  sheetToObjects(SHEETS.PROCESS_DC).filter(function (r) {
    return String(r['JO Number']).trim() === String(joNumber).trim() && (r['Status'] === 'Approved' || r['Status'] === 'Billed');
  }).forEach(function (dc) {
    var bal = getDCChainBalance(dc['DC No']);
    if (bal > 0.01) {
      sources.push({
        sourceType: 'DC', sourceDcNo: dc['DC No'], fromProcess: dc['To Process'], fromSupplier: dc['To Supplier'],
        isReprocess: dc['To Process'] === JIWIN_INWARD_LABEL,
        counts: dc['Counts'] || '', content: dc['Content'] || '', gsm: dc['GSM'] || '', color: dc['Color'] || '', knittingDia: dc['Knitting Dia'] || '',
        styleNumber: dc['Style Number'] || styleNumber, availableKg: bal,
        fabricTypes: dcFabOf[dc['DC No']] ? [dcFabOf[dc['DC No']]] : null
      });
    }
  });

  sources.forEach(function (s) { s.otherJO = false; });

  // Process DC sources are JO-specific. Stock belonging to another JO must not appear here;
  // use JO Transfer when material is intentionally moved/reassigned between JOs.

  // Tag every source with the Style Tracker fabric type(s) it can be. A DC row that already carries its
  // fabric type keeps it; anything else is matched on Counts + Content (+ GSM once knitted).
  // An empty list = "could not classify" (e.g. lycra yarn) and the screen keeps showing it under every fabric.
  sources.forEach(function (src) {
    if (src.fabricTypes && src.fabricTypes.length) return;
    src.fabricTypes = dcFabTypesForSpec_(stRowsForJo, src.counts, src.content, src.gsm, src.knittingDia);
  });

  return sources;
}

/*************** PROCESS DC — FABRIC TYPE + LOCKED QUANTITY ***************
 * Fabric type = Style Tracker "Fabric type" (S/J, 1*1 LYCRA RIB, NECK TAPE ...).
 * For one JO + fabric type (+ colour) the most that may ever go to one To Process is the Style Tracker
 * "Fabric Required (kg)"; whatever earlier (non-rejected, non-re-process) DCs already sent there is deducted.
 * The screen fills the quantity with that balance and locks it; saveProcessDC re-checks on the server.
 ***********************************************************************/
function dcFabNorm_(v) { return dcxS_(v).toUpperCase().replace(/\s+/g, ' '); }

function dcJoRowsAll_(joNumber) {
  return getStyleTracker().filter(function (r) { return dcxS_(r['JO Number']) === dcxS_(joNumber); });
}

/* Style Tracker rows of one JO for a fabric type (and colour, when given) */
function dcJoFabricRows_(joNumber, fabricType, color, knittingDia) {
  var ft = dcFabNorm_(fabricType), col = dcFabNorm_(color), kd = dcFabNorm_(knittingDia);
  return dcJoRowsAll_(joNumber).filter(function (r) {
    if (dcFabNorm_(r['Fabric type']) !== ft) return false;
    if (col && dcFabNorm_(r['Color']) !== col) return false;
    if (kd && dcFabNorm_(r['Knitting Dia']) !== kd) return false;
    return true;
  });
}

function dcJoFabricTypes_(joNumber) {
  var seen = {}, out = [];
  dcJoRowsAll_(joNumber).forEach(function (r) {
    var f = dcxS_(r['Fabric type']), k = dcFabNorm_(f);
    if (f && !seen[k]) { seen[k] = 1; out.push(f); }
  });
  return out;
}

/* Fabric types of a JO for the Process DC screen, with the quantity each one needs */
function getDcFabricTypes(joNumber) {
  var order = [], map = {};
  dcJoRowsAll_(joNumber).forEach(function (r) {
    var f = dcxS_(r['Fabric type']), dia = dcxS_(r['Knitting Dia']), gsm = dcxS_(r['GSM']);
    var k = [dcFabNorm_(f), dcFabNorm_(dia), dcFabNorm_(gsm)].join('||');
    if (!f) return;
    if (!map[k]) { map[k] = { fabricType: f, knittingDia: dia, gsm: gsm, requiredKg: 0, colors: [], parts: [] }; order.push(k); }
    var m = map[k], c = dcxS_(r['Color']), p = dcxS_(r['Part']);
    m.requiredKg += dcxN_(r['Fabric Required (kg)']);
    if (c && m.colors.indexOf(c) < 0) m.colors.push(c);
    if (p && m.parts.indexOf(p) < 0) m.parts.push(p);
  });
  return order.map(function (k) { map[k].requiredKg = dcxR3_(map[k].requiredKg); return map[k]; });
}

/* Fabric types a stock/WIP lot can belong to, matched on Counts + Content (+ GSM when it has one) */
function dcFabTypesForSpec_(stRows, counts, content, gsm, knittingDia) {
  var c = dcFabNorm_(counts), ct = dcFabNorm_(content), g = dcFabNorm_(gsm), kd = dcFabNorm_(knittingDia), seen = {}, out = [];
  stRows.forEach(function (r) {
    var f = dcxS_(r['Fabric type']); if (!f) return;
    if (c && dcFabNorm_(r['Counts']) !== c) return;
    if (ct && dcFabNorm_(r['Content']) !== ct) return;
    if (g && dcFabNorm_(r['GSM']) !== g) return;
    if (kd && dcFabNorm_(r['Knitting Dia']) !== kd) return;
    var k = dcFabNorm_(f);
    if (!seen[k]) { seen[k] = 1; out.push(f); }
  });
  return out;
}

/* Locked quantity for one JO + fabric type (+ colour) going to one To Process */
function getDcFabricBalance(joNumber, fabricType, color, toProcess, knittingDia) {
  var rows = dcJoFabricRows_(joNumber, fabricType, color, knittingDia);
  if (!rows.length) return { found: false, fabricType: fabricType, requiredKg: 0, issuedKg: 0, balanceKg: 0, bodyExtraKg: 0 };
  var required = rows.reduce(function (s, r) { return s + dcxN_(r['Fabric Required (kg)']); }, 0);
  var ft = dcFabNorm_(fabricType), col = dcFabNorm_(color), dia = dcFabNorm_(knittingDia), tp = dcxS_(toProcess);
  var hasBody = rows.some(function (r) { return /\bBODY\b/i.test(dcxS_(r['Part'])) || /BODY/i.test(dcxS_(r['Part'])); });
  var issued = dcxReadDc_().filter(function (r) {
    if (dcxS_(r['JO Number']) !== dcxS_(joNumber) || r['Status'] === 'Rejected') return false;
    if (dcxS_(r['To Process']) !== tp) return false;
    if (dcxS_(r['Process Type']) === 'Re-process') return false;
    if (dcFabNorm_(r['Fabric type']) !== ft) return false;
    if (col && dcFabNorm_(r['Color']) !== col) return false;
    if (dia && dcFabNorm_(r['Knitting Dia']) !== dia) return false;
    return true;
  }).reduce(function (s, r) { return s + dcxN_(r['DC Issue Qty KG']); }, 0);
  required = dcxR3_(required); issued = dcxR3_(issued);
  var baseBalance = Math.max(0, required - issued);
  var bodyExtraTotal = hasBody ? DCX_BODY_EXTRA_KG : 0;
  var bodyExtraUsed = hasBody ? Math.max(0, issued - required) : 0;
  var bodyExtraRemaining = Math.max(0, bodyExtraTotal - bodyExtraUsed);
  return { found: true, fabricType: dcxS_(rows[0]['Fabric type']), requiredKg: required, issuedKg: issued, baseBalanceKg: dcxR3_(baseBalance), bodyExtraKg: dcxR3_(bodyExtraRemaining), balanceKg: dcxR3_(baseBalance + bodyExtraRemaining) };
}

/* Writes the fabric type and its full spec (design / gauge / loop / dias) onto a Process DC row, so the printed DC shows the fabric that was picked */
function dcFabStamp_(rowNo, joNumber, fabricType, color, knittingDia) {
  var rows = dcJoFabricRows_(joNumber, fabricType, color, knittingDia);
  if (!rows.length) rows = dcJoFabricRows_(joNumber, fabricType, '', knittingDia);
  if (!rows.length) rows = dcJoFabricRows_(joNumber, fabricType, color);
  if (!rows.length) rows = dcJoFabricRows_(joNumber, fabricType, '');
  if (!rows.length) return;
  var sp = dcxSpec_(rows[0]), h = dcxEnsureHeaders_();
  var vals = { 'Fabric type': dcxS_(rows[0]['Fabric type']), 'Design': sp.design, 'Gauge': sp.gg, 'Loop length': sp.ll, 'Knitting Dia': sp.kDia, 'Finishing dia': sp.fDia };
  Object.keys(vals).forEach(function (k) { h.sheet.getRange(rowNo, h.col[k] + 1).setValue(vals[k]); });
}

/** Every process a JO's chain can move to next: rated processes from Style Tracker, plus the final "Jiwin Inward" stage. */
function getProcessDestinations(joNumber) {
  var list = getAllowedProcessesForJO(joNumber).map(function (x) { return x.process; });
  list.push(JIWIN_INWARD_LABEL);
  return list;
}

/** Every Counts (+ Color for Dyed Fabric) with a nonzero balance, across every JO — the Stock tab and the Transfer picker. */
function getStockList(stockType) {
  return withSheetCache_(function () {
  var jos = getStyleTrackerJOs();
  var rows = [];
  var useGsm = stockType !== 'Yarn'; // GSM applies once knitted (Cora/Dyed Fabric), not to raw Yarn.
  jos.forEach(function (jo) {
    if (stockType === 'Dyed Fabric') {
      getColorsForJO(jo).forEach(function (color) {
        getCountsForJO(jo, color, useGsm).forEach(function (meta) {
          var bal = getStockBalance(jo, meta.counts, stockType, color, meta.content, meta.gsm);
          if (bal.received > 0.001 || Math.abs(bal.transferredIn) > 0.001 || Math.abs(bal.transferredOut) > 0.001) {
            var rate = getWeightedStockRate(jo, meta.counts, stockType, color, meta.content, meta.gsm);
            rows.push(Object.assign({ joNumber: jo, color: color, rate: rate, value: rate * bal.available }, meta, bal));
          }
        });
      });
    } else {
      getCountsForJO(jo, '', useGsm).forEach(function (meta) {
        var bal = getStockBalance(jo, meta.counts, stockType, '', meta.content, meta.gsm);
        if (bal.received > 0.001 || Math.abs(bal.transferredIn) > 0.001 || Math.abs(bal.transferredOut) > 0.001) {
          var rate = getWeightedStockRate(jo, meta.counts, stockType, '', meta.content, meta.gsm);
          rows.push(Object.assign({ joNumber: jo, color: '', rate: rate, value: rate * bal.available }, meta, bal));
        }
      });
    }
  });
  return rows;
  });
}

/** Available (>0) Yarn/Cora Fabric stock across every JO — feeds the Transfer picker (Counts+Content first, JO second). */
function getAvailableStockForTransfer(stockType) {
  return getStockList(stockType || 'Yarn').filter(function (r) { return r.available > 0.01; });
}

function saveJOTransfer(d) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var qty = Number(d.qtyKg) || 0;
    var counts = String(d.counts || '').trim();
    var stockType = d.stockType || 'Yarn';
    var content = String(d.content || '').trim();
    var gsm = String(d.gsm || '').trim();

    if (counts) {
      var srcBal = getStockBalance(d.sourceJO, counts, stockType, '', content, gsm);
      if (qty > srcBal.available + 0.01) {
        return {
          success: false,
          error: 'Only ' + srcBal.available.toFixed(2) + ' kg of ' + stockType + ' Counts ' + counts +
            (gsm ? ' / GSM ' + gsm : '') + (content ? ' / ' + content : '') + ' is available at ' + d.sourceJO +
            ' (received: ' + srcBal.received.toFixed(2) + ' kg, already transferred out: ' + srcBal.transferredOut.toFixed(2) + ' kg).'
        };
      }
      var destCounted = getCountsForJO(d.destinationJO, '', stockType !== 'Yarn').find(function (c) {
        return c.counts === counts && c.content === content && c.gsm === gsm;
      });
      if (destCounted) {
        var destOrdered = getYarnPOUsage(d.destinationJO, counts, '', content, gsm);
        var destBal = getStockBalance(d.destinationJO, counts, stockType, '', content, gsm);
        var destCovered = destOrdered + destBal.transferredIn - destBal.transferredOut;
        var destBalance = destCounted.requiredKg - destCovered;
        if (qty > destBalance + 0.01) {
          return {
            success: false,
            error: 'Transferring ' + qty + ' kg would exceed ' + d.destinationJO + '\'s required qty for Counts ' + counts +
              ' (required: ' + destCounted.requiredKg.toFixed(2) + ' kg, already covered by PO/transfers: ' + destCovered.toFixed(2) +
              ' kg, room left: ' + Math.max(0, destBalance).toFixed(2) + ' kg).'
          };
        }
      }
    }

    var transferNo = nextId('TRF-', SHEETS.JO_TRANSFER, 1);
    appendRowSafe_(SHEETS.JO_TRANSFER, [
      transferNo, d.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      d.sourceJO, d.destinationJO, qty, d.remarks || '', 'Recorded', counts, stockType, content, gsm
    ]);
    return { success: true, transferNo: transferNo, message: 'Transfer ' + transferNo + ' recorded — ' + qty + ' kg of ' + (counts || stockType) + ' moved from ' + d.sourceJO + ' to ' + d.destinationJO + '.' };
  } catch (err) {
    return { success: false, error: err.message };
  } finally { lock.releaseLock(); }
}
function getJOTransfers() { return sheetToObjects(SHEETS.JO_TRANSFER); }

/*************** 13/14. BILL PASSING + BILL APPROVAL ***************/
/**
 * Fabric Program for one JO, matching Jiwin's own sheet shape: a YARN
 * PURCHASE table (Counts, Yarn Colour, Reqd Kgs) followed by one table per
 * process the JO actually uses, each broken down by Part/Color/Fabric —
 * not a single blended input/output figure — since real fabric programs
 * diverge by color once dyeing starts. Each row is walked forward
 * independently through every stage's own Loss %.
 */
function getFabricProgramData(joNumber) {
  var rows = getStyleTracker().filter(function (r) { return String(r['JO Number']).trim() === String(joNumber).trim(); });
  if (!rows.length) return null;
  var first = rows[0];
  var jo = getJOInfo(joNumber);
  var maps = getProcessMaps();
  // Everything before "Dyeing" in the process chain is still grey/greige —
  // the buyer's actual color only exists from the dyeing stage onward.
  // Yarn purchase and Knitting print grey here regardless of the buyer
  // color recorded on the Style Tracker row.
  var dyeingIdx = PROCESS_SEQUENCE.indexOf('Dyeing');
  function colorForProcess(processName, actualColor) {
    var idx = PROCESS_SEQUENCE.indexOf(processName);
    if (idx === -1 || dyeingIdx === -1) return actualColor; // unknown process — don't guess, keep the buyer color
    return idx < dyeingIdx ? 'GREY' : actualColor;
  }

  // Consumption: how each row's Fabric Required (kg) is worked out —
  // (Order Qty + Excess Qty) x Consumption per pc x (1 + Loss %).
  var consumptionLines = rows.filter(function (r) { return !isSampleWtRow_(r); }).map(function (r) {
    var orderQty = Number(r['Order Quantity']) || 0;
    var excessPct = Number(r['Excess %']) || 0;
    var excessQty = r['Excess Quantity'] !== '' && r['Excess Quantity'] != null
      ? (Number(r['Excess Quantity']) || 0) : Math.round(orderQty * excessPct) / 100;
    var consKg = Number(r['CAD Weight (g)']) || 0;
    var hasLy = String(r['Has Lycra'] || '').trim().toLowerCase() === 'yes';
    var lyPct = Number(r['Lycra %']) || 0;
    if (lyPct > 0 && lyPct < 1) lyPct = lyPct * 100;
    return {
      part: r['Part'] || '', color: r['Color'] || '', fabricType: r['Fabric type'] || '', gsm: r['GSM'] || '',
      orderQty: orderQty, excessPct: excessPct, excessQty: excessQty, totalPcs: orderQty + excessQty,
      consKg: consKg, consDozKg: consKg * 12, lossPct: Number(r['Loss %']) || 0,
      lycra: hasLy ? lyPct : 0, reqdKg: Number(r['Fabric Required (kg)']) || 0
    };
  });
  consumptionLines.forEach(function (l) { l.reqdKg = kgRound_(l.reqdKg); });
  var consumptionTotal = consumptionLines.reduce(function (s, l) { return s + l.reqdKg; }, 0);

  var yarnLines = [];
  rows.forEach(function (r) {
    var hasLycra = String(r['Has Lycra'] || '').trim().toLowerCase() === 'yes';
    var part = r['Part'] || '', counts = r['Counts'] || '', yarnColor = r['Yarn color'] || '';
    var color = 'GREY'; // yarn purchase always happens before dyeing
    if (hasLycra) {
      yarnLines.push({ part: part, color: color, counts: counts, yarnColor: yarnColor, reqdKg: Number(r['Main Fiber Reqd (kg)']) || 0 });
      yarnLines.push({ part: part, color: color, counts: counts, yarnColor: 'Lycra', reqdKg: Number(r['Lycra Reqd (kg)']) || 0 });
    } else {
      yarnLines.push({ part: part, color: color, counts: counts, yarnColor: yarnColor, reqdKg: Number(r['Fabric Required (kg)']) || 0 });
    }
  });
  yarnLines = yarnLines.filter(function (l) { return l.reqdKg > 0; });
  yarnLines.forEach(function (l) { l.reqdKg = kgRound_(l.reqdKg); });
  var yarnTotal = yarnLines.reduce(function (s, l) { return s + l.reqdKg; }, 0);

  var activeProcesses = maps.sequence.filter(function (p) { return Number(jo[maps.rateMap[p]]) > 0; });

  var rowState = rows.map(function (r) {
    return {
      part: r['Part'] || '', color: r['Color'] || '', fabricType: r['Fabric type'] || '',
      counts: r['Counts'] || '', gsm: r['GSM'] || '', qty: Number(r['Fabric Required (kg)']) || 0
    };
  });

  var stages = activeProcesses.map(function (p) {
    var lossPct = Number(jo[maps.lossMap[p]]) || 0;
    var lines = rowState.map(function (rs) {
      var reqdKg = rs.qty * (1 - lossPct / 100);
      rs.qty = reqdKg;
      return { part: rs.part, color: colorForProcess(p, rs.color), fabricType: rs.fabricType, counts: rs.counts, gsm: rs.gsm, reqdKg: reqdKg };
    });
    lines.forEach(function (l) { l.reqdKg = kgRound_(l.reqdKg); });   // whole kg in every department
    var total = lines.reduce(function (s, l) { return s + l.reqdKg; }, 0);
    return { process: p, lossPct: lossPct, rate: Number(jo[maps.rateMap[p]]) || 0, lines: lines, total: total };
  });

  // Final hop: whatever the last stage delivers still has to survive the
  // trip back to Jiwin's own store — apply that stage's own Loss % (from
  // Style Tracker's "Jiwin Inward loss %" column) with no rate, since
  // receiving your own goods back isn't job-work.
  var jiwinInwardLossPct = Number(jo['Jiwin Inward loss %']) || 0;
  if (jiwinInwardLossPct > 0) {
    var jiwinLines = rowState.map(function (rs) {
      var reqdKg = rs.qty * (1 - jiwinInwardLossPct / 100);
      rs.qty = reqdKg;
      return { part: rs.part, color: rs.color, fabricType: rs.fabricType, counts: rs.counts, gsm: rs.gsm, reqdKg: reqdKg };
    });
    jiwinLines.forEach(function (l) { l.reqdKg = kgRound_(l.reqdKg); });
    var jiwinTotal = jiwinLines.reduce(function (s, l) { return s + l.reqdKg; }, 0);
    stages.push({ process: JIWIN_INWARD_LABEL, lossPct: jiwinInwardLossPct, rate: 0, lines: jiwinLines, total: jiwinTotal });
  }

  var finishedLines = rowState.map(function (rs) {
    return { part: rs.part, color: rs.color, fabricType: rs.fabricType, counts: rs.counts, gsm: rs.gsm, reqdKg: rs.qty };
  });
  finishedLines.forEach(function (l) { l.reqdKg = kgRound_(l.reqdKg); });
  var finishedTotal = finishedLines.reduce(function (s, l) { return s + l.reqdKg; }, 0);

  var fpLink = stImgLinkForJo_(joNumber, rows), fpImg = stImgDataUrl_(fpLink);
  return {
    image: fpImg, imageWhy: fpImg ? '' : stImgWhy_(fpLink),
    consumption: consumptionForRows_(rows),
    joNumber: joNumber, buyer: first['Buyer Name'], orderOcnNo: first['Order/OCN No'],
    buyerOrderNo: first['Buyer Order No'], styleNo: first['Style Number'],
    consumptionLines: consumptionLines, consumptionTotal: consumptionTotal,
    yarnLines: yarnLines, yarnTotal: yarnTotal,
    stages: stages, finishedLines: finishedLines, finishedTotal: finishedTotal
  };
}

function generateFabricProgramHtml(d) {
  function lineRows(lines, showColor) {
    return lines.map(function (l) {
      return '<tr><td>' + (l.part) + '</td>' + (showColor ? '<td>' + (l.color) + '</td>' : '') +
        '<td>' + (l.fabricType || '-') + '</td><td>' + (l.counts) + '</td><td>' + (l.gsm || '-') + '</td>' +
        '<td>' + l.reqdKg.toLocaleString('en-IN') + '</td></tr>';
    }).join('');
  }
  var yarnRows = d.yarnLines.map(function (l) {
    return '<tr><td>' + (l.part) + '</td><td>' + (l.color) + '</td><td>' + (l.counts) + '</td><td>' + (l.yarnColor || '-') + '</td><td>' + l.reqdKg.toLocaleString('en-IN') + '</td></tr>';
  }).join('');
  var stageBlocks = d.stages.map(function (s) {
    return '<div class="title-bar" style="margin-top:16px;">' + s.process + ' — Loss ' + s.lossPct.toFixed(2) + '% — Total: <strong>' + s.total.toLocaleString('en-IN') + ' kg</strong></div>' +
      '<table><thead><tr><th>Part</th><th>Color</th><th>Fabric</th><th>Counts</th><th>GSM</th><th>Reqd Kgs</th></tr></thead><tbody>' +
      lineRows(s.lines, true) + '</tbody></table>';
  }).join('');
  return pdfShell('Fabric Program', '#8e44ad',
    '<div class="title-bar">Order/OCN No: <strong>' + (d.orderOcnNo || '-') + '</strong> | Buyer Order No: ' + (d.buyerOrderNo || '-') + ' | Style: ' + d.styleNo + '</div>' +
    (d.image ? '<div style="margin:10px 0;">' + imgTag_(d.image, 320, 220) + '</div>' : '<div style="margin:6px 0;font-size:10px;color:#B45309;">No garment photo &mdash; ' + (d.imageWhy || '') + '</div>') +
    '<div class="title-bar" style="margin-top:16px;">CONSUMPTION — Total Fabric: <strong>' + (d.consumptionTotal || 0).toLocaleString('en-IN') + ' kg</strong></div>' +
    '<table><thead><tr><th>Part</th><th>Color</th><th>Fabric</th><th>GSM</th><th>Order Qty</th><th>Excess</th><th>Total Pcs</th>' +
    '<th>Cons / Pc (g)</th><th>Cons / Dozen (kg)</th><th>Loss %</th><th>Lycra %</th><th>Fabric Reqd Kgs</th></tr></thead><tbody>' +
    (d.consumptionLines || []).map(function (l) {
      return '<tr><td>' + l.part + '</td><td>' + l.color + '</td><td>' + (l.fabricType || '-') + '</td><td>' + (l.gsm || '-') + '</td>' +
        '<td>' + l.orderQty + '</td><td>' + l.excessQty + ' (' + l.excessPct + '%)</td><td>' + l.totalPcs + '</td>' +
        '<td>' + (l.consKg * 1000).toFixed(1) + '</td><td>' + l.consDozKg.toFixed(3) + '</td><td>' + l.lossPct + '%</td>' +
        '<td>' + (l.lycra ? l.lycra + '%' : '-') + '</td><td>' + l.reqdKg.toLocaleString('en-IN') + '</td></tr>';
    }).join('') + '</tbody></table>' +
    '<div style="font-size:11px;color:#7f8c8d;margin-top:4px;">Fabric Reqd = Total Pcs x Cons / Pc x (1 + Loss %) — shown in whole kg</div>' +
    '<div class="title-bar" style="margin-top:16px;">YARN PURCHASE — Total: <strong>' + d.yarnTotal.toLocaleString('en-IN') + ' kg</strong></div>' +
    '<table><thead><tr><th>Part</th><th>Color</th><th>Counts</th><th>Yarn Colour</th><th>Reqd Kgs</th></tr></thead><tbody>' + yarnRows + '</tbody></table>' +
    stageBlocks +
    '<div class="title-bar" style="margin-top:16px;background:#f4ecf7;">FINISHED REQUIREMENT — Total: <strong>' + d.finishedTotal.toLocaleString('en-IN') + ' kg</strong></div>' +
    '<table><thead><tr><th>Part</th><th>Color</th><th>Fabric</th><th>Counts</th><th>GSM</th><th>Reqd Kgs</th></tr></thead><tbody>' + lineRows(d.finishedLines, true) + '</tbody></table>');
}

/**
 * Buyer-facing Order Sheet for one JO — header (Buyer, Order/OCN No, Buyer
 * Order No, Style, Despatch Date, Season, Merchandiser, Fabric, GSM) plus a
 * per-color size breakdown with Order Qty and Production Qty (Order Qty x
 * (1 + Excess %), matching how "P.Qty" is derived on the buyer's own sheet).
 */
/**
 * Buyer-facing Order Sheet for a whole Order/OCN No — one block per JO
 * (Style) that shares it, each with its own full color/size breakdown and
 * Accessories BOM (live from whatever's been entered — no separate step
 * needed), matching how a real multi-style buyer order sheet reads. If the
 * JO has no Order/OCN No set, it just shows that one JO on its own.
 */
/* Parts that are only a trim/secondary fabric of the garment — never the garment's own piece count.
   Used ONLY when a style has no part named BODY, so a JO whose body part is called something else
   (GARMENT, TOP, FRONT...) still gets its order quantity instead of an empty sheet. */
var ORDER_SECONDARY_PART_ = /RIB|TAPE|SERI|LOOSE|COLLAR|CUFF|BIND|ELASTIC|PIPING|WAISTBAND|LINING|POCKET|MESH|^FOLDING$/;

/* Builds the colour x size table for ONE style from the rows chosen to carry the garment quantity.
   - size columns = every size label those rows use (matched by LABEL, not by column position)
   - rows of the same colour are added together (batches), but a row that repeats an earlier row of that
     colour exactly (same sizes + quantities) is the same garment listed again (body + rib, a body in two
     fabrics) and is counted once
   - a row with no colour is kept as "-" rather than dropped
   - if no size breakdown was entered at all, the Order Quantity itself is shown (one "Qty" column)
   - every colour also carries the Order Quantity the Style Tracker holds for it, so a size breakdown that
     does not add up to it can be reported instead of passing silently */
function orderQtyBlock_(useRows) {
  var norm = function (v) { return String(v || '').trim().toUpperCase().replace(/\s+/g, ' '); };
  var sizeLabels = [], sizeIdx = {}, anySizeQty = 0, anyOrderQty = 0;
  useRows.forEach(function (r) {
    anyOrderQty += Number(r['Order Quantity']) || 0;
    for (var i = 1; i <= 10; i++) {
      var lbl = String(r['Size ' + i] || '').trim();
      if (!lbl) continue;
      if (sizeIdx[lbl.toUpperCase()] === undefined) { sizeIdx[lbl.toUpperCase()] = sizeLabels.length; sizeLabels.push({ label: lbl, key: 'Qty ' + i }); }
      anySizeQty += Number(r['Qty ' + i]) || 0;
    }
  });
  var qtyOnly = !(anySizeQty > 0) && anyOrderQty > 0;
  if (qtyOnly) { sizeLabels = [{ label: 'Qty', key: 'Order Quantity' }]; sizeIdx = { QTY: 0 }; }

  var byColor = {}, colorOrder = [];
  useRows.forEach(function (r) {
    var rawColor = String(r['Color'] || '').trim();
    var rowQty = (Number(r['Order Quantity']) || 0);
    for (var k = 1; k <= 10; k++) rowQty += Number(r['Qty ' + k]) || 0;
    if (!rawColor && !(rowQty > 0)) return;          // an empty placeholder row, not a colour
    var color = rawColor || '-', key = norm(color);
    if (!byColor[key]) { byColor[key] = { color: color, oQty: sizeLabels.map(function () { return 0; }), excessPct: 0, seen: {}, trackerQty: 0 }; colorOrder.push(key); }
    var c = byColor[key], sig = [], add = sizeLabels.map(function () { return 0; });
    if (qtyOnly) {
      var oq = Number(r['Order Quantity']) || 0;
      add[0] = oq; sig.push('QTY:' + oq);
    } else {
      for (var i = 1; i <= 10; i++) {
        var lbl = String(r['Size ' + i] || '').trim();
        if (!lbl) continue;
        var q = Number(r['Qty ' + i]) || 0;
        add[sizeIdx[lbl.toUpperCase()]] += q;
        sig.push(lbl.toUpperCase() + ':' + q);
      }
    }
    var sigKey = sig.join('|'), partN = norm(r['Part']);
    // Same colour, same sizes, same quantities again = the same garment entered twice: counted once, and
    // reported below. A row with NO colour can't be told apart from another by colour, so it is never merged.
    if (rawColor && c.seen[sigKey] !== undefined) { if (c.seen[sigKey] === partN) c.dups = (c.dups || 0) + 1; return; }
    c.seen[sigKey] = partN;
    add.forEach(function (q, idx) { c.oQty[idx] += q; });
    c.trackerQty += Number(r['Order Quantity']) || 0;
    var ex = Number(r['Excess %']) || 0;
    if (ex > c.excessPct) c.excessPct = ex;
  });

  var orderQty = 0, mismatches = [], dupNotes = [];
  var colors = colorOrder.map(function (key) {
    var c = byColor[key];
    if (c.dups) dupNotes.push({ color: c.color, extra: c.dups });
    var pQty = c.oQty.map(function (q) { return Math.round(q * (1 + c.excessPct / 100)); });
    var oTotal = c.oQty.reduce(function (sum, q) { return sum + q; }, 0);
    var pTotal = pQty.reduce(function (sum, q) { return sum + q; }, 0);
    orderQty += oTotal;
    if (!qtyOnly && c.trackerQty > 0 && Math.round(c.trackerQty) !== oTotal) mismatches.push({ color: c.color, sheet: oTotal, tracker: Math.round(c.trackerQty) });
    return { color: c.color, oQty: c.oQty, pQty: pQty, oTotal: oTotal, pTotal: pTotal, excessPct: c.excessPct, partUsed: 'BODY' };
  });
  return { sizeLabels: sizeLabels, colors: colors, orderQty: orderQty, qtyOnly: qtyOnly, mismatches: mismatches, dupNotes: dupNotes };
}

function getOrderSheetData(joNumber, lightweight) {
  /*
   * IMPORTANT ORDER SHEET RULE:
   * Only PART = BODY contributes to garment Order Qty.
   * RIB / LOOSE KNIT / FOLDING / any other part is ignored for O.Qty.
   *
   * This implementation deliberately reads Style Tracker only ONCE because
   * this function is also called from the standalone Order Sheet page.
   * Re-reading the whole sheet was causing slow/hanging requests on larger data.
   */
  var targetJo = String(joNumber || '').trim();
  if (!targetJo) return null;

  var trackerRows = getStyleTracker();
  var thisJoRows = trackerRows.filter(function (r) {
    return String(r['JO Number'] || '').trim() === targetJo && !isSampleWtRow_(r);
  });
  if (!thisJoRows.length) return null;

  var ocnNo = String(thisJoRows[0]['Order/OCN No'] || '').trim();
  var allRows = ocnNo ? trackerRows.filter(function (r) {
    return String(r['Order/OCN No'] || '').trim() === ocnNo && !isSampleWtRow_(r);
  }) : thisJoRows;

  var joOrder = [], joRowsMap = {};
  allRows.forEach(function (r) {
    var jo = String(r['JO Number'] || '').trim();
    if (!jo) return;
    if (!joRowsMap[jo]) { joOrder.push(jo); joRowsMap[jo] = []; }
    joRowsMap[jo].push(r);
  });

  function normColor_(v) {
    return String(v || '').trim().toUpperCase().replace(/\s+/g, ' ');
  }

  // One block per STYLE NUMBER within each JO. A JO often carries several styles (JO-75 has six), each
  // with its own colours and its own size run (14/16/18/18X, CH/M/G/EG, 36/38/40/42...) — lumping them
  // into one table with the first row's sizes silently dropped most of the order.
  var groups = [];
  joOrder.forEach(function (jo) {
    var order = [], map = {};
    joRowsMap[jo].forEach(function (r) {
      var sk = normColor_(r['Style Number']);
      if (!map[sk]) { map[sk] = []; order.push(sk); }
      map[sk].push(r);
    });
    order.forEach(function (sk) { groups.push({ jo: jo, styleKey: sk, rows: map[sk], styleKeysOfJo: order }); });
  });

  var bomByJo = {};   // the JO's accessories BOM, read once per JO, handed out per style below
  var styles = groups.map(function (g) {
    var jo = g.jo, rows = g.rows, first = rows[0];

    // Which rows carry the garment's piece count: the BODY part ("BODY+FOLDING" counts — the folding is
    // just part of the body fabric). If the style has no part called BODY at all (GARMENT, TOP, FRONT...),
    // every part that is not a rib / tape / trim fabric. RIB, LOOSE KNIT, NECK TAPE etc. never add to it.
    var partOf = function (r) { return String(r['Part'] || '').trim().toUpperCase().replace(/\s+/g, ' '); };
    var useRows = rows.filter(function (r) { return /^BODY(?![A-Z0-9])/.test(partOf(r)); });
    var basis = 'BODY';
    if (!useRows.length) { useRows = rows.filter(function (r) { return !ORDER_SECONDARY_PART_.test(partOf(r)); }); basis = 'GARMENT'; }
    var qb = orderQtyBlock_(useRows);
    var sizeLabels = qb.sizeLabels, colors = qb.colors, joOrderQty = qb.orderQty;

    // Dashboard/detail view can request a lightweight response. This avoids
    // expensive Accessories BOM recalculation and Drive image downloading
    // while still keeping the full Order Sheet/PDF output unchanged.
    var accessories = [];
    var imgLink = '';
    var imgData = '';
    var imageWhy = '';
    if (!lightweight) {
      if (!bomByJo[jo]) bomByJo[jo] = getAccessoriesBOM(jo);
      // A BOM line belongs to the style it was planned for; one with no style (or a style that isn't in
      // this JO) goes with the JO's first style, so nothing is listed twice and nothing goes missing.
      accessories = bomByJo[jo].filter(function (a) {
        var bk = normColor_(a['Style Number']);
        var target = (bk && g.styleKeysOfJo.indexOf(bk) > -1) ? bk : g.styleKeysOfJo[0];
        return target === g.styleKey;
      }).map(function (a) {
        return {
          name: a['Trims / Accessories'] || '', color: a['Color'] || '', size: a['Size'] || '',
          qtyNeeded: Number(a['Qty Needed'] || 0), perGarments: Number(a['Per How Many Garments'] || 0),
          requiredQty: Number(a['Required Quantity'] || 0), rate: Number(a['Rate (Rs/unit)'] || 0),
          amount: Number(a['Amount'] || 0), ihNeededDate: a['I/H Needed Date'] || ''
        };
      });
      imgLink = stImgLinkForJo_(jo, rows);
      imgData = stImgDataUrl_(imgLink);
      imageWhy = imgData ? '' : stImgWhy_(imgLink);
    }

    return {
      image: imgData, imageWhy: imageWhy,
      consumption: consumptionForRows_(rows),
      jo: jo, styleNo: first['Style Number'], styleDesc: first['Style Description'],
      fabricType: first['Fabric type'], gsm: first['GSM'], despatchDate: first['Despatch Date'],
      sizeLabels: sizeLabels, colors: colors, orderQty: joOrderQty, accessories: accessories,
      basis: basis, qtyOnly: qb.qtyOnly, mismatches: qb.mismatches, dupNotes: qb.dupNotes, rowsUsed: useRows.length, _rows: rows
    };
  });

  // SAFETY NET: if no style of a JO produced any quantity (an unusual part naming, a rib-only JO...),
  // rebuild that JO from every row it has rather than print an empty Order Sheet.
  var joHasQty = {};
  styles.forEach(function (st) { if (st.orderQty > 0) joHasQty[st.jo] = true; });
  styles.forEach(function (st) {
    if (joHasQty[st.jo] || st.orderQty > 0) return;
    var qb = orderQtyBlock_(st._rows);
    if (qb.orderQty > 0) {
      st.sizeLabels = qb.sizeLabels; st.colors = qb.colors; st.orderQty = qb.orderQty; st.basis = 'ALL'; st.qtyOnly = qb.qtyOnly; st.mismatches = qb.mismatches; st.dupNotes = qb.dupNotes;
    }
  });
  styles.forEach(function (st) { if (st.orderQty > 0) joHasQty[st.jo] = true; });

  // A "style" that is only a rib / neck-tape / trim fabric under its own style code has no quantity of
  // its own — leave it out rather than print an empty table, as long as another style of the JO has it.
  styles = styles.filter(function (st) { return st.colors.length || !joHasQty[st.jo]; });

  // Never silent: anything that does not add up, or no quantity at all, is written on the sheet.
  styles.forEach(function (st) {
    st.warnings = (st.mismatches || []).map(function (m) {
      return 'Style ' + st.styleNo + ' / ' + m.color + ': sizes add up to ' + m.sheet + ' but the Style Tracker Order Quantity is ' + m.tracker + ' — correct whichever is wrong.';
    });
    // not a problem on the printed sheet (a part can legitimately carry two consumption lines) — just
    // noted on the dashboard so a genuine double entry can be spotted
    st.notes = (st.dupNotes || []).map(function (n) {
      return 'Style ' + st.styleNo + ' / ' + n.color + ' is on ' + (n.extra + 1) + ' rows with identical quantities — counted once.';
    });
    if (!(st.orderQty > 0)) st.warnings.push('Style ' + st.styleNo + ': no order quantity found in the Style Tracker rows for this style.');
    delete st._rows;
  });

  var totalOrderQty = styles.reduce(function (sum, st) { return sum + st.orderQty; }, 0);
  var firstJo = thisJoRows[0];

  return {
    joNumber: targetJo,
    buyer: firstJo['Buyer Name'],
    orderOcnNo: ocnNo,
    buyerOrderNo: firstJo['Buyer Order No'],
    merchandiser: firstJo['Merchandiser'],
    season: firstJo['Season'],
    totalOrderQty: totalOrderQty,
    styles: styles
  };
}

/* Run from the Apps Script editor (or any time): walks EVERY JO and lists the ones whose Order Sheet has
   no quantity or whose size breakdown does not add up to the Order Quantity in Style Tracker, so a
   missing/wrong quantity is found here instead of on a printed sheet. Returns [] when everything is fine. */
function checkOrderSheetQuantities() {
  var seen = {}, out = [];
  getStyleTracker().forEach(function (r) {
    var jo = String(r['JO Number'] || '').trim();
    if (!jo || seen[jo]) return;
    seen[jo] = true;
    try {
      var d = getOrderSheetData(jo, true);
      if (!d) return;
      var mine = d.styles.filter(function (st) { return String(st.jo).trim() === jo; });
      var notes = [];
      mine.forEach(function (st) { (st.warnings || []).forEach(function (w) { notes.push(w); }); });
      var qty = mine.reduce(function (a, st) { return a + st.orderQty; }, 0);
      if (notes.length || !(qty > 0)) out.push({ jo: jo, orderQty: qty, problems: notes.length ? notes : ['no order quantity found'] });
    } catch (e) { out.push({ jo: jo, orderQty: 0, problems: ['could not be read: ' + e.message] }); }
  });
  Logger.log(out.length ? JSON.stringify(out, null, 1) : 'All JOs: Order Sheet quantities look right.');
  return out;
}

/* A JO that carries several styles repeats colours (JO-73: WASHED NAVY is in SPG019 AND SPG004). Each style has
   its own size table, but dyeing, the fabric program and the buyer all work colour by colour — so this adds a
   colour up across every style of the JO: [{color, oTotal, pTotal}], in order of first appearance. */
function orderColourTotals_(styles) {
  var key = function (v) { return String(v || '').trim().toUpperCase().replace(/\s+/g, ' '); };
  var order = [], map = {};
  styles.forEach(function (st) {
    st.colors.forEach(function (c) {
      var k = key(c.color);
      if (!map[k]) { map[k] = { color: String(c.color || '-').trim() || '-', oTotal: 0, pTotal: 0, styles: [] }; order.push(k); }
      map[k].oTotal += c.oTotal; map[k].pTotal += c.pTotal;
      if (st.styleNo && map[k].styles.indexOf(st.styleNo) < 0) map[k].styles.push(st.styleNo);
    });
  });
  return order.map(function (k) { return map[k]; });
}
function colourTotalsHtml_(styles, esc, cls) {
  var e = esc || function (v) { return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  var rows = orderColourTotals_(styles), o = 0, p = 0;
  var body = rows.map(function (r) {
    o += r.oTotal; p += r.pTotal;
    return '<tr><td class="l"><b>' + e(r.color) + '</b></td><td class="l">' + e(r.styles.join(', ')) + '</td><td><b>' + r.oTotal + '</b></td><td>' + r.pTotal + '</td></tr>';
  }).join('');
  return '<table class="' + (cls || 'os') + '" style="margin-top:6px"><thead><tr><th class="l">Colour &mdash; total of all styles</th><th class="l">In styles</th><th>O.Qty</th><th>P.Qty</th></tr></thead><tbody>' + body +
    '<tr class="tot"><td class="l" colspan="2">TOTAL</td><td>' + o + '</td><td>' + p + '</td></tr></tbody></table>';
}

/* The red "quantity check" box printed under an Order Sheet when something does not add up. */
function qtyWarnHtml_(warnings, esc) {
  var e = esc || function (v) { return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  return '<div style="border:1px solid #DC2626;background:#FEF2F2;color:#991B1B;padding:3px 6px;margin:4px 0;font-size:9px;"><b>QUANTITY CHECK:</b> ' +
    warnings.map(function (w) { return e(w); }).join('<br>') + '</div>';
}

function generateOrderSheetHtml(d) {
  function styleBlock(st) {
    var sizeHeaderCells = st.sizeLabels.map(function (s) { return '<th>' + s.label + '</th>'; }).join('') + '<th>Total</th>';
    var colorRows = st.colors.map(function (c) {
      var oCells = c.oQty.map(function (q) { return '<td>' + q + '</td>'; }).join('') + '<td><strong>' + c.oTotal + '</strong></td>';
      var pCells = c.pQty.map(function (q) { return '<td>' + q + '</td>'; }).join('') + '<td><strong>' + c.pTotal + '</strong></td>';
      return '<tr><td rowspan="2" style="font-weight:bold;">' + (c.color || '-') + '</td><th>O.Qty</th>' + oCells + '</tr>' +
        '<tr><th>P.Qty (+' + c.excessPct + '%)</th>' + pCells + '</tr>';
    }).join('');
    var accRows = st.accessories.map(function (a) {
      return '<tr><td>' + a.name + '</td><td>' + (a.color || '-') + '</td><td>' + (a.size || '-') + '</td><td>' + a.qtyNeeded +
        '</td><td>' + a.perGarments + '</td><td><strong>' + a.requiredQty + '</strong></td><td>' + (a.ihNeededDate || '-') + '</td></tr>';
    }).join('');
    var accTotal = st.accessories.reduce(function (s, a) { return s + a.amount; }, 0);
    return '<div class="title-bar" style="margin-top:20px;">JO: <strong>' + st.jo + '</strong> | Style No: <strong>' + st.styleNo +
      '</strong> | Despatch Date: ' + (st.despatchDate || '-') + '</div>' +
      (st.image ? '<div style="margin:8px 0;">' + imgTag_(st.image, 320, 220) + '</div>' : '<div style="margin:6px 0;font-size:10px;color:#B45309;">No garment photo &mdash; ' + (st.imageWhy || '') + '</div>') +
      '<table><tr><th>Fabric</th><td>' + (st.fabricType || '-') + '</td><th>GSM</th><td>' + (st.gsm || '-') + '</td><th>Order Qty</th><td><strong>' + st.orderQty + ' PCS</strong></td></tr></table>' +
      '<div style="margin-top:8px;font-weight:bold;">Consumption per garment (grams)</div>' + consTableHtml_(st.consumption, '') +
      '<p style="margin:8px 0;">' + (st.styleDesc || '') + '</p>' +
      '<table><thead><tr><th>Color</th><th></th>' + sizeHeaderCells + '</tr></thead><tbody>' + colorRows + '</tbody></table>' +
      ((st.warnings && st.warnings.length) ? qtyWarnHtml_(st.warnings) : '') +
      (st.accessories.length ?
        '<div style="margin-top:10px;font-weight:bold;">Accessories BOM</div>' +
        '<table><thead><tr><th>Trims / Accessories</th><th>Color</th><th>Size</th><th>Qty Needed</th><th>Per Garments</th><th>Required Qty</th><th>I/H Needed Date</th></tr></thead>' +
        '<tbody>' + accRows + '</tbody></table>'
        : '<div style="margin-top:6px;color:#7f8c8d;font-size:11px;">No accessories entered yet for this JO.</div>');
  }
  return pdfShell('Order Sheet', '#1a5276',
    '<div class="title-bar">Order/OCN No: <strong>' + (d.orderOcnNo || '-') + '</strong> | Buyer Order No: ' + (d.buyerOrderNo || '-') + '</div>' +
    '<table><tr><th>Total Order Qty (all styles)</th><td><strong>' + d.totalOrderQty + ' PCS</strong></td></tr></table>' +
    d.styles.map(function (st, i) {
      var html = styleBlock(st);
      var sameJo = d.styles.filter(function (x) { return x.jo === st.jo; });
      var isLast = sameJo[sameJo.length - 1] === st;
      if (sameJo.length > 1 && isLast) {
        html += '<div class="title-bar" style="margin-top:20px;">JO: <strong>' + st.jo + '</strong> | Colour-wise total of all ' + sameJo.length + ' styles</div>' + colourTotalsHtml_(sameJo, null, '');
      }
      return html;
    }).join(''));
}

/**
 * A DC is billable to whoever performed the process it just came FROM — the
 * "From Supplier" who ran "From Process" and is now shipping that output
 * onward (its own DC Issue Qty KG is what they delivered and get paid for).
 * The very first DC in a chain (From = Yarn/Cora Fabric/Dyed Fabric Store)
 * is never billable — nothing's been processed yet at that point.
 */
var CHAIN_STORE_LABELS = [YARN_STORE_LABEL, CORA_STORE_LABEL, DYED_STORE_LABEL];

function getSuppliersWithBillableDCs() {
  var suppliers = {};
  sheetToObjects(SHEETS.PROCESS_DC).forEach(function (r) {
    if (r['Status'] !== 'Approved') return;
    if (CHAIN_STORE_LABELS.indexOf(r['From Process']) !== -1) return;
    var supplier = String(r['From Supplier'] || '').trim();
    if (!supplier) return;
    suppliers[supplier] = true;
  });
  return Object.keys(suppliers);
}

/* Every job-work supplier for the Bill Passing dropdown: those with DCs ready to bill AND those who
   were sent work that has not come out yet (so the list is never empty after the first DC). */
function getBillPassingSuppliers() {
  var ready = {}, atSupplier = {};
  getSuppliersWithBillableDCs().forEach(function (s) { ready[s] = getBillableDCsForSupplier(s).length; });
  sheetToObjects(SHEETS.PROCESS_DC).forEach(function (r) {
    if ((r['Status'] !== 'Approved' && r['Status'] !== 'Billed') || r['To Process'] === JIWIN_INWARD_LABEL) return;
    var s = String(r['To Supplier'] || '').trim(); if (!s) return;
    if (!(s in ready)) ready[s] = ready[s] || 0;
    if (!r['Balance Status'] && getDCChainBalance(r['DC No']) > 0.01) atSupplier[s] = (atSupplier[s] || 0) + 1;
  });
  return Object.keys(ready).sort().map(function (s) { return { name: s, ready: ready[s] || 0, atSupplier: atSupplier[s] || 0 }; });
}

/** Approved, not-yet-Billed DCs for one supplier — across every JO — for Bill Passing's DC picker. */
function getBillableDCsForSupplier(supplier) {
  return sheetToObjects(SHEETS.PROCESS_DC).filter(function (r) {
    if (r['Status'] !== 'Approved') return false;
    if (CHAIN_STORE_LABELS.indexOf(r['From Process']) !== -1) return false;
    if (String(r['From Supplier'] || '').trim() !== String(supplier).trim()) return false;
    // A DC with a real leftover balance can't be billed until that's
    // resolved — Return Requested blocks outright; a blank Balance Status
    // with a genuine un-forwarded balance also blocks (force a decision);
    // Closed/Debited/Pending Approval (already decided, awaiting sign-off)
    // are fine to bill against.
    // Leftover that matters is on the lot this supplier RECEIVED (Previous DC No), not on the DC
    // being billed (that one's balance sits with the next supplier / Jiwin Inward). Part-deliveries
    // are billable; only a lot marked "Return Requested" is held back. Open leftovers show in the
    // supplier balance card for a Close / Debit / Return decision.
    var prev = String(r['Previous DC No'] || '').trim();
    if (prev) {
      var src = sheetToObjects(SHEETS.PROCESS_DC).filter(function (x) { return x['DC No'] === prev; })[0];
      if (src && src['Balance Status'] === 'Return Requested') return false;
    }
    if (r['Balance Status'] === 'Return Requested') return false;
    return true;
  });
}

/**
 * Fabric left sitting un-forwarded at a supplier — every Approved DC that
 * supplier RECEIVED (To Supplier) whose output hasn't fully moved to the
 * next stage yet, and hasn't already been resolved (Closed/Debited/Return
 * Requested). This is exactly what shows up at Bill Passing time: the
 * knitter/dyer/etc. was sent X kg, only some of it has been forwarded
 * onward — the rest may be a real weight loss, or they may still have it.
 */
function getSupplierFabricBalance(supplier) {
  return sheetToObjects(SHEETS.PROCESS_DC).filter(function (r) {
    if ((r['Status'] !== 'Approved' && r['Status'] !== 'Billed')) return false;
    if (String(r['To Supplier'] || '').trim() !== String(supplier).trim()) return false;
    if (r['Balance Status']) return false; // already resolved
    return true;
  }).map(function (dc) {
    var remaining = getDCChainBalance(dc['DC No']);
    return {
      dcNo: dc['DC No'], date: dc['Date'], joNumber: dc['JO Number'], color: dc['Color'] || '',
      counts: dc['Counts'] || '', content: dc['Content'] || '', gsm: dc['GSM'] || '',
      toProcess: dc['To Process'], issuedKg: Number(dc['DC Issue Qty KG']) || 0, remainingKg: remaining
    };
  }).filter(function (row) { return row.remainingKg > 0.01; });
}

/**
 * Resolves one DC's leftover un-forwarded balance, three ways:
 * - 'CLOSE': accepted as-is (e.g. within tolerance, or will move later) — just
 *   recorded, no financial or physical follow-up.
 * - 'DEBIT': treated as unaccounted loss — raises a Debit Note against the
 *   supplier for remainingKg x rate, same mechanism as excess weight loss.
 * - 'RETURN': flags that the supplier still has this fabric and should send
 *   it back — recorded for follow-up, not billed or written off.
 */
function resolveDcBalance(dcNo, action, rate, note) {
  var rowIdx = findRowIndex(SHEETS.PROCESS_DC, 'DC No', dcNo);
  if (rowIdx === -1) return { success: false, error: 'DC not found: ' + dcNo };
  var dc = sheetToObjects(SHEETS.PROCESS_DC).find(function (r) { return r['DC No'] === dcNo; });
  var remaining = getDCChainBalance(dcNo);
  if (remaining <= 0.01) return { success: false, error: 'DC ' + dcNo + ' has no un-forwarded balance left to resolve.' };

  if (action === 'CLOSE') {
    // Closing now requires admin approval before it's final — until then the
    // DC stays un-billable, same as any other unresolved balance.
    setCell(SHEETS.PROCESS_DC, rowIdx, 'Balance Status', 'Pending Approval');
    setCell(SHEETS.PROCESS_DC, rowIdx, 'Balance Note', note || '');
    return { success: true, message: 'Close requested for DC ' + dcNo + '\'s balance of ' + remaining.toFixed(2) + ' kg — awaiting approval.' };
  }

  if (action === 'DEBIT') {
    var debitRate = Number(rate) || 0;
    if (debitRate <= 0) return { success: false, error: 'Enter a Rate (Rs/kg) to raise a Debit Note.' };
    var debitNoteNo = nextId('DN-', SHEETS.DEBIT_NOTE, 1);
    var debitAmount = remaining * debitRate;
    appendRowSafe_(SHEETS.DEBIT_NOTE, [
      debitNoteNo, Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      dc['JO Number'], dc['To Supplier'], dcNo, '', remaining, debitRate, debitAmount, 'Raised'
    ]);
    setCell(SHEETS.PROCESS_DC, rowIdx, 'Balance Status', 'Debited');
    setCell(SHEETS.PROCESS_DC, rowIdx, 'Balance Note', note || '');
    return {
      success: true, debitNoteNo: debitNoteNo,
      message: 'Debit Note ' + debitNoteNo + ' raised for ' + remaining.toFixed(2) + ' kg (₹' + debitAmount.toFixed(2) + ') against ' + dc['To Supplier'] + '. This DC is now billable — the debited amount is already excluded from what it can be billed for.'
    };
  }

  if (action === 'RETURN') {
    // A Return Requested DC stays un-billable until someone explicitly
    // clears it (once the fabric actually comes back) via clearDcReturn.
    setCell(SHEETS.PROCESS_DC, rowIdx, 'Balance Status', 'Return Requested');
    setCell(SHEETS.PROCESS_DC, rowIdx, 'Balance Note', note || '');
    return { success: true, message: 'Return requested from ' + dc['To Supplier'] + ' for ' + remaining.toFixed(2) + ' kg (DC ' + dcNo + '). This DC cannot be billed until the fabric is returned and this is cleared.' };
  }

  return { success: false, error: 'Unknown action: ' + action };
}

/** Approves or rejects a "Close" balance request — Approve finalizes it as Closed; Reject sends it back to unresolved so it can be re-decided. */
function approveDcBalance(dcNo, decision, approver) {
  var rowIdx = findRowIndex(SHEETS.PROCESS_DC, 'DC No', dcNo);
  if (rowIdx === -1) return { success: false, error: 'DC not found: ' + dcNo };
  if (decision === 'APPROVE') {
    setCell(SHEETS.PROCESS_DC, rowIdx, 'Balance Status', 'Closed');
    setCell(SHEETS.PROCESS_DC, rowIdx, 'Approved By', approver || '');
    return { success: true, message: 'DC ' + dcNo + '\'s balance close approved.' };
  }
  setCell(SHEETS.PROCESS_DC, rowIdx, 'Balance Status', '');
  return { success: true, message: 'DC ' + dcNo + '\'s balance close was rejected — back to unresolved.' };
}

/** Once fabric requested back from a supplier actually arrives, clears the flag so the DC becomes billable/forwardable again. */
function clearDcReturn(dcNo) {
  var rowIdx = findRowIndex(SHEETS.PROCESS_DC, 'DC No', dcNo);
  if (rowIdx === -1) return { success: false, error: 'DC not found: ' + dcNo };
  setCell(SHEETS.PROCESS_DC, rowIdx, 'Balance Status', '');
  return { success: true, message: 'DC ' + dcNo + ' return cleared — billable again.' };
}

function getJoSupplierPairs() {
  var pairs = {};
  sheetToObjects(SHEETS.PROCESS_DC).forEach(function (r) {
    if (r['Status'] !== 'Approved') return;
    var jo = String(r['JO Number']).trim(), supplier = String(r['To Supplier']).trim();
    if (!jo || !supplier) return;
    if (!pairs[jo]) pairs[jo] = {};
    pairs[jo][supplier] = true;
  });
  var result = {};
  Object.keys(pairs).forEach(function (jo) { result[jo] = Object.keys(pairs[jo]); });
  return result;
}

function getDCsForJoSupplier(joNumber, supplier) {
  return sheetToObjects(SHEETS.PROCESS_DC).filter(function (r) {
    return r['Status'] === 'Approved' && String(r['JO Number']).trim() === String(joNumber).trim() && String(r['To Supplier']).trim() === String(supplier).trim();
  });
}

function saveBillPassing(d) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var dcNos;
    try { dcNos = JSON.parse(d.dcNos || '[]'); } catch (e) { dcNos = d.dcNo ? [d.dcNo] : []; }
    if (!dcNos.length) return { success: false, error: 'Select at least one DC to bill.' };

    var alreadyBilledSet = {};
    sheetToObjects(SHEETS.BILL_PASSING).forEach(function (b) {
      if (b['Status'] === 'Rejected') return;
      String(b['DC Numbers'] || '').split(',').map(function (s) { return s.trim(); }).forEach(function (n) { if (n) alreadyBilledSet[n] = true; });
    });

    var allDCs = sheetToObjects(SHEETS.PROCESS_DC);
    var debitedByDc = {};
    sheetToObjects(SHEETS.DEBIT_NOTE).forEach(function (dn) {
      if (dn['DC No']) debitedByDc[dn['DC No']] = (debitedByDc[dn['DC No']] || 0) + (Number(dn['Amount']) || 0);
    });

    var lines = [];
    var joSet = {};
    var totalQty = 0, totalExpected = 0;
    for (var i = 0; i < dcNos.length; i++) {
      var dcNo = dcNos[i];
      var dc = allDCs.find(function (r) { return r['DC No'] === dcNo; });
      if (!dc) return { success: false, error: 'DC "' + dcNo + '" not found.' };
      if (dc['Status'] !== 'Approved') return { success: false, error: 'DC ' + dcNo + ' is "' + dc['Status'] + '" — only an Approved DC can be billed.' };
      if (String(dc['From Supplier'] || '').trim() !== String(d.supplier).trim()) {
        return { success: false, error: 'DC ' + dcNo + ' was processed by ' + dc['From Supplier'] + ', not ' + d.supplier + '.' };
      }
      if (dc['Balance Status'] === 'Return Requested') {
        return { success: false, error: 'DC ' + dcNo + ' has a Return Requested — it cannot be billed until the fabric is returned and that\'s cleared.' };
      }
      if (!dc['Balance Status'] && getDCChainBalance(dcNo) > 0.01) {
        return { success: false, error: 'DC ' + dcNo + ' has an unresolved fabric balance — Close, Debit, or Return it first before billing.' };
      }
      if (alreadyBilledSet[dcNo]) return { success: false, error: 'DC ' + dcNo + ' has already been billed (or is pending billing).' };

      var qty = Number(dc['DC Issue Qty KG']) || 0;
      var rate = getRateForJOProcess(dc['JO Number'], dc['From Process']);
      var lineExpected = Math.max(0, qty * rate - (debitedByDc[dcNo] || 0));
      lines.push({ dcNo: dcNo, joNumber: dc['JO Number'], qty: qty, rate: rate, expected: lineExpected });
      joSet[dc['JO Number']] = true;
      totalQty += qty;
      totalExpected += lineExpected;
    }

    var billAmount = Number(d.billAmount) || 0;
    if (billAmount > totalExpected + 0.01) {
      return {
        success: false,
        error: 'Bill Amount (₹' + billAmount.toFixed(2) + ') cannot exceed the combined Qty × Rate across all selected DCs (₹' + totalExpected.toFixed(2) + ', already net of any Debit Notes).'
      };
    }
    var variance = billAmount - totalExpected;
    var blendedRate = totalQty > 0 ? (totalExpected / totalQty) : 0; // for display only — real per-DC rates may differ across JOs
    var joNumber = Object.keys(joSet).join(', ');
    var dcNumbersStr = dcNos.join(', ');

    var gstPct = Number(d.gstPct) || 0;
    var gstAmount = billAmount * gstPct / 100;
    var totalInclGst = billAmount + gstAmount;
    // If the person typed in what the physical bill/invoice actually shows as
    // its final total, check our own Bill Amount + GST against it — catches
    // a mistyped Bill Amount or GST % before it goes to approval.
    var billStatedTotal = d.billStatedTotal !== undefined && d.billStatedTotal !== '' ? Number(d.billStatedTotal) : null;
    var gstMatch = '';
    if (billStatedTotal !== null) {
      gstMatch = Math.abs(totalInclGst - billStatedTotal) <= 0.5 ? 'Match' : 'Mismatch';
    }

    var billNo = nextId('BILL-', SHEETS.BILL_PASSING, 1);
    appendRowSafe_(SHEETS.BILL_PASSING, [
      billNo, d.billDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      d.supplier, joNumber, dcNumbersStr, totalQty, blendedRate, totalExpected, billAmount, variance,
      d.remarks || '', 'Pending Approval', '', '', '', '',
      gstPct, gstAmount, totalInclGst, billStatedTotal === null ? '' : billStatedTotal, gstMatch, d.supplierInvoiceNo || ''
    ]);
    var gstNote = billStatedTotal !== null
      ? (gstMatch === 'Match'
        ? ' GST check: matches the bill (₹' + totalInclGst.toFixed(2) + ').'
        : ' GST check: MISMATCH — our total (₹' + totalInclGst.toFixed(2) + ') vs bill\'s stated total (₹' + billStatedTotal.toFixed(2) + ').')
      : '';
    return {
      success: true, billNo: billNo, expectedAmount: totalExpected, variance: variance, gstMatch: gstMatch,
      message: 'Bill ' + billNo + ' submitted for approval — ' + dcNos.length + ' DC(s) across ' + Object.keys(joSet).length + ' JO(s). Expected ₹' + totalExpected.toFixed(2) + ', variance ₹' + variance.toFixed(2) + '.' + gstNote
    };
  } catch (err) {
    return { success: false, error: err.message };
  } finally { lock.releaseLock(); }
}

function approveBillPassing(billNo, decision, approver, reason) {
  var rowIdx = findRowIndex(SHEETS.BILL_PASSING, 'Bill No', billNo);
  if (rowIdx === -1) return { success: false, error: 'Bill not found: ' + billNo };
  var status = decision === 'APPROVE' ? 'Approved' : 'Rejected';
  setCell(SHEETS.BILL_PASSING, rowIdx, 'Status', status);
  setCell(SHEETS.BILL_PASSING, rowIdx, 'Approved By', approver || '');
  setCell(SHEETS.BILL_PASSING, rowIdx, 'Approval Date', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  setCell(SHEETS.BILL_PASSING, rowIdx, 'Rejection Reason', decision === 'REJECT' ? (reason || '') : '');

  var result = { success: true, message: 'Bill ' + billNo + ' ' + status + '.' };
  if (decision === 'APPROVE') {
    var bill = sheetToObjects(SHEETS.BILL_PASSING).find(function (r) { return r['Bill No'] === billNo; });
    var dcNumbers = String(bill['DC Numbers'] || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
    dcNumbers.forEach(function (dcNo) {
      var dcRow = findRowIndex(SHEETS.PROCESS_DC, 'DC No', dcNo);
      if (dcRow !== -1) setCell(SHEETS.PROCESS_DC, dcRow, 'Status', 'Billed');
    });
    var wl = generateWeightLoss(billNo);
    result.weightLoss = wl;
  }
  return result;
}

function getBillPassingQueue() { return sheetToObjects(SHEETS.BILL_PASSING); }

/*************** 15. WEIGHT / LOSS + DEBIT NOTE ***************/
function generateWeightLoss(billNo) {
  var bill = sheetToObjects(SHEETS.BILL_PASSING).find(function (r) { return r['Bill No'] === billNo; });
  if (!bill) return { success: false, error: 'Bill not found.' };
  var dcNumbers = String(bill['DC Numbers'] || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
  var createdRows = [];

  dcNumbers.forEach(function (dcNo) {
    var dc = sheetToObjects(SHEETS.PROCESS_DC).find(function (r) { return r['DC No'] === dcNo; });
    if (!dc) return;
    // Each DC keeps its own JO's allowed loss % and process rate — a bill
    // can now span several JOs, so nothing here may assume they're all the same.
    var joInfo = getJOInfo(dc['JO Number']);
    var allowedLossPct = joInfo ? Number(joInfo['Loss %']) || 0 : 0;
    // Loss is now computed and stored on the DC itself at creation time
    // (Available KG in vs DC Issue Qty KG out for that stage) — read it
    // directly rather than re-deriving it here.
    var prevQty = Number(dc['Available KG']) || 0;
    var availQty = Number(dc['DC Issue Qty KG']) || 0;
    var lossKg = Number(dc['Loss KG']) || Math.max(0, prevQty - availQty);
    var lossPct = Number(dc['Loss %']) || (prevQty > 0 ? (lossKg / prevQty * 100) : 0);
    var result = lossPct <= allowedLossPct ? 'OK' : 'EXCESS';

    var wlNo = nextId('WL-', SHEETS.WEIGHT_LOSS, 1);
    var debitNoteNo = '';

    if (result === 'EXCESS') {
      var excessLossKg = lossKg - (prevQty * allowedLossPct / 100);
      var rate = getRateForJOProcess(dc['JO Number'], dc['From Process']);
      debitNoteNo = nextId('DN-', SHEETS.DEBIT_NOTE, 1);
      appendRowSafe_(SHEETS.DEBIT_NOTE, [
        debitNoteNo, Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        dc['JO Number'], bill['Supplier'], dcNo, billNo, excessLossKg, rate, excessLossKg * rate, 'Raised'
      ]);
    }

    appendRowSafe_(SHEETS.WEIGHT_LOSS, [
      wlNo, Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      dc['JO Number'], dcNo, billNo, prevQty, availQty, lossKg, lossPct, allowedLossPct, result, debitNoteNo
    ]);
    createdRows.push({ wlNo: wlNo, dcNo: dcNo, result: result, debitNoteNo: debitNoteNo });
  });

  return { success: true, rows: createdRows };
}

function getWeightLossReport() { return sheetToObjects(SHEETS.WEIGHT_LOSS); }
function getDebitNotes() { return sheetToObjects(SHEETS.DEBIT_NOTE); }

/*************** COMBINED APPROVALS QUEUE ***************/
function getPendingApprovals() {
  return {
    pendingPOs: (function(){var all=sheetToObjects(SHEETS.YARN_PO), groups={};all.filter(function(r){return r['Status']==='Pending Approval';}).forEach(function(r){var k=String(r['PO No']);if(!groups[k])groups[k]=[];groups[k].push(r);});return Object.keys(groups).map(function(k){var a=groups[k],o=Object.assign({},a[0]);o['JO Number']=Array.from(new Set(a.map(function(r){return r['JO Number'];}).filter(Boolean))).join(', ');o['Counts']=Array.from(new Set(a.map(function(r){return r['Counts'];}).filter(Boolean))).join(', ');o['Qty KG']=a.reduce(function(s,r){return s+(Number(r['Qty KG'])||0);},0);o['PO Value']=a.reduce(function(s,r){return s+(Number(r['PO Value'])||0);},0);o['Rate (Rs/kg)']=o['Qty KG']?o['PO Value']/o['Qty KG']:0;o['Line Count']=a.length;return o;});})(),
    pendingDCs: sheetToObjects(SHEETS.PROCESS_DC).filter(function (r) { return r['Status'] === 'Pending Approval'; }),
    pendingBills: sheetToObjects(SHEETS.BILL_PASSING).filter(function (r) { return r['Status'] === 'Pending Approval'; }),
    pendingAccessoryPOs: sheetToObjects(SHEETS.ACCESSORIES_PO).filter(function (r) { return r['Status'] === 'Pending Approval'; }),
    pendingAccessoryBills: sheetToObjects(SHEETS.ACCESSORIES_BILL).filter(function (r) { return r['Status'] === 'Pending Approval'; }),
    pendingDcBalances: sheetToObjects(SHEETS.PROCESS_DC).filter(function (r) { return r['Balance Status'] === 'Pending Approval'; }),
    pendingAccExtras: getPendingAccExtras_(),
    pendingYarnExtras: getPendingYarnExtras_(),
    pendingYarnBills: sheetToObjects(SHEETS.YARN_BILL).filter(function (r) { return r['Status'] === 'Pending Approval'; }),
    pendingCmtBills: sheetToObjects(SHEETS.CMT_BILL).filter(function (r) { return r['Status'] === 'Pending Approval'; }),
    pendingGeneralBills: sheetToObjects(SHEETS.GENERAL_BILL).filter(function (r) { return r['Status'] === 'Pending Approval'; })
  };
}

/*************** PDFs ***************/
/** Returns the full HTML for an Order Sheet / Fabric Program so the client can preview it on-screen (an iframe) before deciding to save as PDF. */
function previewOrderSheet(joNumber) {
  var fz = lockedSheetHtml_('ORDERSHEET', joNumber);
  if (fz && !fz.missing) return fz.html;
  var data = getOrderSheetData(joNumber);
  return data ? generateOrderSheetHtml(data) : null;
}
function previewDc(dcNo) {
  var dc = sheetToObjects(SHEETS.PROCESS_DC).find(function (r) { return r['DC No'] === dcNo; });
  return dc ? generateDcHtml(dc) : null;
}
function previewFabricProgram(joNumber) {
  var fz = lockedSheetHtml_('FABRICPROGRAM', joNumber);
  if (fz && !fz.missing) return fz.html;
  var data = getFabricProgramData(joNumber);
  return data ? generateFabricProgramHtml(data) : null;
}

/* Same documents as exportDocumentPdf, but returns the HTML itself for viewing in a browser tab
   instead of converting to a PDF to download — "View" without "Download". Scoped to the JO-based
   documents (Order Sheet, Fabric Program, Fabric Consumption, Order + Fabric); everything else
   still works as Download only. */
/* Yarn / Cora Fabric / Dyed Fabric PO -> printable HTML (shared by View and Download).
   Rejected lines are left out; Approved and Billed (after bill passing) lines can be printed;
   anything still Pending Approval blocks printing. */
function buildPrintablePoHtml_(id) {
  var poNo = String(id || '').trim();
  var poLines = sheetToObjects(SHEETS.YARN_PO).filter(function (r) { return String(r['PO No']).trim() === poNo; });
  if (!poLines.length) throw new Error('PO "' + poNo + '" not found.');
  var live = poLines.filter(function (p) { return String(p['Status'] || '').trim() !== 'Rejected'; });
  if (!live.length) throw new Error('PO ' + poNo + ' is Rejected — nothing to print.');
  var notReady = live.filter(function (p) { var st = String(p['Status'] || '').trim(); return st !== 'Approved' && st !== 'Billed'; });
  if (notReady.length) throw new Error('PO ' + poNo + ' is "' + (notReady[0]['Status'] || 'blank') + '" — only an Approved PO can be printed.');
  return live.length > 1 ? generatePoMultiHtml_(live) : generatePoHtml(live[0]);
}

function viewDocumentHtml(type, id) {
  try {
    var htmlContent = '';
    var frozen = lockedSheetHtml_(type, id);
    if (frozen && !frozen.missing) return { success: true, html: frozen.html, savedCopy: { at: frozen.at, by: frozen.by } };
    var note = frozen && frozen.missing ? 'The saved copy of this sheet could not be found in Drive — showing a freshly generated one instead.' : '';
    if (type === 'PO') {
      htmlContent = buildPrintablePoHtml_(id);
    } else if (type === 'ORDERSHEET') {
      var osData = getOrderSheetData(id);
      if (!osData) throw new Error('JO "' + id + '" not found in Style Tracker.');
      htmlContent = generateOrderSheetHtml(osData);
    } else if (type === 'FABRICPROGRAM') {
      var fpData = getFabricProgramData(id);
      if (!fpData) throw new Error('JO "' + id + '" not found in Style Tracker.');
      htmlContent = generateFabricProgramHtml(fpData);
    } else if (type === 'PROCESSREQ') {
      htmlContent = generateProcessReqHtml(id);
    } else if (type === 'ORDERFABRIC') {
      htmlContent = generateOrderFabricHtml(id);
    } else {
      return { success: false, error: 'Viewing isn\'t available for this document type yet — use Download instead.' };
    }
    return { success: true, html: htmlContent, note: note };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function exportDocumentPdf(type, id) {
  try {
    var htmlContent = '', fileName = '', frozenNote = '';
    var frozen = lockedSheetHtml_(type, id);
    if (frozen && !frozen.missing) {
      // a locked JO's sheet: the PDF is made from the saved copy, never from today's data/layout
      fileName = ({ ORDERSHEET: 'Order_Sheet_', PROCESSREQ: 'Fabric_Program_', FABRICPROGRAM: 'Fabric_Program_', ORDERFABRIC: 'Order_and_Fabric_' })[type] + id;
      htmlContent = frozen.html;
    } else if (type === 'PO') {
      htmlContent = buildPrintablePoHtml_(id);
      var poTypeName = String((sheetToObjects(SHEETS.YARN_PO).find(function (r) { return String(r['PO No']).trim() === String(id).trim(); }) || {})['PO Type'] || 'Yarn');
      fileName = poTypeName.replace(/\s+/g, '_') + '_PO_' + String(id).replace(/[^\w-]+/g, '_');
    } else if (type === 'DC') {
      var dc = sheetToObjects(SHEETS.PROCESS_DC).find(function (r) { return r['DC No'] === id; });
      if (!dc) throw new Error('DC "' + id + '" not found.');
      fileName = 'Process_DC_' + id;
      htmlContent = generateDcHtml(dc);
    } else if (type === 'JODETAILS') {
      var details = getJODetails(id);
      if (!details.jo) throw new Error('JO "' + id + '" not found in Style Tracker.');
      fileName = 'JO_Details_' + id;
      htmlContent = generateJoDetailsHtml(details, id);
    } else if (type === 'TRANSFER') {
      var t = sheetToObjects(SHEETS.JO_TRANSFER).find(function (r) { return r['Transfer No'] === id; });
      if (!t) throw new Error('Transfer "' + id + '" not found.');
      fileName = 'JO_Transfer_' + id;
      htmlContent = generateTransferHtml(t);
    } else if (type === 'DEBITNOTE') {
      var dn = sheetToObjects(SHEETS.DEBIT_NOTE).find(function (r) { return r['Debit Note No'] === id; });
      if (!dn) throw new Error('Debit Note "' + id + '" not found.');
      fileName = 'Debit_Note_' + id;
      htmlContent = generateDebitNoteHtml(dn);
    } else if (type === 'RECEIPT') {
      var inw = sheetToObjects(SHEETS.YARN_INWARD).find(function (r) { return r['Inward No'] === id; });
      if (!inw) throw new Error('Yarn Receipt "' + id + '" not found.');
      fileName = 'Yarn_Receipt_' + id;
      htmlContent = generateYarnReceiptHtml(inw);
    } else if (type === 'ORDERSHEET') {
      var osData = getOrderSheetData(id);
      if (!osData) throw new Error('JO "' + id + '" not found in Style Tracker.');
      fileName = 'Order_Sheet_' + id;
      htmlContent = generateOrderSheetHtml(osData);
    } else if (type === 'FABRICPROGRAM') {
      var fpData = getFabricProgramData(id);
      if (!fpData) throw new Error('JO "' + id + '" not found in Style Tracker.');
      fileName = 'Fabric_Program_' + id;
      htmlContent = generateFabricProgramHtml(fpData);
    } else if (type === 'PROCESSREQ') {
      fileName = 'Fabric_Program_' + id;
      htmlContent = generateProcessReqHtml(id);
    } else if (type === 'OIH') {
      var oihF = {}; try { oihF = JSON.parse(id || '{}'); } catch (pe) { }
      fileName = 'Order_In_Hand' + (oihF.buyer ? '_' + String(oihF.buyer).replace(/[^\w-]+/g, '_') : '_All_Buyers');
      htmlContent = generateOihPdfHtml_(oihF);
      id = 'OIH';
    } else if (type === 'ORDERFABRIC') {
      fileName = 'Order_and_Fabric_' + id;
      htmlContent = generateOrderFabricHtml(id);
    } else if (type === 'ACCESSORYPO') {
      var apo = sheetToObjects(SHEETS.ACCESSORIES_PO).find(function (r) { return r['PO No'] === id; });
      if (!apo) throw new Error('Accessories PO "' + id + '" not found.');
      if (apo['Status'] !== 'Approved') throw new Error('Accessories PO ' + id + ' is "' + apo['Status'] + '" — only an Approved PO can be printed.');
      var grpNos = apoGroupRows_(id).map(function (g) { return g.no; });
      var lines = grpNos.length > 1 ? sheetToObjects(SHEETS.ACCESSORIES_PO).filter(function (r) { return grpNos.indexOf(r['PO No']) > -1; }) : null;
      if (lines) apo = lines[0];
      fileName = 'Accessories_PO_' + apo['PO No'];
      htmlContent = lines ? generateAccessoryPoMultiHtml_(lines) : generateAccessoryPoHtml(apo);
    } else if (type === 'ACCESSORYRECEIPT') {
      var arec = sheetToObjects(SHEETS.ACCESSORIES_RECEIPT).find(function (r) { return r['Receipt No'] === id; });
      if (!arec) throw new Error('Accessories Receipt "' + id + '" not found.');
      fileName = 'Accessories_Receipt_' + id;
      htmlContent = generateAccessoryReceiptHtml(arec);
    } else if (type === 'GENERALBILL') {
      var gbill = sheetToObjects(SHEETS.GENERAL_BILL).find(function (r) { return r['Bill No'] === id; });
      if (!gbill) throw new Error('Bill "' + id + '" not found.');
      if (gbill['Status'] !== 'Approved') throw new Error('Bill ' + id + ' is "' + gbill['Status'] + '" — only an Approved bill can be printed.');
      fileName = 'General_Bill_' + id;
      htmlContent = generateGeneralBillHtml_(gbill);
    } else if (type === 'YARNBILL') {
      var ybill = sheetToObjects(SHEETS.YARN_BILL).find(function (r) { return r['Bill No'] === id; });
      if (!ybill) throw new Error('Yarn Bill "' + id + '" not found.');
      if (ybill['Status'] !== 'Approved') throw new Error('Yarn Bill ' + id + ' is "' + ybill['Status'] + '" — only an Approved bill can be printed.');
      fileName = 'Yarn_Bill_' + id;
      htmlContent = generateYarnBillHtml(ybill);
    } else if (type === 'CMTBILL') {
      var grpNos = cmtGroupRows_(id).map(function (r) { return r['Bill No']; });
      var clines = sheetToObjects(SHEETS.CMT_BILL).filter(function (r) { return grpNos.indexOf(r['Bill No']) > -1; });
      if (!clines.length) throw new Error('CMT Bill "' + id + '" not found.');
      if (clines[0]['Status'] !== 'Approved') throw new Error('CMT Bill ' + id + ' is "' + clines[0]['Status'] + '" — only an Approved bill can be printed.');
      fileName = 'CMT_Bill_' + id;
      htmlContent = generateCmtBillHtml_(clines);
    } else if (type === 'ACCESSORYBILL') {
      var abill = sheetToObjects(SHEETS.ACCESSORIES_BILL).find(function (r) { return r['Bill No'] === id; });
      if (!abill) throw new Error('Accessories Bill "' + id + '" not found.');
      if (abill['Status'] !== 'Approved') throw new Error('Accessories Bill ' + id + ' is "' + abill['Status'] + '" — only an Approved bill can be printed.');
      fileName = 'Accessories_Bill_' + id;
      htmlContent = generateAccessoryBillHtml(abill);
    } else {
      throw new Error('Invalid document type.');
    }

    var folder = getArchiveFolder();
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    var fullFileName = fileName + '_' + timestamp + '.pdf';
    var blob = Utilities.newBlob(htmlContent, 'text/html', fullFileName).getAs('application/pdf');
    blob.setName(fullFileName);
    var savedFile = folder.createFile(blob);

    return {
      success: true, fileUrl: savedFile.getUrl(),
      downloadUrl: 'https://drive.google.com/uc?export=download&id=' + savedFile.getId(),
      fileName: fullFileName, base64: Utilities.base64Encode(blob.getBytes()),
      note: (frozen && frozen.missing) ? 'The saved copy of this sheet could not be found in Drive — this PDF was generated fresh.' : ''
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function pdfShell(title, accent, bodyHtml) {
  return '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Arial,sans-serif;margin:25px;color:#2c3e50;}' +
    '.header{text-align:center;border-bottom:2px solid ' + accent + ';padding-bottom:10px;margin-bottom:20px;}' +
    '.header h1{margin:0;color:' + accent + ';font-size:24px;text-transform:uppercase;}' +
    '.title-bar{background:#f4f6f8;padding:10px 14px;font-weight:bold;border-left:5px solid ' + accent + ';margin-bottom:20px;}' +
    'table{width:100%;border-collapse:collapse;margin-top:10px;}' +
    'th,td{border:1px solid #bdc3c7;padding:8px 10px;font-size:12px;text-align:left;}' +
    'th{background:#f2f4f4;color:#34495e;}' +
    '.footer{margin-top:30px;font-size:11px;color:#7f8c8d;border-top:1px dashed #bdc3c7;padding-top:10px;}' +
    '</style></head><body><div class="header"><h1>Jiwin Apparels</h1><p>' + title + '</p></div>' +
    bodyHtml + '<div class="footer">Generated on ' + new Date().toLocaleString() + ' | Jiwin Apparels ERP</div></body></html>';
}

/** Converts a rupee amount to words, Indian numbering (Lakh/Crore), for the Yarn PO print format. */
function numberToWordsIndian(num) {
  var ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten',
    'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  var tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
  function twoDigits(n) {
    if (n < 20) return ones[n];
    return tens[Math.floor(n / 10)] + (n % 10 ? ' ' + ones[n % 10] : '');
  }
  function threeDigits(n) {
    if (n < 100) return twoDigits(n);
    return ones[Math.floor(n / 100)] + ' Hundred' + (n % 100 ? ' and ' + twoDigits(n % 100) : '');
  }
  num = Math.round(num);
  if (num === 0) return 'Zero';
  var crore = Math.floor(num / 10000000); num %= 10000000;
  var lakh = Math.floor(num / 100000); num %= 100000;
  var thousand = Math.floor(num / 1000); num %= 1000;
  var rest = num;
  var parts = [];
  if (crore) parts.push(threeDigits(crore) + ' Crore');
  if (lakh) parts.push(threeDigits(lakh) + ' Lakh');
  if (thousand) parts.push(threeDigits(thousand) + ' Thousand');
  if (rest) parts.push(threeDigits(rest));
  return parts.join(' ');
}


function generatePoMultiHtml_(lines) {
  var po=lines[0], escH=function(v){return String(v==null?'':v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');};
  var qty=lines.reduce(function(s,r){return s+(Number(r['Qty KG'])||0);},0);
  var amount=lines.reduce(function(s,r){return s+(Number(r['PO Value'])||0);},0);
  var gst=lines.reduce(function(s,r){return s+(Number(r['PO Value'])||0)*(Number(r['GST %'])||0)/100;},0);
  var gstPct=Number(po['GST %'])||0, net=Math.round(amount+gst), round=net-(amount+gst);
  // The Yarn PO sheet itself never stored Fabric Type or Finishing Dia — both are looked up fresh
  // from Style Tracker here (matched by JO + Counts + Content + GSM, the same matching used
  // everywhere else in the app), so this works for POs raised before this fix too, not just new ones.
  var allSt = getStyleTracker();
  // Matching has to tolerate small formatting differences between what's stored on the Yarn PO line
  // and what's in Style Tracker — "200" vs 200 vs "200.0", "30s" vs "30'S" — or two fabrics that
  // share a count (body vs rib) but differ only by GSM would never actually match.
  var normCounts_ = function (v) { try { return normalizeCounts_(v); } catch (e) { return String(v || '').trim().toUpperCase(); } };
  var specFor_ = function(r){
    var joRows = allSt.filter(function(s){ return String(s['JO Number']).trim() === String(r['JO Number']).trim(); });
    var wantCounts = normCounts_(r['Counts']), wantContent = String(r['Content'] || '').trim().toLowerCase(), wantGsm = Number(r['GSM']) || 0;
    // Pass 1: Counts + Content + GSM (GSM is what tells body vs rib apart when both share a count)
    var hit = joRows.find(function(s){
      return normCounts_(s['Counts']) === wantCounts && String(s['Content'] || '').trim().toLowerCase() === wantContent &&
        (!wantGsm || Number(s['GSM']) === wantGsm);
    });
    if (hit) return hit;
    // Pass 2: Counts + GSM only (content wording sometimes differs slightly between the two sheets)
    if (wantGsm) {
      hit = joRows.find(function(s){ return normCounts_(s['Counts']) === wantCounts && Number(s['GSM']) === wantGsm; });
      if (hit) return hit;
    }
    // Pass 3: Counts + Content, no GSM to narrow by
    hit = joRows.find(function(s){ return normCounts_(s['Counts']) === wantCounts && String(s['Content'] || '').trim().toLowerCase() === wantContent; });
    return hit || joRows[0] || {};
  };
  var rows=lines.map(function(r,i){
    var spec = specFor_(r);
    // Fabric Type is only unambiguous when GSM narrowed the match (Cora Fabric / Dyed Fabric lines
    // carry GSM) — a plain Yarn PO line usually has no GSM, and the same yarn can feed both body
    // and rib, so guessing one fabric type for it would be misleading.
    var desc=(r['GSM'] ? [spec['Fabric type'], r['Counts'], r['Content']] : [r['Counts'], r['Content']]).filter(Boolean).join(' / ');
    var finDia = r['GSM'] ? (spec['Finishing dia'] || r['Knitting Dia'] || '') : (r['Knitting Dia'] || '');
    return '<tr><td>'+ (i+1)+'</td><td>'+escH(r['JO Number'])+'</td><td>'+escH(r['Style Number']||'-')+'</td><td>'+escH(desc)+'</td><td>'+escH(finDia||'-')+'</td><td>'+escH(r['GSM']||'-')+'</td><td>'+escH(r['Yarn Color']||'-')+'</td><td class="num">'+(Number(r['Qty KG'])||0).toFixed(2)+'</td><td class="num">'+(Number(r['Rate (Rs/kg)'])||0).toFixed(2)+'</td><td class="num">'+(Number(r['PO Value'])||0).toFixed(2)+'</td></tr>';
  }).join('');
  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>body{font-family:Arial,sans-serif;font-size:11px;color:#000;margin:20px}table{width:100%;border-collapse:collapse}td,th{border:1px solid #000;padding:6px;vertical-align:top;text-align:left}.num{text-align:right}.title{text-align:center;font-weight:bold;font-size:16px;margin:12px}.bold{font-weight:bold}</style></head><body>'+
    '<table><tr><td style="width:60%"><b style="font-size:15px">JIWIN APPARELS</b><br>10/227, SF NO 4/470, GANAPATHYPALAYAM ROAD,<br>KUNNAKKAPALAYAM, VEERAPANDI PO,<br>TIRUPUR, TAMILNADU, CODE: 33<br><b>GSTIN:</b> 33BOWPS1618P2Z2 | Phone: 9092998674</td><td><b>PO No:</b> '+escH(po['PO No'])+'<br><b>Date:</b> '+escH(po['PO Date'])+'<br><b>JO(s):</b> '+escH(Array.from(new Set(lines.map(function(r){return r['JO Number'];}))).join(', '))+'<br><b>Delivery Date:</b> '+escH(po['Delivery Date']||'-')+'<br><b>Payment Terms:</b> '+escH(po['Payment Terms']||'-')+'</td></tr></table>'+
    '<div class="title">'+escH(po['PO Type']||'YARN').toUpperCase()+' PURCHASE ORDER</div>'+
    '<table><tr><td style="width:50%"><b>To Vendor:</b><br>'+escH(po['Supplier']||'-')+'<br>'+escH(po['Supplier Address']||'-')+'<br><b>GSTIN:</b> '+escH(po['GST No']||'-')+'</td><td><b>Knitting Unit:</b> '+escH(po['Knitting Unit Name']||'-')+'<br><b>Delivery At:</b> '+escH(po['Delivery Address']||'-')+'<br><b>Knitting Unit GSTIN:</b> '+escH(po['Knitting Unit GSTIN']||'-')+'<br><b>Phone:</b> '+escH(po['Knitting Unit Phone']||'-')+'<br><b>Buyer(s):</b> '+escH(Array.from(new Set(lines.map(function(r){return r['Buyer'];}))).filter(Boolean).join(', '))+'</td></tr></table>'+
    '<p>We hereby request you to supply the following items as per the terms &amp; conditions mentioned below.</p>'+
    '<table><tr><th>SI</th><th>JO</th><th>Style</th><th>Fabric / Yarn Description / Count</th><th>Finishing Dia</th><th>GSM</th><th>Yarn Color</th><th>Qty (Kg)</th><th>Rate/Kg</th><th>Amount (Rs)</th></tr>'+rows+
    '<tr><td colspan="7" class="bold">TOTAL</td><td class="num bold">'+qty.toFixed(2)+'</td><td></td><td class="num bold">'+amount.toFixed(2)+'</td></tr></table>'+
    '<table style="margin-top:0"><tr><td rowspan="4" style="width:60%"><b>Rupees in Words:</b><br>Rupees '+numberToWordsIndian(net)+' Only<br><br><b>Terms &amp; Conditions:</b><ol><li>Please quote our PO number and JO number(s) in delivery and invoices.</li><li>Quality shall be as per approved sample/specification.</li><li>Quantity shall be as per this PO.</li></ol></td><td>Subtotal</td><td class="num">₹ '+amount.toFixed(2)+'</td></tr><tr><td>GST</td><td class="num">₹ '+gst.toFixed(2)+'</td></tr><tr><td>Rounding</td><td class="num">₹ '+round.toFixed(2)+'</td></tr><tr><td class="bold">Net Amount</td><td class="num bold">₹ '+net.toFixed(2)+'</td></tr></table>'+
    '<table style="margin-top:12px"><tr style="height:60px"><td style="width:33%;vertical-align:bottom"><b>Prepared By:</b><br>'+escH(po['Prepared By']||'-')+'</td><td style="width:33%;vertical-align:bottom;text-align:center"><b>Authorized By:</b><br>'+escH(po['Approved By']||'(Pending)')+poSealHtml_(po['Approved By'])+'</td><td style="width:34%;vertical-align:bottom;text-align:right"><b>For JIWIN APPARELS</b><br><br>Status: '+escH(po['Status'])+'</td></tr></table></body></html>';
}

function generatePoHtml(po) {
  var qty = Number(po['Qty KG']) || 0;
  var rate = Number(po['Rate (Rs/kg)']) || 0;
  var gstPct = Number(po['GST %']) || 0;
  var amount = qty * rate;
  var gstAmount = amount * gstPct / 100;
  var rawNet = amount + gstAmount;
  var netAmount = Math.round(rawNet);
  var rounding = netAmount - rawNet;
  var isCoraFabric = String(po['PO Type'] || '').trim() === 'Cora Fabric';
  // Fabric Type / Finishing Dia aren't stored on the Yarn PO sheet itself — looked up fresh from
  // Style Tracker (matched by JO + Counts + Content + GSM), so this works for POs raised before
  // this fix too, not just new ones.
  var specRowForPo_ = (function () {
    var joRows = getStyleTracker().filter(function (s) { return String(s['JO Number']).trim() === String(po['JO Number']).trim(); });
    return joRows.find(function (s) {
      return String(s['Counts'] || '').trim() === String(po['Counts'] || '').trim() &&
        String(s['Content'] || '').trim() === String(po['Content'] || '').trim() &&
        (!po['GSM'] || String(s['GSM'] || '').trim() === String(po['GSM'] || '').trim());
    }) || joRows[0] || {};
  })();
  var desc = [specRowForPo_['Fabric type'], po['Counts'], po['Content']].filter(Boolean).join(' / ');
  var knittingDia = specRowForPo_['Finishing dia'] || po['Knitting Dia'] || '';

  // "Required Qty" — what Style Tracker says this JO/Counts(/Color/Content/GSM)
  // actually needs, for reference alongside what's actually being ordered here.
  var requiredKg = 0;
  var isColorSpecific = po['PO Type'] === 'Dyed Fabric' && po['Color'];
  var poUseGsm = String(po['PO Type'] || 'Yarn').trim() !== 'Yarn';
  var counted = null;
  try {
    counted = getCountsForJO(po['JO Number'], isColorSpecific ? po['Color'] : '', poUseGsm).find(function (c) {
      return String(c.counts) === String(po['Counts'] || '').trim() && String(c.content) === String(po['Content'] || '').trim() &&
        String(c.gsm || '') === (poUseGsm ? String(po['GSM'] || '').trim() : '') && String(c.yarnColor || '') === String(po['Yarn Color'] || '').trim();
    });
  } catch (e) { counted = null; }
  if (counted) requiredKg = counted.requiredKg;

  return '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Arial,sans-serif;font-size:12px;color:#000;margin:20px;}' +
    'table{width:100%;border-collapse:collapse;}' +
    'td,th{border:1px solid #000;padding:6px 8px;vertical-align:top;text-align:left;}' +
    '.po-title{text-align:center;font-weight:bold;font-size:16px;margin:10px 0;letter-spacing:0.5px;}' +
    '.num{text-align:right;}.bold{font-weight:bold;}' +
    '</style></head><body>' +
    '<table><tr>' +
    '<td style="width:60%;"><strong style="font-size:15px;">JIWIN APPARELS</strong><br>' +
    '10/227, SF NO 4/470, GANAPATHYPALAYAM ROAD,<br>KUNNAKKAPALAYAM, VEERAPANDI PO,<br>TIRUPUR, TAMILNADU, CODE: 33<br>' +
    '<strong>GSTIN NO:</strong> 33BOWPS1618P2Z2 | <strong>Phone No:</strong> 9092998674</td>' +
    '<td style="width:40%;"><strong>PO No:</strong> ' + po['PO No'] + '<br>' +
    '<strong>Date:</strong> ' + po['PO Date'] + '<br>' +
    '<strong>Order No / JO:</strong> ' + (po['JO Number'] || '-') + '<br>' +
    '<strong>Delivery Date:</strong> ' + (po['Delivery Date'] || '-') + '<br>' +
    '<strong>Payment Terms:</strong> ' + (po['Payment Terms'] || '-') + '</td>' +
    '</tr></table>' +
    '<div class="po-title">' + (po['PO Type'] || 'YARN').toUpperCase() + ' PURCHASE ORDER</div>' +
    '<table style="margin-top:-1px;"><tr>' +
    '<td style="width:50%;"><strong>To Vendor:</strong><br><span class="bold">' + (po['Supplier'] || '-') + '</span><br>' +
    (po['Supplier Address'] || '-') + '<br><strong>GSTIN NO:</strong> ' + (po['GST No'] || '-') + '</td>' +
    '<td style="width:50%;"><strong>Knitting Unit:</strong> <span class="bold">' + (po['Knitting Unit Name'] || '-') + '</span><br>' +
    '<strong>Delivery At (Knitting Unit Address):</strong><br><span class="bold">' + (po['Delivery Address'] || '-') + '</span><br>' +
    '<strong>Knitting Unit GSTIN:</strong> ' + (po['Knitting Unit GSTIN'] || '-') + '<br>' +
    '<strong>Knitting Contact Phone No:</strong> ' + (po['Knitting Unit Phone'] || '-') + '<br>' +
    '<strong>JO / Buyer:</strong> ' + (po['JO Number'] || '-') + ' / ' + (po['Buyer'] || '-') + '<br>' +
    '<strong>Style:</strong> ' + (po['Style Number'] || '-') + (po['Color'] ? ' / ' + po['Color'] : '') + '</td>' +
    '</tr><tr><td colspan="2"><strong>Required Qty (per Style Tracker):</strong> ' + requiredKg.toFixed(2) + ' kg — this PO covers ' + qty.toFixed(2) + ' kg of it.</td></tr>' +
    '<tr><td colspan="2">We hereby request you to supply the following items as per the terms &amp; conditions mentioned below.</td></tr></table>' +
    '<table style="margin-top:-1px;"><tr style="background:#F1F1F1;">' +
    (isCoraFabric
      ? '<th style="width:30px;">SI</th><th>Fabric / Yarn Description / Count</th><th>Finishing Dia</th><th>GSM</th><th>Yarn Color</th><th class="num">Bags</th><th class="num">Qty Kgs</th><th class="num">Rate/Kg</th><th class="num">Amount</th><th class="num">GST %</th></tr>' +
        '<tr><td>1</td><td>' + desc + '</td><td>' + (knittingDia || '-') + '</td><td>' + (po['GSM'] || '-') + '</td><td>' + (po['Yarn Color'] || '-') + '</td><td class="num">' + (po['No of Bags'] || 0) + '</td>'
      : '<th style="width:30px;">SI</th><th>Yarn Description</th><th>Yarn Color</th><th class="num">Bags</th><th class="num">Qty Kgs</th><th class="num">Rate/Kg</th><th class="num">Amount</th><th class="num">GST %</th></tr>' +
        '<tr><td>1</td><td>' + [po['Counts'], po['Content'], po['GSM'] ? po['GSM'] + ' GSM' : ''].filter(Boolean).join(' / ') + '</td><td>' + (po['Yarn Color'] || '-') + '</td><td class="num">' + (po['No of Bags'] || 0) + '</td>') +
    '<td class="num">' + qty.toFixed(2) + '</td><td class="num">' + rate.toFixed(2) + '</td><td class="num">' + amount.toFixed(2) + '</td><td class="num">' + gstPct.toFixed(2) + '</td></tr>' +
    '</table>' +
    '<table style="margin-top:-1px;"><tr>' +
    '<td colspan="1" rowspan="4" style="width:55%;"><strong>Rupees in Words:</strong><br><span class="bold">Rupees ' + numberToWordsIndian(netAmount) + ' Only</span>' +
    '<br><br><strong>Terms &amp; Conditions:</strong><ol style="margin:5px 0 0 15px;padding:0;">' +
    '<li>Please Quote Our P.O. No and Order No: in all your delivery and Invoices.</li>' +
    '<li>Quality of Items should be exactly as per sample.</li>' +
    '<li>Quantity Should be as per P.O.</li></ol></td>' +
    '<td class="bold" style="width:25%;">Total Qty / Amt</td><td class="num bold" style="width:20%;">' + qty.toFixed(2) + ' KGS<br>₹ ' + amount.toFixed(2) + '</td></tr>' +
    '<tr><td>GST Total</td><td class="num">₹ ' + gstAmount.toFixed(2) + '</td></tr>' +
    '<tr><td>Rounding</td><td class="num">₹ ' + rounding.toFixed(2) + '</td></tr>' +
    '<tr><td class="bold" style="font-size:13px;">Net Amount</td><td class="num bold" style="font-size:13px;">₹ ' + netAmount.toFixed(2) + '</td></tr>' +
    '</table>' +
    '<table style="margin-top:-1px;"><tr style="height:60px;">' +
    '<td style="width:33%;vertical-align:bottom;"><strong>Prepared By:</strong><br>' + (po['Prepared By'] || '-') + '</td>' +
    '<td style="width:33%;vertical-align:bottom;text-align:center;"><strong>Authorized By:</strong><br>' + (po['Approved By'] || '(Pending)') + poSealHtml_(po['Approved By']) + '</td>' +
    '<td style="width:34%;vertical-align:bottom;text-align:right;"><strong>For JIWIN APPARELS</strong><br><br>Status: ' + po['Status'] + '</td>' +
    '</tr></table>' +
    '</body></html>';
}

function generateDcHtmlClassic_(dc) {
  // Vendor block: To Supplier's own Address/GST from Supplier Master.
  var supplierRow = getSuppliers().find(function (s) { return String(s['Supplier Name'] || '').trim() === String(dc['To Supplier'] || '').trim(); });
  var supplierAddress = supplierRow ? (supplierRow['Address'] || '') : '';
  var supplierGst = supplierRow ? (supplierRow['GST'] || '') : '';

  // Fabric spec + Order No come from Style Tracker (Knitting Dia / Loop
  // Length / Fabric type aren't stored on the DC itself, only the fabric
  // spec — matched against JO + Counts + Content + GSM, same as everywhere
  // else in the system).
  var styleRows = getStyleTracker().filter(function (r) { return String(r['JO Number']).trim() === String(dc['JO Number']).trim(); });
  var specRow = styleRows.find(function (r) {
    return String(r['Counts'] || '').trim() === String(dc['Counts'] || '').trim() &&
      String(r['Content'] || '').trim() === String(dc['Content'] || '').trim() &&
      (!dc['GSM'] || String(r['GSM'] || '').trim() === String(dc['GSM'] || '').trim());
  }) || styleRows[0] || {};
  var buyerOrderNo = specRow['Buyer Order No'] || '';
  var fabricType = specRow['Fabric type'] || '';
  var knittingDia = specRow['Knitting Dia'] || '';
  var loopLength = specRow['Loop length'] || '';
  var countsContent = [dc['Counts'], dc['Content']].filter(Boolean).join(' ');

  return '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Arial,sans-serif;font-size:12px;color:#000;margin:20px;}' +
    'table{width:100%;border-collapse:collapse;}' +
    'td,th{border:1px solid #000;padding:6px 8px;vertical-align:top;text-align:left;}' +
    '.dc-title{text-align:center;font-weight:bold;font-size:16px;margin:10px 0;letter-spacing:0.5px;}' +
    '.num{text-align:right;}.bold{font-weight:bold;}' +
    '</style></head><body>' +
    '<table><tr>' +
    '<td style="width:60%;"><strong style="font-size:15px;">JIWIN APPARELS</strong><br>' +
    '10/227, SF NO 4/470, GANAPATHYPALAYAM ROAD,<br>KUNNAKKAPALAYAM, VEERAPANDI PO,<br>' +
    '<strong>Phone No:</strong> 9092998674<br><strong>GSTIN NO:</strong> 33BOWPS1618P2Z2<br>' +
    '<strong>STATE NAME:</strong> TAMILNADU &nbsp; <strong>CODE:</strong> 33</td>' +
    '<td style="width:40%;"><strong>D.C.NO:</strong> ' + dc['DC No'] + (dc['Supplier DC No'] ? ' (Supplier DC: ' + dc['Supplier DC No'] + ')' : '') + '<br>' +
    '<strong>Date:</strong> ' + dc['Date'] + '<br>' +
    '<strong>Order No:</strong> ' + (buyerOrderNo || dc['JO Number'] || '-') + '</td>' +
    '</tr></table>' +
    '<div class="dc-title">FABRIC DELIVERY - ' + (dc['To Process'] || '').toUpperCase() + '</div>' +
    '<table style="margin-top:-1px;"><tr><td>' +
    '<strong>To</strong><br><span class="bold">' + (dc['To Supplier'] || '-') + '</span><br>' +
    (supplierAddress || '') + '<br><strong>GSTIN NO:</strong> ' + (supplierGst || '-') +
    '</td></tr></table>' +
    '<table style="margin-top:-1px;"><thead><tr style="background:#F1F1F1;">' +
    '<th style="width:30px;">Sl</th><th>Fab.Type</th><th>Colour</th><th>K.Dia</th><th>GSM</th><th>LL</th><th>Lot No</th><th class="num">Rolls</th><th class="num">Kgs</th></tr></thead>' +
    '<tbody><tr><td>1</td><td>' + (fabricType || '-') + '<br><span style="font-size:10px;">' + (countsContent || '') + '</span></td>' +
    '<td>' + (dc['Color'] || '-') + '</td><td>' + (knittingDia || '-') + '</td><td>' + (dc['GSM'] || '-') + '</td><td>' + (loopLength || '-') + '</td>' +
    '<td>' + (dc['Lot No'] || '-') + '</td><td class="num">' + (dc['Rolls'] || 0) + '</td><td class="num">' + Number(dc['DC Issue Qty KG'] || 0).toFixed(2) + '</td></tr>' +
    '<tr class="bold"><td colspan="7" style="text-align:right;">Total</td><td class="num">' + (dc['Rolls'] || 0) + '</td><td class="num">' + Number(dc['DC Issue Qty KG'] || 0).toFixed(2) + '</td></tr>' +
    '</tbody></table>' +
    '<table style="margin-top:-1px;"><tr>' +
    '<td style="width:50%;"><strong>Remarks:</strong> ' + (dc['Remarks'] || '-') + '</td>' +
    '<td style="width:50%;"><strong>Vech.No:</strong> ' + (dc['Vehicle No'] || '-') + '</td>' +
    '</tr></table>' +
    '<table style="margin-top:-1px;"><tr style="height:55px;">' +
    '<td style="width:33%;vertical-align:bottom;">Receiver Sign:</td>' +
    '<td style="width:34%;vertical-align:bottom;text-align:center;"><strong>For JIWIN APPARELS</strong></td>' +
    '<td style="width:33%;vertical-align:bottom;text-align:right;">Status: ' + dc['Status'] + '</td>' +
    '</tr><tr>' +
    '<td>Prepared By:</td><td style="text-align:center;">FABRIC SECURITY RECEIVED</td>' +
    '<td style="text-align:right;">Authorised By: ' + (dc['Approved By'] || '-') + '</td>' +
    '</tr></table>' +
    '</body></html>';
}

function generateJoDetailsHtml(details, joNumber) {
  var jo = details.jo || {};
  var rows = details.timeline.map(function (t) {
    return '<tr><td>' + t.date + '</td><td>' + t.stage + '</td><td>' + t.supplier + '</td><td>' + t.status + '</td><td>' + t.reference + '</td></tr>';
  }).join('');
  return pdfShell('JO Details — ' + joNumber, '#8e44ad',
    '<table><tr><th>Style</th><td>' + (jo['Style Number'] || '-') + '</td><th>Buyer</th><td>' + (jo['Buyer'] || '-') + '</td></tr>' +
    '<tr><th>Color</th><td>' + (jo['Color'] || '-') + '</td><th>Fabric Required</th><td>' + (jo.totalFabricRequiredKg || 0) + ' KG</td></tr>' +
    '<tr><th>Current Process</th><td>' + details.currentProcess + '</td><th>Current Supplier</th><td>' + details.currentSupplier + '</td></tr>' +
    '<tr><th>Last Updated</th><td colspan="3">' + details.lastUpdated + '</td></tr></table>' +
    '<h3>Timeline</h3><table><thead><tr><th>Date</th><th>Stage</th><th>Supplier</th><th>Status</th><th>Reference</th></tr></thead><tbody>' + rows + '</tbody></table>');
}

function generateTransferHtml(t) {
  return pdfShell('JO Transfer', '#d35400',
    '<div class="title-bar">Transfer #: <strong>' + t['Transfer No'] + '</strong> | Date: ' + t['Date'] + '</div>' +
    '<table><tr><th>Source JO</th><td>' + t['Source JO'] + '</td><th>Destination JO</th><td>' + t['Destination JO'] + '</td></tr>' +
    '<tr><th>Stock Type</th><td>' + (t['Stock Type'] || 'Yarn') + '</td><th>Counts</th><td>' + (t['Counts'] || '-') + '</td></tr>' +
    '<tr><th>Qty</th><td colspan="3">' + Number(t['Qty KG']).toFixed(2) + ' KG</td></tr>' +
    '<tr><th>Status</th><td colspan="3">' + t['Status'] + '</td></tr>' +
    '<tr><th>Remarks</th><td colspan="3">' + (t['Remarks'] || '-') + '</td></tr></table>');
}

function generateDebitNoteHtml(dn) {
  return pdfShell('Debit Note', '#c0392b',
    '<div class="title-bar">Debit Note #: <strong>' + dn['Debit Note No'] + '</strong> | Date: ' + dn['Date'] + '</div>' +
    '<table><tr><th>JO Number</th><td>' + dn['JO Number'] + '</td><th>Supplier</th><td>' + dn['Supplier'] + '</td></tr>' +
    '<tr><th>DC No</th><td>' + dn['DC No'] + '</td><th>Bill No</th><td>' + dn['Bill No'] + '</td></tr>' +
    '<tr><th>Excess Loss</th><td>' + Number(dn['Excess Loss KG']).toFixed(2) + ' KG</td><th>Rate</th><td>₹' + Number(dn['Rate']).toFixed(2) + '</td></tr>' +
    '<tr><th>Amount</th><td colspan="3"><strong>₹' + Number(dn['Amount']).toFixed(2) + '</strong></td></tr></table>');
}

function generateYarnReceiptHtml(inw) {
  return pdfShell('Yarn Receipt Summary', '#1a5276',
    '<div class="title-bar">Inward #: <strong>' + inw['Inward No'] + '</strong> | Date: ' + inw['Date'] + '</div>' +
    '<table><tr><th>JO Number</th><td>' + (inw['JO Number'] || '-') + '</td><th>Yarn PO No</th><td>' + (inw['Yarn PO No'] || '-') + '</td></tr>' +
    '<tr><th>Mill Name</th><td>' + (inw['Mill Name'] || '-') + '</td><th>Mill DC / Invoice No</th><td>' + (inw['Mill DC/Invoice No'] || '-') + '</td></tr>' +
    '<tr><th>Counts</th><td>' + (inw['Counts'] || '-') + '</td><th>Yarn Type</th><td>' + (inw['Yarn Type'] || '-') + '</td></tr>' +
    '<tr><th>Content</th><td>' + (inw['Content'] || '-') + '</td><th>GSM</th><td>' + (inw['GSM'] || '-') + '</td></tr>' +
    '<tr><th>Received Bags</th><td>' + (inw['Received Bags'] || 0) + '</td><th>Received KG</th><td><strong>' + Number(inw['Received KG'] || 0).toFixed(2) + ' KG</strong></td></tr>' +
    (inw['Knitting Unit'] ? '<tr><th>Delivered To (Knitting Unit)</th><td colspan="3"><strong>' + inw['Knitting Unit'] + '</strong></td></tr>' : '') +
    '<tr><th>Remarks</th><td colspan="3">' + (inw['Remarks'] || '-') + '</td></tr></table>');
}

function generateAccessoryPoHtml(po) {
  return generateAccessoryPoDocumentHtml_([po]);
}

/* Accessories PO print layout based on the Stores Purchase Order sample supplied. */
function generateAccessoryPoDocumentHtml_(lines) {
  lines = lines || [];
  if (!lines.length) throw new Error('No Accessories PO lines found.');
  var f = lines[0], gstPct = Number(f['GST %']) || 0, amount = 0, totalQty = 0;
  function h(v) { return String(v === null || v === undefined || v === '' ? '-' : v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;'); }
  function money(v) { return (Number(v) || 0).toFixed(2); }
  var rowsHtml = lines.map(function (l, i) {
    var q = Number(l['Qty Ordered']) || 0, rt = Number(l['Rate (Rs/unit)']) || 0, lineAmount = q * rt;
    totalQty += q; amount += lineAmount;
    var desc = [l['Trims / Accessories'] || 'Accessories', l['Color'] || '', l['Size'] ? 'Size ' + l['Size'] : '', l['Description'] || ''].filter(Boolean).join(' ');
    return '<tr><td class="num">' + (i + 1) + '</td><td class="desc">' + h(desc) + '</td><td class="center">NOS</td><td class="num">' + q.toLocaleString('en-IN') + '</td><td class="num">' + money(rt) + '</td><td class="num">' + money(lineAmount) + '</td><td class="num">' + gstPct.toFixed(2) + '</td></tr>';
  }).join('');
  var gstAmount = amount * gstPct / 100, rawNet = amount + gstAmount, netAmount = Math.round(rawNet), rounding = netAmount - rawNet;
  var date = h(f['PO Date'] || '');
  var amountWords = numberToWordsIndian(netAmount) + ' Rupees Only';
  var remarks = f['Remarks'] || f['Trims / Accessories'] || 'Accessories as per approved sample and specifications.';
  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' +
    '@page{size:A4 portrait;margin:10mm}*{box-sizing:border-box}body{font-family:Arial,sans-serif;font-size:10px;color:#111;margin:0}' +
    '.doc{width:100%;border:1px solid #222}.title{text-align:center;font-size:15px;font-weight:bold;text-decoration:underline;padding:7px;border-bottom:1px solid #222}' +
    'table{width:100%;border-collapse:collapse;table-layout:fixed}td,th{border:1px solid #555;padding:5px 6px;vertical-align:top}th{font-weight:bold;text-align:center;background:#f7f7f7}' +
    '.company{font-weight:bold;font-size:12px}.small{font-size:9px;line-height:1.45}.bold{font-weight:bold}.center{text-align:center}.num{text-align:right;white-space:nowrap}.meta td{height:24px}.items td{height:27px}.desc{width:44%;white-space:normal}.totals td{height:20px}.words{padding:7px;border-bottom:1px solid #555;font-weight:bold}.section{padding:6px;border-bottom:1px solid #555}.sign{height:88px;vertical-align:bottom;font-weight:bold}.sign div{padding-top:55px;text-align:center}.no-border{border:0!important}.supplier{height:74px}.terms{height:180px;vertical-align:top}.muted{font-size:9px}' +
    '</style></head><body><div class="doc"><div class="title">STORES PURCHASE ORDER</div>' +
    '<table><tr><td class="supplier" colspan="4"><div class="company">JIWIN APPARELS</div><div class="small">10/227, SF NO 4/470,<br>GANAPATHYPALAYAM ROAD,<br>KUNNAKKAPALAYAM, VEERAPANDI PO,<br>9092998674<br>Phone No :</div></td><td colspan="3"><b>GSTIN NO :33BOWPS1618P2Z2</b><br><br><b>STATE NAME :TAMILNADU CODE :33</b></td></tr>' +
    '<tr><td colspan="4"><b>To</b><br><b>' + h(f['Supplier'] || '') + '</b><br><span class="small">GSTIN NO : ' + h(f['GST No'] || '') + '</span></td><td colspan="3"><b>PO No:</b>　<strong>' + h(f['PO No']) + '</strong><br><br><b>Date:</b>　<strong>' + date + '</strong></td></tr>' +
    '<tr><td colspan="7"><b>JO / Order No:</b> ' + h(f['JO Number'] || '') + '　　<b>Style:</b> ' + h(f['Style Number'] || '') + '</td></tr>' +
    '<tr><td colspan="7" class="center bold">We hereby request you to supply the following items as per the terms &amp; conditions mentioned below.</td></tr>' +
    '<tr><td colspan="3"><b>PO No:</b> ' + h(f['PO No']) + '</td><td colspan="2"><b>Date:</b> ' + date + '</td><td colspan="2"><b>Page:</b> 1 of 1</td></tr>' +
    '<tr><th style="width:5%">Sl</th><th class="desc">Description</th><th style="width:8%">Unit</th><th style="width:12%">Qty</th><th style="width:12%">Rate</th><th style="width:14%">Amount</th><th style="width:9%">GST %</th></tr>' + rowsHtml +
    '<tr class="totals"><td colspan="3" class="bold">Total</td><td class="num bold">' + totalQty.toLocaleString('en-IN') + '</td><td></td><td class="num bold">' + money(amount) + '</td><td></td></tr>' +
    '<tr><td colspan="4" rowspan="4" class="small">' + h(remarks) + '</td><td colspan="2">Discount</td><td class="num">0.00</td></tr>' +
    '<tr><td colspan="2">GST Amount</td><td class="num">' + money(gstAmount) + '</td></tr><tr><td colspan="2">Rounding</td><td class="num">' + money(rounding) + '</td></tr><tr><td colspan="2"><b>Net Amount</b></td><td class="num"><b>' + money(netAmount) + '</b></td></tr>' +
    '<tr><td colspan="7" class="words">' + h(amountWords) + '</td></tr>' +
    '<tr><td colspan="7"><b>Remarks:</b> ' + h(remarks) + '<br><b>Delivery At:</b> ' + h(f['Delivery At'] || '') + '<br><b>Payment Terms:</b> ' + h(f['Payment Terms'] || '') + '<br><b>Delivery Date:</b> ' + h(f['Delivery Date'] || date) + '</td></tr>' +
    '<tr><td colspan="7" class="terms"><b>Terms &amp; Conditions:</b><br>Please quote our P.O. No and Order No. in all your delivery challans and invoices.<br>Quality of items should be exactly as per approved sample.<br>Quantity should be as per P.O.</td></tr>' +
    '<tr><td colspan="2" class="sign"><div>Prepared By</div></td><td colspan="2" class="sign"><div>ACC INCHARGE</div></td><td colspan="2" class="sign"><div>MERCH</div></td><td class="sign"><div>Authorised By</div></td></tr></table></div></body></html>';
}

function generateAccessoryReceiptHtml(r) {
  return pdfShell('Accessories Receipt Summary', '#1a5276',
    '<div class="title-bar">Receipt #: <strong>' + r['Receipt No'] + '</strong> | Date: ' + r['Date'] + '</div>' +
    '<table><tr><th>Accessories PO No</th><td>' + (r['Accessories PO No'] || '-') + '</td><th>JO Number</th><td>' + (r['JO Number'] || '-') + '</td></tr>' +
    '<tr><th>Trims / Accessories</th><td>' + (r['Trims / Accessories'] || '-') + '</td><th>Supplier</th><td>' + (r['Supplier'] || '-') + '</td></tr>' +
    '<tr><th>Ordered Qty</th><td>' + (r['Ordered Qty'] || 0) + '</td><th>Received Qty</th><td><strong>' + (r['Received Qty'] || 0) + '</strong></td></tr>' +
    '<tr><th>Invoice No</th><td>' + (r['Invoice No'] || '-') + '</td><th>Remarks</th><td>' + (r['Remarks'] || '-') + '</td></tr></table>');
}

function generateAccessoryBillHtml(b) {
  return pdfShell('Accessories Bill', '#1a5276',
    '<div class="title-bar">Bill #: <strong>' + b['Bill No'] + '</strong> | Date: ' + b['Date'] + ' | Status: ' + b['Status'] + '</div>' +
    '<table><tr><th>Supplier</th><td>' + (b['Supplier'] || '-') + '</td><th>JO Number</th><td>' + (b['JO Number'] || '-') + '</td></tr>' +
    '<tr><th>Accessories PO No</th><td>' + (b['Accessories PO No'] || '-') + '</td><th>Supplier Invoice No</th><td>' + (b['Supplier Invoice No'] || '-') + '</td></tr>' +
    '<tr><th>Trims / Accessories</th><td colspan="3">' + (b['Trims / Accessories'] || '-') + '</td></tr>' +
    '<tr><th>Ordered Qty</th><td>' + (b['Ordered Qty'] || 0) + '</td><th>Applicable Rate</th><td>₹' + Number(b['Applicable Rate'] || 0).toFixed(2) + '</td></tr>' +
    '<tr><th>Expected Amount</th><td>₹' + Number(b['Expected Amount'] || 0).toFixed(2) + '</td><th>Bill Amount</th><td><strong>₹' + Number(b['Bill Amount'] || 0).toFixed(2) + '</strong></td></tr>' +
    '<tr><th>GST %</th><td>' + Number(b['GST %'] || 0).toFixed(2) + '</td><th>GST Amount</th><td>₹' + Number(b['GST Amount'] || 0).toFixed(2) + '</td></tr>' +
    '<tr><th>Total incl. GST</th><td>₹' + Number(b['Total Incl GST'] || 0).toFixed(2) + '</td><th>Variance</th><td>₹' + Number(b['Variance'] || 0).toFixed(2) + '</td></tr>' +
    '<tr><th>Approved By</th><td>' + (b['Approved By'] || '-') + '</td><th>Approval Date</th><td>' + (b['Approval Date'] || '-') + '</td></tr>' +
    '<tr><th>Remarks</th><td colspan="3">' + (b['Remarks'] || '-') + '</td></tr></table>');
}

function generateYarnBillHtml(b) {
  return pdfShell('Yarn Bill', '#1a5276',
    '<div class="title-bar">Bill #: <strong>' + b['Bill No'] + '</strong> | Date: ' + b['Date'] + ' | Status: ' + b['Status'] + '</div>' +
    '<table><tr><th>Supplier</th><td>' + (b['Supplier'] || '-') + '</td><th>JO Number(s)</th><td>' + (b['JO Number'] || '-') + '</td></tr>' +
    '<tr><th>PO Numbers</th><td colspan="3">' + (b['PO Numbers'] || '-') + '</td></tr>' +
    '<tr><th>Supplier Invoice No</th><td>' + (b['Supplier Invoice No'] || '-') + '</td><th>PO Qty KG</th><td>' + Number(b['PO Qty KG'] || 0).toFixed(2) + '</td></tr>' +
    '<tr><th>Blended Rate</th><td>₹' + Number(b['Applicable Rate'] || 0).toFixed(2) + '</td><th>Expected Amount</th><td>₹' + Number(b['Expected Amount'] || 0).toFixed(2) + '</td></tr>' +
    '<tr><th>Bill Amount</th><td><strong>₹' + Number(b['Bill Amount'] || 0).toFixed(2) + '</strong></td><th>Variance</th><td>₹' + Number(b['Variance'] || 0).toFixed(2) + '</td></tr>' +
    '<tr><th>GST %</th><td>' + Number(b['GST %'] || 0).toFixed(2) + '</td><th>GST Amount</th><td>₹' + Number(b['GST Amount'] || 0).toFixed(2) + '</td></tr>' +
    '<tr><th>Total incl. GST</th><td>₹' + Number(b['Total Incl GST'] || 0).toFixed(2) + '</td><th>GST Check</th><td>' + (b['GST Match'] || '-') + '</td></tr>' +
    '<tr><th>Approved By</th><td>' + (b['Approved By'] || '-') + '</td><th>Approval Date</th><td>' + (b['Approval Date'] || '-') + '</td></tr>' +
    '<tr><th>Remarks</th><td colspan="3">' + (b['Remarks'] || '-') + '</td></tr></table>');
}
function generateCmtBillHtml_(lines) {
  var f = lines[0], total = 0;
  var rowsHtml = lines.map(function (l, i) {
    var q = Number(l['Qty']) || 0, rt = Number(l['Rate Per Pc']) || 0; total += q * rt;
    return '<tr><td>' + (i + 1) + '</td><td>' + l['JO Number'] + '</td><td>' + (l['Style Number'] || '-') + '</td><td>' + (l['Color'] || '-') + '</td>' +
      '<td style="text-align:right">' + q + '</td><td style="text-align:right">₹' + rt.toFixed(2) + '</td><td style="text-align:right">₹' + (q * rt).toFixed(2) + '</td></tr>';
  }).join('');
  return pdfShell('CMT Bill', '#8e44ad',
    '<div class="title-bar">Bill #: <strong>' + f['Bill No'] + '</strong> | Date: ' + f['Date'] + ' | Status: ' + f['Status'] + '</div>' +
    '<table><tr><th>Department</th><td>' + (f['Department'] || '-') + '</td><th>Contractor</th><td>' + (f['Contractor'] || '-') + '</td></tr></table>' +
    '<table style="margin-top:10px"><thead><tr><th>#</th><th>JO</th><th>Style</th><th>Colour</th><th>Qty</th><th>Rate/Pc</th><th>Amount</th></tr></thead><tbody>' + rowsHtml +
    '<tr><th colspan="6" style="text-align:right">Total</th><td style="text-align:right"><strong>₹' + total.toFixed(2) + '</strong></td></tr></tbody></table>' +
    '<table style="margin-top:10px"><tr><th>Approved By</th><td>' + (f['Approved By'] || '-') + '</td><th>Approval Date</th><td>' + (f['Approval Date'] || '-') + '</td></tr></table>');
}
function authorizeJiwinERP() {
  DriveApp.getRootFolder().getName();
  SpreadsheetApp.getActiveSpreadsheet().getName();
  return 'JIWIN ERP authorization completed.';
}

/*************** OIH (ORDER IN HAND) — MERGED FROM STANDALONE OIH APP ***************/
/**
 * The ORDERS/PRICING/IMAGES/CONFIG/TECHPACKS/REQS sheets below are OIH's own
 * data (Orders, Samples, Planning, Requirements, buyer pricing) — same Sheet
 * as everything else in this project now, so JO Number is the shared key
 * between OIH's order lines and Fabric ERP's Style Tracker.
 */

/**
 * JIWIN APPARELS - ORDER IN HAND (OIH)  |  Code.gs
 * Bound to the OIH Google Sheet.  Front end = Index.html
 *
 * WHO SEES WHAT
 *  - Everyone (merch team) : Dashboard (buyer-wise qty in hand), Orders, Planning. NO price / budget / amount data is ever sent to them.
 *  - Management only       : Management tab. Needs the management PIN. Checked here on the server for every call.
 *
 * SETUP (once)
 *  1. Sheet menu  Jiwin OIH > 1. Setup sheets
 *  2. Sheet menu  Jiwin OIH > 2. Load seed data (one time)   -> then DELETE Seed.gs
 *  3. Sheet menu  Jiwin OIH > 3. Set management PIN
 *  4. Deploy > New deployment > Web app > Execute as: Me > Who has access: your team
 *  5. Keep the Google Sheet itself shared with management only (merch use the web app link, not the sheet).
 */

var SHEET = { ORDERS: 'ORDERS', PRICING: 'PRICING', IMAGES: 'IMAGES', CONFIG: 'CONFIG', TECHPACKS: 'TECHPACKS', REQS: 'REQS' };

// ORDERS columns (order matters)
var OCOLS = ['id', 'status', 'jo', 'ship', 'buyer', 'style', 'styling', 'fabric', 'color', 'qty',
  'yarnIH', 'fabricIH', 'trimsIH', 'ppAD', 'cutDate', 'cutDays', 'printStart', 'printDays',
  'proStart', 'proDays', 'fiDays', 'remarks', 'updatedBy', 'updatedAt', 'kind'];
var OHEAD = ['ID', 'Status', 'JO Number', 'Shipment Date', 'Buyer', 'Style', 'Styling', 'Fabric', 'Color', 'Order Qty',
  'Yarn Inhouse', 'Fabric Inhouse', 'Trims IH Date', 'PP AD', 'Cutting Date', 'Cut Days', 'Print Start', 'Print Days',
  'Pro Start', 'Pro Days', 'FI Days', 'Remarks', 'Updated By', 'Updated At', 'Type (ORDER/SAMPLE)'];
var DATE_FIELDS = ['ship', 'yarnIH', 'fabricIH', 'trimsIH', 'ppAD', 'cutDate', 'printStart', 'proStart'];
var NUM_FIELDS = ['qty', 'cutDays', 'printDays', 'proDays', 'fiDays'];
var TEXT_FIELDS = ['status', 'jo', 'buyer', 'style', 'styling', 'fabric', 'color', 'remarks', 'kind'];

var PCOLS = ['id', 'price', 'wc', 'ok'];
var PHEAD = ['ID', 'Buyer Price (USD)', '10% WC', 'Budget OK'];

// Fabric / Trims requirements per style
var RCOLS = ['id', 'key', 'team', 'item', 'spec', 'colour', 'perPc', 'unit', 'wastage', 'needBy', 'status', 'remarks', 'updatedBy', 'updatedAt', 'sentAt', 'manualQty', 'tpx', 'createdBy', 'kind'];
var RHEAD = ['ID', 'Style Key (BUYER|STYLE)', 'Team', 'Item', 'Spec / Details', 'Colour', 'Per Pc', 'Unit', 'Wastage %', 'Needed By', 'Status', 'Remarks', 'Updated By', 'Updated At', 'Sent To Team At', 'Required Qty (typed)', 'TPX No.', 'Requested By', 'For (ORDER/SAMPLE)'];
var RTEXT = ['key', 'team', 'item', 'spec', 'colour', 'unit', 'status', 'remarks', 'sentAt', 'tpx', 'createdBy', 'kind'];
var DEFAULT_REQ = 'REQUESTED,ORDERED,PARTIAL,SUBMITTED,ACCEPTED,CANCELLED';
// Reminders go to the fabric / trims team while an item is in one of these, and to merch while it is SUBMITTED.
var PENDING_TEAM = { 'REQUESTED': 1, 'ORDERED': 1, 'PARTIAL': 1 };

var DEFAULT_STATUSES = 'TBS,IN PROGRESS,HOLD,SHIPPED,CANCELLED';
var DEFAULT_CLOSED = 'SHIPPED,CANCELLED';

/* ============================== MENU / SETUP ============================== */


function setupSheets() {
  var ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, SHEET.ORDERS, OHEAD);
  ss.getSheetByName(SHEET.ORDERS).getRange(1, 1, 1, OHEAD.length).setValues([OHEAD]).setBackground('#1B2A4A').setFontColor('#FFFFFF').setFontWeight('bold');
  ensureSheet_(ss, SHEET.PRICING, PHEAD);
  ensureSheet_(ss, SHEET.IMAGES, ['Line ID', 'Image (data URL)', 'Updated At']);
  var cfg = ensureSheet_(ss, SHEET.CONFIG, ['Key', 'Value', 'Note']);
  if (cfg.getLastRow() < 2) {
    cfg.getRange(2, 1, 3, 3).setValues([
      ['statuses', DEFAULT_STATUSES, 'Comma separated. Order shown in dropdowns.'],
      ['closedStatuses', DEFAULT_CLOSED, 'Not counted in "quantity in hand".'],
      ['usdInr', 93, 'USD to INR rate used in Management tab.']
    ]);
  }
  sheetOrCreate_(SHEET.TECHPACKS);
  var rq = ensureSheet_(ss, SHEET.REQS, RHEAD);
  rq.getRange(1, 1, 1, RHEAD.length).setValues([RHEAD]).setBackground('#1B2A4A').setFontColor('#FFFFFF').setFontWeight('bold');
  ensureConfig_(cfg, 'reqStatuses', DEFAULT_REQ, 'Comma separated statuses for fabric / trims requirements.');
  ensureConfig_(cfg, 'fabricTeamEmail', '', 'Email(s) for the "Email fabric team" button (comma separated).');
  ensureConfig_(cfg, 'trimsTeamEmail', '', 'Email(s) for the "Email trims team" button (comma separated).');
  ensureConfig_(cfg, 'fabricTeamMobile', '', 'WhatsApp / SMS number(s) of the fabric team, comma separated (10 digits or with country code).');
  ensureConfig_(cfg, 'trimsTeamMobile', '', 'WhatsApp / SMS number(s) of the trims team, comma separated.');
  ensureConfig_(cfg, 'merchTeamMobile', '', 'WhatsApp / SMS number(s) of the merch team, comma separated.');
  ensureConfig_(cfg, 'merchTeamEmail', '', 'Merch team email(s). Reminded daily while a requirement is SUBMITTED but not yet ACCEPTED.');
  ensureConfig_(cfg, 'techpackShare', 'ANYONE', 'ANYONE = anyone with the link can view tech pack PDFs. DOMAIN = only your Google Workspace domain.');
  ensureConfig_(cfg, 'productionTeamEmail', '', 'Production team email(s), comma separated — for JO checklist notifications.');
  ensureConfig_(cfg, 'storeTeamEmail', '', 'Store team email(s), comma separated — for JO checklist notifications.');
  RCOLS.forEach(function (k, i) {
    var r = rq.getRange(2, i + 1, rq.getMaxRows() - 1, 1);
    if (RTEXT.indexOf(k) > -1 || k === 'id' || k === 'updatedBy') r.setNumberFormat('@');
    else if (k === 'needBy') r.setNumberFormat('dd-mmm-yy');
  });
  repairIds_(ss);
  // text columns as plain text so nothing typed can turn into a formula
  var o = ss.getSheetByName(SHEET.ORDERS);
  OCOLS.forEach(function (k, i) {
    var col = i + 1, rng = o.getRange(2, col, o.getMaxRows() - 1, 1);
    if (TEXT_FIELDS.indexOf(k) > -1 || k === 'id' || k === 'updatedBy') rng.setNumberFormat('@');
    else if (DATE_FIELDS.indexOf(k) > -1) rng.setNumberFormat('dd-mmm-yy');
    else if (k === 'qty') rng.setNumberFormat('#,##0');
  });
  o.setFrozenRows(1);
  SpreadsheetApp.getActive().toast('Sheets ready. Next: Load seed data, then Set management PIN.', 'Jiwin OIH', 8);
}

/** Returns the sheet, creating REQS / TECHPACKS on the fly if Setup was not re-run after an update. */
function sheetOrCreate_(name) {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(name);
  if (sh) return sh;
  if (name === SHEET.REQS) return ensureSheet_(ss, name, RHEAD);
  if (name === SHEET.TECHPACKS) return ensureSheet_(ss, name, ['Style Key', 'Drive File ID', 'File Name', 'Updated By', 'Updated At']);
  throw new Error('Sheet ' + name + ' is missing. Run Jiwin OIH > Setup sheets.');
}

/** IDs like 932451e3 get turned into numbers by Sheets. Restore them and force text format. */
function repairIds_(ss) {
  var o = ss.getSheetByName(SHEET.ORDERS), last = o.getLastRow(), map = {}, fixed = 0;
  if (last > 1) o.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
    var id = String(r[0]); if (/^\d+e\d+$/i.test(id)) map[String(Number(id))] = id;
  });
  [SHEET.PRICING, SHEET.IMAGES].forEach(function (n) {
    var sh = ss.getSheetByName(n); if (!sh) return;
    sh.getRange(2, 1, sh.getMaxRows() - 1, 1).setNumberFormat('@');
    if (sh.getLastRow() < 2) return;
    var rng = sh.getRange(2, 1, sh.getLastRow() - 1, 1), v = rng.getValues(), ch = false;
    v.forEach(function (r) { if (typeof r[0] === 'number') { var id = map[String(r[0])]; if (id) { r[0] = id; ch = true; fixed++; } } });
    if (ch) rng.setValues(v);
  });
  return fixed;
}

function ensureConfig_(sh, key, val, note) {
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (v[i][0] === key) return;
  sh.appendRow([key, val, note]);
}

/** Run once from the menu so the web app may use Drive (tech pack PDFs) and Mail (requirement emails). */
function authorizeServices() {
  DriveApp.getRootFolder();
  MailApp.getRemainingDailyQuota();
  SpreadsheetApp.getUi().alert('Drive and Mail access authorised.\n\nNow re-deploy: Deploy > Manage deployments > pencil > Version: New version > Deploy.');
}

function ensureSheet_(ss, name, head) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, head.length).setValues([head])
      .setBackground('#1B2A4A').setFontColor('#FFFFFF').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function setAnthropicApiKey() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Anthropic API Key', 'Paste your Anthropic API key (from console.anthropic.com). Used only for the "Import PO" auto-fill feature — stored privately in this script\'s properties, never shown in the sheet.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var key = String(r.getResponseText()).trim();
  if (!key) { ui.alert('No key entered.'); return; }
  PropertiesService.getScriptProperties().setProperty('ANTHROPIC_API_KEY', key);
  ui.alert('Anthropic API key saved. "Import PO" auto-fill is now active.');
}

function setManagementPin() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Management PIN', 'Enter a new PIN (min 4 characters). Only management should know it.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var pin = String(r.getResponseText()).trim();
  if (pin.length < 4) { ui.alert('PIN must be at least 4 characters.'); return; }
  var salt = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperties({ MGMT_SALT: salt, MGMT_HASH: hashPin_(pin, salt) });
  CacheService.getScriptCache().remove('mgmt_fail');
  ui.alert('Management PIN saved.');
}

function hashPin_(pin, salt) {
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pin, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(d);
}

/* ================================== WEB APP ================================= */


/* ============================ PUBLIC API (no price) ========================== */

function getInitData() {
  var cfg = readConfig_();
  return {
    statuses: cfg.statuses,
    closed: cfg.closed,
    reqStatuses: cfg.reqStatuses,
    mobiles: cfg.mobiles,
    orders: readOrders_(),
    images: readImages_(),
    techpacks: readTechpacks_(),
    reqs: readReqs_()
  };
}

function getOrders() { return readOrders_(); }

/**
 * Upsert lines. Each item: {id?, ...fields}. Only the fields present are changed on existing lines.
 * Returns the saved lines as stored.
 */
function saveLines(items, userName) {
  if (!items || !items.length) return { ok: true, rows: [] };
  if (items.length > 300) throw new Error('Too many lines in one save (max 300).');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET.ORDERS);
    var last = sh.getLastRow();
    var idIndex = {};
    if (last > 1) {
      sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r, i) { idIndex[String(r[0])] = i + 2; });
    }
    var now = new Date(), who = clean_(userName, 40) || 'unknown';
    var saved = [], newRows = [];
    items.forEach(function (it) {
      var id = it.id ? String(it.id) : '';
      var rowNo = id ? idIndex[id] : 0;
      var cur;
      if (rowNo) cur = rowToObj_(sh.getRange(rowNo, 1, 1, OCOLS.length).getValues()[0]);
      else { cur = {}; OCOLS.forEach(function (k) { cur[k] = ''; }); cur.id = id || Utilities.getUuid().slice(0, 8); cur.status = 'TBS'; cur.kind = 'ORDER'; }
      applyFields_(cur, it);
      if (!cur.kind) cur.kind = 'ORDER';
      if (!cur.buyer || (!cur.jo && cur.kind !== 'SAMPLE')) throw new Error('JO number and Buyer are required.');
      cur.updatedBy = who; cur.updatedAt = Utilities.formatDate(now, tz_(), 'yyyy-MM-dd HH:mm');
      var arr = objToRow_(cur);
      if (rowNo) sh.getRange(rowNo, 1, 1, OCOLS.length).setValues([arr]);
      else newRows.push(arr);
      saved.push(rowToObj_(arr));
    });
    if (newRows.length) {
      var start = Math.max(sh.getLastRow(), 1) + 1;
      sh.getRange(start, 1, newRows.length, OCOLS.length).setValues(newRows);
    }
    SpreadsheetApp.flush();
    return { ok: true, rows: saved.map(fmtOut_) };
  } finally {
    lock.releaseLock();
  }
}

function saveImage(lineId, dataUrl) {
  lineId = String(lineId || '');
  if (!lineId) throw new Error('Missing line id.');
  if (!/^data:image\/(jpeg|png|webp);base64,/.test(dataUrl || '')) throw new Error('Bad image.');
  if (dataUrl.length > 48000) throw new Error('Image too large after compression.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET.IMAGES);
    var last = sh.getLastRow(), row = 0;
    if (last > 1) {
      var ids = sh.getRange(2, 1, last - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === lineId) { row = i + 2; break; }
    }
    if (!row) row = last + 1;
    sh.getRange(row, 1, 1, 3).setValues([[lineId, dataUrl, new Date()]]);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ============================ TECH PACKS (PDF in Drive) ============================ */

function normKey_(k) { return String(k || '').replace(/[\r\n\t]+/g, ' ').trim().toUpperCase().slice(0, 160); }

function readTechpacks_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET.TECHPACKS), out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().forEach(function (r) {
    if (r[0] && r[1]) out[String(r[0])] = { id: String(r[1]), name: String(r[2]), by: String(r[3]), at: String(r[4] instanceof Date ? Utilities.formatDate(r[4], tz_(), 'yyyy-MM-dd') : r[4]) };
  });
  return out;
}

function techFolder_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('TP_FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { } }
  var f = DriveApp.createFolder('OIH Techpacks');
  p.setProperty('TP_FOLDER', f.getId());
  return f;
}

/** One PDF per style (BUYER|STYLE). Uploading again replaces the old file. */
function saveTechpack(key, fileName, b64, userName) {
  key = normKey_(key);
  if (key.indexOf('|') < 1) throw new Error('Bad style key.');
  var name = clean_(fileName, 100).replace(/[\\/:*?"<>|]+/g, '_');
  if (!/\.pdf$/i.test(name)) throw new Error('Only PDF files are allowed.');
  if (!b64 || b64.length > 14000000) throw new Error('PDF too large (max 10 MB).');
  var bytes = Utilities.base64Decode(b64);
  if (bytes.length > 10500000) throw new Error('PDF too large (max 10 MB).');
  if (!(bytes[0] === 37 && bytes[1] === 80 && bytes[2] === 68 && bytes[3] === 70)) throw new Error('That file is not a valid PDF.');
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var file = techFolder_().createFile(Utilities.newBlob(bytes, 'application/pdf', key.replace('|', ' - ') + ' - ' + name));
    var shared = true;
    try {
      file.setSharing(readConfig_().techShare === 'DOMAIN' ? DriveApp.Access.DOMAIN_WITH_LINK : DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) { shared = false; }
    var sh = sheetOrCreate_(SHEET.TECHPACKS), last = sh.getLastRow(), row = 0, oldId = '';
    if (last > 1) {
      var v = sh.getRange(2, 1, last - 1, 2).getValues();
      for (var i = 0; i < v.length; i++) if (String(v[i][0]) === key) { row = i + 2; oldId = String(v[i][1]); break; }
    }
    if (!row) row = last + 1;
    var at = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm'), by = clean_(userName, 40) || 'unknown';
    sh.getRange(row, 1, 1, 5).setValues([[key, file.getId(), name, by, at]]);
    if (oldId) { try { DriveApp.getFileById(oldId).setTrashed(true); } catch (e) { } }
    return { ok: true, shared: shared, tp: { id: file.getId(), name: name, by: by, at: at } };
  } finally { lock.releaseLock(); }
}

/* ========================= FABRIC / TRIMS REQUIREMENTS ========================= */

function readReqs_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET.REQS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, RCOLS.length).getValues()
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) { var o = {}; RCOLS.forEach(function (k, i) { o[k] = r[i]; }); return fmtReq_(o); });
}

function fmtReq_(o) {
  var out = {};
  RCOLS.forEach(function (k) {
    var v = o[k];
    if (k === 'needBy') out[k] = dateOut_(v);
    else if (k === 'perPc' || k === 'wastage' || k === 'manualQty') out[k] = numOrBlank_(v);
    else if (k === 'kind') out[k] = String(v).toUpperCase() === 'SAMPLE' ? 'SAMPLE' : 'ORDER';
    else if (k === 'updatedAt' || k === 'sentAt') out[k] = (v instanceof Date) ? Utilities.formatDate(v, tz_(), 'yyyy-MM-dd HH:mm') : String(v || '');
    else out[k] = String(v === null || v === undefined ? '' : v);
  });
  return out;
}

/** items: [{id, key, team, item, spec, colour, perPc, unit, wastage, needBy, status, remarks}] - upsert, only provided fields change. */
function saveReq(items, userName) {
  if (!items || !items.length) return { ok: true, rows: [] };
  if (items.length > 100) throw new Error('Too many rows in one save.');
  var statuses = readConfig_().reqStatuses;
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheetOrCreate_(SHEET.REQS), last = sh.getLastRow(), idx = {};
    if (last > 1) sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r, i) { idx[String(r[0])] = i + 2; });
    var now = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm'), who = clean_(userName, 40) || 'unknown';
    var saved = [], fresh = [];
    items.forEach(function (it) {
      var id = it.id ? String(it.id) : '', rowNo = id ? idx[id] : 0, cur;
      if (rowNo) { cur = {}; sh.getRange(rowNo, 1, 1, RCOLS.length).getValues()[0].forEach(function (v, i) { cur[RCOLS[i]] = v; }); }
      else { cur = {}; RCOLS.forEach(function (k) { cur[k] = ''; }); cur.id = id || Utilities.getUuid().slice(0, 8); cur.status = statuses[0]; cur.createdBy = who; cur.kind = 'ORDER'; }
      if (it.hasOwnProperty('key')) cur.key = normKey_(it.key);
      if (it.hasOwnProperty('team')) cur.team = String(it.team).toUpperCase();
      if (it.hasOwnProperty('item')) cur.item = clean_(it.item, 80);
      if (it.hasOwnProperty('spec')) cur.spec = clean_(it.spec, 200);
      if (it.hasOwnProperty('colour')) cur.colour = clean_(it.colour, 60);
      if (it.hasOwnProperty('tpx')) cur.tpx = clean_(it.tpx, 40);
      if (it.hasOwnProperty('kind')) cur.kind = String(it.kind).toUpperCase() === 'SAMPLE' ? 'SAMPLE' : 'ORDER';
      if (it.hasOwnProperty('unit')) cur.unit = clean_(it.unit, 10).toUpperCase();
      if (it.hasOwnProperty('remarks')) cur.remarks = clean_(it.remarks, 200);
      if (it.hasOwnProperty('status')) {
        var st = String(it.status).trim().toUpperCase();
        if (statuses.indexOf(st) < 0) throw new Error('Unknown status: ' + st);
        cur.status = st;
      }
      ['perPc', 'wastage', 'manualQty'].forEach(function (k) {
        if (!it.hasOwnProperty(k)) return;
        if (it[k] === '' || it[k] === null) { cur[k] = ''; return; }
        var n = Number(it[k]);
        if (!isFinite(n) || n < 0 || n > 100000000) throw new Error('Bad number for ' + k);
        cur[k] = n;
      });
      if (it.hasOwnProperty('needBy')) {
        var d = String(it.needBy || '').trim();
        if (d && !/^\d{4}-\d{2}-\d{2}$/.test(d)) throw new Error('Bad date.');
        cur.needBy = d;
      }
      if (cur.key.indexOf('|') < 1) throw new Error('Bad style key.');
      if (cur.team !== 'FABRIC' && cur.team !== 'TRIMS') throw new Error('Team must be FABRIC or TRIMS.');
      if (!cur.item) throw new Error('Item name is required.');
      cur.updatedBy = who; cur.updatedAt = now;
      var arr = RCOLS.map(function (k) {
        if (k === 'needBy') return dateIn_(cur[k]);
        if (k === 'perPc' || k === 'wastage' || k === 'manualQty') return numOrBlank_(cur[k]);
        return cur[k] === null || cur[k] === undefined ? '' : cur[k];
      });
      if (rowNo) sh.getRange(rowNo, 1, 1, RCOLS.length).setValues([arr]); else fresh.push(arr);
      var o = {}; RCOLS.forEach(function (k, i) { o[k] = arr[i]; });
      saved.push(fmtReq_(o));
    });
    if (fresh.length) sh.getRange(Math.max(sh.getLastRow(), 1) + 1, 1, fresh.length, RCOLS.length).setValues(fresh);
    SpreadsheetApp.flush();
    return { ok: true, rows: saved };
  } finally { lock.releaseLock(); }
}

function deleteReq(ids) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var set = {}; (ids || []).forEach(function (i) { set[String(i)] = true; });
    var sh = sheetOrCreate_(SHEET.REQS), last = sh.getLastRow();
    if (last < 2) return { ok: true };
    var v = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = v.length - 1; i >= 0; i--) if (set[String(v[i][0])]) sh.deleteRow(i + 2);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/**
 * ONE click for the whole list: emails every NEW requirement (status REQUESTED, not yet sent) to the
 * fabric team and to the trims team - one email per team, grouped by style. Marks them as sent.
 */
function emailAllRequirements(userName) {
  var cfg = readConfig_();
  var allOrders = readOrders_(), allBy = {};
  allOrders.forEach(function (o) { allBy[kk_(normKey_(o.buyer + '|' + o.style), o.kind)] = true; });
  var open = allOrders.filter(function (o) { return cfg.closed.indexOf(o.status) < 0; });
  var byKey = {};
  open.forEach(function (o) { var k = kk_(normKey_(o.buyer + '|' + o.style), o.kind); (byKey[k] = byKey[k] || []).push(o); });
  // styles with no entry yet (e.g. sample requests) are allowed; styles whose lines are all shipped/cancelled are skipped
  var all = readReqs_().filter(function (r) { var k = kk_(r.key, r.kind); return (byKey[k] || !allBy[k]) && r.status === cfg.reqStatuses[0] && !r.sentAt; });
  if (!all.length) throw new Error('Nothing new to send.');
  var teams = [['FABRIC', cfg.fabricEmail, 'fabricTeamEmail', 'Fabric'], ['TRIMS', cfg.trimsEmail, 'trimsTeamEmail', 'Trims']];
  teams.forEach(function (t) {
    if (all.some(function (r) { return r.team === t[0]; }) && !t[1]) throw new Error('Add ' + t[2] + ' in the CONFIG sheet first.');
  });
  var tps = readTechpacks_(), url = ''; try { url = ScriptApp.getService().getUrl(); } catch (e) { }
  var who = clean_(userName, 40) || 'Merchandising';
  var num = function (n) { return n >= 100 ? String(Math.round(n)) : String(Math.round(n * 100) / 100); };
  var sentIds = [], summary = [], failure = '';
  teams.forEach(function (t) {
    if (failure) return;
    var items = all.filter(function (r) { return r.team === t[0]; });
    if (!items.length) return;
    var keys = [];
    items.forEach(function (r) { var k = kk_(r.key, r.kind); if (keys.indexOf(k) < 0) keys.push(k); });
    keys.sort();
    var body = 'Hello ' + t[3] + ' team,\n\nPlease arrange the following ' + t[0].toLowerCase() + ' requirements (' + items.length + ' items, ' + keys.length + ' styles).\n';
    keys.forEach(function (k) {
      var kind = k.split('#')[1], kkey = k.split('#')[0];
      var ls = byKey[k] || [], f = ls[0] || { buyer: kkey.split('|')[0], style: kkey.split('|').slice(1).join('|'), styling: '' }, jos = [], ships = [];
      ls.forEach(function (o) { if (jos.indexOf(o.jo) < 0) jos.push(o.jo); if (o.ship) ships.push(o.ship); });
      ships.sort();
      var qtyFor = function (colour) { return ls.reduce(function (n, o) { return n + ((!colour || o.color.toLowerCase() === colour.toLowerCase()) ? Number(o.qty) || 0 : 0); }, 0); };
      body += '\n' + (kind === 'SAMPLE' ? '[SAMPLE] ' : '[ORDER] ') + f.buyer + ' ' + f.style + (f.styling ? ' (' + f.styling + ')' : '') +
        (ls.length ? ' | ' + jos.filter(Boolean).join(', ') + ' | ' + num(qtyFor('')) + ' pcs' + (ships.length ? ' | ship ' + ships[0] : '') : (kind === 'SAMPLE' ? ' | sample request' : ' | (no order entered yet)')) + '\n';
      if (tps[kkey]) body += 'Tech pack: https://drive.google.com/file/d/' + tps[kkey].id + '/view\n';
      items.filter(function (r) { return kk_(r.key, r.kind) === k; }).forEach(function (r) {
        var total = r.manualQty !== '' ? ' | REQUIRED QTY ' + num(Number(r.manualQty)) + ' ' + r.unit
          : (r.perPc === '' ? '' : ' | total ' + num(Number(r.perPc) * qtyFor(r.colour) * (1 + (Number(r.wastage) || 0) / 100)) + ' ' + r.unit);
        body += '  - ' + r.item + (r.spec ? ' - ' + r.spec : '') + (r.colour ? ' | ' + r.colour : ' | all colours') + (r.tpx ? ' | TPX/TCX ' + r.tpx : '') +
          (r.manualQty === '' && r.perPc !== '' ? ' | ' + r.perPc + ' ' + r.unit + '/pc' : '') + (r.manualQty === '' && r.wastage !== '' ? ' + ' + r.wastage + '% wastage' : '') + total +
          (r.needBy ? ' | needed by ' + r.needBy : '') + (r.remarks ? ' | ' + r.remarks : '') + (r.createdBy ? ' | requested by ' + r.createdBy : '') + '\n';
      });
    });
    body += '\n' + (url ? 'Set the status to SUBMITTED here when done: ' + url + '\n\n' : '') + 'Regards,\n' + who + '\nJiwin Apparels';
    try {
      MailApp.sendEmail({ to: t[1], subject: t[0] + ' requirements - ' + items.length + ' items, ' + keys.length + ' styles', body: body, name: 'Jiwin OIH' });
      items.forEach(function (r) { sentIds.push(r.id); });
      summary.push({ team: t[0], to: t[1], count: items.length, styles: keys.length });
    } catch (e) { failure = t[3] + ' email failed: ' + e.message; }
  });
  var rows = markSent_(sentIds);
  if (failure) throw new Error(failure + (sentIds.length ? ' (the other team was sent and marked)' : ''));
  return { ok: true, count: sentIds.length, teams: summary, rows: rows };
}

/** Used after the merch team sent the message by WhatsApp / SMS themselves. */
function kk_(key, kind) { return key + '#' + (String(kind).toUpperCase() === 'SAMPLE' ? 'SAMPLE' : 'ORDER'); }

function markRequirementsSent(ids, userName) {
  if (!ids || !ids.length || ids.length > 300) throw new Error('Nothing to mark.');
  return { ok: true, rows: markSent_(ids.map(String)) };
}

function markSent_(ids) {
  if (!ids.length) return [];
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheetOrCreate_(SHEET.REQS), last = sh.getLastRow(), set = {}, col = RCOLS.indexOf('sentAt') + 1;
    ids.forEach(function (i) { set[i] = true; });
    var at = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm');
    sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r, i) { if (set[String(r[0])]) sh.getRange(i + 2, col).setValue(at); });
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  return readReqs_().filter(function (r) { return set[r.id]; });
}

/* ============================== DAILY REMINDERS ============================== */
/**
 * Every day (trigger) until an item is ACCEPTED (or CANCELLED):
 *   - REQUESTED / ORDERED / PARTIAL -> email to the fabric / trims team ("please submit")
 *   - SUBMITTED                     -> email to the merch team ("please accept")
 * Styles whose lines are all shipped / cancelled are skipped.
 */
function remindersOn() {
  remindersOff_();
  ScriptApp.newTrigger('sendDailyReminders').timeBased().everyDays(1).atHour(9).create();
  SpreadsheetApp.getUi().alert('Daily reminders are ON (around 9 AM every day).\n\nFill fabricTeamEmail, trimsTeamEmail and merchTeamEmail in the CONFIG tab.');
}
function remindersOff() { remindersOff_(); SpreadsheetApp.getUi().alert('Daily reminders are OFF.'); }
function remindersOff_() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'sendDailyReminders') ScriptApp.deleteTrigger(t); });
}
function sendRemindersNow() {
  var r = remind_(true);
  SpreadsheetApp.getUi().alert('Reminder emails sent: ' + r.sent.length + (r.sent.length ? '\n' + r.sent.join('\n') : '\n(nothing pending, or team emails missing in CONFIG)'));
}
/** Trigger entry point. Guarded so it cannot be spammed if called from the browser. */
function sendDailyReminders() { return remind_(false); }

function daysBetween_(a, b) {
  var p = function (s) { var x = s.split('-'); return Date.UTC(+x[0], +x[1] - 1, +x[2]); };
  return Math.round((p(a) - p(b)) / 864e5);
}

function remind_(force) {
  var cache = CacheService.getScriptCache();
  if (!force) { if (cache.get('rem_last')) return { ok: true, sent: [], skipped: true }; cache.put('rem_last', '1', 19800); }
  var cfg = readConfig_();
  var allOrders = readOrders_(), allBy = {};
  allOrders.forEach(function (o) { allBy[kk_(normKey_(o.buyer + '|' + o.style), o.kind)] = true; });
  var open = allOrders.filter(function (o) { return cfg.closed.indexOf(o.status) < 0; });
  var byKey = {};
  open.forEach(function (o) { var k = kk_(normKey_(o.buyer + '|' + o.style), o.kind); (byKey[k] = byKey[k] || []).push(o); });
  var reqs = readReqs_().filter(function (r) { var k = kk_(r.key, r.kind); return byKey[k] || !allBy[k]; });
  var today = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
  var url = ''; try { url = ScriptApp.getService().getUrl(); } catch (e) { }
  var num = function (n) { return n >= 100 ? String(Math.round(n)) : String(Math.round(n * 100) / 100); };
  var sent = [];

  var line = function (r) {
    var ls = byKey[kk_(r.key, r.kind)] || [], f = ls[0] || { buyer: r.key.split('|')[0], style: r.key.split('|').slice(1).join('|') }, jos = [];
    ls.forEach(function (o) { if (o.jo && jos.indexOf(o.jo) < 0) jos.push(o.jo); });
    var basis = ls.reduce(function (t, o) { return t + ((!r.colour || o.color.toLowerCase() === r.colour.toLowerCase()) ? Number(o.qty) || 0 : 0); }, 0);
    var total = r.manualQty !== '' ? ' | qty ' + num(Number(r.manualQty)) + ' ' + r.unit
      : (r.perPc === '' ? '' : ' | total ' + num(Number(r.perPc) * basis * (1 + (Number(r.wastage) || 0) / 100)) + ' ' + r.unit);
    var due = '';
    if (r.needBy) { var d = daysBetween_(r.needBy, today); due = ' | needed by ' + r.needBy + (d < 0 ? ' (OVERDUE ' + (-d) + ' days)' : d === 0 ? ' (TODAY)' : ' (' + d + ' days left)'); }
    return '- ' + (r.kind === 'SAMPLE' ? '[SAMPLE] ' : '[ORDER] ') + f.buyer + ' ' + f.style + (jos.length ? ' (' + jos.join(', ') + ')' : '') + ' | ' + r.item + (r.spec ? ' - ' + r.spec : '') + (r.colour ? ' | ' + r.colour : '') + (r.tpx ? ' | TPX/TCX ' + r.tpx : '') + total + due + ' | ' + r.status + (r.remarks ? ' | ' + r.remarks : '') + (r.createdBy ? ' | requested by ' + r.createdBy : '');
  };
  var order = function (a, b) {
    if (a.needBy !== b.needBy) { if (!a.needBy) return 1; if (!b.needBy) return -1; return a.needBy < b.needBy ? -1 : 1; }
    return a.key < b.key ? -1 : 1;
  };
  var overdue = function (list) { return list.filter(function (r) { return r.needBy && r.needBy < today; }).length; };
  var mail = function (to, subject, intro, list) {
    if (!to || !list.length) return;
    list.sort(order);
    var body = intro + '\n\n' + list.map(line).join('\n') + '\n\n' + (url ? 'Update status here: ' + url + '\n\n' : '') + 'This reminder comes every day until the items are closed.\nJiwin Apparels - OIH';
    MailApp.sendEmail({ to: to, subject: subject, body: body, name: 'Jiwin OIH' });
    sent.push(subject + ' -> ' + to);
  };

  [['FABRIC', cfg.fabricEmail, 'Fabric'], ['TRIMS', cfg.trimsEmail, 'Trims']].forEach(function (t) {
    var list = reqs.filter(function (r) { return r.team === t[0] && PENDING_TEAM[r.status] && r.sentAt; });
    mail(t[1], 'Reminder: ' + list.length + ' ' + t[0].toLowerCase() + ' requirement(s) pending' + (overdue(list) ? ' (' + overdue(list) + ' overdue)' : ''),
      'Hello ' + t[2] + ' team,\n\nThese requirements are still open. Please arrange them and set the status to SUBMITTED once done. Merch will then accept.', list);
  });
  var sub = reqs.filter(function (r) { return r.status === 'SUBMITTED'; });
  mail(cfg.merchEmail, 'Reminder: ' + sub.length + ' submitted item(s) waiting for merch acceptance',
    'Hello Merch team,\n\nThe fabric / trims teams have submitted these items. Please check and set the status to ACCEPTED (or send back).', sub);
  return { ok: true, sent: sent };
}

/* ========================= MANAGEMENT API (PIN + token) ====================== */

function mgmtLogin(pin) {
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('mgmt_fail') || 0);
  if (fails >= 8) return { ok: false, msg: 'Too many wrong attempts. Try again in 5 minutes.' };
  var p = PropertiesService.getScriptProperties();
  var salt = p.getProperty('MGMT_SALT'), hash = p.getProperty('MGMT_HASH');
  if (!salt || !hash) return { ok: false, msg: 'Management PIN not set yet (Sheet menu > Jiwin OIH > Set management PIN).' };
  if (hashPin_(String(pin || '').trim(), salt) !== hash) {
    cache.put('mgmt_fail', String(fails + 1), 300);
    return { ok: false, msg: 'Wrong PIN.' };
  }
  cache.remove('mgmt_fail');
  var token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('mt_' + token, '1', 7200);
  return { ok: true, token: token };
}

function mgmtLogout(token) { CacheService.getScriptCache().remove('mt_' + token); return { ok: true }; }

function getPricing(token) {
  requireMgmt_(token);
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET.PRICING);
  var last = sh.getLastRow(), map = {};
  if (last > 1) {
    sh.getRange(2, 1, last - 1, PCOLS.length).getValues().forEach(function (r) {
      if (r[0] !== '') map[String(r[0])] = { price: numOrBlank_(r[1]), wc: numOrBlank_(r[2]), ok: String(r[3] || '') };
    });
  }
  return { ok: true, pricing: map, usdInr: readConfig_().usdInr };
}

/** items: [{id, price, wc, ok}] */
function savePricing(token, items) {
  requireMgmt_(token);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET.PRICING);
    var last = sh.getLastRow(), idx = {};
    if (last > 1) sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r, i) { idx[String(r[0])] = i + 2; });
    (items || []).forEach(function (it) {
      var id = String(it.id || ''); if (!id) return;
      var arr = [id, numOrBlank_(it.price), numOrBlank_(it.wc), String(it.ok || '').slice(0, 10)];
      var row = idx[id];
      if (!row) { row = sh.getLastRow() + 1; idx[id] = row; }
      sh.getRange(row, 1, 1, 4).setValues([arr]);
    });
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function saveUsdInr(token, rate) {
  requireMgmt_(token);
  rate = Number(rate);
  if (!(rate > 0)) throw new Error('Invalid rate.');
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET.CONFIG);
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (v[i][0] === 'usdInr') { sh.getRange(i + 1, 2).setValue(rate); return { ok: true }; }
  sh.appendRow(['usdInr', rate, '']);
  return { ok: true };
}

function deleteLines(token, ids) {
  requireMgmt_(token);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var set = {}; (ids || []).forEach(function (i) { set[String(i)] = true; });
    [SHEET.ORDERS, SHEET.PRICING, SHEET.IMAGES].forEach(function (name) {
      var sh = SpreadsheetApp.getActive().getSheetByName(name), last = sh.getLastRow();
      if (last < 2) return;
      var v = sh.getRange(2, 1, last - 1, 1).getValues();
      for (var i = v.length - 1; i >= 0; i--) if (set[String(v[i][0])]) sh.deleteRow(i + 2);
    });
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function requireMgmt_(token) {
  if (!token || !CacheService.getScriptCache().get('mt_' + token)) {
    var e = new Error('AUTH'); throw e;
  }
}

/* ================================== HELPERS ================================== */

function tz_() { return Session.getScriptTimeZone(); }

/** '98765 43210, +91 99946 79999' -> ['919876543210','919994679999'] (India default, digits only, for wa.me links). */
function phones_(v) {
  return String(v === null || v === undefined ? '' : v).split(/[,;\n]+/).map(function (x) {
    var d = x.replace(/\D/g, '');
    if (d.length === 10) d = '91' + d;
    else if (d.length === 11 && d.charAt(0) === '0') d = '91' + d.slice(1);
    return d;
  }).filter(function (d) { return d.length >= 11 && d.length <= 15; });
}

function readConfig_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET.CONFIG);
  var c = { statuses: DEFAULT_STATUSES, closedStatuses: DEFAULT_CLOSED, usdInr: 93, reqStatuses: DEFAULT_REQ, techpackShare: 'ANYONE', fabricTeamEmail: '', trimsTeamEmail: '', merchTeamEmail: '', fabricTeamMobile: '', trimsTeamMobile: '', merchTeamMobile: '' };
  if (sh) sh.getDataRange().getValues().slice(1).forEach(function (r) { if (r[0]) c[String(r[0])] = r[1]; });
  var split = function (s) { return String(s).split(',').map(function (x) { return x.trim().toUpperCase(); }).filter(String); };
  var rs = split(c.reqStatuses);
  ['SUBMITTED', 'ACCEPTED'].forEach(function (x) {
    if (rs.length && rs.indexOf(x) < 0) { var i = rs.indexOf('CANCELLED'); if (i < 0) rs.push(x); else rs.splice(i, 0, x); }
  });
  return {
    statuses: split(c.statuses), closed: split(c.closedStatuses), usdInr: Number(c.usdInr) || 93,
    reqStatuses: rs.length ? rs : split(DEFAULT_REQ),
    techShare: String(c.techpackShare || 'ANYONE').trim().toUpperCase(),
    fabricEmail: String(c.fabricTeamEmail || '').trim(), trimsEmail: String(c.trimsTeamEmail || '').trim(), merchEmail: String(c.merchTeamEmail || '').trim(),
    mobiles: { FABRIC: phones_(c.fabricTeamMobile), TRIMS: phones_(c.trimsTeamMobile), MERCH: phones_(c.merchTeamMobile) }
  };
}

function readOrders_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET.ORDERS);
  if (!sh) throw new Error('Run Jiwin OIH > Setup sheets first.');
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, OCOLS.length).getValues()
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) { return fmtOut_(rowToObj_(r)); });
}

function readImages_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET.IMAGES);
  var out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) { if (r[0] && r[1]) out[String(r[0])] = String(r[1]); });
  return out;
}

function rowToObj_(r) {
  var o = {};
  OCOLS.forEach(function (k, i) { o[k] = r[i]; });
  return o;
}

function objToRow_(o) {
  return OCOLS.map(function (k) {
    var v = o[k];
    if (DATE_FIELDS.indexOf(k) > -1) return dateIn_(v);
    if (NUM_FIELDS.indexOf(k) > -1) return (v === '' || v === null || v === undefined) ? '' : Number(v);
    return v === null || v === undefined ? '' : v;
  });
}

// Sheet values -> plain JSON-safe strings/numbers for the browser
function fmtOut_(o) {
  var out = {};
  OCOLS.forEach(function (k) {
    var v = o[k];
    if (DATE_FIELDS.indexOf(k) > -1) out[k] = dateOut_(v);
    else if (NUM_FIELDS.indexOf(k) > -1) out[k] = (v === '' || v === null || v === undefined) ? '' : Number(v);
    else if (k === 'updatedAt') out[k] = (v instanceof Date) ? Utilities.formatDate(v, tz_(), 'yyyy-MM-dd HH:mm') : String(v || '');
    else if (k === 'kind') out[k] = String(v).toUpperCase() === 'SAMPLE' ? 'SAMPLE' : 'ORDER';
    else out[k] = String(v === null || v === undefined ? '' : v);
  });
  return out;
}

function applyFields_(cur, it) {
  TEXT_FIELDS.forEach(function (k) {
    if (it.hasOwnProperty(k)) {
      cur[k] = clean_(it[k], k === 'fabric' || k === 'remarks' ? 200 : 80);
      if (k === 'buyer' || k === 'status') cur[k] = cur[k].toUpperCase();
      if (k === 'kind') cur[k] = cur[k].toUpperCase() === 'SAMPLE' ? 'SAMPLE' : 'ORDER';
    }
  });
  DATE_FIELDS.forEach(function (k) {
    if (it.hasOwnProperty(k)) {
      var s = String(it[k] || '').trim();
      if (s && !/^\d{4}-\d{2}-\d{2}$/.test(s)) throw new Error('Bad date for ' + k);
      cur[k] = s;
    }
  });
  NUM_FIELDS.forEach(function (k) {
    if (it.hasOwnProperty(k)) {
      var v = it[k];
      if (v === '' || v === null) { cur[k] = ''; return; }
      v = Number(v);
      if (!isFinite(v) || v < 0 || v > 100000000) throw new Error('Bad number for ' + k);
      cur[k] = Math.round(v);
    }
  });
}

function clean_(s, max) { return String(s === null || s === undefined ? '' : s).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max); }

function dateIn_(v) {
  if (!v) return '';
  if (v instanceof Date) return v;
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(v));
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0) : '';
}

function dateOut_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  return /^\d{4}-\d{2}-\d{2}$/.test(String(v)) ? String(v) : '';
}

function numOrBlank_(v) {
  if (v === '' || v === null || v === undefined) return '';
  var n = Number(v);
  return isFinite(n) ? n : '';
}

/* ================================= SEED LOAD ================================= */

function loadSeedData() {
  var ui = SpreadsheetApp.getUi();
  if (typeof SEED_ORDERS === 'undefined') { ui.alert('Seed.gs not found (already deleted?).'); return; }
  setupSheets();
  var ss = SpreadsheetApp.getActive();
  var o = ss.getSheetByName(SHEET.ORDERS);
  if (o.getLastRow() > 1) { ui.alert('ORDERS already has data. Seed load cancelled.'); return; }
  var now = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm');
  var orows = [], prow = [], irow = [];
  SEED_ORDERS.forEach(function (s, i) {
    var id = 'S' + Utilities.getUuid().slice(0, 7);
    var obj = {};
    OCOLS.forEach(function (k) { obj[k] = ''; });
    obj.id = id; obj.status = s[0]; obj.jo = s[1]; obj.ship = s[2]; obj.buyer = String(s[3]).toUpperCase();
    obj.style = s[4]; obj.styling = s[5]; obj.fabric = s[6]; obj.color = s[7]; obj.qty = s[8];
    obj.updatedBy = 'seed'; obj.updatedAt = now;
    orows.push(objToRow_(obj));
    var p = SEED_PRICING[i];
    if (p && (p[0] !== '' || p[1] !== '' || p[2] !== '')) prow.push([id, p[0], p[1], p[2]]);
    if (SEED_IMAGES[i]) irow.push([id, SEED_IMAGES[i], new Date()]);
  });
  o.getRange(2, 1, orows.length, OCOLS.length).setValues(orows);
  if (prow.length) ss.getSheetByName(SHEET.PRICING).getRange(2, 1, prow.length, 4).setValues(prow);
  if (irow.length) ss.getSheetByName(SHEET.IMAGES).getRange(2, 1, irow.length, 3).setValues(irow);
  ui.alert('Loaded ' + orows.length + ' order lines, ' + prow.length + ' price rows, ' + irow.length + ' images.\n\nNow DELETE Seed.gs (it contains prices).');
}

/*************** ONE-TIME OIH DATA MIGRATION ***************/
/**
 * Copies every row from the original standalone OIH Google Sheet into the
 * matching tabs here (created fresh by setupSheets() if not already
 * present). Safe to run more than once — a tab already holding data is
 * left untouched and reported as skipped, rather than duplicating rows.
 * Run once from Sheet menu: Jiwin OIH > 1. Migrate OIH data (one time).
 */
var OIH_SOURCE_SHEET_ID = '1Bt1kr3ejk8gUTRF5Wdf3NwgfKTfJfhLO3qBj5UhtZC8';

function migrateOihData() {
  setupSheets();
  var src;
  try {
    src = SpreadsheetApp.openById(OIH_SOURCE_SHEET_ID);
  } catch (e) {
    var msg = 'Could not open the source OIH sheet (' + OIH_SOURCE_SHEET_ID + '): ' + e.message;
    SpreadsheetApp.getUi().alert(msg);
    return msg;
  }
  var dst = SpreadsheetApp.getActiveSpreadsheet();
  var names = [SHEET.ORDERS, SHEET.PRICING, SHEET.IMAGES, SHEET.CONFIG, SHEET.TECHPACKS, SHEET.REQS];
  var report = [];
  names.forEach(function (name) {
    var srcSh = src.getSheetByName(name);
    if (!srcSh) { report.push(name + ': not found in source sheet, skipped.'); return; }
    var dstSh = dst.getSheetByName(name);
    if (!dstSh) { report.push(name + ': destination tab missing (setupSheets should have created it) — skipped.'); return; }
    var lastRow = srcSh.getLastRow(), lastCol = srcSh.getLastColumn();
    if (lastRow < 2) { report.push(name + ': no data rows in source.'); return; }
    if (dstSh.getLastRow() > 1) { report.push(name + ': destination already has data — skipped (this tab was already migrated, or has new entries; delete its rows first to re-migrate from scratch).'); return; }
    var values = srcSh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    dstSh.getRange(2, 1, values.length, lastCol).setValues(values);
    report.push(name + ': migrated ' + values.length + ' row(s).');
  });
  repairIds_(dst); // fix any sci-notation-mangled IDs carried over from the source
  SpreadsheetApp.flush();
  var summary = 'OIH data migration:\n\n' + report.join('\n');
  SpreadsheetApp.getUi().alert(summary);
  return summary;
}

/*************** COMBINED DASHBOARD (OIH + FABRIC ERP, joined by JO Number) ***************/
/**
 * For every open OIH order line, attaches its Fabric ERP status (required
 * kg, ordered, received, current process stage) by matching JO Number
 * against Style Tracker. An OIH order with no matching Style Tracker entry
 * yet still shows (fabric side blank) rather than being dropped — the two
 * apps' data entry doesn't have to happen in lockstep.
 */
function getCombinedDashboardData() {
  var oihOrders = readOrders_().filter(function (o) {
    var cfg = readConfig_();
    return cfg.closed.indexOf(o.status) === -1 && String(o.kind).toUpperCase() !== 'SAMPLE';
  });
  var styleTracker = getStyleTracker();
  var stByJO = {};
  styleTracker.forEach(function (r) {
    var jo = String(r['JO Number'] || '').trim();
    if (!jo) return;
    if (!stByJO[jo]) stByJO[jo] = { requiredKg: 0, rows: 0 };
    stByJO[jo].requiredKg += Number(r['Fabric Required (kg)']) || 0;
    stByJO[jo].rows++;
  });
  var allDCs = sheetToObjects(SHEETS.PROCESS_DC);
  var lastDcByJO = {};
  allDCs.forEach(function (dc) {
    var jo = String(dc['JO Number'] || '').trim();
    if (!jo) return;
    if (!lastDcByJO[jo] || String(dc['Date']) >= String(lastDcByJO[jo]['Date'])) lastDcByJO[jo] = dc;
  });
  var yarnUsageByJO = {};
  sheetToObjects(SHEETS.YARN_PO).forEach(function (po) {
    if (po['Status'] === 'Rejected') return;
    var jo = String(po['JO Number'] || '').trim();
    yarnUsageByJO[jo] = (yarnUsageByJO[jo] || 0) + (Number(po['Qty KG']) || 0);
  });

  var rows = oihOrders.map(function (o) {
    var jo = String(o.jo || '').trim();
    var st = stByJO[jo];
    var lastDc = lastDcByJO[jo];
    return {
      jo: jo, buyer: o.buyer, style: o.style, styling: o.styling, color: o.color, qty: o.qty,
      ship: o.ship, status: o.status,
      fabricRequiredKg: st ? Math.round(st.requiredKg * 100) / 100 : null,
      yarnOrderedKg: yarnUsageByJO[jo] ? Math.round(yarnUsageByJO[jo] * 100) / 100 : 0,
      currentProcess: lastDc ? lastDc['To Process'] : (st ? 'Not started' : ''),
      currentSupplier: lastDc ? lastDc['To Supplier'] : '',
      inStyleTracker: !!st
    };
  });
  return { rows: rows, generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') };
}

/*************** JO CHECKLIST / TASKS ***************/
var TASK_SHEET = 'TASKS';
var TASK_TEMPLATE_SHEET = 'TASK_TEMPLATES';
var TASK_HEAD = ['Task ID', 'JO Number', 'Buyer', 'Department', 'Task Text', 'Assigned To', 'Due Date', 'Status', 'Completed By', 'Completed At', 'Created By', 'Created At'];
var TASK_TEMPLATE_HEAD = ['Template ID', 'Buyer', 'Department', 'Task Text'];
var TASK_DEPARTMENTS = ['Merch', 'Production', 'Store', 'Purchase', 'Fabric'];

function ensureTaskSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, TASK_SHEET, TASK_HEAD);
  ensureSheet_(ss, TASK_TEMPLATE_SHEET, TASK_TEMPLATE_HEAD);
}

function taskRowToObj_(r) {
  var o = {}, tz = Session.getScriptTimeZone();
  TASK_HEAD.forEach(function (h, i) {
    var v = r[i];
    // Dates can't travel to the browser (google.script.run returns null) - send text.
    if (v instanceof Date) v = Utilities.formatDate(v, tz, /At$/.test(h) ? 'yyyy-MM-dd HH:mm' : 'yyyy-MM-dd');
    o[h] = v;
  });
  return o;
}

function getAllTasks_() {
  ensureTaskSheets_();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TASK_SHEET);
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, TASK_HEAD.length).getValues()
    .filter(function (r) { return r[0] !== ''; })
    .map(taskRowToObj_);
}

/** Buyer's saved template tasks, so creating a new JO's checklist can start from what usually applies to that buyer. */
function getTaskTemplatesForBuyer(buyer) {
  ensureTaskSheets_();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TASK_TEMPLATE_SHEET);
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, TASK_TEMPLATE_HEAD.length).getValues()
    .filter(function (r) { return r[0] !== '' && String(r[1]).trim().toUpperCase() === String(buyer).trim().toUpperCase(); })
    .map(function (r) { return { id: r[0], buyer: r[1], department: r[2], taskText: r[3] }; });
}

/** Saves (or updates) one buyer template task — freely typed, not a fixed list. */
function saveTaskTemplate(buyer, department, taskText, templateId) {
  ensureTaskSheets_();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TASK_TEMPLATE_SHEET);
  if (templateId) {
    var idx = findRowIndex(TASK_TEMPLATE_SHEET, 'Template ID', templateId);
    if (idx !== -1) { setCell(TASK_TEMPLATE_SHEET, idx, 'Buyer', buyer); setCell(TASK_TEMPLATE_SHEET, idx, 'Department', department); setCell(TASK_TEMPLATE_SHEET, idx, 'Task Text', taskText); return { success: true }; }
  }
  var id = 'TPL-' + Utilities.getUuid().slice(0, 8);
  sh.appendRow([id, buyer, department, taskText]);
  return { success: true, id: id };
}

function deleteTaskTemplate(templateId) {
  var idx = findRowIndex(TASK_TEMPLATE_SHEET, 'Template ID', templateId);
  if (idx === -1) return { success: false, error: 'Template not found.' };
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TASK_TEMPLATE_SHEET).deleteRow(idx);
  return { success: true };
}

/**
 * Creates the checklist for a JO — one row per task — and emails the
 * techpack/spec sheet/order details + task list to the 5 departments the
 * moment it's created (Merch, Production, Store, Purchase, Fabric).
 * tasks: [{department, taskText, assignedTo, dueDate}]
 */
function createChecklistForJO(joNumber, buyer, tasks, createdBy) {
  ensureTaskSheets_();
  if (!tasks || !tasks.length) return { success: false, error: 'Add at least one task.' };
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TASK_SHEET);
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var rows = tasks.map(function (t) {
    var id = 'TSK-' + Utilities.getUuid().slice(0, 8);
    return [id, joNumber, buyer, t.department, t.taskText, t.assignedTo || '', t.dueDate || '', 'Pending', '', '', createdBy || '', now];
  });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, TASK_HEAD.length).setValues(rows);

  // Notify the 5 departments — techpack/spec sheet/order details + the tasks assigned to each.
  var cfg = readConfig_();
  var deptEmail = {
    'Merch': cfg.merchEmail, 'Fabric': cfg.fabricEmail, 'Purchase': cfg.trimsEmail,
    'Production': String(readConfigRaw_('productionTeamEmail') || ''), 'Store': String(readConfigRaw_('storeTeamEmail') || '')
  };
  var byDept = {};
  tasks.forEach(function (t) { (byDept[t.department] = byDept[t.department] || []).push(t); });
  var jo = getJOInfo(joNumber);
  var url = ''; try { url = ScriptApp.getService().getUrl(); } catch (e) { }
  var sent = [];
  Object.keys(byDept).forEach(function (dept) {
    var to = deptEmail[dept];
    if (!to) return;
    var body = 'Hello ' + dept + ' team,\n\nA new JO has been created: ' + joNumber + (buyer ? ' (' + buyer + ')' : '') +
      (jo ? '\nStyle: ' + (jo['Style Number'] || '') + ' | Color: ' + (jo['Color'] || '') : '') +
      '\n\nYour checklist for this JO:\n' + byDept[dept].map(function (t) { return '- ' + t.taskText + (t.assignedTo ? ' (assigned: ' + t.assignedTo + ')' : '') + (t.dueDate ? ' | due ' + t.dueDate : ''); }).join('\n') +
      '\n\n' + (url ? 'Open the ERP: ' + url + '\n\n' : '') + 'Regards,\nJiwin Apparels ERP';
    try { MailApp.sendEmail({ to: to, subject: 'New JO ' + joNumber + ' — ' + dept + ' checklist', body: body, name: 'Jiwin ERP' }); sent.push(dept); } catch (e) { }
  });
  return { success: true, message: 'Checklist created (' + rows.length + ' task(s)). Notified: ' + (sent.length ? sent.join(', ') : 'no department emails configured yet') + '.' };
}

function readConfigRaw_(key) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.CONFIG);
  if (!sh) return '';
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (v[i][0] === key) return v[i][1];
  return '';
}

function updateTask(taskId, fields) {
  ensureTaskSheets_();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TASK_SHEET);
  var last = sh.getLastRow();
  if (last < 2) return { success: false, error: 'Task not found.' };
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  var row = -1;
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === String(taskId).trim()) { row = i + 2; break; }
  if (row === -1) return { success: false, error: 'Task not found.' };
  Object.keys(fields).forEach(function (k) {
    var c = TASK_HEAD.indexOf(k);
    if (c !== -1) sh.getRange(row, c + 1).setValue(fields[k]);
  });
  return { success: true };
}

/*************** TASKS AGAINST A JO (Fabric ERP dashboard) ***************/
function getJoTaskList(joNumber) {
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var list = getTasksForJO(joNumber).map(function (t) {
    var due = String(t['Due Date'] || '').slice(0, 10);
    t.overdue = t['Status'] !== 'Completed' && !!due && due < today;
    return t;
  });
  list.sort(function (a, b) {
    var ca = a['Status'] === 'Completed' ? 1 : 0, cb = b['Status'] === 'Completed' ? 1 : 0;
    if (ca !== cb) return ca - cb;
    var da = String(a['Due Date'] || '9999'), db = String(b['Due Date'] || '9999');
    return da < db ? -1 : da > db ? 1 : 0;
  });
  return { tasks: list, departments: TASK_DEPARTMENTS };
}

/* task = { department, taskText, assignedTo, dueDate } */
function addJoTask(joNumber, task, createdBy) {
  if (!joNumber) return { success: false, error: 'Select a JO first.' };
  if (!task || !String(task.taskText || '').trim()) return { success: false, error: 'Enter the task.' };
  var jo = getJOInfo(joNumber);
  var buyer = jo ? (jo['Buyer Name'] || '') : '';
  return createChecklistForJO(joNumber, buyer, [{
    department: task.department || 'Merch', taskText: String(task.taskText).trim(),
    assignedTo: task.assignedTo || '', dueDate: task.dueDate || ''
  }], createdBy || '');
}

/* Adds every standard task saved for this JO's buyer (Task Templates) */
function addBuyerTemplateTasks(joNumber, createdBy) {
  var jo = getJOInfo(joNumber);
  if (!jo) return { success: false, error: 'JO not found in Style Tracker.' };
  var buyer = jo['Buyer Name'] || '';
  var tpl = getTaskTemplatesForBuyer(buyer) || [];
  if (!tpl.length) return { success: false, error: 'No standard tasks saved for buyer ' + buyer + '.' };
  var have = {};
  getTasksForJO(joNumber).forEach(function (t) { have[(t['Department'] + '|' + t['Task Text']).toLowerCase()] = 1; });
  var tasks = tpl.map(function (t) {
    return { department: t['Department'] || t.department || 'Merch', taskText: t['Task Text'] || t.taskText || '' };
  }).filter(function (t) { return t.taskText && !have[(t.department + '|' + t.taskText).toLowerCase()]; });
  if (!tasks.length) return { success: false, error: 'All standard tasks for ' + buyer + ' are already on this JO.' };
  return createChecklistForJO(joNumber, buyer, tasks, createdBy || '');
}

function deleteJoTask(taskId) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TASK_SHEET);
  if (!sh || sh.getLastRow() < 2) return { success: false, error: 'Task not found.' };
  var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(taskId).trim()) { sh.deleteRow(i + 2); return { success: true }; }
  }
  return { success: false, error: 'Task not found.' };
}

function completeTask(taskId, completedBy) {
  return updateTask(taskId, {
    'Status': 'Completed', 'Completed By': completedBy || '',
    'Completed At': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
  });
}

function reopenTask(taskId) {
  return updateTask(taskId, { 'Status': 'Pending', 'Completed By': '', 'Completed At': '' });
}

function getTasksForJO(joNumber) {
  return getAllTasks_().filter(function (t) { return String(t['JO Number']).trim() === String(joNumber).trim(); });
}

/**
 * Every task, bucketed: Overdue (due date passed, not done), Pending (due
 * today), Upcoming (due later this week), Completed. A task with no due
 * date at all falls into Pending until one is set.
 */
function getTasksBucketed(filterDept, filterPerson) {
  var all = getAllTasks_().filter(function (t) {
    if (filterDept && t['Department'] !== filterDept) return false;
    if (filterPerson && String(t['Assigned To'] || '').trim().toLowerCase() !== String(filterPerson).trim().toLowerCase()) return false;
    return true;
  });
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var buckets = { overdue: [], pending: [], upcoming: [], completed: [] };
  all.forEach(function (t) {
    if (t['Status'] === 'Completed') { buckets.completed.push(t); return; }
    var due = String(t['Due Date'] || '').slice(0, 10);
    if (due && due < today) buckets.overdue.push(t);
    else if (!due || due === today) buckets.pending.push(t);
    else buckets.upcoming.push(t); // this week or later — all treated as upcoming until it becomes today
  });
  return buckets;
}

/*************** IMPORT PO — AUTO-FILL STYLE TRACKER VIA AI EXTRACTION ***************/
/**
 * Converts an uploaded Excel file to CSV text server-side (via a temporary
 * Google Sheet) so it can be sent to Claude as plain text — the Messages
 * API doesn't accept raw .xlsx binary. Requires the Advanced Drive Service
 * to be enabled once (Apps Script editor > Services > + > Drive API).
 */
function convertExcelToCsv_(bytes, fileName) {
  var blob = Utilities.newBlob(bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', fileName);
  var file;
  try {
    file = Drive.Files.create({ name: fileName, mimeType: MimeType.GOOGLE_SHEETS }, blob, { fields: 'id' });
  } catch (e) {
    throw new Error('Could not convert the Excel file. The Advanced Drive Service may not be enabled for this project — in the Apps Script editor, click Services (+), add "Drive API", then try again. Original error: ' + e.message);
  }
  var fileId = file.id;
  var ss = SpreadsheetApp.openById(fileId);
  var allText = ss.getSheets().map(function (sh) {
    var vals = sh.getDataRange().getValues();
    return vals.map(function (row) { return row.join(','); }).join('\n');
  }).join('\n\n');
  DriveApp.getFileById(fileId).setTrashed(true);
  return allText;
}

function buildStyleTrackerExtractionPrompt_() {
  return 'You are reading a buyer\'s Purchase Order / techpack / spec sheet for a knitwear garment manufacturer. ' +
    'Extract every distinct Part+Color line item you find (e.g. a style with BODY, RIB and NECK TAPE parts, each color, is one row per Part+Color combination) into a JSON array. ' +
    'Each object in the array must use exactly these keys (leave a key as an empty string "" if genuinely not found in the document — never omit a key):\n\n' +
    '"JO Number","Style Number","Buyer Name","Style Description","Part","Color","Order Quantity","Excess %",' +
    '"Size 1","Qty 1","Size 2","Qty 2","Size 3","Qty 3","Size 4","Qty 4","Size 5","Qty 5","Size 6","Qty 6","Size 7","Qty 7","Size 8","Qty 8","Size 9","Qty 9","Size 10","Qty 10",' +
    '"CAD Weight (g)","Loss %","Counts","Yarn Type","Content","Has Lycra","Lycra %","Fabric type","GSM",' +
    '"Knitting Dia","Finishing dia","Loop length","Gauge","Yarn color","Yarn rate",' +
    '"Yarn dyeing rate","Yarn dyeing loss %","Cone winding rate","Cone winding loss %","Knitting rate","Knitting loss %",' +
    '"Heatsetting rate","Heatsetting loss %","Dyeing rate","Dyeing loss %","Tubular compacting rate","Tubular compacting loss %",' +
    '"Stener rate","Stenter loss %","Compacting rate","Compacting loss %","Rasing /Peach fnish","Raising loss %",' +
    '"Bio washing rate","Bio washing loss %","Jiwin Inward loss %","Order/OCN No","Buyer Order No","Despatch Date","Season","Merchandiser"\n\n' +
    '"Has Lycra" must be exactly "Yes" or "No". Numbers should be plain numbers (no units, no commas) as strings or numbers, not formulas. ' +
    'Rate and loss-% fields for processes the document doesn\'t mention should stay "". ' +
    'RULES FOR MULTI-ORDER POs (e.g. LPP / Sinsay): one PDF can hold several order "entities" (separate order numbers, E-COMM orders) for the SAME model and colour. ' +
    'Add their size-wise quantities together into ONE row per Part+Colour (so Qty per size = sum over all entities, Order Quantity = grand total). ' +
    'Put every order number in "Buyer Order No" separated by " / ", and use the EARLIEST shipment date as "Despatch Date" (format YYYY-MM-DD). ' +
    'Use the Model No. as "Style Number", the VPN / model name as "Style Description", the brand or buying company as "Buyer Name", ' +
    'the colour code with its name as "Color" (e.g. "04X LAVENDER"), and Fabric Composition as "Content". Sizes are the size labels in order (e.g. 116, 122, 128, 134, 140). ' +
    'If no garment parts are listed, use Part "BODY". Ignore terms & conditions, packing, carton and chemical pages. ' +
    'Respond with ONLY the JSON array — no prose before or after it, no markdown code fences.';
}

/** Sizewise Total, Size Check, Fabric Required (kg), Main Fiber Reqd, Lycra Reqd, Excess Quantity — verified against real Style Tracker rows (e.g. JO-69 RIB: (15120+453.6)*0.006*1.06 = 99; JO-72 BODY: (30016+750.4)*0.152*1.08 = 5051). */
function computeStyleTrackerDerived_(r) {
  var sizewiseTotal = 0;
  for (var i = 1; i <= 10; i++) { sizewiseTotal += Number(r['Qty ' + i]) || 0; }
  r['Sizewise Total'] = sizewiseTotal;
  var orderQty = Number(r['Order Quantity']) || 0;
  r['Size Check'] = (sizewiseTotal === orderQty) ? 'OK' : ('Mismatch: sizes total ' + sizewiseTotal + ' vs order qty ' + orderQty);
  var excessQty = Math.round(orderQty * (Number(r['Excess %']) || 0) / 100 * 100) / 100;
  r['Excess Quantity'] = excessQty;
  var cadWeightKg = (Number(r['CAD Weight (g)']) || 0); // stored as kg-per-piece despite the "(g)" label, matching existing Style Tracker rows
  var lossPct = Number(r['Loss %']) || 0;
  // Fabric Required is based on Order Quantity + Excess Quantity (the buffer already built into the order), not the raw sizewise total.
  var fabricRequired = Math.round((orderQty + excessQty) * cadWeightKg * (1 + lossPct / 100) * 100) / 100;
  r['Fabric Required (kg)'] = fabricRequired;
  if (String(r['Has Lycra']).trim().toLowerCase() === 'yes') {
    var lycraPct = Number(r['Lycra %']) || 0;
    r['Lycra Reqd (kg)'] = Math.round(fabricRequired * lycraPct / 100 * 100) / 100;
    r['Main Fiber Reqd (kg)'] = Math.round((fabricRequired - r['Lycra Reqd (kg)']) * 100) / 100;
  } else {
    r['Main Fiber Reqd (kg)'] = fabricRequired;
    r['Lycra Reqd (kg)'] = 0;
  }
  r['No Of Parts'] = r['No Of Parts'] || '';
  return r;
}

/**
 * fileName/mimeType/base64Data come from the browser's file picker.
 * Routes to the right Claude content type per format and returns the
 * extracted rows (with derived fields already computed) for review before
 * anything is saved to Style Tracker.
 */
function extractStyleTrackerFromUpload(fileName, mimeType, base64Data) {
  var props = PropertiesService.getScriptProperties();
  var gKey = props.getProperty('GEMINI_API_KEY'), aKey = props.getProperty('ANTHROPIC_API_KEY');
  var provider = props.getProperty('AI_PROVIDER') || (gKey ? 'GEMINI' : (aKey ? 'ANTHROPIC' : ''));
  if (!provider || (provider === 'GEMINI' && !gKey) || (provider === 'ANTHROPIC' && !aKey)) {
    return { success: false, needKey: true, error: 'PO import needs an AI key (one time).' };
  }
  var promptText = buildStyleTrackerExtractionPrompt_();

  // what to send: PDF / image as a file, Excel (already turned into CSV text in the browser) as text
  var kind, textData = '';
  if (/csv|text\/plain/i.test(mimeType) || /\.csv$/i.test(fileName)) {
    kind = 'TEXT'; textData = Utilities.newBlob(Utilities.base64Decode(base64Data)).getDataAsString().slice(0, 60000);
  } else if (/spreadsheet|sheet|excel/i.test(mimeType) || /\.xlsx?$/i.test(fileName)) {
    try { kind = 'TEXT'; textData = convertExcelToCsv_(Utilities.base64Decode(base64Data), fileName).slice(0, 60000); }
    catch (e) { return { success: false, error: e.message }; }
  } else if (/pdf/i.test(mimeType) || /\.pdf$/i.test(fileName)) { kind = 'PDF'; mimeType = 'application/pdf'; }
  else if (/image/i.test(mimeType)) { kind = 'IMAGE'; }
  else return { success: false, error: 'Unsupported file type (' + mimeType + '). Upload a PDF, Excel/CSV, or a photo (JPG/PNG).' };

  var text = '';
  try {
    if (provider === 'GEMINI') {
      var parts = [{ text: promptText + (kind === 'TEXT' ? '\n\nHere is the PO spreadsheet as CSV:\n\n' + textData : '') }];
      if (kind !== 'TEXT') parts.push({ inline_data: { mime_type: mimeType, data: base64Data } });
      var gr = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + encodeURIComponent(gKey), {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        payload: JSON.stringify({ contents: [{ role: 'user', parts: parts }], generationConfig: { temperature: 0, responseMimeType: 'application/json' } })
      });
      if (gr.getResponseCode() !== 200) return { success: false, error: 'Gemini error (' + gr.getResponseCode() + '): ' + gr.getContentText().slice(0, 400) };
      var gd = JSON.parse(gr.getContentText());
      text = ((((gd.candidates || [])[0] || {}).content || {}).parts || []).map(function (p) { return p.text || ''; }).join('');
    } else {
      var userContent = [{ type: 'text', text: promptText + (kind === 'TEXT' ? '\n\nHere is the PO spreadsheet as CSV:\n\n' + textData : '') }];
      if (kind === 'PDF') userContent.push({ type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: base64Data } });
      if (kind === 'IMAGE') userContent.push({ type: 'image', source: { type: 'base64', media_type: mimeType, data: base64Data } });
      var ar = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { 'x-api-key': aKey, 'anthropic-version': '2023-06-01' },
        payload: JSON.stringify({ model: 'claude-sonnet-4-5', max_tokens: 8000, messages: [{ role: 'user', content: userContent }] })
      });
      if (ar.getResponseCode() !== 200) return { success: false, error: 'Anthropic API error (' + ar.getResponseCode() + '): ' + ar.getContentText().slice(0, 400) };
      var ad = JSON.parse(ar.getContentText());
      text = (ad.content || []).filter(function (b) { return b.type === 'text'; }).map(function (b) { return b.text; }).join('');
    }
  } catch (e) { return { success: false, error: 'Could not reach the AI service: ' + e.message }; }

  var jsonMatch = text.match(/\[[\s\S]*\]/);
  if (!jsonMatch) return { success: false, error: 'Could not read line items from this file. Try a clearer PDF/scan. Reply: ' + text.slice(0, 300) };
  var rows;
  try { rows = JSON.parse(jsonMatch[0]); } catch (e) { return { success: false, error: 'The AI reply was not valid JSON: ' + e.message }; }
  if (!rows.length) return { success: false, error: 'No lines were found in this document.' };
  rows.forEach(computeStyleTrackerDerived_);
  return { success: true, rows: rows };
}

/** Writes reviewed/corrected rows to Style Tracker — re-derives the computed fields at save time too, in case the person edited a raw input (qty, CAD weight, loss %, lycra %) during review. */
function saveStyleTrackerRows(rows) {
  if (!rows || !rows.length) return { success: false, error: 'Nothing to save.' };
  var sh = SS.getSheetByName(SHEETS.STYLE_TRACKER);
  var head = stHead_(sh);
  var out = rows.map(function (r) {
    computeStyleTrackerDerived_(r);
    // match by column NAME (the sheet has extra columns after Merchandiser, so positions differ from HEADERS)
    return head.map(function (h) { return h && r[h] !== undefined ? r[h] : ''; });
  });
  sh.getRange(sh.getLastRow() + 1, 1, out.length, head.length).setValues(out);
  return { success: true, message: out.length + ' row(s) saved to Style Tracker.' };
}

/*************** STYLE IMAGE (stored against the JO in Style Tracker) ***************
 * The picture file is saved in Drive (folder JIWIN_Style_Images) and its link is
 * written into a "Style Image" column on every Style Tracker row of that JO.
 * You can also paste any Drive link / image URL into that column by hand.
 ***********************************************************************************/
var STYLE_IMAGE_COL = 'Style Image';

function stImgFolder_() {
  var folders = DriveApp.getFoldersByName('JIWIN_Style_Images');
  return folders.hasNext() ? folders.next() : DriveApp.createFolder('JIWIN_Style_Images');
}

/* 1-based column of "Style Image" in Style Tracker; adds it after the last column if missing */
function stImgCol_(sh, create) {
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h || '').trim(); });
  var idx = head.indexOf(STYLE_IMAGE_COL);
  if (idx > -1) return idx + 1;
  if (!create) return 0;
  var col = lastCol + 1;
  if (sh.getMaxColumns() < col) sh.insertColumnsAfter(sh.getMaxColumns(), col - sh.getMaxColumns());
  sh.getRange(1, col).setValue(STYLE_IMAGE_COL).setFontWeight('bold').setBackground('#1B2A4A').setFontColor('#FFFFFF');
  return col;
}

function stImgWrite_(joNumber, value) {
  var sh = SS.getSheetByName(SHEETS.STYLE_TRACKER);
  var col = stImgCol_(sh, true);
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var jos = sh.getRange(2, 1, last - 1, 1).getValues();
  var n = 0;
  jos.forEach(function (r, i) {
    if (String(r[0]).trim() === String(joNumber).trim()) { sh.getRange(i + 2, col).setValue(value); n++; }
  });
  try { CacheService.getScriptCache().remove('oihth:' + String(joNumber).trim()); } catch (e) { }   // picture changed -> drop the cached thumbnail
  return n;
}

/* Small pictures for the Order In Hand list, one per JO (cached 6 h so the page stays fast) */
function oihThumbFor_(jo, link) {
  var cache = CacheService.getScriptCache(), key = 'oihth:' + jo, hit = null;
  try { hit = cache.get(key); } catch (e) { }
  if (hit) return hit;
  var d = stImgThumbDataUrl_(link);
  if (d && d.length < 90000) { try { cache.put(key, d, /^cell:/.test(link) ? 3600 : 21600); } catch (e2) { } }
  return d;
}
function getOihThumbs() {
  var out = {}, seen = {};
  getStyleTracker().forEach(function (r) {
    var jo = String(r['JO Number'] || '').trim();
    if (!jo || seen[jo]) return;
    var link = stImgLinkForJo_(jo, [r]);
    if (!link) return;
    seen[jo] = 1;
    var d = oihThumbFor_(jo, link);
    if (d) out[jo] = d;
  });
  return out;
}

/* Turn a stored link into something an <img> can show (Drive files are sent as data, no sharing needed) */
/* Download a picture that was pasted into a Style Tracker cell (Insert > Image > Image in cell) */
function stCellImageData_(row, maxPx) {
  try {
    var sh = SS.getSheetByName(SHEETS.STYLE_TRACKER), col = stImgCol_(sh, false);
    if (!col) { STIMG_ERR_ = 'the Style Image column was not found'; return ''; }
    var v = sh.getRange(row, col).getValue();
    if (!isCellImage_(v)) { STIMG_ERR_ = 'the Style Image cell (row ' + row + ') no longer holds a picture'; return ''; }
    var url = ''; try { url = v.getContentUrl(); } catch (e0) { }
    if (!url) { try { url = v.getUrl(); } catch (e1) { } }
    if (!url) { STIMG_ERR_ = 'the picture in the cell has no downloadable address'; return ''; }
    var tok = ScriptApp.getOAuthToken(), tries = [];
    if (maxPx && /googleusercontent\.com/.test(url)) tries.push(url + '=w' + maxPx);   // let Google shrink it
    tries.push(url);
    var last = 0;
    for (var i = 0; i < tries.length; i++) {
      var r = UrlFetchApp.fetch(tries[i], { headers: { Authorization: 'Bearer ' + tok }, muteHttpExceptions: true, followRedirects: true });
      last = r.getResponseCode();
      if (last !== 200) continue;
      var b = r.getBlob(), bytes = b.getBytes(), ct = String(b.getContentType() || '');
      if (!/^image\//.test(ct)) { if (imgDims_(bytes)) ct = 'image/png'; else continue; }
      return 'data:' + ct + ';base64,' + Utilities.base64Encode(bytes);
    }
    STIMG_ERR_ = 'the picture in the cell could not be downloaded (HTTP ' + last + ')';
  } catch (e) { STIMG_ERR_ = 'the picture in the cell could not be read: ' + String(e.message || e).slice(0, 80); }
  return '';
}

var STIMG_ERR_ = '';   // reason the last picture could not be read (shown on the sheet / in Style Tracker)
function stImgDataUrl_(val, maxPx) {
  STIMG_ERR_ = '';
  var s = String(val || '').trim();
  if (!s) return '';
  var cm = /^cell:R(\d+)$/.exec(s);
  if (cm) return stCellImageData_(Number(cm[1]), maxPx || 800);
  var m = /\/d\/([A-Za-z0-9_-]{20,})/.exec(s) || /[?&]id=([A-Za-z0-9_-]{20,})/.exec(s) || /^([A-Za-z0-9_-]{25,})$/.exec(s);
  if (m) {
    try {
      var blob = DriveApp.getFileById(m[1]).getBlob();
      var ct = String(blob.getContentType() || '');
      if (!/^image\//.test(ct)) { STIMG_ERR_ = 'the saved file is not a picture (' + (ct || 'unknown type') + ')'; return ''; }
      return 'data:' + ct + ';base64,' + Utilities.base64Encode(blob.getBytes());
    } catch (e) {
      STIMG_ERR_ = /permission|access|No item/i.test(String(e.message)) ?
        'this account cannot open the picture file in Drive (not shared with it, or deleted)' : String(e.message || e).slice(0, 90);
      return '';
    }
  }
  if (/^data:image\//i.test(s)) return s;
  if (/^https?:\/\//i.test(s)) {   // an outside link: embed it (the PDF maker cannot fetch it reliably)
    try {
      var r = UrlFetchApp.fetch(s, { muteHttpExceptions: true, followRedirects: true });
      var rb = r.getBlob();
      if (r.getResponseCode() === 200 && /^image\//.test(String(rb.getContentType() || ''))) return 'data:' + rb.getContentType() + ';base64,' + Utilities.base64Encode(rb.getBytes());
      STIMG_ERR_ = 'the picture link did not return an image (HTTP ' + r.getResponseCode() + ')';
    } catch (e2) { STIMG_ERR_ = 'the picture link could not be opened'; }
    return '';
  }
  STIMG_ERR_ = 'the saved picture link is not recognised: "' + s.slice(0, 40) + '"';
  return '';
}

/* Pixel size of a PNG / GIF / JPEG from its bytes (Apps Script bytes are signed, hence & 0xFF) */
function imgDims_(bytes) {
  var b = function (i) { return bytes[i] & 0xFF; };
  if (bytes.length > 24 && b(0) === 0x89 && b(1) === 0x50) {
    return { w: (b(16) * 16777216) + (b(17) << 16) + (b(18) << 8) + b(19), h: (b(20) * 16777216) + (b(21) << 16) + (b(22) << 8) + b(23) };
  }
  if (bytes.length > 10 && b(0) === 0x47 && b(1) === 0x49) return { w: b(6) | (b(7) << 8), h: b(8) | (b(9) << 8) };
  if (bytes.length > 4 && b(0) === 0xFF && b(1) === 0xD8) {
    var i = 2;
    while (i + 9 < bytes.length) {
      if (b(i) !== 0xFF) { i++; continue; }
      var m = b(i + 1);
      if (m === 0xD8 || m === 0x01 || m === 0xFF || (m >= 0xD0 && m <= 0xD7)) { i += (m === 0xFF ? 1 : 2); continue; }
      var len = (b(i + 2) << 8) | b(i + 3);
      if (m >= 0xC0 && m <= 0xCF && m !== 0xC4 && m !== 0xC8 && m !== 0xCC) return { h: (b(i + 5) << 8) | b(i + 6), w: (b(i + 7) << 8) | b(i + 8) };
      i += 2 + len;
    }
  }
  return null;
}
/* <img> that fits inside maxW x maxH with EXPLICIT width/height attributes. The PDF converter ignores
   max-width / max-height, which is why pictures sometimes vanished or overflowed the page. */
function imgTag_(dataUrl, maxW, maxH, style) {
  if (!dataUrl) return '';
  var w = maxW, h = '';
  try {
    var m = /^data:[^;,]+;base64,(.+)$/.exec(dataUrl);
    if (m) {
      var d = imgDims_(Utilities.base64Decode(m[1]));
      if (d && d.w > 0 && d.h > 0) { var sc = Math.min(maxW / d.w, maxH / d.h, 1); w = Math.max(1, Math.round(d.w * sc)); h = Math.max(1, Math.round(d.h * sc)); }
    }
  } catch (e) { }
  return '<img src="' + dataUrl + '" width="' + w + '"' + (h ? ' height="' + h + '"' : '') + ' style="' + (style === undefined ? 'border:1px solid #ccc;' : style) + '">';
}

/* Small thumbnail for lists / PDFs (falls back to the full picture if Drive has no usable thumbnail) */
function stImgThumbDataUrl_(val) {
  STIMG_ERR_ = '';
  var s = String(val || '').trim();
  if (!s) return '';
  if (/^cell:R\d+$/.test(s)) return stImgDataUrl_(s, 160);
  var m = /\/d\/([A-Za-z0-9_-]{20,})/.exec(s) || /[?&]id=([A-Za-z0-9_-]{20,})/.exec(s) || /^([A-Za-z0-9_-]{25,})$/.exec(s);
  if (m) {
    try {
      var t = DriveApp.getFileById(m[1]).getThumbnail();
      if (t) {
        var tb = t.getBytes(), ct = String(t.getContentType() || '');
        if (tb && tb.length > 200) return 'data:' + (/^image\//.test(ct) ? ct : 'image/png') + ';base64,' + Utilities.base64Encode(tb);
      }
    } catch (e) { }
  }
  return stImgDataUrl_(s);
}

/* Fabric consumption per garment in GRAMS, part-wise (Style Tracker keeps it as kg per piece) */
function consumptionForRows_(rows) {
  rows = (rows || []).filter(function (r) { return !isSampleWtRow_(r); });
  var byPart = {}, order = [];
  rows.forEach(function (r) {
    var part = String(r['Part'] || '').trim() || '-', g = Math.round((Number(r['CAD Weight (g)']) || 0) * 10000) / 10;
    if (!(g > 0)) return;
    var partKey = part.toUpperCase();   // "BODY" and "Body" are the same part — group by it regardless of how it was typed
    if (!byPart[partKey]) { byPart[partKey] = { part: part, vals: {}, order: [], fabric: String(r['Fabric type'] || ''), gsm: String(r['GSM'] || ''), loss: Number(r['Loss %']) || 0 }; order.push(partKey); }
    var P = byPart[partKey], k = String(g);
    if (!P.vals[k]) { P.vals[k] = { grams: g, colors: [] }; P.order.push(k); }
    var c = String(r['Color'] || '').trim(); if (c && P.vals[k].colors.indexOf(c) < 0) P.vals[k].colors.push(c);
  });
  var lines = [], total = 0, varies = false;
  order.forEach(function (pt) {
    var P = byPart[pt];
    P.order.forEach(function (k) { lines.push({ part: P.part, grams: P.vals[k].grams, colors: P.order.length > 1 ? P.vals[k].colors : [], fabric: P.fabric, gsm: P.gsm, loss: P.loss }); });
    total += P.vals[P.order[0]].grams; if (P.order.length > 1) varies = true;
  });
  return { lines: lines, totalG: Math.round(total * 10) / 10, varies: varies };
}
function consTableHtml_(c, cls) {
  if (!c || !c.lines || !c.lines.length) return '<div style="color:#777;font-size:10px;">Consumption not entered in Style Tracker yet.</div>';
  var E = function (v) { return String(v == null ? '' : v).replace(/[&<>"]/g, function (x) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[x]; }); };
  var g = function (n) { return (Math.round(n * 10) / 10).toLocaleString('en-IN'); };
  return '<table' + (cls ? ' class="' + cls + '"' : '') + '><thead><tr><th class="l" style="text-align:left">Part</th><th class="l" style="text-align:left">Fabric</th><th>GSM</th><th style="text-align:right">Consumption (g / pc)</th><th style="text-align:right">Loss %</th></tr></thead><tbody>' +
    c.lines.map(function (l) {
      return '<tr><td class="l" style="text-align:left"><b>' + E(l.part) + '</b>' + (l.colors.length ? ' <span style="color:#777">(' + E(l.colors.join(', ')) + ')</span>' : '') + '</td><td class="l" style="text-align:left">' + E(l.fabric || '-') + '</td><td>' + E(l.gsm || '-') +
        '</td><td style="text-align:right"><b>' + g(l.grams) + ' g</b></td><td style="text-align:right">' + (l.loss ? g(l.loss) + '%' : '-') + '</td></tr>';
    }).join('') +
    '<tr><td class="l" style="text-align:left" colspan="3"><b>Total per garment' + (c.varies ? ' (first colour)' : '') + '</b></td><td style="text-align:right"><b>' + g(c.totalG) + ' g</b></td><td></td></tr></tbody></table>';
}

/* Plain-words reason a JO has no picture on a printed sheet */
function stImgWhy_(link) {
  return link ? ('a picture is saved but ' + (STIMG_ERR_ || 'it could not be read')) : 'no picture has been uploaded for this JO yet';
}

function stImgLinkForJo_(joNumber, rows) {
  rows = rows || getStyleTracker().filter(function (r) { return String(r['JO Number']).trim() === String(joNumber).trim(); });
  for (var i = 0; i < rows.length; i++) {
    var v = String(rows[i][STYLE_IMAGE_COL] || '').trim();
    if (!v) continue;
    if (v === 'CellImage' && rows[i]._row) return 'cell:R' + rows[i]._row;   // picture pasted into the cell itself
    return v;
  }
  return '';
}

function getStyleImage(joNumber) {
  var link = stImgLinkForJo_(joNumber);
  var d = stImgDataUrl_(link);
  return { jo: joNumber, link: link, dataUrl: d, error: d ? '' : (link ? STIMG_ERR_ : '') };
}

function uploadStyleImage(joNumber, fileName, base64Data, mimeType) {
  try {
    if (!joNumber) return { success: false, error: 'Select a JO first.' };
    if (!getJOInfo(joNumber)) return { success: false, error: 'JO ' + joNumber + ' not found in Style Tracker.' };
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType || 'image/jpeg',
      String(joNumber).replace(/[^\w-]/g, '_') + '_' + (fileName || 'style.jpg'));
    var file = stImgFolder_().createFile(blob);
    stImgShare_(file);
    var link = 'https://drive.google.com/file/d/' + file.getId() + '/view';
    var n = stImgWrite_(joNumber, link);
    return { success: true, message: 'Image saved against ' + joNumber + ' (' + n + ' Style Tracker rows).', link: link,
      dataUrl: 'data:' + blob.getContentType() + ';base64,' + base64Data };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/* view-only sharing so every user of the app can read the picture (falls back to domain-only if outside sharing is blocked) */
function stImgShare_(file) {
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); return true; } catch (e) { }
  try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); return true; } catch (e2) { }
  return false;
}
/* RUN ONCE from the Apps Script editor (as the file owner): shares every picture already saved in Style Tracker */
function shareAllStyleImages() {
  var seen = {}, ok = 0, bad = [];
  getStyleTracker().forEach(function (r) {
    var link = String(r[STYLE_IMAGE_COL] || '').trim();
    var m = /\/d\/([A-Za-z0-9_-]{20,})/.exec(link) || /[?&]id=([A-Za-z0-9_-]{20,})/.exec(link);
    if (!m || seen[m[1]]) return;
    seen[m[1]] = 1;
    try { if (stImgShare_(DriveApp.getFileById(m[1]))) ok++; else bad.push(String(r['JO Number']) + ' (sharing blocked)'); }
    catch (e) { bad.push(String(r['JO Number']) + ' (' + String(e.message).slice(0, 50) + ')'); }
  });
  var msg = ok + ' picture(s) shared.' + (bad.length ? ' Problems: ' + bad.join('; ') : '');
  Logger.log(msg); return msg;
}

function removeStyleImage(joNumber) {
  stImgWrite_(joNumber, '');
  return { success: true, message: 'Image removed from ' + joNumber + '.' };
}

/*************** ORDER IN HAND (built from Style Tracker) ***************/
function getOrderInHand() {
  var st = getStyleTracker();
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var num = function (v) { return Number(v) || 0; };
  var s = function (v) { return String(v == null ? '' : v).trim(); };
  var sumBy = function (rows, keyFn, valFn) {
    var o = {};
    rows.forEach(function (r) { var k = keyFn(r); if (k) o[k] = (o[k] || 0) + valFn(r); });
    return o;
  };
  var poKg = sumBy(sheetToObjects(SHEETS.YARN_PO).filter(function (r) { return r['Status'] !== 'Rejected'; }),
    function (r) { return s(r['JO Number']); }, function (r) { return num(r['Qty KG']); });
  var rcvKg = sumBy(sheetToObjects(SHEETS.YARN_INWARD), function (r) { return s(r['JO Number']); }, function (r) { return num(r['Received KG']); });
  var dcs = sheetToObjects(SHEETS.PROCESS_DC).filter(function (r) { return r['Status'] !== 'Rejected'; });
  var toProc = function (p) {
    return sumBy(dcs.filter(function (r) { return String(r['To Process'] || '').indexOf(p) === 0; }),
      function (r) { return s(r['JO Number']); }, function (r) { return num(r['DC Issue Qty KG']); });
  };
  var knit = toProc('Knitting'), dye = toProc('Dyeing'), inward = toProc(JIWIN_INWARD_LABEL);
  var status = {};
  sheetToObjects(SHEETS.JO_STATUS).forEach(function (r) { status[s(r['JO Number'])] = r['Status']; });
  var tasks = {};
  try {
    getAllTasks_().forEach(function (t) {
      var jo = s(t['JO Number']); if (!jo || t['Status'] === 'Completed') return;
      tasks[jo] = tasks[jo] || { open: 0, overdue: 0 };
      tasks[jo].open++;
      var due = String(t['Due Date'] || '').slice(0, 10);
      if (due && due < today) tasks[jo].overdue++;
    });
  } catch (e) { }

  var jos = {}, joOrder = [], lines = [], lmap = {};
  st.forEach(function (r) {
    var jo = s(r['JO Number']); if (!jo) return;
    if (!jos[jo]) {
      jos[jo] = { jo: jo, buyer: s(r['Buyer Name']), fabricKg: 0, status: status[jo] || 'Open', merch: s(r['Merchandiser']),
        ocn: s(r['Order/OCN No']), season: s(r['Season']) };
      joOrder.push(jo);
    }
    jos[jo].fabricKg += num(r['Fabric Required (kg)']);

    // OIH must show every real piece-based order from Style Tracker. Do not
    // discard a JO merely because one row has Sample Weight (kg). Weight-only
    // sample rows are ignored only when they contain no garment/style/color/qty.
    var style = s(r['Style Number']), color = s(r['Color']), part = s(r['Part']).toUpperCase();
    var qty = num(r['Order Quantity']);
    var hasSizeQty = false;
    for (var qi = 1; qi <= 10; qi++) {
      if (num(r['Qty ' + qi]) > 0) { hasSizeQty = true; break; }
    }
    if (!style && !color && !qty && !hasSizeQty) return;

    // Prefer BODY / BODY+FOLDING over RIB when the same JO + Style + Colour
    // appears on multiple component rows. RIB must never replace the garment
    // quantity shown in Order In Hand.
    var key = jo + '|' + style + '|' + color;
    if (lmap[key]) {
      if (part.indexOf('RIB') === -1 && lmap[key]._isRib) {
        lmap[key].qty = qty;
        lmap[key].fabric = s(r['Fabric type']);
        lmap[key].orderStatus = s(r['Order Status']);
        lmap[key].remarks = s(r['Order Remarks']);
        lmap[key]._isRib = false;
      }
      if (s(r[STYLE_IMAGE_COL])) lmap[key].hasImage = true;
      return;
    }
    var L = {
      jo: jo, style: style, desc: s(r['Style Description']), color: color, buyer: s(r['Buyer Name']),
      qty: qty, despatch: s(r['Despatch Date']).slice(0, 10), fabric: s(r['Fabric type']),
      type: s(r['Order Type']) || (/^S\s*-|^SAMPLE/i.test(jo) ? 'SAMPLE' : 'ORDER'), orderStatus: s(r['Order Status']),
      remarks: s(r['Order Remarks']), hasImage: !!s(r[STYLE_IMAGE_COL]), tna: {}, _isRib: part.indexOf('RIB') !== -1
    };
    ST_ORDER_DATE_COLS.forEach(function (h) { L.tna[h] = s(r[h]).slice(0, 10); });
    lmap[key] = L; lines.push(L);
  });

  // Internal helper flag is not needed by the browser.
  lines.forEach(function (l) { delete l._isRib; });

  var r2 = function (n) { return Math.round((n || 0) * 100) / 100; };
  joOrder.forEach(function (jo) {
    var o = jos[jo];
    o.fabricKg = r2(o.fabricKg); o.yarnPoKg = r2(poKg[jo]); o.yarnRcvKg = r2(rcvKg[jo]);
    o.knitKg = r2(knit[jo]); o.dyeKg = r2(dye[jo]); o.inwardKg = r2(inward[jo]);
    o.tasksOpen = tasks[jo] ? tasks[jo].open : 0; o.tasksOverdue = tasks[jo] ? tasks[jo].overdue : 0;
    o.stage = o.inwardKg > 0 ? 'Fabric Inward' : o.dyeKg > 0 ? 'Dyeing' : o.knitKg > 0 ? 'Knitting' :
      o.yarnRcvKg > 0 ? 'Yarn received' : o.yarnPoKg > 0 ? 'Yarn ordered' : 'Not started';
  });
  return { jos: jos, lines: lines, dateCols: ST_ORDER_DATE_COLS };
}

/* Picture for one JO + Style + Colour line (falls back to the JO picture) */
function getLineImage(jo, style, color) {
  var rows = getStyleTracker().filter(function (r) { return String(r['JO Number']).trim() === String(jo).trim(); });
  var own = rows.filter(function (r) { return String(r['Style Number']).trim() === String(style).trim() && String(r['Color']).trim() === String(color).trim(); });
  var link = stImgLinkForJo_(jo, own.length ? own : rows) || stImgLinkForJo_(jo, rows);
  var d = stImgDataUrl_(link);
  return { dataUrl: d, error: d ? '' : (link ? STIMG_ERR_ : 'no picture uploaded') };
}

/*************** STYLE TRACKER + BUDGET ENTRY (from the web app) ***************
 * Reads/writes Style Tracker by its REAL row-1 headings (never by position),
 * so extra columns (Process 6-10, Style Image ...) are never shifted.
 *******************************************************************************/
var ST_JO_FIELDS = ['Style Number', 'Buyer Name', 'Style Description', 'Order/OCN No', 'Buyer Order No',
  'Despatch Date', 'Season', 'Merchandiser', 'Excess %'];
var ST_ROW_FIELDS = ['Part', 'Color', 'CAD Weight (g)', 'Loss %', 'Counts', 'Yarn Type', 'Content', 'Has Lycra', 'Lycra %',
  'Fabric type', 'GSM', 'Knitting Dia', 'Finishing dia', 'Loop length', 'Gauge', 'Yarn color', 'Yarn rate'];

function stHead_(sh) {
  var lastCol = Math.max(sh.getLastColumn(), 1);
  return sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h || '').trim(); });
}

/* Budget columns = every process rate / loss % column (sheet order), paired per process */
function stBudgetGroups_(head) {
  var a = head.indexOf('Yarn rate'), b = head.indexOf('Order/OCN No'), m = head.indexOf('Merchandiser');
  var cols = [];
  head.forEach(function (h, i) {
    if (!h) return;
    if ((i > a && i < b) || (i > m && /(rate|loss %)$/i.test(h))) cols.push(h);
  });
  var groups = [];
  cols.forEach(function (h) {
    var isLoss = /loss %$/i.test(h);
    var last = groups[groups.length - 1];
    if (isLoss && last && !last.loss) { last.loss = h; return; }
    var name = h.replace(/\s*rate$/i, '').replace(/\s*loss %$/i, '').replace(/^Stener$/i, 'Stenter').replace(/^Rasing \/Peach fnish$/i, 'Raising / Peach finish');
    groups.push(isLoss ? { name: name, rate: '', loss: h } : { name: name, rate: h, loss: '' });
  });
  return groups;
}

function stNextJo_(values, joIdx) {
  var max = 0;
  values.forEach(function (r) {
    var s = String(r[joIdx] || '').trim();
    if (/^S\s*-/i.test(s)) return;                       // sample JOs don't count
    var m = /^JO\s*-?\s*(\d+)/i.exec(s);
    if (m && Number(m[1]) > max) max = Number(m[1]);
  });
  return 'JO-' + (max + 1);
}

function getStyleTrackerEditor(joNumber) {
  var sh = SS.getSheetByName(SHEETS.STYLE_TRACKER);
  var head = stHead_(sh);
  var last = sh.getLastRow();
  var values = last > 1 ? sh.getRange(2, 1, last - 1, head.length).getValues() : [];
  var joIdx = head.indexOf('JO Number');
  var tz = Session.getScriptTimeZone();
  var toObj = function (r) {
    var o = {};
    head.forEach(function (h, i) { if (h) o[h] = r[i] instanceof Date ? Utilities.formatDate(r[i], tz, 'yyyy-MM-dd') : stPlainCell_(r[i]); });
    return o;
  };
  var lists = { buyers: {}, parts: {}, fabrics: {}, counts: {}, yarnTypes: {}, contents: {}, merch: {}, colors: {}, sizes: {} };
  var sizeSets = {}, sizeSetOrder = [];
  values.forEach(function (r) {
    var o = toObj(r);
    if (Number(o[SAMPLE_WT_COL]) > 0) return;   // weight-only sample lines don't feed the pick-lists
    // size runs used before (e.g. "S M L XL", "98 104 110 116"), newest first, with buyer + JO
    var run = [];
    for (var z = 1; z <= 10; z++) { var lb = String(o['Size ' + z] || '').trim(); if (lb) { run.push(lb); lists.sizes[lb] = 1; } }
    if (run.length) {
      var key = run.join('|');
      if (!sizeSets[key]) { sizeSets[key] = { sizes: run, buyers: [], jos: [] }; sizeSetOrder.push(key); }
      var b = String(o['Buyer Name'] || '').trim(), j = String(o['JO Number'] || '').trim();
      if (b && sizeSets[key].buyers.indexOf(b) < 0) sizeSets[key].buyers.push(b);
      if (j && sizeSets[key].jos.indexOf(j) < 0) sizeSets[key].jos.push(j);
    }
    var add = function (k, v) { v = String(v || '').trim(); if (v) lists[k][v] = 1; };
    add('buyers', o['Buyer Name']); add('parts', o['Part']); add('fabrics', o['Fabric type']); add('counts', o['Counts']);
    add('yarnTypes', o['Yarn Type']); add('contents', o['Content']); add('merch', o['Merchandiser']); add('colors', o['Color']);
  });
  Object.keys(lists).forEach(function (k) { lists[k] = Object.keys(lists[k]).sort(); });
  lists.sizeSets = sizeSetOrder.reverse().map(function (k) { return sizeSets[k]; });
  var joRows = joNumber ? values.filter(function (r) { return String(r[joIdx]).trim() === String(joNumber).trim(); }).map(toObj) : [];
  var rows = joRows.filter(function (r) { return !(Number(r[SAMPLE_WT_COL]) > 0); });
  var samples = joRows.filter(function (r) { return Number(r[SAMPLE_WT_COL]) > 0; })
    .map(function (r) { return { Part: r['Part'] || '', kg: Number(r[SAMPLE_WT_COL]) || 0 }; });
  var budget = stBudgetGroups_(head), known = {};
  [SAMPLE_WT_COL, 'JO Number', STYLE_IMAGE_COL, 'Order Quantity', 'Sizewise Total', 'Size Check', 'Excess Quantity', 'Fabric Required (kg)',
    'Main Fiber Reqd (kg)', 'Lycra Reqd (kg)', 'No Of Parts'].concat(ST_JO_FIELDS, ST_ROW_FIELDS, ST_VIEW_COLS).forEach(function (h) { known[h] = 1; });
  for (var z = 1; z <= 10; z++) { known['Size ' + z] = 1; known['Qty ' + z] = 1; }
  budget.forEach(function (g) { if (g.rate) known[g.rate] = 1; if (g.loss) known[g.loss] = 1; });
  var extra = head.filter(function (h) { return h && !known[h]; });
  ST_ORDER_COLS.forEach(function (h) { if (extra.indexOf(h) < 0) extra.push(h); });   // always ask order / TNA fields
  return {
    budgetLock: joNumber ? getBudgetLock(joNumber) : null,
    dataLock: joNumber ? lockWithSheets_(joNumber) : null,
    extra: extra, orderCols: ST_ORDER_COLS, orderDateCols: ST_ORDER_DATE_COLS,
    jo: joNumber || '', nextJo: stNextJo_(values, joIdx), exists: joRows.length > 0,
    rows: rows, samples: samples, budget: stBudgetGroups_(head), lists: lists,
    hasImage: joRows.some(function (r) { return String(r[STYLE_IMAGE_COL] || '').trim(); })
  };
}

/*
 * payload = { jo, isNew, header:{Style Number..Excess %}, sizes:[10 labels],
 *             rows:[{Part, Color, qty:[10], CAD Weight (g) ... Yarn rate}], rates:{<rate/loss column>: value} }
 */
function saveStyleTrackerJo(p) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var jo = String(p.jo || '').trim();
    if (!jo) return { success: false, error: 'Enter the JO Number.' };
    if (getJoDataLock(jo)) return { success: false, error: 'JO ' + jo + ' is locked (Order Sheet + Fabric Program are final). Unlock it first if this change is genuinely needed.' };
    var rows = (p.rows || []).filter(function (r) { return String(r.Part || '').trim() || String(r.Color || '').trim(); });
    if (!rows.length) return { success: false, error: 'Add at least one Part / Colour line.' };
    for (var i = 0; i < rows.length; i++) {
      if (!String(rows[i].Part || '').trim() || !String(rows[i].Color || '').trim()) {
        return { success: false, error: 'Line ' + (i + 1) + ': Part and Colour are both needed.' };
      }
    }
    var sh = SS.getSheetByName(SHEETS.STYLE_TRACKER);
    var head = stHead_(sh);
    // make sure every column we write exists (added at the end, never inserted)
    var need = ['JO Number', 'Sizewise Total', 'Size Check', 'Order Quantity', 'Excess Quantity', 'Fabric Required (kg)',
      'Main Fiber Reqd (kg)', 'Lycra Reqd (kg)', 'No Of Parts', SAMPLE_WT_COL, 'Order Type'].concat(ST_JO_FIELDS, ST_ROW_FIELDS);
    for (var s = 1; s <= 10; s++) need.push('Size ' + s, 'Qty ' + s);
    Object.keys(p.rates || {}).forEach(function (k) { need.push(k); });
    Object.keys(p.partRates || {}).forEach(function (pt) { Object.keys(p.partRates[pt] || {}).forEach(function (k) { if (need.indexOf(k) < 0) need.push(k); }); });
    rows.forEach(function (r) { Object.keys(r.extra || {}).forEach(function (k) { if (need.indexOf(k) < 0) need.push(k); }); });
    var added = false;
    need.forEach(function (h) { if (head.indexOf(h) < 0) { head.push(h); added = true; } });
    if (added) {
      if (sh.getMaxColumns() < head.length) sh.insertColumnsAfter(sh.getMaxColumns(), head.length - sh.getMaxColumns());
      sh.getRange(1, 1, 1, head.length).setValues([head]);
    }
    var col = {}; head.forEach(function (h, i) { if (h && !(h in col)) col[h] = i; });

    var last = sh.getLastRow();
    var values = last > 1 ? sh.getRange(2, 1, last - 1, head.length).getValues() : [];
    var existing = [];
    values.forEach(function (r, i) { if (String(r[col['JO Number']]).trim() === jo) existing.push({ row: i + 2, vals: r }); });
    if (p.isNew && existing.length) return { success: false, error: 'JO ' + jo + ' already exists. Open it from the list to edit.' };
    // weight-only sample lines live in the same JO but are managed separately from the piece-based lines
    var existingSamples = existing.filter(function (e) { return Number(e.vals[col[SAMPLE_WT_COL]]) > 0; });
    existing = existing.filter(function (e) { return !(Number(e.vals[col[SAMPLE_WT_COL]]) > 0); });

    var parts = {}; rows.forEach(function (r) { parts[String(r.Part).trim().toUpperCase()] = 1; });
    var nParts = Object.keys(parts).length;
    var lockInfo = getBudgetLock(jo), lockedCols = [], lockedByPart = {};
    if (lockInfo) {
      stBudgetGroups_(head).forEach(function (g) { if (g.rate) lockedCols.push(g.rate); if (g.loss) lockedCols.push(g.loss); });
      lockedCols.push('Yarn rate');
      existing.forEach(function (e) {
        var pk = String(e.vals[col['Part']] || '').trim().toUpperCase() + '|' + String(e.vals[col['Color']] || '').trim().toUpperCase();
        var pp = String(e.vals[col['Part']] || '').trim().toUpperCase();
        var snap = {}; lockedCols.forEach(function (c) { if (col[c] !== undefined) snap[c] = e.vals[col[c]]; });
        lockedByPart[pk] = snap; if (!lockedByPart[pp]) lockedByPart[pp] = snap;
      });
    }
    var keepImg = existing.length && col[STYLE_IMAGE_COL] !== undefined ? existing[0].vals[col[STYLE_IMAGE_COL]] : '';
    if (!keepImg && p.copyImageFrom) {            // "Copy as new JO": reuse the old JO's picture
      if (col[STYLE_IMAGE_COL] === undefined) { head.push(STYLE_IMAGE_COL); col[STYLE_IMAGE_COL] = head.length - 1;
        if (sh.getMaxColumns() < head.length) sh.insertColumnsAfter(sh.getMaxColumns(), head.length - sh.getMaxColumns());
        sh.getRange(1, head.length).setValue(STYLE_IMAGE_COL); }
      try { keepImg = stImgLinkForJo_(p.copyImageFrom) || ''; } catch (e) { }
    }
    var num = function (v) { return v === '' || v === null || v === undefined ? '' : (isNaN(Number(v)) ? v : Number(v)); };

    var out = rows.map(function (r, i) {
      var base = existing[i] ? existing[i].vals.slice() : head.map(function () { return ''; });
      if (!existing[i] && keepImg) base[col[STYLE_IMAGE_COL]] = keepImg;
      var o = { 'JO Number': jo };
      ST_JO_FIELDS.forEach(function (f) { o[f] = (p.header || {})[f] !== undefined ? p.header[f] : ''; });
      ST_ROW_FIELDS.forEach(function (f) { o[f] = r[f] !== undefined ? r[f] : ''; });
      if (Number(o['CAD Weight (g)']) > 5) o['CAD Weight (g)'] = Number(o['CAD Weight (g)']) / 1000;   // grams given -> kg
      Object.keys(r.extra || {}).forEach(function (f) { if (col[f] !== undefined) o[f] = r.extra[f]; });
      var total = 0;
      for (var k = 1; k <= 10; k++) {
        var lbl = String((p.sizes || [])[k - 1] || '').trim();
        var q = lbl ? (Number((r.qty || [])[k - 1]) || 0) : '';
        o['Size ' + k] = lbl; o['Qty ' + k] = lbl ? q : '';
        if (lbl) total += q;
      }
      o['Order Quantity'] = total;
      // process rates: per part (Body / Rib ...) when given, else JO-level
      var pr = (p.partRates || {})[String(r.Part || '').trim().toUpperCase()] || null, rk = {};
      Object.keys(p.rates || {}).forEach(function (k) { rk[k] = 1; });
      Object.keys(p.partRates || {}).forEach(function (pt) { Object.keys(p.partRates[pt] || {}).forEach(function (k) { rk[k] = 1; }); });
      Object.keys(rk).forEach(function (k) { o[k] = pr && pr[k] !== undefined ? pr[k] : ((p.rates || {})[k] !== undefined ? p.rates[k] : ''); });
      if (lockInfo) {   // locked budget: rates / loss % / yarn rate come from what was saved, never from the form
        var pp = String(r.Part || '').trim().toUpperCase(), snap = lockedByPart[pp + '|' + String(r.Color || '').trim().toUpperCase()] || lockedByPart[pp] || {};
        lockedCols.forEach(function (c) { o[c] = snap[c] !== undefined ? snap[c] : ''; });
      }
      computeStyleTrackerDerived_(o);
      o['No Of Parts'] = nParts;
      Object.keys(o).forEach(function (k) {
        if (col[k] === undefined) return;
        var v = o[k];
        if (['Excess %', 'CAD Weight (g)', 'Loss %', 'Lycra %', 'GSM', 'Yarn rate'].indexOf(k) > -1 || /(rate|loss %)$/i.test(k)) v = num(v);
        base[col[k]] = v;
      });
      return base;
    });

    // ---- sample lines (same JO, weight only): copy the part's spec from its piece-based line, no pcs / sizes ----
    var sampleOut = [];
    var samplesIn = [], sampleIdx = {};   // one sample line per Part (same part entered twice = weights added up)
    (p.samples || []).forEach(function (x) {
      var k = String(x.Part || '').trim().toUpperCase();
      if (!k || !(Number(x.kg) > 0)) return;
      if (sampleIdx[k] === undefined) { sampleIdx[k] = samplesIn.length; samplesIn.push({ Part: String(x.Part).trim(), kg: Number(x.kg) }); }
      else samplesIn[sampleIdx[k]].kg += Number(x.kg);
    });
    for (var q = 0; q < samplesIn.length; q++) {
      var sp = String(samplesIn[q].Part).trim().toUpperCase(), srcIdx = -1;
      for (var z = 0; z < rows.length; z++) { if (String(rows[z].Part).trim().toUpperCase() === sp) { srcIdx = z; break; } }
      if (srcIdx < 0) return { success: false, error: 'Sample "' + samplesIn[q].Part + '": no line with that Part in this JO. Add the ' + samplesIn[q].Part + ' line above first, or use the same Part name.' };
      var sb = out[srcIdx].slice(), skg = Number(samplesIn[q].kg);
      sb[col['Color']] = 'SAMPLE';
      for (var kz = 1; kz <= 10; kz++) { sb[col['Size ' + kz]] = ''; sb[col['Qty ' + kz]] = ''; }
      sb[col['Order Quantity']] = 0; sb[col['Sizewise Total']] = 0; sb[col['Size Check']] = 'OK'; sb[col['Excess Quantity']] = 0;
      sb[col['Fabric Required (kg)']] = skg;
      var hasLy = String(sb[col['Has Lycra']]).trim().toLowerCase() === 'yes', lp = Number(sb[col['Lycra %']]) || 0;
      if (lp > 0 && lp < 1) lp = lp * 100;
      var lyKg = hasLy ? Math.round(skg * lp / 100 * 100) / 100 : 0;
      sb[col['Lycra Reqd (kg)']] = lyKg; sb[col['Main Fiber Reqd (kg)']] = Math.round((skg - lyKg) * 100) / 100;
      sb[col[SAMPLE_WT_COL]] = skg; sb[col['Order Type']] = 'SAMPLE';
      if (col['Order Line ID'] !== undefined) sb[col['Order Line ID']] = '';
      sampleOut.push(sb);
    }

    // sheet-only view formula columns (Main Line, Days to Ship ...): keep them as formulas
    var viewCols = ST_VIEW_COLS.filter(function (h) { return col[h] !== undefined; });
    var viewR1C1 = {};
    viewCols.forEach(function (h) {
      var c = col[h] + 1, lr = sh.getLastRow();
      if (lr < 2) return;
      var f = sh.getRange(2, c, lr - 1, 1).getFormulasR1C1();
      for (var k = 0; k < f.length; k++) if (f[k][0]) { viewR1C1[h] = f[k][0]; break; }
    });
    out.concat(sampleOut).forEach(function (row) { viewCols.forEach(function (h) { row[col[h]] = ''; }); });

    // write: update existing rows in place, append extra, delete surplus (bottom-up)
    var n = Math.min(existing.length, out.length), written = [];
    for (var a = 0; a < n; a++) { sh.getRange(existing[a].row, 1, 1, head.length).setValues([out[a]]); written.push(existing[a].row); }
    var appendRows = out.slice(n).concat(sampleOut);
    if (appendRows.length) {
      var startRow = sh.getLastRow() + 1;
      sh.getRange(startRow, 1, appendRows.length, head.length).setValues(appendRows);
      for (var w = 0; w < appendRows.length; w++) written.push(startRow + w);
    }
    viewCols.forEach(function (h) {
      if (!viewR1C1[h]) return;
      written.forEach(function (r) { sh.getRange(r, col[h] + 1).setFormulaR1C1(viewR1C1[h]); });
    });
    var delRows = existing.slice(n).map(function (e) { return e.row; }).concat(existingSamples.map(function (e) { return e.row; }));
    delRows.sort(function (a, b) { return b - a; }).forEach(function (rw) { sh.deleteRow(rw); });
    SpreadsheetApp.flush();

    var kg = out.reduce(function (s2, r) { return s2 + (Number(r[col['Fabric Required (kg)']]) || 0); }, 0);
    var skgTotal = sampleOut.reduce(function (s2, r) { return s2 + (Number(r[col['Fabric Required (kg)']]) || 0); }, 0);
    return { success: true, jo: jo, message: 'Style Tracker saved for ' + jo + ' — ' + out.length + ' line(s), fabric ' + kg.toFixed(2) + ' kg' +
      (sampleOut.length ? ' + sample ' + skgTotal.toFixed(2) + ' kg (' + sampleOut.length + ' sample line(s))' : '') + '.' };
  } catch (err) {
    return { success: false, error: err.message };
  } finally { lock.releaseLock(); }
}

/* Whole Style Tracker (every column) for the "All Style Tracker" view and Excel download */
function getStyleTrackerAll() {
  var sh = SS.getSheetByName(SHEETS.STYLE_TRACKER);
  var head = stHead_(sh);
  var last = sh.getLastRow();
  var tz = Session.getScriptTimeZone();
  var rows = last > 1 ? sh.getRange(2, 1, last - 1, head.length).getValues() : [];
  rows = rows.filter(function (r) { return String(r[0]).trim(); }).map(function (r) {
    return r.map(function (v) { return v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : stPlainCell_(v); });
  });
  var imgCol = head.indexOf(STYLE_IMAGE_COL);
  return { headers: head, rows: rows, imageCol: imgCol };
}

/*************** ORDERS  ->  STYLE TRACKER (one combined sheet) ***************
 * Style Tracker is the single master. The order-follow-up fields that used to
 * live in ORDERS (status, TNA dates, remarks, price, type, picture) are now
 * columns at the end of Style Tracker, stored on every row (part) of that
 * JO + Style + Colour.
 *******************************************************************************/
var ST_ORDER_COLS = ['Order Type', 'Order Status', 'Yarn Inhouse', 'Fabric Inhouse', 'Trims IH Date', 'PP AD',
  'Cutting Date', 'Cut Days', 'Print Start', 'Print Days', 'Pro Start', 'Pro Days', 'FI Days', 'Order Remarks',
  'Buyer Price (USD)', 'Order Line ID'];
/* Sheet-only "Order In Hand view" formula columns - the app never edits or overwrites them */
var ST_VIEW_COLS = ['Main Line', 'Line Qty (pcs)', 'Days to Ship', 'JO Yarn PO (kg)', 'JO Yarn Rcvd (kg)', 'JO Sent to Knitting (kg)', 'JO Sent to Dyeing (kg)'];
var ST_ORDER_DATE_COLS = ['Yarn Inhouse', 'Fabric Inhouse', 'Trims IH Date', 'PP AD', 'Cutting Date', 'Print Start', 'Pro Start'];

function stEnsureOrderCols_() {
  var sh = SS.getSheetByName(SHEETS.STYLE_TRACKER);
  var head = stHead_(sh), add = [];
  ST_ORDER_COLS.concat([STYLE_IMAGE_COL]).forEach(function (h) { if (head.indexOf(h) < 0) add.push(h); });
  if (add.length) {
    var start = head.length + 1;
    if (sh.getMaxColumns() < head.length + add.length) sh.insertColumnsAfter(sh.getMaxColumns(), head.length + add.length - sh.getMaxColumns());
    sh.getRange(1, start, 1, add.length).setValues([add]).setFontWeight('bold').setBackground('#1B2A4A').setFontColor('#FFFFFF');
  }
  return stHead_(sh);
}

function stNormJo_(v) {
  var s = String(v || '').trim();
  if (/^S\s*-/i.test(s) || /^SAMPLE/i.test(s)) return 'SAMPLE:' + s.toUpperCase();
  var m = /JO\s*-?\s*(\d+)/i.exec(s);
  return m ? 'JO-' + Number(m[1]) : s.toUpperCase();
}
function stNorm_(v) { return String(v || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function stLike_(a, b) {            // 2 = same, 1 = one starts with the other, 0 = different
  a = stNorm_(a); b = stNorm_(b);
  if (!a || !b) return 0;
  if (a === b) return 2;
  return (a.indexOf(b) === 0 || b.indexOf(a) === 0) ? 1 : 0;
}

/* dryRun = true only reports what would happen */
function mergeOrdersIntoStyleTracker(dryRun) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ord = SS.getSheetByName('ORDERS');
    if (!ord || ord.getLastRow() < 2) return { success: false, error: 'ORDERS sheet is empty or missing.' };
    var tz = Session.getScriptTimeZone();
    var dv = function (v) { return v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : v; };
    var readObj = function (sh) {
      if (!sh || sh.getLastRow() < 2) return [];
      var v = sh.getDataRange().getValues(), h = v[0].map(function (x) { return String(x).trim(); });
      return v.slice(1).map(function (r) { var o = {}; h.forEach(function (k, i) { o[k] = r[i]; }); return o; });
    };
    var orders = readObj(ord).filter(function (o) { return String(o['ID'] || '').trim() && (o['Style'] || o['Color']); });
    var img = {}, price = {};
    readObj(SS.getSheetByName('IMAGES')).forEach(function (r) { if (r['Line ID']) img[String(r['Line ID']).trim()] = r['Image (data URL)']; });
    readObj(SS.getSheetByName('PRICING')).forEach(function (r) { if (r['ID']) price[String(r['ID']).trim()] = r['Buyer Price (USD)']; });

    var sh = SS.getSheetByName(SHEETS.STYLE_TRACKER);
    var head = dryRun ? stHead_(sh).concat(ST_ORDER_COLS.concat([STYLE_IMAGE_COL]).filter(function (h) { return stHead_(sh).indexOf(h) < 0; })) : stEnsureOrderCols_();
    var col = {}; head.forEach(function (h, i) { if (h && !(h in col)) col[h] = i; });
    var last = sh.getLastRow();
    var width = Math.min(head.length, sh.getMaxColumns());
    var vals = last > 1 ? sh.getRange(2, 1, last - 1, width).getValues() : [];
    vals.forEach(function (r) { while (r.length < head.length) r.push(''); });

    // group Style Tracker rows by JO + Style + Colour
    var groups = [], gmap = {};
    vals.forEach(function (r, i) {
      var jo = String(r[col['JO Number']] || '').trim(); if (!jo) return;
      var k = jo + '|' + String(r[col['Style Number']]).trim() + '|' + String(r[col['Color']]).trim();
      if (!gmap[k]) { gmap[k] = { jo: jo, nj: stNormJo_(jo), style: r[col['Style Number']], color: r[col['Color']], idx: [] }; groups.push(gmap[k]); }
      gmap[k].idx.push(i);
    });

    var rep = { orderLines: orders.length, matched: 0, rowsUpdated: 0, newRows: 0, images: 0, prices: 0, unmatched: [] };
    var appendRows = [];
    orders.forEach(function (o) {
      var type = String(o['Type (ORDER/SAMPLE)'] || '').trim().toUpperCase() || 'ORDER';
      var nj = stNormJo_(o['JO Number'] || (type === 'SAMPLE' ? 'SAMPLE' : ''));
      var best = null, bestScore = 0;
      groups.forEach(function (g) {
        if (!o['JO Number'] || g.nj !== nj) return;
        var c = stLike_(g.color, o['Color']); if (!c) return;
        var s = stLike_(g.style, o['Style']);
        var score = c * 3 + s * 2;
        if (score > bestScore) { bestScore = score; best = g; }
      });
      var id = String(o['ID']).trim();
      var set = {
        'Order Type': type, 'Order Status': o['Status'] || '', 'Order Remarks': o['Remarks'] || '', 'Order Line ID': id,
        'Cut Days': o['Cut Days'] || '', 'Print Days': o['Print Days'] || '', 'Pro Days': o['Pro Days'] || '', 'FI Days': o['FI Days'] || ''
      };
      ST_ORDER_DATE_COLS.forEach(function (h) { set[h] = dv(o[h]) || ''; });
      if (price[id] !== undefined && price[id] !== '') { set['Buyer Price (USD)'] = price[id]; rep.prices++; }
      var picture = img[id] || '';
      if (best) {
        rep.matched++;
        best.idx.forEach(function (i) {
          Object.keys(set).forEach(function (h) { if (set[h] !== '' || !vals[i][col[h]]) vals[i][col[h]] = set[h]; });
          if (!vals[i][col['Despatch Date']] && o['Shipment Date']) vals[i][col['Despatch Date']] = dv(o['Shipment Date']);
          if (picture && !String(vals[i][col[STYLE_IMAGE_COL]] || '').trim()) { vals[i][col[STYLE_IMAGE_COL]] = picture; }
          rep.rowsUpdated++;
        });
        if (picture) rep.images++;
      } else {
        var r = head.map(function () { return ''; });
        var put = function (h, v) { if (col[h] !== undefined) r[col[h]] = v; };
        var sameJo = groups.filter(function (g) { return o['JO Number'] && g.nj === nj; })[0];
        put('JO Number', sameJo ? sameJo.jo : (o['JO Number'] || (type === 'SAMPLE' ? 'SAMPLE' : '')));
        put('Style Number', o['Style'] || ''); put('Buyer Name', o['Buyer'] || ''); put('Style Description', o['Styling'] || '');
        put('Part', 'GARMENT'); put('Color', o['Color'] || ''); put('Order Quantity', Number(o['Order Qty']) || 0);
        put('Despatch Date', dv(o['Shipment Date']) || ''); put('Fabric type', o['Fabric'] || ''); put('Excess %', 0);
        Object.keys(set).forEach(function (h) { put(h, set[h]); });
        if (picture) { put(STYLE_IMAGE_COL, picture); rep.images++; }
        appendRows.push(r); rep.newRows++;
        rep.unmatched.push((o['JO Number'] || type) + ' · ' + (o['Style'] || '') + ' · ' + (o['Color'] || ''));
      }
    });

    if (!dryRun) {
      if (vals.length) sh.getRange(2, 1, vals.length, head.length).setValues(vals);
      if (appendRows.length) sh.getRange(sh.getLastRow() + 1, 1, appendRows.length, head.length).setValues(appendRows);
      ['ORDERS', 'PRICING', 'IMAGES'].forEach(function (n) { var s = SS.getSheetByName(n); if (s) try { s.hideSheet(); } catch (e) { } });
      SpreadsheetApp.flush();
    }
    rep.success = true;
    rep.message = (dryRun ? 'Preview: ' : 'Done: ') + rep.matched + ' of ' + rep.orderLines + ' order lines matched to Style Tracker (' +
      rep.rowsUpdated + ' rows updated), ' + rep.newRows + ' new line(s) added, ' + rep.images + ' picture(s), ' + rep.prices + ' price(s).' +
      (dryRun ? '' : ' ORDERS, PRICING and IMAGES sheets are now hidden (not deleted).');
    return rep;
  } catch (err) {
    return { success: false, error: err.message };
  } finally { lock.releaseLock(); }
}

/* Only the settings sheets the app still needs (ORDERS / PRICING / IMAGES / REQS are no longer created). */
function setupCoreSheets_() {
  var ss = SpreadsheetApp.getActive();
  var cfg = ensureSheet_(ss, SHEET.CONFIG, ['Key', 'Value', 'Note']);
  if (cfg.getLastRow() < 2) {
    cfg.getRange(2, 1, 3, 3).setValues([
      ['statuses', DEFAULT_STATUSES, 'Comma separated. Order shown in dropdowns.'],
      ['closedStatuses', DEFAULT_CLOSED, 'Not counted in "quantity in hand".'],
      ['usdInr', 93, 'USD to INR rate used in Management tab.']
    ]);
  }
  ensureConfig_(cfg, 'productionTeamEmail', '', 'Production team email(s), comma separated — for JO checklist notifications.');
  ensureConfig_(cfg, 'storeTeamEmail', '', 'Store team email(s), comma separated — for JO checklist notifications.');
  sheetOrCreate_(SHEET.TECHPACKS);
}

/*************** STYLE DETAIL: tech pack, CAD, accessories BOM (Order In Hand click) ***************
 * Files live in Drive (folder "OIH Techpacks"); links in the TECHPACKS sheet.
 *   Tech pack key : BUYER|STYLE          (same key the old OIH app used, so existing tech packs show)
 *   CAD key       : CAD|BUYER|STYLE
 ***************************************************************************************************/
function sdKey_(kind, buyer, style) {
  var k = normKey_(String(buyer || '').trim() + '|' + String(style || '').trim());
  return kind === 'CAD' ? normKey_('CAD|' + k) : k;
}
function sdFileInfo_(tp) {
  if (!tp) return null;
  return { id: tp.id, name: tp.name, by: tp.by, at: tp.at,
    url: 'https://drive.google.com/file/d/' + tp.id + '/view', preview: 'https://drive.google.com/file/d/' + tp.id + '/preview' };
}

function getStyleDetail(jo, style, color, buyer) {
  var tps = readTechpacks_();
  var find = function (kind) {
    var exact = tps[sdKey_(kind, buyer, style)];
    if (exact) return exact;
    // same style under any buyer spelling
    var want = stNorm_(style), pre = kind === 'CAD' ? 'CAD|' : '';
    var hit = Object.keys(tps).filter(function (k) {
      if (pre ? k.indexOf(pre) !== 0 : k.indexOf('CAD|') === 0) return false;
      return stNorm_(k.split('|').pop()) === want;
    })[0];
    return hit ? tps[hit] : null;
  };
  // cut qty of THIS style (per colour and size) = order qty x (1 + excess %), first part row of each colour only
  var cut = { total: 0, byColor: {} };
  var seenC = {};
  getStyleTracker().forEach(function (r) {
    if (String(r['JO Number']).trim() !== String(jo).trim() || stLike_(r['Style Number'], style) === 0) return;
    var c = String(r['Color'] || '').trim().toUpperCase();
    if (seenC[c]) return; seenC[c] = true;
    var ex = Number(r['Excess %']) || 0, cc = cut.byColor[c] = { total: 0, bySize: {} };
    var any = false;
    for (var i = 1; i <= 10; i++) {
      var sz = String(r['Size ' + i] || '').trim(), q = Number(r['Qty ' + i]) || 0;
      if (!sz) continue;
      any = true;
      var v = Math.round(q * (1 + ex / 100));
      cc.bySize[sz.toUpperCase()] = (cc.bySize[sz.toUpperCase()] || 0) + v; cc.total += v;
    }
    if (!any) cc.total = Math.round((Number(r['Order Quantity']) || 0) * (1 + ex / 100));
    cut.total += cc.total;
  });
  var bom = [];
  try {
    bom = sheetToObjects(SHEETS.ACCESSORIES_BOM).filter(function (r) {
      if (String(r['JO Number']).trim() !== String(jo).trim()) return false;
      var s = String(r['Style Number'] || '').trim();
      return !s || stLike_(s, style) > 0;
    }).map(function (r) {
      var c = String(r['Color'] || '').trim().toUpperCase(), sz = String(r['Size'] || '').trim().toUpperCase();
      var base = c ? (cut.byColor[c] || { total: 0, bySize: {} }) : null;
      var pcs;
      if (base) pcs = sz ? (base.bySize[sz] || 0) : base.total;
      else if (sz) pcs = Object.keys(cut.byColor).reduce(function (t, k) { return t + (cut.byColor[k].bySize[sz] || 0); }, 0);
      else pcs = cut.total;
      var per = Number(r['Per How Many Garments']) || 1, need = Number(r['Qty Needed']) || 0, rate = Number(r['Rate (Rs/unit)']) || 0;
      var req = Math.ceil(pcs / per) * need;
      return { trims: r['Trims / Accessories'] || '', color: r['Color'] || '', size: r['Size'] || '',
        qtyNeeded: need, perGarments: per, rate: rate, pcs: pcs, required: req,
        amount: Math.round(req * rate * 100) / 100, ihDate: String(r['I/H Needed Date'] || '').slice(0, 10) };
    });
  } catch (e) { bom = []; }
  var img = '';
  try { img = getLineImage(jo, style, color).dataUrl; } catch (e) { }
  var photos = [];
  try { photos = getStylePhotos(jo, style); } catch (e) { }
  return { jo: jo, style: style, color: color, buyer: buyer,
    techpack: sdFileInfo_(find('TECHPACK')), cad: sdFileInfo_(find('CAD')), bom: bom, image: img, photos: photos };
}

/* kind = TECHPACK (PDF only) or CAD (PDF / image / any CAD file); one file per kind per style, re-upload replaces */
function saveStyleFile(kind, buyer, style, fileName, b64, userName) {
  kind = kind === 'CAD' ? 'CAD' : 'TECHPACK';
  if (!style) return { success: false, error: 'Style number missing.' };
  var name = clean_(fileName, 100).replace(/[\\/:*?"<>|]+/g, '_') || (kind + '.pdf');
  if (kind === 'TECHPACK' && !/\.pdf$/i.test(name)) return { success: false, error: 'Tech pack must be a PDF.' };
  if (!b64 || b64.length > 28000000) return { success: false, error: 'File too large (max 20 MB).' };
  var bytes = Utilities.base64Decode(b64);
  var mime = /\.pdf$/i.test(name) ? 'application/pdf' : /\.(png)$/i.test(name) ? 'image/png' : /\.(jpe?g)$/i.test(name) ? 'image/jpeg' : 'application/octet-stream';
  var key = sdKey_(kind, buyer, style);
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var file = techFolder_().createFile(Utilities.newBlob(bytes, mime, (kind === 'CAD' ? 'CAD - ' : '') + key.replace(/\|/g, ' - ') + ' - ' + name));
    try {
      file.setSharing(readConfig_().techShare === 'DOMAIN' ? DriveApp.Access.DOMAIN_WITH_LINK : DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) { }
    var sh = sheetOrCreate_(SHEET.TECHPACKS), last = sh.getLastRow(), row = 0, oldId = '';
    if (last > 1) {
      var v = sh.getRange(2, 1, last - 1, 2).getValues();
      for (var i = 0; i < v.length; i++) if (String(v[i][0]) === key) { row = i + 2; oldId = String(v[i][1]); break; }
    }
    if (!row) row = last + 1;
    var at = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), by = clean_(userName, 40) || 'unknown';
    sh.getRange(row, 1, 1, 5).setValues([[key, file.getId(), name, by, at]]);
    if (oldId) { try { DriveApp.getFileById(oldId).setTrashed(true); } catch (e) { } }
    return { success: true, message: (kind === 'CAD' ? 'CAD' : 'Tech pack') + ' saved for ' + style + '.',
      file: sdFileInfo_({ id: file.getId(), name: name, by: by, at: at }) };
  } catch (err) {
    return { success: false, error: err.message };
  } finally { lock.releaseLock(); }
}

/*************** PURCHASE / PROCESS REQUIRED (fabric program PDF like the OCN report) ***************/
var JIWIN_LOGO_PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEUAAABgCAYAAABL7L+sAAAfaklEQVR42u2ceZhdZX34P+97zrn7nX3PZLJNMtkTkrBkISYsFUEBBRcoID+t9qkttT+XturTVrRYn/60rbVV3NqiYkEU2QMKSAhZSEjIJJPZktm3zHLvnTt3v+ec9/39cSaIRStCFqg9z3Pz5HnuPM8953O++/IKHpjW/O/1S5f8XwSvvMw3+wMYQiMADbha/C8UAbi2BUqDFAjTxkP0OwpFINBKsq1mlNXRLAeTAXbH5yCkRqN+96BIAdqVfKH2Hj417zjMjGHXhfki5/PXU+9GzgrP7wwUiULJCE2pA7x34m6Otx+maV2IsSMpfr+qk69llnOyYh3CyaCFfI2/8Sa7tNYIU1KMJ+lqPYmyolhWCTnXz2jPIE46C4Z8neDfhNZEF7OM12xgx+Q8An5Bsm8ENTHFN9uXMFV3PsJ+7VLyJoSi0dJPicijw9X8eP4fUeIr8LOdkj8+8X/43tZvIvwGWr2+eNR888iHQssgNTMd7FnxTZ7TaxnyFRDjM3x24oN03PRVsDNouwBC/o5A0QptBagdegGn406uXR2g9LyNJIwaPnljhK7Sx/mX4WVkfZWgndcVr4g3T+6jQRj48wnKBzpoih3gZ+98iJ62Ec5rmGakV7PGfYjYkm3IfAr1OqTlTeSSBYa2KQRrOLl4Lp/aMkyJr8Bnuq5jeGoViwaeILb9PChmXxeQcwLFQCGEwNXe2xdebIqLQACG8AIv9V/EX6JxVYigG+eORYf4k8hukkfHqWioYEfLhzi6/BagCEq97ns0z8471rM6LnBVCFwFPu19ozRoA4wi2tU4jgGmBunMZjceEKUCbA8f5s7VHcyf2kc6N0wub7KSLkRqGiOgcLTx5igdSBQaA4lGuIoby57l60sepsHupynbxkNNX+PLVd9B5LNUlzjcuriPhcEpwIdAYaBAhmgqHOOu6BfRP76dVGwnkdp5pFM5Lgj3408P4xghhD495vHMSooQKMKEdZxMoI5w3wHe+9O3c/V7qsnGL2bXzgzzlzzFtlVhvtkb4uK2Xfz9xh6OzThclvs37Jq5uNk8+AUNfTsp5g4wMjJKy3t+n6ItyeczLKzPMm/kGF31KxBk0RhvXCiG0ChbcpP8IZ9b08/DR9N8bOZKutVcxm0/t6ya4KqGPEHrPDJ+lx+vuYvhtgT9J0YJ6xx+M06h2MC3G+9l6eJm/nLHCZ5PKH5vWwOYPtTUCFVzysgWsszkBULqV6jsqX/fEFCEVrjCArtAy31/ipjIcW3NfFpWZbn0tmuIjQzR9kI3J7tjTE84uDmHYLhAeUOAfGWYOSUl3Gd+ja/+PMoC/T02v62eawR86/4iTjbDZeED5BM5jhx0eazdIvGe89G2A0L8woZpE4T7Mnt2DuMUgUYLP2EjQ9QUyHv/ka/W/Sdvve0GwhVF7rtzD89+7RiVMT+LaxupMgL43ALTtstQLsGIexJjpeaSSwOEHIfHnrJZOAeu/qtPEktZPPRvPyc9cQLXdUiOT/PhL36A6PzF/N5Tq+ir3IzQOWwZosGIMWWHKSo/UuRnPZ44+1AEGowAVdPt/GThvbTUS+4/mua6ddWYYowv3voIvieHuaVuIYuWNINUcP1KmLcMdh+B9gnyymVvVw8PjPVSul1y8WaDHT8y+dyjnyZSLsBqRGVaUakT9D57CDU1hY5luS59Ox2XfALiBf44dwe3X1bKcHqGG7qupIP1oLKeVzvb6iOUiwoGKO09gPuzvyd2QYibNl6KsBz+6rr7qd49wW2RWoxskZ5nn8P+gyiV2yUhf5ZMXYzco3upGfGzfckCtkQjPLijix89nMBYUE76ZCv5A49R3rKQxItHqDhvI7vagjxzb5IL1ltcveU410TuJpFIcs3gt5jZqWl0ElQet2m+vIUGptiVrkUoF/0bbM1phaKEAYUUJ5a9l78d2MVN/c/RdJWPHXcdwNk9wXW+MKPpDA4OHYaiakGe8P6dVK/pYGoozfAaC3U0RmTfJH5fhOujEfyJBPvKLezUODpnY2TGmJ4o4up5PPuNg2x6VxmLlilWZH5ASft9lDfWEnzfFirDQQZf6KUmM8ZnOq7giktLuWHX5dxT9Qkk8dnSgjjzUKQA5VhcKx/hk3+yFn9nJ7nYCId2jHGBlOSVS0ZoTKkYdTWjhxxWrijB9dUyM5Xl2KFpIsJAG4JUMU3KyeDgUlJXgpNKEi0NkRzLEmlZx1M/HqZhIkHg3/z0W1FUtIq0ASlfmmzZXqLrq7ny5lX8+JPQvd/HI//yJB2BzSxYNM1IUlMUFkL/6kL3afY+XtK2ZMcdlESO0ZOVRKLLURM5SrQmoVxcNAiotDT37rDRQYdLNw1w179nKO3SaKFIuw6WENgSJhQEqkzIxQg3+Zk8nsVe2kz7Nx7nRmGxWJQQKK+COTVQXQmuZmImzt7HT/L1e+5n3i1r+YsvbGVELeDaYwf5VMs/8US/5Pr+W3GiFQj1yoz6tELRGvBp7mn6OM7J7+D0dDFvOkc4YJLRGildlBbkHRcHELZLljoWX74c81t7CLgzpABD+jE1WGiSQGWZjeGkUEUfZlkdB3YnqOuPETFDDBUyyOE0/uEBDCThWy6h7N3v4pqhE7x153H+7s5n+cTgDP9877WsrJqk9Wu30zdWgb31/SAEQr9Siwze95efPZ0RLMom2biWvSv/kKlDbVxd04lR0UD/C5MEtSStNZHLG1n2ycsI19YTdPuZv6qMXFxiUIdyBcV0CqTElooTWtCwLUDLYok/AJnwPH76d52sGUthCQMbhW0KpjWkbghQ+vklWOdVkFk9TW5vJ1dl63h234vsS1u87Q+3ct/TNg/N/SiRlauZyJiIX0HltOc+UiuElizve4gvb3yR2rXnsfWmZm7t/hzG2+aRqDOxLnPo6HySjdf7mbe2iZO9U2x871yuvnM+FzywjdjyehLKZtR2SAuXqC+NGa1BEaBv2KJwcJgSYZBXiryAonLp0Q4zC4qYXc9gDf8n5vizdNSPMNU3zMei9Rz612doPZjiphsWcNfG59gd+ihNsefRZgSp1ZmFojVoyyI9laR/VzvTHfuJHdhFYue3aE9MMBiVhNMx/IUZnv/B41RXpYlENPnO5xl54FHE1HPUXxWkT5o4G5diL65i4Zp6gktvIu+vZfe9PVQpRdEQZIVLRhXJCsWg1Ox6Nk2+6EfW1FOYztF7NM2wO0OZq1jlujx09/NEyyp57vsPcPc/3EMub4DUr3DPpz3M19JA5KYZvOBm/u7FF5jffSf+RY3c8We9zC8K0oZB+2CAB5+V2FmH9RdnqCrx8d0fFXEdH9evyNLZlcW62OKmp2/kLUeHSR/bQ/bEE/S0dnL8mRznI0jYNi4WvvIqCokpZnAING9hZHSMvf9vL0e7FOIZh3VCMFnM0AS8cKSfdLIGc+XFPGrexGTjhYhi9hWV/9NrU17KfTRYQcIqSf2u+9jwjiX07UxTPZ2nREqefdHFmVYkMy6bNggCEcmn/yrDoTZBbkax58k8VrVmgThIpreVhk3XEqxaSGEmxpVf+EsmuztITOZY/ZVNLP/r7ajFVYwMTPPJ721i78MH+OyXUwS7FdtdTVBIXKEZ1w4n60Oct0GzZG4pH1o9Rtuh47SXXo6hs+iXKc0ZypI9g5uN1OHzwYvPTfHhf9rOD971MJWOwyZgBMhV+imrEMzYFivLNaGEzeQTiq1C0XpI8MG/yDK3QvFZvsHJhEm0opLF6XuYNBOUvD3EvDnPo3v30/LW1fhaltP7xA6e2+vj8pIqLpqZwi9NMlrhCpgWEIm46PQkOx86RLRo01P9BVig0Tn9S4bkjEgKaDAMiqk8V8W/y9AL44Tm1nLhjecxPpShoEwqWiK8+yOC7qNZiNawcuN8hn86ygoU5Wh8StGbE7zvg1Eu2iD49JcVjz4zw5Gf9vPcU0UaWwSJNLQezNO8rIpoucHwkOCS96xny83zmB5XZE7EMEwDJTUvuoLGKypYWjlNh1jP5+q+zNEVNyNU7uyoD4AwDMhmecfMT1i3MsBE3xj1Cy0u/dP1rHv/OlZsEowP2cSmLKZHcizbUkvVkhISqSBF08C/tprrP1bLnLI0ZiREPqvZ//MctSN+WqTiaL/miZ2KF9tc3nFdCaadx01MEYofpqwsTrHOz7EH4xjKYdJVTDQGueJ6g/anY9z8gYVUpfv4cfESRNCP1uqX3PKZUR8hwClCRT1d4TWUjz1BUQbY+3Q7R/cMEynxI0w/oZIKlm2ZgzAk0ycdll7exLY/W4qdcxk9McrhRw7z0I5pSitNbvjUGkb3DBLpmKRCQTDtEgLCq4NUlmR58J4k//SDNNvO8/O5v/Zx8OeDGFfMYevtt7L/7mf5o6tr6Gh9gW8nNiO/8gyPOpvQ14R/ZaH7jPR9vBKCD39smO+X/yOb5zk88zd3M7fQSPCSKQJzDPIqhF2QdE0EKNEmmUSadDKLz9SU14aJVNVSXj+HcNhivC+Nv8SkcVUDO//lOMlnujGkJrqxjDVvL2Lm43z+U0n6Bh0WtPjZvFly9915PnZbGde8fxl25TKcGbjhb07wyA1PYA23Ylc1QTgMjv1ScerMSorWaNNCJsYoPnkn4xdV03nIpXooQdVGh5YNFoSC3LVfcIeq54ZFTdz+lm3kplMkT8aY6BsiNj5K14F2CvkUZdUBIvkw2SnB739lGzMz19C9/0Um2nbyw6/2c8WtTVTm4jQZMNlrs7tLMF9oPv+lGN0de/jI502+98/HuCLosCvfR3LxekQ+hf4VQM6cTREC4Rawq+bTliyjuvUpZLXJQIMiUutn/HiB3IkYu/fEePC+XhbLUd6+pQUzVEa0ronaJRfDcI5V+zKsLFuM2dxAwRfi8FNH2Hv/U0z1dhApCRIKz8HOGVTN9REqCTGyP06jUsxB4UOy7NoFbL1lBQceOc4Pfi7Ztf7PGFx4GaKY8fI0Ic5u29TQDq5ZySVtd/CnyX/EFZqKcoNQaRWRaDn+gJ+AqYgnHcpCGiNsos0QwpKUzK3jmx95jPLHT7IOg/D3PkbzTe8DNUi8r5Pugwc40X0c3AJ1c+dT1VDOyk21PP9gloFdg0QrBbUbJMmpCcYOdrG3U/Pt9zyDs2wDIpX4jWMaZyZ4Q6NEgPluFx/1P0jAKEOXtTDtzGUqGWLopM3QQJKBwRjZ6Rgz02m0owkGLAKGBdks1c01tAs/Q7U5ll0RJBCxMawgoZqlNDQ309LhUq00U8VuOna3cexgjKWbm1iyrZl8foh99+0iPTBM2CjwjaqPE7/kFozkJEqar+L+z4CkGNrB9VVy/a4P8OH49yhUlrNmpUmwJIS0IhRlObZVQcqtJDFjMTFeYGpolHx8FL8dpyxaZN7CKI1L5hKobEDlCuTTKYr5HOGqMMePxDjxJ/tZ3NTAhW3vxMpP0vrwPp4/OIkVrWXZyrmYwiGQOs6f313Ojo/uRvgDaNf9tSpz5r2PVmh/mPCRn3HNAx+ksjhO0YHaSphTD/XV0FAHdQ0+quqi+GvnQHgeGVXHxEw5QyOCgc4xTh7vwswPU1dts2BxDXMWNOILlRCbyvLA908wx+rl2g9alDQ1IcpqcH5q0dp1hMePxqleVEujOc6toe8wefEtiHQMLc1XKelnahRDK/BHIHaS8LGdRHoOEBo8Rmiii8rCOA1+m6ZKaJ4PixZC3Vw/dQ0ByqsCiNIqKFmEMhYylqyip9+h52g/yd6DBNwRFi0pZen6ZsLlVeRiCVSqg8H2Ifo+a1J5cQmXPPoxHviHn/DR9k1M3XIHynF+q6bYGZ1PEVqB6UMHQrMBnY2ZnsI3MYBvpIvw0DEqRtuoip+gzh6jyp+jcY7Flgtg3nyN39T4AwaBskas+o1kI+sYGA9xbG8nAwd24ct1s3h5lKUbVmKVVPL4D7tJDbbx/uslzx+Dt684hltRhyhkf6sZuDM/tKO1BwcNQqINC0yf9xGA40B6GiM9hTz0HO8/9mnuuC2IkxxB+k0Mvx+tvGqc1IpwZQ0li7ega7YwMFnG4V1H6Nv3JCXGBFvesZJgdA5f+tpRftLyfxnb8E60cn7rocAzP4ohBFq8rOntOuDaiHz6JVDCMnAXLUX0HGX75goi0RzpGRefoQkETfzlNahoMzZR0rEc/Qf3oDP3E6mq4aotb8O5+kN0duc48OxTZAae4fHt32Sk+QpkOv7Lv/2GnWQS3pzKyzuYQjmQUyzofoK1V0fJJ4aQ1uzfKYFyHKRr468IE5i/lqrNN2LnQsz07KH9wX9H8F385QvYevFqJi44n/hAJTL/yzWSNzaUX3kXFsTjrOU41bUSZzSL6TcQUiKkOFXSQzsKnbdBmFjhFsoWO4w8/UPufyRDZaSX8BO9LKqHqsXXE2/YgHTyr2k049xD0Qrlj8Kxg2yunyFoOeQESKkxDIE0TKRhgGkirABYEbQRAe0w3fooRzuKfKnpDoJDbSzJdBMXy+nedjnYr30Y8JxDEVqhDZPK/t1sWBeAfA+GBUJopCEwDOGNlRuG58nMIIgyEEmmW/fS5juPzLs/TVYo9qSyEAiCnfNsl3htY6PnfOJaCAm5IssSh1g4P4KTiiNNAVKCNNDS8P5vWGgjAKIEaZaQT/aQ6IvR2vBWRF5hJKcQwkXkZ0C9diDnHorW4AvASA8XRU9SVp4Dx0UIAykkUgqENMCwwPCBEUKLEhCS+LGdjMSgtfpStFS40kTDf9s4f1NAEVqh/AGs3gNcuEhiqSTCACkE0pAIQ3plTcMEw4eQfpBlwDSJo/vp8zcxVrYK4WZf14LCGwsKGpRg3skXWNUSwk2fxLAMkCANiTRMz5YYFlg+tBnFMEsppk4wc3ycFysuQwcqkMp53dLxxlEfw4TpadbRy5xGicolEaZv1hWfkhATTBOMAFAGwiJ+4gDJBLxYc7nX4TvNMfm5g6IV2h+EoQ4urJkmFEh7NkYIhBQIKZGGMas+PjBDaLMayDDdtp9RK0p3xYWgcqjT/BjyXNoTbfooHdzPRUsDiELMsxucAiNBSJQ0vJTfiIJVhVPoJ9nZR2vpZgrhuRhu4XV5mjcWFCGg6LIs2cqS5hBOJoYwfIBECgMhTkW0EiEttCzBEEFm+g+TmtQcrLwcLNMrUfA/Qn00WD6YGOUC/xAVVS66mEMID4j0XBBaevGKNn1oowLIk2jbw4Rrcqx6G+jCaVedcwZFKI3yhZADrVw410YaSRCm50AkaMOLTwTeRxt+sKpQziCJ9m66o6uYLm1BuvnT6orPsaQoEJL64X2saQlBIYY0LORshOupj0RIzyULI4i0SkgPtZIatjlSvg38YS+7PhP56TmRFCnRmRxr8+0saA7hZpNgmJ7z0bN2UwpPYgyBMsox0cTanyeWgYPVl4JwX7ET9OaVFK3RviCM9nJR+TShMoWyiy8FXxqBFtKTFGGAtBBWJaiTxDva6ffXMly+1psWOEO3f9ahnFqQDA6+wIUtPlAJhJZorbzRMHEq1pUo4blj4a8hP95B/ESO9oqtqEg90i2edld87qCgQcGiyYOsWF6Gm5kEaaFPhaWal1XmNMqMIn1+pjoOMDMDh2ouBVP+4u//R6iPacJ0gvX0UDvXQmUTXsyi1ctOs5hdrpMCfNXgxJhsO0LMCtNdvRFU/oypztmHohXaF4KRDi6qL2IEc94ci/Iq/kIrxOw6gRASJIhAOW78BOn+LN3RNeRKms+YKz4nULzQ3qJ88AXWLw2APYnERGsHpdXsvKmatccChA/ps5jpaycdh0OVl0AghFDuGb3PswtFCCgqViYOsGR5BToTQ0hj1shq74O3gqtxUWYYoTPEujoZL0iO1Vx8Rl3xOYCivQbY5CgXRkYpnePHzae8TT+tkEqjlQKtENr15tB8EVRigKnjMfr885isWnPaC0qnDYq3VK0whcIUGlMo5G/o1gqlUf4Q1sAhNs23wMqB4yCUBu2itetV9pVCa4VQYFgGM319JMcEbRWb0eEapGOf1oLSaYFiCIVG4rohHCeIY/tw3AAKP0II5K85D0nMhvaNo/tYvbIG8jEQFhqFVqCVQitPhZSr0MJCODkmjo+RKmjaai4BE/RZOJnBfPXS4W1tuk4EU8Y5P9xBiz9HQFvEC9CajdDl1KN9pUiRQen/wlsakM2zJnuMuS3NqMwxr7qmXcCTEq211/QCMKAwNUFqNMmYKKW/9iJQ+dd9jsFpgyJRKO1DCPhQzePctjDFyupKMCXoAihNITfDzr42Pt05h4O+jRgyg3sKjNZovx8Ge9hYOYOvJoDTl8YQ0nvzSqO08mBohXY9L5Qen8Isao6F1lIsX4i0cyghzj0UgUIRoNqa4T/qv86VSxeDvxY7sReV60SIIgTrMQNN/N7KRWypaeX6nePsCLwLSQqFfCm0jwwc4PwWrxovFC95m1PnIGivR4EwJE4+TyrtkHPgUNlbIOCHXBrkmc9h5W9SGSH9RApx7jX/gCubq8jjYPd8AWP8J/gyvVjFGKYbA+GQS7QSrJ3PXWt20xh/ztulQb1UtW+e2M+KtY2QG0VIa1Z11EsfiQvaRRqS3HQaZZUxkvLTWbsVcM6413l1UGYXJT/a/4dsr4eCaWD1fAXTziB0AJT0giylENrGV7qC9NA+qpsa+Yj9dXR+NmczTEhOs070Ut1chZuZ8o4IUhrcWVesNFrNxiiOTbZoYYoAh1NzSNWvQTo59FlQnf8WitAuyorSNPgYVyUfJVVSjm/0RwhteA/hFLw5Ez0bXbp5pFkFRZvCdJrLIwP4pjtxpR/tCyBG2rmoXiFKJLqYAwVauZ690S87M0UK7HwRJ1BFdirJC8ELobQKcRZc8W+EIgXgCrYM/oDEFGSGuhCq4BlBx/Xcp3I9QLOWQZPECIRJTQwQEoLKZJeXARt+qgf3sH5VDRTHwHFB/SIuUUqj8NwxQlMoaGSwitH+aY7UvxWkRp3FQ5Lkr4s+lTQhN03D+CHahgzMQBC0i9AOQjkIZSOU60mKKoLOIvQ0qBSFmTh2JocsZL3muO2wInGIlpWN6GSvd5aKssF10a6D1i5CaYRykcol7wSQVpjW0QAn524GJ3fW7Mmvh6JBCxOjOEPuZJyk6yNaakLRE3eUN6KlXduLSp08wskiSFOM96KUj5GRHBkiYBkwOczGaIzw3BBuesI7KlW5s1AUuLO/icTJKUTpPDJjCfa4K1DVC5BO4YwVlH47Qyu8nCSXVzjKOy/Jc5ka7bqzD+WgnQLazqOKM1DoY2awG9Mf5kB7ipna5SAh3L+Py9eWgBFDzOY4wlXgurNhvad8QkK+AKJ6CT2HjtM65yrwGaflnKXXD0UAysXxRSgGouRTBbIZB2FoULYnKbNQcIroQhZD2mR6WqGknvHONp4orEI1NEPO5YLJnVy4aQ460Y0w/KAUWrsoNWubTpUibRcdriGdVhzqdehruRqcnHeGwjmHgvA6+cEyklWLycVdBgay4De8N4uaTeBccGwMwyU3PcVE3wiFVIZvf7uVQ5fdjvCZ0H+UmxtHCM5VuPkkQggvaFOnoHjhvZQSuwBu3VqGdu5lV9kVOLVLkMXcWVWd/159lAK/j+MLriBkwr6nx8AUs45Ge2/bLSKlTXwoQfdjhxje081X79jP3du/g71kPToPbznxfa57RwMq3Yc0guC6CFeD64LrJYEIgVN0EdVNTIwVOHw4zu7zP46QjhcHne2K6a9lIgxEMcPxle/l4he+xPDhMdr3R1m+wsKZynkDw4bm6YfH2H/AJkeY3dZ2nn/vZ8iu3oYqCuoOfJ+vbJ+gpKGAM5PDAK9Ool3UrF3SgLIVMhBlKrSS9u88xk/m/gGZhZuR+ThKGm8cKAgBdh63uoknz/tzPvzcJ+g4GKOiJExdmcDJOIhIGXVbtnPPeD2tqz8MTUu8AGd8gnXjT/KvG3ax5kIbNzaIEShFuzkvI3Y9d+4qjbJd/EE/sbIVHL77OXbE57Pvur9FuOkz0id+dRWB3zCGLtBoM8SF372BDzr3s+HGrTQG+igPZxHSwFh9M/3Tq/lhp5+uQjk+J81Gfw/vWjxMpKQXJz6IEakEaaKdAmSmIZXATk3jFrJIy8dwoYHDj43ws1aTu276KYXmDYhcyhsCfCNCQWuEYaCVYN23buQD4d1ccdvbqA4MEci0I0yB1XgBRJtB+cBNgZ5Ap4Zw0jMYoVIIRtBaeJMFmSRGbhqdj5OcgaMnwvTsmeKZ8Wruuf4+CssvRmQSr3oN5dxAma1xCMNEGz7m3/MXXN37H7zvQ5tZsaWeqO5H54dw7DxgcGo3QWqJlBYiWo4IlSKEC04GkpPEB+Oc6Cww0J3heIfmydK3sPOdX0fNXYLITJ9TIK8eyimJEaDDZQQP7mDdQ5/hyqpBtr5tHcs2LaW8WiJFDOwp7+HtPOgi+EK4yiQVzzDRM8Vge4zhzgzDI3DEmseu8z/Oya0fAhNEPnPOgfx2UE517pSLDpeBXSS6534WP/9tVhc6WTc/xMLmMirr/IQjAily2DMTpMYTxMeyTJ6EqQSM2gE6y9bTtuzdjG14D1TVI7LTs8ngG+N06de07yOU600YhUq9+tBIJyVduykfeJ7K2HGihTh+VUBKgWMGKAQrSJQtZLJhLbH5F+E0roBw0FtvsQtvCOl43VB+CY4QaF8AfMHZpSYFxSK4thfkSW+mHsua/b6AKOZAud4ujhC80a7X9YpecpnFAqKQ89RLePMlnFpJwUHYRSjOFqaFnP3+jXto8v8HhjRPBuH+XAIAAAAASUVORK5CYII=';
/* Stamps the Jiwin Apparels / Authorised Signatory / Pavithra seal next to "Authorized By" whenever
   that field is Pavithra (any spelling/case) — add more name -> seal mappings here the same way if
   another signatory ever needs their own stamp. */
function poSealHtml_(approvedBy) {
  var name = String(approvedBy || '').trim().toLowerCase();
  if (name === 'pavithra' || name.indexOf('pavithra') === 0) {
    return '<br><img src="' + PAVITHRA_SEAL_PNG + '" style="width:70px;height:auto;margin-top:2px;">';
  }
  return '';
}
var PAVITHRA_SEAL_PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAJYAAACQCAMAAADdhmLNAAAAP1BMVEX39/oAAAAKCpyfoNhnaMtZWrQKCrIAAHwFBfkICbILDLIAAH82N8FcXOG/wOb///8AAAAAAAAAAAAAAAAAAABOqbYnAAAAEHRSTlP9APz9/f0P/QKiXgL9Bf0B/uMx8wAADqpJREFUeNq9XImasyoMTQmorZ353/9tr5CFBHDpMtfvm6W21WMSwskCcPvoWKbtWP25dXDu1QPeR7RM68L/T/VY/tVzy/8Na5MS3/vxuN+DP+73xzQBCW75/2CtRUU/06PF48E9HgXTW8jgdTnl2zhIMeJ2wHb8bn9jTAZakdrryOBVQTlMsYCR49f8jwruXr4y/SGsZTKYGBLGZ6e/iPIevX5M9N0/gZUtaiJMT7R3zWokRW4aVJTPgps/c38RGFz3BwrK3C56NaoKBXH96P2xXWP9MqyivgIq6o3GiJx9KTIFtnwT1mpB6c0uHIoMRZXT12Btj0ig8uVhD1MqB46RGWCXvAVcERU8wrNICjvdobzCmU09FYxoP4VBvx0elwR2Cmub96ZQLxu9hEKYsd66HPkEYTSiQ5JYZIF9DGu6/XvkW+iV3RHoLbkx/U4Fbgj+Ifjr+ex0bvlwhmoK+qC93VStMSyg1/l3NzKwwIyXFAknCnyo/uLAzKtzYmlsv2eChQLRHJE/cq5IOBQV3PUGYk3oYaXE76lthUZyBU5qBXaCC45GYHEL5WJoDciqJcnNFRaSaSVMYlwp1JFRBEaKnN6CRWZVHu7ZjHTzEoHFRb6D/EIwkmP3O6P9Dp7hgn1UbFbeqpy4Uq84cLCwfgjNJcpVN1zry7AIFVYFJoIXLK5kfJWBxb6CJJddW/K6j/nVMS44RhXlemkOArM+OPnMYlUWlkVRzMzLuOAGeB7pEU5QxWZuweCVVX3DNiHK6XlODhUap5su4YJ9VGBQWTuq3kju9tuyLXEj9CV1IjIJnOOCoWfwqEIzvMyjH/ObMDdPUuww8L9sX9NFWOQZVINoVYUO19kRC4bitpJ9Motr7Fdh5EWdXSW5EKki9bAKhVdS73mPCAllhNQBobjWC7CW6aexdsXFukh2pqthhg992qkT63XEtyKP32VADKFX4T17KNGgxRVbSjAMMyRC9GSLnb1zMUiD9D5QI4zM/UkjRcdzfU6jPjwKMwrip5lwzFQezBSbn39k9g0soqKkd8MBKq5UPfVJ6INWuOxzSY7q1barlNhg6iI1aBMM5Vrq2y0JSJW5x2uxj/1YsSjL9kM1+16NMDYsNOauXit5/nspIKsfTY63yeXIYDo1Qu+xisrLJdHjinZcXT0sr0Xr7SWwLeY1NewefC6tMSx2M1h9/R59PlZlnYuSZbJY3h2pEXoVQrCkKXnvGV8SVSMwDHOsM6V6w4EawQYUVoV+RDsK9/rx1GuiPK/KLt8vO9V1GcMqwqqjcIAL30QFlbdZBbK7F3GNYRV711GYUOh5sjEXwpsHCi5lbSy7fKq3ejD+/V7sPXr6hpa7vY2q4kod82erd+ICR/1i+YjJc1hH/RGqiosApTnYqbwVFwyF5bgoync/QgV1niVchkTTHGQHI3jLejZTsWVtH6JStpuCCbKFiTXiAjsMRViqt6Rply+gAhlOooGkM6RY19TAWmUYlm/PlaNjCuZBPz0YzxyqNnQwbgEHtNIqBo8gkmmnPbzK3s9xsTbQx1Cx+K7K6xnWUgZIrGlhnwT6FiqdQlJ3h1hmxtXBYoO3mT0rMiE6hW7lIH4LBDHFnMHFLWjNP9lkUxlZJa2bgpzJBl4GGqV7n8YBIXQ+QnwXOO9gTaqJnegRsQR/2euU6H8OeVbZbsJPAkLjNXaj0DCh5ixdtGLFSD5idUpkHTbpWp3jfw11S4XtUPoqYaQZLUsj+9vIE8v2J20TF78RM8YocQoOfTNrcTKwig4RKI+V055CrWY09gAy8RcSBvUX3R0JSyrTckEXISXQMoaMQxVX6zysFkHHIetQDR7FgZnYvocVC5aAHaxf9polp2NhqbiSmgsaLU4VFpkW6TDNXaYhilYxlexfA4spQYEVC71DFh7HzWRo6rVYXHPy9NlpESQK03GIyeOKwQ6WdAorKAI1+WClJeLCMrDrnWgsGiU6X8rQahrWBCuJBlsDy9sWGZXmJ5k5WlhQ6S8No8SJf/Koq4OlpMyMxGR8lhjdGSwSZ8EDwGLUuBqtv2FPEjgqdsYFOk3HMAwMbOoxmwgLJHWwsoMAByvRG9mRYg0phJWiwmLzt8YF7LViM8HUNE30CeKkU2aajYOo7rScF3eaxJ06ty565URwdbRxg7UwLGYPniNIqqCqtqRGc/oxawr5/+1XpLfoken9KGdyVp+qjClF9FMsSuGAxC6eC8TkxZm2sALoHPblI2oWjkMGjmCqzcOOxXNW/XvcoecRbR2BbJ7IjYF1RERatK7wOr71Ie2qvM5c3wxF2BuIdtC0ZU0x1u2ayIEbT3PVi0hAV8ZA8erY8beUYjdbX4LVkQo3tUjxAjX0EyY8h5rPSigf2bmswnqqhwCaEZ87tu1NqwzYOYGZ5FDTdFETt8iuK/+UYZa93Zz6WKO/Waywbrdyu+cVWKEYjcBKpKeEkhkXJ1syUHNi7JD6+jAOEygFFo/E20+BBTtybZKO5PgQbQKNbx91Sipci+BGUe3RSLKUq/hT2H7/NN4U2yHT+n6GxYxFT4KBJdLSonIcDMUxrBvDWhpYYf/ryBOKgZW0cAdcjO1hxdb92RqXyTuXq+zAsvboYSUq2jcmDya4SAaWTN5hEGea1/VfoTYMCzwsG5S0M1KUkWVhJaV7bEsyEtE4jz1YJthoYHnb8vXV2Lc9JOO3UNiDuFiCVbhz8VunsAahNY/EfwrL5o/iwLaMl48lKytSC3r3OAfNUp9Jq8mCksn/u/V+KzVZrc7k60R4MiUiHs2QIcCoM8GMxE2L98oCY5Nri/AnR42mTDLC+a1mToyObfwJr/Gct7paP/l4WG2RXOtd3zmw8fINrKeFZflWsMGADRK/dPg5kScJbNMQLTtFio3C4VT/rWwlh47SsOT51sJcnkLe6Pqh8G9g1SiPo2VoshAGVjy0zb+B5edb8qaL5iD6OBH+ciju8poN1qKw1kPWjM5HonOX0lJtcub88+t9KO5ftMlTavh6OfTBpilL6GpJQs5a+hJ6w12ebY/SrgraYH+yNn+gRbRZHwuLkhMdrOoC276IuBNkudRITwTHAucklpYqUeIJysYg1kY0OQf1KdKJDrtE0mFI5vgh3bVQg8h/kbPB2nDpYRELnEMaEr/etODW5k7xTItRBSQJZcngSG5c4sUYnLSYqB676C5JSTmbPc9VLSGqbWmehQhgamEFG2GDxJaHwhqkdA89F5iZKBcI5P5kV7Oa+B4siivBVOtwX4cuAX6sRZNqDtRwjiYHpFX6oW3ZDEA4dNCDcsGxFm3HDTpthlNYUIMi5Zm4l+zoiytHjr6p+YADETTE34dVQpFkyhm7Olx8KUrrGI24ks2kV1ima43z5GmeuUpFJZntdXVzfJodK+wbfFu4y8BGRo++RCYzImoFT6ZJ1MwqZ1Klnrf9wsSnUau4HbgwLnOyR+2MvnZRfok6PKnqHPpRtd2qNi/CzRWr+yy4lGbiF3lW8cgJW2E9elgDcdUc0bdq6LUVjCpejbCmQXuGcNSmSVrj0Ofn9Nn4l2RDi0GvFNzG4sqLKrTh5zv9GcGlOYLtfSNHso5af5bFWpcykaTt3B/iElRUYp+TNbWusazt39LGEXFRGRXGz3EJqu2CxTZq5pI6bO7/poNut+hcvRC7GT/pdbNZrE1MXBHc1BBs+9Zt2u0NfHirx2JfZh3Ru7hqbi1rj43KlG+6nkXwS+uCdRLa8l1JXHxLkeZbmaj6pjJu1HWtgYO+U1VjuQIXYe1Kn/i6qIQZJQQm23Nyc/TjoMGzWn3QDG7SVj7UAfmawKqoModkHaBlX4PuYWgXvNseT5JVyoMHQR8wviIw8+Fs7RJjomFfo15rGCx4qE3NFObMudoa5jC61ymoKlq29jKy5zrVDjvTYdTFX0cj2VXSClOsjuMUmO9M51wWxQK8GJVTvvfTfvkN1481L7GlJFG9n8dPBGUXKqFqz56l1Rj9qgcYrxAxBH42MWsQvqo3Hrby8wLnjrdRkc8OQlojcmEtBjnVp+lDikC0N8xtrQuj3UuAl6G3i0SQ17wmKdCKkWZOsrcCCXbWREXL/DRSzdCSp5b90hW/ZIS7JVItEMxRB2HcW38Ew7Va92ddvMJRcQkikqi3NSuV1mCK1uXYjpkwqvvlVVF5JbxdVCOtfFwvj8XC5qaJbmBgCZtwIDaJIEK1XF7atqxwd0uQeEkarQmbpUPlsL5iM2GollDrv4QK1uWFhYDtgi3KMiZ2PMVeSsVuHxqxYuBaMfLcY/1F3FuodbBscuGFry4S0zU1xdZyuDzTeA+p8xO5oijtX3XR7lVUu4tMDS7DSYxKSZdS6pxD3/2FzD2SLko0MXEsi6uXl5fk0q4G0VJlxGo2c+03zX+xW4UX2BeUj1prj7xi8j69sSQ34/q5Mw/0099vWYFrdFl6yebQOwbynv5NWSd8/3eA6nC5dy40utXCxkGKiQTpD5mTa7CqORsE4+bMqup1eXNx/DJeWy19KVQP1pVhJgLUAdJa3FO6Xh4nu1XAlQ0OGg6PdQigtpYm4cURrdFjx1PD4RC8vPFCt5uADWdCbQ3EPYfqKD1vVHG2Rcv5NhW3uk1Fm5TbSOaMvqfUd3O0KzplW4/T7UZON/VYWGBxwJWRA/dwlgBr9vQ43wTlwhYoIjC6bs/7cufk4SKgWEFd3AHl0oYx62o2jAkvBoroQF3cL+bi9jq65w9eYPEtJt2S5/ruOq9sRvTw+/6c70X05E+hglqvbsR1eeumfEUGVgsZcYfWyIpr1NpCkdT13cHe2+jK7mGVt7TSGENXybMwzQZcL22O99K2YPlxl4fbOqfGOl2IIbDvD3gR1MubqC20iVq7LVmvxKi7uy3vbO72jS3nshJBdBjdZnjLG/u6vb9BX2lVmPodA932fBNtLPjWzoFvbme4rnU7ww3d3WxlmAFNsk773X0WP9n8cbV2/PPzs9i09SdbP34CiyPKvC1mhVdeT8ty+/CA27eO5XMwfwHrq8d/uYB0fh8HtNYAAAAASUVORK5CYII=';

/* Whole-kg display for the Fabric Program: nearest kg, but a positive need never shows as 0 */
function kgRound_(n) { n = Number(n) || 0; return n > 0 ? Math.max(1, Math.round(n)) : 0; }
function kgFmt_(n) { return kgRound_(n).toLocaleString('en-IN'); }

/* Shared pieces of the Fabric Program (Purchase / Process Required): used by the Fabric Program PDF
   and by the combined Order Sheet + Fabric Program sheet. Every kg figure is rounded to whole kg, and
   each column adds up exactly as printed (totals = sum of the rounded lines). */
function processReqParts_(jo, rows) {
  var S = function (v) { return String(v == null ? '' : v).trim(); }, N = function (v) { return Number(v) || 0; };
  var E = function (v) { return S(v).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var first = rows[0], maps = getProcessMaps();

  // order qty = first part row of each style + colour (weight-only sample lines carry no pcs)
  var seen = {}, orderQty = 0;
  rows.forEach(function (r) { var k = S(r['Style Number']) + '|' + S(r['Color']); if (!seen[k]) { seen[k] = 1; orderQty += N(r['Order Quantity']); } });

  var lines = [];
  var yarn = {}, yOrder = [], yarnLabel = {};
  // Grouped by Counts + Yarn Colour — two colourways on the same count still need separately
  // coloured/dyed yarn bought separately, so they must never be summed into one purchase line.
  var addY = function (countsKey, yarnColor, v) {
    if (!(v > 0)) return;
    var k = countsKey + (yarnColor ? '|' + yarnColor : '');
    if (!(k in yarn)) { yarn[k] = 0; yOrder.push(k); yarnLabel[k] = { counts: countsKey, color: yarnColor }; }
    yarn[k] += v;
  };
  rows.forEach(function (r) {
    var ly = dcxLycra_(r), kg = N(r['Fabric Required (kg)']), yc = S(r['Yarn color']);
    if (ly.pure) { addY(S(r['Counts']) || 'LYCRA', yc, kg); return; }
    addY(S(r['Counts']), yc, ly.mainKg); if (ly.lycraKg > 0) addY('LYCRA', '', ly.lycraKg);
    var procs = [];
    maps.sequence.forEach(function (p) {
      var rc = maps.rateMap[p]; if (!rc || /^yarn rate$/i.test(rc) || !(N(r[rc]) > 0)) return;
      procs.push({ p: p, loss: maps.lossMap[p] ? N(r[maps.lossMap[p]]) : 0 });
    });
    lines.push({ part: S(r['Part']), style: S(r['Style Number']), color: S(r['Color']) || '-', fabric: S(r['Fabric type']), counts: S(r['Counts']) + (ly.lycraKg > 0 ? ',LYCRA' : ''),
      gg: S(r['Gauge']), ll: S(r['Loop length']), gsm: S(r['GSM']), kd: S(r['Knitting Dia']), fd: S(r['Finishing dia']),
      kg: kg, fin: kg / (1 + N(r['Loss %']) / 100), procs: procs });
  });

  // A process entered on one row also applies to the other rows of the SAME Part and SAME Colour in this JO
  // (WASHED NAVY is in two styles: if the Dyeing rate is on one, the other is dyed too). Without this a row
  // that simply never got the rate typed in silently drops out of that department's table and its total
  // comes up short of the knitting total. Different colours never lend each other a process — a process that
  // is deliberately colour-specific (AOP on one colourway) stays exactly as entered.
  var normKey = function (v) { return S(v).toUpperCase().replace(/\s+/g, ' '); };
  var sibKey = function (l) { return normKey(l.part) + '|' + normKey(l.color); };
  var sibs = {};
  lines.forEach(function (l) { (sibs[sibKey(l)] = sibs[sibKey(l)] || []).push(l); });
  var inheritedNotes = {};        // process -> [{style, color, part, kg}]
  lines.forEach(function (l) {
    var group = sibs[sibKey(l)];
    if (group.length < 2) return;
    maps.sequence.forEach(function (p) {
      if (p === 'Knitting' || l.procs.some(function (x) { return x.p === p; })) return;
      var src = null;
      group.forEach(function (g) { if (!src && g !== l && g.procs.some(function (x) { return x.p === p; })) src = g; });
      if (!src) return;
      var sp = src.procs.filter(function (x) { return x.p === p; })[0];
      l.procs.push({ p: p, loss: sp.loss, inherited: true });
      (inheritedNotes[p] = inheritedNotes[p] || []).push({ style: l.style, color: l.color, part: l.part, kg: l.kg });
    });
  });
  // Dyeing is the one step every fabric goes through. A row with no Dyeing rate while the rest of its Part
  // has one, and no same-colour row to take it from, is almost certainly an omission — say so on the sheet.
  var dyeMissing = [];
  var partHasDye = {};
  lines.forEach(function (l) { if (l.procs.some(function (x) { return x.p === 'Dyeing'; })) partHasDye[normKey(l.part)] = true; });
  lines.forEach(function (l) {
    if (partHasDye[normKey(l.part)] && !l.procs.some(function (x) { return x.p === 'Dyeing'; })) dyeMissing.push({ style: l.style, color: l.color, part: l.part, kg: l.kg });
  });
  var noteKg = function (list) {
    var by = {}, ord = [];
    list.forEach(function (n) { var k = n.style + '|' + n.color; if (!by[k]) { by[k] = { style: n.style, color: n.color, bits: [] }; ord.push(k); } by[k].bits.push(n.part + ' ' + kgRound_(n.kg).toLocaleString('en-IN') + ' kg'); });
    return ord.map(function (k) { return E(by[k].style) + ' / ' + E(by[k].color) + ' (' + by[k].bits.map(E).join('; ') + ')'; }).join(' &middot; ');
  };

  var f2 = function (n) { return (Math.round(n * 100) / 100).toFixed(2); };
  // Every department table is a proper gridded table: header row, one merged block per colour/fabric
  // (Colour, Fabric, Counts, GG, LL, GSM span all of that block's dia lines), blocks alternately shaded
  // so a line is never confused with the next colour's, a sub-total under any block with several dias.
  var head = '<thead><tr><th>Colour</th><th>Fabric Descr</th><th>Counts</th><th class="c">GG</th><th class="c">LL</th><th class="c">K.GSM</th><th class="c">F.GSM</th><th class="c">K.Dia/Wid</th><th class="c">F.Dia/Wid</th><th class="r">Reqd Kgs</th></tr></thead>';
  var table = function (list, withColour, valFn) {
    var order = [], map = {}, tot = 0;
    list.forEach(function (l) {
      var key = (withColour ? l.color : '') + '|' + [l.fabric, l.counts, l.gg, l.ll, l.gsm].join('|');
      if (!map[key]) { map[key] = { l: l, dias: [], dk: {} }; order.push(key); }
      var g = map[key], dk = l.kd + '|' + l.fd;
      if (!(dk in g.dk)) { g.dk[dk] = g.dias.length; g.dias.push({ kd: l.kd, fd: l.fd, kg: 0 }); }
      g.dias[g.dk[dk]].kg += valFn(l);
    });
    var lead = withColour ? 9 : 8;   // columns before "Reqd Kgs"
    var h = '<table class="pt">' + (withColour ? head : head.replace('<th>Colour</th>', ''));
    order.forEach(function (k, gi) {
      var g = map[k], l = g.l, sub = 0, n = g.dias.length, z = gi % 2 ? ' g1' : '';
      var span = n > 1 ? ' rowspan="' + n + '"' : '';
      h += '<tbody>';
      g.dias.forEach(function (d, i) {
        var rk = kgRound_(d.kg); sub += rk; tot += rk;
        h += '<tr>';
        if (i === 0) {
          h += (withColour ? '<td class="b' + z + '"' + span + '>' + E(l.color) + '</td>' : '') +
            '<td class="' + z + '"' + span + '>' + E(l.fabric) + '</td><td class="' + z + '"' + span + '>' + E(l.counts) + '</td>' +
            '<td class="c' + z + '"' + span + '>' + E(l.gg) + '</td><td class="c' + z + '"' + span + '>' + E(l.ll) + '</td>' +
            '<td class="c' + z + '"' + span + '></td><td class="c' + z + '"' + span + '>' + E(l.gsm) + '</td>';
        }
        h += '<td class="c' + z + '">' + E(d.kd) + '</td><td class="c' + z + '">' + E(d.fd) + '</td><td class="r' + z + '">' + rk.toLocaleString('en-IN') + '</td></tr>';
      });
      if (n > 1) h += '<tr class="sub"><td colspan="' + lead + '" class="r">Sub-total' + (withColour ? ' &mdash; ' + E(l.color) : '') + '</td><td class="r">' + sub.toLocaleString('en-IN') + '</td></tr>';
      h += '</tbody>';
    });
    return h + '<tr class="tot"><td colspan="' + lead + '">Total</td><td class="r">' + tot.toLocaleString('en-IN') + '</td></tr></table>';
  };

  var body = '<div class="sec">YARN PURCHASE</div><table class="pt"><thead><tr><th>Yarn Count</th><th>Yarn Colour</th><th class="r">Reqd Kgs</th></tr></thead>';
  var yt = 0;
  yOrder.forEach(function (k, yi) {
    var yk = kgRound_(yarn[k]); yt += yk;
    var z = yi % 2 ? ' g1' : '';
    body += '<tr><td class="' + z + '">' + E(yarnLabel[k].counts) + '</td><td class="' + z + '">' + E(yarnLabel[k].color || '-') + '</td><td class="r' + z + '">' + yk.toLocaleString('en-IN') + '</td></tr>';
  });
  body += '<tr class="tot"><td colspan="2">Total</td><td class="r">' + yt.toLocaleString('en-IN') + '</td></tr></table>';

  body += '<div class="sec">KNITTING</div>' + table(lines, false, function (l) { return l.kg; });
  maps.sequence.forEach(function (p) {
    if (p === 'Knitting') return;
    var list = lines.filter(function (l) { return l.procs.some(function (x) { return x.p === p; }); });
    if (!list.length) return;
    var losses = [];
    list.forEach(function (l) { l.procs.forEach(function (x) { if (x.p === p && losses.indexOf(x.loss) < 0) losses.push(x.loss); }); });
    var label = p === 'Compacting' ? 'OPENWIDTH COMPACTING' : p === 'Stenter' ? 'STENTER' : p.toUpperCase();
    body += '<div class="sec">' + E(label) + (losses.some(function (x) { return x > 0; }) ? '<span>Loss Details : ' + losses.map(f2).join(' / ') + '</span>' : '') + '</div>' +
      table(list, true, function (l) { return l.kg; });
    if (inheritedNotes[p]) body += '<div style="font-size:8px;color:#475569;margin:2px 0 3px;">Included although this row has no ' + E(p) + ' rate of its own &mdash; the same colour and part carries it in another style: ' + noteKg(inheritedNotes[p]) + '.</div>';
    if (p === 'Dyeing' && dyeMissing.length) body += '<div style="border:1px solid #DC2626;background:#FEF2F2;color:#991B1B;padding:3px 6px;margin:3px 0;font-size:8.5px;"><b>FABRIC TOTAL CHECK:</b> not in Dyeing &mdash; no Dyeing rate entered, unlike the rest of the same part: ' + noteKg(dyeMissing) + '. Enter the rate in Style Tracker or this total is short.</div>';
  });
  body += '<div class="sec">Finished Requirement</div>' + table(lines, true, function (l) { return l.fin; });

  var consHtml = consTableHtml_(consumptionForRows_(rows), 'os');
  var css = '<style>@page{size:A4;margin:7mm}body{font-family:Arial,sans-serif;font-size:8.5px;color:#000;margin:0;line-height:1.25}' +
    '.co{font-size:13px;font-weight:bold;color:#0B6E74}.meta td{padding:0px 7px 0px 0;font-size:9px}' +
    '.ttl{text-align:center;font-weight:bold;font-size:10px;margin:5px 0;background:#E3F6F4;padding:3px}' +
    '.sec{font-weight:bold;font-size:9px;margin:6px 0 2px;color:#0B6E74}.sec span{font-weight:normal;color:#000;margin-left:10px}' +
    'table{border-collapse:collapse;width:100%}th{border-top:1px solid #000;border-bottom:1px solid #000;text-align:left;padding:1.5px 3px;font-size:8px}' +
    'td{padding:1px 3px;font-size:8.5px}.r{text-align:right}.sub td{font-weight:bold}.sub td.r{border-top:1px solid #999}.tot td{font-weight:bold;border-top:1px solid #000;border-bottom:1px solid #000}' +
    '.os th,.os td{border:1px solid #999;text-align:center;padding:1.5px 3px}.os th{background:#F1F5F9}.os td.l,.os th.l{text-align:left}.os{page-break-inside:avoid}.sec{page-break-after:avoid}' +
    '.pt{border:1px solid #64748B}.pt th{background:#D9ECEC;border:1px solid #64748B;padding:2px 3px;font-size:8px}.pt td{border:1px solid #B6C2CC;padding:1.5px 3px;vertical-align:middle}' +
    '.pt td.g1,.pt td.c.g1,.pt td.r.g1{background:#EAF4F4}.pt td.b{font-weight:bold}.pt .c,.pt th.c{text-align:center}.pt .r,.pt th.r{text-align:right}' +
    '.pt tr.sub td{background:#DCE8E8;font-weight:bold;border-top:1px solid #64748B}.pt tr.tot td{background:#C9DEDE;font-weight:bold;border:1px solid #64748B}.pt tbody{page-break-inside:avoid}.pt tr{page-break-inside:avoid}' +
    '</style>';
  var footer = '<table style="margin-top:16px"><tr><td style="font-weight:bold">Prepared By</td><td style="font-weight:bold;text-align:center">MERCH</td><td style="font-weight:bold;text-align:center">COSTING</td><td style="font-weight:bold;text-align:right">G.M./C.E.O</td></tr></table>' +
    '<div style="text-align:right;margin-top:5px">For JIWIN APPARELS</div>' +
    '<div style="font-size:7.5px;color:#555;margin-top:5px">Print Date ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm') + '</div>';
  return { css: css, body: body, footer: footer, first: first, orderQty: orderQty, E: E, S: S, consHtml: consHtml };
}
function processReqHeader_(jo, first, orderQty, E, S) {
  return '<table style="width:100%;border-bottom:2px solid #0B6E74;margin-bottom:6px"><tr><td style="width:64px"><img src="' + JIWIN_LOGO_PNG + '" style="height:54px"></td>' +
    '<td><div class="co">JIWIN APPARELS</div><div>10/227, SF NO 4/470, GANAPATHYPALAYAM ROAD, TIRUPUR</div><div>9092998674</div></td>' +
    '<td style="text-align:right;vertical-align:top"><table class="meta" style="width:auto;margin-left:auto">' +
    '<tr><td>ORDER / OCN No</td><td><b>' + E(first['Order/OCN No'] || jo) + '</b></td></tr>' +
    '<tr><td>JO No</td><td><b>' + E(jo) + '</b></td></tr>' +
    '<tr><td>Buyer OrderNo</td><td><b>' + E(first['Buyer Order No']) + '</b></td></tr>' +
    '<tr><td>Order Qty</td><td><b>' + orderQty.toLocaleString('en-IN') + ' PCS</b></td></tr>' +
    '<tr><td>Des.Dt</td><td><b>' + E(S(first['Despatch Date']).slice(0, 10)) + '</b></td></tr></table></td></tr></table>';
}
function generateProcessReqHtml(jo) {
  var rows = getStyleTracker().filter(function (r) { return String(r['JO Number']).trim() === String(jo).trim(); });
  if (!rows.length) throw new Error('JO "' + jo + '" not found in Style Tracker.');
  var P = processReqParts_(jo, rows);
  var img = stImgDataUrl_(stImgLinkForJo_(jo, rows));
  return '<!DOCTYPE html><html><head><meta charset="utf-8">' + P.css + '</head><body>' +
    processReqHeader_(jo, P.first, P.orderQty, P.E, P.S) +
    '<div class="ttl">PURCHASE / PROCESS REQUIRED Details</div>' +
    (img ? '<div style="margin:8px 0;">' + imgTag_(img, 260, 180) + '</div>' : '<div style="margin:6px 0;font-size:10px;color:#B45309;">No garment photo &mdash; ' + stImgWhy_(stImgLinkForJo_(jo, rows)) + '</div>') +
    '<div class="sec">CONSUMPTION PER GARMENT (grams)</div>' + P.consHtml +
    P.body + P.footer + '</body></html>';
}

/* ONE sheet: Order Sheet (colour x size, order vs production qty, trims list) + Fabric Program
   (yarn purchase, every department, finished requirement) for a JO, with the garment photo on top. */
function generateOrderFabricHtml(jo) {
  var rows = getStyleTracker().filter(function (r) { return String(r['JO Number']).trim() === String(jo).trim(); });
  if (!rows.length) throw new Error('JO "' + jo + '" not found in Style Tracker.');
  var P = processReqParts_(jo, rows), E = P.E;
  var os = getOrderSheetData(jo);
  // every style block that belongs to THIS JO (a JO can carry several style numbers; the Order Sheet
  // shares an OCN with other JOs too, but this document is for one JO)
  var jStyles = os ? os.styles.filter(function (x) { return String(x.jo).trim() === String(jo).trim(); }) : [];
  var st = jStyles[0] || (os ? os.styles[0] : null);
  var img = (st && st.image) || stImgDataUrl_(stImgLinkForJo_(jo, rows));

  var h = '<!DOCTYPE html><html><head><meta charset="utf-8">' + P.css + '</head><body>' + processReqHeader_(jo, P.first, P.orderQty, E, P.S);
  // garment photo + style summary side by side
  h += '<table style="margin:4px 0 3px"><tr>' +
    '<td style="width:110px;vertical-align:top">' + (img ? imgTag_(img, 100, 100) : '<div style="width:88px;padding:16px 4px;border:1px dashed #bbb;color:#999;text-align:center;font-size:8px;">No photo<br><span style="font-size:7px;color:#B45309">' + E(stImgWhy_(stImgLinkForJo_(jo, rows))) + '</span></div>') + '</td>' +
    '<td style="vertical-align:top"><table class="meta" style="width:auto">' +
    '<tr><td>Style No</td><td><b>' + E(P.first['Style Number']) + '</b></td></tr>' +
    '<tr><td>Description</td><td><b>' + E(P.first['Style Description']) + '</b></td></tr>' +
    '<tr><td>Fabric / GSM</td><td><b>' + E(P.first['Fabric type']) + ' / ' + E(P.first['GSM']) + '</b></td></tr></table>' +
    '<div class="sec" style="margin-top:8px">CONSUMPTION PER GARMENT (grams)</div>' + P.consHtml + '</td></tr></table>';

  // ---- ORDER SHEET ----
  h += '<div class="ttl">ORDER SHEET</div>';
  if (jStyles.length) {
    var multi = jStyles.length > 1, grandO = 0, grandP = 0;
    jStyles.forEach(function (sx) {
      if (multi) h += '<div class="sec" style="margin-top:6px">Style ' + E(sx.styleNo || '-') + (sx.styleDesc ? ' &mdash; ' + E(sx.styleDesc) : '') + '</div>';
      h += '<table class="os"><thead><tr><th class="l">Colour</th><th></th>' + sx.sizeLabels.map(function (z) { return '<th>' + E(z.label) + '</th>'; }).join('') + '<th>Total</th></tr></thead><tbody>';
      var gO = 0, gP = 0;
      sx.colors.forEach(function (c) {
        gO += c.oTotal; gP += c.pTotal;
        h += '<tr><td class="l" rowspan="2"><b>' + E(c.color || '-') + '</b></td><td class="l">O.Qty</td>' + c.oQty.map(function (q) { return '<td>' + q + '</td>'; }).join('') + '<td><b>' + c.oTotal + '</b></td></tr>' +
          '<tr><td class="l">P.Qty (+' + c.excessPct + '%)</td>' + c.pQty.map(function (q) { return '<td>' + q + '</td>'; }).join('') + '<td><b>' + c.pTotal + '</b></td></tr>';
      });
      grandO += gO; grandP += gP;
      h += '<tr class="tot"><td class="l" colspan="2">TOTAL' + (multi ? ' &mdash; ' + E(sx.styleNo || '') : '') + '</td><td colspan="' + sx.sizeLabels.length + '" style="text-align:right">Order / Production</td><td>' + gO + ' / ' + gP + '</td></tr></tbody></table>';
      if (sx.accessories && sx.accessories.length) {
        h += '<div class="sec">TRIMS / ACCESSORIES' + (multi ? ' &mdash; ' + E(sx.styleNo || '') : '') + '</div><table class="os"><thead><tr><th class="l">Trim</th><th>Colour</th><th>Size</th><th>Required Qty</th><th>I/H Needed</th></tr></thead><tbody>' +
          sx.accessories.map(function (a) { return '<tr><td class="l">' + E(a.name) + '</td><td>' + E(a.color || '-') + '</td><td>' + E(a.size || '-') + '</td><td><b>' + Number(a.requiredQty || 0).toLocaleString('en-IN') + '</b></td><td>' + E(a.ihNeededDate || '-') + '</td></tr>'; }).join('') + '</tbody></table>';
      }
    });
    if (multi) {
      h += '<div class="sec" style="margin-top:6px">COLOUR-WISE TOTAL &mdash; all styles added up</div>' + colourTotalsHtml_(jStyles, E, 'os');
      h += '<table class="os" style="margin-top:4px"><tbody><tr class="tot"><td class="l">ALL STYLES &mdash; Order / Production</td><td style="text-align:right;width:130px">' + grandO + ' / ' + grandP + '</td></tr></tbody></table>';
    }
    var warns = [];
    jStyles.forEach(function (sx) { (sx.warnings || []).forEach(function (w) { warns.push(w); }); });
    if (warns.length) h += qtyWarnHtml_(warns, E);
  } else {
    h += '<div style="color:#777;">No piece-based order lines for this JO.</div>';
  }

  // ---- FABRIC PROGRAM ----
  h += '<div class="ttl" style="margin-top:14px;">FABRIC PROGRAM &mdash; PURCHASE / PROCESS REQUIRED</div>' + P.body + P.footer + '</body></html>';
  return h;
}

/*************** STYLE PHOTOS: lab dips, approved artworks ... (small files, thumbnails from Drive) ***************/
var STYLE_FILES_SHEET = 'STYLE FILES';
var STYLE_FILES_HEAD = ['File ID', 'JO Number', 'Style Number', 'Buyer', 'Category', 'File Name', 'Drive ID', 'Note', 'Uploaded By', 'Uploaded At'];
function sfSheet_() {
  var sh = SS.getSheetByName(STYLE_FILES_SHEET);
  if (!sh) {
    sh = SS.insertSheet(STYLE_FILES_SHEET);
    sh.getRange(1, 1, 1, STYLE_FILES_HEAD.length).setValues([STYLE_FILES_HEAD]).setFontWeight('bold').setBackground('#0B6E74').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }
  return sh;
}
function sfFolder_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('SF_FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { } }
  var f = DriveApp.createFolder('JIWIN_Style_Files');
  p.setProperty('SF_FOLDER', f.getId());
  return f;
}
function getStylePhotos(jo, style) {
  var sh = SS.getSheetByName(STYLE_FILES_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  var tz = Session.getScriptTimeZone();
  return sh.getRange(2, 1, sh.getLastRow() - 1, STYLE_FILES_HEAD.length).getValues().filter(function (r) {
    return String(r[1]).trim() === String(jo).trim() && stLike_(r[2], style) === 2;
  }).map(function (r) {
    return { fid: String(r[0]), category: String(r[4]), name: String(r[5]), id: String(r[6]), note: String(r[7]), by: String(r[8]),
      at: r[9] instanceof Date ? Utilities.formatDate(r[9], tz, 'yyyy-MM-dd') : String(r[9]).slice(0, 10),
      isPdf: /\.pdf$/i.test(String(r[5])),
      thumb: 'https://drive.google.com/thumbnail?id=' + r[6] + '&sz=w320', url: 'https://drive.google.com/file/d/' + r[6] + '/view' };
  });
}
function saveStylePhoto(jo, style, buyer, category, fileName, b64, mime, note, userName) {
  try {
    if (!jo || !style) return { success: false, error: 'JO / style missing.' };
    if (!b64 || b64.length > 8000000) return { success: false, error: 'File too large (max 5 MB).' };
    var name = clean_(fileName, 100).replace(/[\\/:*?"<>|]+/g, '_') || 'photo.jpg';
    var file = sfFolder_().createFile(Utilities.newBlob(Utilities.base64Decode(b64), mime || 'image/jpeg',
      String(jo).replace(/[^\w-]/g, '_') + '_' + String(style).replace(/[^\w-]/g, '_') + '_' + String(category).replace(/[^\w-]/g, '_') + '_' + name));
    try { file.setSharing(readConfig_().techShare === 'DOMAIN' ? DriveApp.Access.DOMAIN_WITH_LINK : DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { }
    var fid = 'SF-' + Utilities.getUuid().slice(0, 8);
    sfSheet_().appendRow([fid, jo, style, buyer || '', category || 'Other', name, file.getId(), clean_(note, 200), clean_(userName, 40) || '', new Date()]);
    return { success: true, message: name + ' added to ' + (category || 'Other') + '.', photos: getStylePhotos(jo, style) };
  } catch (err) {
    return { success: false, error: err.message };
  }
}
function deleteStylePhoto(fid, jo, style) {
  var sh = SS.getSheetByName(STYLE_FILES_SHEET);
  if (!sh || sh.getLastRow() < 2) return { success: false, error: 'Not found.' };
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]) === String(fid)) {
      try { DriveApp.getFileById(String(v[i][6])).setTrashed(true); } catch (e) { }
      sh.deleteRow(i + 2);
      return { success: true, photos: getStylePhotos(jo, style) };
    }
  }
  return { success: false, error: 'Not found.' };
}

/*************** SAFE APPEND + SHEET REPAIR ***************/
/* Writes the row just below the last row that has a value in column A.
   (appendRow() goes below the last row with ANY content - formulas copied down
   to row 200 made new POs land at row 202, out of sight.) */
function appendRowSafe_(sheetName, row) {
  var sh = SS.getSheetByName(sheetName);
  var last = sh.getLastRow(), r = 2;
  if (last >= 2) {
    var a = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = a.length - 1; i >= 0; i--) { if (String(a[i][0]).trim() !== '') { r = i + 3; break; } }
  }
  if (sh.getMaxColumns() < row.length) sh.insertColumnsAfter(sh.getMaxColumns(), row.length - sh.getMaxColumns());
  sh.getRange(r, 1, 1, row.length).setValues([row]);
  delete _sheetValuesCache_[sheetName];   // this sheet changed — the next sheetToObjects() call for it must re-read, not use a stale cache
  return r;
}

/* If a tab the app uses (e.g. "Yarn PO") has a different column layout (a template's),
   rename it to "<name> (old layout)", let the app create a clean tab, and move over
   the rows the app itself wrote there. Duplicate PO numbers are renumbered. */
function repairStoreLoss_() {
  // Before this fix, part-issues from Yarn / Cora / Dyed store were saved as "loss". Clear them.
  var sh = SS.getSheetByName(SHEETS.PROCESS_DC);
  if (!sh || sh.getLastRow() < 2) return;
  var H = HEADERS[SHEETS.PROCESS_DC], fp = H.indexOf('From Process'), lk = H.indexOf('Loss KG'), lp = H.indexOf('Loss %');
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, H.length).getValues();
  var stores = [YARN_STORE_LABEL, CORA_STORE_LABEL, DYED_STORE_LABEL];
  v.forEach(function (r, i) {
    if (stores.indexOf(String(r[fp])) > -1 && Number(r[lk]) > 0) {
      sh.getRange(i + 2, lk + 1, 1, 2).setValues([[0, 0]]);
    }
  });
}
function repairForeignSheets_() {
  try { repairStoreLoss_(); } catch (e) { }
  var name = SHEETS.YARN_PO, sh = SS.getSheetByName(name);
  if (!sh || sh.getLastRow() < 1) return;
  var want = HEADERS[name];
  var first = String(sh.getRange(1, 1).getValue()).trim();
  if (first === want[0]) return;
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var old = sh.setName(name + ' (old layout)');
    var nsh = SS.insertSheet(name);
    nsh.getRange(1, 1, 1, want.length).setValues([want]).setFontWeight('bold').setBackground('#1B2A4A').setFontColor('#FFFFFF');
    nsh.setFrozenRows(1);
    var last = old.getLastRow();
    if (last < 2) return;
    var width = Math.min(want.length, old.getMaxColumns());
    var vals = old.getRange(2, 1, last - 1, width).getValues();
    var types = ['Yarn', 'Cora Fabric', 'Dyed Fabric'], seen = {}, max = 0, keep = [];
    vals.forEach(function (r) {
      if (!/^YPO-/i.test(String(r[0])) || types.indexOf(String(r[2]).trim()) < 0) return;   // only rows written by the app
      var m = String(r[0]).match(/(\d+)$/); if (m) max = Math.max(max, Number(m[1]));
      while (r.length < want.length) r.push('');
      keep.push(r);
    });
    keep.forEach(function (r) {
      var id = String(r[0]).trim();
      if (seen[id]) { max++; r[0] = 'YPO-' + Utilities.formatString('%04d', max); }
      seen[String(r[0])] = 1;
    });
    if (keep.length) nsh.getRange(2, 1, keep.length, want.length).setValues(keep);
  } finally { lock.releaseLock(); }
}

/*************** API KEY FROM THE APP (no sheet menu needed) ***************/
function hasAnthropicApiKey() {
  var p = PropertiesService.getScriptProperties();
  return !!(p.getProperty('ANTHROPIC_API_KEY') || p.getProperty('GEMINI_API_KEY'));
}
function saveAnthropicApiKey(key) {   // saves either a Google Gemini key (AIza...) or an Anthropic key (sk-ant-...)
  key = String(key || '').trim();
  var props = PropertiesService.getScriptProperties();
  if (/^AIza/.test(key)) {
    try {
      var g = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?key=' + encodeURIComponent(key), { muteHttpExceptions: true });
      if (g.getResponseCode() === 400 || g.getResponseCode() === 403) return { success: false, error: 'Google says this key is not valid. Copy it again from aistudio.google.com > Get API key.' };
    } catch (e) { }
    props.setProperty('GEMINI_API_KEY', key); props.setProperty('AI_PROVIDER', 'GEMINI');
    return { success: true, message: 'Google Gemini key saved. PO import is ready.' };
  }
  if (/^sk-ant-/.test(key)) {
    try {
      var r = UrlFetchApp.fetch('https://api.anthropic.com/v1/models', { method: 'get', headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' }, muteHttpExceptions: true });
      if (r.getResponseCode() === 401) return { success: false, error: 'Anthropic says this key is not valid. Copy it again from console.anthropic.com > API Keys.' };
    } catch (e) { }
    props.setProperty('ANTHROPIC_API_KEY', key); props.setProperty('AI_PROVIDER', 'ANTHROPIC');
    return { success: true, message: 'Anthropic key saved. PO import is ready.' };
  }
  return { success: false, error: 'Paste a Google Gemini key (starts with AIza...) or an Anthropic key (starts with sk-ant-...).' };
}

/*************** PERMANENT BUDGET LOCK ***************
 * Once a JO's fabric budget is locked, its process rates, loss % and yarn rates can never be
 * changed from the app - not even by saving the JO in Style Tracker. Only the owner can unlock,
 * by running unlockJoBudget('JO-xx') from the Apps Script editor.
 *****************************************************/
/*************** JO DATA LOCK — freezes a JO's Style Tracker rows completely (every field, not just
 * rates), once the Order Sheet + Fabric Program are final. App-level only: saveStyleTrackerJo checks
 * this before anything else and refuses to save a locked JO. There's no Google Sheets range
 * protection here — Sheets protection can never block the spreadsheet's owner anyway, so it added
 * a point of failure (and a stale-row-number risk as the sheet grows) without buying real security.
 * Editing a locked JO directly in the sheet is still possible for anyone with sheet edit access;
 * this lock only governs the app. ***************/
var JO_DATA_LOCK_SHEET = 'JO DATA LOCK';
function joDataLockSheet_() {
  var sh = SS.getSheetByName(JO_DATA_LOCK_SHEET);
  if (!sh) {
    sh = SS.insertSheet(JO_DATA_LOCK_SHEET);
    sh.getRange(1, 1, 1, 5).setValues([['JO Number', 'Locked By', 'Locked At', 'Protection ID', 'Note']])
      .setFontWeight('bold').setBackground('#0B6E74').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    try { sh.protect().setWarningOnly(true).setDescription('JO data locks - do not edit'); } catch (e) { }
  }
  return sh;
}
function getJoDataLock(jo) {
  var sh = SS.getSheetByName(JO_DATA_LOCK_SHEET);
  if (!sh || sh.getLastRow() < 2) return null;
  var tz = Session.getScriptTimeZone();
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim() === String(jo).trim()) {
      return { jo: v[i][0], by: v[i][1], at: v[i][2] instanceof Date ? Utilities.formatDate(v[i][2], tz, 'dd-MM-yyyy HH:mm') : String(v[i][2]) };
    }
  }
  return null;
}
/* App-level lock only — no Google Sheets range protection. Protection can't block the spreadsheet's
   owner anyway (Sheets never allows that, by design), so it added a point of failure without buying
   real security. The real protection is here: saveStyleTrackerJo checks getJoDataLock() before any
   write, so no save for a locked JO can ever go through the app, by anyone. */
/*************** SAVED COPY OF A LOCKED JO'S SHEETS — locking a JO also saves its Order Sheet, Fabric Program,
 * Order + Fabric and Fabric Consumption documents exactly as they are at that moment (as HTML files in
 * Drive). From then on, viewing or downloading any of them for that JO serves the saved copy — never a
 * fresh regeneration — so later changes to the app, or edits made straight in the Google Sheet, cannot
 * alter a sheet that was finalised. Unlocking discards the saved copy; locking again saves a new one. ***************/
var JO_LOCKED_SHEETS_SHEET = 'JO LOCKED SHEETS';
var LOCKED_SHEET_TYPES_ = ['ORDERSHEET', 'PROCESSREQ', 'ORDERFABRIC', 'FABRICPROGRAM'];
var LOCKED_SHEET_LABELS_ = { ORDERSHEET: 'Order Sheet', PROCESSREQ: 'Fabric Program', ORDERFABRIC: 'Order + Fabric', FABRICPROGRAM: 'Fabric Consumption' };
function getLockedSheetsFolder_() {
  var folders = DriveApp.getFoldersByName('JIWIN_Locked_JO_Sheets');
  return folders.hasNext() ? folders.next() : DriveApp.createFolder('JIWIN_Locked_JO_Sheets');
}
function lockedSheetsSheet_() {
  var sh = SS.getSheetByName(JO_LOCKED_SHEETS_SHEET);
  if (!sh) {
    sh = SS.insertSheet(JO_LOCKED_SHEETS_SHEET);
    sh.getRange(1, 1, 1, 5).setValues([['JO Number', 'Doc Type', 'File ID', 'Saved At', 'Saved By']])
      .setFontWeight('bold').setBackground('#0B6E74').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }
  return sh;
}
function lockedSheetRecords_(jo) {
  var sh = SS.getSheetByName(JO_LOCKED_SHEETS_SHEET), out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  var tz = Session.getScriptTimeZone();
  sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().forEach(function (r, i) {
    if (String(r[0]).trim() !== String(jo).trim()) return;
    out.push({ row: i + 2, type: String(r[1]), fileId: String(r[2]), by: String(r[4] || ''),
      at: r[3] instanceof Date ? Utilities.formatDate(r[3], tz, 'dd-MM-yyyy HH:mm') : String(r[3]) });
  });
  return out;
}
function getLockedSheetTypes_(jo) {
  return lockedSheetRecords_(jo).map(function (r) { return r.type; });
}
/* Builds the document fresh from the current data — only ever used to take the saved copy at lock time
   (or for a JO that isn't locked), never to answer a request for a locked JO that has a saved copy. */
function liveJoDocHtml_(type, jo) {
  if (type === 'ORDERSHEET') {
    var os = getOrderSheetData(jo);
    if (!os) throw new Error('JO "' + jo + '" not found in Style Tracker.');
    return generateOrderSheetHtml(os);
  }
  if (type === 'FABRICPROGRAM') {
    var fp = getFabricProgramData(jo);
    if (!fp) throw new Error('JO "' + jo + '" not found in Style Tracker.');
    return generateFabricProgramHtml(fp);
  }
  if (type === 'PROCESSREQ') return generateProcessReqHtml(jo);
  if (type === 'ORDERFABRIC') return generateOrderFabricHtml(jo);
  throw new Error('No saved copy is kept for ' + type + '.');
}
/* The saved copy of this document for a locked JO: { html, at, by } — or null when the JO isn't locked or
   nothing was saved (the caller then falls back to generating it live, as before). */
function lockedSheetHtml_(type, jo) {
  if (LOCKED_SHEET_TYPES_.indexOf(type) < 0) return null;
  jo = String(jo || '').trim();
  if (!jo || !getJoDataLock(jo)) return null;
  var rec = lockedSheetRecords_(jo).filter(function (r) { return r.type === type; })[0];
  if (!rec) return null;
  try {
    return { html: DriveApp.getFileById(rec.fileId).getBlob().getDataAsString('UTF-8'), at: rec.at, by: rec.by };
  } catch (e) {
    return { missing: true, at: rec.at, by: rec.by };   // the saved file was deleted from Drive — say so, don't pretend
  }
}
function saveLockedSheets_(jo, by) {
  var saved = [], failed = [], folder = getLockedSheetsFolder_(), sh = lockedSheetsSheet_();
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  LOCKED_SHEET_TYPES_.forEach(function (type) {
    try {
      var html = liveJoDocHtml_(type, jo);
      var file = folder.createFile('LOCKED_' + String(jo).replace(/[^\w-]+/g, '_') + '_' + type + '_' + stamp + '.html', html, 'text/html');
      sh.appendRow([jo, type, file.getId(), new Date(), String(by || '').slice(0, 40)]);
      saved.push(type);
    } catch (e) {
      failed.push({ type: type, error: e && e.message ? e.message : String(e) });
    }
  });
  delete _sheetValuesCache_[JO_LOCKED_SHEETS_SHEET];
  return { saved: saved, failed: failed };
}
function deleteLockedSheets_(jo) {
  var recs = lockedSheetRecords_(jo), sh = SS.getSheetByName(JO_LOCKED_SHEETS_SHEET);
  recs.forEach(function (r) { try { DriveApp.getFileById(r.fileId).setTrashed(true); } catch (e) { } });
  recs.map(function (r) { return r.row; }).sort(function (a, b) { return b - a; }).forEach(function (row) { sh.deleteRow(row); });
}
function lockWithSheets_(jo) {
  var lock = getJoDataLock(jo);
  return lock ? Object.assign({}, lock, { savedSheets: getLockedSheetTypes_(jo) }) : null;
}
function lockedSheetsSummary_(r) {
  var parts = [];
  if (r.saved.length) parts.push(r.saved.map(function (t) { return LOCKED_SHEET_LABELS_[t]; }).join(', ') + ' saved exactly as they are now — later changes to the app or the sheet will not alter them.');
  if (r.failed.length) parts.push('Could not save a copy of: ' + r.failed.map(function (f) { return LOCKED_SHEET_LABELS_[f.type] + ' (' + f.error + ')'; }).join('; ') + ' — those will still be generated fresh.');
  return parts.join(' ');
}
/* For a JO that was locked before saved copies existed (or whose copy could not be made): saves them now.
   This captures the sheets as they generate TODAY. Refused if a copy already exists — to replace it, unlock
   the JO and lock it again. */
function saveLockedSheetsNow(jo) {
  jo = String(jo || '').trim();
  if (!jo) return { success: false, error: 'Pick a JO.' };
  if (!getJoDataLock(jo)) return { success: false, error: jo + ' is not locked.' };
  var have = getLockedSheetTypes_(jo);
  if (have.length >= LOCKED_SHEET_TYPES_.length) return { success: false, error: 'Saved copies already exist for ' + jo + '. To replace them, unlock the JO and lock it again.' };
  var r = saveLockedSheets_(jo, (getJoDataLock(jo) || {}).by);
  if (!r.saved.length) return { success: false, error: lockedSheetsSummary_(r) };
  return { success: true, message: lockedSheetsSummary_(r), lock: lockWithSheets_(jo) };
}

function lockJoData(jo, lockedBy) {
  jo = String(jo || '').trim();
  if (!jo) return { success: false, error: 'Pick a JO.' };
  if (getJoDataLock(jo)) return { success: false, error: jo + ' is already locked.' };
  var sh = SS.getSheetByName(SHEETS.STYLE_TRACKER);
  var head = stHead_(sh), joCol = head.indexOf('JO Number') + 1;
  var last = sh.getLastRow();
  var rowNums = [];
  if (last >= 2) {
    var vals = sh.getRange(2, joCol, last - 1, 1).getValues();
    vals.forEach(function (r, i) { if (String(r[0]).trim() === jo) rowNums.push(i + 2); });
  }
  if (!rowNums.length) return { success: false, error: 'No Style Tracker rows found for ' + jo + '.' };
  var minRow = Math.min.apply(null, rowNums), maxRow = Math.max.apply(null, rowNums);
  joDataLockSheet_().appendRow([jo, String(lockedBy || '').slice(0, 40) || 'unknown', new Date(), '', 'Locked from app — rows ' + minRow + '-' + maxRow]);
  var snap = { saved: [], failed: [] };
  try { snap = saveLockedSheets_(jo, lockedBy); } catch (e) { snap.failed = LOCKED_SHEET_TYPES_.map(function (t) { return { type: t, error: e.message }; }); }
  return { success: true, message: jo + ' is now locked — no further edits through the app. ' + lockedSheetsSummary_(snap), lock: lockWithSheets_(jo) };
}
/* Lets someone unlock a JO from inside the app itself, without needing Apps Script editor access —
   gated by its own password (change it here), separate from Management/Accessories/Bill Passing. */
var JO_UNLOCK_PASSWORD = '7410';
function unlockJoDataWithPassword(jo, password) {
  if (String(password || '') !== JO_UNLOCK_PASSWORD) return { success: false, error: 'Incorrect password.' };
  jo = String(jo || '').trim();
  if (!getJoDataLock(jo)) return { success: false, error: jo + ' is not locked.' };
  unlockJoData(jo);
  return { success: true, message: jo + ' is unlocked — it can be edited again through the app.' };
}
/* Run from the Apps Script editor only (e.g. unlockJoData('JO-73 ADONIS')) — no password needed here,
   since being in the script editor already means you're trusted. For an in-app unlock (with the
   password prompt), use unlockJoDataWithPassword instead. */
function unlockJoData(jo) {
  jo = String(jo || '').trim();
  var sh = SS.getSheetByName(JO_DATA_LOCK_SHEET);
  if (sh && sh.getLastRow() >= 2) {
    var v = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (var i = v.length - 1; i >= 0; i--) if (String(v[i][0]).trim() === jo) sh.deleteRow(i + 2);
  }
  try { deleteLockedSheets_(jo); } catch (e) { }   // the saved copy belongs to the locked version — gone with it
  Logger.log('Unlocked ' + jo);
  return 'Unlocked ' + jo;
}

var BUDGET_LOCK_SHEET = 'BUDGET LOCK';
function budgetLockSheet_() {
  var sh = SS.getSheetByName(BUDGET_LOCK_SHEET);
  if (!sh) {
    sh = SS.insertSheet(BUDGET_LOCK_SHEET);
    sh.getRange(1, 1, 1, 5).setValues([['JO Number', 'Locked By', 'Locked At', 'Budget Total (Rs)', 'Note']])
      .setFontWeight('bold').setBackground('#0B6E74').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    try { sh.protect().setWarningOnly(true).setDescription('Budget locks - do not edit'); } catch (e) { }
  }
  return sh;
}
function getBudgetLock(jo) {
  var sh = SS.getSheetByName(BUDGET_LOCK_SHEET);
  if (!sh || sh.getLastRow() < 2) return null;
  var tz = Session.getScriptTimeZone();
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim() === String(jo).trim()) {
      return { jo: v[i][0], by: v[i][1], at: v[i][2] instanceof Date ? Utilities.formatDate(v[i][2], tz, 'dd-MM-yyyy HH:mm') : String(v[i][2]), total: v[i][3] };
    }
  }
  return null;
}
function lockJoBudget(jo, lockedBy, budgetTotal) {
  jo = String(jo || '').trim();
  if (!jo) return { success: false, error: 'Pick a JO.' };
  if (getBudgetLock(jo)) return { success: false, error: 'Budget for ' + jo + ' is already locked.' };
  budgetLockSheet_().appendRow([jo, String(lockedBy || '').slice(0, 40) || 'unknown', new Date(), Number(budgetTotal) || 0, 'Locked from app']);
  return { success: true, message: 'Budget for ' + jo + ' is now locked permanently.', lock: getBudgetLock(jo) };
}
/* Run from the Apps Script editor only (e.g. unlockJoBudget('JO-73 ADONIS')). */
function unlockJoBudget(jo) {
  var sh = SS.getSheetByName(BUDGET_LOCK_SHEET);
  if (!sh || sh.getLastRow() < 2) return 'No locks.';
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = v.length - 1; i >= 0; i--) if (String(v[i][0]).trim() === String(jo).trim()) sh.deleteRow(i + 2);
  Logger.log('Unlocked ' + jo);
  return 'Unlocked ' + jo;
}

/* One DC with several lines (e.g. a style with 2 fabrics / colours going to the same supplier).
   common = JO, date, to process, to supplier, lot, supplier DC no, vehicle, remarks
   lines  = [{ sourceType, sourceDcNo, fromProcess, fromSupplier, counts, content, gsm, color, dcQtyKg, rolls, closeLot }]
   All lines are checked first; nothing is saved if any line is over its available kg. */
function saveProcessDCMulti(common, lines) {
  var tok = common && common.submitToken;
  var c = Object.assign({}, common || {}); delete c.submitToken;   // the lines inside must not each carry the token
  return once_(tok, function () { return saveProcessDCMulti_(c, lines); });
}
function saveProcessDCMulti_(common, lines) {
  lines = (lines || []).filter(function (l) { return Number(l.dcQtyKg) > 0; });
  if (!lines.length) return { success: false, error: 'Enter a qty on at least one line.' };
  if (lines.length === 1) return saveProcessDC(Object.assign({}, common, lines[0]));
  var need = {};
  for (var i = 0; i < lines.length; i++) {
    var l = lines[i];
    var origin = CHAIN_ORIGINS.filter(function (o) { return o.sourceType === l.sourceType; })[0];
    var key = l.sourceType === 'DC' ? 'DC|' + l.sourceDcNo : [l.sourceType, l.counts, l.content, origin && origin.colorSpecific ? l.color : '', origin && origin.useGsm ? l.gsm : ''].join('|');
    if (!need[key]) {
      var avail = l.sourceType === 'DC' ? getDCChainBalance(l.sourceDcNo)
        : (origin ? getStockBalance(common.joNumber, l.counts, origin.stockType, origin.colorSpecific ? l.color : '', l.content, origin.useGsm ? l.gsm : '').available : 0);
      need[key] = { avail: avail, qty: 0, label: l.sourceType === 'DC' ? (l.fromProcess + ' ' + l.sourceDcNo) : (l.counts + ' ' + (l.content || '')) };
    }
    need[key].qty += Number(l.dcQtyKg) || 0;
  }
  var over = Object.keys(need).filter(function (k) { return need[k].qty > need[k].avail + 0.01; })[0];
  if (over) return { success: false, error: 'Lines from ' + need[over].label + ' add up to ' + need[over].qty.toFixed(2) + ' kg, but only ' + need[over].avail.toFixed(2) + ' kg is available.' };

  // Locked fabric quantity across all lines of this DC (same fabric type + colour + To Process share one balance)
  if (common.processType !== 'Re-process' && dcJoFabricTypes_(common.joNumber).length) {
    var fneed = {};
    for (var q = 0; q < lines.length; q++) {
      var fl = lines[q], ft = dcxS_(fl.fabricType);
      if (!ft) return { success: false, error: 'Line ' + (q + 1) + ': select the Fabric Type.' };
      var fk = [ft.toUpperCase(), dcxS_(fl.knittingDia).toUpperCase(), dcxS_(fl.color).toUpperCase()].join('|');
      if (!fneed[fk]) {
        var fb = getDcFabricBalance(common.joNumber, ft, fl.color || '', common.toProcess, fl.knittingDia || '');
        if (!fb.found) return { success: false, error: 'Line ' + (q + 1) + ': fabric type \"' + ft + '\" is not in Style Tracker for ' + common.joNumber + '.' };
        fneed[fk] = { bal: fb.balanceKg, qty: 0, label: fb.fabricType + (fl.knittingDia ? ' / Dia ' + fl.knittingDia : '') + (fl.color ? ' / ' + fl.color : '') };
      }
      fneed[fk].qty += Number(fl.dcQtyKg) || 0;
    }
    var fover = Object.keys(fneed).filter(function (k) { return fneed[k].qty > fneed[k].bal + 0.01; })[0];
    if (fover) return { success: false, error: 'Fabric ' + fneed[fover].label + ': lines add up to ' + fneed[fover].qty.toFixed(2) + ' kg, but only ' + Math.max(0, fneed[fover].bal).toFixed(2) + ' kg is left to send to ' + common.toProcess + ' (locked quantity).' };
  }

  if (common.photoB64) {   // upload the supplier DC photo once for the whole DC
    common = Object.assign({}, common);
    common.photoUrl = dcPhotoSave_(common.photoB64, common.photoMime, common.photoName, common.joNumber);
    delete common.photoB64;
  }
  var base = '', saved = [];
  for (var j = 0; j < lines.length; j++) {
    var d = Object.assign({}, common, lines[j]);
    if (j > 0) { d._dcNo = dcxRowId_(base, j); d._dcGroup = base; }
    else d._dcGroup = '__FIRST__';
    var r = saveProcessDC(d);
    if (!r.success) return { success: false, error: 'Line ' + (j + 1) + ': ' + r.error + (saved.length ? ' (lines ' + saved.join(', ') + ' were saved)' : '') };
    if (j === 0) {
      base = r.dcNo;
      var hx = dcxEnsureHeaders_(), rn = findRowIndex(SHEETS.PROCESS_DC, 'DC No', base);
      if (rn > 0) hx.sheet.getRange(rn, hx.col['DC Group'] + 1).setValue(base);
    }
    saved.push(r.dcNo);
  }
  return { success: true, dcNo: base, message: 'DC ' + base + ' created with ' + saved.length + ' lines (' + saved.join(', ') + ').' };
}

/* Accessories: plan (BOM line) and raise the PO in one step.
   The PO qty = required qty of the line(s) just planned (never more than the trim's balance),
   rate = the planned rate. The PO goes to Management approval like any accessories PO. */
function saveAccessoryBOMAndPO(d) {
  var res = saveAccessoryBOM(d);
  if (!res.success || !d.makePO) return res;
  if (!String(d.supplier || '').trim()) return { success: true, message: res.message + ' PO not made: no supplier chosen.' };
  var bom = getAccessoriesBOM(d.joNumber);
  var same = function (a, b) { return String(a || '').trim().toUpperCase() === String(b || '').trim().toUpperCase(); };
  var mine = bom.filter(function (r) { return same(r['Trims / Accessories'], d.name) && same(r['Color'], d.color); });
  var bySize = {}, order = [];
  mine.forEach(function (r) {
    var sz = String(r['Size'] || '').trim();
    if (!(sz in bySize)) { bySize[sz] = { req: 0, rate: Number(r['Rate (Rs/unit)']) || 0 }; order.push(sz); }
    bySize[sz].req += Number(r['Required Quantity']) || 0;
  });
  var pos = [], notes = [];
  order.forEach(function (sz) {
    var total = bom.filter(function (r) { return same(r['Trims / Accessories'], d.name) && String(r['Size'] || '').trim() === sz; })
      .reduce(function (t, r) { return t + (Number(r['Required Quantity']) || 0); }, 0);
    var already = getAccessoryPOUsage(d.joNumber, d.name, sz);
    var qty = Math.min(bySize[sz].req, Math.max(0, total - already));
    if (qty <= 0) { notes.push((sz || 'all sizes') + ': already fully on PO'); return; }
    var p = saveAccessoryPO({ joNumber: d.joNumber, trimName: d.name, size: sz, requiredQty: total, qtyOrdered: qty,
      rate: bySize[sz].rate, supplier: d.supplier, gstNo: d.gstNo || '', gstPct: d.gstPct || 0, poDate: d.poDate || '' });
    if (p.success) pos.push(p.poNo + (sz ? ' (' + sz + ')' : '') + ' ' + qty.toLocaleString('en-IN') + ' pcs');
    else notes.push((sz || 'all sizes') + ': ' + p.error);
  });
  return { success: true, pos: pos,
    message: res.message + (pos.length ? ' PO raised: ' + pos.join(', ') + ' — awaiting approval.' : '') + (notes.length ? ' Note: ' + notes.join('; ') : '') };
}

/*************** ACCESSORIES PO - several lines in one PO (like CMT manual bill entry) ***************/
function apoGroupCol_() {
  var sh = SS.getSheetByName(SHEETS.ACCESSORIES_PO), base = HEADERS[SHEETS.ACCESSORIES_PO].length;
  var lastCol = Math.max(sh.getLastColumn(), base);
  if (sh.getMaxColumns() < base + 1) sh.insertColumnsAfter(sh.getMaxColumns(), base + 1 - sh.getMaxColumns());
  var head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (x) { return String(x).trim(); });
  var i = head.indexOf('PO Group');
  if (i > -1) return i + 1;
  sh.getRange(1, base + 1).setValue('PO Group').setFontWeight('bold').setBackground('#1B2A4A').setFontColor('#FFFFFF');
  return base + 1;
}
/* rows of the same multi-line PO as the given PO No: [{no, row}] */
function apoGroupRows_(poNo) {
  var sh = SS.getSheetByName(SHEETS.ACCESSORIES_PO);
  if (!sh || sh.getLastRow() < 2) return [];
  var head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(function (x) { return String(x).trim(); });
  var gi = head.indexOf('PO Group');
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(sh.getLastColumn(), 1)).getValues();
  var me = v.filter(function (r) { return String(r[0]) === String(poNo); })[0];
  if (!me) return [];
  var grp = gi > -1 ? String(me[gi] || '').trim() : '';
  if (!grp) return [{ no: poNo, row: v.indexOf(me) + 2 }];
  var out = [];
  v.forEach(function (r, i) { if (gi > -1 && String(r[gi] || '').trim() === grp) out.push({ no: String(r[0]), row: i + 2 }); });
  return out;
}
/* h = { poDate, supplier, gstNo, gstPct }; rows = [{ joNumber, trimName, size, qty, rate }] */
function saveAccessoryPOMulti(h, rows) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    if (!String(h.supplier || '').trim()) return { success: false, error: 'Choose the supplier.' };
    if (!String(h.quotationUrl || '').trim()) return { success:false, error:'Attach the supplier quotation before submitting for approval.' };
    rows = (rows || []).filter(function (r) { return r.trimName && Number(r.qty) > 0; });
    if (!rows.length) return { success: false, error: 'Add at least one row with a trim and qty.' };
    // balance per JO + trim + size (all colours), checked for ALL rows before anything is saved
    var need = {}, itemsCache = {};
    rows.forEach(function (r) {
      if (!r.joNumber) return; // general purchase is not checked against a JO BOM
      var key = r.joNumber + '|' + String(r.trimName).trim() + '|' + String(r.size || '').trim();
      if (!need[key]) {
        var items = itemsCache[r.joNumber] || (itemsCache[r.joNumber] = getAccessoryItemsForJO(r.joNumber));
        var req = items.filter(function (it) { return String(it.name).trim() === String(r.trimName).trim() && String(it.size || '').trim() === String(r.size || '').trim(); })
          .reduce(function (t, it) { return t + it.requiredQty; }, 0);
        var already = getAccessoryPOUsage(r.joNumber, r.trimName, r.size);
        need[key] = { req: req, already: already, qty: 0, label: r.joNumber + ' / ' + r.trimName + (r.size ? ' / ' + r.size : '') };
      }
      need[key].qty += Number(r.qty) || 0;
    });
    var over = Object.keys(need).filter(function (k) { return need[k].qty > need[k].req - need[k].already + 0.5; })[0];
    if (over) {
      var n = need[over];
      return { success: false, error: n.label + ': ' + n.qty + ' ordered, but only ' + Math.max(0, n.req - n.already) + ' is left to order (required ' + n.req + ', already on PO ' + n.already + ').' };
    }
    var gcol = apoGroupCol_(), sh = SS.getSheetByName(SHEETS.ACCESSORIES_PO);
    var base = nextId('APO-', SHEETS.ACCESSORIES_PO, 1), date = h.poDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var nos = [];
    rows.forEach(function (r, i) {
      var jo = r.joNumber ? (getJOInfo(r.joNumber) || {}) : {};
      var key = r.joNumber + '|' + String(r.trimName).trim() + '|' + String(r.size || '').trim();
      var no = i === 0 ? base : dcxRowId_(base, i), qty = Number(r.qty) || 0, rate = Number(r.rate) || 0;
      var rowNo = appendRowSafe_(SHEETS.ACCESSORIES_PO, [
        no, date, r.joNumber || '', r.styleNo || jo['Style Number'] || '', r.trimName, r.size || '', (need[key] ? need[key].req : 0), qty, rate, rate * qty,
        h.supplier, h.gstNo || '', Number(h.gstPct) || 0, 'Pending Approval', '', '', '', (r.joNumber ? (h.purchaseType || 'JO-linked') : 'General Purchase'), h.quotationUrl || '', r.generalDescription || '', ''
      ]);
      if (rows.length > 1) sh.getRange(rowNo, gcol).setValue(base);
      nos.push(no);
    });
    return { success: true, poNo: base, message: 'Accessories PO ' + base + (nos.length > 1 ? ' with ' + nos.length + ' lines' : '') + ' submitted — awaiting approval.' };
  } catch (err) {
    return { success: false, error: err.message };
  } finally { lock.releaseLock(); }
}

function saveAccessoriesQuotation(fileName, mimeType, base64Data) {
  try {
    if (!base64Data) return { success:false, error:'No quotation file received.' };
    var bytes = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(bytes, mimeType || 'application/pdf', fileName || 'Accessories_Quotation');
    var folderName = 'Jiwin Accessories Quotations';
    var it = DriveApp.getFoldersByName(folderName), folder = it.hasNext() ? it.next() : DriveApp.createFolder(folderName);
    var file = folder.createFile(blob);
    return { success:true, url:file.getUrl(), name:file.getName() };
  } catch(e) { return {success:false,error:e.message}; }
}

function generateAccessoryPoMultiHtml_(lines) {
  return generateAccessoryPoDocumentHtml_(lines || []);
}

/*************** Accessories planning line: edit / delete ***************/
function accBomFindRow_(k) {
  var sh = SS.getSheetByName(SHEETS.ACCESSORIES_BOM), H = HEADERS[SHEETS.ACCESSORIES_BOM];
  if (!sh || sh.getLastRow() < 2) return null;
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, H.length).getValues();
  var col = function (h) { return H.indexOf(h); }, S = function (x) { return String(x == null ? '' : x).trim().toUpperCase(); };
  for (var i = 0; i < v.length; i++) {
    var r = v[i];
    if (S(r[col('JO Number')]) === S(k.jo) && S(r[col('Trims / Accessories')]) === S(k.trim) &&
        S(r[col('Color')]) === S(k.color) && S(r[col('Size')]) === S(k.size)) return { sh: sh, row: i + 2, col: col };
  }
  return null;
}
function accBomHasPo_(k) {
  return getAccessoryPOUsage(k.jo, k.trim, k.size) > 0;
}
/* k = { jo, trim, color, size }; f = { qtyNeeded, perGarments, rate } */
function updateAccessoryBOMRow(k, f) {
  var hit = accBomFindRow_(k);
  if (!hit) return { success: false, error: 'Planning line not found.' };
  var q = Number(f.qtyNeeded), p = Number(f.perGarments) || 1, rt = Number(f.rate) || 0;
  if (!(q > 0)) return { success: false, error: 'Qty per garment must be more than 0.' };
  if (accBomHasPo_(k) && q > 0) {
    // a PO exists: allow only rate / usage fixes that do NOT drop required qty below what is already ordered
    var cut = getTotalCutQtyForJO(k.jo, k.color).total, newReq = Math.ceil(cut / p) * q;
    var already = getAccessoryPOUsage(k.jo, k.trim, k.size);
    if (newReq < already) return { success: false, error: 'Already ' + already + ' on PO for ' + k.trim + '. The new requirement (' + newReq + ') would be less - reject / cancel that PO first.' };
  }
  hit.sh.getRange(hit.row, hit.col('Qty Needed') + 1).setValue(q);
  hit.sh.getRange(hit.row, hit.col('Per How Many Garments') + 1).setValue(p);
  hit.sh.getRange(hit.row, hit.col('Rate (Rs/unit)') + 1).setValue(rt);
  return { success: true, message: k.trim + ' / ' + (k.color || 'all colours') + ' updated.' };
}
function deleteAccessoryBOMRow(k) {
  var hit = accBomFindRow_(k);
  if (!hit) return { success: false, error: 'Planning line not found.' };
  if (accBomHasPo_(k)) return { success: false, error: 'A PO already exists for ' + k.trim + (k.size ? ' / ' + k.size : '') + '. Reject / cancel it first, or edit the line instead.' };
  hit.sh.deleteRow(hit.row);
  return { success: true, message: k.trim + ' / ' + (k.color || 'all colours') + ' removed from planning.' };
}

/*************** ACCESSORIES PLANNING - many rows at once (CMT-style entry) ***************/
function getAccPlanContext(joNumber) {
  var rows = getStyleTracker().filter(function (r) { return String(r['JO Number']).trim() === String(joNumber).trim() && !isSampleWtRow_(r); });
  if (!rows.length) throw new Error('JO ' + joNumber + ' not found in Style Tracker.');
  var styles = [], smap = {}, sizes = [];
  rows.forEach(function (r) {
    for (var i = 1; i <= 10; i++) { var s = String(r['Size ' + i] || '').trim(); if (s && sizes.indexOf(s) < 0) sizes.push(s); }
    var st = String(r['Style Number'] || '').trim(), c = String(r['Color'] || '').trim();
    if (!smap[st]) { smap[st] = { style: st, colors: [], _c: {} }; styles.push(smap[st]); }
    if (c && !smap[st]._c[c.toUpperCase()]) { smap[st]._c[c.toUpperCase()] = 1; smap[st].colors.push({ color: c }); }
  });
  styles.forEach(function (st) {
    st.colors.forEach(function (c) { var cut = getTotalCutQtyForJO(joNumber, c.color, st.style); c.total = cut.total; c.bySize = cut.bySize; });
    var all = getTotalCutQtyForJO(joNumber, '', st.style); st.total = all.total; st.bySize = all.bySize; delete st._c;
  });
  return { jo: joNumber, buyer: String(rows[0]['Buyer Name'] || ''), styles: styles, sizes: sizes };
}
/* h = { joNumber, ihDate }; rows = [{ color, trim, size, qtyNeeded, perGarments, rate }]
   po = { make, supplier, gstNo, gstPct, poDate } */
function saveAccessoryPlanMulti(h, rows, po) {
  var jo = String(h.joNumber || '').trim();
  if (!jo) return { success: false, error: 'Pick the JO.' };
  var info = getJOInfo(jo);
  if (!info) return { success: false, error: 'JO ' + jo + ' not found in Style Tracker.' };
  rows = (rows || []).filter(function (r) { return String(r.trim || '').trim(); });
  if (!rows.length) return { success: false, error: 'Add at least one trim row.' };
  var U = function (x) { return String(x == null ? '' : x).trim().toUpperCase(); };
  var existing = sheetToObjects(SHEETS.ACCESSORIES_BOM).filter(function (r) { return U(r['JO Number']) === U(jo); });
  var keys = {};
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i], q = Number(r.qtyNeeded);
    if (!(q > 0)) return { success: false, error: 'Row ' + (i + 1) + ': qty per garment must be more than 0.' };
    if (q > 50 && !r.confirmBigQty) return { success: false, error: 'Row ' + (i + 1) + ': qty per garment ' + q + ' is too high (it is per ONE garment, usually 1).' };
    var k = U(r.style) + '|' + U(r.color) + '|' + U(r.trim) + '|' + U(r.size);
    if (keys[k]) return { success: false, error: 'Row ' + (i + 1) + ': ' + r.trim + ' is entered twice for the same style / colour / size.' };
    keys[k] = 1;
    if (existing.some(function (e) { return U(e['Style Number']) + '|' + U(e['Color']) + '|' + U(e['Trims / Accessories']) + '|' + U(e['Size']) === k; })) {
      return { success: false, error: 'Row ' + (i + 1) + ': ' + r.trim + ' is already planned for ' + jo + ' / ' + (r.color || 'all colours') + '. Edit that line instead.' };
    }
  }
  for (var j = 0; j < rows.length; j++) {
    var rr = rows[j];
    if (Number(rr.extraQty) > 0 && !String(rr.extraReason || '').trim()) return { success: false, error: 'Row ' + (j + 1) + ': give a reason for the extra qty.' };
    if (Number(rr.plannedQty) > 0) {
      var cut = getTotalCutQtyForJO(jo, rr.color, rr.style), pcs = rr.size ? (cut.bySize[rr.size] || 0) : cut.total;
      var auto = Math.ceil(pcs / (Number(rr.perGarments) || 1)) * (Number(rr.qtyNeeded) || 1);
      if (Number(rr.plannedQty) > auto) return { success: false, error: 'Row ' + (j + 1) + ': ' + rr.plannedQty + ' is more than the order needs (' + auto + '). Use "Extra" for more - it goes to Management.' };
    }
  }
  var exCols = accBomExtraCols_(), bomSh = SS.getSheetByName(SHEETS.ACCESSORIES_BOM), extraCount = 0;
  rows.forEach(function (r) {
    var rowNo = appendRowSafe_(SHEETS.ACCESSORIES_BOM, [jo, r.style || info['Style Number'] || '', info['Buyer'] || info['Buyer Name'] || '',
      r.color || '', String(r.trim).trim(), r.size || '', Number(r.qtyNeeded) || 1, Number(r.perGarments) || 1, Number(r.rate) || 0, h.ihDate || '']);
    if (Number(r.plannedQty) > 0) bomSh.getRange(rowNo, exCols['Planned Qty']).setValue(Number(r.plannedQty));
    if (Number(r.extraQty) > 0) {
      bomSh.getRange(rowNo, exCols['Extra Qty']).setValue(Number(r.extraQty));
      bomSh.getRange(rowNo, exCols['Extra Reason']).setValue(String(r.extraReason).trim());
      bomSh.getRange(rowNo, exCols['Extra Status']).setValue('Pending Approval');
      bomSh.getRange(rowNo, exCols['Extra By']).setValue(String(h.by || ''));
      extraCount++;
    }
  });
  var msg = rows.length + ' trim line(s) planned for ' + jo + '.' + (extraCount ? ' ' + extraCount + ' extra-qty request(s) sent to Management for approval.' : '');
  if (po && po.make) {
    if (!String(po.supplier || '').trim()) return { success: true, message: msg + ' PO not made: no supplier chosen.' };
    var bom = getAccessoriesBOM(jo), agg = {}, order = [];
    rows.forEach(function (r) {
      var line = bom.filter(function (b) { return U(b['Style Number']) === U(r.style) && U(b['Color']) === U(r.color) && U(b['Trims / Accessories']) === U(r.trim) && U(b['Size']) === U(r.size); })[0];
      if (!line) return;
      var key = U(r.trim) + '|' + U(r.size);
      if (!agg[key]) { agg[key] = { joNumber: jo, styleNo: line['Style Number'] || '', trimName: String(r.trim).trim(), size: r.size || '', qty: 0, rate: Number(r.rate) || 0 }; order.push(key); }
      agg[key].qty += Number(line['Required Quantity']) || 0;
    });
    var poRows = order.map(function (k) {
      var a = agg[k], left = a.qty;
      var req = bom.filter(function (b) { return U(b['Trims / Accessories']) === U(a.trimName) && U(b['Size']) === U(a.size); }).reduce(function (t, b) { return t + (Number(b['Required Quantity']) || 0); }, 0);
      var already = getAccessoryPOUsage(jo, a.trimName, a.size);
      a.qty = Math.min(left, Math.max(0, req - already));
      return a;
    }).filter(function (a) { return a.qty > 0; });
    if (poRows.length) {
      var p = saveAccessoryPOMulti({ supplier: po.supplier, gstNo: po.gstNo, gstPct: po.gstPct, poDate: po.poDate }, poRows);
      msg += p.success ? ' ' + p.message : ' PO not made: ' + p.error;
    } else msg += ' Nothing left to put on a PO.';
  }
  return { success: true, message: msg };
}

/*************** Accessories: planned qty (<= order) + extra qty (needs Management approval) ***************/
var ACC_BOM_EXTRA_HEADERS = ['Planned Qty', 'Extra Qty', 'Extra Reason', 'Extra Status', 'Extra By', 'Extra Approved By'];
function accBomExtraCols_() {
  var sh = SS.getSheetByName(SHEETS.ACCESSORIES_BOM), base = HEADERS[SHEETS.ACCESSORIES_BOM].length;
  var lastCol = Math.max(sh.getLastColumn(), base);
  var head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (x) { return String(x).trim(); });
  var out = {}, add = [];
  ACC_BOM_EXTRA_HEADERS.forEach(function (h) {
    var i = head.indexOf(h);
    if (i > -1) out[h] = i + 1; else { head.push(h); out[h] = head.length; add.push(h); }
  });
  if (add.length) {
    if (sh.getMaxColumns() < head.length) sh.insertColumnsAfter(sh.getMaxColumns(), head.length - sh.getMaxColumns());
    add.forEach(function (h) { sh.getRange(1, out[h]).setValue(h).setFontWeight('bold').setBackground('#1B2A4A').setFontColor('#FFFFFF'); });
  }
  return out;
}
function accBomExtrasByRow_() {
  var sh = SS.getSheetByName(SHEETS.ACCESSORIES_BOM);
  if (!sh || sh.getLastRow() < 2) return {};
  var lastCol = sh.getLastColumn();
  var head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (x) { return String(x).trim(); });
  var ci = {}; ACC_BOM_EXTRA_HEADERS.forEach(function (h) { ci[h] = head.indexOf(h); });
  if (ci['Extra Qty'] < 0 && ci['Planned Qty'] < 0) return {};
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues(), out = {};
  v.forEach(function (r, i) {
    var g = function (h) { return ci[h] > -1 ? r[ci[h]] : ''; };
    if (g('Planned Qty') !== '' || g('Extra Qty') !== '') out[i + 2] = { planned: g('Planned Qty'), extra: g('Extra Qty'), reason: g('Extra Reason'), status: String(g('Extra Status') || ''), by: g('Extra By') };
  });
  return out;
}
function getPendingAccExtras_() {
  var ex = accBomExtrasByRow_(), out = [];
  sheetToObjects(SHEETS.ACCESSORIES_BOM).forEach(function (r) {
    var e = ex[r._row]; if (!e || e.status !== 'Pending Approval') return;
    out.push({ row: r._row, jo: r['JO Number'], style: r['Style Number'], color: r['Color'], trim: r['Trims / Accessories'], size: r['Size'],
      extra: Number(e.extra) || 0, reason: e.reason, by: e.by, rate: Number(r['Rate (Rs/unit)']) || 0 });
  });
  return out;
}
/* request extra qty on an existing planning line (goes to Management) */
function requestAccessoryExtra(k, extraQty, reason, by) {
  var hit = accBomFindRow_(k);
  if (!hit) return { success: false, error: 'Planning line not found.' };
  if (!(Number(extraQty) > 0)) return { success: false, error: 'Enter the extra qty.' };
  if (!String(reason || '').trim()) return { success: false, error: 'Give a reason for the extra qty.' };
  var c = accBomExtraCols_();
  hit.sh.getRange(hit.row, c['Extra Qty']).setValue(Number(extraQty));
  hit.sh.getRange(hit.row, c['Extra Reason']).setValue(String(reason).trim());
  hit.sh.getRange(hit.row, c['Extra Status']).setValue('Pending Approval');
  hit.sh.getRange(hit.row, c['Extra By']).setValue(String(by || ''));
  hit.sh.getRange(hit.row, c['Extra Approved By']).setValue('');
  return { success: true, message: 'Extra ' + extraQty + ' of ' + k.trim + ' sent to Management for approval.' };
}
function approveAccessoryExtra(row, decision, approver) {
  var sh = SS.getSheetByName(SHEETS.ACCESSORIES_BOM), c = accBomExtraCols_();
  if (!sh || row < 2 || row > sh.getLastRow()) return { success: false, error: 'Line not found.' };
  var status = decision === 'APPROVE' ? 'Approved' : 'Rejected';
  sh.getRange(row, c['Extra Status']).setValue(status);
  sh.getRange(row, c['Extra Approved By']).setValue(String(approver || '') + ' ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  return { success: true, message: 'Extra qty ' + status.toLowerCase() + '.' };
}

/*************** YARN EXTRA REQUESTS (beyond required kg — needs Management approval) ***************/
var YARN_EXTRA_SHEET = 'YARN EXTRA REQUESTS';
var YARN_EXTRA_HEAD = ['JO Number', 'Counts', 'Content', 'GSM', 'Color', 'Yarn Color', 'Extra KG', 'Reason', 'Status', 'Requested By', 'Approved By', 'Approved At'];
function yarnExtraSheet_() {
  var sh = SS.getSheetByName(YARN_EXTRA_SHEET);
  if (!sh) {
    sh = SS.insertSheet(YARN_EXTRA_SHEET);
    sh.getRange(1, 1, 1, YARN_EXTRA_HEAD.length).setValues([YARN_EXTRA_HEAD]).setFontWeight('bold').setBackground('#1B2A4A').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }
  return sh;
}
function yarnExtraKey_(k) { return [k.joNumber, k.counts, k.content || '', k.gsm || '', k.color || '', k.yarnColor || ''].map(function (x) { return String(x || '').trim().toUpperCase(); }).join('|'); }
function yarnExtraRows_() {
  var sh = SS.getSheetByName(YARN_EXTRA_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, YARN_EXTRA_HEAD.length).getValues();
  return v.map(function (r, i) {
    var o = { _row: i + 2 }; YARN_EXTRA_HEAD.forEach(function (h, j) { o[h] = r[j]; }); return o;
  });
}
/* net approved-but-unused extra kg for one spec (approved total minus what's already been put on a Yarn PO as "extra") */
function getYarnExtraApproved(k) {
  var key = yarnExtraKey_(k), total = 0;
  yarnExtraRows_().forEach(function (r) { if (r['Status'] === 'Approved' && yarnExtraKey_({ joNumber: r['JO Number'], counts: r['Counts'], content: r['Content'], gsm: r['GSM'], color: r['Color'], yarnColor: r['Yarn Color'] }) === key) total += Number(r['Extra KG']) || 0; });
  return total;
}
function getPendingYarnExtras_() {
  return yarnExtraRows_().filter(function (r) { return r['Status'] === 'Pending Approval'; })
    .map(function (r) { return { row: r._row, jo: r['JO Number'], counts: r['Counts'], content: r['Content'], gsm: r['GSM'], color: r['Color'], yarnColor: r['Yarn Color'], extra: Number(r['Extra KG']) || 0, reason: r['Reason'], by: r['Requested By'] }; });
}
function requestYarnExtra(k, extraKg, reason, by) {
  if (!k.joNumber || !k.counts) return { success: false, error: 'JO / Counts missing.' };
  if (!(Number(extraKg) > 0)) return { success: false, error: 'Enter the extra kg.' };
  if (!String(reason || '').trim()) return { success: false, error: 'Give a reason for the extra yarn.' };
  var sh = yarnExtraSheet_();
  sh.appendRow([k.joNumber, k.counts, k.content || '', k.gsm || '', k.color || '', k.yarnColor || '', Number(extraKg), String(reason).trim(), 'Pending Approval', String(by || ''), '', '']);
  return { success: true, message: 'Extra ' + extraKg + ' kg requested for ' + k.joNumber + ' / ' + k.counts + ' — sent to Management for approval.' };
}
function approveYarnExtra(row, decision, approver) {
  var sh = SS.getSheetByName(YARN_EXTRA_SHEET);
  if (!sh || row < 2 || row > sh.getLastRow()) return { success: false, error: 'Request not found.' };
  var status = decision === 'APPROVE' ? 'Approved' : 'Rejected';
  var col = function (h) { return YARN_EXTRA_HEAD.indexOf(h) + 1; };
  sh.getRange(row, col('Status')).setValue(status);
  sh.getRange(row, col('Approved By')).setValue(String(approver || ''));
  sh.getRange(row, col('Approved At')).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  return { success: true, message: 'Extra yarn request ' + status.toLowerCase() + '.' };
}

/*************** YARN BILL PASSING (supplier's invoice against Approved Yarn POs) ***************/
function getYarnBillSuppliers() {
  var ready = {};
  sheetToObjects(SHEETS.YARN_PO).forEach(function (r) {
    if (r['Status'] !== 'Approved' && r['Status'] !== 'Billed') return;
    var s = String(r['Supplier'] || '').trim(); if (!s) return;
    ready[s] = (ready[s] || 0) + (r['Status'] === 'Approved' ? 1 : 0);
  });
  return Object.keys(ready).sort().map(function (s) { return { name: s, ready: ready[s] || 0 }; });
}
function getBillableYarnPOsForSupplier(supplier) {
  var billed = {};
  sheetToObjects(SHEETS.YARN_BILL).forEach(function (b) {
    if (b['Status'] === 'Rejected') return;
    String(b['PO Numbers'] || '').split(',').map(function (x) { return x.trim(); }).forEach(function (n) { if (n) billed[n] = true; });
  });
  return sheetToObjects(SHEETS.YARN_PO).filter(function (r) {
    return r['Status'] === 'Approved' && String(r['Supplier'] || '').trim() === String(supplier).trim() && !billed[r['PO No']];
  });
}
function saveYarnBill(d) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var poNos = (d.poNos || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean);
    if (!poNos.length) return { success: false, error: 'Tick at least one Yarn PO.' };
    var already = {};
    sheetToObjects(SHEETS.YARN_BILL).forEach(function (b) {
      if (b['Status'] === 'Rejected') return;
      String(b['PO Numbers'] || '').split(',').map(function (x) { return x.trim(); }).forEach(function (n) { if (n) already[n] = true; });
    });
    var allPOs = sheetToObjects(SHEETS.YARN_PO), joSet = {}, totalQty = 0, totalExpected = 0;
    for (var i = 0; i < poNos.length; i++) {
      var po = allPOs.filter(function (r) { return r['PO No'] === poNos[i]; })[0];
      if (!po) return { success: false, error: 'PO "' + poNos[i] + '" not found.' };
      if (po['Status'] !== 'Approved') return { success: false, error: 'PO ' + poNos[i] + ' is "' + po['Status'] + '" — only an Approved PO can be billed.' };
      if (String(po['Supplier'] || '').trim() !== String(d.supplier).trim()) return { success: false, error: 'PO ' + poNos[i] + ' belongs to ' + po['Supplier'] + ', not ' + d.supplier + '.' };
      if (already[poNos[i]]) return { success: false, error: 'PO ' + poNos[i] + ' has already been billed (or is pending billing).' };
      var qty = Number(po['Qty KG']) || 0, rate = Number(po['Rate (Rs/kg)']) || 0;
      joSet[po['JO Number']] = true; totalQty += qty; totalExpected += qty * rate;
    }
    var billAmount = Number(d.billAmount) || 0;
    if (billAmount > totalExpected + 0.01) return { success: false, error: 'Bill Amount (₹' + billAmount.toFixed(2) + ') cannot exceed the combined PO value (₹' + totalExpected.toFixed(2) + ').' };
    var variance = billAmount - totalExpected;
    var blendedRate = totalQty > 0 ? (totalExpected / totalQty) : 0;
    var gstPct = Number(d.gstPct) || 0, gstAmount = billAmount * gstPct / 100, totalInclGst = billAmount + gstAmount;
    var billStatedTotal = d.billStatedTotal !== undefined && d.billStatedTotal !== '' ? Number(d.billStatedTotal) : null;
    var gstMatch = billStatedTotal !== null ? (Math.abs(totalInclGst - billStatedTotal) <= 0.5 ? 'Match' : 'Mismatch') : '';
    var billNo = nextId('YBILL-', SHEETS.YARN_BILL, 1);
    appendRowSafe_(SHEETS.YARN_BILL, [
      billNo, d.billDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      d.supplier, Object.keys(joSet).join(', '), poNos.join(', '), totalQty, blendedRate, totalExpected, billAmount, variance,
      d.remarks || '', 'Pending Approval', '', '', '', '',
      gstPct, gstAmount, totalInclGst, billStatedTotal === null ? '' : billStatedTotal, gstMatch, d.supplierInvoiceNo || ''
    ]);
    return { success: true, billNo: billNo, message: 'Yarn Bill ' + billNo + ' submitted for approval — ' + poNos.length + ' PO(s) across ' + Object.keys(joSet).length + ' JO(s). Expected ₹' + totalExpected.toFixed(2) + ', variance ₹' + variance.toFixed(2) + '.' };
  } catch (err) {
    return { success: false, error: err.message };
  } finally { lock.releaseLock(); }
}
function approveYarnBill(billNo, decision, approver, reason) {
  var rowIdx = findRowIndex(SHEETS.YARN_BILL, 'Bill No', billNo);
  if (rowIdx < 0) return { success: false, error: 'Bill not found.' };
  var status = decision === 'APPROVE' ? 'Approved' : 'Rejected';
  setCell(SHEETS.YARN_BILL, rowIdx, 'Status', status);
  setCell(SHEETS.YARN_BILL, rowIdx, 'Approved By', approver || '');
  setCell(SHEETS.YARN_BILL, rowIdx, 'Approval Date', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  setCell(SHEETS.YARN_BILL, rowIdx, 'Rejection Reason', decision === 'REJECT' ? (reason || '') : '');
  if (status === 'Approved') {
    var bill = sheetToObjects(SHEETS.YARN_BILL).filter(function (r) { return r['Bill No'] === billNo; })[0];
    String(bill['PO Numbers'] || '').split(',').map(function (x) { return x.trim(); }).forEach(function (n) {
      if (!n) return;
      sheetToObjects(SHEETS.YARN_PO).forEach(function (po) {
        if (String(po['PO No']).trim() === n && String(po['Status']).trim() === 'Approved') setCell(SHEETS.YARN_PO, po._row, 'Status', 'Billed');
      });
    });
  }
  return { success: true, message: 'Yarn Bill ' + billNo + ' ' + status + '.' };
}

/*************** CMT BILL PASSING (garmenting / job-work wages — several styles in one bill) ***************/
function getCmtContractors(department) {
  var all = sheetToObjects(SHEETS.SUPPLIERS);
  if (String(department || '').trim()) {
    var list = suppliersTaggedWith_(all, department);   // e.g. Process = "Stitching" or "Stitching, Checking"
    if (list.length) return list;
  }
  var list = suppliersTaggedWith_(all, 'CMT');
  if (!list.length) list = suppliersTaggedWith_(all, 'Job Work');
  if (!list.length) list = all.map(function (s) { return s['Supplier Name']; });
  return list;
}
function suppliersTaggedWith_(suppliers, tag) {
  var want = String(tag).toLowerCase();
  return suppliers.filter(function (s) {
    return String(s['Process'] || '').split(',').map(function (p) { return p.trim().toLowerCase(); })
      .some(function (p) { return p === want || p.indexOf(want) !== -1; });
  }).map(function (s) { return s['Supplier Name']; });
}
/* pieces already billed (any status except Rejected) for one JO+Style+Colour+Department */
function getCmtBilledQty(jo, style, color, dept) {
  var U = function (x) { return String(x == null ? '' : x).trim().toUpperCase(); };
  return sheetToObjects(SHEETS.CMT_BILL).filter(function (r) {
    return r['Status'] !== 'Rejected' && U(r['JO Number']) === U(jo) && U(r['Style Number']) === U(style) &&
      U(r['Color']) === U(color) && U(r['Department']) === U(dept);
  }).reduce(function (t, r) { return t + (Number(r['Qty']) || 0); }, 0);
}
/* h = { billDate, department, contractor }; rows = [{ jo, style, color, qty, rate }] */
function saveCmtBillMulti(h, rows) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    if (!String(h.contractor || '').trim()) return { success: false, error: 'Choose the contractor.' };
    if (!String(h.department || '').trim()) return { success: false, error: 'Choose the department / process.' };
    rows = (rows || []).filter(function (r) { return r.jo && r.style && Number(r.qty) > 0; });
    if (!rows.length) return { success: false, error: 'Add at least one row with style and qty.' };
    var need = {};
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var cut = getTotalCutQtyForJO(r.jo, r.color, r.style);
      var order = cut.total;
      var key = r.jo + '|' + r.style + '|' + (r.color || '') + '|' + h.department;
      if (!need[key]) { var billed = getCmtBilledQty(r.jo, r.style, r.color, h.department); need[key] = { order: order, billed: billed, qty: 0, label: r.jo + ' / ' + r.style + (r.color ? ' / ' + r.color : '') }; }
      need[key].qty += Number(r.qty) || 0;
    }
    var over = Object.keys(need).filter(function (k) { return need[k].qty > need[k].order - need[k].billed + 0.5; })[0];
    if (over) {
      var n = need[over];
      return { success: false, error: n.label + ' (' + h.department + '): ' + n.qty + ' entered, but only ' + Math.max(0, n.order - n.billed) + ' pcs left to bill (order ' + n.order + ', already billed ' + n.billed + ').' };
    }
    var gcol = cmtGroupCol_(), base = nextId('CMT-', SHEETS.CMT_BILL, 1), date = h.billDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var sh = SS.getSheetByName(SHEETS.CMT_BILL), nos = [];
    rows.forEach(function (r, i) {
      var qty = Number(r.qty) || 0, rate = Number(r.rate) || 0, no = i === 0 ? base : dcxRowId_(base, i);
      var rowNo = appendRowSafe_(SHEETS.CMT_BILL, [no, date, h.department, h.contractor, r.jo, r.style, r.color || '', qty, rate, qty * rate, h.remarks || '', 'Pending Approval', '', '', '', '', '']);
      if (rows.length > 1) sh.getRange(rowNo, gcol).setValue(base);
      nos.push(no);
    });
    return { success: true, billNo: base, message: 'CMT Bill ' + base + (nos.length > 1 ? ' with ' + nos.length + ' lines' : '') + ' submitted — awaiting approval.' };
  } catch (err) {
    return { success: false, error: err.message };
  } finally { lock.releaseLock(); }
}
function cmtGroupCol_() {
  var sh = SS.getSheetByName(SHEETS.CMT_BILL), base = HEADERS[SHEETS.CMT_BILL].length;
  return base; // 'Bill Group' is already the last header column
}
function cmtGroupRows_(billNo) {
  var rows = sheetToObjects(SHEETS.CMT_BILL);
  var me = rows.filter(function (r) { return r['Bill No'] === billNo; })[0];
  if (!me) return [];
  var grp = String(me['Bill Group'] || '').trim();
  if (!grp) return [me];
  return rows.filter(function (r) { return String(r['Bill Group'] || '').trim() === grp; });
}
function approveCmtBill(billNo, decision, approver, reason) {
  var status = decision === 'APPROVE' ? 'Approved' : 'Rejected';
  var grp = cmtGroupRows_(billNo);
  if (!grp.length) return { success: false, error: 'Bill not found.' };
  grp.forEach(function (r) {
    var rowIdx = findRowIndex(SHEETS.CMT_BILL, 'Bill No', r['Bill No']);
    if (rowIdx < 0) return;
    setCell(SHEETS.CMT_BILL, rowIdx, 'Status', status);
    setCell(SHEETS.CMT_BILL, rowIdx, 'Approved By', approver || '');
    setCell(SHEETS.CMT_BILL, rowIdx, 'Approval Date', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
    setCell(SHEETS.CMT_BILL, rowIdx, 'Rejection Reason', decision === 'REJECT' ? (reason || '') : '');
  });
  return { success: true, message: 'CMT Bill ' + billNo + (grp.length > 1 ? ' (' + grp.length + ' lines)' : '') + ' ' + status + '.' };
}
function getCmtBillsForJO(jo) {
  return sheetToObjects(SHEETS.CMT_BILL).filter(function (r) { return !jo || String(r['JO Number']).trim() === String(jo).trim(); });
}

/*************** JO TOTAL SPEND (yarn + process/job-work + accessories + CMT, Approved only) ***************/
function getJOSpendSummary(jo) {
  var U = function (x) { return String(x == null ? '' : x).trim().toUpperCase(); };
  var sum = function (sheet, filterFn, amtField) {
    return sheetToObjects(sheet).filter(function (r) { return U(r['JO Number']) === U(jo) && r['Status'] === 'Approved' && (!filterFn || filterFn(r)); })
      .reduce(function (t, r) { return t + (Number(r[amtField]) || 0); }, 0);
  };
  // multi-JO bills (yarn / process) list several JOs comma-separated in one row — split evenly isn't
  // accurate, so instead count the PO/DC-level value that belongs to this JO specifically.
  var yarnSpend = 0;
  sheetToObjects(SHEETS.YARN_BILL).forEach(function (b) {
    if (b['Status'] !== 'Approved') return;
    var jos = String(b['JO Number'] || '').split(',').map(function (x) { return x.trim(); });
    if (jos.indexOf(String(jo).trim()) === -1) return;
    var poNos = String(b['PO Numbers'] || '').split(',').map(function (x) { return x.trim(); });
    var poVal = 0, myVal = 0;
    poNos.forEach(function (n) {
      var po = sheetToObjects(SHEETS.YARN_PO).filter(function (r) { return r['PO No'] === n; })[0]; if (!po) return;
      var v = (Number(po['Qty KG']) || 0) * (Number(po['Rate (Rs/kg)']) || 0);
      poVal += v; if (U(po['JO Number']) === U(jo)) myVal += v;
    });
    if (poVal > 0) yarnSpend += (Number(b['Bill Amount']) || 0) * (myVal / poVal);
  });
  var processSpend = 0;
  sheetToObjects(SHEETS.BILL_PASSING).forEach(function (b) {
    if (b['Status'] !== 'Approved') return;
    var jos = String(b['JO Number'] || '').split(',').map(function (x) { return x.trim(); });
    if (jos.indexOf(String(jo).trim()) === -1) return;
    var dcNos = String(b['DC Numbers'] || '').split(',').map(function (x) { return x.trim(); });
    var dcVal = 0, myVal = 0;
    dcNos.forEach(function (n) {
      var dc = sheetToObjects(SHEETS.PROCESS_DC).filter(function (r) { return r['DC No'] === n; })[0]; if (!dc) return;
      var v = (Number(dc['DC Issue Qty KG']) || 0) * (Number(b['Applicable Rate']) || 0);
      dcVal += v; if (U(dc['JO Number']) === U(jo)) myVal += v;
    });
    if (dcVal > 0) processSpend += (Number(b['Bill Amount']) || 0) * (myVal / dcVal);
  });
  var accSpend = sum(SHEETS.ACCESSORIES_BILL, null, 'Bill Amount');
  var cmtSpend = sum(SHEETS.CMT_BILL, null, 'Amount');
  var generalSpend = sum(SHEETS.GENERAL_BILL, null, 'Total Incl GST');
  var total = yarnSpend + processSpend + accSpend + cmtSpend + generalSpend;
  return { jo: jo, yarn: yarnSpend, process: processSpend, accessories: accSpend, cmt: cmtSpend, general: generalSpend, total: total };
}

function getYarnBillRows() { return sheetToObjects(SHEETS.YARN_BILL); }

/*************** CMT RATES (per JO, per department) — feeds Fabric Budget Entry + CMT Bill default rate ***************/
var CMT_DEPARTMENTS = ['Cutting', 'Stitching', 'Checking', 'Ironing', 'Packing', 'Embroidery', 'Printing', 'Other'];
/* returns { Department: { costing: <rate quoted at costing time, never changes once set>, rate: <current budget rate> } } */
function getCmtRatesForJO(jo) {
  var out = {};
  sheetToObjects(SHEETS.CMT_RATES).forEach(function (r) {
    if (String(r['JO Number']).trim() === String(jo).trim()) {
      out[r['Department']] = { costing: Number(r['Costing Rate Per Pc']) || 0, rate: Number(r['Rate Per Pc']) || 0 };
    }
  });
  return out;
}
/* rates = { Department: { costing, rate } }. A costing rate, once saved, is never overwritten from here —
   it's the number you quoted the buyer with, kept as a fixed reference point. Only the current
   ("actual/budget") rate can be revised. Send costing only when setting it for the first time. */
function saveCmtRates(jo, rates, by) {
  if (!jo) return { success: false, error: 'JO missing.' };
  var sh = SS.getSheetByName(SHEETS.CMT_RATES);
  var existing = sheetToObjects(SHEETS.CMT_RATES).filter(function (r) { return String(r['JO Number']).trim() === String(jo).trim(); });
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  Object.keys(rates || {}).forEach(function (dept) {
    var v = rates[dept] || {}, rate = Number(v.rate) || 0, costing = Number(v.costing) || 0;
    var row = existing.filter(function (r) { return r['Department'] === dept; })[0];
    if (rate <= 0 && costing <= 0) { if (row) sh.deleteRow(row._row); return; }
    if (row) {
      var keepCosting = (Number(row['Costing Rate Per Pc']) || 0) || costing;   // never blank out an existing costing rate
      sh.getRange(row._row, 3, 1, 4).setValues([[keepCosting, rate, by || '', now]]);
    } else {
      appendRowSafe_(SHEETS.CMT_RATES, [jo, dept, costing, rate, by || '', now]);
    }
  });
  return { success: true, message: 'CMT rates saved for ' + jo + '.' };
}
/* CMT budget estimate for the whole JO: each department's CURRENT rate x the JO's total order pcs
   (incl. excess), alongside the same total at the ORIGINAL COSTING rate, for comparison. */
function getCmtBudgetForJO(jo) {
  var rates = getCmtRatesForJO(jo);
  var pcs = getTotalCutQtyForJO(jo).total;
  var lines = Object.keys(rates).map(function (d) {
    return { department: d, costing: rates[d].costing, rate: rates[d].rate, pcs: pcs, costingAmount: rates[d].costing * pcs, amount: rates[d].rate * pcs };
  });
  var total = lines.reduce(function (t, l) { return t + l.amount; }, 0);
  var costingTotal = lines.reduce(function (t, l) { return t + l.costingAmount; }, 0);
  return { pcs: pcs, lines: lines, total: total, costingTotal: costingTotal, variance: total - costingTotal };
}

/*************** GENERAL BILL PASSING (any purchase / payment — against a JO, or general overhead) ***************/
var GENERAL_BILL_CATEGORIES = ['Transport / Freight', 'Packing Material', 'Electricity', 'Repairs & Maintenance',
  'Consultancy / Professional Fees', 'Rent', 'Office / Admin', 'Testing / Certification', 'Miscellaneous'];
function saveGeneralBill(d) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    if (!String(d.paidTo || '').trim()) return { success: false, error: 'Enter who this was paid to.' };
    var amount = Number(d.amount) || 0;
    if (amount <= 0) return { success: false, error: 'Enter the amount.' };
    var gstPct = Number(d.gstPct) || 0, gstAmount = amount * gstPct / 100, totalInclGst = amount + gstAmount;
    var billNo = nextId('GBILL-', SHEETS.GENERAL_BILL, 1);
    var photoUrl = d.photoB64 ? dcPhotoSave_(d.photoB64, d.photoMime, d.photoName, billNo) : '';
    appendRowSafe_(SHEETS.GENERAL_BILL, [
      billNo, d.billDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      d.category || 'Miscellaneous', d.paidTo, d.joNumber || '', d.description || '', amount,
      gstPct, gstAmount, totalInclGst, d.remarks || '', 'Pending Approval', '', '', '', photoUrl
    ]);
    return { success: true, billNo: billNo, message: 'Bill ' + billNo + ' (' + (d.joNumber ? d.joNumber : 'general, no JO') + ') submitted for approval — ₹' + totalInclGst.toFixed(2) + ' incl. GST.' };
  } catch (err) {
    return { success: false, error: err.message };
  } finally { lock.releaseLock(); }
}
function approveGeneralBill(billNo, decision, approver, reason) {
  var rowIdx = findRowIndex(SHEETS.GENERAL_BILL, 'Bill No', billNo);
  if (rowIdx < 0) return { success: false, error: 'Bill not found.' };
  var status = decision === 'APPROVE' ? 'Approved' : 'Rejected';
  setCell(SHEETS.GENERAL_BILL, rowIdx, 'Status', status);
  setCell(SHEETS.GENERAL_BILL, rowIdx, 'Approved By', approver || '');
  setCell(SHEETS.GENERAL_BILL, rowIdx, 'Approval Date', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  setCell(SHEETS.GENERAL_BILL, rowIdx, 'Rejection Reason', decision === 'REJECT' ? (reason || '') : '');
  return { success: true, message: 'Bill ' + billNo + ' ' + status + '.' };
}
function getGeneralBillRows() { return sheetToObjects(SHEETS.GENERAL_BILL); }
function generateGeneralBillHtml_(b) {
  return pdfShell('General Bill', '#1a5276',
    '<div class="title-bar">Bill #: <strong>' + b['Bill No'] + '</strong> | Date: ' + b['Date'] + ' | Status: ' + b['Status'] + '</div>' +
    '<table><tr><th>Category</th><td>' + (b['Category'] || '-') + '</td><th>Paid To</th><td>' + (b['Paid To'] || '-') + '</td></tr>' +
    '<tr><th>JO Number</th><td>' + (b['JO Number'] || 'General (no JO)') + '</td><th>Description</th><td>' + (b['Description'] || '-') + '</td></tr>' +
    '<tr><th>Amount</th><td>₹' + Number(b['Amount'] || 0).toFixed(2) + '</td><th>GST %</th><td>' + Number(b['GST %'] || 0).toFixed(2) + '</td></tr>' +
    '<tr><th>GST Amount</th><td>₹' + Number(b['GST Amount'] || 0).toFixed(2) + '</td><th>Total incl. GST</th><td><strong>₹' + Number(b['Total Incl GST'] || 0).toFixed(2) + '</strong></td></tr>' +
    '<tr><th>Approved By</th><td>' + (b['Approved By'] || '-') + '</td><th>Approval Date</th><td>' + (b['Approval Date'] || '-') + '</td></tr>' +
    '<tr><th>Remarks</th><td colspan="3">' + (b['Remarks'] || '-') + '</td></tr></table>');
}

/*************** REPORTS: fabric status, yarn needed (drill-down), yarn PO pending ***************/

/* One row per JO: fabric progress at every stage (kg), for the "Fabric Status" report. */
function getFabricStatusReport() {
  var oh = getOrderInHand();
  return Object.keys(oh.jos).map(function (jo) {
    var j = oh.jos[jo];
    var link = stImgLinkForJo_(jo), m = /\/d\/([A-Za-z0-9_-]{20,})/.exec(link) || /[?&]id=([A-Za-z0-9_-]{20,})/.exec(link) || /^([A-Za-z0-9_-]{25,})$/.exec(link);
    return { jo: jo, buyer: j.buyer, style: (oh.lines.filter(function (l) { return l.jo === jo; }).map(function (l) { return l.style; })
        .filter(function (v, i, a) { return a.indexOf(v) === i; }).join(', ')),
      fabricKg: j.fabricKg, yarnPoKg: j.yarnPoKg, yarnRcvKg: j.yarnRcvKg, knitKg: j.knitKg, dyeKg: j.dyeKg, inwardKg: j.inwardKg, stage: j.stage,
      photoThumb: m ? 'https://drive.google.com/thumbnail?id=' + m[1] + '&sz=w160' : (/^cell:/.test(link) ? oihThumbFor_(jo, link) : '') };
  }).sort(function (a, b) { return a.jo.localeCompare(b.jo, undefined, { numeric: true }); });
}

/* Yarn needed, drill-down: Counts -> Buyers -> Styles. Uses the same main-fibre/lycra split as
   Yarn PO (dcxLycra_), across every JO in Style Tracker. */
/* "30's", "30s", "30", "30S" all mean the same yarn count — normalize to "30's" for grouping/display.
   A genuine spinning-type suffix (OE, Combed, Carded...) is kept, e.g. "30's OE" stays distinct. */
function normalizeCounts_(raw) {
  var s = String(raw || '').trim();
  var m = /^(\d+(?:\.\d+)?)/.exec(s);
  if (!m) return s;
  var rest = s.slice(m[0].length).replace(/^['’]?s\b/i, '').trim();
  return m[1] + "'s" + (rest ? ' ' + rest : '');
}
function getYarnNeedReport() {
  var rows = getStyleTracker();
  var countsMap = {}, countsOrder = [];
  rows.forEach(function (r) {
    var counts = normalizeCounts_(r['Counts']);
    if (!counts) return;
    var ly = dcxLycra_(r), kg = ly.pure ? Number(r['Fabric Required (kg)']) || 0 : ly.mainKg;
    if (!(kg > 0)) return;
    var buyer = String(r['Buyer Name'] || '').trim() || '(no buyer)';
    var style = String(r['Style Number'] || '').trim() || '(no style)';
    var jo = String(r['JO Number'] || '').trim();
    if (!countsMap[counts]) { countsMap[counts] = { counts: counts, totalKg: 0, buyers: {}, buyerOrder: [] }; countsOrder.push(counts); }
    var c = countsMap[counts]; c.totalKg += kg;
    if (!c.buyers[buyer]) { c.buyers[buyer] = { buyer: buyer, totalKg: 0, styles: {}, styleOrder: [] }; c.buyerOrder.push(buyer); }
    var b = c.buyers[buyer]; b.totalKg += kg;
    if (!b.styles[style]) { b.styles[style] = { style: style, jo: jo, totalKg: 0 }; b.styleOrder.push(style); }
    b.styles[style].totalKg += kg;
  });
  return countsOrder.map(function (ck) {
    var c = countsMap[ck];
    return {
      counts: c.counts, totalKg: Math.round(c.totalKg * 1000) / 1000,
      buyers: c.buyerOrder.map(function (bk) {
        var b = c.buyers[bk];
        return { buyer: b.buyer, totalKg: Math.round(b.totalKg * 1000) / 1000,
          styles: b.styleOrder.map(function (sk) { var st = b.styles[sk]; return { style: st.style, jo: st.jo, totalKg: Math.round(st.totalKg * 1000) / 1000 }; }) };
      })
    };
  }).sort(function (a, b) { return b.totalKg - a.totalKg; });
}

/* Every JO + Counts spec where a Yarn PO is still needed (required > what's already on PO). */
/* Distinct JO numbers that still have a pending Yarn PO balance on at least one Counts/Content spec —
   used to hide JOs that are already fully covered from the "Create Purchase Order" JO picker, so a
   JO that's done doesn't clutter the list once nothing more needs to be raised for it. */
function getJosNeedingYarnPO() {
  return withSheetCache_(function () {
  var seen = {}, out = [];
  getYarnPoPendingReport().forEach(function (r) { if (!seen[r.jo]) { seen[r.jo] = true; out.push(r.jo); } });
  return out;
  });
}
/* JOs/specs that still need a Yarn PO raised. Every PO type counts as cover — a Yarn PO, a Cora Fabric PO
   and a Dyed Fabric PO all reduce what is still to be ordered for that JO's Counts + Content.
   A spec is done once what is left is less than one bag of what has already been ordered: POs go in whole
   bags, so 3 kg short on a 3,293 kg requirement is not "pending". Nothing ordered yet always shows. */
function getYarnPoPendingReport() {
  return withSheetCache_(function () {
  var rows = getStyleTracker(), groups = {}, order = [];
  rows.forEach(function (r) {
    var jo = String(r['JO Number'] || '').trim(), counts = String(r['Counts'] || '').trim();
    if (!jo || !counts) return;
    var content = String(r['Content'] || '').trim(), key = jo + '|' + counts + '|' + content;
    if (!groups[key]) { groups[key] = { jo: jo, counts: counts, content: content, buyer: String(r['Buyer Name'] || '').trim(), styles: [] }; order.push(key); }
    var st = String(r['Style Number'] || '').trim();
    if (st && groups[key].styles.indexOf(st) < 0) groups[key].styles.push(st);
  });

  var specsByJo = {}, poRows = sheetToObjects(SHEETS.YARN_PO), out = [];
  order.forEach(function (key) {
    var g = groups[key];
    if (!specsByJo[g.jo]) specsByJo[g.jo] = getCountsForJO(g.jo, '', false);
    // THIS spec's own requirement: same Counts AND same Content (not every content under that count)
    var required = 0;
    specsByJo[g.jo].forEach(function (c) {
      if (String(c.counts || '').trim() === g.counts && String(c.content || '').trim() === g.content) required += Number(c.requiredKg) || 0;
    });
    if (!(required > 0)) return;
    var already = getYarnPOUsage(g.jo, g.counts, '', g.content, '', '');
    var extra = getYarnExtraApproved({ joNumber: g.jo, counts: g.counts, content: g.content, gsm: '', color: '', yarnColor: '' });
    var balance = required + extra - already;
    if (!(balance > 0.5)) return;

    var bagKg = 0;
    poRows.forEach(function (p) {
      if (bagKg > 0 || p['Status'] === 'Rejected' || String(p['JO Number']).trim() !== g.jo) return;
      if (String(p['Counts'] || '').trim() === g.counts && String(p['Content'] || '').trim() === g.content) bagKg = Number(p['Bag Weight (KG)']) || 0;
    });
    if (!(bagKg > 0)) bagKg = 50;
    if (already > 0.5 && balance < bagKg) return;      // within one bag of fully covered

    out.push({ jo: g.jo, buyer: g.buyer, style: g.styles.join(', '),
      counts: normalizeCounts_(g.counts), content: g.content, requiredKg: Math.round(required * 1000) / 1000,
      onPoKg: Math.round(already * 1000) / 1000, balanceKg: Math.round(balance * 1000) / 1000 });
  });
  return out.sort(function (a, b) { return b.balanceKg - a.balanceKg; });
  });
}

/* PO made vs Inward — one row per released (Approved / Billed) PO line that is still waiting for material,
   Yarn, Cora Fabric and Dyed Fabric alike. Received is matched the same way the Yarn Inward screen matches
   it (PO No + JO + Counts + Content + GSM), so this agrees with the "remaining" shown when receiving.
   Oldest PO first, so the longest-waiting ones are at the top. */
function getPoVsInwardReport() {
  return withSheetCache_(function () {
  var norm = function (v) { return String(v == null ? '' : v).trim().toLowerCase(); };
  var cnt = function (v) { try { return norm(normalizeCounts_(v)); } catch (e) { return norm(v); } };
  var gsmN = function (v) { var n = Number(v); return isNaN(n) || v === '' ? norm(v) : String(n); };
  var key = function (po, jo, counts, content, gsm) { return [norm(po), norm(jo), cnt(counts), norm(content), gsmN(gsm)].join('|'); };

  var received = {}, lastIn = {};
  sheetToObjects(SHEETS.YARN_INWARD).forEach(function (r) {
    var k = key(r['Yarn PO No'], r['JO Number'], r['Counts'], r['Content'], r['GSM']);
    received[k] = (received[k] || 0) + (Number(r['Received KG']) || 0);
    var d = String(r['Date'] || '').slice(0, 10);
    if (d && (!lastIn[k] || d > lastIn[k])) lastIn[k] = d;
  });

  var today = new Date(), out = [];
  sheetToObjects(SHEETS.YARN_PO).forEach(function (p) {
    var st = String(p['Status'] || '').trim();
    if (st !== 'Approved' && st !== 'Billed') return;
    var ordered = Number(p['Qty KG']) || 0;
    if (!(ordered > 0)) return;
    var k = key(p['PO No'], p['JO Number'], p['Counts'], p['Content'], p['GSM']);
    var got = received[k] || 0, pending = ordered - got;
    if (!(pending > 0.5)) return;                      // fully received
    var d = String(p['PO Date'] || '').slice(0, 10), days = '';
    if (/^\d{4}-\d{2}-\d{2}$/.test(d)) days = Math.max(0, Math.floor((today - new Date(d + 'T00:00:00')) / 86400000));
    out.push({
      poNo: p['PO No'], poDate: d, poType: p['PO Type'] || 'Yarn', supplier: p['Supplier'] || '',
      jo: p['JO Number'], counts: normalizeCounts_(p['Counts'] || ''), content: p['Content'] || '', gsm: p['GSM'] || '',
      orderedKg: Math.round(ordered * 1000) / 1000, receivedKg: Math.round(got * 1000) / 1000, pendingKg: Math.round(pending * 1000) / 1000,
      lastInward: lastIn[k] || '', days: days
    });
  });
  return out.sort(function (a, b) { return String(a.poDate).localeCompare(String(b.poDate)) || String(a.poNo).localeCompare(String(b.poNo)); });
  });
}

/*************** DASHBOARD: Orders In Hand, month-wise delivery drill-down (Month -> Buyer -> Style) ***************/
function getOihMonthlyReport() {
  var oh = getOrderInHand();
  var monthsMap = {}, monthOrder = [];
  oh.lines.forEach(function (l) {
    var qty = Number(l.qty) || 0;
    if (!(qty > 0)) return;
    var d = String(l.despatch || '').trim();
    var month = /^\d{4}-\d{2}/.test(d) ? d.slice(0, 7) : 'No delivery date';
    var buyer = l.buyer || '(no buyer)';
    var style = l.style || '(no style)';
    if (!monthsMap[month]) { monthsMap[month] = { month: month, totalQty: 0, buyers: {}, buyerOrder: [] }; monthOrder.push(month); }
    var m = monthsMap[month]; m.totalQty += qty;
    if (!m.buyers[buyer]) { m.buyers[buyer] = { buyer: buyer, totalQty: 0, styles: {}, styleOrder: [] }; m.buyerOrder.push(buyer); }
    var b = m.buyers[buyer]; b.totalQty += qty;
    if (!b.styles[style]) { b.styles[style] = { style: style, jo: l.jo, color: l.color, totalQty: 0 }; b.styleOrder.push(style); }
    b.styles[style].totalQty += qty;
  });
  monthOrder.sort(function (a, b) { return a === 'No delivery date' ? 1 : b === 'No delivery date' ? -1 : a.localeCompare(b); });
  var fmt = function (m) {
    if (m === 'No delivery date') return m;
    var d = new Date(m + '-01');
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM yyyy');
  };
  return monthOrder.map(function (mk) {
    var m = monthsMap[mk];
    return {
      month: mk, monthLabel: fmt(mk), totalQty: Math.round(m.totalQty),
      buyers: m.buyerOrder.map(function (bk) {
        var b = m.buyers[bk];
        return { buyer: b.buyer, totalQty: Math.round(b.totalQty),
          styles: b.styleOrder.map(function (sk) { var st = b.styles[sk]; return { style: st.style, jo: st.jo, color: st.color, totalQty: Math.round(st.totalQty) }; }) };
      })
    };
  });
}

/*************** SAMPLES: priority flag on JO dropdowns + samples tracker for the fabric team ***************/
/* every JO with its sample flag (reuses the same rule as isSampleJO_: JO number starting
   S-/SAMPLE, or any of its Style Tracker rows marked Order Type = SAMPLE) — computed in one
   pass over rows already in hand, rather than re-querying the sheet per JO. */
function getStyleTrackerJOsSampleFlag() {
  var rows = getStyleTracker(), seen = {}, out = [], info = {};
  rows.forEach(function (r) {
    var jo = String(r['JO Number'] || '').trim();
    if (!jo) return;
    var i = info[jo] = info[jo] || { bulk: 0, bulkSample: 0, sampleKg: 0 };
    if (isSampleWtRow_(r)) { i.sampleKg += Number(r['Fabric Required (kg)']) || 0; return; }
    i.bulk++;
    if (String(r['Order Type'] || '').trim().toUpperCase() === 'SAMPLE') i.bulkSample++;
  });
  var poKg = {};
  sheetToObjects(SHEETS.YARN_PO).forEach(function (r) {
    if (r['Status'] === 'Rejected') return;
    var j = String(r['JO Number'] || '').trim(); poKg[j] = (poKg[j] || 0) + (Number(r['Qty KG']) || 0);
  });
  rows.forEach(function (r) {
    var jo = String(r['JO Number'] || '').trim();
    if (!jo || seen[jo]) return;
    seen[jo] = true;
    var i = info[jo];
    var wholeSample = /^S\s*-/i.test(jo) || /^SAMPLE/i.test(jo) || (i.bulk > 0 && i.bulkSample === i.bulk);
    var pendingWeight = i.sampleKg > 0 && (i.sampleKg - (poKg[jo] || 0)) > 0.5;   // sample weight not yet on a Yarn PO
    out.push({ jo: jo, isSample: wholeSample || pendingWeight, hasSampleWeight: i.sampleKg > 0 });
  });
  return out;
}

/* For the fabric team: every SAMPLE JO with its fabric/yarn need and its trims (accessories) need,
   in one place — mirrors what Accessories Planning already shows for trims. */
function getSamplesTrackerReport() {
  var flags = getStyleTrackerJOsSampleFlag();
  var jos = flags.filter(function (x) { return x.isSample || x.hasSampleWeight; }).map(function (x) { return x.jo; });
  var accByJo = {};
  sheetToObjects(SHEETS.ACCESSORIES_BOM).forEach(function (r) {
    var jo = String(r['JO Number'] || '').trim();
    if (jos.indexOf(jo) === -1) return;
    (accByJo[jo] = accByJo[jo] || []).push(String(r['Trims / Accessories'] || '').trim());
  });
  var all = getStyleTracker();
  return jos.map(function (jo) {
    var rows = all.filter(function (r) { return String(r['JO Number']).trim() === jo; });
    var first = rows[0] || {};
    var wt = rows.filter(isSampleWtRow_);
    var poKg = sheetToObjects(SHEETS.YARN_PO).filter(function (r) { return String(r['JO Number']).trim() === jo && r['Status'] !== 'Rejected'; })
      .reduce(function (t, r) { return t + (Number(r['Qty KG']) || 0); }, 0);
    var counts = {}, fabricKg, parts = '';
    if (wt.length) {   // sample weight inside a normal order JO: only the sample kg, part-wise
      getCountsForJO(jo, '', false).forEach(function (c) { if (c.sampleKg > 0) counts[c.counts] = (counts[c.counts] || 0) + c.sampleKg; });
      fabricKg = wt.reduce(function (t, r) { return t + (Number(r[SAMPLE_WT_COL]) || 0); }, 0);
      parts = wt.map(function (r) { return String(r['Part'] || '') + ' ' + (Number(r[SAMPLE_WT_COL]) || 0) + ' kg'; }).join(', ');
    } else {
      getCountsForJO(jo, '', false).forEach(function (c) { counts[c.counts] = (counts[c.counts] || 0) + c.requiredKg; });
      fabricKg = rows.reduce(function (t, r) { return t + (Number(r['Fabric Required (kg)']) || 0); }, 0);
    }
    var trims = (accByJo[jo] || []).filter(function (v, i, a) { return a.indexOf(v) === i; });
    return {
      jo: jo, buyer: first['Buyer Name'] || '', style: rows.map(function (r) { return r['Style Number']; }).filter(function (v, i, a) { return v && a.indexOf(v) === i; }).join(', '),
      counts: Object.keys(counts).map(function (c) { return { counts: c, requiredKg: Math.round(counts[c] * 1000) / 1000 }; }),
      fabricKg: Math.round(fabricKg * 1000) / 1000, yarnPoKg: Math.round(poKg * 1000) / 1000, sampleParts: parts,
      trims: trims, trimCount: trims.length
    };
  });
}

/*************** ORDER IN HAND - buyer-wise PDF (same filters as the screen, with garment photos) ***************/
function generateOihPdfHtml_(f) {
  f = f || {};
  var oh = getOrderInHand(), E = function (v) { return String(v == null ? '' : v).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var q = String(f.q || '').trim().toLowerCase(), fb = String(f.buyer || '').trim(), fs = String(f.status || ''), ft = String(f.type || ''), sort = f.sort || 'despatch';
  var lines = oh.lines.filter(function (l) {
    var J = oh.jos[l.jo] || {};
    if (fb && l.buyer !== fb) return false;
    if (fs && J.status !== fs) return false;
    if (ft && l.type !== ft) return false;
    if (q && [l.jo, l.buyer, l.style, l.desc, l.color, l.fabric, l.orderStatus, l.remarks, J.ocn, J.merch].join(' ').toLowerCase().indexOf(q) < 0) return false;
    return true;
  });
  if (!lines.length) throw new Error('No Order In Hand lines match the selected filters.');

  // buyer -> JO -> lines
  var byBuyer = {}, buyerList = [];
  lines.forEach(function (l) {
    var b = l.buyer || '(no buyer)';
    if (!byBuyer[b]) { byBuyer[b] = { jos: {}, joList: [], pcs: 0, lines: 0 }; buyerList.push(b); }
    var B = byBuyer[b];
    if (!B.jos[l.jo]) { B.jos[l.jo] = []; B.joList.push(l.jo); }
    B.jos[l.jo].push(l); B.pcs += l.qty; B.lines++;
  });
  buyerList.sort(function (a, b) { return a.localeCompare(b); });
  var minDes = function (B, jo) { return B.jos[jo].map(function (l) { return l.despatch || '9999'; }).sort()[0]; };

  var imgCache = {}, tz = Session.getScriptTimeZone(), today = new Date(); today.setHours(0, 0, 0, 0);
  var allRows = getStyleTracker();
  var noPic = [];
  var thumbForJo = function (jo) {
    var link = stImgLinkForJo_(jo);
    if (!link) { if (noPic.indexOf(jo) < 0) noPic.push(jo); return ''; }
    if (!(link in imgCache)) imgCache[link] = stImgThumbDataUrl_(link);
    if (!imgCache[link] && noPic.indexOf(jo) < 0) noPic.push(jo + ' (' + (STIMG_ERR_ || 'picture could not be read') + ')');
    return imgCache[link];
  };
  var n0 = function (n) { return n ? Number(n).toLocaleString('en-IN', { maximumFractionDigits: 0 }) : '-'; };
  var dm = function (d) { return d ? String(d).slice(5).split('-').reverse().join('/') : '-'; };
  var days = function (d) { if (!d) return ''; var x = new Date(d); if (isNaN(x)) return ''; x.setHours(0, 0, 0, 0); return Math.round((x - today) / 86400000); };

  var css = '<style>@page{size:A4;margin:8mm}body{font-family:Arial,sans-serif;font-size:9.5px;color:#000;margin:0}' +
    'h1{font-size:15px;margin:0;color:#0B6E74}.sub{color:#555;font-size:9px;margin:2px 0 6px}' +
    'table{border-collapse:collapse;width:100%}th{background:#0B6E74;color:#fff;text-align:left;padding:3px 4px;font-size:8.5px}' +
    'td{padding:3px 4px;border-bottom:1px solid #ddd;vertical-align:middle}.r{text-align:right}' +
    '.bh{background:#1B2F5E;color:#fff;font-weight:bold;font-size:12px;padding:5px 8px;margin-top:6px}' +
    '.jh td{background:#E3F6F4;font-weight:bold;border-top:1px solid #0B6E74;font-size:9px}' +
    '.pg{page-break-before:always}.nop{width:52px;height:52px;border:1px dashed #bbb;color:#aaa;font-size:7px;text-align:center;line-height:52px}' +
    '.sm{color:#777;font-size:8px}.late{color:#DC2626;font-weight:bold}.soon{color:#D97706;font-weight:bold}.ok{color:#16A34A;font-weight:bold}</style>';

  var totJos = 0, totPcs = 0;
  buyerList.forEach(function (b) { totJos += byBuyer[b].joList.length; totPcs += byBuyer[b].pcs; });
  var filt = [fb ? 'Buyer: ' + fb : 'All buyers', fs ? 'JO status: ' + fs : 'All JO status', ft ? 'Type: ' + ft : 'Orders + Samples', q ? 'Search: "' + f.q + '"' : ''].filter(Boolean).join(' · ');
  var h = '<!DOCTYPE html><html><head><meta charset="utf-8">' + css + '</head><body>' +
    '<table style="border-bottom:2px solid #0B6E74;margin-bottom:6px"><tr><td style="width:56px;border:0"><img src="' + JIWIN_LOGO_PNG + '" style="height:46px"></td>' +
    '<td style="border:0"><h1>ORDER IN HAND' + (fb ? ' &mdash; ' + E(fb) : ' &mdash; Buyer-wise') + '</h1><div class="sub">JIWIN APPARELS &middot; ' + E(filt) + ' &middot; printed ' + Utilities.formatDate(new Date(), tz, 'dd-MM-yyyy HH:mm') + '</div></td>' +
    '<td class="r" style="border:0;font-size:11px"><b>' + totJos + '</b> JO(s) &middot; <b>' + lines.length + '</b> line(s)<br><b>' + totPcs.toLocaleString('en-IN') + '</b> pcs</td></tr></table>';

  if (buyerList.length > 1) {   // buyer summary on the first page
    h += '<table style="width:60%;margin-bottom:4px"><tr><th>Buyer</th><th class="r">JOs</th><th class="r">Lines</th><th class="r">Pcs</th></tr>' +
      buyerList.map(function (b) { var B = byBuyer[b]; return '<tr><td><b>' + E(b) + '</b></td><td class="r">' + B.joList.length + '</td><td class="r">' + B.lines + '</td><td class="r"><b>' + B.pcs.toLocaleString('en-IN') + '</b></td></tr>'; }).join('') +
      '<tr><td><b>TOTAL</b></td><td class="r"><b>' + totJos + '</b></td><td class="r"><b>' + lines.length + '</b></td><td class="r"><b>' + totPcs.toLocaleString('en-IN') + '</b></td></tr></table>';
  }

  buyerList.forEach(function (b, bi) {
    var B = byBuyer[b];
    B.joList.sort(function (x, y) {
      if (sort === 'jo') return x.localeCompare(y, undefined, { numeric: true });
      return minDes(B, x).localeCompare(minDes(B, y));
    });
    h += '<div class="bh' + (bi > 0 ? ' pg' : '') + '">' + E(b) + ' &nbsp;&middot;&nbsp; ' + B.joList.length + ' JO(s) &nbsp;&middot;&nbsp; ' + B.pcs.toLocaleString('en-IN') + ' pcs</div>' +
      '<table><tr><th style="width:60px">Photo</th><th>Style</th><th>Colour</th><th>Fabric</th><th class="r">Qty</th><th>Shipment</th><th class="r">Days</th><th>Key dates (Yarn / Fabric / Trims IH · PP · Cut)</th></tr>';
    B.joList.forEach(function (jo) {
      var J = oh.jos[jo] || {}, L = B.jos[jo], jq = L.reduce(function (t, l) { return t + l.qty; }, 0);
      h += '<tr class="jh"><td colspan="8">' + E(jo) + (J.ocn ? ' &middot; OCN ' + E(J.ocn) : '') + (J.merch ? ' &middot; ' + E(J.merch) : '') + ' &middot; ' + jq.toLocaleString('en-IN') + ' pcs' +
        ' &middot; <span class="sm" style="color:#0B6E74">Fabric ' + n0(J.fabricKg) + ' kg &middot; Yarn PO ' + n0(J.yarnPoKg) + ' / Rcvd ' + n0(J.yarnRcvKg) + ' &middot; Knit ' + n0(J.knitKg) + ' &middot; Dyed ' + n0(J.dyeKg) + ' &middot; Inward ' + n0(J.inwardKg) + ' &middot; ' + E(J.stage || '') + '</span></td></tr>';
      var jimg = thumbForJo(jo);
      L.forEach(function (l, li) {
        var d = days(l.despatch), t = l.tna || {};
        h += '<tr>' + (li === 0 ? '<td rowspan="' + L.length + '" style="vertical-align:top;text-align:center;border-bottom:1px solid #ccc">' + (jimg ? imgTag_(jimg, 52, 62, 'border:1px solid #ccc;') : '<div class="nop">no picture</div>') + '</td>' : '') +
          '<td><b>' + E(l.style) + '</b>' + (l.type === 'SAMPLE' ? ' <span style="color:#B45309;font-weight:bold">[SAMPLE]</span>' : '') + '<div class="sm">' + E(l.desc) + '</div></td>' +
          '<td>' + E(l.color) + '</td><td>' + E(l.fabric) + '</td><td class="r"><b>' + Number(l.qty).toLocaleString('en-IN') + '</b></td>' +
          '<td>' + E(l.despatch || '-') + '</td><td class="r ' + (d === '' ? '' : d < 0 ? 'late' : d <= 15 ? 'soon' : 'ok') + '">' + d + '</td>' +
          '<td class="sm">' + [t['Yarn IH'], t['Fabric IH'], t['Trims IH']].map(dm).join(' / ') + ' &middot; ' + dm(t['PP AD']) + ' &middot; ' + dm(t['Cutting']) + (l.remarks ? '<br>' + E(l.remarks) : '') + '</td></tr>';
      });
    });
    h += '</table>';
  });
  if (noPic.length) h += '<div style="margin-top:10px;font-size:8.5px;color:#B45309">No garment picture on file for: ' + E(noPic.join(', ')) + ' &mdash; add it in Style Tracker (JO Entry &rarr; Add / Change picture).</div>';
  return h + '</body></html>';
}


/*************** STYLE CALENDAR -> GOOGLE TASKS ***************/
function getStyleCalendarEvents(year, month) {
  year = Number(year); month = Number(month);
  if (!year || month < 1 || month > 12) throw new Error('Invalid calendar month.');
  var tz = Session.getScriptTimeZone();
  var start = new Date(year, month - 1, 1); var end = new Date(year, month, 1);
  var rows = getStyleTracker(); var events = []; var dateFields = (typeof ST_ORDER_DATE_COLS !== 'undefined' && ST_ORDER_DATE_COLS.length) ? ST_ORDER_DATE_COLS : ['Yarn Inhouse','Fabric Inhouse','Trims IH Date','PP AD','Cutting Date','Print Start','Pro Start'];
  var departmentMap = {'Yarn Inhouse':'Purchase','Fabric Inhouse':'Fabric','Trims IH Date':'Purchase','PP AD':'Merch','Cutting Date':'Production','Print Start':'Production','Pro Start':'Production'};
  rows.forEach(function(r) {
    var jo=String(r['JO Number']||'').trim(); if(!jo) return;
    dateFields.forEach(function(field) {
      var raw=r[field]; if(!raw) return;
      var d = raw instanceof Date ? raw : new Date(String(raw).slice(0,10)+'T00:00:00');
      if(isNaN(d.getTime()) || d < start || d >= end) return;
      events.push({date:Utilities.formatDate(d,tz,'yyyy-MM-dd'),jo:jo,style:String(r['Style Number']||''),buyer:String(r['Buyer Name']||''),color:String(r['Color']||''),milestone:field,assignedTo:String(r['Merchandiser']||''),department:departmentMap[field]||'Merch',remarks:String(r['Order Remarks']||'')});
    });
  });
  events.sort(function(a,b){return a.date.localeCompare(b.date)||a.jo.localeCompare(b.jo)||a.milestone.localeCompare(b.milestone);});
  return {events:events};
}

/**
 * Creates a real Google Task in the default task list of the Google account
 * authorizing/running this Apps Script. Google Tasks API does not support assigning
 * a task directly to another user's account; the assignee is included in the task notes.
 * Enable the Advanced Google service "Tasks API" (Tasks) and its Google Cloud API.
 */
function createGoogleTaskFromStyleCalendar(p) {
  if (!p || !p.date || !p.milestone) return {success:false,error:'Task date and task description are required.'};
  try {
    if (typeof Tasks === 'undefined' || !Tasks.Tasks) {
      return {success:false,error:'Google Tasks API is not enabled. In Apps Script, open Services (+), add Tasks API, then enable Google Tasks API in the linked Google Cloud project and redeploy.'};
    }
    var title = (p.style ? String(p.style)+' — ' : '') + String(p.milestone);
    if (p.jo) title += ' ('+String(p.jo)+')';
    var notes = [];
    if (p.jo) notes.push('JO: '+p.jo);
    if (p.style) notes.push('Style: '+p.style);
    if (p.buyer) notes.push('Buyer: '+p.buyer);
    if (p.color) notes.push('Colour: '+p.color);
    if (p.department) notes.push('Department: '+p.department);
    if (p.assignedTo) notes.push('Assigned to: '+p.assignedTo);
    if (p.remarks) notes.push('Remarks: '+p.remarks);
    notes.push('Created from Jiwin Fabric ERP Style Calendar.');
    var dueDate = String(p.date).slice(0,10);
    var resource = {title:title,notes:notes.join('\n'),due:dueDate+'T00:00:00.000Z'};
    var created = Tasks.Tasks.insert(resource,'@default');
    return {success:true,taskId:created&&created.id?created.id:''};
  } catch (e) {
    return {success:false,error:'Google Tasks API error: '+(e&&e.message?e.message:e)};
  }
}


/*************** FABRIC REQUIREMENT REPORTS (Reports tab) ***************/
function getFabricRequirementReports(){
  return getStyleTracker().map(function(r){
    var kg=Number(r['Fabric Required (kg)'])||0;
    if(!(kg>0)) return null;
    var part=String(r['Part']||'').trim();
    var isColorKnitting=/LOOSE\s*(KNIT|RIB)/i.test(part);
    return {
      buyer:String(r['Buyer Name']||'').trim()||'(no buyer)',
      style:String(r['Style Number']||'').trim()||'(no style)',
      content:String(r['Content']||'').trim(),
      counts:normalizeCounts_(r['Counts']),
      part:part,
      color:String(r['Color']||'').trim(),
      kg:kg,
      colorKnittingKg:isColorKnitting?kg:0,
      colorKnitting:isColorKnitting?'Color Knitting — After Dyeing':''
    };
  }).filter(function(r){return r&&r.counts;});
}
