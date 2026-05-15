/**
 * CF7 → LEADS SYNC
 * Source: WordPress CF7 + Google Sheets connector (one shared raw tab)
 * Dest:   Active Leads sheet
 *
 * TRIGGERS TO SET UP:
 *   onNewCF7Row   → On change  (auto-adds checkbox when CF7 writes a new row)
 *   installableOnEdit → On edit   (fires when you tick the checkbox)
 */

// ─────────────────────────────────────────────
//  SECTION 1 — GLOBAL CONFIG
//  Things you rarely change
// ─────────────────────────────────────────────
const CF7_CONFIG = {
  SOURCE_SHEET_ID:  SpreadsheetApp.getActiveSpreadsheet().getId(), // raw sheet (this file)
  SOURCE_TAB_NAME:  'Website Forms',          // tab the WP connector writes into
  SYNC_TRIGGER_COL: 'Sync to Leads',    // header of the checkbox column (auto-created)

  DEST_SHEET_ID:    '',                 // ← paste destination spreadsheet ID
  DEST_TAB_NAME:    'Leads',

  TIMEZONE:         'Australia/Adelaide',
  FOLLOW_UP_DAYS:   3,

  // Fields checked for duplicates (matched against dest sheet)
  DUP_FIELDS: ['Phone Number', 'Email Address'],
};

// ─────────────────────────────────────────────
//  SECTION 2 — FIELD CONFIG
//  Edit this when forms change.
//
//  Each entry:
//    cf7Key   : exact column header the WP connector writes (CF7 field name)
//    destCol  : matching column header in the Leads sheet
//    transform: optional fn to clean/format the value (or null)
//
//  ORDER doesn't matter — matching is by header name.
//  To ignore a CF7 field, simply omit it.
//  To hard-code a dest value regardless of input, use the DEFAULTS section below.
// ─────────────────────────────────────────────
const FIELD_MAP = [
  // CF7 raw column        → Leads column          transform
  { cf7Key: 'fname',        destCol: 'Contact Name',  transform: null },
  { cf7Key: 'phone',        destCol: 'Phone Number',  transform: forcePhoneString },
  { cf7Key: 'email',        destCol: 'Email Address', transform: null },
  { cf7Key: 'dropdown',     destCol: 'Equipment',     transform: null },  // "I'm interested in"
  // { cf7Key: 'product',      destCol: 'Equipment',     transform: null },  // "I'm interested in"
  { cf7Key: 'country',      destCol: 'Country',       transform: null },
  // { cf7Key: 'message',      destCol: 'Notes',         transform: null },
  // { cf7Key: 'checkbox',     destCol: 'Pref. Contact', transform: null },  // "Preferred Contact Method"
  // { cf7Key: 'page-title',   destCol: 'Page Title',    transform: null },  // page the form was submitted from
  // { cf7Key: 'form-name',    destCol: 'Source Form',   transform: null },  // CF7 form title (verify header, may differ)
  // ── Add/remove rows here as forms evolve ──
];

// Hard-coded values written to dest regardless of source data
const FIELD_DEFAULTS = {
  'Status': '4 - New Lead',
  // 'Rep': 'Unassigned',  // ← uncomment & set if needed
};

// ─────────────────────────────────────────────
//  SECTION 3 — TRANSFORM HELPERS
//  Add new ones here and reference in FIELD_MAP
// ─────────────────────────────────────────────

/** Prefix phone with apostrophe so Sheets treats it as text */
function forcePhoneString(val) {
  return val && val !== '' ? "'" + val : val;
}

/** Example: trim + title-case a name */
function titleCase(val) {
  if (!val) return val;
  return val.toString().trim().replace(/\w\S*/g, w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase());
}

// ─────────────────────────────────────────────
//  SECTION 4 — TRIGGER FUNCTIONS
//  Point your installable triggers at these
// ─────────────────────────────────────────────

/**
 * AUTO-CHECKBOX
 * Trigger: Time-driven → every minute → onNewCF7Row
 *
 * Scans all data rows for missing checkboxes and fills them in.
 * More reliable than INSERT_ROW change event for third-party connectors
 * (CF7, Zapier, etc.) which don't always fire that trigger.
 */
function onNewCF7Row() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CF7_CONFIG.SOURCE_TAB_NAME);
  if (!sheet || sheet.getLastRow() < 2) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let triggerCol = headers.indexOf(CF7_CONFIG.SYNC_TRIGGER_COL) + 1;

  // Auto-create the header column if missing
  if (triggerCol === 0) {
    triggerCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, triggerCol).setValue(CF7_CONFIG.SYNC_TRIGGER_COL);
  }

  // Scan all data rows — add checkbox only if the cell is empty
  const lastRow   = sheet.getLastRow();
  const colValues = sheet.getRange(2, triggerCol, lastRow - 1, 1).getValues();

  const toFill = [];
  colValues.forEach((r, i) => {
    if (r[0] === '' || r[0] === null) toFill.push(i + 2); // +2: 1-indexed + skip header
  });

  toFill.forEach(r => {
    sheet.getRange(r, triggerCol).insertCheckboxes().setBackground('#f3f3f3');
  });
}

/**
 * SYNC ON CHECKBOX TICK
 * Trigger: On edit → installableOnEdit
 */
function installableOnEdit(e) {
  if (!e || e.value !== 'TRUE') return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== CF7_CONFIG.SOURCE_TAB_NAME) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const triggerCol = headers.indexOf(CF7_CONFIG.SYNC_TRIGGER_COL) + 1;
  if (e.range.getColumn() !== triggerCol) return;

  e.range.setBackground('#fff2cc').setValue('Syncing...');
  SpreadsheetApp.flush();

  syncRow(e.range.getRow(), sheet, headers, e.range);
}

// ─────────────────────────────────────────────
//  SECTION 5 — CORE SYNC LOGIC
//  No need to edit unless the Leads sheet structure changes
// ─────────────────────────────────────────────

function syncRow(row, srcSheet, srcHeaders, triggerCell) {
  const ui = SpreadsheetApp.getUi();

  // Build key→value map from source row
  const srcValues = srcSheet.getRange(row, 1, 1, srcSheet.getLastColumn()).getValues()[0];
  const srcData = {};
  srcHeaders.forEach((h, i) => { if (h) srcData[h] = srcValues[i]; });

  // Helper: look up a CF7 value by cf7Key
  const getCF7Val = (cf7Key) => {
    const entry = FIELD_MAP.find(f => f.cf7Key === cf7Key);
    if (!entry) return '';
    let val = srcData[cf7Key] !== undefined ? srcData[cf7Key] : '';
    return entry.transform ? entry.transform(val) : val;
  };

  // Formatted timestamps
  const now      = new Date();
  const fmtDate  = (d, fmt) => Utilities.formatDate(d, CF7_CONFIG.TIMEZONE, fmt);
  const dateStr  = fmtDate(now, 'dd/MM/yyyy HH:mm');
  const followUp = fmtDate(new Date(now.getTime() + CF7_CONFIG.FOLLOW_UP_DAYS * 86400000), 'dd/MM/yyyy');

  try {
    const destSS    = SpreadsheetApp.openById(CF7_CONFIG.DEST_SHEET_ID);
    const destSheet = destSS.getSheetByName(CF7_CONFIG.DEST_TAB_NAME);
    if (!destSheet) throw new Error(`Dest tab "${CF7_CONFIG.DEST_TAB_NAME}" not found.`);

    const destHeaders = destSheet.getRange(1, 1, 1, destSheet.getLastColumn()).getValues()[0];
    const destData    = destSheet.getDataRange().getValues();

    // ── Duplicate check ──
    for (const dupField of CF7_CONFIG.DUP_FIELDS) {
      const destColIdx = destHeaders.indexOf(dupField);
      const mapEntry   = FIELD_MAP.find(f => f.destCol === dupField);
      if (destColIdx === -1 || !mapEntry) continue;

      const incomingVal = srcData[mapEntry.cf7Key];
      if (!incomingVal) continue;

      let isDup = false;

      if (dupField === 'Phone Number') {
        const cleanPhone = (p) => p ? p.toString().split(/[\/,]/).map(x => x.replace(/\D/g, '')).filter(x => x.length > 7) : [];
        const incomingNums = cleanPhone(incomingVal);
        isDup = destData.some((r, i) => i > 0 && incomingNums.some(p => cleanPhone(r[destColIdx]).includes(p)));
      } else {
        isDup = destData.some((r, i) => i > 0 &&
          r[destColIdx].toString().toLowerCase().trim() === incomingVal.toString().toLowerCase().trim());
      }

      if (isDup) {
        const proceed = ui.alert(
          'Duplicate Found',
          `${dupField} match for "${incomingVal}" already exists in Leads.\n\nSync anyway?`,
          ui.ButtonSet.YES_NO
        );
        if (proceed !== ui.Button.YES) {
          triggerCell.setValue(false).setBackground(null);
          return;
        }
      }
    }

    // ── Build destination row ──
    const newRow = destHeaders.map(destHeader => {
      // Auto fields
      if (destHeader === 'Date')      return dateStr;
      if (destHeader === 'Follow Up') return followUp;

      // Check FIELD_MAP
      const mapEntry = FIELD_MAP.find(f => f.destCol === destHeader);
      if (mapEntry) {
        const raw = srcData[mapEntry.cf7Key];
        return mapEntry.transform ? mapEntry.transform(raw) : (raw !== undefined ? raw : '');
      }

      // Check DEFAULTS
      if (FIELD_DEFAULTS[destHeader] !== undefined) return FIELD_DEFAULTS[destHeader];

      return '';
    });

    destSheet.appendRow(newRow);
    triggerCell.setValue('SYNCED').setBackground('#d9ead3');

  } catch (err) {
    ui.alert('Sync Error', err.message, ui.ButtonSet.OK);
    triggerCell.setValue(false).setBackground('#f4cccc');
  }
}
