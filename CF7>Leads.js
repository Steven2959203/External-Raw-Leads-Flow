/**
 * CF7 → LEADS SYNC  (fully sheet-driven config)
 *
 * TRIGGERS:
 *   onNewCF7Row       → Time-driven, every minute   (auto-adds checkbox)
 *   installableOnEdit → On edit                     (fires on checkbox tick)
 *
 * CONFIG TAB LAYOUT  ("CF7 Config" tab):
 *
 *   A              B                 D                 E               F
 *   ─────────────────────           ──────────────────────────────────────────
 *   Setting        Value             CF7 Raw Field     Leads Column    Default Value
 *   (instruction row)                (instruction row)
 *   Source Tab     CF7 Raw           fname             Contact Name
 *   Dest Sheet ID  abc123...         phone             Phone Number
 *   Dest Tab       Leads             ...               ...
 *   Timezone       Australia/...
 *   Follow Up Days 3
 *   Dup Fields     Phone Number, Email Address
 */

// ─────────────────────────────────────────────
//  ONLY THING LEFT IN SCRIPT:
//  The config tab name — everything else lives in the sheet
// ─────────────────────────────────────────────
const CONFIG_TAB_NAME = '(Config)Website Forms';

// ─────────────────────────────────────────────
//  TRANSFORM HELPERS  (script-only, client never touches)
//  Preserve leading zero on phone numbers
// ─────────────────────────────────────────────
function forcePhoneString(val) {
  if (!val || val === '') return val;
  const s = val.toString().trim();
  return s.startsWith("'") ? s : "'" + s;
}

// Dest column header → auto-applied transform function
const AUTO_TRANSFORMS = {
  'Phone Number': forcePhoneString,
};

// ─────────────────────────────────────────────
//  CONFIG TAB READER
// ─────────────────────────────────────────────

/**
 * Reads the config tab and returns:
 * {
 *   settings: { sourceTab, destSheetId, destTab, timezone, followUpDays, dupFields[] }
 *   fieldConfig: [{ cf7Key, destCol, defaultVal }, ...]
 * }
 *
 * Settings block  → cols A:B, starting row 3 (row 1 = header, row 2 = instruction)
 * Field map block → cols D:F, starting row 3 (row 1 = header, row 2 = instruction)
 */
function loadConfig() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG_TAB_NAME);
  if (!configSheet) throw new Error(`Config tab "${CONFIG_TAB_NAME}" not found.`);

  const data = configSheet.getDataRange().getValues();

  // ── Settings (cols A:B, skip rows 0–1) ──
  const settingsRaw = {};
  for (let i = 2; i < data.length; i++) {
    const key = data[i][0] ? data[i][0].toString().trim() : '';
    const val = data[i][1] !== undefined ? data[i][1] : '';
    if (key) settingsRaw[key] = val.toString().trim();
  }

  const settings = {
    sourceTab:    settingsRaw['Source Tab']     || '',
    destSheetId:  settingsRaw['Dest Sheet ID']  || '',
    destTab:      settingsRaw['Dest Tab']        || '',
    timezone:     settingsRaw['Timezone']        || 'Australia/Adelaide',
    followUpDays: parseInt(settingsRaw['Follow Up Days']) || 3,
    dupFields:    settingsRaw['Dup Fields']
                    ? settingsRaw['Dup Fields'].split(',').map(s => s.trim())
                    : [],
    syncTriggerCol: settingsRaw['Sync Trigger Col'] || 'Sync to Leads',
  };

  // ── Field map (cols D:F, skip rows 0–1) ──
  const fieldConfig = [];
  for (let i = 2; i < data.length; i++) {
    const cf7Key     = data[i][3] ? data[i][3].toString().trim() : '';
    const destCol    = data[i][4] ? data[i][4].toString().trim() : '';
    const defaultVal = data[i][5] !== undefined ? data[i][5] : '';
    if (!destCol) continue; // skip blank/incomplete rows
    fieldConfig.push({ cf7Key, destCol, defaultVal });
  }

  return { settings, fieldConfig };
}

// ─────────────────────────────────────────────
//  TRIGGER: AUTO-CHECKBOX
//  Time-driven → every minute
// ─────────────────────────────────────────────
function onNewCF7Row() {
  const { settings } = loadConfig();

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(settings.sourceTab);
  if (!sheet || sheet.getLastRow() < 2) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let triggerCol = headers.indexOf(settings.syncTriggerCol) + 1;

  // Auto-create the column if missing
  if (triggerCol === 0) {
    triggerCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, triggerCol).setValue(settings.syncTriggerCol);
  }

  const lastRow = sheet.getLastRow();
  const colVals = sheet.getRange(2, triggerCol, lastRow - 1, 1).getValues();

  colVals.forEach((r, i) => {
    if (r[0] === '' || r[0] === null) {
      sheet.getRange(i + 2, triggerCol).insertCheckboxes().setBackground('#f3f3f3');
    }
  });
}

// ─────────────────────────────────────────────
//  TRIGGER: SYNC ON CHECKBOX TICK
//  On edit
// ─────────────────────────────────────────────
function installableOnEdit(e) {
  if (!e || e.value !== 'TRUE') return;

  let settings;
  try {
    ({ settings } = loadConfig());
  } catch (err) {
    SpreadsheetApp.getUi().alert('Config Error', err.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const sheet = e.range.getSheet();
  if (sheet.getName() !== settings.sourceTab) return;

  const headers    = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const triggerCol = headers.indexOf(settings.syncTriggerCol) + 1;
  if (e.range.getColumn() !== triggerCol) return;

  e.range.setBackground('#fff2cc').setValue('Syncing...');
  SpreadsheetApp.flush();

  syncRow(e.range.getRow(), sheet, headers, e.range, settings);
}

// ─────────────────────────────────────────────
//  CORE SYNC
// ─────────────────────────────────────────────
function syncRow(row, srcSheet, srcHeaders, triggerCell, settings) {
  const ui = SpreadsheetApp.getUi();

  const srcValues = srcSheet.getRange(row, 1, 1, srcSheet.getLastColumn()).getValues()[0];
  const srcData   = {};
  srcHeaders.forEach((h, i) => { if (h) srcData[h] = srcValues[i]; });

  let fieldConfig;
  try {
    ({ fieldConfig } = loadConfig());
  } catch (err) {
    ui.alert('Config Error', err.message, ui.ButtonSet.OK);
    triggerCell.setValue(false).setBackground('#f4cccc');
    return;
  }

  const fmtDate  = (d, fmt) => Utilities.formatDate(d, settings.timezone, fmt);
  const now      = new Date();
  const dateStr  = fmtDate(now, 'dd/MM/yyyy HH:mm');
  const followUp = fmtDate(new Date(now.getTime() + settings.followUpDays * 86400000), 'dd/MM/yyyy');

  try {
    const destSS    = SpreadsheetApp.openById(settings.destSheetId);
    const destSheet = destSS.getSheetByName(settings.destTab);
    if (!destSheet) throw new Error(`Dest tab "${settings.destTab}" not found.`);

    const destHeaders = destSheet.getRange(1, 1, 1, destSheet.getLastColumn()).getValues()[0];
    const destData    = destSheet.getDataRange().getValues();

    // ── Duplicate check ──
    for (const dupField of settings.dupFields) {
      const destColIdx = destHeaders.indexOf(dupField);
      const mapEntry   = fieldConfig.find(f => f.destCol === dupField);
      if (destColIdx === -1 || !mapEntry || !mapEntry.cf7Key) continue;

      const incomingVal = srcData[mapEntry.cf7Key];
      if (!incomingVal) continue;

      let isDup = false;
      if (dupField === 'Phone Number') {
        const clean = p => p ? p.toString().split(/[\/,]/).map(x => x.replace(/\D/g, '')).filter(x => x.length > 7) : [];
        isDup = destData.some((r, i) => i > 0 && clean(incomingVal).some(p => clean(r[destColIdx]).includes(p)));
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
      if (destHeader === 'Date')      return dateStr;
      if (destHeader === 'Follow Up') return followUp;

      const entry = fieldConfig.find(f => f.destCol === destHeader);
      if (!entry) return '';

      let val = (entry.cf7Key && srcData[entry.cf7Key] !== undefined && srcData[entry.cf7Key] !== '')
        ? srcData[entry.cf7Key]
        : entry.defaultVal;

      if (AUTO_TRANSFORMS[destHeader]) val = AUTO_TRANSFORMS[destHeader](val);

      return val;
    });

    destSheet.appendRow(newRow);
    triggerCell.setValue('SYNCED').setBackground('#d9ead3');

  } catch (err) {
    ui.alert('Sync Error', err.message, ui.ButtonSet.OK);
    triggerCell.setValue(false).setBackground('#f4cccc');
  }
}
