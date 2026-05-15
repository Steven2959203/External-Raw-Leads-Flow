/**
 * LEAD SYNC ENGINE (Unified & Optimized)
 * 1. Auto-adds checkbox on form submit
 * 2. Deduplicates with "Clean Phone" algorithm
 * 3. Handles SA Timezone and 3-day follow-up
 */

const CONFIG = {
  DESTINATION_ID: '', 
  DEST_TAB_NAME: 'Leads',
  SOURCE_SHEETS: ['QB Form Test', ], 
  SYNC_TRIGGER: 'Sync to Leads',
  TIMEZONE: "Australia/Adelaide",
  DUP_FIELDS: ['Phone Number', 'Email Address'],
  MAP: {
    "Rep": "Rep",
    "Contact Name": "Contact Name",
    "Business Name": "Business Name",
    "Phone Number": "Phone Number",
    "Email Address": "Email Address",
    "Suburb": "Suburb",
    "Equipment": "Equipment",
    "Source": "Source"
  },
  DEFAULTS: { "Status": "4 - New Lead" }
};

/**
 * Helper to get AU State from Postcode
 */
function getAUState(postcode) {
  const pc = parseInt(postcode, 10);
  if (isNaN(pc) || postcode.toString().length !== 4) return null;

  if (pc >= 2000 && pc <= 2999) return "NSW";
  if (pc >= 3000 && pc <= 3999) return "VIC";
  if (pc >= 4000 && pc <= 4999) return "QLD";
  if (pc >= 5000 && pc <= 5999) return "SA";
  if (pc >= 6000 && pc <= 6999) return "WA";
  if (pc >= 7000 && pc <= 7999) return "TAS";
  if (pc >= 800 && pc <= 999) return "NT";
  if (pc >= 200 && pc <= 299) return "ACT"; // Handling edge cases for ACT
  return null;
}

/**
 * AUTO-ADD CHECKBOX: Set this to "On form submit" trigger
 */
function onFormSubmit(e) {
  const sheet = e.range.getSheet();
  if (!CONFIG.SOURCE_SHEETS.includes(sheet.getName())) return;
  const col = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf(CONFIG.SYNC_TRIGGER) + 1;
  if (col > 0) sheet.getRange(e.range.getRow(), col).insertCheckboxes().setBackground("#f3f3f3");
}

/**
 * SYNC HANDLER: Set this to "On edit" trigger
 */
function installableOnEdit(e) {
  if (!e || e.value !== "TRUE") return;
  const sheet = e.range.getSheet();
  if (!CONFIG.SOURCE_SHEETS.includes(sheet.getName())) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (e.range.getColumn() !== (headers.indexOf(CONFIG.SYNC_TRIGGER) + 1)) return;

  e.range.setBackground("#fff2cc").setValue("Syncing...");
  SpreadsheetApp.flush();
  processSync(e.range.getRow(), sheet, headers, e.range);
}

function processSync(row, srcSheet, headers, trigger) {
  const ui = SpreadsheetApp.getUi();
  const srcData = {};
  srcSheet.getRange(row, 1, 1, srcSheet.getLastColumn()).getValues()[0].forEach((v, i) => srcData[headers[i]] = v);

  const clean = (p) => p ? p.toString().split(/[\/,]/).map(x => x.replace(/\D/g, '')).filter(x => x.length > 7) : [];
  const now = (d, fmt) => Utilities.formatDate(d, CONFIG.TIMEZONE, fmt);

  try {
    const dSheet = SpreadsheetApp.openById(CONFIG.DESTINATION_ID).getSheetByName(CONFIG.DEST_TAB_NAME);
    const dHeaders = dSheet.getRange(1, 1, 1, dSheet.getLastColumn()).getValues()[0];
    const dData = dSheet.getDataRange().getValues();

    // Check for Duplicates
    for (let f of CONFIG.DUP_FIELDS) {
      const dIdx = dHeaders.indexOf(f);
      const sKey = Object.keys(CONFIG.MAP).find(k => CONFIG.MAP[k] === f);
      const sVal = srcData[sKey];
      if (dIdx === -1 || !sVal) continue;

      let match = (f === "Phone Number") 
        ? dData.some((r, i) => i > 0 && clean(sVal).some(p => clean(r[dIdx]).includes(p)))
        : dData.some((r, i) => i > 0 && r[dIdx].toString().toLowerCase().trim() === sVal.toString().toLowerCase().trim());

      if (match && ui.alert('Duplicate', `${f} match found for "${sVal}". Sync anyway?`, ui.ButtonSet.YES_NO) !== ui.Button.YES) {
        trigger.setValue(false).setBackground(null);
        return;
      }
    }

    // Determine AU State/Country logic
    const postcodeVal = srcData["Suburb"]; // Assuming the 4-digit input comes through 'Suburb'
    const detectedState = getAUState(postcodeVal);

    const newRow = dHeaders.map(h => {
      // Special Fields
      if (h === "Date") return now(new Date(), "dd/MM/yyyy HH:mm");
      if (h === "Follow Up") return now(new Date(Date.now() + 259200000), "dd/MM/yyyy");
      
      // Auto-logic for State/Country
      if (h === "State" && detectedState) return detectedState;
      if (h === "Country" && detectedState) return "Australia";

      // Standard Mapping
      const sK = Object.keys(CONFIG.MAP).find(k => CONFIG.MAP[k] === h);
      let v = (sK && srcData[sK]) ? srcData[sK] : (CONFIG.DEFAULTS[h] || "");
      
      // Force Phone to String for Sheets
      return (h === "Phone Number" && v !== "") ? "'" + v : v;
    });

    dSheet.appendRow(newRow);
    trigger.setValue("SYNCED").setBackground("#d9ead3");
  } catch (err) {
    ui.alert("Error: " + err.message);
    trigger.setValue(false).setBackground("#f4cccc");
  }
}
