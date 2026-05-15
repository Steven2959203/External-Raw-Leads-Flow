/**
 * GOOGLE APPS SCRIPT: LEAD SYNC ENGINE (Optimized)
 */

const CONFIG = {
  DESTINATION_ID: '', 
  DEST_TAB_NAME: 'Leads',
  SOURCE_SHEETS: ['QB Form Test'], 
  SYNC_TRIGGER_COLUMN: 'Sync to Leads',
  TIMEZONE: "Australia/Adelaide",
  DUP_CHECK_FIELDS: ['Phone Number', 'Email Address'],

  FIELD_MAP: {
    "Rep": "Rep",
    "Contact Name": "Contact Name",
    "Business Name": "Business Name",
    "Phone Number": "Phone Number",
    "Email Address": "Email Address",
    "Suburb": "Suburb",
    "Equipment": "Equipment",
    "Source": "Source"
  },

  DEFAULT_VALUES: {
    "Status": "4 - Early"
  }
};

/**
 * Helper to strip non-numeric characters for comparison
 * Also handles strings with "/" by returning an array of numbers
 */
function cleanPhone(phoneStr) {
  if (!phoneStr) return [];
  // Split by common separators like slash or comma
  const parts = phoneStr.toString().split(/[\/,]/);
  // Remove all non-digits from each part
  return parts.map(p => p.replace(/\D/g, '')).filter(p => p.length > 0);
}

function installableOnEdit(e) {
  if (!e) return;
  const range = e.range;
  const sheet = range.getSheet();
  if (!CONFIG.SOURCE_SHEETS.includes(sheet.getName())) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const syncColIndex = headers.indexOf(CONFIG.SYNC_TRIGGER_COLUMN) + 1;

  if (range.getColumn() === syncColIndex && e.value === "TRUE") {
    // 4. VISUAL FEEDBACK: Start Syncing
    range.setBackground("#fff2cc"); // Light Yellow
    range.setValue("Syncing...");
    SpreadsheetApp.flush(); // Force UI update immediately

    processSync(range.getRow(), sheet, headers, range);
  }
}

function processSync(rowNum, sourceSheet, sourceHeaders, triggerRange) {
  const ui = SpreadsheetApp.getUi();
  const rowDataRaw = sourceSheet.getRange(rowNum, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  const sourceData = {};
  sourceHeaders.forEach((h, i) => sourceData[h] = rowDataRaw[i]);

  const todayStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "dd/MM/yyyy HH:mm");
  const followUpStr = Utilities.formatDate(new Date(Date.now() + 3*24*60*60*1000), CONFIG.TIMEZONE, "dd/MM/yyyy");

  try {
    const destSs = SpreadsheetApp.openById(CONFIG.DESTINATION_ID);
    const destSheet = destSs.getSheetByName(CONFIG.DEST_TAB_NAME);
    const destHeaders = destSheet.getRange(1, 1, 1, destSheet.getLastColumn()).getValues()[0];
    const destData = destSheet.getDataRange().getValues();

    // --- 1 & 3: ADVANCED DEDUPLICATION ---
    for (let fieldName of CONFIG.DUP_CHECK_FIELDS) {
      const destColIdx = destHeaders.indexOf(fieldName);
      const sourceKey = Object.keys(CONFIG.FIELD_MAP).find(key => CONFIG.FIELD_MAP[key] === fieldName);
      let valToSearch = sourceData[sourceKey];
      
      if (destColIdx === -1 || !valToSearch) continue;

      let hasMatch = false;
      let matchInfo = "";

      if (fieldName === "Phone Number") {
        const sourcePhones = cleanPhone(valToSearch);
        // Look through every row in destination
        for (let i = 1; i < destData.length; i++) {
          const existingPhones = cleanPhone(destData[i][destColIdx]);
          // Check if any number from source exists in the destination row
          const overlap = sourcePhones.filter(p => existingPhones.includes(p));
          if (overlap.length > 0) {
            hasMatch = true;
            matchInfo = overlap.join(", ");
            break;
          }
        }
      } else {
        // Standard check for Email/others
        const searchStr = valToSearch.toString().toLowerCase().trim();
        hasMatch = destData.some(row => row[destColIdx] && row[destColIdx].toString().toLowerCase().trim() === searchStr);
        matchInfo = valToSearch;
      }

      if (hasMatch) {
        const response = ui.alert('Duplicate Warning', `Match found for ${fieldName}: "${matchInfo}". \n\nSync anyway?`, ui.ButtonSet.YES_NO);
        if (response !== ui.Button.YES) {
          triggerRange.setValue(false).setBackground(null);
          return;
        }
      }
    }

    // --- 2: MAPPING & FORMATTING ---
    const newRow = destHeaders.map(destHeader => {
      if (destHeader === "Date") return todayStr;
      if (destHeader === "Follow Up") return followUpStr;

      const sourceKey = Object.keys(CONFIG.FIELD_MAP).find(key => CONFIG.FIELD_MAP[key] === destHeader);
      let value = "";
      
      if (sourceKey && sourceData[sourceKey] !== undefined && sourceData[sourceKey] !== "") {
        value = sourceData[sourceKey];
      } else if (CONFIG.DEFAULT_VALUES[destHeader]) {
        value = CONFIG.DEFAULT_VALUES[destHeader];
      }

      // 2. FIX LEADING ZEROS: If it looks like a phone number, force it to text
      if (destHeader === "Phone Number" && value !== "") {
        return "'" + value.toString(); 
      }

      return value;
    });

    destSheet.appendRow(newRow);
    
    // 4. VISUAL FEEDBACK: Success
    triggerRange.setValue("SYNCED").setBackground("#d9ead3"); // Light Green
    SpreadsheetApp.getActiveSpreadsheet().toast("Lead synced successfully", "Success");

  } catch (err) {
    console.error(err);
    ui.alert("Error during sync: " + err.message);
    triggerRange.setValue(false).setBackground("#f4cccc"); // Light Red
  }
}

// #### V2 ####

