/**
 * ==========================================
 * 1. CONFIGURATION SECTION (HARDCODED MAPPING)
 * ==========================================
 */
const CONFIG = {
  // --- GOOGLE SHEET 1 ---
  SHEET_1: {
    ID: "1qoi8Lhg2gpmQTNn058hCLnrK3QpI3YKkIzRPMD8gyO0", 
    TABS: [
      {
        NAME: "DATABASE",
        DATA_START_ROW: 7, 
        MAP: {
          "Employee ID": "C",
          "Full Name": "D",
          "Last Name": "E",
          "First Name": "F",
          "Middle Name": "H",
          "Suffix": "G",
          "Nickname": "P",
          "Gender": "Q",
          "Date of Birth": "S",
          "Civil Status": "U",
          "Active": "I",
          "Employment Type": "BG",
          "Company Name": "DL",
          "Division": "AT",
          "Group": "DM", 
          "Department": "AQ",
          "Section": "AU",
          "Work Location": "AP",
          "Position": "AV",
          "Job Group": "BA",
          "Superior ID": "BE",
          "Superior Name": "BF",
          "Date Hired": "BH",
          "Date Regular": "BI",
          "Date Resigned": "BL",
          "Reason for Leaving": "BS"
        }
      }
    ]
  },

  // --- GOOGLE SHEET 2 ---
  SHEET_2: {
    ID: "1y5Ao8kTzVxXc3WCSdthCbrvVmOHGkAZAAdQNjLRNLj8", 
    TABS: [
      {
        NAME: "as of June 2025",
        DATA_START_ROW: 3, 
        MAP: {
          "Employee ID": "C",
          "Full Name": "D",
          "Last Name": "E",
          "First Name": "F",
          "Middle Name": "H",
          "Suffix": "G",
          "Nickname": "J",
          "Gender": "K",
          "Date of Birth": "L",
          "Civil Status": "M",
          "Active": "I",
          "Employment Type": "X",
          "Company Name": "O",
          "Division": "P",
          "Group": "Q",
          "Department": "R",
          "Section": "S",
          "Work Location": "N",
          "Position": "T",
          "Job Group": "U",
          "Superior ID": "V",
          "Superior Name": "W",
          "Date Hired": "Y",
          "Date Regular": "Z",
          "Date Resigned": "AA",
          "Reason for Leaving": "AB"
        }
      }
    ]
  },

  // --- GOOGLE SHEET 3 ---
  SHEET_3: {
    ID: "1TubY9eVVFyFOy-B_KFKlclmgOZPdASS3Ua4nYtdeJeg", 
    TABS: [
      {
        NAME: "Personal Information",
        DATA_START_ROW: 2, 
        MAP: {
          "Employee ID": "B", 
          // Name Parts (Concatenation Sources)
          "Last Name": "D",
          "First Name": "E",
          "Middle Name": "G",
          "Suffix": "F",
          // Demographics
          "Nickname": "I",
          "Gender": "J",
          "Date of Birth": "K",
          "Civil Status": "Q"
        }
      },
      {
        NAME: "HR FIELDS",
        DATA_START_ROW: 2, 
        MAP: {
          "Employee ID": "B",
          "Active": "P",
          "Employment Type": "O",
          "Company Name": "C",
          "Division": "D",
          "Group": "E",
          "Department": "F",
          "Section": "G",
          "Work Location": "H",
          "Position": "AI",
          "Superior ID": "U",
          "Superior Name": "V",
          "Date Hired": "W",
          "Date Resigned": "Q",
          "Reason for Leaving": "S"
        }
      }
    ]
  },

  MASTER: {
    ID: "1Ll9L8D7rkze9vQVgFxs48besFwBiStz0eAIEUPrpTAY", 
    TAB_NAME: "MasterDatabase"
  }
};

/**
 * ==========================================
 * 2. WEB APP SERVING
 * ==========================================
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Employee Data Consolidator')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * ==========================================
 * 3. MASTER HEADERS (Standard List)
 * ==========================================
 */
function getMasterHeaders() {
  return [
    "Employee ID", "Full Name", "Last Name", "First Name", "Middle Name", "Suffix", "Nickname",
    "Gender", "Date of Birth", "Civil Status", "Active", "Employment Type",
    "Company Name", "Division", "Group", "Department", "Section", "Work Location",
    "Position", "Job Group", "Superior ID", "Superior Name",
    "Date Hired", "Date Regular", "Date Resigned", "Reason for Leaving"
  ];
}

/**
 * ==========================================
 * 4. FETCH & CONSOLIDATE DATA (HARDCODED)
 * ==========================================
 */
function getConsolidatedData() {
  console.log("Starting Hardcoded Data Extraction...");
  const fields = getMasterHeaders();
  
  const sources = [
    { key: 'SHEET_1', config: CONFIG.SHEET_1 },
    { key: 'SHEET_2', config: CONFIG.SHEET_2 },
    { key: 'SHEET_3', config: CONFIG.SHEET_3 }
  ];

  let groupedData = {};
  let detectedSourceNames = [];

  sources.forEach(sourceObj => {
    const sheetConfig = sourceObj.config;
    
    if (!sheetConfig.ID || sheetConfig.ID.includes("REPLACE")) {
      console.warn(`Skipping ${sourceObj.key}: ID is missing.`);
      return;
    }

    try {
      const ss = SpreadsheetApp.openById(cleanId(sheetConfig.ID));
      const fileName = ss.getName();
      if (!detectedSourceNames.includes(fileName)) detectedSourceNames.push(fileName);

      sheetConfig.TABS.forEach(tabConfig => {
        const sheet = ss.getSheetByName(tabConfig.NAME);
        if (!sheet) {
          console.warn(`Tab "${tabConfig.NAME}" not found in ${fileName}`);
          return;
        }

        const allValues = sheet.getDataRange().getValues();
        const startRow = tabConfig.DATA_START_ROW - 1; 

        if (allValues.length <= startRow) return; 

        console.log(`Processing ${fileName} - ${tabConfig.NAME}`);

        for (let i = startRow; i < allValues.length; i++) {
          const row = allValues[i];
          if (row.every(c => c === "")) continue; 

          const getVal = (fieldName) => {
            const colLetter = tabConfig.MAP[fieldName];
            if (!colLetter) return ""; 
            const colIndex = letterToColumn(colLetter) - 1; 
            return (colIndex >= 0 && colIndex < row.length) ? row[colIndex] : "";
          };

          const rawID = getVal("Employee ID");
          const cleanIdVal = normalizeEmployeeID(rawID);
          if (!cleanIdVal) continue;

          // --- NAME CONCATENATION LOGIC ---
          let rawName = getVal("Full Name");
          if (!rawName) {
            const last = getVal("Last Name");
            const first = getVal("First Name");
            const mid = getVal("Middle Name");
            const suffix = getVal("Suffix");
            
            if (last || first) {
               rawName = `${last}, ${first} ${mid} ${suffix}`.replace(/\s+/g, ' ').trim();
               // Clean up stray commas if any parts are missing
               if (rawName.startsWith(",")) rawName = rawName.substring(1).trim();
               if (rawName.endsWith(",")) rawName = rawName.substring(0, rawName.length - 1).trim();
            }
          }
          // -----------------------------------------

          const compositeKey = cleanIdVal;
          if (!groupedData[compositeKey]) {
            groupedData[compositeKey] = {
              key: compositeKey,
              normalizedId: cleanIdVal,
              normalizedName: String(rawName || "").toUpperCase(),
              sources: {}
            };
          } else {
            // Update name if it was missing before but we found it now
            if ((!groupedData[compositeKey].normalizedName || groupedData[compositeKey].normalizedName === "") && rawName) {
               groupedData[compositeKey].normalizedName = String(rawName).toUpperCase();
            }
          }

          let sourceRecord = {};
          fields.forEach(field => {
            let val = getVal(field);

            // === FIX: INJECT CONSTRUCTED NAME ===
            // If the field is "Full Name" and the sheet didn't have a column for it (val is empty),
            // use the 'rawName' we constructed above.
            if (field === "Full Name" && !val && rawName) {
              val = rawName;
            }
            
            // === FIX: N/A LOGIC FOR SHEET 3 ===
            if (sourceObj.key === 'SHEET_3') {
              // Define fields that MUST stay blank if empty (removed "Full Name" from here)
              const keepBlank = [
                "Date Resigned", "Reason for Leaving", "Date Hired", 
                "Last Name", "First Name", "Middle Name", "Suffix", "Date of Birth"
              ];

              // For everything else, if it's empty, make it "N/A"
              if (!keepBlank.includes(field)) {
                 if (val === "" || val === null || val === undefined) {
                   val = "N/A";
                 }
              }
            }
            // -----------------------------------

            if (val instanceof Date) {
              val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
            }
            sourceRecord[field] = val;
          });

          // MERGE LOGIC (Tabs within same file)
          if (!groupedData[compositeKey].sources[fileName]) {
            groupedData[compositeKey].sources[fileName] = sourceRecord;
          } else {
             const existing = groupedData[compositeKey].sources[fileName];
             fields.forEach(f => {
               if (sourceRecord[f] && sourceRecord[f] !== "" && sourceRecord[f] !== "N/A") {
                 existing[f] = sourceRecord[f];
               } else if ((!existing[f] || existing[f] === "") && sourceRecord[f] === "N/A") {
                 existing[f] = "N/A";
               }
             });
          }
        }
      });

    } catch (e) {
      console.error(`Error processing ${sourceObj.key}: ${e.message}`);
    }
  });

  return {
    headers: fields,
    sourceNames: detectedSourceNames,
    data: Object.values(groupedData)
  };
}

/**
 * ==========================================
 * 5. SAVE TO MASTER
 * ==========================================
 */
function saveToMaster(record) {
  try {
    const fields = getMasterHeaders();
    const masterId = cleanId(CONFIG.MASTER.ID);
    
    if (!masterId || masterId.includes("REPLACE")) return { success: false, message: "Master ID not set." };

    const ss = SpreadsheetApp.openById(masterId);
    let sheet = ss.getSheetByName(CONFIG.MASTER.TAB_NAME);
    
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.MASTER.TAB_NAME);
      sheet.appendRow(fields);
      sheet.getRange(1, 1, 1, fields.length).setFontWeight("bold");
    }

    const data = sheet.getDataRange().getValues();
    const idKey = "Employee ID";
    const idColIndex = fields.indexOf(idKey); 

    let rowIndex = -1;
    if (idColIndex !== -1 && data.length > 1) {
       for (let i = 1; i < data.length; i++) {
         if (String(data[i][idColIndex]) === String(record[idKey])) {
           rowIndex = i + 1; 
           break;
         }
       }
    }

    const rowData = fields.map(f => record[f] || "");

    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
      return { success: true, message: "Record Updated" };
    } else {
      sheet.appendRow(rowData);
      return { success: true, message: "Record Added" };
    }
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * ==========================================
 * 6. HELPER FUNCTIONS
 * ==========================================
 */

function letterToColumn(letter) {
  let column = 0;
  const length = letter.length;
  for (let i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

function cleanId(idOrUrl) {
  if (!idOrUrl) return null;
  const match = idOrUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : idOrUrl;
}

function normalizeEmployeeID(rawId) {
  if (!rawId) return "";
  let idStr = String(rawId).trim().replace(/[^0-9]/g, '');
  if (idStr.length > 0 && idStr.length <= 5) {
    idStr = "2400700" + idStr; 
  }
  return idStr;
}
