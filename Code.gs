/**
 * ==========================================
 * 1. CONFIGURATION SECTION
 * ==========================================
 */
const CONFIG = {
  SHEET_1: {
    ID: "1qoi8Lhg2gpmQTNn058hCLnrK3QpI3YKkIzRPMD8gyO0",
    // Example: Pulling from 3 different tabs in this one file
    TABS: ["DATABASE"] 
  },
  SHEET_2: {
    ID: "1y5Ao8kTzVxXc3WCSdthCbrvVmOHGkAZAAdQNjLRNLj8",
    // Example: Only 1 tab
    TABS: ["as of june 2025"] 
  },
  SHEET_3: {
    ID: "1TubY9eVVFyFOy-B_KFKlclmgOZPdASS3Ua4nYtdeJeg",
    // Example: Pulling from 2 tabs
    TABS: ["Personal Information", "HR Fields", "Contact Information", "Contact Information", "Address"] 
  },
  MASTER: {
    ID: "1Ll9L8D7rkze9vQVgFxs48besFwBiStz0eAIEUPrpTAY",
    TAB_NAME: "MasterDatabase", 
    CONFIG_TAB: "Config"   
  }
};

const HARDCODED_MAPPING = {
  "2 Consolidated_Employee Database_v3": {
    "DATABASE": {
      headerRow: 6,
      columns: {
        "Employee ID": 2, "Full Name": 3, "Last Name": 4, "First Name": 5, "Middle Name": 7, "Suffix": 6,
        "Nickname": 15, "Gender": 16, "Date of Birth": 18, "Civil Status": 20, "Active": 8,
        "Employment Type": 58, "Company Name": 115, "Division": 45, "Group": 116, "Department": 42,
        "Section": 46, "Work Location": 41, "Position": 47, "Job Group": 52, "Superior ID": 56,
        "Superior Name": 57, "Date Hired": 59, "Date Regular": 60, "Date Resigned": 63,
        "Reason for Leaving": 70
      }
    }
  },
  "Masterlist": {
    "as of june 2025": {
      headerRow: 2,
      columns: {
        "Employee ID": 2, "Full Name": 3, "Last Name": 4, "First Name": 5, "Middle Name": 7, "Suffix": 6,
        "Nickname": 9, "Gender": 10, "Date of Birth": 11, "Civil Status": 12, "Active": 8,
        "Employment Type": 23, "Company Name": 14, "Division": 15, "Group": 16, "Department": 17,
        "Section": 18, "Work Location": 13, "Position": 19, "Job Group": 20, "Superior ID": 21,
        "Superior Name": 22, "Date Hired": 24, "Date Regular": 25, "Date Resigned": 26,
        "Reason for Leaving": 27
      }
    }
  },
  "Personal Information": {
    "Personal Information": {
      headerRow: 1,
      columns: {
        "Last Name": 3, "First Name": 4, "Middle Name": 6, "Suffix": 5,
        "Civil Status": 16, "Date of Birth": 10, "Nickname": 8, "Gender": 9
      }
    },
    "HR FIELDS": {
      headerRow: 1,
      columns: {
        "Active": 15, "Employee ID": 1, "Employment Type": 14,
        "Company Name": 2, "Division": 3, "Group": 4, "Department": 5,
        "Section": 6, "Work Location": 7, "Position": 34,
        "Superior ID": 20, "Superior Name": 21, "Date Hired": 22,
        "Date Resigned": 16, "Reason for Leaving": 18
      }
    }
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
 * 3. DYNAMIC CONFIGURATION
 * ==========================================
 */
function getColumnConfig() {
  const masterId = cleanId(CONFIG.MASTER.ID);
  if (!masterId) return getDefaults();

  let ss;
  try {
    ss = SpreadsheetApp.openById(masterId);
  } catch (e) {
    console.error("Error opening Master: " + e.message);
    return getDefaults();
  }

  let sheet = ss.getSheetByName(CONFIG.MASTER.CONFIG_TAB);

  // If Config tab missing, create it
  if (!sheet) {
    try {
      sheet = ss.insertSheet(CONFIG.MASTER.CONFIG_TAB);
      sheet.getRange(1, 1, 1, 2)
        .setValues([["Master Column Name", "Search Keywords (comma separated)"]])
        .setFontWeight("bold")
        .setBackground("#e5e7eb");
      
      const defaultMap = getDefaults();
      const defaultsArr = Object.entries(defaultMap).map(([k, v]) => [k, v.join(', ')]);
      sheet.getRange(2, 1, defaultsArr.length, 2).setValues(defaultsArr);
      return defaultMap;
    } catch (e) {
      return getDefaults();
    }
  }

  // Read existing config
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return getDefaults();
  return convertArrayToConfig(sheet.getRange(2, 1, lastRow - 1, 2).getValues());
}

function getDefaults() {
  // These are the DEFAULT keywords. 
  // IF DATA IS MISSING: Go to the 'Config' tab in your Master Sheet and add the exact header name there.
  return {
    "Employee ID":   ["employee id", "employee no", "id no", "badge", "id"],
    "Full Name":     ["full name", "name", "name of staff", "employee name"],
    "Last Name":     ["last name", "family name", "surname"],
    "First Name":    ["first name", "given name"],
    "Middle Name":   ["middle name", "mi"],
    "Suffix":        ["suffix"],
    "Nickname":      ["nickname"],
    "Gender":        ["gender", "sex"],
    "Date of Birth": ["date of birth", "birth", "dob"],
    "Civil Status":  ["civil status", "marital status"],
    "Active":        ["active", "status", "employee status"],
    "Employment Type": ["employment type", "emp type", "contract type"],
    "Company Name":  ["company name", "company"],
    "Division":      ["division", "division name", "div"],
    "Group":         ["group", "group name", "grp"], 
    "Department":    ["department", "dept"],
    "Section":       ["section", "section name", "sec"],
    "Work Location": ["work location", "base code", "location", "site"],
    "Position":      ["position", "position title", "job title"],
    "Job Group":     ["job group", "job level"],
    "Superior ID":   ["immediate superior code", "superior id", "manager id", "immediate superior employee number"],
    "Superior Name": ["immediate superior name", "superior name", "manager name"],
    "Date Hired":    ["date hired", "hiring date"],
    "Date Regular":  ["date regular"],
    "Date Resigned": ["date resigned", "separation date", "date of separation"],
    "Reason for Leaving": ["reason for leaving", "reason", "reason for resignation"]
  };
}

function convertArrayToConfig(rows) {
  const map = {};
  rows.forEach(row => {
    const key = String(row[0]).trim();
    const keywords = String(row[1]).split(',').map(v => v.trim());
    if (key) map[key] = keywords;
  });
  return map;
}

/**
 * ==========================================
 * 4. FETCH & CONSOLIDATE DATA
 * ==========================================
 */
function getConsolidatedData(filters = {}) {
  const cache = CacheService.getScriptCache();
  const cacheKey = "consolidated_data";
  const cachedValue = cache.get(cacheKey);

  if (cachedValue) {
    console.log("Serving from cache...");
    let jsonString;
    try {
      // Check if data is compressed
      if (cachedValue.startsWith("COMPRESSED:")) {
        const base64String = cachedValue.substring("COMPRESSED:".length);
        const bytes = Utilities.base64Decode(base64String);
        const blob = Utilities.newBlob(bytes, 'application/zip');
        jsonString = Utilities.unzip(blob)[0].getDataAsString();
      } else {
        jsonString = cachedValue; // Backwards compatibility
      }
      const allData = JSON.parse(jsonString);
      return filterData(allData, filters);
    } catch (e) {
      console.error("Error reading from cache, fetching fresh data. Error: " + e.toString());
      // If cache is corrupted or in a bad state, proceed to fetch fresh data.
    }
  }

  console.log("Fetching fresh data...");
  const columnMapping = getColumnConfig();
  const fields = Object.keys(columnMapping);

  // Get all records from the Master DB to check for saved status and get master data
  const masterId = cleanId(CONFIG.MASTER.ID);
  const savedRecords = new Map();
  if (masterId) {
    try {
      const masterSheet = SpreadsheetApp.openById(masterId).getSheetByName(CONFIG.MASTER.TAB_NAME);
      if (masterSheet && masterSheet.getLastRow() > 1) {
        const masterData = masterSheet.getDataRange().getValues();
        const masterHeaders = masterData[0].map(h => String(h).trim()); // Trim headers to be safe
        const idColIndex = masterHeaders.indexOf("Employee ID");

        if (idColIndex !== -1) {
          // Create a map from the master header name to its column index for efficient lookup
          const masterHeaderMap = {};
          masterHeaders.forEach((h, i) => {
              masterHeaderMap[h] = i;
          });

          for (let i = 1; i < masterData.length; i++) {
            const row = masterData[i];
            const id = String(row[idColIndex]);
            if (id) {
              const record = {};
              // Iterate over the app's standard fields (e.g., "Full Name", "Position")
              fields.forEach(field => {
                // Find the column index for that standard field in the Master Sheet
                const idx = masterHeaderMap[field];
                let val = (idx !== undefined) ? row[idx] : "";
                 if (val instanceof Date) {
                  val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
                }
                record[field] = val;
              });
              savedRecords.set(id, record);
            }
          }
        }
      }
    } catch (e) { console.error("Could not read master records: " + e.message); }
  }
  
  const fileSources = [
    { id: cleanId(CONFIG.SHEET_1.ID), tabs: CONFIG.SHEET_1.TABS },
    { id: cleanId(CONFIG.SHEET_2.ID), tabs: CONFIG.SHEET_2.TABS },
    { id: cleanId(CONFIG.SHEET_3.ID), tabs: CONFIG.SHEET_3.TABS }
  ];

  let groupedData = {};
  let detectedSourceNames = []; // We will collect real file names here

  fileSources.forEach(source => {
    try {
      if (!source.id) return;
      const ss = SpreadsheetApp.openById(source.id);
      const fileName = ss.getName(); // GET REAL FILE NAME
      
      // Add to our list of sources if not already there
      if (!detectedSourceNames.includes(fileName)) {
        detectedSourceNames.push(fileName);
      }

      let tabsToProcess = (source.tabs && source.tabs.length > 0) ? source.tabs : [ss.getSheets()[0].getName()];

      tabsToProcess.forEach(tabName => {
        const sheet = ss.getSheetByName(tabName);
        if (!sheet) return;

        const allValues = sheet.getDataRange().getValues();
        const mapping = HARDCODED_MAPPING[fileName]?.[tabName];

        if (!mapping || allValues.length < mapping.headerRow) return;

        const columnIndices = mapping.columns;
        const headerRowIndex = mapping.headerRow - 1;

        for (let i = headerRowIndex + 1; i < allValues.length; i++) {
          const row = allValues[i];
          if (row.every(cell => cell === "")) continue;

          const idCol = columnIndices["Employee ID"];
          // If Employee ID isn't in this tab's mapping, we can't process the row.
          if (idCol === undefined) continue;

          const rawID = row[idCol];
          const cleanIdVal = normalizeEmployeeID(rawID);
          if(!cleanIdVal) continue;

          const compositeKey = cleanIdVal;

          let rawName = "";
          if (columnIndices["Full Name"] !== undefined) {
             rawName = row[columnIndices["Full Name"]];
          } else {
            const last = row[columnIndices["Last Name"]] || "";
            const first = row[columnIndices["First Name"]] || "";
            const middle = row[columnIndices["Middle Name"]] || "";
            const suffix = row[columnIndices["Suffix"]] || "";
            let parts = [last, first, middle].filter(Boolean);
            if (parts.length > 0) rawName = parts.join(", ");
            if (suffix) rawName += `, ${suffix}`;
          }

          if (!groupedData[compositeKey]) {
            groupedData[compositeKey] = {
              key: compositeKey,
              normalizedId: cleanIdVal,
              normalizedName: normalizeText(rawName),
              isSaved: savedRecords.has(cleanIdVal),
              masterRecord: savedRecords.get(cleanIdVal) || null,
              sources: {}
            };
          }

          if (!groupedData[compositeKey].sources[fileName]) {
            groupedData[compositeKey].sources[fileName] = {};
          }

          const sourceRecord = groupedData[compositeKey].sources[fileName];
          fields.forEach(field => {
            const colIdx = columnIndices[field];
            if (colIdx !== undefined && !sourceRecord[field]) {
              let val = row[colIdx];
              if (val instanceof Date) {
                val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
              }
              sourceRecord[field] = val;
            }
          });
        }
      });

    } catch (e) {
      console.error(`Error reading source: ${e.message}`);
    }
  });

  let finalData = Object.values(groupedData);

  const fullDataObject = {
    headers: fields,
    sourceNames: detectedSourceNames,
    data: Object.values(groupedData)
  };

  try {
    const jsonString = JSON.stringify(fullDataObject);
    const blob = Utilities.newBlob(jsonString, 'text/plain', 'data.json');
    const zippedBlob = Utilities.zip([blob]);
    const base64String = Utilities.base64Encode(zippedBlob.getBytes());

    // Prefix to identify compressed content later
    cache.put(cacheKey, "COMPRESSED:" + base64String, 3600);
    console.log("Data cached successfully (compressed).");

  } catch(e) {
    console.error("Could not cache data: " + e.toString());
    // If caching fails, just return the data without caching
  }

  return filterData(fullDataObject, filters);
}

function filterData(allData, filters) {
  let finalData = allData.data;
  const filterKeys = Object.keys(filters).filter(key => filters[key]);

  if (filterKeys.length > 0) {
    finalData = finalData.filter(employee => {
      return filterKeys.every(key => {
        return Object.values(employee.sources).some(source => source[key] === filters[key]);
      });
    });
  }

  return {
    headers: allData.headers,
    sourceNames: allData.sourceNames,
    data: finalData
  };
}


function getSingleEmployee(employeeId) {
  // This is a simplified version; in a real app, you might want to optimize this
  // to avoid re-scanning all sheets just for one employee.
  const allData = getConsolidatedData().data;
  const employee = allData.find(e => e.normalizedId === employeeId);
  return employee;
}

function getFilterOptions() {
  const fieldsToFilter = ["Work Location", "Division", "Group", "Department", "Section"];
  const options = {};

  // This is not the most performant way for large datasets, but it's simple.
  // A better way would be to cache these values.
  const allData = getConsolidatedData().data;

  fieldsToFilter.forEach(field => {
    const values = new Set();
    allData.forEach(employee => {
      Object.values(employee.sources).forEach(source => {
        if (source[field]) {
          values.add(source[field]);
        }
      });
    });
    options[field] = Array.from(values).sort();
  });

  return options;
}

/**
 * ==========================================
 * 5. SAVE TO MASTER DATABASE
 * ==========================================
 */
function saveToMaster(record) {
  try {
    const columnMapping = getColumnConfig();
    const fields = Object.keys(columnMapping);
    const masterId = cleanId(CONFIG.MASTER.ID);

    if (!masterId) throw new Error("Invalid Master Spreadsheet ID");

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
      if (sheet.getMaxColumns() < rowData.length) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), rowData.length - sheet.getMaxColumns());
      }
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
      CacheService.getScriptCache().remove("consolidated_data"); // Invalidate cache
      return { success: true, message: "Record Updated in Master" };
    } else {
      sheet.appendRow(rowData);
      CacheService.getScriptCache().remove("consolidated_data"); // Invalidate cache
      return { success: true, message: "Record Added to Master" };
    }

  } catch (e) {
    return { success: false, message: "Error saving: " + e.toString() };
  }
}

/**
 * ==========================================
 * 6. HELPER FUNCTIONS
 * ==========================================
 */
function cleanId(idOrUrl) {
  if (!idOrUrl || idOrUrl.includes("REPLACE")) return null;
  const match = idOrUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) return match[1];
  return idOrUrl;
}

function normalizeEmployeeID(rawId) {
  if (!rawId) return "";
  let idStr = String(rawId).trim().replace(/[^0-9]/g, '');
  if (idStr.length > 0 && idStr.length <= 5) {
    idStr = "2400700" + idStr; 
  }
  return idStr;
}

function normalizeText(text) {
  return String(text || "").trim().toUpperCase();
}
