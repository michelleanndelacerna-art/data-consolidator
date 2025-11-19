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
 * 4. FETCH & CONSOLIDATE DATA (WITH CACHING)
 * ==========================================
 */

// This "private" function does the heavy lifting of reading the sheets.
function _fetchRawDataFromSources() {
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
    if (!sheetConfig.ID || sheetConfig.ID.includes("REPLACE")) return;

    try {
      const ss = SpreadsheetApp.openById(cleanId(sheetConfig.ID));
      const fileName = ss.getName();
      if (!detectedSourceNames.includes(fileName)) detectedSourceNames.push(fileName);

      sheetConfig.TABS.forEach(tabConfig => {
        const sheet = ss.getSheetByName(tabConfig.NAME);
        if (!sheet) return;

        const allValues = sheet.getDataRange().getValues();
        const startRow = tabConfig.DATA_START_ROW - 1;
        if (allValues.length <= startRow) return;

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

          let rawName = getVal("Full Name");
          if (!rawName) {
            const last = getVal("Last Name"), first = getVal("First Name"), mid = getVal("Middle Name"), suffix = getVal("Suffix");
            if (last || first) {
               rawName = `${last}, ${first} ${mid} ${suffix}`.replace(/\s+/g, ' ').trim();
               if (rawName.startsWith(",")) rawName = rawName.substring(1).trim();
               if (rawName.endsWith(",")) rawName = rawName.substring(0, rawName.length - 1).trim();
            }
          }

          if (!groupedData[cleanIdVal]) {
            groupedData[cleanIdVal] = {
              key: cleanIdVal,
              normalizedId: cleanIdVal,
              normalizedName: String(rawName || "").toUpperCase(),
              sources: {}
            };
          } else if ((!groupedData[cleanIdVal].normalizedName || groupedData[cleanIdVal].normalizedName === "") && rawName) {
            groupedData[cleanIdVal].normalizedName = String(rawName).toUpperCase();
          }

          let sourceRecord = {};
          fields.forEach(field => {
            let val = getVal(field);
            if (field === "Full Name" && !val && rawName) val = rawName;
            
            if (sourceObj.key === 'SHEET_3') {
              const keepBlank = ["Date Resigned", "Reason for Leaving", "Date Hired", "Last Name", "First Name", "Middle Name", "Suffix", "Date of Birth"];
              if (!keepBlank.includes(field) && (val === "" || val === null || val === undefined)) {
                 val = "N/A";
              }
            }
            if (val instanceof Date) val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
            sourceRecord[field] = val;
          });

          if (!groupedData[cleanIdVal].sources[fileName]) {
            groupedData[cleanIdVal].sources[fileName] = sourceRecord;
          } else {
             const existing = groupedData[cleanIdVal].sources[fileName];
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
    } catch (e) { console.error(`Error processing ${sourceObj.key}: ${e.message}`); }
  });

  return { groupedData, detectedSourceNames };
}

function getConsolidatedData(filters = {}) {
  const fields = getMasterHeaders();
  const cache = CacheService.getScriptCache();
  const cacheKeyMeta = 'consolidated_data_v4_meta';
  const cacheKeyChunkPrefix = 'consolidated_data_v4_chunk_';
  const CHUNK_SIZE = 90000; // 90KB, well below the 100KB limit

  let groupedData, detectedSourceNames;

  // --- 1. TRY TO GET FROM CACHE USING CHUNKING ---
  const metaCached = cache.get(cacheKeyMeta);
  if (metaCached) {
    console.log("CACHE HIT: Found metadata. Reconstructing from chunks...");
    try {
      const meta = JSON.parse(metaCached);
      const chunkKeys = [];
      for (let i = 0; i < meta.chunkCount; i++) {
        chunkKeys.push(cacheKeyChunkPrefix + i);
      }

      const cachedChunks = cache.getAll(chunkKeys);
      let base64Encoded = "";
      let success = true;

      for (let i = 0; i < meta.chunkCount; i++) {
        const key = cacheKeyChunkPrefix + i;
        if (cachedChunks[key]) {
          base64Encoded += cachedChunks[key];
        } else {
          console.error("CACHE ERROR: Missing chunk " + i + ". Aborting cache read.");
          success = false;
          break;
        }
      }

      if (success) {
        const decompressed = Utilities.unzip(Utilities.base64Decode(base64Encoded));
        const dataObj = JSON.parse(decompressed.getDataAsString());
        groupedData = dataObj.groupedData;
        detectedSourceNames = dataObj.detectedSourceNames;
      }
    } catch(e) {
      console.error("CACHE ERROR: Could not parse or reconstruct from cache. Refetching. Error: " + e.message);
    }
  }

  // --- 2. IF CACHE MISS OR CORRUPT, FETCH FRESH DATA ---
  if (!groupedData) {
    console.log("CACHE MISS: Fetching fresh data from sources...");
    const freshData = _fetchRawDataFromSources();
    groupedData = freshData.groupedData;
    detectedSourceNames = freshData.detectedSourceNames;

    console.log("STORING IN CACHE: Zipping, chunking, and storing fresh data...");
    try {
      const dataToCache = { groupedData, detectedSourceNames };
      const blob = Utilities.newBlob(JSON.stringify(dataToCache), 'application/json');
      const compressed = Utilities.zip([blob]);
      const encodedData = Utilities.base64Encode(compressed.getBytes());

      const numChunks = Math.ceil(encodedData.length / CHUNK_SIZE);
      const chunks = {};

      for (let i = 0; i < numChunks; i++) {
        const chunk = encodedData.substring(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
        chunks[cacheKeyChunkPrefix + i] = chunk;
      }

      chunks[cacheKeyMeta] = JSON.stringify({ chunkCount: numChunks });
      cache.putAll(chunks, 3600); // Cache all chunks and meta for 1 hour
    } catch (e) {
      console.error("CACHE WRITE ERROR: Could not store chunks. Error: " + e.message);
      // This may be due to total cache size limit, but we'll proceed with the fresh data for this request.
    }
  }

  // --- 3. FETCH MASTER DB FOR REAL-TIME SAVED STATUS (THIS IS FAST) ---
  const masterRecords = {};
  try {
    const masterId = cleanId(CONFIG.MASTER.ID);
    const ss = SpreadsheetApp.openById(masterId);
    const sheet = ss.getSheetByName(CONFIG.MASTER.TAB_NAME);
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      const headers = data.shift();
      const idColIndex = headers.indexOf("Employee ID");
      if (idColIndex !== -1) {
        data.forEach(row => {
          const id = normalizeEmployeeID(row[idColIndex]);
          if (id) {
            const record = {};
            headers.forEach((h, i) => record[h] = row[i]);
            masterRecords[id] = record;
          }
        });
      }
    }
  } catch(e) { console.error("Could not fetch Master Database: " + e.message); }

  // --- 3. PROCESS AND FILTER THE DATA ---
  let finalData = Object.values(groupedData);

  finalData.forEach(item => {
    if (masterRecords[item.normalizedId]) {
      item.masterRecord = masterRecords[item.normalizedId];
    }
  });

  const activeFilters = Object.entries(filters).filter(([_, value]) => value);
  if (activeFilters.length > 0) {
    finalData = finalData.filter(item => {
      return activeFilters.every(([field, value]) => {
        for (const sourceName in item.sources) {
          if (item.sources[sourceName] && item.sources[sourceName][field] === value) return true;
        }
        return false;
      });
    });
  }

  return {
    headers: fields,
    sourceNames: detectedSourceNames,
    data: finalData
  };
}

/**
 * ==========================================
 * 5. GET FILTER OPTIONS (WITH CACHING)
 * ==========================================
 */
function getFilterOptions() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'filter_options_v2';
  const cached = cache.get(cacheKey);

  if (cached) {
    console.log("CACHE HIT: Returning filter options from cache.");
    return JSON.parse(cached);
  }

  console.log("CACHE MISS: Fetching fresh filter options...");
  const fieldsToFilter = ["Work Location", "Division", "Group", "Department", "Section"];
  const options = {};
  fieldsToFilter.forEach(f => options[f] = new Set());

  const sources = [
    { key: 'SHEET_1', config: CONFIG.SHEET_1 },
    { key: 'SHEET_2', config: CONFIG.SHEET_2 },
    { key: 'SHEET_3', config: CONFIG.SHEET_3 }
  ];

  sources.forEach(sourceObj => {
    const sheetConfig = sourceObj.config;
    if (!sheetConfig.ID || sheetConfig.ID.includes("REPLACE")) return;

    try {
      const ss = SpreadsheetApp.openById(cleanId(sheetConfig.ID));
      sheetConfig.TABS.forEach(tabConfig => {
        const sheet = ss.getSheetByName(tabConfig.NAME);
        if (!sheet) return;
        const allValues = sheet.getDataRange().getValues();
        const startRow = tabConfig.DATA_START_ROW - 1;
        if (allValues.length <= startRow) return;

        for (let i = startRow; i < allValues.length; i++) {
          const row = allValues[i];
          if (row.every(c => c === "")) continue;
          fieldsToFilter.forEach(field => {
            const colLetter = tabConfig.MAP[field];
            if (colLetter) {
              const colIndex = letterToColumn(colLetter) - 1;
              const val = (colIndex >= 0 && colIndex < row.length) ? row[colIndex] : "";
              if (val && String(val).trim() !== "" && String(val).trim() !== "N/A") {
                options[field].add(String(val).trim());
              }
            }
          });
        }
      });
    } catch (e) { console.error(`Error getting filter options from ${sourceObj.key}: ${e.message}`); }
  });

  const finalOptions = {};
  for (const field in options) {
    finalOptions[field] = Array.from(options[field]).sort();
  }

  console.log("STORING IN CACHE: Storing filter options.");
  cache.put(cacheKey, JSON.stringify(finalOptions), 3600); // Cache for 1 hour

  return finalOptions;
}

/**
 * ==========================================
 * 6. SAVE TO MASTER
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
    let message = "";

    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
      message = "Record Updated";
    } else {
      sheet.appendRow(rowData);
      message = "Record Added";
    }

    // --- CLEAR CACHE ON SUCCESSFUL SAVE ---
    console.log("Record saved. Clearing cache...");
    const cache = CacheService.getScriptCache();
    const cacheKeyMeta = 'consolidated_data_v4_meta';
    const cacheKeyChunkPrefix = 'consolidated_data_v4_chunk_';
    const filterCacheKey = 'filter_options_v2';

    const metaCached = cache.get(cacheKeyMeta);
    if (metaCached) {
        try {
            const meta = JSON.parse(metaCached);
            const keysToRemove = [cacheKeyMeta];
            for (let i = 0; i < meta.chunkCount; i++) {
                keysToRemove.push(cacheKeyChunkPrefix + i);
            }
            cache.removeAll(keysToRemove);
            console.log(`Removed ${keysToRemove.length} data cache keys.`);
        } catch (e) {
            console.error("Cache clear error, removing meta key only.", e);
            cache.remove(cacheKeyMeta);
        }
    }
    cache.remove(filterCacheKey);
    console.log("Removed filter options cache.");
    // ------------------------------------

    return { success: true, message: message };
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
