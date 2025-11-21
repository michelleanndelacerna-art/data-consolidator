/**
 * ==========================================
 * 1. CONFIGURATION SECTION
 * ==========================================
 */
const CONFIG = {
  SHEET_1: {
    ID: "1qoi8Lhg2gpmQTNn058hCLnrK3QpI3YKkIzRPMD8gyO0", 
    TABS: [{ NAME: "DATABASE", DATA_START_ROW: 7, MAP: { "Employee ID": "C", "Full Name": "D", "Last Name": "E", "First Name": "F", "Middle Name": "H", "Suffix": "G", "Nickname": "P", "Gender": "Q", "Date of Birth": "S", "Civil Status": "U", "Active": "I", "Employment Type": "BG", "Company Name": "DL", "Division": "AT", "Group": "DM", "Department": "AQ", "Section": "AU", "Work Location": "AP", "Position": "AV", "Job Group": "BA", "Superior ID": "BE", "Superior Name": "BF", "Date Hired": "BH", "Date Regular": "BI", "Date Resigned": "BL", "Reason for Leaving": "BS" } }]
  },
  SHEET_2: {
    ID: "1y5Ao8kTzVxXc3WCSdthCbrvVmOHGkAZAAdQNjLRNLj8", 
    TABS: [{ NAME: "as of June 2025", DATA_START_ROW: 3, MAP: { "Employee ID": "C", "Full Name": "D", "Last Name": "E", "First Name": "F", "Middle Name": "H", "Suffix": "G", "Nickname": "J", "Gender": "K", "Date of Birth": "L", "Civil Status": "M", "Active": "I", "Employment Type": "X", "Company Name": "O", "Division": "P", "Group": "Q", "Department": "R", "Section": "S", "Work Location": "N", "Position": "T", "Job Group": "U", "Superior ID": "V", "Superior Name": "W", "Date Hired": "Y", "Date Regular": "Z", "Date Resigned": "AA", "Reason for Leaving": "AB" } }]
  },
  SHEET_3: {
    ID: "1TubY9eVVFyFOy-B_KFKlclmgOZPdASS3Ua4nYtdeJeg", 
    TABS: [
      { NAME: "Personal Information", DATA_START_ROW: 2, MAP: { "Employee ID": "B", "Last Name": "D", "First Name": "E", "Middle Name": "G", "Suffix": "F", "Nickname": "I", "Gender": "J", "Date of Birth": "K", "Civil Status": "Q" } },
      { NAME: "HR FIELDS", DATA_START_ROW: 2, MAP: { "Employee ID": "B", "Active": "P", "Employment Type": "O", "Company Name": "C", "Division": "D", "Group": "E", "Department": "F", "Section": "G", "Work Location": "H", "Position": "AI", "Superior ID": "U", "Superior Name": "V", "Date Hired": "W", "Date Resigned": "Q", "Reason for Leaving": "S" } }
    ]
  },
  // --- NEW: SHEET 4 (REFERENCE SOURCE) ---
  SHEET_4: {
    ID: "1UIwtAixLz0wd5lXQD6PvvYkv-pOkOggriMTF_BGdkQk", // <--- PUT YOUR SHEET 4 ID HERE
    TABS: [{ 
        NAME: "Reference Data", // <--- Check your tab name
        DATA_START_ROW: 2, 
        MAP: { 
           "Employee ID": "B", // Adjust Columns as needed
           "Position": "S", 
           "Department": "Q", 
           "Division": "O", 
           "Group": "P", 
           "Section": "R", 
           "Superior Name": "V", 
           "Superior ID": "U" 
        } 
    }]
  },
  MASTER: { ID: "1Ll9L8D7rkze9vQVgFxs48besFwBiStz0eAIEUPrpTAY", TAB_NAME: "MasterDatabase", LOG_TAB: "Logs" }
};

const HEADER_ALIASES = {
  "LAST NAME": ["SURNAME", "FAMILY NAME", "LASTNAME", "LNAME", "L.NAME"],
  "FIRST NAME": ["GIVEN NAME", "FIRSTNAME", "FNAME", "F.NAME"],
  "MIDDLE NAME": ["MIDDLE", "MIDDLENAME", "M.I.", "MI", "MNAME"],
  "SUFFIX": ["EXTENSION", "EXT", "NAME EXTENSION"],
  "EMPLOYEE ID": ["EMP ID", "ID", "EMPLOYEE NO", "EMPLOYEE #", "EMP NO"]
};

function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate().setTitle('Employee Data Consolidator').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getMasterHeaders() {
  return ["Employee ID", "Full Name", "Last Name", "First Name", "Middle Name", "Suffix", "Nickname", "Gender", "Date of Birth", "Civil Status", "Active", "Employment Type", "Company Name", "Division", "Group", "Department", "Section", "Work Location", "Position", "Job Group", "Superior ID", "Superior Name", "Date Hired", "Date Regular", "Date Resigned", "Reason for Leaving"];
}

function _fetchRawDataFromSources() {
  const fields = getMasterHeaders();
  // Added SHEET_4 with isReference flag
  const sources = [
      {key:'SHEET_1',config:CONFIG.SHEET_1},
      {key:'SHEET_2',config:CONFIG.SHEET_2},
      {key:'SHEET_3',config:CONFIG.SHEET_3},
      {key:'SHEET_4',config:CONFIG.SHEET_4, isReference: true} 
  ];
  let groupedData = {};
  let detectedSourceNames = []; 

  sources.forEach(sourceObj => {
    const sheetConfig = sourceObj.config;
    if (!sheetConfig.ID || sheetConfig.ID.includes("REPLACE")) return;
    try {
      const ss = SpreadsheetApp.openById(cleanId(sheetConfig.ID));
      const fileName = ss.getName();
      const fileUrl = ss.getUrl(); 
      
      if (!detectedSourceNames.some(s => s.name === fileName)) {
         // Pass reference flag to frontend
         detectedSourceNames.push({ name: fileName, url: fileUrl, isReference: sourceObj.isReference || false });
      }

      sheetConfig.TABS.forEach(tabConfig => {
        const sheet = ss.getSheetByName(tabConfig.NAME);
        if (!sheet) return;
        
        const headerRowIdx = Math.max(0, tabConfig.DATA_START_ROW - 2);
        const allValues = sheet.getDataRange().getValues();
        const headerMap = {};
        if (allValues.length > headerRowIdx) {
            allValues[headerRowIdx].forEach((h, i) => { 
                if (h) headerMap[String(h).trim().toUpperCase()] = i; 
            });
        }

        const startRow = tabConfig.DATA_START_ROW - 1;
        if (allValues.length <= startRow) return;

        for (let i = startRow; i < allValues.length; i++) {
          const row = allValues[i];
          if (row.every(c => c === "")) continue;

          const getVal = (fieldName) => {
            const key = fieldName.toUpperCase();
            let colIndex = headerMap[key];
            if (colIndex === undefined && HEADER_ALIASES[key]) {
                for (const alias of HEADER_ALIASES[key]) {
                    if (headerMap[alias] !== undefined) { colIndex = headerMap[alias]; break; }
                }
            }
            if (colIndex === undefined) {
                 const colLetter = tabConfig.MAP[fieldName];
                 if (colLetter) colIndex = letterToColumn(colLetter) - 1;
            }
            return (colIndex !== undefined && colIndex >= 0 && colIndex < row.length) ? row[colIndex] : "";
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
                sources: {},
                originalIds: [cleanIdVal] 
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
              if (!keepBlank.includes(field) && (val === "" || val === null || val === undefined)) val = "N/A";
            }
            
            if (val instanceof Date) val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
            else val = String(val ?? "").trim();
            
            sourceRecord[field] = val;
          });

          if (!groupedData[cleanIdVal].sources[fileName]) {
            groupedData[cleanIdVal].sources[fileName] = sourceRecord;
          } else {
             const existing = groupedData[cleanIdVal].sources[fileName];
             fields.forEach(f => {
               if (sourceRecord[f] && sourceRecord[f] !== "" && sourceRecord[f] !== "N/A") existing[f] = sourceRecord[f];
               else if ((!existing[f] || existing[f] === "") && sourceRecord[f] === "N/A") existing[f] = "N/A";
             });
          }
        }
      });
    } catch (e) { console.error(`Error processing ${sourceObj.key}: ${e.message}`); }
  });

  let recordList = Object.values(groupedData);
  recordList.sort((a, b) => a.normalizedName.localeCompare(b.normalizedName));

  const mergedData = {};
  const skipIds = new Set();

  for (let i = 0; i < recordList.length; i++) {
      if (skipIds.has(recordList[i].key)) continue;

      let current = recordList[i];
      for (let j = i + 1; j < recordList.length; j++) {
          let next = recordList[j];
          if (next.normalizedName[0] !== current.normalizedName[0]) break; 
          const dist = getLevenshteinDistance(current.normalizedName, next.normalizedName);
          if (dist <= 3 && current.key !== next.key) {
              for (const [srcName, srcData] of Object.entries(next.sources)) {
                  if (!current.sources[srcName]) {
                      current.sources[srcName] = srcData;
                  } else {
                      fields.forEach(f => {
                          const val = srcData[f];
                          if (val && val !== "N/A" && (!current.sources[srcName][f] || current.sources[srcName][f] === "N/A")) {
                              current.sources[srcName][f] = val;
                          }
                      });
                  }
              }
              current.originalIds.push(next.key);
              skipIds.add(next.key); 
          }
      }
      mergedData[current.key] = current;
  }

  groupedData = mergedData;
  return { groupedData, detectedSourceNames };
}

function getLevenshteinDistance(s1, s2) {
  if (s1.length === 0) return s2.length;
  if (s2.length === 0) return s1.length;
  const matrix = [];
  for (let i = 0; i <= s2.length; i++) { matrix[i] = [i]; }
  for (let j = 0; j <= s1.length; j++) { matrix[0][j] = j; }
  for (let i = 1; i <= s2.length; i++) {
    for (let j = 1; j <= s1.length; j++) {
      if (s2.charAt(i - 1) === s1.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(matrix[i - 1][j - 1] + 1, Math.min(matrix[i][j - 1] + 1, matrix[i - 1][j] + 1));
      }
    }
  }
  return matrix[s2.length][s1.length];
}

function getConsolidatedData(filters = {}) {
  const fields = getMasterHeaders();
  const cache = CacheService.getScriptCache();
  const cacheKeyMeta = 'consolidated_data_v17_ref'; // Version Bump
  const cacheKeyChunkPrefix = 'consolidated_data_v17_chunk_';
  const CHUNK_SIZE = 90000; 

  let groupedData, detectedSourceNames;
  const metaCached = cache.get(cacheKeyMeta);
  if (metaCached) {
    try {
      const meta = JSON.parse(metaCached);
      const chunkKeys = [];
      for (let i = 0; i < meta.chunkCount; i++) chunkKeys.push(cacheKeyChunkPrefix + i);
      const cachedChunks = cache.getAll(chunkKeys);
      let base64Encoded = "";
      let success = true;
      for (let i = 0; i < meta.chunkCount; i++) {
        if (cachedChunks[cacheKeyChunkPrefix + i]) base64Encoded += cachedChunks[cacheKeyChunkPrefix + i];
        else { success = false; break; }
      }
      if (success) {
        const decompressed = Utilities.unzip(Utilities.base64Decode(base64Encoded));
        const dataObj = JSON.parse(decompressed.getDataAsString());
        groupedData = dataObj.groupedData;
        detectedSourceNames = dataObj.detectedSourceNames;
      }
    } catch(e) {}
  }

  if (!groupedData) {
    const freshData = _fetchRawDataFromSources();
    groupedData = freshData.groupedData;
    detectedSourceNames = freshData.detectedSourceNames;
    try {
      const dataToCache = { groupedData, detectedSourceNames };
      const blob = Utilities.newBlob(JSON.stringify(dataToCache), 'application/json');
      const compressed = Utilities.zip([blob]);
      const encodedData = Utilities.base64Encode(compressed.getBytes());
      const numChunks = Math.ceil(encodedData.length / CHUNK_SIZE);
      const chunks = {};
      for (let i = 0; i < numChunks; i++) chunks[cacheKeyChunkPrefix + i] = encodedData.substring(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
      chunks[cacheKeyMeta] = JSON.stringify({ chunkCount: numChunks });
      cache.putAll(chunks, 3600); 
    } catch (e) {}
  }

  const masterRecords = {};
  try {
    const masterId = cleanId(CONFIG.MASTER.ID);
    const ss = SpreadsheetApp.openById(masterId);
    const sheet = ss.getSheetByName(CONFIG.MASTER.TAB_NAME);
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      const sheetHeaders = data[0].map(h => String(h).trim().toUpperCase());
      const appFields = getMasterHeaders(); 
      const colMap = {};
      appFields.forEach(field => {
        const index = sheetHeaders.indexOf(field.toUpperCase());
        if (index > -1) colMap[field] = index;
      });
      const idKey = "Employee ID";
      if (colMap[idKey] !== undefined) {
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const normalizedId = normalizeEmployeeID(row[colMap[idKey]]);
          if (normalizedId) {
            const record = {};
            Object.entries(colMap).forEach(([field, colIndex]) => {
              let val = row[colIndex];
              if (val instanceof Date) val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
              else val = String(val ?? "").trim(); 
              record[field] = val;
            });
            masterRecords[normalizedId] = record;
          }
        }
      }
    }
  } catch(e) {}

  let finalData = Object.values(groupedData);
  finalData.forEach(item => {
    if (masterRecords[item.normalizedId]) item.masterRecord = masterRecords[item.normalizedId];
  });

  const activeFilters = Object.entries(filters).filter(([_, value]) => value);
  if (activeFilters.length > 0) {
    finalData = finalData.filter(item => {
      return activeFilters.every(([field, value]) => {
        for (const sourceName in item.sources) {
          if (item.sources[sourceName] && String(item.sources[sourceName][field] ?? "").trim() === String(value).trim()) return true;
        }
        return false;
      });
    });
  }
  return { headers: fields, sourceNames: detectedSourceNames, data: finalData };
}

function getFilterOptions() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'filter_options_v17_ref';
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const fieldsToFilter = ["Work Location", "Division", "Group", "Department", "Section"];
  const rawOptions = {};
  fieldsToFilter.forEach(f => rawOptions[f] = new Set());

  const sources = [{key:'SHEET_1',config:CONFIG.SHEET_1},{key:'SHEET_2',config:CONFIG.SHEET_2},{key:'SHEET_3',config:CONFIG.SHEET_3}];

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
              let val = (colIndex >= 0 && colIndex < row.length) ? row[colIndex] : "";
              val = String(val ?? "").trim();
              if (val && val !== "N/A") rawOptions[field].add(val);
            }
          });
        }
      });
    } catch (e) {}
  });

  let orgChartRows = [];
  try {
    const masterId = cleanId(CONFIG.MASTER.ID);
    const ss = SpreadsheetApp.openById(masterId);
    const orgSheet = ss.getSheetByName("orgchartcode");
    if (orgSheet) {
      const data = orgSheet.getDataRange().getValues();
      for(let i = 1; i < data.length; i++) {
         orgChartRows.push({
           Division: String(data[i][0] ?? "").trim(),
           Group: String(data[i][1] ?? "").trim(),
           Department: String(data[i][2] ?? "").trim(),
           Section: String(data[i][3] ?? "").trim()
         });
      }
    }
  } catch (e) {}

  const finalOutput = { raw: {}, orgChart: orgChartRows };
  for (const field in rawOptions) finalOutput.raw[field] = Array.from(rawOptions[field]).sort();
  cache.put(cacheKey, JSON.stringify(finalOutput), 3600);
  return finalOutput;
}

function saveToMaster(record) { return saveBulkToMaster([record]); }

function saveBulkToMaster(records) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(50000)) return { success: false, message: "System busy. Try again." };
  try {
    const fields = getMasterHeaders();
    const masterId = cleanId(CONFIG.MASTER.ID);
    if (!masterId) return { success: false, message: "Master ID not set." };
    
    const ss = SpreadsheetApp.openById(masterId);
    let sheet = ss.getSheetByName(CONFIG.MASTER.TAB_NAME);
    if (!sheet) { sheet = ss.insertSheet(CONFIG.MASTER.TAB_NAME); sheet.appendRow(fields); }

    const data = sheet.getDataRange().getValues();
    const idKey = "Employee ID";
    const idColIndex = fields.indexOf(idKey); 
    const idMap = new Map();
    if (idColIndex !== -1 && data.length > 1) { for (let i = 1; i < data.length; i++) idMap.set(String(data[i][idColIndex]), i); }

    let added = 0, updated = 0;
    const newRows = [];
    records.forEach(record => {
       const id = String(record[idKey]);
       const rowData = fields.map(f => record[f] || "");
       if (idMap.has(id)) { data[idMap.get(id)] = rowData; updated++; }
       else { newRows.push(rowData); added++; }
       const logMsg = `Action: ${idMap.has(id) ? 'UPDATE' : 'CREATE'} - ${record[fields[1]]}`;
       logChange(ss, idMap.has(id) ? "UPDATE" : "CREATE", id, logMsg);
    });

    if (updated > 0) sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    if (newRows.length > 0) sheet.getRange(data.length + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    
    const cache = CacheService.getScriptCache();
    const keys = ['consolidated_data_v17_ref', 'filter_options_v17_ref'];
    cache.removeAll(keys);
    return { success: true, message: `Bulk Saved: ${updated} Updated, ${added} Added.` };
  } catch (e) { return { success: false, message: e.toString() }; } finally { lock.releaseLock(); }
}

function logChange(ss, action, id, details) {
  try {
    let logSheet = ss.getSheetByName(CONFIG.MASTER.LOG_TAB);
    if (!logSheet) {
      logSheet = ss.insertSheet(CONFIG.MASTER.LOG_TAB);
      logSheet.appendRow(["Timestamp", "User", "Action", "Employee ID", "Details"]);
      logSheet.getRange(1, 1, 1, 5).setFontWeight("bold");
    }
    const user = Session.getActiveUser().getEmail() || "Unknown";
    logSheet.appendRow([new Date(), user, action, id, details]);
  } catch(e) { console.error("Logging failed", e); }
}

function getEmployeeLogs(employeeId) {
  try {
      const masterId = cleanId(CONFIG.MASTER.ID);
      const ss = SpreadsheetApp.openById(masterId);
      const sheet = ss.getSheetByName(CONFIG.MASTER.LOG_TAB);
      if (!sheet) return [];
      const data = sheet.getDataRange().getValues();
      const logs = data.slice(1).filter(row => String(row[3]) === String(employeeId)).map(row => ({
           timestamp: row[0] instanceof Date ? row[0].toLocaleString() : String(row[0]),
           user: row[1],
           action: row[2],
           details: row[4]
      }));
      return logs.reverse(); 
  } catch(e) { return []; }
}

function letterToColumn(letter) {
  let column = 0, length = letter.length;
  for (let i = 0; i < length; i++) column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  return column;
}
function cleanId(idOrUrl) { if (!idOrUrl) return null; const match = idOrUrl.match(/\/d\/([a-zA-Z0-9-_]+)/); return match ? match[1] : idOrUrl; }
function normalizeEmployeeID(rawId) { if (!rawId) return ""; let idStr = String(rawId).trim().replace(/[^0-9]/g, ''); if (idStr.length > 0 && idStr.length <= 5) idStr = "2400700" + idStr; return idStr; }
