/**
 * =========================================================================
 * QC-DESU DISTRICT 2 MASTER SURVEILLANCE & AUTHENTICATION BACKEND
 * =========================================================================
 */

// PASTE YOUR ACTUAL SPREADSHEET ID HERE
const SPREADSHEET_ID = "1BgSSx1eaF0Sxy4yNaXV60sEu5enUL06FyqPOrweYAmI";

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('QC-DESU District 2 Master Surveillance Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSpreadsheet() {
  try {
    if (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== "") {
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    }
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    Logger.log("getSpreadsheet error: " + e.message);
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

function getSheetByNameInsensitive(ss, name) {
  if (!ss) return null;
  const sheets = ss.getSheets();
  const cleanTarget = name.trim().toUpperCase();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().trim().toUpperCase() === cleanTarget) {
      return sheets[i];
    }
  }
  return null;
}

function getMdbSheet(ss) {
  const activeSs = ss || getSpreadsheet();
  if (!activeSs) return null;
  return getSheetByNameInsensitive(activeSs, "MDB DISTRICT 2 2026") || activeSs.getSheets()[0];
}

function getEpisenseSheet(ss) {
  const activeSs = ss || getSpreadsheet();
  if (!activeSs) return null;
  return getSheetByNameInsensitive(activeSs, "EPISENSE") || activeSs.getSheets()[0];
}

function getUsersSheet(ss) {
  const activeSs = ss || getSpreadsheet();
  if (!activeSs) return null;
  return getSheetByNameInsensitive(activeSs, "users") || getSheetByNameInsensitive(activeSs, "USERS") || getSheetByNameInsensitive(activeSs, "ACCOUNTS");
}

/**
 * 0. AUTHENTICATION SERVICE
 */
function processLogin(data) {
  try {
    const usernameInput = String(data.username || '').trim().toLowerCase();
    const passwordInput = String(data.password || '').trim();

    if (!usernameInput || !passwordInput) {
      return { success: false, message: 'Please enter both username and password.' };
    }

    const ss = getSpreadsheet();
    if (!ss) {
      return { success: false, message: 'DATABASE_ERROR: Cannot access spreadsheet. Check SPREADSHEET_ID.' };
    }

    const sheet = getUsersSheet(ss);
    if (!sheet) {
      return { success: false, message: 'DATABASE_OFFLINE: "users" tab not found in the spreadsheet.' };
    }

    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) {
      return { success: false, message: 'DATABASE_EMPTY: No user accounts registered.' };
    }

    // Row 0 is header: [Username, Password, Status, ...]
    for (let i = 1; i < rows.length; i++) {
      const dbUser = String(rows[i][0] || '').trim().toLowerCase();
      const dbPass = String(rows[i][1] || '').trim();
      const dbStatus = String(rows[i][2] || '').trim().toLowerCase();

      if (dbUser === usernameInput) {
        if (dbStatus !== 'active' && dbStatus !== '') {
          return { success: false, message: 'ACCESS_DENIED: Account status is ' + (dbStatus.toUpperCase() || 'INACTIVE') + '.' };
        }

        if (dbPass === passwordInput) {
          return { success: true, message: 'ACCESS_GRANTED: Authentication successful.', user: rows[i][0] };
        } else {
          return { success: false, message: 'SECURITY_ALERT: Incorrect password.' };
        }
      }
    }

    return { success: false, message: 'NODE_NOT_FOUND: User ID / Email not recognized.' };
  } catch (error) {
    Logger.log("Error in processLogin: " + error.toString());
    return { success: false, message: 'CRITICAL_FAIL: ' + error.toString() };
  }
}

/**
 * 1. DASHBOARD METRICS
 */
function getMetricsData() {
  try {
    const ss = getSpreadsheet();
    const sheet = getMdbSheet(ss); 
    if (!sheet) return { data: [], dropdowns: { healthCenters: [], morbWeeks: [] }, baselines: {}, syncTime: getFormattedTime() };

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow <= 1 || lastCol === 0) {
      return { data: [], dropdowns: { healthCenters: [], morbWeeks: [] }, baselines: {}, syncTime: getFormattedTime() };
    }
    
    const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = values[0].map(h => String(h).trim().toUpperCase());
    
    let idxCenter = headers.indexOf("HEALTH CENTER DESIGNATION"); 
    if (idxCenter === -1) idxCenter = headers.findIndex(h => h.includes("HEALTH CENTER") || h.includes("CENTER"));

    let idxRemarks = headers.indexOf("INVESTIGATION REMARKS"); 
    if (idxRemarks === -1) idxRemarks = headers.findIndex(h => h.includes("REMARKS"));

    let idxBarangay = headers.indexOf("BARANGAY");                 
    if (idxBarangay === -1) idxBarangay = headers.findIndex(h => h.includes("BRGY") || h.includes("BGY"));

    let idxOutcome = headers.indexOf("OUTCOME"); 
    if (idxOutcome === -1) idxOutcome = headers.findIndex(h => h.includes("OUTCOME"));

    let idxDisease = headers.indexOf("DISEASE_NAME");
    if (idxDisease === -1) idxDisease = headers.findIndex(h => h.includes("DISEASE"));
    
    let idxSex = headers.indexOf("GENDER");
    if (idxSex === -1) idxSex = headers.indexOf("SEX");
    
    let idxAge = headers.indexOf("AGE_IN_YEARS");
    if (idxAge === -1) idxAge = headers.indexOf("AGE IN YEARS");
    
    let idxAgeGroup = headers.indexOf("AGE_GROUP");
    if (idxAgeGroup === -1) idxAgeGroup = headers.findIndex(h => h.includes("AGE_GROUP") || h.includes("AGE GROUP"));
    
    let idxMorbWeek = headers.indexOf("MORBIDITY_WEEK");           
    if (idxMorbWeek === -1) idxMorbWeek = headers.indexOf("MORBIDITY WEEK");
    
    let idxCoords = headers.indexOf("GEO_MAPPING (CORDINATES THRU G-MAPS)");
    if (idxCoords === -1) idxCoords = headers.findIndex(h => h.includes("GEO_MAPPING") || h.includes("CORDINATES") || h.includes("COORDINATES"));

    let idxClass = headers.indexOf("LABORATORY_STATUS");
    if (idxClass === -1) {
      idxClass = headers.findIndex(h => h.includes("CLASSIFICATION") || h.includes("LABORATORY") || h.includes("LAB_STATUS") || h.includes("STATUS"));
    }

    let idxCategory = headers.indexOf("CATEGORY");
    if (idxCategory === -1) idxCategory = headers.findIndex(h => h.includes("CATEGORY"));
    
    const dashboardRows = [];
    const dynamicCenters = new Set();
    const dynamicWeeks = new Set();
    
    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      if (!row || row.join("").trim() === "") continue;

      const healthCenter = idxCenter !== -1 ? String(row[idxCenter]).trim() : "";
      const investigationRemarks = idxRemarks !== -1 ? String(row[idxRemarks]).trim() : "";
      const barangay = idxBarangay !== -1 ? String(row[idxBarangay]).trim() : "";
      const outcome = idxOutcome !== -1 ? String(row[idxOutcome]).trim() : "";
      const diseaseName = idxDisease !== -1 ? String(row[idxDisease]).trim().toUpperCase() : "";
      const rawClassValue = idxClass !== -1 ? String(row[idxClass]).trim().toUpperCase() : "";
      const rawRemarksValue = investigationRemarks.toUpperCase();

      if (!barangay && !diseaseName && !healthCenter) continue;
      if (!healthCenter || healthCenter === "") continue;
      
      if (
        rawRemarksValue.includes("DELIST") || rawRemarksValue.includes("NON-RES") || 
        rawRemarksValue.includes("NON RES") || rawRemarksValue.includes("DISCARD") || 
        rawRemarksValue.includes("DUPLICATE") || rawClassValue.includes("DISCARD") || 
        rawClassValue.includes("DUPLICATE") || rawClassValue.includes("DELIST")
      ) {
        continue; 
      }
      
      let morbWeek = null;
      if (idxMorbWeek !== -1 && row[idxMorbWeek] !== "") {
        let rawValue = String(row[idxMorbWeek]).trim();
        if (rawValue !== "") {
          let parsedWeek = parseInt(rawValue, 10);
          if (!isNaN(parsedWeek)) morbWeek = parsedWeek;
        }
      }
      
      let statusClass = "Pending"; 
      if (idxClass !== -1 && row[idxClass]) {
        let rawClass = String(row[idxClass]).trim().toUpperCase();
        if (rawClass.includes("CONFIRM") || rawClass.includes("POSITIVE") || rawClass === "POS" || rawClass === "C" || rawClass === "+") {
          statusClass = "Confirmed"; 
        } else if (rawClass.includes("NEGATIVE") || rawClass === "NEG" || rawClass === "N" || rawClass === "-") {
          statusClass = "Negative";
        } else if (rawClass.includes("SUSPECT")) {
          statusClass = "Suspect";
        } else if (rawClass.includes("PROBABLE")) {
          statusClass = "Probable";
        }
      }
      
      let lat = null, lng = null;
      if (idxCoords !== -1 && row[idxCoords]) {
        let coordString = String(row[idxCoords]).trim();
        let parts = coordString.split(/[\s,]+/);
        if (parts.length >= 2) {
          let p1 = parseFloat(parts[0]);
          let p2 = parseFloat(parts[1]);
          if (!isNaN(p1) && !isNaN(p2)) {
            if (p1 > p2) { lat = p2; lng = p1; } else { lat = p1; lng = p2; }
          }
        }
      }
      
      const rawSex = idxSex !== -1 ? String(row[idxSex]).trim().toUpperCase() : "";
      let cleanSex = "UNKNOWN";
      if (rawSex.startsWith("MALE") || rawSex === "M") cleanSex = "MALE";
      else if (rawSex.startsWith("FEMALE") || rawSex === "F") cleanSex = "FEMALE";
      
      let cleanAge = null;
      if (idxAge !== -1 && row[idxAge] !== "" && row[idxAge] !== undefined) {
        let parsed = parseFloat(row[idxAge]);
        if (!isNaN(parsed)) cleanAge = parsed;
      }
      
      let ageGroup = (idxAgeGroup !== -1 && row[idxAgeGroup]) ? String(row[idxAgeGroup]).trim() : "";
      let category = (idxCategory !== -1 && row[idxCategory]) ? String(row[idxCategory]).trim().toUpperCase() : "";
      
      let district = "Other Districts";
      if (["COMMONWEALTH", "BATASAN HILLS", "HOLY SPIRIT", "PAYATAS", "BAGONG SILANGAN"].includes(barangay.toUpperCase())) {
        district = "District 2";
      }

      dashboardRows.push({
        district: district,
        barangay: barangay,
        healthCenter: healthCenter,
        investigationRemarks: investigationRemarks,
        outcome: outcome,
        morbWeek: morbWeek,
        diseaseName: diseaseName,
        sex: cleanSex,
        age: cleanAge,
        ageGroup: ageGroup,
        category: category,
        lat: lat,
        lng: lng,
        classification: statusClass
      });
      
      if (healthCenter) dynamicCenters.add(healthCenter);
      if (morbWeek !== null && morbWeek !== undefined) dynamicWeeks.add(morbWeek);
    }
    
    const sortedWeeks = Array.from(dynamicWeeks).sort((a, b) => a - b);
    
    const baselines = {
      DENGUE: fetchThresholdArrays(ss, "DENGUE BASELINE"),
      MEASLES: fetchThresholdArrays(ss, "MEASLES BASELINE"),
      LEPTO: fetchThresholdArrays(ss, "LEPTO BASELINE"),
      LEPTOSPIROSIS: fetchThresholdArrays(ss, "LEPTO BASELINE"),
      COVID: fetchThresholdArrays(ss, "COVID BASELINE"),
      "COVID-19": fetchThresholdArrays(ss, "COVID BASELINE"),
      COVID19: fetchThresholdArrays(ss, "COVID BASELINE"),
      "SARS-COV-2": fetchThresholdArrays(ss, "COVID BASELINE"),
      CORONAVIRUS: fetchThresholdArrays(ss, "COVID BASELINE")
    };
    
    return {
      data: dashboardRows,
      dropdowns: {
        healthCenters: Array.from(dynamicCenters).sort(),
        morbWeeks: sortedWeeks
      },
      baselines: baselines,
      syncTime: getFormattedTime()
    };
  } catch (err) {
    Logger.log("Error in getMetricsData: " + err.toString());
    return { data: [], dropdowns: { healthCenters: [], morbWeeks: [] }, baselines: {}, syncTime: getFormattedTime(), error: err.toString() };
  }
}

/**
 * 2. FEEDBACK MONITORING
 */
function getDashboardData() {
  try {
    const ss = getSpreadsheet();
    const sheet = getEpisenseSheet(ss);
    if (!sheet) return { cases: [] };

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol === 0) return { cases: [] };

    const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = values[0].map(h => String(h).trim().toUpperCase());

    let idxDisease = headers.indexOf("DISEASE_TYPE");
    if (idxDisease === -1) idxDisease = headers.findIndex(h => h.includes("DISEASE"));

    let idxDate = headers.indexOf("DATE_VALIDATED");
    if (idxDate === -1) idxDate = headers.findIndex(h => h.includes("VALIDAT") || h.includes("DATE"));

    let idxCaseId = headers.indexOf("CASE_ID");
    if (idxCaseId === -1) idxCaseId = headers.findIndex(h => h.includes("CASE") || h.includes("ID"));

    let idxHC = headers.indexOf("HEALTH_CENTER");
    if (idxHC === -1) idxHC = headers.findIndex(h => h.includes("HEALTH") || h.includes("CENTER"));

    let idxInvStatus = headers.indexOf("INV_STATUS");
    if (idxInvStatus === -1) idxInvStatus = headers.findIndex(h => h.includes("STATUS") || h.includes("REMARK"));

    let idxSurveillance = headers.indexOf("OVER_2DAY_SURVEILLANCE");
    if (idxSurveillance === -1) idxSurveillance = headers.findIndex(h => h.includes("SURVEILLANCE") || h.includes("2DAY") || h.includes("OVER"));

    const now = new Date();
    const nowMs = now.getTime();
    const cases = [];

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      if (!row || row.join("").trim() === "") continue;

      const caseId = (idxCaseId !== -1 && row[idxCaseId]) ? String(row[idxCaseId]).trim() : `CASE-${r}`;
      const disease = (idxDisease !== -1 && row[idxDisease]) ? String(row[idxDisease]).trim().toUpperCase() : "DENGUE";
      let healthCenter = (idxHC !== -1 && row[idxHC]) ? String(row[idxHC]).trim() : "Unassigned Health Center";
      const invStatus = (idxInvStatus !== -1 && row[idxInvStatus]) ? String(row[idxInvStatus]).trim().toUpperCase() : "";
      const surveillanceStr = (idxSurveillance !== -1 && row[idxSurveillance]) ? String(row[idxSurveillance]).trim().toUpperCase() : "";

      if (invStatus.includes("DELIST") || invStatus.includes("DISCARD") || invStatus.includes("DUPLICATE")) continue;

      let rawDate = (idxDate !== -1) ? row[idxDate] : null;
      let dateObj = (rawDate instanceof Date) ? rawDate : (rawDate ? new Date(String(rawDate).trim()) : now);
      if (!dateObj || isNaN(dateObj.getTime())) dateObj = now;

      let hoursElapsed = (nowMs - dateObj.getTime()) / (1000 * 60 * 60);
      let category = "Under 24 Hours";
      if (surveillanceStr.includes("OVER 48") || surveillanceStr.includes("OVER 2") || surveillanceStr.includes(">48")) {
        category = "Over 48 Hours";
        if (hoursElapsed < 48) hoursElapsed = 49.0;
      } else if (surveillanceStr.includes("24 TO 48") || surveillanceStr.includes("24-48")) {
        category = "24 to 48 Hours";
        if (hoursElapsed < 24 || hoursElapsed > 48) hoursElapsed = 36.0;
      } else if (hoursElapsed > 48) {
        category = "Over 48 Hours";
      } else if (hoursElapsed >= 24) {
        category = "24 to 48 Hours";
      } else {
        category = "Under 24 Hours";
        if (hoursElapsed < 0) hoursElapsed = 12.0;
      }

      const yyyy = dateObj.getFullYear();
      const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
      const dd = String(dateObj.getDate()).padStart(2, '0');

      cases.push({
        caseId: caseId,
        disease: disease,
        healthCenter: healthCenter,
        dateValidated: `${yyyy}-${mm}-${dd}`,
        hoursElapsed: Number(hoursElapsed).toFixed(1),
        category: category
      });
    }

    return { cases: cases };
  } catch (err) {
    Logger.log("Error in getDashboardData: " + err.toString());
    return { cases: [], error: err.toString() };
  }
}

/**
 * 3. D2 NAVDPCP 2026 REPORT
 */
function getNavdpcpData(filters) {
  try {
    const ss = getSpreadsheet();
    const sheet = getMdbSheet(ss); 
    if (!sheet) return { table1: { rows: [], totals: {} }, table21: { testTypes: [], rows: [], totals: {} }, kpi22: 0, table34: { rows: [], totals: {} }, table4Star: { rows: [], totals: {} } };

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol === 0) {
      return { table1: { rows: [], totals: {} }, table21: { testTypes: [], rows: [], totals: {} }, kpi22: 0, table34: { rows: [], totals: {} }, table4Star: { rows: [], totals: {} } };
    }

    const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = values[0].map(h => String(h).trim().toUpperCase());
    
    let idxHC = headers.indexOf("HEALTH CENTER DESIGNATION");
    if (idxHC === -1) idxHC = headers.findIndex(h => h.includes("HEALTH CENTER") || h.includes("CENTER"));

    let idxBrgy = headers.indexOf("BARANGAY");
    if (idxBrgy === -1) idxBrgy = headers.findIndex(h => h.includes("BRGY") || h.includes("BGY"));

    let idxRemarks = headers.indexOf("INVESTIGATION REMARKS");
    if (idxRemarks === -1) idxRemarks = headers.findIndex(h => h.includes("REMARKS"));

    let idxDisease = headers.indexOf("DISEASE_NAME");
    if (idxDisease === -1) idxDisease = headers.findIndex(h => h.includes("DISEASE"));

    let idxClass = headers.indexOf("LABORATORY_STATUS");
    if (idxClass === -1) idxClass = headers.findIndex(h => h.includes("CLASSIFICATION") || h.includes("LABORATORY") || h.includes("STATUS"));

    let idxTesting = headers.indexOf("TYPE OF TESTING");
    if (idxTesting === -1) idxTesting = headers.findIndex(h => h.includes("TESTING") || h.includes("TEST"));

    let idxClinical = headers.indexOf("CLINICAL_CLASSIFICATION");
    if (idxClinical === -1) idxClinical = headers.findIndex(h => h.includes("CLINICAL"));

    let idxOutcome = headers.indexOf("OUTCOME");
    if (idxOutcome === -1) idxOutcome = headers.findIndex(h => h.includes("OUTCOME"));

    let idxCategory = headers.indexOf("CATEGORY");
    if (idxCategory === -1) idxCategory = headers.findIndex(h => h.includes("CATEGORY"));

    let idxDate = headers.indexOf("CESU DATE ADDED");
    if (idxDate === -1) idxDate = headers.findIndex(h => h.includes("CESU") || h.includes("DATE"));

    const targetBrgy = filters && filters.barangay ? filters.barangay.toUpperCase().trim() : "ALL BARANGAY";
    const targetHC = filters && filters.healthCenter ? filters.healthCenter.toUpperCase().trim() : "ALL HEALTH CENTERS";
    const targetRemarks = filters && filters.remarks ? filters.remarks.toUpperCase().trim() : "ALL REMARKS";
    const targetDisease = filters && filters.disease ? filters.disease.toUpperCase().trim() : "DENGUE";

    const startDateStr = filters && filters.startDate ? filters.startDate.trim() : "";
    const endDateStr = filters && filters.endDate ? filters.endDate.trim() : "";

    let startMs = startDateStr ? new Date(startDateStr).getTime() : 0;
    let endMs = endDateStr ? new Date(endDateStr + "T23:59:59").getTime() : Infinity;

    const caseFindingMap = {};
    const testingMap = {};
    let confirmatoryCount = 0;
    const clinicalOutcomeMap = {};
    const cesdMap = {};

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      if (!row || row.join("").trim() === "") continue;
      
      const hc = idxHC !== -1 && row[idxHC] ? String(row[idxHC]).trim() : "null";
      const brgy = idxBrgy !== -1 ? String(row[idxBrgy]).trim().toUpperCase() : "";
      const remarks = idxRemarks !== -1 ? String(row[idxRemarks]).trim().toUpperCase() : "";
      const disease = idxDisease !== -1 ? String(row[idxDisease]).trim().toUpperCase() : "";
      const rawClass = idxClass !== -1 ? String(row[idxClass]).trim().toUpperCase() : "";

      if (remarks.includes("DELIST") || remarks.includes("NON-RES") || remarks.includes("DISCARD") || remarks.includes("DUPLICATE")) continue;

      if (targetBrgy !== "ALL BARANGAY" && !brgy.includes(targetBrgy)) continue;
      if (targetHC !== "ALL HEALTH CENTERS" && hc.toUpperCase() !== targetHC) continue;
      if (targetRemarks !== "ALL REMARKS" && !remarks.includes(targetRemarks)) continue;
      if (targetDisease !== "ALL DISEASES" && !disease.includes(targetDisease)) continue;

      if (startDateStr || endDateStr) {
        let rawD = idxDate !== -1 ? row[idxDate] : null;
        let dObj = rawD instanceof Date ? rawD : new Date(rawD);
        let rowTime = dObj.getTime();
        if (isNaN(rowTime) || rowTime < startMs || rowTime > endMs) continue;
      }

      let classification = "null";
      if (rawClass.includes("SUSPECT")) classification = "SUSPECT";
      else if (rawClass.includes("CONFIRM") || rawClass.includes("POS")) classification = "CONFIRM";
      else if (rawClass.includes("PROBABLE")) classification = "PROBABLE";

      if (!caseFindingMap[hc]) caseFindingMap[hc] = { SUSPECT: 0, CONFIRM: 0, PROBABLE: 0, null: 0 };
      caseFindingMap[hc][classification] = (caseFindingMap[hc][classification] || 0) + 1;

      let testType = idxTesting !== -1 && row[idxTesting] ? String(row[idxTesting]).trim() : "null";
      if (!testType || testType === "") testType = "null";
      if (!testingMap[hc]) testingMap[hc] = {};
      testingMap[hc][testType] = (testingMap[hc][testType] || 0) + 1;

      if (testType.includes("PCR") || testType.includes("RTPCR") || testType.includes("POLYMERASE") || testType.includes("SERUM") || testType.includes("SWAB")) {
        confirmatoryCount++;
      }

      let clinical = idxClinical !== -1 && row[idxClinical] ? String(row[idxClinical]).trim().toUpperCase() : "null";
      let outcome = idxOutcome !== -1 && row[idxOutcome] ? String(row[idxOutcome]).trim().toUpperCase() : "null";
      if (outcome.includes("ALIVE")) outcome = "ALIVE";
      else if (outcome.includes("DIE") || outcome.includes("DEATH")) outcome = "DIED";
      else outcome = "null";

      const clinKey = `${hc}||${clinical}`;
      if (!clinicalOutcomeMap[clinKey]) clinicalOutcomeMap[clinKey] = { hc: hc, clinical: clinical, ALIVE: 0, DIED: 0, null: 0 };
      clinicalOutcomeMap[clinKey][outcome]++;

      let category = idxCategory !== -1 && row[idxCategory] ? String(row[idxCategory]).trim().toUpperCase() : "HC DETECTED";
      if (category.includes("PHSU")) category = "PHSU ENDORSEMENT";
      else category = "HC DETECTED";

      if (!cesdMap[hc]) cesdMap[hc] = { "PHSU ENDORSEMENT": 0, "HC DETECTED": 0 };
      cesdMap[hc][category]++;
    }

    return {
      table1: formatTable1(caseFindingMap),
      table21: formatTable21(testingMap),
      kpi22: confirmatoryCount,
      table34: formatTable34(clinicalOutcomeMap),
      table4Star: formatTable4Star(cesdMap)
    };
  } catch(err) {
    Logger.log("Error in getNavdpcpData: " + err.toString());
    return { table1: { rows: [], totals: {} }, table21: { testTypes: [], rows: [], totals: {} }, kpi22: 0, table34: { rows: [], totals: {} }, table4Star: { rows: [], totals: {} } };
  }
}

/**
 * 4. PATIENT PROFILE RECORDS FETCH
 */
function getRecords() {
  try {
    const ss = getSpreadsheet();
    const sheet = getMdbSheet(ss); 
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol === 0) return [];

    const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = values[0].map(h => String(h).trim());
    const records = [];

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      if (!row || row.join("").trim() === "") continue;

      const recordObj = { _rowIndex: r + 1 };
      headers.forEach((h, idx) => {
        let val = row[idx];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        recordObj[h] = (val !== undefined && val !== null) ? String(val).trim() : "";
      });
      records.push(recordObj);
    }
    return records;
  } catch(err) {
    Logger.log("Error in getRecords: " + err.toString());
    return [];
  }
}

/**
 * 5. SAVE / UPDATE PATIENT RECORD
 */
function saveRecord(formData, rowIndex) {
  try {
    const ss = getSpreadsheet();
    const sheet = getMdbSheet(ss);
    if (!sheet) return { success: false, message: "MDB Sheet not found." };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim().toUpperCase());

    let targetRow = rowIndex;
    if (!targetRow || isNaN(targetRow) || targetRow <= 1) {
      targetRow = sheet.getLastRow() + 1;
    }

    const rowValues = [];
    headers.forEach(header => {
      let val = "";
      for (let key in formData) {
        if (key.trim().toUpperCase() === header) {
          val = formData[key];
          break;
        }
      }
      rowValues.push(val);
    });

    sheet.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
    return { success: true, message: `Record successfully ${rowIndex ? 'updated' : 'saved'}!` };
  } catch (err) {
    return { success: false, message: `Failed to save: ${err.message}` };
  }
}

// Helpers
function formatTable1(map) {
  let list = [];
  let totals = { SUSPECT: 0, CONFIRM: 0, null: 0, PROBABLE: 0, grandTotal: 0 };
  for (let hc in map) {
    let row = map[hc];
    let sum = (row.SUSPECT || 0) + (row.CONFIRM || 0) + (row.null || 0) + (row.PROBABLE || 0);
    list.push({ hc: hc, suspect: row.SUSPECT || 0, confirm: row.CONFIRM || 0, nullVal: row.null || 0, probable: row.PROBABLE || 0, grandTotal: sum });
    totals.SUSPECT += (row.SUSPECT || 0);
    totals.CONFIRM += (row.CONFIRM || 0);
    totals.null += (row.null || 0);
    totals.PROBABLE += (row.PROBABLE || 0);
    totals.grandTotal += sum;
  }
  return { rows: list, totals: totals };
}

function formatTable21(map) {
  let testTypesSet = new Set();
  let list = [];
  for (let hc in map) {
    for (let tt in map[hc]) testTypesSet.add(tt);
  }
  let testTypes = Array.from(testTypesSet).sort();
  let totals = { grandTotal: 0 };
  testTypes.forEach(tt => totals[tt] = 0);

  for (let hc in map) {
    let rowObj = { hc: hc, grandTotal: 0 };
    testTypes.forEach(tt => {
      let val = map[hc][tt] || 0;
      rowObj[tt] = val;
      rowObj.grandTotal += val;
      totals[tt] += val;
      totals.grandTotal += val;
    });
    list.push(rowObj);
  }
  return { testTypes: testTypes, rows: list, totals: totals };
}

function formatTable34(map) {
  let list = [];
  let totals = { ALIVE: 0, nullVal: 0, DIED: 0, grandTotal: 0 };
  for (let key in map) {
    let item = map[key];
    let sum = item.ALIVE + item.null + item.DIED;
    list.push({ hc: item.hc, clinical: item.clinical, alive: item.ALIVE, nullVal: item.null, died: item.DIED, grandTotal: sum });
    totals.ALIVE += item.ALIVE;
    totals.nullVal += item.null;
    totals.DIED += item.DIED;
    totals.grandTotal += sum;
  }
  return { rows: list, totals: totals };
}

function formatTable4Star(map) {
  let list = [];
  let totals = { phsu: 0, hcDetected: 0, grandTotal: 0 };
  for (let hc in map) {
    let phsu = map[hc]["PHSU ENDORSEMENT"] || 0;
    let hcDet = map[hc]["HC DETECTED"] || 0;
    let sum = phsu + hcDet;
    list.push({ hc: hc, phsu: phsu, hcDetected: hcDet, grandTotal: sum });
    totals.phsu += phsu;
    totals.hcDetected += hcDet;
    totals.grandTotal += sum;
  }
  return { rows: list, totals: totals };
}

function fetchThresholdArrays(ss, sheetName) {
  const targetSheet = ss.getSheetByName(sheetName);
  const alertCurve = new Array(53).fill(0);
  const epidemicCurve = new Array(53).fill(0);
  const currentYearCurve = new Array(53).fill(0);
  
  if (!targetSheet) return { alert: alertCurve, epidemic: epidemicCurve, currentYear: currentYearCurve, hasNegOne: false };
  
  const dataset = targetSheet.getDataRange().getValues();
  let hasNegOne = false;
  for (let i = 1; i < dataset.length; i++) {
    const row = dataset[i];
    const wNum = parseInt(row[0], 10);
    if (!isNaN(wNum)) {
      const epidemicThreshold = Number(row[1] || 0);
      const alertThreshold = Number(row[2] || 0);
      const currentYearValue = Number(row[3] || 0);

      if (wNum === -1) {
        hasNegOne = true;
        alertCurve[0] = alertThreshold;
        epidemicCurve[0] = epidemicThreshold;
        currentYearCurve[0] = currentYearValue;
      } else if (wNum >= 1 && wNum <= 52) {
        alertCurve[wNum] = alertThreshold;
        epidemicCurve[wNum] = epidemicThreshold;
        currentYearCurve[wNum] = currentYearValue;
      }
    }
  }
  return { alert: alertCurve, epidemic: epidemicCurve, currentYear: currentYearCurve, hasNegOne: hasNegOne };
}

function getFormattedTime() {
  const now = new Date();
  return Utilities.formatDate(now, Session.getScriptTimeZone(), "hh:mm a");
}