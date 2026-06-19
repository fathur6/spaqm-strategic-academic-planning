/**
 * ====================================================================
 * UNISZA STRATEGIC PORTAL BACKEND - CODE.JS (SECURE VERSION)
 * ====================================================================
 * ID Spreadsheet disimpan dengan selamat di dalam Google Script Properties.
 * Sila rujuk nota di bawah untuk cara memasukkan kunci ID di Google Apps Script.
 */

// Menarik ID Spreadsheet secara rahsia daripada persekitaran pelayan Google (Environment Variable)
const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');

/**
 * Mendapatkan Spreadsheet secara selamat tanpa mendedahkan ID pada kod sumber
 */
function getSpreadsheet() {
  if (SPREADSHEET_ID && SPREADSHEET_ID !== "") {
    try {
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    } catch (e) {
      Logger.log("Ralat openById: " + e.toString());
    }
  }
  // Fallback sekiranya diakses sebagai skrip aktif kontena
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * MENAMPILKAN PORTAL (doGet)
 * Memuatkan fail index.html secara selamat
 */
function doGet(e) {
  if (!SPREADSHEET_ID) {
    return ContentService.createTextOutput(
      "⚠️ RALAT AMAN: SPREADSHEET_ID belum ditetapkan di dalam Script Properties UniSZA.\n" +
      "Sila ke Project Settings (ikon gear) di Google Apps Script untuk menetapkannya."
    );
  }
  
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle("UniSZA Strategic Command Portal")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Mendapatkan Alamat Emel Pengguna Aktif UniSZA untuk Log Audit
 */
function getActiveUserEmail() {
  try {
    const email = Session.getActiveUser().getEmail();
    return email && email !== "" ? email : "Pengguna UniSZA Workspace";
  } catch (e) {
    return "Sesi Luar Talian";
  }
}

/**
 * CRUD - READ DATA
 */
function getInitiatives() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName("QS_Strategic_Plan");
    if (!sheet) return { success: false, error: "Helaian QS_Strategic_Plan tidak ditemui." };
    
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    
    const list = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const obj = {};
      headers.forEach((header, index) => {
        obj[header.toLowerCase()] = row[index];
      });
      list.push(obj);
    }
    
    return { 
      success: true, 
      data: list, 
      userEmail: getActiveUserEmail() 
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * CRUD - CREATE RECORD
 */
function createInitiative(item) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName("QS_Strategic_Plan");
    const values = sheet.getDataRange().getValues();
    
    let maxId = 0;
    for (let i = 1; i < values.length; i++) {
      const idVal = parseInt(values[i][0]);
      if (idVal > maxId) maxId = idVal;
    }
    const newId = maxId + 1;
    
    const newRow = [
      newId,
      item.program || "",
      item.objective || "",
      item.quarter || "Q1",
      item.kpi || "",
      Number(item.target) || 10,
      Number(item.current) || 0,
      item.status || "Not Started",
      item.category || "GAVP",
      item.weight || "",
      item.icon || "Layers"
    ];
    
    sheet.appendRow(newRow);
    logActivity("CREATE", "Inisiatif ID: " + newId + " (" + (item.program || "") + ") dicipta.");
    
    return { success: true, id: newId };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * CRUD - UPDATE RECORD
 */
function updateInitiative(item) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName("QS_Strategic_Plan");
    const values = sheet.getDataRange().getValues();
    const targetId = parseInt(item.id);
    let rowIndex = -1;
    
    for (let i = 1; i < values.length; i++) {
      if (parseInt(values[i][0]) === targetId) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) return { success: false, error: "ID tidak dijumpai." };
    
    sheet.getRange(rowIndex, 2).setValue(item.program);
    sheet.getRange(rowIndex, 3).setValue(item.objective);
    sheet.getRange(rowIndex, 4).setValue(item.quarter);
    sheet.getRange(rowIndex, 5).setValue(item.kpi);
    sheet.getRange(rowIndex, 6).setValue(Number(item.target));
    sheet.getRange(rowIndex, 7).setValue(Number(item.current));
    sheet.getRange(rowIndex, 8).setValue(item.status);
    sheet.getRange(rowIndex, 9).setValue(item.category);
    sheet.getRange(rowIndex, 10).setValue(item.weight);
    sheet.getRange(rowIndex, 11).setValue(item.icon);
    
    logActivity("UPDATE", "Inisiatif ID: " + targetId + " dikemaskini.");
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * CRUD - DELETE RECORD
 */
function deleteInitiative(id) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName("QS_Strategic_Plan");
    const values = sheet.getDataRange().getValues();
    const targetId = parseInt(id);
    let rowIndex = -1;
    
    for (let i = 1; i < values.length; i++) {
      if (parseInt(values[i][0]) === targetId) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) return { success: false, error: "ID tidak dijumpai." };
    
    sheet.deleteRow(rowIndex);
    logActivity("DELETE", "Inisiatif ID: " + targetId + " dipadam.");
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * REKOD LOG AUDIT AKTIVITI
 */
function logActivity(action, details) {
  try {
    const ss = getSpreadsheet();
    let logSheet = ss.getSheetByName("Audit_Logs");
    if (!logSheet) {
      logSheet = ss.insertSheet("Audit_Logs");
      logSheet.appendRow(["Timestamp", "Pengguna (Emel)", "Aksi", "Perincian Aktiviti"]);
    }
    logSheet.appendRow([new Date(), getActiveUserEmail(), action, details]);
  } catch (e) {}
}