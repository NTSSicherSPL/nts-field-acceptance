/**
 * NTS Field Acceptance — Google Apps Script Backend
 * ═══════════════════════════════════════════════════
 * GATA CONFIGURAT — nu trebuie modificat nimic.
 *
 * PAȘI DEPLOY:
 * 1. Lipește acest cod în editorul script.google.com
 * 2. Salvează (Ctrl+S)
 * 3. Run > setupDatabase  (o singură dată, pentru a crea spreadsheet-ul)
 * 4. Deploy > Manage Deployments > editeaza deployment-ul existent
 *    SAU Deploy > New Deployment > Web App
 *    - Execute as: Me
 *    - Who has access: Anyone
 */

// ═══════════════════════════════════════════════
//  CONFIGURARE HARDCODATĂ
// ═══════════════════════════════════════════════
var DRIVE_FOLDER_ID  = '1ug6huCS_OZEbKzaO5d-WQ1hWC-XBFHMy';
var SPREADSHEET_NAME = 'NTS Field Acceptance — Database';

var SHEET_USERS      = 'Users';
var SHEET_CATEGORIES = 'Categories';
var SHEET_PROJECTS   = 'Projects';
var SHEET_NA_LOG     = 'NA_Log';

// ═══════════════════════════════════════════════
//  OBȚINE SAU CREEAZĂ SPREADSHEET-UL
// ═══════════════════════════════════════════════
function getSpreadsheet() {
  var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  var files  = folder.getFilesByName(SPREADSHEET_NAME);
  if (files.hasNext()) {
    return SpreadsheetApp.openById(files.next().getId());
  }
  // Creează spreadsheet nou și inițializează
  var ss     = SpreadsheetApp.create(SPREADSHEET_NAME);
  var ssFile = DriveApp.getFileById(ss.getId());
  folder.addFile(ssFile);
  try { DriveApp.getRootFolder().removeFile(ssFile); } catch(e) {}
  initSpreadsheet(ss);
  return ss;
}

// ═══════════════════════════════════════════════
//  INIȚIALIZARE SPREADSHEET
// ═══════════════════════════════════════════════
function initSpreadsheet(ss) {

  // ── USERS ──────────────────────────────────
  var usersSheet = ss.getSheetByName(SHEET_USERS);
  if (!usersSheet) usersSheet = ss.insertSheet(SHEET_USERS);
  else usersSheet.clearContents();
  usersSheet.appendRow(['Name', 'Email', 'Status', 'Date Requested']);
  styleHeader(usersSheet, 4);
  usersSheet.setColumnWidth(1, 200);
  usersSheet.setColumnWidth(2, 250);
  usersSheet.setColumnWidth(3, 130);
  usersSheet.setColumnWidth(4, 200);
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Approved', 'Rejected'], true).build();
  usersSheet.getRange('C2:C1000').setDataValidation(statusRule);

  // ── CATEGORIES ─────────────────────────────
  var catSheet = ss.getSheetByName(SHEET_CATEGORIES);
  if (!catSheet) catSheet = ss.insertSheet(SHEET_CATEGORIES);
  else catSheet.clearContents();
  catSheet.appendRow(['Category', 'Subcategory', 'Required Photos', 'Active']);
  styleHeader(catSheet, 4);
  catSheet.setColumnWidth(1, 220);
  catSheet.setColumnWidth(2, 220);
  catSheet.setColumnWidth(3, 150);
  catSheet.setColumnWidth(4, 100);
  var initCats = [
    ['Overview Location', 'Key Safe',      5, 'TRUE'],
    ['Overview Location', 'Access Gate',   3, 'TRUE'],
    ['Overview Location', 'Site Entrance', 2, 'TRUE'],
    ['Equipment',         'Cabinet',       4, 'TRUE'],
    ['Equipment',         'Power Supply',  3, 'TRUE'],
    ['Equipment',         'Labels',        2, 'TRUE'],
    ['Final Status',      'Site Overview', 5, 'TRUE'],
  ];
  initCats.forEach(function(row) { catSheet.appendRow(row); });
  var activeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'], true).build();
  catSheet.getRange('D2:D1000').setDataValidation(activeRule);

  // ── PROJECTS ───────────────────────────────
  var projSheet = ss.getSheetByName(SHEET_PROJECTS);
  if (!projSheet) projSheet = ss.insertSheet(SHEET_PROJECTS);
  else projSheet.clearContents();
  projSheet.appendRow(['Site Number','User Name','User Email','Date','Total Photos','N/A Items','Status','N/A Details']);
  styleHeader(projSheet, 8);
  [150,180,220,200,130,100,120,400].forEach(function(w,i){ projSheet.setColumnWidth(i+1,w); });

  // ── NA_LOG ─────────────────────────────────
  var naSheet = ss.getSheetByName(SHEET_NA_LOG);
  if (!naSheet) naSheet = ss.insertSheet(SHEET_NA_LOG);
  else naSheet.clearContents();
  naSheet.appendRow(['Site Number','Category','Subcategory','User Name','User Email','Date & Time','Status']);
  styleHeader(naSheet, 7);
  [150,200,200,180,220,200,100].forEach(function(w,i){ naSheet.setColumnWidth(i+1,w); });

  // Șterge Sheet1 implicit dacă există
  var def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) { try { ss.deleteSheet(def); } catch(e) {} }

  Logger.log('Spreadsheet initializat: ' + ss.getUrl());
}

function styleHeader(sheet, numCols) {
  var r = sheet.getRange(1, 1, 1, numCols);
  r.setFontWeight('bold');
  r.setBackground('#0a1628');
  r.setFontColor('#29AAE1');
  r.setFontSize(11);
  sheet.setFrozenRows(1);
}

// ═══════════════════════════════════════════════
//  HTTP HANDLERS
//  Notă: GAS Web App deployed cu "Anyone" nu
//  necesită headere CORS manuale — platforma
//  le adaugă automat. setHeader() nu este
//  disponibil pe ContentService.TextOutput.
// ═══════════════════════════════════════════════
function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleRequest(e) {
  try {
    var data = {};
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    var result = {};
    switch (data.action) {
      case 'login':         result = actionLogin(data);         break;
      case 'requestAccess': result = actionRequestAccess(data); break;
      case 'getCategories': result = actionGetCategories();     break;
      case 'uploadPhoto':   result = actionUploadPhoto(data);   break;
      case 'submitProject': result = actionSubmitProject(data); break;
      case 'logNA':         result = actionLogNA(data);         break;
      case 'getNALog':      result = actionGetNALog(data);      break;
      case 'setupDB':       result = actionSetupDB();           break;
      default:
        result = { error: 'Unknown action: ' + (data.action || 'none') };
    }
    return jsonOut(result);

  } catch (err) {
    return jsonOut({ error: err.message, stack: err.stack });
  }
}

// ═══════════════════════════════════════════════
//  SETUP DB — rulează din editor o singură dată
// ═══════════════════════════════════════════════
function setupDatabase() {
  var ss = getSpreadsheet();
  Logger.log('SUCCESS — Spreadsheet URL: ' + ss.getUrl());
  Logger.log('Spreadsheet ID: ' + ss.getId());
}

function actionSetupDB() {
  var ss = getSpreadsheet();
  return { success: true, spreadsheetId: ss.getId(), spreadsheetUrl: ss.getUrl() };
}

// ═══════════════════════════════════════════════
//  LOGIN
// ═══════════════════════════════════════════════
function actionLogin(data) {
  var email = (data.email || '').toLowerCase().trim();
  if (!email) return { status: 'error', message: 'Email required' };

  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) return { status: 'error', message: 'Users sheet not found' };

  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var rowEmail  = (rows[i][1] || '').toLowerCase().trim();
    var rowStatus = (rows[i][2] || '').trim().toLowerCase();
    var rowName   = (rows[i][0] || '').trim();
    if (rowEmail === email) {
      if (rowStatus === 'approved') return { status: 'approved', name: rowName, email: email };
      if (rowStatus === 'pending')  return { status: 'pending' };
      return { status: 'rejected' };
    }
  }
  return { status: 'not_found' };
}

// ═══════════════════════════════════════════════
//  REQUEST ACCESS
// ═══════════════════════════════════════════════
function actionRequestAccess(data) {
  var name  = (data.name  || '').trim();
  var email = (data.email || '').toLowerCase().trim();
  if (!name || !email) return { success: false, message: 'Name and email required' };

  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) return { success: false, message: 'Users sheet not found' };

  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if ((rows[i][1] || '').toLowerCase().trim() === email) {
      return { success: false, message: 'This email already has a request on file.' };
    }
  }
  sheet.appendRow([name, email, 'Pending', new Date().toISOString()]);
  return { success: true, message: 'Request submitted successfully.' };
}

// ═══════════════════════════════════════════════
//  GET CATEGORIES
// ═══════════════════════════════════════════════
function actionGetCategories() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_CATEGORIES);
  if (!sheet) return { categories: [] };

  var rows       = sheet.getDataRange().getValues();
  var categories = [];
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    if (!row[0] || !row[1]) continue;
    var active = String(row[3]).toUpperCase();
    if (active === 'FALSE') continue;
    categories.push({
      category:    String(row[0]).trim(),
      subcategory: String(row[1]).trim(),
      required:    Number(row[2]) || 1,
      active:      true,
    });
  }
  return { categories: categories };
}

// ═══════════════════════════════════════════════
//  UPLOAD PHOTO
// ═══════════════════════════════════════════════
function actionUploadPhoto(data) {
  var siteNumber  = (data.siteNumber  || '').trim();
  var category    = (data.category    || '').trim();
  var subcategory = (data.subcategory || '').trim();
  var fileName    = (data.fileName    || 'photo.jpg').trim();
  var base64Data  = data.base64Data   || '';
  var mimeType    = data.mimeType     || 'image/jpeg';

  if (!siteNumber || !base64Data) {
    return { success: false, message: 'Missing required fields' };
  }

  var rootFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  var siteFolder = getOrCreateFolder(rootFolder, 'Site_' + siteNumber);
  var catFolder  = getOrCreateFolder(siteFolder, category);
  var subFolder  = getOrCreateFolder(catFolder, subcategory);

  var existing = subFolder.getFilesByName(fileName);
  if (existing.hasNext()) {
    return { success: true, fileId: existing.next().getId(), message: 'Already exists' };
  }

  var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
  var file = subFolder.createFile(blob);
  return { success: true, fileId: file.getId(), fileName: fileName };
}

// ═══════════════════════════════════════════════
//  LOG N/A
// ═══════════════════════════════════════════════
function actionLogNA(data) {
  var siteNumber  = (data.siteNumber  || '').trim();
  var category    = (data.category    || '').trim();
  var subcategory = (data.subcategory || '').trim();
  var userName    = (data.userName    || '').trim();
  var userEmail   = (data.userEmail   || '').trim();
  var markedAt    = data.markedAt || new Date().toISOString();

  if (!siteNumber || !category || !subcategory) {
    return { success: false, message: 'Missing required fields' };
  }

  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NA_LOG);
  if (!sheet) return { success: false, message: 'NA_Log sheet not found' };

  sheet.appendRow(['Site_' + siteNumber, category, subcategory, userName, userEmail, markedAt, 'N/A']);
  return { success: true };
}

// ═══════════════════════════════════════════════
//  GET N/A LOG
// ═══════════════════════════════════════════════
function actionGetNALog(data) {
  var siteNumber = (data.siteNumber || '').trim();
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NA_LOG);
  if (!sheet) return { entries: [] };

  var rows    = sheet.getDataRange().getValues();
  var entries = [];
  var filter  = siteNumber ? ('Site_' + siteNumber) : null;

  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    if (!row[0]) continue;
    if (filter && row[0] !== filter) continue;
    entries.push({
      siteNumber: row[0], category: row[1], subcategory: row[2],
      userName: row[3], userEmail: row[4], markedAt: row[5], status: row[6],
    });
  }
  return { entries: entries };
}

// ═══════════════════════════════════════════════
//  SUBMIT PROJECT
// ═══════════════════════════════════════════════
function actionSubmitProject(data) {
  var siteNumber  = (data.siteNumber  || '').trim();
  var userName    = (data.userName    || '').trim();
  var userEmail   = (data.userEmail   || '').trim();
  var totalPhotos = Number(data.totalPhotos) || 0;
  var totalNA     = Number(data.totalNA)     || 0;
  var date        = data.date || new Date().toISOString();
  var naDetails   = data.naDetails || [];

  var ss        = getSpreadsheet();
  var projSheet = ss.getSheetByName(SHEET_PROJECTS);
  if (!projSheet) return { success: false, message: 'Projects sheet not found' };

  var naDetailsStr = naDetails.length
    ? naDetails.map(function(n){ return n.category + ' > ' + n.subcategory; }).join('; ')
    : 'None';

  projSheet.appendRow([
    'Site_' + siteNumber, userName, userEmail,
    date, totalPhotos, totalNA, 'Completed', naDetailsStr
  ]);

  // Flush N/A logs lipsă (scenariu offline)
  if (naDetails.length) {
    var naSheet = ss.getSheetByName(SHEET_NA_LOG);
    if (naSheet) {
      var existing    = naSheet.getDataRange().getValues();
      var existingSet = {};
      existing.slice(1).forEach(function(r){ existingSet[r[0]+'|'+r[1]+'|'+r[2]] = true; });
      naDetails.forEach(function(n) {
        var k = 'Site_' + siteNumber + '|' + n.category + '|' + n.subcategory;
        if (!existingSet[k]) {
          naSheet.appendRow(['Site_' + siteNumber, n.category, n.subcategory, userName, userEmail, date, 'N/A']);
        }
      });
    }
  }

  return { success: true, message: 'Project submitted successfully.' };
}

// ═══════════════════════════════════════════════
//  HELPER: Obține sau creează subfolder
// ═══════════════════════════════════════════════
function getOrCreateFolder(parent, name) {
  var folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(name);
}
