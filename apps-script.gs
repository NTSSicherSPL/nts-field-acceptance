/**
 * NTS Field Acceptance — Google Apps Script Backend v2.0
 * ═══════════════════════════════════════════════════════
 * PASI DEPLOY:
 * 1. Inlocuieste codul in script.google.com
 * 2. Salveaza (Ctrl+S)
 * 3. Run > upgradeSheets   (adauga coloanele noi la sheet-urile existente)
 * 4. Deploy > Manage Deployments > Edit > New Version > Deploy
 *
 * STRUCTURA SPREADSHEET:
 *   Users      — Name | Email | Status | Date Requested
 *   Projects   — Project Name | Description | Active
 *   Categories — Project | Category | Subcategory | Required Photos | Active | Instructions | Image URLs
 *   Submissions— Site Number | Project | User Name | User Email | Date | Total Photos | N/A Items | Status | N/A Details
 *   NA_Log     — Site Number | Project | Category | Subcategory | User Name | User Email | Date & Time | Status
 */

var SPREADSHEET_ID  = '1GRUPV9eAIKlv9P2JFpWEPUZ6iVZ394l8_nz9fN2yypc';
var DRIVE_FOLDER_ID = '1ug6huCS_OZEbKzaO5d-WQ1hWC-XBFHMy';

var SHEET_USERS       = 'Users';
var SHEET_PROJECTS    = 'Projects';
var SHEET_CATEGORIES  = 'Categories';
var SHEET_SUBMISSIONS = 'Submissions';
var SHEET_NA_LOG      = 'NA_Log';

// -----------------------------------------------
//  SPREADSHEET ACCESS
// -----------------------------------------------
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// -----------------------------------------------
//  UPGRADE SHEETS
//  Adauga coloane/sheet-uri noi fara a sterge date
//  Run > upgradeSheets din editorul GAS
// -----------------------------------------------
function upgradeSheets() {
  var ss = getSpreadsheet();

  // ── PROJECTS sheet (nou) ──
  var proj = ss.getSheetByName(SHEET_PROJECTS);
  if (!proj) {
    proj = ss.insertSheet(SHEET_PROJECTS);
    proj.appendRow(['Project Name', 'Description', 'Active']);
    styleHeader(proj, 3);
    // Date initiale
    proj.appendRow(['Overview Acceptance', 'Standard telecom site overview acceptance', 'TRUE']);
    proj.appendRow(['Full Site Audit',     'Complete audit with all equipment categories', 'TRUE']);
    proj.setColumnWidth(1, 250);
    proj.setColumnWidth(2, 350);
    proj.setColumnWidth(3, 80);
    var activeRule = SpreadsheetApp.newDataValidation().requireValueInList(['TRUE','FALSE'],true).build();
    proj.getRange('C2:C1000').setDataValidation(activeRule);
    Logger.log('Created Projects sheet');
  }

  // ── CATEGORIES — adauga coloana Project (col A) daca lipseste ──
  var c = ss.getSheetByName(SHEET_CATEGORIES);
  if (!c) {
    c = ss.insertSheet(SHEET_CATEGORIES);
    c.appendRow(['Project', 'Category', 'Subcategory', 'Required Photos', 'Active', 'Instructions', 'Image URLs']);
    styleHeader(c, 7);
    // Date initiale cu proiect
    c.appendRow(['Overview Acceptance', 'Overview Location', 'Key Safe',      5, 'TRUE', 'Photo the key safe box mounted on the wall or fence. Include full view and close-up of lock.', '']);
    c.appendRow(['Overview Acceptance', 'Overview Location', 'Access Gate',   3, 'TRUE', 'Capture access gate from outside and inside. Include padlock and signage.', '']);
    c.appendRow(['Overview Acceptance', 'Overview Location', 'Site Entrance', 2, 'TRUE', 'Photograph main site entrance, barriers, signs and road markings.', '']);
    c.appendRow(['Overview Acceptance', 'Equipment',         'Cabinet',       4, 'TRUE', 'Open cabinet doors and photograph interior. Include all shelves and cabling.', '']);
    c.appendRow(['Overview Acceptance', 'Equipment',         'Power Supply',  3, 'TRUE', 'Photograph PSU unit, input connections and output terminals.', '']);
    c.appendRow(['Overview Acceptance', 'Final Status',      'Site Overview', 5, 'TRUE', 'General overview photos of complete site from all four directions.', '']);
    c.appendRow(['Full Site Audit',     'Overview Location', 'Key Safe',      5, 'TRUE', 'Photo the key safe box.', '']);
    c.appendRow(['Full Site Audit',     'Equipment',         'Cabinet',       4, 'TRUE', 'Cabinet interior.', '']);
    c.appendRow(['Full Site Audit',     'Equipment',         'Power Supply',  3, 'TRUE', 'PSU unit photos.', '']);
    c.appendRow(['Full Site Audit',     'Equipment',         'Labels',        2, 'TRUE', 'All equipment labels and serial numbers.', '']);
    c.appendRow(['Full Site Audit',     'Antenna',           'Sector A',      6, 'TRUE', 'Antenna sector A, all angles.', '']);
    c.appendRow(['Full Site Audit',     'Antenna',           'Sector B',      6, 'TRUE', 'Antenna sector B, all angles.', '']);
    c.appendRow(['Full Site Audit',     'Grounding',         'Earth Bar',     4, 'TRUE', 'Earth bar and all connections.', '']);
    c.appendRow(['Full Site Audit',     'Final Status',      'Site Overview', 5, 'TRUE', 'Complete site overview.', '']);
    Logger.log('Created Categories sheet with projects');
  } else {
    // Sheet existent - verifica daca are coloana Project
    var headers = c.getRange(1, 1, 1, c.getLastColumn()).getValues()[0];
    var hasProject = headers[0] === 'Project';
    if (!hasProject) {
      // Insereaza coloana Project la inceput
      c.insertColumnBefore(1);
      c.getRange(1, 1).setValue('Project').setFontWeight('bold').setBackground('#0a1628').setFontColor('#29AAE1');
      // Completeaza cu proiectul default pentru randurile existente
      var lastRow = c.getLastRow();
      for (var i = 2; i <= lastRow; i++) {
        c.getRange(i, 1).setValue('Overview Acceptance');
      }
      Logger.log('Added Project column to existing Categories sheet');
    }
    // Asigura coloanele Instructions si Image URLs (col 6 si 7 in noua structura)
    var newHeaders = c.getRange(1, 1, 1, c.getLastColumn()).getValues()[0];
    var lastCol = newHeaders.length;
    if (lastCol < 6) { c.getRange(1, 6).setValue('Instructions'); styleHeaderCell(c, 1, 6); }
    if (lastCol < 7) { c.getRange(1, 7).setValue('Image URLs');   styleHeaderCell(c, 1, 7); }
  }
  c.setColumnWidth(1, 200);
  c.setColumnWidth(2, 200);
  c.setColumnWidth(3, 200);
  c.setColumnWidth(4, 130);
  c.setColumnWidth(5, 80);
  c.setColumnWidth(6, 400);
  c.setColumnWidth(7, 350);
  if (c.getFrozenRows() < 1) c.setFrozenRows(1);
  var activeRule2 = SpreadsheetApp.newDataValidation().requireValueInList(['TRUE','FALSE'],true).build();
  c.getRange('E2:E2000').setDataValidation(activeRule2);

  // ── SUBMISSIONS sheet (inlocuieste Projects din v1) ──
  var sub = ss.getSheetByName(SHEET_SUBMISSIONS);
  if (!sub) {
    sub = ss.insertSheet(SHEET_SUBMISSIONS);
    sub.appendRow(['Site Number','Project','User Name','User Email','Date','Total Photos','N/A Items','Status','N/A Details']);
    styleHeader(sub, 9);
    [120,200,160,200,180,120,90,100,400].forEach(function(w,i){ sub.setColumnWidth(i+1,w); });
    Logger.log('Created Submissions sheet');
  }

  // ── NA_LOG — adauga coloana Project daca lipseste ──
  var na = ss.getSheetByName(SHEET_NA_LOG);
  if (!na) {
    na = ss.insertSheet(SHEET_NA_LOG);
    na.appendRow(['Site Number','Project','Category','Subcategory','User Name','User Email','Date & Time','Status']);
    styleHeader(na, 8);
    Logger.log('Created NA_Log sheet');
  }

  // ── USERS — verifica ──
  var u = ss.getSheetByName(SHEET_USERS);
  if (!u) {
    u = ss.insertSheet(SHEET_USERS);
    u.appendRow(['Name','Email','Status','Date Requested']);
    styleHeader(u, 4);
    var sRule = SpreadsheetApp.newDataValidation().requireValueInList(['Pending','Approved','Rejected'],true).build();
    u.getRange('C2:C1000').setDataValidation(sRule);
    [200,260,140,200].forEach(function(w,i){ u.setColumnWidth(i+1,w); });
    Logger.log('Created Users sheet');
  }

  Logger.log('upgradeSheets DONE');
}

function styleHeader(sheet, numCols) {
  var r = sheet.getRange(1, 1, 1, numCols);
  r.setFontWeight('bold').setBackground('#0a1628').setFontColor('#29AAE1').setFontSize(11);
  sheet.setFrozenRows(1);
}

function styleHeaderCell(sheet, row, col) {
  sheet.getRange(row, col).setFontWeight('bold').setBackground('#0a1628').setFontColor('#29AAE1').setFontSize(11);
}

// -----------------------------------------------
//  HTTP HANDLERS
// -----------------------------------------------
function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function handleRequest(e) {
  try {
    var data = {};
    if (e && e.parameter && e.parameter.payload) {
      data = JSON.parse(e.parameter.payload);
    } else if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      data = e.parameter;
    }
    Logger.log('action=' + data.action);
    var result;
    switch (data.action) {
      case 'login':          result = actionLogin(data);          break;
      case 'requestAccess':  result = actionRequestAccess(data);  break;
      case 'getProjects':    result = actionGetProjects();         break;
      case 'getCategories':  result = actionGetCategories(data);  break;
      case 'uploadPhoto':    result = actionUploadPhoto(data);    break;
      case 'submitProject':  result = actionSubmitProject(data);  break;
      case 'logNA':          result = actionLogNA(data);          break;
      case 'ping':           result = { ok: true };               break;
      default:               result = { error: 'Unknown: ' + data.action };
    }
    return jsonOut(result);
  } catch(err) {
    Logger.log('ERROR: ' + err.message);
    return jsonOut({ error: err.message });
  }
}

// -----------------------------------------------
//  LOGIN
// -----------------------------------------------
function actionLogin(data) {
  var email = (data.email || '').toLowerCase().trim();
  if (!email) return { status: 'error' };
  var sheet = getSpreadsheet().getSheetByName(SHEET_USERS);
  if (!sheet) return { status: 'not_found' };
  var last = sheet.getLastRow();
  if (last < 2) return { status: 'not_found' };
  var rows = sheet.getRange(2, 1, last - 1, 4).getValues();
  for (var i = 0; i < rows.length; i++) {
    if ((rows[i][1] || '').toLowerCase().trim() !== email) continue;
    var st = (rows[i][2] || '').trim().toLowerCase();
    if (st === 'approved') return { status: 'approved', name: String(rows[i][0]).trim() };
    if (st === 'pending')  return { status: 'pending' };
    return { status: 'rejected' };
  }
  return { status: 'not_found' };
}

// -----------------------------------------------
//  REQUEST ACCESS
// -----------------------------------------------
function actionRequestAccess(data) {
  var name  = (data.name  || '').trim();
  var email = (data.email || '').toLowerCase().trim();
  if (!name || !email) return { success: false, message: 'Name and email required' };
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) { sheet = ss.insertSheet(SHEET_USERS); sheet.appendRow(['Name','Email','Status','Date Requested']); styleHeader(sheet, 4); }
  var last = sheet.getLastRow();
  if (last >= 2) {
    var rows = sheet.getRange(2, 1, last - 1, 3).getValues();
    for (var i = 0; i < rows.length; i++) {
      if ((rows[i][1] || '').toLowerCase().trim() === email) {
        if ((rows[i][2] || '').toLowerCase() === 'approved') return { success: false, message: 'already_approved' };
        return { success: false, message: 'Request already exists.' };
      }
    }
  }
  sheet.appendRow([name, email, 'Pending', new Date().toISOString()]);
  SpreadsheetApp.flush();
  Logger.log('New request: ' + name + ' <' + email + '>');
  return { success: true };
}

// -----------------------------------------------
//  GET PROJECTS
//  Returneaza lista proiectelor active
// -----------------------------------------------
function actionGetProjects() {
  var sheet = getSpreadsheet().getSheetByName(SHEET_PROJECTS);
  if (!sheet) return { projects: [] };
  var last = sheet.getLastRow();
  if (last < 2) return { projects: [] };
  var rows = sheet.getRange(2, 1, last - 1, 3).getValues();
  var projects = [];
  for (var i = 0; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    if (String(rows[i][2]).toUpperCase() === 'FALSE') continue;
    projects.push({
      name:        String(rows[i][0]).trim(),
      description: String(rows[i][1] || '').trim()
    });
  }
  return { projects: projects };
}

// -----------------------------------------------
//  GET CATEGORIES (filtrat pe proiect)
// -----------------------------------------------
function actionGetCategories(data) {
  var projectName = (data.project || '').trim();
  var sheet = getSpreadsheet().getSheetByName(SHEET_CATEGORIES);
  if (!sheet) return { categories: [] };
  var last = sheet.getLastRow();
  if (last < 2) return { categories: [] };
  var totalCols = Math.max(sheet.getLastColumn(), 7);
  var rows = sheet.getRange(2, 1, last - 1, totalCols).getValues();
  var cats = [];
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (!row[0] || !row[1] || !row[2]) continue;
    if (String(row[4]).toUpperCase() === 'FALSE') continue;
    // Filtreaza pe proiect daca e specificat
    if (projectName && String(row[0]).trim() !== projectName) continue;
    cats.push({
      project:      String(row[0]).trim(),
      category:     String(row[1]).trim(),
      subcategory:  String(row[2]).trim(),
      required:     Number(row[3]) || 1,
      active:       true,
      instructions: totalCols > 5 ? String(row[5] || '').trim() : '',
      imageUrls:    totalCols > 6 ? String(row[6] || '').trim() : ''
    });
  }
  return { categories: cats };
}

// -----------------------------------------------
//  UPLOAD PHOTO
//  Folder: Root / Site_XXXXX / Project / Category / Subcategory
// -----------------------------------------------
function actionUploadPhoto(data) {
  var siteNumber  = (data.siteNumber  || '').trim();
  var project     = (data.project     || '').trim();
  var category    = (data.category    || '').trim();
  var subcategory = (data.subcategory || '').trim();
  var fileName    = (data.fileName    || 'photo.jpg').trim();
  var base64Data  = data.base64Data   || '';
  var mimeType    = data.mimeType     || 'image/jpeg';
  if (!siteNumber || !base64Data) return { success: false, message: 'Missing fields' };

  var root    = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  var siteFol = getOrCreateFolder(root, 'Site_' + siteNumber);
  var projFol = project ? getOrCreateFolder(siteFol, project) : siteFol;
  var catFol  = getOrCreateFolder(projFol, category);
  var subFol  = getOrCreateFolder(catFol, subcategory);

  var ex = subFol.getFilesByName(fileName);
  if (ex.hasNext()) return { success: true, fileId: ex.next().getId() };

  var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
  var file = subFol.createFile(blob);
  return { success: true, fileId: file.getId(), fileName: fileName };
}

// -----------------------------------------------
//  LOG N/A
// -----------------------------------------------
function actionLogNA(data) {
  var sheet = getSpreadsheet().getSheetByName(SHEET_NA_LOG);
  if (!sheet) return { success: false };
  sheet.appendRow([
    'Site_' + (data.siteNumber || ''),
    data.project     || '',
    data.category    || '',
    data.subcategory || '',
    data.userName    || '',
    data.userEmail   || '',
    data.markedAt    || new Date().toISOString(),
    'N/A'
  ]);
  SpreadsheetApp.flush();
  return { success: true };
}

// -----------------------------------------------
//  SUBMIT PROJECT
// -----------------------------------------------
function actionSubmitProject(data) {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_SUBMISSIONS);
  if (!sheet) return { success: false, message: 'Submissions sheet not found' };

  var naDetails    = data.naDetails || [];
  var naDetailsStr = naDetails.length
    ? naDetails.map(function(n){ return n.category + ' > ' + n.subcategory; }).join('; ')
    : 'None';

  sheet.appendRow([
    'Site_' + (data.siteNumber || ''),
    data.project     || '',
    data.userName    || '',
    data.userEmail   || '',
    data.date        || new Date().toISOString(),
    Number(data.totalPhotos) || 0,
    Number(data.totalNA)     || 0,
    'Completed',
    naDetailsStr
  ]);

  // Flush NA logs lipsа (offline scenario)
  if (naDetails.length) {
    var naSheet = ss.getSheetByName(SHEET_NA_LOG);
    if (naSheet) {
      var lastRow = naSheet.getLastRow();
      var existSet = {};
      if (lastRow >= 2) {
        naSheet.getRange(2, 1, lastRow - 1, 4).getValues()
          .forEach(function(r){ existSet[r[0]+'|'+r[1]+'|'+r[2]+'|'+r[3]] = true; });
      }
      naDetails.forEach(function(n) {
        var k = 'Site_'+(data.siteNumber||'')+'|'+(data.project||'')+'|'+n.category+'|'+n.subcategory;
        if (!existSet[k]) {
          naSheet.appendRow(['Site_'+(data.siteNumber||''), data.project||'', n.category, n.subcategory,
            data.userName||'', data.userEmail||'', data.date||new Date().toISOString(), 'N/A']);
        }
      });
    }
  }
  SpreadsheetApp.flush();
  return { success: true };
}

// -----------------------------------------------
//  HELPER
// -----------------------------------------------
function getOrCreateFolder(parent, name) {
  var f = parent.getFoldersByName(name);
  if (f.hasNext()) return f.next();
  return parent.createFolder(name);
}
