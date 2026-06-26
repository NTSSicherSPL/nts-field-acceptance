/**
 * NTS Field Acceptance — Google Apps Script Backend v3.1
 * ═══════════════════════════════════════════════════════
 * PASI DEPLOY:
 * 1. Înlocuiește codul în script.google.com
 * 2. Salvează (Ctrl+S)
 * 3. Run > setupSheets   ← rulează DOAR PRIMA DATĂ (migrare + creare)
 * 4. Deploy > Manage Deployments > Edit > New Version > Deploy
 *
 * STRUCTURA SPREADSHEET:
 *   Users              — Name | Email | Status | Date Requested
 *   Projects           — Project Name | Description | Active
 *   Category_Master    — ID | Category | Subcategory | Default Required Photos | Instructions | Image URLs
 *   Project_Categories — Project | Category ID | Required Photos Override | Active
 *   Submissions        — Site Number | Project | User Name | User Email | Date | Total Photos | N/A Items | Status | N/A Details
 *   NA_Log             — Site Number | Project | Category | Subcategory | User Name | User Email | Date & Time | Status
 */

var SPREADSHEET_ID  = '1GRUPV9eAIKlv9P2JFpWEPUZ6iVZ394l8_nz9fN2yypc';
var DRIVE_FOLDER_ID = '1ug6huCS_OZEbKzaO5d-WQ1hWC-XBFHMy';

var SHEET_USERS       = 'Users';
var SHEET_PROJECTS    = 'Projects';
var SHEET_CAT_MASTER  = 'Category_Master';
var SHEET_PROJ_CATS   = 'Project_Categories';
var SHEET_SUBMISSIONS = 'Submissions';
var SHEET_NA_LOG      = 'NA_Log';
var SHEET_USER_PROJECTS = 'User_Projects';
var SHEET_ACTIVE_SITES  = 'Active_Sites';

// Sheet-uri vechi care vor fi șterse după migrare
var SHEET_OLD_CATEGORIES = 'Categories';

// ═══════════════════════════════════════════════════════
//  SETUP SHEETS — rulează PRIMA DATĂ după deploy
//  1. Creează sheet-urile noi
//  2. Migrează datele din vechiul Categories (dacă există)
//  3. Curăță sheet-urile vechi/duplicate
//  4. Formatează tot
// ═══════════════════════════════════════════════════════
function setupSheets() {
  var ss = getSpreadsheet();
  var log = [];

  // ── STEP 1: Asigură sheet-urile de bază (Users, Projects) ──
  ensureUsers(ss, log);
  ensureProjects(ss, log);

  // ── STEP 2: Creează Category_Master ──
  var cm = ensureCategoryMaster(ss, log);

  // ── STEP 3: Migrează din vechiul Categories dacă există ──
  var oldCats = ss.getSheetByName(SHEET_OLD_CATEGORIES);
  if (oldCats) {
    migrateOldCategories(ss, oldCats, cm, log);
  }

  // ── STEP 4: Creează Project_Categories ──
  ensureProjectCategories(ss, log);

  // ── STEP 5: Asigură asignările, lucrările active, Submissions și NA_Log ──
  ensureUserProjects(ss, log);
  ensureActiveSites(ss, log);
  ensureSubmissions(ss, log);
  ensureNALog(ss, log);

  // ── STEP 6: Curăță sheet-urile vechi ──
  cleanupOldSheets(ss, log);

  // ── STEP 7: Ordonează sheet-urile ──
  orderSheets(ss);

  Logger.log(log.join('\n'));
  SpreadsheetApp.getUi().alert(
    '✅ Setup complet!\n\n' + log.join('\n') +
    '\n\n📋 Pași următori:\n' +
    '1. Verifică "Category_Master" — adaugă/editează categoriile tale\n' +
    '2. Verifică "Project_Categories" — alocă categoriile la proiecte\n' +
    '3. Deploy > New Version în Apps Script'
  );
}

// -----------------------------------------------
//  ENSURE USERS
// -----------------------------------------------
function ensureUsers(ss, log) {
  var sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_USERS);
    sheet.appendRow(['Name', 'Email', 'Status', 'Date Requested']);
    styleHeader(sheet, 4);
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pending', 'Approved', 'Rejected'], true).build();
    sheet.getRange('C2:C1000').setDataValidation(rule);
    [200, 260, 140, 200].forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
    log.push('✓ Created: Users');
  } else {
    log.push('→ Kept existing: Users (' + (sheet.getLastRow() - 1) + ' users)');
  }
}

// -----------------------------------------------
//  ENSURE PROJECTS
// -----------------------------------------------
function ensureProjects(ss, log) {
  var sheet = ss.getSheetByName(SHEET_PROJECTS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PROJECTS);
    sheet.appendRow(['Project Name', 'Description', 'Active']);
    styleHeader(sheet, 3);
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['TRUE', 'FALSE'], true).build();
    sheet.getRange('C2:C1000').setDataValidation(rule);
    sheet.setColumnWidth(1, 250);
    sheet.setColumnWidth(2, 350);
    sheet.setColumnWidth(3, 80);
    log.push('✓ Created: Projects (empty — adaugă proiectele tale)');
  } else {
    log.push('→ Kept existing: Projects (' + (sheet.getLastRow() - 1) + ' proiecte)');
  }
}

// -----------------------------------------------
//  ENSURE CATEGORY_MASTER
//  Returnează sheet-ul (nou sau existent)
// -----------------------------------------------
function ensureCategoryMaster(ss, log) {
  var sheet = ss.getSheetByName(SHEET_CAT_MASTER);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CAT_MASTER);
    sheet.appendRow(['ID', 'Category', 'Subcategory', 'Default Required Photos', 'Instructions', 'Image URLs']);
    styleHeader(sheet, 6);
    sheet.setColumnWidth(1, 50);
    sheet.setColumnWidth(2, 180);
    sheet.setColumnWidth(3, 200);
    sheet.setColumnWidth(4, 170);
    sheet.setColumnWidth(5, 420);
    sheet.setColumnWidth(6, 300);
    sheet.setFrozenRows(1);
    // Validare: Required Photos să fie număr
    var numRule = SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThan(0).build();
    sheet.getRange('D2:D5000').setDataValidation(numRule);
    log.push('✓ Created: Category_Master (empty — datele vor fi migrate din Categories)');
  } else {
    log.push('→ Kept existing: Category_Master (' + (sheet.getLastRow() - 1) + ' categorii)');
  }
  return sheet;
}

// -----------------------------------------------
//  MIGRARE din vechiul sheet Categories
//  Structura veche: Project | Category | Subcategory | Required Photos | Active | Instructions | Image URLs
//  sau:             Category | Subcategory | Required Photos | Active | Instructions | Image URLs
// -----------------------------------------------
function migrateOldCategories(ss, oldSheet, masterSheet, log) {
  var lastRow = oldSheet.getLastRow();
  if (lastRow < 2) {
    log.push('⚠ Vechiul Categories e gol — nimic de migrat');
    return;
  }

  var totalCols = oldSheet.getLastColumn();
  var rows = oldSheet.getRange(2, 1, lastRow - 1, totalCols).getValues();
  var headers = oldSheet.getRange(1, 1, 1, totalCols).getValues()[0];

  // Detectează automat structura: are coloana "Project" la start?
  var hasProject = headers[0].toString().toLowerCase().indexOf('project') >= 0;
  var offset = hasProject ? 1 : 0; // câte coloane de shift

  // Colectează subcategorii unice (Category + Subcategory)
  var seen = {};       // key: "Category|||Subcategory" → row data
  var masterSheet_lastRow = masterSheet.getLastRow();

  // Verifică ce există deja în Category_Master
  var existingInMaster = {};
  if (masterSheet_lastRow >= 2) {
    masterSheet.getRange(2, 1, masterSheet_lastRow - 1, 3).getValues().forEach(function(r) {
      if (r[1] && r[2]) existingInMaster[r[1] + '|||' + r[2]] = true;
    });
  }

  // Găsește ultimul ID din Category_Master
  var nextId = 1;
  if (masterSheet_lastRow >= 2) {
    var ids = masterSheet.getRange(2, 1, masterSheet_lastRow - 1, 1).getValues()
      .map(function(r) { return Number(r[0]) || 0; });
    nextId = Math.max.apply(null, ids) + 1;
  }

  var migrated = 0;
  var skipped  = 0;

  rows.forEach(function(row) {
    var category    = String(row[0 + offset] || '').trim();
    var subcategory = String(row[1 + offset] || '').trim();
    var required    = Number(row[2 + offset]) || 1;
    var active      = String(row[3 + offset] || 'TRUE').toUpperCase();
    var instructions= String(row[4 + offset] || '').trim();
    var imageUrls   = String(row[5 + offset] || '').trim();

    if (!category || !subcategory) return;
    if (active === 'FALSE') { skipped++; return; }

    var key = category + '|||' + subcategory;
    if (existingInMaster[key]) { skipped++; return; } // deja există în master
    if (seen[key]) return; // duplicat în sursă

    seen[key] = true;
    existingInMaster[key] = true;
    masterSheet.appendRow([nextId, category, subcategory, required, instructions, imageUrls]);
    nextId++;
    migrated++;
  });

  log.push('✓ Migrat din Categories → Category_Master: ' + migrated + ' subcategorii unice (' + skipped + ' sărite)');
}

// -----------------------------------------------
//  ENSURE PROJECT_CATEGORIES
// -----------------------------------------------
function ensureProjectCategories(ss, log) {
  var sheet = ss.getSheetByName(SHEET_PROJ_CATS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PROJ_CATS);
    sheet.appendRow(['Project', 'Category ID', 'Required Photos Override', 'Active']);
    styleHeader(sheet, 4);
    sheet.setColumnWidth(1, 220);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 200);
    sheet.setColumnWidth(4, 80);
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['TRUE', 'FALSE'], true).build();
    sheet.getRange('D2:D5000').setDataValidation(rule);
    sheet.setFrozenRows(1);

    // Instrucțiuni în sheet — rândul 2 ca note vizibile
    sheet.getRange('A2').setNote(
      'Completează manual:\n' +
      '- Coloana A: numele exact al proiectului (din sheet-ul Projects)\n' +
      '- Coloana B: ID-ul categoriei (din sheet-ul Category_Master)\n' +
      '- Coloana C: lasă gol pentru valoarea default, sau pune număr custom\n' +
      '- Coloana D: TRUE / FALSE'
    );

    log.push('✓ Created: Project_Categories (completează manual alocările)');
  } else {
    log.push('→ Kept existing: Project_Categories (' + (sheet.getLastRow() - 1) + ' alocări)');
  }
}

// -----------------------------------------------
//  ENSURE USER_PROJECTS
// -----------------------------------------------
function ensureUserProjects(ss, log) {
  var sheet = ss.getSheetByName(SHEET_USER_PROJECTS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_USER_PROJECTS);
    sheet.appendRow(['Email', 'Project', 'Active']);
    styleHeader(sheet, 3);
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['TRUE', 'FALSE'], true).build();
    sheet.getRange('C2:C5000').setDataValidation(rule);
    sheet.setColumnWidth(1, 260);
    sheet.setColumnWidth(2, 240);
    sheet.setColumnWidth(3, 80);
    sheet.setFrozenRows(1);
    sheet.getRange('A2').setNote(
      'Completeaza manual:\n' +
      '- Coloana A: email utilizator aprobat\n' +
      '- Coloana B: proiect asignat\n' +
      '- Coloana C: TRUE / FALSE\n' +
      'Adauga cate un rand pentru fiecare proiect asignat.'
    );
    log.push('✓ Created: User_Projects (asigneaza utilizatori la proiecte)');
  } else {
    log.push('→ Kept existing: User_Projects (' + (sheet.getLastRow() - 1) + ' asignari)');
  }
}

// -----------------------------------------------
//  ENSURE ACTIVE_SITES
// -----------------------------------------------
function ensureActiveSites(ss, log) {
  var sheet = ss.getSheetByName(SHEET_ACTIVE_SITES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_ACTIVE_SITES);
    sheet.appendRow(['Site Number','Project','Created By Name','Created By Email','Created At','Status','Last Updated']);
    styleHeader(sheet, 7);
    [120,220,180,240,180,100,180].forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
    sheet.setFrozenRows(1);
    log.push('✓ Created: Active_Sites');
  } else {
    log.push('→ Kept existing: Active_Sites (' + (sheet.getLastRow() - 1) + ' lucrari)');
  }
}

// -----------------------------------------------
//  ENSURE SUBMISSIONS
// -----------------------------------------------
function ensureSubmissions(ss, log) {
  var sheet = ss.getSheetByName(SHEET_SUBMISSIONS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SUBMISSIONS);
    sheet.appendRow(['Site Number','Project','User Name','User Email','Date','Total Photos','N/A Items','Status','N/A Details']);
    styleHeader(sheet, 9);
    [120,200,160,200,180,120,90,100,400].forEach(function(w, i) { sheet.setColumnWidth(i+1, w); });
    log.push('✓ Created: Submissions');
  } else {
    // Asigură că are coloana Project (v2 → v3 upgrade)
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers[1] !== 'Project') {
      sheet.insertColumnAfter(1);
      sheet.getRange(1, 2).setValue('Project');
      styleHeaderCell(sheet, 1, 2);
      log.push('✓ Added Project column to Submissions');
    } else {
      log.push('→ Kept existing: Submissions (' + (sheet.getLastRow() - 1) + ' înregistrări)');
    }
  }
}

// -----------------------------------------------
//  ENSURE NA_LOG
// -----------------------------------------------
function ensureNALog(ss, log) {
  var sheet = ss.getSheetByName(SHEET_NA_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NA_LOG);
    sheet.appendRow(['Site Number','Project','Category','Subcategory','User Name','User Email','Date & Time','Status']);
    styleHeader(sheet, 8);
    log.push('✓ Created: NA_Log');
  } else {
    log.push('→ Kept existing: NA_Log (' + (sheet.getLastRow() - 1) + ' înregistrări)');
  }
}

// -----------------------------------------------
//  CLEANUP — șterge sheet-urile vechi/duplicate
// -----------------------------------------------
function cleanupOldSheets(ss, log) {
  var toDelete = [
    SHEET_OLD_CATEGORIES,   // vechiul Categories
    'Categories_OLD',        // dacă există din setup anterior
    'Sheet1', 'Sheet2', 'Sheet3', 'Foaie1', 'Foaie2', 'Foaie3'  // sheet-uri default goale
  ];

  toDelete.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    // Nu șterge dacă are date importante (mai mult de 1 rând = header + date)
    if (name === SHEET_OLD_CATEGORIES || name === 'Categories_OLD') {
      // Șterge întotdeauna vechiul Categories după migrare
      ss.deleteSheet(sheet);
      log.push('🗑 Deleted: ' + name + ' (migrat în Category_Master)');
    } else {
      // Șterge sheet-urile default goale
      if (sheet.getLastRow() <= 1 && sheet.getLastColumn() <= 1) {
        try { ss.deleteSheet(sheet); log.push('🗑 Deleted: ' + name + ' (gol)'); } catch(e) {}
      }
    }
  });
}

// -----------------------------------------------
//  ORDONEAZĂ SHEET-URILE în spreadsheet
// -----------------------------------------------
function orderSheets(ss) {
  var order = [
    SHEET_USERS,
    SHEET_PROJECTS,
    SHEET_CAT_MASTER,
    SHEET_PROJ_CATS,
    SHEET_USER_PROJECTS,
    SHEET_ACTIVE_SITES,
    SHEET_SUBMISSIONS,
    SHEET_NA_LOG
  ];
  order.forEach(function(name, idx) {
    var sheet = ss.getSheetByName(name);
    if (sheet) ss.setActiveSheet(sheet), ss.moveActiveSheet(idx + 1);
  });
}

// -----------------------------------------------
//  HELPERS DE FORMATARE
// -----------------------------------------------
function styleHeader(sheet, numCols) {
  var r = sheet.getRange(1, 1, 1, numCols);
  r.setFontWeight('bold')
   .setBackground('#0a1628')
   .setFontColor('#29AAE1')
   .setFontSize(11)
   .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sheet.setFrozenRows(1);
}

function styleHeaderCell(sheet, row, col) {
  sheet.getRange(row, col)
    .setFontWeight('bold')
    .setBackground('#0a1628')
    .setFontColor('#29AAE1')
    .setFontSize(11);
}

// -----------------------------------------------
//  HTTP HANDLERS
// -----------------------------------------------
function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
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
      case 'login':          result = actionLogin(data);         break;
      case 'requestAccess':  result = actionRequestAccess(data); break;
      case 'getProjects':    result = actionGetProjects(data);    break;
      case 'getOpenSites':   result = actionGetOpenSites(data);   break;
      case 'startSite':      result = actionStartSite(data);      break;
      case 'getSiteState':   result = actionGetSiteState(data);   break;
      case 'getCategories':  result = actionGetCategories(data); break;
      case 'uploadPhoto':    result = actionUploadPhoto(data);   break;
      case 'submitProject':  result = actionSubmitProject(data); break;
      case 'logNA':          result = actionLogNA(data);         break;
      case 'ping':           result = { ok: true };              break;
      default:               result = { error: 'Unknown action: ' + data.action };
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
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_USERS);
    sheet.appendRow(['Name','Email','Status','Date Requested']);
    styleHeader(sheet, 4);
  }
  var last = sheet.getLastRow();
  if (last >= 2) {
    var rows = sheet.getRange(2, 1, last - 1, 3).getValues();
    for (var i = 0; i < rows.length; i++) {
      if ((rows[i][1] || '').toLowerCase().trim() === email) {
        if ((rows[i][2] || '').toLowerCase() === 'approved')
          return { success: false, message: 'already_approved' };
        return { success: false, message: 'Request already exists.' };
      }
    }
  }
  sheet.appendRow([name, email, 'Pending', new Date().toISOString()]);
  SpreadsheetApp.flush();
  return { success: true };
}

// -----------------------------------------------
//  GET PROJECTS
//  Citește proiectele unice direct din Project_Categories
//  + descrierea din sheet-ul Projects (dacă există)
// -----------------------------------------------
function actionGetProjects(data) {
  data = data || {};
  var email = (data.email || '').toLowerCase().trim();
  var ss = getSpreadsheet();
  var allowedMap = getAllowedProjectMap(ss, email);

  // Citește descrierile din Projects (opțional)
  var descMap = {};
  var projSheet = ss.getSheetByName(SHEET_PROJECTS);
  if (projSheet && projSheet.getLastRow() >= 2) {
    projSheet.getRange(2, 1, projSheet.getLastRow() - 1, 3).getValues()
      .forEach(function(r) {
        if (r[0]) descMap[String(r[0]).trim()] = {
          description: String(r[1] || '').trim(),
          active: String(r[2] || 'TRUE').toUpperCase()
        };
      });
  }

  // Citește proiectele unice din Project_Categories
  var pc = ss.getSheetByName(SHEET_PROJ_CATS);
  if (!pc) return { projects: [], error: 'Sheet Project_Categories lipseste. Ruleaza setupSheets().' };
  var last = pc.getLastRow();
  if (last < 2) return { projects: [], error: 'Sheet Project_Categories este gol.' };
  var rows = pc.getRange(2, 1, last - 1, 4).getValues();

  var seen = {};
  var projects = [];
  rows.forEach(function(r) {
    var name   = String(r[0] || '').trim();
    var active = String(r[3] || 'TRUE').toUpperCase();
    if (!name || seen[name]) return;
    if (active === 'FALSE') return;
    if (descMap[name] && descMap[name].active === 'FALSE') return;
    if (allowedMap && !allowedMap[name]) return;
    seen[name] = true;
    projects.push({
      name:        name,
      description: descMap[name] ? descMap[name].description : ''
    });
  });

  Logger.log('getProjects: ' + projects.length + ' proiecte din Project_Categories');
  return { projects: projects };
}

function getAllowedProjectMap(ss, email) {
  var up = ss.getSheetByName(SHEET_USER_PROJECTS);
  if (!up || up.getLastRow() < 2) return null;
  var rows = up.getRange(2, 1, up.getLastRow() - 1, 3).getValues();
  var hasAssignments = false;
  var allowed = {};
  rows.forEach(function(r) {
    var rowEmail = String(r[0] || '').toLowerCase().trim();
    var project  = String(r[1] || '').trim();
    var active   = String(r[2] || 'TRUE').toUpperCase();
    if (!rowEmail || !project || active === 'FALSE') return;
    hasAssignments = true;
    if (email && rowEmail === email) allowed[project] = true;
  });
  if (!hasAssignments) return null;
  return allowed;
}

function userCanAccessProject(ss, email, project) {
  var allowedMap = getAllowedProjectMap(ss, email);
  return !allowedMap || !!allowedMap[project];
}

// -----------------------------------------------
//  GET OPEN SITES
// -----------------------------------------------
function actionGetOpenSites(data) {
  var email = (data.email || '').toLowerCase().trim();
  var projectFilter = (data.project || '').trim();
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_ACTIVE_SITES);
  if (!sheet || sheet.getLastRow() < 2) return { sites: [] };

  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  var sites = [];
  rows.forEach(function(r) {
    var site    = String(r[0] || '').replace(/^Site_/, '').trim();
    var project = String(r[1] || '').trim();
    var status  = String(r[5] || 'Open').toLowerCase();
    if (!site || !project || status !== 'open') return;
    if (projectFilter && project !== projectFilter) return;
    if (!userCanAccessProject(ss, email, project)) return;
    sites.push({
      siteNumber: site,
      project: project,
      createdByName: String(r[2] || '').trim(),
      createdByEmail: String(r[3] || '').trim(),
      createdAt: r[4] ? String(r[4]) : '',
      lastUpdated: r[6] ? String(r[6]) : ''
    });
  });
  return { sites: sites };
}

// -----------------------------------------------
//  START SITE
// -----------------------------------------------
function actionStartSite(data) {
  var siteNumber = (data.siteNumber || '').replace(/^Site_/, '').trim();
  var project    = (data.project || '').trim();
  var userEmail  = (data.userEmail || '').toLowerCase().trim();
  var userName   = (data.userName || '').trim();
  if (!siteNumber || !project) return { success: false, message: 'Missing site/project' };

  var ss = getSpreadsheet();
  if (!userCanAccessProject(ss, userEmail, project)) {
    return { success: false, message: 'User is not assigned to this project' };
  }

  var sheet = ss.getSheetByName(SHEET_ACTIVE_SITES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_ACTIVE_SITES);
    sheet.appendRow(['Site Number','Project','Created By Name','Created By Email','Created At','Status','Last Updated']);
    styleHeader(sheet, 7);
  }

  var now = new Date().toISOString();
  var last = sheet.getLastRow();
  if (last >= 2) {
    var rows = sheet.getRange(2, 1, last - 1, 7).getValues();
    for (var i = 0; i < rows.length; i++) {
      var rowSite = String(rows[i][0] || '').replace(/^Site_/, '').trim();
      var rowProject = String(rows[i][1] || '').trim();
      var status = String(rows[i][5] || 'Open').toLowerCase();
      if (rowSite === siteNumber && rowProject === project && status === 'open') {
        sheet.getRange(i + 2, 7).setValue(now);
        return { success: true, existing: true };
      }
    }
  }

  sheet.appendRow(['Site_' + siteNumber, project, userName, userEmail, now, 'Open', now]);
  SpreadsheetApp.flush();
  return { success: true, existing: false };
}
// -----------------------------------------------
//  GET CATEGORIES
//  JOIN: Category_Master + Project_Categories
// -----------------------------------------------
function actionGetCategories(data) {
  var projectName = (data.project || '').trim();
  var userEmail = (data.userEmail || data.email || '').toLowerCase().trim();
  var ss = getSpreadsheet();
  if (projectName && !userCanAccessProject(ss, userEmail, projectName)) {
    return { categories: [], error: 'User is not assigned to this project' };
  }

  // Citește Category_Master → map ID → detalii
  var cm = ss.getSheetByName(SHEET_CAT_MASTER);
  if (!cm) return { categories: [] };
  var cmLast = cm.getLastRow();
  if (cmLast < 2) return { categories: [] };
  var cmRows = cm.getRange(2, 1, cmLast - 1, 6).getValues();
  var masterMap = {};
  cmRows.forEach(function(r) {
    if (!r[0]) return;
    masterMap[String(r[0]).trim()] = {
      category:     String(r[1] || '').trim(),
      subcategory:  String(r[2] || '').trim(),
      required:     Number(r[3]) || 1,
      instructions: String(r[4] || '').trim(),
      imageUrls:    String(r[5] || '').trim()
    };
  });

  // Citește Project_Categories → filtrează pe proiect
  var pc = ss.getSheetByName(SHEET_PROJ_CATS);
  if (!pc) return { categories: [] };
  var pcLast = pc.getLastRow();
  if (pcLast < 2) return { categories: [] };
  var pcRows = pc.getRange(2, 1, pcLast - 1, 4).getValues();

  var categories = [];
  pcRows.forEach(function(r) {
    var rowProject  = String(r[0] || '').trim();
    var catId       = String(r[1] || '').trim();
    var reqOverride = r[2] !== '' && r[2] !== null ? Number(r[2]) : null;
    var active      = String(r[3] || 'TRUE').toUpperCase();

    if (projectName && rowProject !== projectName) return;
    if (active === 'FALSE') return;
    if (!catId || !masterMap[catId]) return;

    var m = masterMap[catId];
    categories.push({
      project:      rowProject,
      category:     m.category,
      subcategory:  m.subcategory,
      required:     reqOverride !== null ? reqOverride : m.required,
      instructions: m.instructions,
      imageUrls:    m.imageUrls
    });
  });

  return { categories: categories };
}

// -----------------------------------------------
//  UPLOAD PHOTO
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
  var userEmail = (data.userEmail || '').toLowerCase().trim();
  if (project && !userCanAccessProject(getSpreadsheet(), userEmail, project)) {
    return { success: false, message: 'User is not assigned to this project' };
  }

  var root    = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  var siteFol = getOrCreateFolder(root, 'Site_' + siteNumber);
  var projFol = project ? getOrCreateFolder(siteFol, project) : siteFol;
  var catFol  = getOrCreateFolder(projFol, category);

  fileName = getNextAvailablePhotoName(catFol, fileName);

  var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
  var file = catFol.createFile(blob);
  file.setDescription(JSON.stringify({
    siteNumber: siteNumber,
    project: project,
    category: category,
    subcategory: subcategory
  }));
  touchActiveSite(siteNumber, project);
  return { success: true, fileId: file.getId(), fileName: fileName };
}

function getNextAvailablePhotoName(folder, fileName) {
  if (!folder.getFilesByName(fileName).hasNext()) return fileName;

  var dot = fileName.lastIndexOf('.');
  var ext = dot >= 0 ? fileName.substring(dot) : '.jpg';
  var base = dot >= 0 ? fileName.substring(0, dot) : fileName;
  var m = base.match(/^(.*_)(\d{3})$/);
  if (!m) {
    var fallback = base + '_001' + ext;
    return folder.getFilesByName(fallback).hasNext()
      ? getNextAvailablePhotoName(folder, fallback)
      : fallback;
  }

  var prefix = m[1];
  var n = Number(m[2]) || 1;
  var nextName = fileName;
  do {
    n++;
    nextName = prefix + String(n).padStart(3, '0') + ext;
  } while (folder.getFilesByName(nextName).hasNext());
  return nextName;
}

// -----------------------------------------------
//  GET SITE STATE FROM DRIVE + N/A LOG
// -----------------------------------------------
function actionGetSiteState(data) {
  var siteNumber = (data.siteNumber || '').replace(/^Site_/, '').trim();
  var project    = (data.project || '').trim();
  var userEmail  = (data.userEmail || data.email || '').toLowerCase().trim();
  if (!siteNumber || !project) return { photos: [], naStatus: {} };

  var ss = getSpreadsheet();
  if (!userCanAccessProject(ss, userEmail, project)) {
    return { photos: [], naStatus: {}, error: 'User is not assigned to this project' };
  }

  var photos = [];
  var subSlugMap = buildSubcategorySlugMap(ss, project);
  try {
    var root = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    var siteFolders = root.getFoldersByName('Site_' + siteNumber);
    if (siteFolders.hasNext()) {
      var siteFol = siteFolders.next();
      var projectFolders = siteFol.getFoldersByName(project);
      if (projectFolders.hasNext()) {
        var projFol = projectFolders.next();
        var catFolders = projFol.getFolders();
        while (catFolders.hasNext()) {
          var catFol = catFolders.next();
          var catName = catFol.getName();
          var directFiles = catFol.getFiles();
          while (directFiles.hasNext()) {
            var directFile = directFiles.next();
            var directMeta = getPhotoMetaFromDescription(directFile);
            var directSub = directMeta.subcategory || inferSubcategoryFromPhotoName(
              directFile.getName(),
              siteNumber,
              catName,
              subSlugMap[catName] || {}
            );
            if (directSub) {
              photos.push({
                category: directMeta.category || catName,
                subcategory: directSub,
                key: (directMeta.category || catName) + '|||' + directSub,
                name: directFile.getName(),
                fileId: directFile.getId(),
                thumbUrl: 'https://drive.google.com/thumbnail?id=' + directFile.getId() + '&sz=w400'
              });
            }
          }

          var subFolders = catFol.getFolders();
          while (subFolders.hasNext()) {
            var subFol = subFolders.next();
            var files = subFol.getFiles();
            while (files.hasNext()) {
              var file = files.next();
              photos.push({
                category: catName,
                subcategory: subFol.getName(),
                key: catName + '|||' + subFol.getName(),
                name: file.getName(),
                fileId: file.getId(),
                thumbUrl: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w400'
              });
            }
          }
        }
      }
    }
  } catch(e) {
    Logger.log('getSiteState Drive error: ' + e.message);
  }

  var naStatus = {};
  var naSheet = ss.getSheetByName(SHEET_NA_LOG);
  if (naSheet && naSheet.getLastRow() >= 2) {
    naSheet.getRange(2, 1, naSheet.getLastRow() - 1, 8).getValues().forEach(function(r) {
      var rowSite = String(r[0] || '').replace(/^Site_/, '').trim();
      var rowProject = String(r[1] || '').trim();
      if (rowSite !== siteNumber || rowProject !== project) return;
      var cat = String(r[2] || '').trim();
      var sub = String(r[3] || '').trim();
      if (!cat || !sub) return;
      naStatus[cat + '|||' + sub] = {
        markedBy: String(r[4] || '').trim(),
        markedAt: String(r[6] || '').trim()
      };
    });
  }

  return { photos: photos, naStatus: naStatus };
}

function getPhotoMetaFromDescription(file) {
  try {
    var desc = file.getDescription();
    if (!desc) return {};
    var meta = JSON.parse(desc);
    return {
      category: String(meta.category || '').trim(),
      subcategory: String(meta.subcategory || '').trim()
    };
  } catch(e) {
    return {};
  }
}

function buildSubcategorySlugMap(ss, projectName) {
  var result = {};
  var cm = ss.getSheetByName(SHEET_CAT_MASTER);
  var pc = ss.getSheetByName(SHEET_PROJ_CATS);
  if (!cm || !pc || cm.getLastRow() < 2 || pc.getLastRow() < 2) return result;

  var masterMap = {};
  cm.getRange(2, 1, cm.getLastRow() - 1, 3).getValues().forEach(function(r) {
    if (!r[0]) return;
    masterMap[String(r[0]).trim()] = {
      category: String(r[1] || '').trim(),
      subcategory: String(r[2] || '').trim()
    };
  });

  pc.getRange(2, 1, pc.getLastRow() - 1, 4).getValues().forEach(function(r) {
    var rowProject = String(r[0] || '').trim();
    var catId = String(r[1] || '').trim();
    var active = String(r[3] || 'TRUE').toUpperCase();
    if (rowProject !== projectName || active === 'FALSE' || !masterMap[catId]) return;
    var category = masterMap[catId].category;
    var subcategory = masterMap[catId].subcategory;
    if (!result[category]) result[category] = {};
    result[category][slugForPhotoName(subcategory)] = subcategory;
  });
  return result;
}

function inferSubcategoryFromPhotoName(fileName, siteNumber, category, slugMap) {
  var dot = fileName.lastIndexOf('.');
  var base = dot >= 0 ? fileName.substring(0, dot) : fileName;
  base = base.replace(/_\d{3}$/, '');

  var prefix = slugForPhotoName(siteNumber) + '_' + slugForPhotoName(category) + '_';
  if (base.indexOf(prefix) !== 0) return '';

  var subSlug = base.substring(prefix.length);
  return slugMap[subSlug] || subSlug.replace(/_/g, ' ');
}

function slugForPhotoName(str) {
  return String(str || '')
    .trim()
    .replace(/[^a-zA-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
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
  if (!sheet) return { success: false };
  if (data.project && !userCanAccessProject(ss, (data.userEmail || '').toLowerCase().trim(), data.project)) {
    return { success: false, message: 'User is not assigned to this project' };
  }

  var naDetails    = data.naDetails || [];
  var naDetailsStr = naDetails.length
    ? naDetails.map(function(n){ return n.category + ' > ' + n.subcategory; }).join('; ')
    : 'None';

  sheet.appendRow([
    'Site_' + (data.siteNumber || ''),
    data.project   || '',
    data.userName  || '',
    data.userEmail || '',
    data.date      || new Date().toISOString(),
    Number(data.totalPhotos) || 0,
    Number(data.totalNA)     || 0,
    'Completed',
    naDetailsStr
  ]);

  if (naDetails.length) {
    var naSheet = ss.getSheetByName(SHEET_NA_LOG);
    if (naSheet) {
      var existSet = {};
      var lastRow = naSheet.getLastRow();
      if (lastRow >= 2) {
        naSheet.getRange(2, 1, lastRow - 1, 4).getValues()
          .forEach(function(r) { existSet[r[0]+'|'+r[1]+'|'+r[2]+'|'+r[3]] = true; });
      }
      naDetails.forEach(function(n) {
        var k = 'Site_'+(data.siteNumber||'')+'|'+(data.project||'')+'|'+n.category+'|'+n.subcategory;
        if (!existSet[k]) {
          naSheet.appendRow([
            'Site_'+(data.siteNumber||''), data.project||'',
            n.category, n.subcategory,
            data.userName||'', data.userEmail||'',
            data.date||new Date().toISOString(), 'N/A'
          ]);
        }
      });
    }
  }
  SpreadsheetApp.flush();
  completeActiveSite(data.siteNumber || '', data.project || '');
  return { success: true };
}

function touchActiveSite(siteNumber, project) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_ACTIVE_SITES);
  if (!sheet || !project) return;
  var cleanSite = String(siteNumber || '').replace(/^Site_/, '').trim();
  var last = sheet.getLastRow();
  if (last < 2) return;
  var rows = sheet.getRange(2, 1, last - 1, 7).getValues();
  for (var i = 0; i < rows.length; i++) {
    var rowSite = String(rows[i][0] || '').replace(/^Site_/, '').trim();
    var rowProject = String(rows[i][1] || '').trim();
    var status = String(rows[i][5] || 'Open').toLowerCase();
    if (rowSite === cleanSite && rowProject === project && status === 'open') {
      sheet.getRange(i + 2, 7).setValue(new Date().toISOString());
      return;
    }
  }
}

function completeActiveSite(siteNumber, project) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_ACTIVE_SITES);
  if (!sheet || !project) return;
  var cleanSite = String(siteNumber || '').replace(/^Site_/, '').trim();
  var last = sheet.getLastRow();
  if (last < 2) return;
  var rows = sheet.getRange(2, 1, last - 1, 7).getValues();
  for (var i = 0; i < rows.length; i++) {
    var rowSite = String(rows[i][0] || '').replace(/^Site_/, '').trim();
    var rowProject = String(rows[i][1] || '').trim();
    var status = String(rows[i][5] || 'Open').toLowerCase();
    if (rowSite === cleanSite && rowProject === project && status === 'open') {
      sheet.getRange(i + 2, 6, 1, 2).setValues([['Completed', new Date().toISOString()]]);
      return;
    }
  }
}

// -----------------------------------------------
//  HELPER
// -----------------------------------------------
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getOrCreateFolder(parent, name) {
  var f = parent.getFoldersByName(name);
  if (f.hasNext()) return f.next();
  return parent.createFolder(name);
}
