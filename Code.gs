/********************************************************
 * Task Manager - Code.gs
 * - Inserts formulas into Sheets_Master (G:K) using IMPORTRANGE
 * - Creates sheet copies in Created_Sheets
 * - Does NOT compute metrics server-side (formulas do)
 * - Menu: Generate Sheets, Ensure Formulas, Force Reinsert, Open Dashboard
 ********************************************************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Task Manager")
    .addItem("Generate Sheets", "processSheetsMaster")
    .addItem("Ensure Formulas / Update Metrics", "updateMetrics")
    .addItem("Force Reinsert Formulas", "forceReinsertFormulas")
    .addItem("Open Dashboard (Sidebar)", "showDashboardSidebar")
    .addToUi();
}

/** Standalone web app entry (deploy as Web App to use this URL) */
function doGet() {
  return HtmlService.createHtmlOutputFromFile("dashboard")
    .setTitle("Task Dashboard");
}

/**
 * Create copies for rows that have template name but no URL.
 * After copy, write D (URL), E (status), F (date), and insert formulas G:K (IMPORTRANGE).
 */
function processSheetsMaster() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName("Sheets_Master");
  var templateList = ss.getSheetByName("Template_List");
  if (!master || !templateList) {
    SpreadsheetApp.getUi().alert("Missing Sheets_Master or Template_List.");
    return;
  }

  ensureHeaders(master);

  // Build template map
  var tRows = templateList.getDataRange().getValues();
  var tMap = {};
  for (var i = 1; i < tRows.length; i++) {
    var nm = (tRows[i][0] || "").toString().trim();
    var url = (tRows[i][1] || "").toString().trim();
    if (nm && url) tMap[nm] = url;
  }

  // Parent folder and Created_Sheets
  var ssFile = DriveApp.getFileById(ss.getId());
  var parentFolder = ssFile.getParents().hasNext() ? ssFile.getParents().next() : DriveApp.getRootFolder();
  var createdFolder = getOrCreateSubFolder(parentFolder, "Created_Sheets");

  var mData = master.getDataRange().getValues();
  for (var r = 1; r < mData.length; r++) {
    var templateName = (mData[r][0] || "").toString().trim();
    var sharedWith = (mData[r][1] || "").toString().trim();
    var email = (mData[r][2] || "").toString().trim();
    var sheetUrl = (mData[r][3] || "").toString().trim();

    if (!templateName || sheetUrl) continue; // create only when template present and URL blank

    try {
      var srcUrl = tMap[templateName];
      if (!srcUrl) {
        master.getRange(r + 1, 5).setValue("Template not found"); // E
        continue;
      }
      var srcId = extractFileId(srcUrl);
      var srcFile = DriveApp.getFileById(srcId);

      // Build new name: FirstWordOfTemplate_Sheet_FirstName (spaces -> underscores)
      var firstWord = (templateName.split(/\s+/)[0] || "Template");
      var firstName = (sharedWith.split(/\s+/)[0] || "User");
      var newName = (firstWord + " Sheet " + firstName).replace(/\s+/g, "_");

      var newFile = srcFile.makeCopy(newName, createdFolder);
      var newId = newFile.getId();
      var newSS = SpreadsheetApp.openById(newId);

      if (email) {
        try { newFile.addEditor(email); } catch(e) { /* ignore share errors */ }
      }

      // Write URL, status, created date
      var newUrl = newSS.getUrl();
      master.getRange(r + 1, 4).setValue(newUrl); // D
      master.getRange(r + 1, 5).setValue("Sheet created successfully"); // E
      master.getRange(r + 1, 6).setValue(new Date()); // F

      // Detect the likely tasks sheet name in child and insert formulas (won't overwrite existing formulas)
      var tasksSheetName = detectTasksSheetName(newSS);
      setFormulasForRow(master, r + 1, tasksSheetName, false /*force*/);

    } catch (err) {
      master.getRange(r + 1, 5).setValue("Error: " + err.message); // E
    }
  }
}

/**
 * Insert formulas for every row that has a Sheet URL (without overwriting existing formulas unless forced).
 * Default: does not overwrite.
 * To force overwrite call updateMetrics(true) or use menu item Force Reinsert Formulas.
 */
function updateMetrics(force) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName("Sheets_Master");
  if (!master) { SpreadsheetApp.getUi().alert("Missing Sheets_Master"); return; }

  ensureHeaders(master);

  var mData = master.getDataRange().getValues();
  for (var r = 1; r < mData.length; r++) {
    var url = (mData[r][3] || "").toString().trim();
    if (!url) continue;
    try {
      var id = extractFileId(url);
      var childSS = SpreadsheetApp.openById(id);
      var tasksSheetName = detectTasksSheetName(childSS);
      setFormulasForRow(master, r + 1, tasksSheetName, !!force);
    } catch (err) {
      master.getRange(r + 1, 5).setValue("Error: " + err.message);
    }
  }
}

/** wrapper for menu to force reinsert */
function forceReinsertFormulas() {
  updateMetrics(true);
}

/** Sidebar view */
function showDashboardSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("dashboard")
    .setTitle("Task Dashboard")
    .setWidth(1100)
    .setHeight(800);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Responds to dashboard.html client with the evaluated values from Sheets_Master.
 * Returns array of objects for rendering.
 */
function getMasterData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName("Sheets_Master");
  if (!master) return [];
  var data = master.getDataRange().getValues();
  var rows = [];
  for (var r = 1; r < data.length; r++) {
    if (!data[r][0]) continue; // skip blank template name
    rows.push({
      template: data[r][0],
      sharedWith: data[r][1],
      email: data[r][2],
      url: data[r][3],
      status: data[r][4],
      createdAt: data[r][5],
      total: safeNumber(data[r][6]),
      completed: safeNumber(data[r][7]),
      pending: safeNumber(data[r][8]),
      overdue: safeNumber(data[r][9]),
      progress: safeNumber(data[r][10])
    });
  }
  return rows;
}

/* ---------------- Helpers ---------------- */

function setFormulasForRow(masterSheet, rowIndex, tasksSheetName, force) {
  // masterSheet: Sheet object
  // rowIndex: 1-based
  // tasksSheetName: name inside child sheet
  // force: if true, overwrite any existing formula/value in G:K

  if (!tasksSheetName) tasksSheetName = "Tasks";
  var safeName = tasksSheetName.replace(/'/g, "\\'"); // escape single quotes in sheet name

  // ranges inside child sheet
  var titleRange = "'" + safeName + "'!B2:B";  // Title column (B)
  var statusRange = "'" + safeName + "'!G2:G"; // Status column (G)
  var plannedRange = "'" + safeName + "'!E2:E"; // Planned date (E)
  var estimatedRange = "'" + safeName + "'!K2:K"; // Estimated? (K)

  var dRef = "D" + rowIndex; // cell in master containing child URL

  // Formulas (they rely on IMPORTRANGE referencing the child URL cell D#)
  var fTotal = '=IFERROR(COUNTA(IMPORTRANGE(' + dRef + ',"' + titleRange + '")),0)';
  var fCompleted = '=IFERROR(SUM(COUNTIF(IMPORTRANGE(' + dRef + ',"' + statusRange + '"),{"Completed","Done","Closed","Complete"})),0)';
  var fOverdue = '=IFERROR(COUNTIFS(IMPORTRANGE(' + dRef + ',"' + plannedRange + '"),"<"&TODAY(),IMPORTRANGE(' + dRef + ',"' + statusRange + '"),"<>Completed"),0)';
  var fPending = '=IF(' + 'G' + rowIndex + '=0,0, G' + rowIndex + '-H' + rowIndex + '-J' + rowIndex + ')';
  var fProgress = '=IF(G' + rowIndex + '=0,0,ROUND(H' + rowIndex + '/G' + rowIndex + '*100,0))';

  var colMap = {
    totalCol: 7,     // G
    completedCol: 8, // H
    pendingCol: 9,   // I
    overdueCol: 10,  // J
    progressCol: 11  // K
  };

  // helper to set formula only if blank or force = true
  function setIfBlankOrForce(r, c, formula) {
    var cell = masterSheet.getRange(r, c);
    var existingFormula = cell.getFormula();
    var existingValue = cell.getValue();

    if (force) {
      cell.setFormula(formula);
    } else {
      if (!existingFormula && (existingValue === "" || existingValue === null)) {
        cell.setFormula(formula);
      }
      // if existingFormula present, leave it alone
      // if existingValue present but no formula, do not overwrite (user prefers formulas)
    }
  }

  setIfBlankOrForce(rowIndex, colMap.totalCol, fTotal);
  setIfBlankOrForce(rowIndex, colMap.completedCol, fCompleted);
  setIfBlankOrForce(rowIndex, colMap.overdueCol, fOverdue);
  setIfBlankOrForce(rowIndex, colMap.pendingCol, fPending);
  setIfBlankOrForce(rowIndex, colMap.progressCol, fProgress);
}

/** Detects the likely tasks sheet inside a child spreadsheet */
function detectTasksSheetName(childSS) {
  var sheets = childSS.getSheets();
  var patterns = [/sr\.?\s*no/i, /\btask\b/i, /\bidea\b/i, /\bstatus\b/i, /\ballocated\b/i, /\bplanned\b/i];
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    try {
      var header = sh.getRange(1, 1, 1, Math.min(20, sh.getLastColumn())).getValues()[0].join(" ");
      for (var p = 0; p < patterns.length; p++) {
        if (patterns[p].test(header)) return sh.getName();
      }
    } catch (e) {
      // ignore read errors and try next
    }
  }
  return sheets.length ? sheets[0].getName() : "Tasks";
}

function extractFileId(url) {
  var m = (url || "").match(/[-\w]{25,}/);
  if (!m) throw new Error("Cannot extract file ID from URL: " + url);
  return m[0];
}

function getOrCreateSubFolder(parentFolder, name) {
  var it = parentFolder.getFoldersByName(name);
  return it.hasNext() ? it.next() : parentFolder.createFolder(name);
}

function ensureHeaders(master) {
  // Ensure E/F and G:K headers exist. G:K are always set (safe to overwrite header row)
  var lastCol = Math.max(master.getLastColumn(), 11);
  var headers = master.getRange(1, 1, 1, lastCol).getValues()[0];

  if (!headers[4]) master.getRange(1, 5).setValue("Sheet Status"); // E
  if (!headers[5]) master.getRange(1, 6).setValue("Date Created"); // F

  master.getRange(1, 7, 1, 5).setValues([["Total tasks","Completed tasks","Pending tasks","Overdue tasks","Progress %"]]); // G:K
}

function safeNumber(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  var s = String(v).replace(/[^0-9.\-]/g, "");
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}
