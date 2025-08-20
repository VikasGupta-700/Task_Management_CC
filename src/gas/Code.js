/**
 * Enhanced Task Management System
 * Features:
 * - Create Single Sheet from Task Master
 * - Bulk Create Pending Sheets 
 * - Enhanced Dashboard with Charts
 * - Comprehensive Error Handling
 */

// =============================================================================
// MENU SYSTEM
// =============================================================================

/**
 * Creates the main menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Task Management System')
    .addItem('Create Sheets', 'createSheets')
    .addSeparator()
    .addItem('Open Dashboard', 'openDashboard')
    .addToUi();
}

/**
 * Web app entry point for standalone dashboard access
 * Deploy this script as a Web App to get a direct URL to the dashboard
 */
function doGet() {
  try {
    const dashboardData = getDashboardData();
    const htmlTemplate = HtmlService.createTemplateFromFile('dashboard');
    htmlTemplate.dashboardData = dashboardData;
    
    return htmlTemplate.evaluate()
      .setTitle('Task Management Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
  } catch (error) {
    return HtmlService.createHtmlOutput('<h1>Error loading dashboard</h1><p>' + error.toString() + '</p>');
  }
}

// =============================================================================
// SHEET CREATION
// =============================================================================

/**
 * Creates sheets for all rows in Sheets_Master where Column E is empty
 * Shows summary popup upon completion
 */
function createSheets() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Confirm before proceeding
    const response = ui.alert(
      'Bulk Create Confirmation',
      'This will create sheets for all rows where status (Column E) is empty.\\n\\nProceed?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    const masterSheet = getSheetsMasterSheet();
    const lastRow = masterSheet.getLastRow();
    
    if (lastRow < 2) {
      ui.alert('No data rows found in Sheets_Master sheet.');
      return;
    }
    
    // Get all data at once for efficiency
    const allData = masterSheet.getRange(2, 1, lastRow - 1, 12).getValues();
    
    let sheetsCreated = 0;
    let sheetsSkipped = 0;
    let errors = 0;
    
    for (let i = 0; i < allData.length; i++) {
      const rowIndex = i + 2; // Actual row number in sheet
      const rowData = allData[i];
      
      // Skip if Column E (status) is not empty
      if (rowData[4]) { // Column E
        sheetsSkipped++;
        continue;
      }
      
      // Skip if no template name in Column A
      if (!rowData[0]) {
        sheetsSkipped++;
        continue;
      }
      
      const result = processSheetCreation(rowIndex, rowData);
      
      if (result.success) {
        sheetsCreated++;
      } else {
        errors++;
      }
      
      // Add small delay to prevent hitting quotas
      Utilities.sleep(100);
    }
    
    // Show summary popup
    showBulkCreationSummary(sheetsCreated, sheetsSkipped, errors);
    
  } catch (error) {
    console.error('Error in createSheets:', error);
    SpreadsheetApp.getUi().alert('Unexpected error: ' + error.toString());
  }
}

/**
 * Shows the bulk creation summary popup
 * @param {number} sheetsCreated - Number of sheets successfully created
 * @param {number} sheetsSkipped - Number of sheets skipped
 * @param {number} errors - Number of errors encountered
 */
function showBulkCreationSummary(sheetsCreated, sheetsSkipped, errors) {
  const ui = SpreadsheetApp.getUi();
  
  const message = 'Bulk Creation Summary\\n' +
                 '-----------------------\\n' +
                 'Sheets Created: ' + sheetsCreated + '\\n' +
                 'Sheets Skipped: ' + sheetsSkipped + '\\n' +
                 'Errors: ' + errors;
  
  ui.alert('Bulk Creation Complete', message, ui.ButtonSet.OK);
}

// =============================================================================
// CORE SHEET PROCESSING
// =============================================================================

/**
 * Processes the creation of a single sheet for a given row
 * @param {number} rowIndex - The row number in the Sheets_Master sheet
 * @param {Array} rowData - The data from the row
 * @return {Object} Result object with success flag and details
 */
function processSheetCreation(rowIndex, rowData) {
  try {
    const templateName = rowData[0]; // Column A
    const sharedWith = rowData[1];   // Column B
    const emailId = rowData[2];      // Column C
    
    // Get template URL from Template_List sheet
    const templateUrl = getTemplateUrl(templateName);
    if (!templateUrl) {
      updateRowStatus(rowIndex, 'Error: Template not found in Template_List');
      return { success: false, error: 'Template "' + templateName + '" not found in Template_List sheet' };
    }
    
    // Create the sheet copy
    const sheetResult = createSheetCopy(templateName, sharedWith, templateUrl);
    if (!sheetResult.success) {
      updateRowStatus(rowIndex, 'Error: ' + sheetResult.error);
      return sheetResult;
    }
    
    // Update Task Master row with results
    updateTaskMasterRow(rowIndex, sheetResult.sheetUrl);
    
    // Share with email if provided
    if (emailId && emailId.includes('@')) {
      shareSheetWithUser(sheetResult.sheetId, emailId);
    }
    
    // Insert IMPORTRANGE formulas
    insertImportRangeFormulas(rowIndex, sheetResult.sheetUrl);
    
    return {
      success: true,
      sheetName: sheetResult.sheetName,
      sheetUrl: sheetResult.sheetUrl,
      sheetId: sheetResult.sheetId
    };
    
  } catch (error) {
    console.error('Error in processSheetCreation:', error);
    updateRowStatus(rowIndex, 'Error: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Creates a copy of the template sheet
 * @param {string} templateName - Name of the template
 * @param {string} sharedWith - Name of the person to share with
 * @param {string} templateUrl - URL of the template sheet
 * @return {Object} Result object with sheet details
 */
function createSheetCopy(templateName, sharedWith, templateUrl) {
  try {
    // Extract file ID from template URL
    const templateId = extractFileIdFromUrl(templateUrl);
    const templateFile = DriveApp.getFileById(templateId);
    
    // Generate new sheet name
    const firstWordTemplate = templateName.split(' ')[0] || 'Template';
    const firstWordShared = (sharedWith.split(' ')[0] || 'User').split('@')[0];
    const sheetName = firstWordTemplate + '_sheet_' + firstWordShared;
    
    // Get or create Created_Sheets folder
    const createdSheetsFolder = getOrCreateFolder('Created_Sheets');
    
    // Create the copy
    const newFile = templateFile.makeCopy(sheetName, createdSheetsFolder);
    const newSpreadsheet = SpreadsheetApp.openById(newFile.getId());
    
    return {
      success: true,
      sheetName: sheetName,
      sheetUrl: newSpreadsheet.getUrl(),
      sheetId: newFile.getId(),
      spreadsheet: newSpreadsheet
    };
    
  } catch (error) {
    console.error('Error creating sheet copy:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Updates the Sheets_Master row with sheet creation results
 * @param {number} rowIndex - Row number to update
 * @param {string} sheetUrl - URL of the created sheet
 */
function updateTaskMasterRow(rowIndex, sheetUrl) {
  try {
    const masterSheet = getSheetsMasterSheet();
    const currentDate = new Date();
    const formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'dd-MMM-yy');
    
    // Update columns D, E, F
    masterSheet.getRange(rowIndex, 4).setValue(sheetUrl);                    // Column D: Sheet URL
    masterSheet.getRange(rowIndex, 5).setValue('Sheet created successfully'); // Column E: Status
    masterSheet.getRange(rowIndex, 6).setValue(formattedDate);               // Column F: Date Created
    
  } catch (error) {
    console.error('Error updating Task Master row:', error);
    throw error;
  }
}

/**
 * Updates only the status column (E) for a row
 * @param {number} rowIndex - Row number to update
 * @param {string} status - Status message to set
 */
function updateRowStatus(rowIndex, status) {
  try {
    const masterSheet = getSheetsMasterSheet();
    masterSheet.getRange(rowIndex, 5).setValue(status); // Column E
  } catch (error) {
    console.error('Error updating row status:', error);
  }
}

/**
 * Inserts IMPORTRANGE formulas into columns G-L for the specified row
 * @param {number} rowIndex - Row number to update
 * @param {string} sheetUrl - URL of the target sheet
 */
function insertImportRangeFormulas(rowIndex, sheetUrl) {
  try {
    const masterSheet = getSheetsMasterSheet();
    const cellRef = 'D' + rowIndex; // Reference to the URL cell
    
    // Define formulas as per specifications
    const formulas = {
      G: `=IFERROR(COUNTA(IMPORTRANGE(${cellRef},"'Tasks'!B2:B"))-1,0)`,
      H: `=IFERROR(SUM(COUNTIF(IMPORTRANGE(${cellRef},"'Tasks'!G2:G"),{"Completed","Done","Closed","Complete"})),0)`,
      I: `=IF(G${rowIndex}=0,0, G${rowIndex}-H${rowIndex})`,
      J: `=IFERROR(COUNTIFS(IMPORTRANGE(${cellRef},"'Tasks'!E2:E"),"<"&TODAY(),IMPORTRANGE(${cellRef},"'Tasks'!G2:G"),"<>Completed"),0)`,
      K: `=IFERROR(SUM(COUNTIF(IMPORTRANGE(${cellRef},"'Tasks'!K2:K"),{"No"})),0)`,
      L: `=IF(G${rowIndex}=0,0,ROUND(H${rowIndex}/G${rowIndex}*100,0))`
    };
    
    // Insert each formula
    Object.keys(formulas).forEach(col => {
      const colIndex = col.charCodeAt(0) - 'A'.charCodeAt(0) + 1;
      masterSheet.getRange(rowIndex, colIndex).setFormula(formulas[col]);
    });
    
  } catch (error) {
    console.error('Error inserting IMPORTRANGE formulas:', error);
    throw error;
  }
}

// =============================================================================
// HELPER FUNCTIONS
// =============================================================================

/**
 * Gets the Sheets_Master sheet, creates it if it doesn't exist
 * @return {Sheet} The Sheets_Master sheet
 */
function getSheetsMasterSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName('Sheets_Master');
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet('Sheets_Master');
  }
  
  return sheet;
}


/**
 * Gets the template URL from the Template_List sheet
 * @param {string} templateName - Name of the template to find
 * @return {string|null} Template URL or null if not found
 */
function getTemplateUrl(templateName) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const templateSheet = spreadsheet.getSheetByName('Template_List');
    
    if (!templateSheet) {
      console.error('Template_List sheet not found');
      return null;
    }
    
    const data = templateSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) { // Skip header row
      if (data[i][0] && data[i][0].toString().trim() === templateName.trim()) {
        return data[i][1]; // Return URL from column B
      }
    }
    
    return null;
    
  } catch (error) {
    console.error('Error getting template URL:', error);
    return null;
  }
}

/**
 * Extracts file ID from Google Sheets URL
 * @param {string} url - Google Sheets URL
 * @return {string} File ID
 */
function extractFileIdFromUrl(url) {
  // Try multiple patterns to extract Google Sheets ID
  
  // Pattern 1: /spreadsheets/d/{ID}/
  let match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match) {
    return match[1];
  }
  
  // Pattern 2: /open?id={ID}
  match = url.match(/[?&]id=([a-zA-Z0-9-_]+)/);
  if (match) {
    return match[1];
  }
  
  // Pattern 3: Look for any Google Drive file ID pattern (44 characters)
  match = url.match(/[a-zA-Z0-9-_]{25,}/);
  if (match) {
    return match[0];
  }
  
  throw new Error('Invalid Google Sheets URL: ' + url);
}

/**
 * Gets or creates a folder with the specified name
 * @param {string} folderName - Name of the folder
 * @return {Folder} The folder object
 */
function getOrCreateFolder(folderName) {
  try {
    // Try to find existing folder in the same location as the current spreadsheet
    const currentFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
    const parentFolders = currentFile.getParents();
    
    let parentFolder = DriveApp.getRootFolder();
    if (parentFolders.hasNext()) {
      parentFolder = parentFolders.next();
    }
    
    const existingFolders = parentFolder.getFoldersByName(folderName);
    if (existingFolders.hasNext()) {
      return existingFolders.next();
    }
    
    // Create new folder
    return parentFolder.createFolder(folderName);
    
  } catch (error) {
    console.error('Error getting/creating folder:', error);
    throw error;
  }
}

/**
 * Shares the created sheet with the specified email
 * @param {string} fileId - ID of the file to share
 * @param {string} email - Email address to share with
 */
function shareSheetWithUser(fileId, email) {
  try {
    const file = DriveApp.getFileById(fileId);
    file.addEditor(email);
  } catch (error) {
    console.error('Error sharing sheet with user:', error);
    // Don't throw error - sharing failure shouldn't stop the process
  }
}

// =============================================================================
// ENHANCED DASHBOARD
// =============================================================================

/**
 * Opens the enhanced dashboard in a new browser window
 */
function openDashboard() {
  try {
    console.log('Opening enhanced dashboard...');
    
    // Get data from Sheets_Master sheet
    const dashboardData = getDashboardData();
    
    // Create HTML template
    const htmlTemplate = HtmlService.createTemplateFromFile('dashboard');
    htmlTemplate.dashboardData = dashboardData;
    
    // Create output for new window with visible URL
    const htmlOutput = htmlTemplate.evaluate()
      .setTitle('Task Management Dashboard')
      .setWidth(1400)
      .setHeight(900)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    // Open in new window - this will open as a separate browser window with URL visible
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Task Management Dashboard');
    
    // Also provide option to open as web app (user needs to deploy)
    try {
      const scriptUrl = ScriptApp.getService().getUrl();
      if (scriptUrl) {
        SpreadsheetApp.getUi().alert('Dashboard opened!\\n\\nFor full browser window with URL bar, deploy this script as a Web App and access: ' + scriptUrl);
      }
    } catch (e) {
      // Ignore if web app not deployed
    }
    
  } catch (error) {
    console.error('Error opening dashboard:', error);
    SpreadsheetApp.getUi().alert('Error opening dashboard: ' + error.toString());
  }
}

/**
 * Gets dashboard data directly from Sheets_Master columns G-L
 * @return {Array} Array of dashboard card data
 */
function getDashboardData() {
  try {
    const masterSheet = getSheetsMasterSheet();
    const lastRow = masterSheet.getLastRow();
    
    if (lastRow < 2) {
      return [];
    }
    
    // Get all data from columns A-L
    const data = masterSheet.getRange(2, 1, lastRow - 1, 12).getValues();
    const result = [];
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      
      // Skip rows without template name or sheet URL
      if (!row[0] || !row[3]) continue;
      
      // Create card title from Column A and B
      const templateName = row[0].toString();
      const sharedWith = row[1].toString();
      const firstWordTemplate = templateName.split(' ')[0] || 'Template';
      const firstWordShared = (sharedWith.split(' ')[0] || 'User').split('@')[0];
      const cardTitle = firstWordTemplate + ' Sheet ' + firstWordShared;
      
      // Extract values directly from columns G-L
      const cardData = {
        cardTitle: cardTitle,
        templateName: templateName,
        sheetUrl: row[3],                    // Column D
        totalTasks: row[6] || 0,            // Column G
        completedTasks: row[7] || 0,        // Column H  
        pendingTasks: row[8] || 0,          // Column I
        overdueTasks: row[9] || 0,          // Column J
        notEstimated: row[10] || 0,         // Column K
        progressPercent: row[11] || 0       // Column L (already as percentage)
      };
      
      result.push(cardData);
    }
    
    return result;
    
  } catch (error) {
    console.error('Error getting dashboard data:', error);
    return [];
  }
}

// =============================================================================
// UTILITY FUNCTIONS
// =============================================================================

/**
 * Include function for HTML templates
 * @param {string} filename - Name of the HTML file to include
 * @return {string} Content of the HTML file
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}