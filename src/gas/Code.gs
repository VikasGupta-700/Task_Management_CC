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
    .addItem('Open Fast Dashboard', 'openFastDashboard')
    .addItem('Generate Standalone Dashboard', 'generateStandaloneDashboard')
    .addItem('Get Dashboard URLs', 'showDashboardUrls')
    .addToUi();
}

/**
 * Web app entry point for standalone dashboard access
 * Deploy this script as a Web App to get a direct URL to the dashboard
 */
function doGet() {
  console.log('=== doGet() START - Full Dashboard ===');
  
  try {
    console.log('Step 1: Starting doGet function for full dashboard');
    
    console.log('Step 2: Getting full dashboard data (with overdue details)');
    const dashboardData = getDashboardData();
    console.log('Step 3: Dashboard data retrieved, length:', dashboardData ? dashboardData.length : 'null');
    
    console.log('Step 4: Creating HTML template from dashboard file');
    const htmlTemplate = HtmlService.createTemplateFromFile('dashboard');
    console.log('Step 5: HTML template created successfully');
    
    console.log('Step 6: Assigning dashboard data to template');
    htmlTemplate.dashboardData = dashboardData;
    console.log('Step 7: Data assigned to template');
    
    console.log('Step 8: Evaluating template');
    const htmlOutput = htmlTemplate.evaluate();
    console.log('Step 9: Template evaluated successfully');
    
    console.log('Step 10: Setting title and options');
    htmlOutput.setTitle('Task Management Dashboard');
    htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    console.log('Step 11: Options configured');
    
    console.log('Step 12: Returning HTML output');
    console.log('=== doGet() SUCCESS - Full Dashboard ===');
    
    return htmlOutput;
      
  } catch (error) {
    console.error('=== doGet() ERROR ===');
    console.error('Error details:', error.toString());
    console.error('Error stack:', error.stack);
    
    // Return detailed error page
    return HtmlService.createHtmlOutput(`
      <html>
        <head>
          <title>Dashboard Error</title>
          <style>
            body { font-family: Arial, sans-serif; margin: 40px; background: #f5f5f5; }
            .error-container { background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
            .error-title { color: #d32f2f; font-size: 24px; margin-bottom: 20px; }
            .error-details { background: #f8f8f8; padding: 15px; border-radius: 4px; font-family: monospace; white-space: pre-wrap; }
            .success-note { background: #e8f5e8; padding: 15px; border-radius: 4px; color: #2e7d32; margin-bottom: 20px; }
          </style>
        </head>
        <body>
          <div class="error-container">
            <div class="success-note">
              ✅ Good News: Web App deployment is working! You're seeing this error page, which means the deployment infrastructure is functional.
            </div>
            <h1 class="error-title">Dashboard Loading Error</h1>
            <p><strong>Error Message:</strong></p>
            <div class="error-details">${error.toString()}</div>
            <p><strong>Error Stack:</strong></p>
            <div class="error-details">${error.stack || 'No stack trace available'}</div>
            <p><strong>Time:</strong> ${new Date().toLocaleString()}</p>
            <p><strong>Next Steps:</strong></p>
            <ul>
              <li>Check the Google Apps Script execution logs</li>
              <li>Verify the dashboard.html file exists</li>
              <li>Ensure getFastDashboardData() function works</li>
            </ul>
          </div>
        </body>
      </html>
    `);
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
      G: `=IFERROR(SUMPRODUCT(--(LEN(IMPORTRANGE(${cellRef},"'Tasks'!B2:B1000"))>0)),0)`,
      H: `=IFERROR(SUM(COUNTIF(IMPORTRANGE(${cellRef},"'Tasks'!G2:G"),{"Completed","Done","Closed","Complete"})),0)`,
      I: `=IF(G${rowIndex}=0,0, G${rowIndex}-H${rowIndex})`,
      J: `=IFERROR(COUNTIFS(IMPORTRANGE(${cellRef},"'Tasks'!E2:E"),"<"&TODAY(),IMPORTRANGE(${cellRef},"'Tasks'!G2:G"),"<>Completed"),0)`,
      K: `=IF(G${rowIndex}=0,0,ROUND(H${rowIndex}/G${rowIndex}*100,0))`
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
 * Opens the fast dashboard (menu-safe version without overdue task details)
 */
function openFastDashboard() {
  try {
    console.log('Opening fast dashboard...');
    
    // Get fast data (no overdue task details)
    const dashboardData = getFastDashboardData();
    
    // Create HTML template
    const htmlTemplate = HtmlService.createTemplateFromFile('dashboard');
    htmlTemplate.dashboardData = dashboardData;
    
    // Create output for modal dialog
    const htmlOutput = htmlTemplate.evaluate()
      .setTitle('Task Management Dashboard (Fast)')
      .setWidth(1400)
      .setHeight(900)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    // Show modal dialog
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Task Management Dashboard (Fast)');
    
  } catch (error) {
    console.error('Error opening fast dashboard:', error);
    SpreadsheetApp.getUi().alert('Error opening fast dashboard: ' + error.toString());
  }
}

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
 * Gets fast dashboard data for web app (skips overdue task details)
 * @return {Array} Array of dashboard card data
 */
function getFastDashboardData() {
  console.log('=== getFastDashboardData() START ===');
  
  try {
    console.log('Fast Data Step 1: Getting master sheet');
    const masterSheet = getSheetsMasterSheet();
    console.log('Fast Data Step 2: Master sheet retrieved');
    
    const lastRow = masterSheet.getLastRow();
    console.log('Fast Data Step 3: Last row:', lastRow);
    
    if (lastRow < 2) {
      console.log('Fast Data Step 4: No data rows, returning empty array');
      return [];
    }
    
    // Get all data from columns A-L
    console.log('Fast Data Step 5: Getting data range from A2 to L' + lastRow);
    const data = masterSheet.getRange(2, 1, lastRow - 1, 12).getValues();
    console.log('Fast Data Step 6: Data retrieved, rows:', data.length);
    
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
      
      // Extract values directly from columns F-L (NO overdue task details fetching)
      const cardData = {
        cardTitle: cardTitle,
        templateName: templateName,
        sharedWith: sharedWith,
        sheetUrl: row[3],                    // Column D
        dateCreated: row[5] || '',           // Column F
        totalTasks: row[6] || 0,            // Column G
        completedTasks: row[7] || 0,        // Column H  
        estimated: row[8] || 0,             // Column I
        onTrack: row[9] || 0,               // Column J
        overdueTasks: row[10] || 0,         // Column K
        upcoming: row[11] || 0,             // Column L
        progressPercent: row[7] && row[6] ? Math.round((row[7]/row[6])*100) : 0, // Calculate progress
        overdueTaskDetails: [],             // Empty for fast loading
        hasError: false,
        errorMessage: ''
      };
      
      result.push(cardData);
    }
    
    console.log('Fast Data Step 7: Processed', result.length, 'cards');
    console.log('=== getFastDashboardData() SUCCESS ===');
    return result;
    
  } catch (error) {
    console.error('=== getFastDashboardData() ERROR ===');
    console.error('Error getting fast dashboard data:', error);
    console.error('Error stack:', error.stack);
    return [];
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
      
      // Extract values directly from columns F-L (updated to include Date Created)
      const cardData = {
        cardTitle: cardTitle,
        templateName: templateName,
        sharedWith: sharedWith,
        sheetUrl: row[3],                    // Column D
        dateCreated: row[5] || '',           // Column F
        totalTasks: row[6] || 0,            // Column G
        completedTasks: row[7] || 0,        // Column H  
        estimated: row[8] || 0,             // Column I
        onTrack: row[9] || 0,               // Column J
        overdueTasks: row[10] || 0,         // Column K
        upcoming: row[11] || 0,             // Column L
        progressPercent: row[7] && row[6] ? Math.round((row[7]/row[6])*100) : 0, // Calculate progress
        overdueTaskDetails: [],             // Will be populated below
        hasError: false,
        errorMessage: ''
      };
      
      // Fetch overdue task details from individual sheet if overdue tasks > 0
      if (cardData.overdueTasks > 0 && cardData.sheetUrl) {
        try {
          const overdueDetails = getOverdueTaskDetails(cardData.sheetUrl);
          cardData.overdueTaskDetails = overdueDetails.tasks;
          cardData.hasError = overdueDetails.hasError;
          cardData.errorMessage = overdueDetails.errorMessage;
        } catch (error) {
          console.error(`Error fetching overdue details for ${cardData.cardTitle}:`, error);
          cardData.hasError = true;
          cardData.errorMessage = `Failed to fetch overdue task details: ${error.toString()}`;
        }
      }
      
      result.push(cardData);
    }
    
    return result;
    
  } catch (error) {
    console.error('Error getting dashboard data:', error);
    return [];
  }
}

/**
 * Fetches overdue task details from individual sheet's Dashboard tab
 * @param {string} sheetUrl - URL of the individual sheet
 * @return {Object} Object containing tasks array and error information
 */
function getOverdueTaskDetails(sheetUrl) {
  try {
    console.log(`Fetching overdue details from: ${sheetUrl}`);
    
    // Extract file ID from the URL
    const fileId = extractFileIdFromUrl(sheetUrl);
    
    // Open the spreadsheet
    const spreadsheet = SpreadsheetApp.openById(fileId);
    
    // Get the first sheet (main task sheet) or try common sheet names
    let taskSheet;
    try {
      // Try common sheet names first
      const commonNames = ['Tasks', 'Task', 'Sheet1', 'Main'];
      for (const name of commonNames) {
        try {
          taskSheet = spreadsheet.getSheetByName(name);
          if (taskSheet) {
            console.log(`Found task sheet: ${name}`);
            break;
          }
        } catch (e) {
          // Continue to next name
        }
      }
      
      // If no common name found, use the first sheet
      if (!taskSheet) {
        taskSheet = spreadsheet.getSheets()[0];
        console.log(`Using first sheet: ${taskSheet.getName()}`);
      }
    } catch (error) {
      console.error(`Error accessing sheets in ${sheetUrl}:`, error);
      return {
        tasks: [],
        hasError: true,
        errorMessage: 'Unable to access task sheet'
      };
    }
    
    if (!taskSheet) {
      return {
        tasks: [],
        hasError: true,
        errorMessage: 'No task sheet found'
      };
    }
    
    // Get all data from the sheet (assuming header in row 1)
    const lastRow = taskSheet.getLastRow();
    const lastCol = taskSheet.getLastColumn();
    
    if (lastRow < 2) {
      console.log(`No task data found in sheet`);
      return {
        tasks: [],
        hasError: false,
        errorMessage: 'No tasks found'
      };
    }
    
    // Get all task data (starting from row 2 to skip header)
    const taskData = taskSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    // Find overdue tasks by checking Status and Delay columns
    // Based on your data structure: 
    // Column B = Task Title, Column G = Status, Column J = Delay Days, Column I = On Time/Delayed
    const overdueTasks = [];
    
    for (let i = 0; i < taskData.length; i++) {
      const row = taskData[i];
      const taskTitle = row[1]?.toString().trim(); // Column B
      const status = row[6]?.toString().trim();    // Column G  
      const delayStatus = row[8]?.toString().trim(); // Column I (On Time/Delayed)
      const delayDays = row[9] || 0;               // Column J (Delay Days)
      
      // Check if task is overdue (has delay days > 0 and not completed)
      if (taskTitle && 
          delayStatus === 'Delayed' && 
          delayDays > 0 && 
          status !== 'Completed') {
        
        overdueTasks.push({
          taskTitle: taskTitle,
          delayDays: parseInt(delayDays) || 0
        });
      }
    }
    
    // Sort by delay days (highest first) and limit to top 3
    overdueTasks.sort((a, b) => b.delayDays - a.delayDays);
    const topThreeTasks = overdueTasks.slice(0, 3);
    
    console.log(`Found ${overdueTasks.length} overdue tasks, returning top ${topThreeTasks.length}`);
    
    return {
      tasks: topThreeTasks,
      hasError: false,
      errorMessage: ''
    };
    
  } catch (error) {
    console.error(`Error accessing sheet ${sheetUrl}:`, error);
    return {
      tasks: [],
      hasError: true,
      errorMessage: `Unable to access sheet: ${error.toString()}`
    };
  }
}

// =============================================================================
// STANDALONE DASHBOARD GENERATOR
// =============================================================================

/**
 * Generates a standalone HTML dashboard file that can be hosted anywhere
 */
function generateStandaloneDashboard() {
  try {
    console.log('Generating standalone dashboard...');
    
    // Get dashboard data
    const dashboardData = getFastDashboardData();
    
    // Read dashboard template
    const dashboardTemplate = HtmlService.createTemplateFromFile('dashboard');
    dashboardTemplate.dashboardData = dashboardData;
    
    // Generate HTML content
    const htmlContent = dashboardTemplate.evaluate().getContent();
    
    // Create standalone HTML file in Google Drive
    const fileName = `Task_Management_Dashboard_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss')}.html`;
    
    // Get or create Dashboards folder
    const dashboardsFolder = getOrCreateFolder('Generated_Dashboards');
    
    // Create HTML file
    const htmlFile = dashboardsFolder.createFile(fileName, htmlContent, MimeType.HTML);
    
    // Make file publicly viewable
    htmlFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const viewUrl = `https://drive.google.com/file/d/${htmlFile.getId()}/view`;
    const downloadUrl = `https://drive.google.com/uc?export=download&id=${htmlFile.getId()}`;
    
    // Show success message with URLs
    const ui = SpreadsheetApp.getUi();
    const message = `Standalone Dashboard Generated Successfully!\n\n` +
                   `File: ${fileName}\n\n` +
                   `View URL: ${viewUrl}\n\n` +
                   `Download URL: ${downloadUrl}\n\n` +
                   `The HTML file is saved in your Drive under "Generated_Dashboards" folder.\n` +
                   `You can download and host it on any web server.`;
                   
    ui.alert('Dashboard Generated!', message, ui.ButtonSet.OK);
    
    console.log('Standalone dashboard generated successfully:', fileName);
    
    return {
      success: true,
      fileName: fileName,
      fileId: htmlFile.getId(),
      viewUrl: viewUrl,
      downloadUrl: downloadUrl
    };
    
  } catch (error) {
    console.error('Error generating standalone dashboard:', error);
    SpreadsheetApp.getUi().alert('Error generating dashboard: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Generates a self-contained HTML dashboard with embedded data (no external calls)
 */
function generateSelfContainedDashboard() {
  try {
    const dashboardData = getFastDashboardData();
    
    // Create self-contained HTML with embedded data
    const htmlContent = `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Task Management Dashboard</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; }
        .header { text-align: center; margin-bottom: 30px; }
        .cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; }
        .card { 
            background: white; 
            border-radius: 8px; 
            padding: 20px; 
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            border-left: 4px solid #4285f4;
        }
        .card-title { font-size: 18px; font-weight: bold; margin-bottom: 15px; color: #333; }
        .metric { display: flex; justify-content: space-between; margin: 8px 0; }
        .metric-label { color: #666; }
        .metric-value { font-weight: bold; color: #333; }
        .progress-bar { 
            width: 100%; 
            height: 8px; 
            background-color: #e0e0e0; 
            border-radius: 4px; 
            overflow: hidden; 
            margin: 10px 0;
        }
        .progress-fill { 
            height: 100%; 
            background-color: #4285f4; 
            transition: width 0.3s ease;
        }
        .overdue { color: #d32f2f; font-weight: bold; }
        .completed { color: #388e3c; font-weight: bold; }
        .timestamp { text-align: center; color: #666; font-size: 12px; margin-top: 20px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 Task Management Dashboard</h1>
            <p>Real-time view of your task management system</p>
        </div>
        
        <div class="cards" id="dashboard-cards">
            ${dashboardData.map(card => `
                <div class="card">
                    <div class="card-title">${card.cardTitle}</div>
                    <div class="metric">
                        <span class="metric-label">Template:</span>
                        <span class="metric-value">${card.templateName}</span>
                    </div>
                    <div class="metric">
                        <span class="metric-label">Shared With:</span>
                        <span class="metric-value">${card.sharedWith}</span>
                    </div>
                    <div class="metric">
                        <span class="metric-label">Total Tasks:</span>
                        <span class="metric-value">${card.totalTasks}</span>
                    </div>
                    <div class="metric">
                        <span class="metric-label">Completed:</span>
                        <span class="metric-value completed">${card.completedTasks}</span>
                    </div>
                    <div class="metric">
                        <span class="metric-label">Overdue:</span>
                        <span class="metric-value overdue">${card.overdueTasks}</span>
                    </div>
                    <div class="progress-bar">
                        <div class="progress-fill" style="width: ${card.progressPercent}%"></div>
                    </div>
                    <div class="metric">
                        <span class="metric-label">Progress:</span>
                        <span class="metric-value">${card.progressPercent}%</span>
                    </div>
                    ${card.sheetUrl ? `
                    <div class="metric">
                        <span class="metric-label">Sheet:</span>
                        <span class="metric-value"><a href="${card.sheetUrl}" target="_blank">Open Sheet</a></span>
                    </div>
                    ` : ''}
                </div>
            `).join('')}
        </div>
        
        <div class="timestamp">
            Generated: ${new Date().toLocaleString()}
        </div>
    </div>
</body>
</html>`;

    return htmlContent;
    
  } catch (error) {
    console.error('Error generating self-contained dashboard:', error);
    return `<html><body><h1>Error</h1><p>${error.toString()}</p></body></html>`;
  }
}

// =============================================================================
// UTILITY FUNCTIONS
// =============================================================================

/**
 * Shows URLs for the most recent dashboard files
 */
function showDashboardUrls() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Get Generated_Dashboards folder
    const dashboardsFolder = getOrCreateFolder('Generated_Dashboards');
    const files = dashboardsFolder.getFiles();
    
    if (!files.hasNext()) {
      ui.alert('No Dashboard Files Found', 'Please generate a dashboard first using "Generate Standalone Dashboard"', ui.ButtonSet.OK);
      return;
    }
    
    // Get the most recent dashboard file
    let mostRecentFile = null;
    let mostRecentDate = new Date(0);
    
    while (files.hasNext()) {
      const file = files.next();
      if (file.getName().includes('Task_Management_Dashboard') && file.getDateCreated() > mostRecentDate) {
        mostRecentFile = file;
        mostRecentDate = file.getDateCreated();
      }
    }
    
    if (!mostRecentFile) {
      ui.alert('No Dashboard Files Found', 'Please generate a dashboard first using "Generate Standalone Dashboard"', ui.ButtonSet.OK);
      return;
    }
    
    const fileId = mostRecentFile.getId();
    const fileName = mostRecentFile.getName();
    
    const downloadUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;
    const viewUrl = `https://drive.google.com/file/d/${fileId}/view`;
    
    const message = `Latest Dashboard File: ${fileName}\n\n` +
                   `📥 DOWNLOAD URL (recommended):\n${downloadUrl}\n\n` +
                   `👁️ VIEW URL (shows code only):\n${viewUrl}\n\n` +
                   `🌐 TO VIEW AS DASHBOARD:\n` +
                   `1. Click download URL above\n` +
                   `2. Save the HTML file to your computer\n` +
                   `3. Double-click the file to open in browser\n\n` +
                   `📤 FOR FREE HOSTING:\n` +
                   `• Upload to GitHub Pages\n` +
                   `• Upload to Netlify\n` +
                   `• Upload to Vercel\n` +
                   `• Use any free HTML hosting service`;
    
    ui.alert('Dashboard URLs Ready!', message, ui.ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error getting URLs: ' + error.toString());
  }
}

/**
 * Include function for HTML templates
 * @param {string} filename - Name of the HTML file to include
 * @return {string} Content of the HTML file
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}