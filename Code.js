function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Task & Idea Manager')
    .addItem('Create Idea Template', 'createIdeaTemplate')
    .addItem('Create Task Template', 'createTaskTemplate')
    .addItem('Open Dashboard', 'openDashboard')
    .addToUi();
}

function createIdeaTemplate() {
  try {
    const newSpreadsheet = SpreadsheetApp.create('Idea Template - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));
    const sheet = newSpreadsheet.getActiveSheet();
    sheet.setName('Ideas');
    
    // Set up headers
    const headers = [
      'Sr. No', 'Idea Title', 'Idea Description', 'Idea Date', 
      'Planned Implementation Date', 'Actual Implementation Date', 
      'Status', 'Remarks / Issues', 'On Time / Delayed', 
      'Delay Days', 'Estimated?', 'Days Since Allocated'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.getRange(1, 1, 1, headers.length).setBackground('#4285f4');
    sheet.getRange(1, 1, 1, headers.length).setFontColor('white');
    
    // Set column widths
    sheet.setColumnWidth(1, 60);   // Sr. No
    sheet.setColumnWidth(2, 150);  // Idea Title
    sheet.setColumnWidth(3, 200);  // Idea Description
    sheet.setColumnWidth(4, 100);  // Idea Date
    sheet.setColumnWidth(5, 150);  // Planned Implementation Date
    sheet.setColumnWidth(6, 150);  // Actual Implementation Date
    sheet.setColumnWidth(7, 100);  // Status
    sheet.setColumnWidth(8, 150);  // Remarks / Issues
    sheet.setColumnWidth(9, 120);  // On Time / Delayed
    sheet.setColumnWidth(10, 100); // Delay Days
    sheet.setColumnWidth(11, 100); // Estimated?
    sheet.setColumnWidth(12, 150); // Days Since Allocated
    
    // Add sample row with formulas
    const sampleData = [
      [1, 'Sample Idea', 'This is a sample idea description', new Date(), '', '', 'Pending', '', '', '', '', '']
    ];
    sheet.getRange(2, 1, 1, 8).setValues(sampleData);
    
    // Set formulas for calculated columns (I, J, K, L)
    sheet.getRange('I2').setFormula('=IF(AND(E2<>"",F2<>""),IF(F2<=E2,"On Time","Delayed"),IF(AND(E2<>"",TODAY()>E2),"Delayed","TBD"))');
    sheet.getRange('J2').setFormula('=IF(AND(E2<>"",F2<>""),IF(F2>E2,F2-E2,0),IF(AND(E2<>"",TODAY()>E2),TODAY()-E2,0))');
    sheet.getRange('K2').setFormula('=IF(E2<>"","Yes","No")');
    sheet.getRange('L2').setFormula('=IF(D2<>"",TODAY()-D2,"")');
    
    // Protect the formula columns
    const formulaRange = sheet.getRange('I:L');
    const protection = formulaRange.protect().setDescription('Formula columns - do not edit');
    protection.setWarningOnly(true);
    
    // Record in Master sheet
    recordTemplateInMaster('Idea Template', newSpreadsheet.getUrl(), newSpreadsheet.getId());
    
    SpreadsheetApp.getUi().alert('Idea Template created successfully!\n\nURL: ' + newSpreadsheet.getUrl());
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error creating Idea Template: ' + error.toString());
  }
}

function createTaskTemplate() {
  try {
    const newSpreadsheet = SpreadsheetApp.create('Task Template - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));
    const sheet = newSpreadsheet.getActiveSheet();
    sheet.setName('Tasks');
    
    // Set up headers
    const headers = [
      'Sr. No', 'Task Title', 'Task Description', 'Allocated Date', 
      'Planned Completion Date', 'Actual Completion Date', 
      'Status', 'Remarks / Issues', 'On Time / Delayed', 
      'Delay Days', 'Estimated?', 'Days Since Allocated'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.getRange(1, 1, 1, headers.length).setBackground('#34a853');
    sheet.getRange(1, 1, 1, headers.length).setFontColor('white');
    
    // Set column widths
    sheet.setColumnWidth(1, 60);   // Sr. No
    sheet.setColumnWidth(2, 150);  // Task Title
    sheet.setColumnWidth(3, 200);  // Task Description
    sheet.setColumnWidth(4, 100);  // Allocated Date
    sheet.setColumnWidth(5, 150);  // Planned Completion Date
    sheet.setColumnWidth(6, 150);  // Actual Completion Date
    sheet.setColumnWidth(7, 100);  // Status
    sheet.setColumnWidth(8, 150);  // Remarks / Issues
    sheet.setColumnWidth(9, 120);  // On Time / Delayed
    sheet.setColumnWidth(10, 100); // Delay Days
    sheet.setColumnWidth(11, 100); // Estimated?
    sheet.setColumnWidth(12, 150); // Days Since Allocated
    
    // Add sample row with formulas
    const sampleData = [
      [1, 'Sample Task', 'This is a sample task description', new Date(), '', '', 'Pending', '', '', '', '', '']
    ];
    sheet.getRange(2, 1, 1, 8).setValues(sampleData);
    
    // Set formulas for calculated columns (I, J, K, L)
    sheet.getRange('I2').setFormula('=IF(AND(E2<>"",F2<>""),IF(F2<=E2,"On Time","Delayed"),IF(AND(E2<>"",TODAY()>E2),"Delayed","TBD"))');
    sheet.getRange('J2').setFormula('=IF(AND(E2<>"",F2<>""),IF(F2>E2,F2-E2,0),IF(AND(E2<>"",TODAY()>E2),TODAY()-E2,0))');
    sheet.getRange('K2').setFormula('=IF(E2<>"","Yes","No")');
    sheet.getRange('L2').setFormula('=IF(D2<>"",TODAY()-D2,"")');
    
    // Protect the formula columns
    const formulaRange = sheet.getRange('I:L');
    const protection = formulaRange.protect().setDescription('Formula columns - do not edit');
    protection.setWarningOnly(true);
    
    // Record in Master sheet
    recordTemplateInMaster('Task Template', newSpreadsheet.getUrl(), newSpreadsheet.getId());
    
    SpreadsheetApp.getUi().alert('Task Template created successfully!\n\nURL: ' + newSpreadsheet.getUrl());
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error creating Task Template: ' + error.toString());
  }
}

function recordTemplateInMaster(templateType, sheetUrl, sheetId) {
  try {
    const masterSheet = getMasterSheet();
    const currentUser = Session.getActiveUser().getEmail();
    
    // Find the next empty row
    const lastRow = masterSheet.getLastRow();
    const nextRow = lastRow + 1;
    
    // Prepare data for Master sheet
    const masterData = [
      templateType,                    // A: Sheet Template
      currentUser,                     // B: Shared With
      currentUser,                     // C: Email ID
      sheetUrl,                        // D: Sheet URL
      'Active',                        // E: Sheet Status
      new Date(),                      // F: Date Created
      '',                              // G: Total tasks (formula will be added)
      '',                              // H: Completed tasks (formula will be added)
      '',                              // I: Pending tasks (formula will be added)
      '',                              // J: Overdue tasks (formula will be added)
      ''                               // K: Progress % (formula will be added)
    ];
    
    // Insert the data
    masterSheet.getRange(nextRow, 1, 1, masterData.length).setValues([masterData]);
    
    // Add formulas for calculated columns (G-K)
    const sheetIdForFormula = sheetId;
    
    // These formulas will count data from the linked spreadsheet
    masterSheet.getRange(`G${nextRow}`).setFormula(`=IF(ISBLANK(D${nextRow}),"",COUNTA(IMPORTRANGE(D${nextRow},"B2:B")))`);
    masterSheet.getRange(`H${nextRow}`).setFormula(`=IF(ISBLANK(D${nextRow}),"",COUNTIF(IMPORTRANGE(D${nextRow},"G:G"),"Completed"))`);
    masterSheet.getRange(`I${nextRow}`).setFormula(`=IF(ISBLANK(D${nextRow}),"",COUNTIFS(IMPORTRANGE(D${nextRow},"G:G"),"Pending")+COUNTIFS(IMPORTRANGE(D${nextRow},"G:G"),"In Progress"))`);
    masterSheet.getRange(`J${nextRow}`).setFormula(`=IF(ISBLANK(D${nextRow}),"",COUNTIF(IMPORTRANGE(D${nextRow},"I:I"),"Delayed"))`);
    masterSheet.getRange(`K${nextRow}`).setFormula(`=IF(G${nextRow}=0,"0%",H${nextRow}/G${nextRow})`);
    
    // Format the progress column as percentage
    masterSheet.getRange(`K${nextRow}`).setNumberFormat('0.00%');
    
  } catch (error) {
    console.error('Error recording template in Master sheet:', error);
    throw error;
  }
}

function getMasterSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let masterSheet;
  
  try {
    masterSheet = spreadsheet.getSheetByName('Sheets_Master');
  } catch (error) {
    // Sheets_Master sheet doesn't exist, create it
    masterSheet = spreadsheet.insertSheet('Sheets_Master');
    initializeMasterSheet(masterSheet);
  }
  
  return masterSheet;
}

function initializeMasterSheet(sheet) {
  // Set up Master sheet headers
  const headers = [
    'Sheet Template', 'Shared With', 'Email ID', 'Sheet URL', 'Sheet Status',
    'Date Created', 'Total tasks', 'Completed tasks', 'Pending tasks', 
    'Overdue tasks', 'Progress %'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.getRange(1, 1, 1, headers.length).setBackground('#ff9900');
  sheet.getRange(1, 1, 1, headers.length).setFontColor('white');
  
  // Set column widths
  sheet.setColumnWidth(1, 120); // Sheet Template
  sheet.setColumnWidth(2, 150); // Shared With
  sheet.setColumnWidth(3, 200); // Email ID
  sheet.setColumnWidth(4, 300); // Sheet URL
  sheet.setColumnWidth(5, 100); // Sheet Status
  sheet.setColumnWidth(6, 120); // Date Created
  sheet.setColumnWidth(7, 100); // Total tasks
  sheet.setColumnWidth(8, 120); // Completed tasks
  sheet.setColumnWidth(9, 100); // Pending tasks
  sheet.setColumnWidth(10, 100); // Overdue tasks
  sheet.setColumnWidth(11, 100); // Progress %
  
  // Protect formula columns
  const formulaRange = sheet.getRange('G:K');
  const protection = formulaRange.protect().setDescription('Formula columns - do not edit');
  protection.setWarningOnly(true);
}

function openDashboard() {
  try {
    console.log('=== DASHBOARD OPENING ===');
    
    // Get the data server-side
    console.log('Step 1: Calling getMasterData()');
    var data = getMasterData();
    console.log('Step 2: Data received:', JSON.stringify(data));
    console.log('Step 3: Data length:', data ? data.length : 'null');
    
    // Create HTML service and pass data directly
    console.log('Step 4: Creating HTML template from file');
    var htmlTemplate = HtmlService.createTemplateFromFile('dashboard');
    
    console.log('Step 5: Assigning data to template');
    htmlTemplate.dashboardData = data;
    console.log('Step 6: Template data assigned. Data preview:', data && data.length > 0 ? data[0].cardTitle : 'No data');
    
    console.log('Step 7: Evaluating template');
    var htmlOutput = htmlTemplate.evaluate()
      .setWidth(1200)
      .setHeight(800)
      .setTitle('Task & Idea Manager Dashboard');
    
    console.log('Step 8: Showing modal dialog');
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Dashboard');
    console.log('=== DASHBOARD OPENED ===');
    
  } catch (error) {
    console.error('Error in openDashboard:', error);
    SpreadsheetApp.getUi().alert('Error opening dashboard: ' + error.toString());
  }
}

function getMasterSheetData() {
  try {
    const masterSheet = getMasterSheet();
    const lastRow = masterSheet.getLastRow();
    
    if (lastRow <= 1) {
      return [];
    }
    
    const range = masterSheet.getRange(2, 1, lastRow - 1, 11);
    const values = range.getValues();
    
    const data = values.map(row => ({
      templateType: row[0],
      sharedWith: row[1],
      emailId: row[2],
      sheetUrl: row[3],
      sheetStatus: row[4],
      dateCreated: row[5],
      totalTasks: row[6] || 0,
      completedTasks: row[7] || 0,
      pendingTasks: row[8] || 0,
      overdueTasks: row[9] || 0,
      progressPercent: row[10] || 0
    }));
    
    return data;
    
  } catch (error) {
    console.error('Error getting Master sheet data:', error);
    return [];
  }
}

function getMasterData() {
  console.log('=== GET MASTER DATA (LIVE) ===');
  
  try {
    // Get the current spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    console.log('Spreadsheet accessed:', ss.getName());
    
    // Get the Sheets_Master sheet
    var sheet = ss.getSheetByName('Sheets_Master');
    console.log('Sheet found:', sheet ? 'Yes' : 'No');
    
    if (!sheet) {
      console.error('Sheets_Master sheet not found');
      return [];
    }
    
    var lastRow = sheet.getLastRow();
    console.log('Last row:', lastRow);
    
    if (lastRow <= 1) {
      console.log('No data rows found');
      return [];
    }
    
    // Get headers from row 1
    var headers = sheet.getRange(1, 1, 1, 11).getValues()[0];
    console.log('Headers:', headers);
    
    // Get data from row 2 onwards
    var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    console.log('Raw data rows:', data.length);
    console.log('Sample row:', data[0]);
    
    var result = [];
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      console.log('Processing row ' + (i + 1) + ':', row);
      
      // Column A and B processing
      var colA = row[0] || '';
      var colB = row[1] || '';
      var firstWordA = colA.toString().split(' ')[0] || '';
      var firstWordB = colB.toString().split(' ')[0] || colB.toString().split('@')[0] || '';
      var cardTitle = firstWordA + ' Sheet ' + firstWordB;
      
      console.log('Card title created:', cardTitle);
      console.log('Values - G:', row[6], 'H:', row[7], 'I:', row[8], 'J:', row[9], 'K:', row[10]);
      
      var item = {
        template: colA,
        cardTitle: cardTitle,
        sharedWith: colB,
        email: row[2] || '',
        url: row[3] || '',
        status: row[4] || 'Active',
        createdAt: row[5] || new Date(),
        total: row[6] || 0,
        completed: row[7] || 0,
        pending: row[8] || 0,
        overdue: row[9] || 0,
        progress: row[10] || 0,
        totalLabel: headers[6] || 'Total',
        completedLabel: headers[7] || 'Completed',
        pendingLabel: headers[8] || 'Pending',
        overdueLabel: headers[9] || 'Overdue',
        progressLabel: headers[10] || 'Progress'
      };
      
      result.push(item);
      console.log('Item added:', item.cardTitle);
    }
    
    console.log('Final result:', result.length + ' items created');
    console.log('=== GET MASTER DATA (LIVE) END ===');
    
    return result;
    
  } catch (error) {
    console.error('Error reading live data:', error.toString());
    console.log('Falling back to test data due to error');
    
    // Return test data as fallback
    return [
      {
        template: 'Error: ' + error.message,
        cardTitle: 'Error Loading Live Data',
        sharedWith: 'System',
        email: '',
        url: '#',
        status: 'Error',
        createdAt: new Date(),
        total: 0,
        completed: 0,
        pending: 0,
        overdue: 0,
        progress: 0,
        totalLabel: 'Total',
        completedLabel: 'Completed',
        pendingLabel: 'Pending',
        overdueLabel: 'Overdue',
        progressLabel: 'Progress'
      }
    ];
  }
}

function getMasterDataFromSheet() {
  // This is the actual sheet reading function - we'll test this separately
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Sheets_Master');
  
  var lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) {
    return [];
  }
  
  // Get headers
  var headers = sheet.getRange(1, 1, 1, 11).getValues()[0];
  
  // Get data
  var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  
  var result = [];
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    
    // Column A and B processing
    var colA = row[0] || '';
    var colB = row[1] || '';
    var firstWordA = colA.toString().split(' ')[0] || '';
    var firstWordB = colB.toString().split(' ')[0] || colB.toString().split('@')[0] || '';
    var cardTitle = firstWordA + ' Sheet ' + firstWordB;
    
    result.push({
      template: colA,
      cardTitle: cardTitle,
      sharedWith: colB,
      email: row[2] || '',
      url: row[3] || '',
      status: row[4] || 'Active',
      createdAt: row[5] || new Date(),
      total: row[6] || 0,
      completed: row[7] || 0,
      pending: row[8] || 0,
      overdue: row[9] || 0,
      progress: row[10] || 0,
      totalLabel: headers[6] || 'Total',
      completedLabel: headers[7] || 'Completed',
      pendingLabel: headers[8] || 'Pending',
      overdueLabel: headers[9] || 'Overdue',
      progressLabel: headers[10] || 'Progress'
    });
  }
  
  return result;
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}