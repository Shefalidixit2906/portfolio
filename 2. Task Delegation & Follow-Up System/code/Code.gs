let spreadsheet = new ManagedSpreadsheet();
let sss = SpreadsheetApp.getActive()

// Helper function to get Doer data as a Map with caching
function getDoerMap() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'doerMap';
  const cached = cache.get(cacheKey);
  
  if (cached != null) {
    // console.log("Using cached Doer Map");
    return JSON.parse(cached); // Parse the stored JSON string
  }

  // console.log("Fetching Doer Map from sheet");
  let sheet = sss.getSheetByName("Doer List");
  let data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  let doerMap = {};
  data.forEach(row => {
    const name = row[0];
    const number = row[1] || ''; // Ensure number is string, default empty
    const email = row[2] || ''; // Ensure email is string, default empty
    if (name) { // Only add if name exists
       doerMap[name] = { number: number, email: email };
    }
  });
  
  // Cache the result for 10 minutes
  cache.put(cacheKey, JSON.stringify(doerMap), 600); 
  
  return doerMap;
}

function currentVersion() {
  return "7.0.1"
}

function migrate() {
  let ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Delegation Sheet URL', 'Please enter the OLD Delegation Sheet URL:', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    var url = response.getResponseText();
    let oldSheet = SpreadsheetApp.openByUrl(url);
    let newSheet = SpreadsheetApp.getActiveSpreadsheet();

    // Get the old sheets
    let oldMaster = oldSheet.getSheetByName("Master");
    let oldArchive = oldSheet.getSheetByName("Archive");
    let oldDoerList = oldSheet.getSheetByName("Doer List");

    // Get the new sheets
    let newMaster = newSheet.getSheetByName("Master");
    let newArchive = newSheet.getSheetByName("Archive");
    let newDoerList = newSheet.getSheetByName("Doer List");

    // Copy data from oldMaster to newMaster
    if (oldMaster && newMaster) {
      let oldMasterData = oldMaster.getDataRange().getValues();
      newMaster.clearContents();
      newMaster.getRange(1, 1, oldMasterData.length, oldMasterData[0].length).setValues(oldMasterData);
    } else {
      ui.alert("Error: 'Master' sheet not found in one of the spreadsheets.");
    }

    // Copy data from oldArchive to newArchive
    if (oldArchive && newArchive) {
      let oldArchiveData = oldArchive.getDataRange().getValues();
      newArchive.clearContents();
      newArchive.getRange(1, 1, oldArchiveData.length, oldArchiveData[0].length).setValues(oldArchiveData);
    } else {
      ui.alert("Error: 'Archive' sheet not found in one of the spreadsheets.");
    }

    ui.alert("Migration completed successfully! Please migrate the Doer List Manually.");

  } else {
    ui.alert('Migration canceled.');
  }
}

function addNew(options) {
  let name = options.name;
  let task = options.task;
  let firstDate = options.fdate;
  let worksheet = spreadsheet.sheet("Master");
  let date = firstDate;
  let taskid = makeid(7);
  
  // Get Doer details to construct the name-number string
  const doerMap = getDoerMap();
  const doerDetails = doerMap[name];
  let nameToStore = name; // Default to just name if not found
  if (doerDetails && doerDetails.number) {
    nameToStore = `${name} - ${doerDetails.number}`;
  }

  worksheet.append({ 
    'Task ID': taskid, 
    'Name': nameToStore, // Store combined Name - Number
    'Task': task, 
    'First Date': date, 
    'Latest Revision': date, 
    'Total Revisions': 0, 
    'Status': "Pending" 
  });
  return "Done";
}

function addNext(options) {
  let id = options.id
  let theDate = options.pl
  let dateParts = theDate.split("/");
  var nextDate = new Date(+dateParts[2], +dateParts[1] - 1, +dateParts[0]); 
  let worksheet = spreadsheet.sheet("Master")
  let date = nextDate.toLocaleDateString("en-GB")
  let task = worksheet.find('Task ID', id)
  let index = worksheet.rowIndex('Task ID', id)
  let revs = task["Total Revisions"]
  let firstDate = task["First Date"]
  let correctWeek = nextDate.getWeek() === firstDate.getWeek()

  if (revs > 1 || !correctWeek) {
    if (revs < 2) {
      task["Total Revisions"] = 2
    }

    task["Status"] = "Week Shifted"

    worksheet.update(index,task)

    worksheet.append({ 'Task ID': makeid(7), 'Name': task["Name"], 'Task': task["Task"], 'First Date': date, 'Latest Revision': date, 'Total Revisions': 0 , 'Status': "Pending"})


  } else {
    if (revs === 0) {
      task["Revision 1"] = date
    } else if (revs === 1) {
      task["Revision 2"] = date
    }
    task["Total Revisions"] = revs + 1
    task["Latest Revision"] = date
  }
  worksheet.update(index, task)

}


function fetch(){
  var tomorrow = new Date(new Date().getTime() + (24 * 60 * 60 * 1000));
  let onlyDate = new Date(tomorrow.getFullYear(),tomorrow.getMonth(),tomorrow.getDate())
  let ss = SpreadsheetApp.getActiveSpreadsheet()
  let sheet = ss.getSheetByName("Master")
  let data = sheet.getRange(2,1,sheet.getLastRow(),9).getValues()
  let pending = data.filter(function(row){
    let task_status = row[8]
    let last_revision = row[6]
    if (task_status === "Pending" && last_revision < onlyDate) {
      return row
    }
  })
  return pending 
}

function fetchAllPendings(){
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Master");
  let data = sheet.getRange(2,1,sheet.getLastRow(),9).getValues();
  
  // Filter pending tasks
  let pending = data.filter(function(row){
    let task_status = row[8];
    if (task_status === "Pending") {
      return row;
    }
  });
  
  // Sort by last revision date (assumed to be in the 7th column, index 6)
  pending.sort(function(a, b) {
    let dateA = new Date(a[6]);
    let dateB = new Date(b[6]);
    return dateA - dateB; // Sorts in ascending order by date
  });
  
  return pending;
}


function taskComplete(taskID){
  let worksheet = spreadsheet.sheet("Master")
  let task = worksheet.find('Task ID', taskID)
  let index = worksheet.rowIndex('Task ID', taskID)
  task["Status"] = "Completed"
  worksheet.update(index, task)

}



function fetchAllPendingData(){
  let data = fetchAllPendings(); // Gets raw data rows from Master sheet
  let doerMap = getDoerMap();
  let array = [];
  data.forEach(function(r){
    // r[0]=taskID, r[1]=Name (potentially 'Name - Number'), r[2]=Task, r[6]=Date object
    const storedName = r[1]; 
    const nameOnly = storedName.includes(' - ') ? storedName.split(' - ')[0] : storedName; // Extract name part
    const doerDetails = doerMap[nameOnly] || { number: '', email: '' }; // Find details
    
    array.push([
      r[0], // Task ID
      nameOnly, // Just the name for display consistency?
      r[2], // Task
      r[6].toLocaleDateString("en-GB"), // Formatted Date
      doerDetails.number, // Number
      doerDetails.email // Email
    ]);
  });
  return array;
}


function fetchRelevantData(){
  let data = fetch(); // Gets raw data rows from Master sheet
  let doerMap = getDoerMap();
  let array = [];
  data.forEach(function(r){
    // r[0]=taskID, r[1]=Name (potentially 'Name - Number'), r[2]=Task, r[6]=Date object
    const storedName = r[1];
    const nameOnly = storedName.includes(' - ') ? storedName.split(' - ')[0] : storedName; // Extract name part
    const doerDetails = doerMap[nameOnly] || { number: '', email: '' }; // Find details
    
    array.push([
      r[0], // Task ID
      nameOnly, // Just the name
      r[2], // Task
      r[6].toLocaleDateString("en-GB"), // Formatted Date
      doerDetails.number, // Number
      doerDetails.email // Email
    ]);
  });
  return array;
}

function getInfoByTaskID(taskID){
  let worksheet = spreadsheet.sheet("Master");
  let task = worksheet.find('Task ID', taskID);
  // Get Doer details from the map
  const doerMap = getDoerMap();
  const storedName = task["Name"];
  const nameOnly = storedName.includes(' - ') ? storedName.split(' - ')[0] : storedName;
  const doerDetails = doerMap[nameOnly] || { number: '', email: '' };
  
  let taskinfo = task["Task"];
  let firstDate = task["First Date"] instanceof Date ? task["First Date"].toLocaleDateString("en-GB") : task["First Date"];
  
  // Return all necessary info for the revise form, including the original name for the dropdown
  return {
    name: nameOnly, // Just the name for dropdown selection
    task: taskinfo,
    taskid: taskID,
    fdate: firstDate, // Use First Date for revise view?
    number: doerDetails.number,
    email: doerDetails.email
    // Note: Revise form currently only uses name, task, id, fdate. 
    // If number/email need displaying there, revise.html needs changes too.
  };
}

function getAllDoers(){
  // This can now potentially just use the Map function if preferred,
  // but keeping it separate allows fetching raw data for preloading if needed.
  let sheet = sss.getSheetByName("Doer List");
  // Read starting row 2, column 1, for (lastRow - 1) rows, and 3 columns
  let data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  // Filter out rows where the name (first column) is empty
  data = data.filter(row => row[0] && String(row[0]).trim() !== ''); 
  // This now returns [[name, number, email], ...]
  return data;
}

function archive() {
  let sheet = sss.getSheetByName("Dashboard")
  let data = sheet.getRange(4, 1, 6, 4).getValues()
  let name = sheet.getRange("A2").getValue()
  let week = sheet.getRange("D2").getValue()
  let arch = sss.getSheetByName("Archive")
  let archLast = arch.getLastRow() - 1
  let ared = data[3][1]
  let ayellow = data[3][2]
  let agreen = data[3][3]
  let pred = data[5][1]
  let pyellow = data[5][2]
  let pgreen = data[5][3]
  arch.appendRow([name,week,pred,pyellow,pgreen,ared,ayellow,agreen])
  let spreadsheet = sss.getSheetByName("Dashboard")
  spreadsheet.getRange('B9:D9').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
}

/**
 * Sends reminder emails for tasks due tomorrow
 * Can be scheduled to run daily with a time-based trigger
 * @param {boolean} manualTrigger - If true, will force send reminders regardless of due date
 * @return {object} Summary of emails sent
 */
function sendTaskReminders(manualTrigger) {
  // Default to false if not provided
  manualTrigger = manualTrigger || false;
  
  // Get all pending tasks
  let data = fetchAllPendings();
  
  // Get tomorrow's date for comparison
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  tomorrow.setHours(0, 0, 0, 0); // Set to start of day
  
  // For tasks due tomorrow (or all tasks if manual trigger)
  let tasksToRemind = manualTrigger ? data : data.filter(function(row) {
    // Compare dates (assuming revision date is in the 7th column at index 6)
    let taskDueDate = new Date(row[6]); 
    return taskDueDate.toDateString() === tomorrow.toDateString();
  });
  
  if (tasksToRemind.length === 0) {
    return { success: true, message: "No tasks due tomorrow to remind about", emailsSent: 0 };
  }
  
  // Get the doer map for emails
  const doerMap = getDoerMap();
  
  // Group tasks by doer
  let tasksByDoer = {};
  
  tasksToRemind.forEach(function(row) {
    // Extract the base name (without the number)
    const storedName = row[1];
    const nameOnly = storedName.includes(' - ') ? storedName.split(' - ')[0] : storedName;
    
    // Skip if no email available for this doer
    if (!doerMap[nameOnly] || !doerMap[nameOnly].email) {
      return;
    }
    
    // Initialize array for this doer if it doesn't exist
    if (!tasksByDoer[nameOnly]) {
      tasksByDoer[nameOnly] = {
        email: doerMap[nameOnly].email,
        number: doerMap[nameOnly].number,
        tasks: []
      };
    }
    
    // Add task info
    tasksByDoer[nameOnly].tasks.push({
      id: row[0],
      task: row[2],
      dueDate: new Date(row[6]).toLocaleDateString("en-GB") // Format date
    });
  });
  
  // Send emails to each doer
  let emailsSent = 0;
  let emailErrors = [];
  
  for (let doerName in tasksByDoer) {
    try {
      const doerInfo = tasksByDoer[doerName];
      if (doerInfo.tasks.length > 0) {
        sendReminderEmail(doerName, doerInfo);
        emailsSent++;
      }
    } catch (error) {
      emailErrors.push({
        doer: doerName,
        error: error.toString()
      });
    }
  }
  
  return {
    success: true,
    message: `Sent ${emailsSent} reminder emails`,
    errors: emailErrors,
    emailsSent: emailsSent
  };
}

/**
 * Sends an email to a specific doer with their pending tasks
 * @param {string} doerName - The name of the doer
 * @param {object} doerInfo - Object containing email, number and tasks array
 */
function sendReminderEmail(doerName, doerInfo) {
  const email = doerInfo.email;
  const tasks = doerInfo.tasks;
  
  // Create email subject
  const subject = tasks.length === 1 
    ? `Reminder: 1 Task Due Tomorrow` 
    : `Reminder: ${tasks.length} Tasks Due Tomorrow`;
  
  // Create HTML for email body
  let htmlBody = `
    <p>Hello ${doerName},</p>
    <p>This is a reminder about ${tasks.length === 1 ? 'your task' : 'your tasks'} due tomorrow:</p>
    <table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse;">
      <tr style="background-color: #f2f2f2;">
        <th>Task</th>
        <th>Due Date</th>
      </tr>
  `;
  
  // Add each task to the email
  tasks.forEach(function(task) {
    htmlBody += `
      <tr>
        <td>${task.task}</td>
        <td>${task.dueDate}</td>
      </tr>
    `;
  });
  
  // Close the table and add footer
  htmlBody += `
    </table>
    <p>Please complete ${tasks.length === 1 ? 'this task' : 'these tasks'} by the due date.</p>
    <p>Thank you!</p>
  `;
  
  // Send the email
  GmailApp.sendEmail(
    email,
    subject,
    // Plain text version
    `Hello ${doerName}, this is a reminder about your task(s) due tomorrow. Please check your email for details.`,
    {
      htmlBody: htmlBody,
      name: "Delegation Sheet Task Reminder"
    }
  );
  
  return true;
}

/**
 * Function to be called from the UI to manually trigger reminders
 * @return {object} Summary of emails sent
 */
function manualSendReminders() {
  return sendTaskReminders(true);
}

/**
 * Function to set up a daily trigger for automatic reminders
 * This should be run once by an admin to set up the automatic schedule
 */
function setupDailyReminderTrigger() {
  // Delete any existing triggers with the same function name
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendTaskReminders') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create a new trigger to run at 9:00 AM every day
  ScriptApp.newTrigger('sendTaskReminders')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();
    
  return { success: true, message: "Daily reminder trigger set for 9:00 AM" };
}

/**
 * Creates a custom menu in the Google Sheets UI when the spreadsheet is opened
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Delegation Tools')
      .addItem('Setup Daily Reminders (9:00 AM)', 'setupDailyReminderTriggerWithUI')
      .addItem('Send Reminders Manually', 'manualSendRemindersWithUI')
      .addSeparator()
      .addItem('Flush Doer Data Cache', 'flushCacheWithUI')
      .addSeparator()
      .addItem('Migrate from Old Sheet', 'migrate')
      .addToUi();
}

/**
 * Wrapper for setupDailyReminderTrigger with UI feedback
 */
function setupDailyReminderTriggerWithUI() {
  try {
    const result = setupDailyReminderTrigger();
    const ui = SpreadsheetApp.getUi();
    ui.alert('Success', 'Daily reminder trigger has been set up to run at 9:00 AM every day.', ui.ButtonSet.OK);
  } catch (error) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Error', 'Failed to set up the daily reminder trigger: ' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Wrapper for manualSendReminders with UI feedback
 */
function manualSendRemindersWithUI() {
  try {
    const result = manualSendReminders();
    const ui = SpreadsheetApp.getUi();
    
    if (result.emailsSent > 0) {
      ui.alert('Success', `Successfully sent ${result.emailsSent} reminder email(s).`, ui.ButtonSet.OK);
    } else {
      ui.alert('Information', 'No reminders sent. There are no pending tasks.', ui.ButtonSet.OK);
    }
  } catch (error) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Error', 'Failed to send reminders: ' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Flushes the cached doer data
 * @return {object} Result of the operation
 */
function flushCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove('doerMap');
    return { success: true, message: "Doer data cache successfully flushed" };
  } catch (error) {
    return { success: false, message: "Failed to flush cache: " + error.toString() };
  }
}

/**
 * Wrapper for flushCache with UI feedback
 */
function flushCacheWithUI() {
  try {
    const result = flushCache();
    const ui = SpreadsheetApp.getUi();
    
    if (result.success) {
      ui.alert('Success', 'Doer data cache has been successfully cleared. New data will be fetched from the sheet on next request.', ui.ButtonSet.OK);
    } else {
      ui.alert('Error', result.message, ui.ButtonSet.OK);
    }
  } catch (error) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Error', 'Failed to flush cache: ' + error.toString(), ui.ButtonSet.OK);
  }
}

