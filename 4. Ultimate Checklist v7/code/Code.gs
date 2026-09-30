function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Ultimate Checklist')
    .addItem('Setup Sheet', 'onOpen')
    .addItem('Assign Buddy', 'showAssignBuddyDialog')
    .addToUi();
}

Date.prototype.addTheDays = function (days) {
  var date = new Date(this.valueOf());
  date.setDate(date.getDate() + days);
  return date;
}

function getWorkingDateOnOrBefore(date, workingDatesStr, calendarFirstDate) {
  let candidate = new Date(date)
  let firstDate = new Date(calendarFirstDate)
  candidate.setHours(0, 0, 0, 0)
  firstDate.setHours(0, 0, 0, 0)

  let candidateStr = Utilities.formatDate(candidate, Session.getScriptTimeZone(), "yyyy-MM-dd")
  while (!workingDatesStr.includes(candidateStr)) {
    if (candidate.valueOf() <= firstDate.valueOf()) {
      throw new Error("No working date found on or before " + Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd") + " in Working Day Calender.")
    }

    candidate = Date.parse(candidate).addDays(-1)
    candidateStr = Utilities.formatDate(candidate, Session.getScriptTimeZone(), "yyyy-MM-dd")
  }

  return candidate
}

function showAssignBuddyDialog() {
  var html = HtmlService.createHtmlOutputFromFile('AssignBuddyDialog')
    .setWidth(460)
    .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Assign Buddy');
}

function getAssignBuddyOptions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var doerSheet = ss.getSheetByName("Doer List");
  if (!doerSheet) {
    throw new Error("Doer List sheet was not found.");
  }

  var lastRow = doerSheet.getLastRow();
  if (lastRow < 1) {
    return { doers: [], buddies: [] };
  }

  var data = doerSheet.getRange(1, 1, lastRow, 3).getValues();
  var doers = [];
  var buddies = [];
  var seenDoers = {};
  var seenEmails = {};

  data.forEach(function(row) {
    var name = String(row[0] || "").trim();
    var department = String(row[1] || "").trim();
    var email = String(row[2] || "").trim();

    if (!name || !department || !email || email.indexOf("@") === -1) {
      return;
    }

    var doerKey = name + "\u0000" + department;
    if (!seenDoers[doerKey]) {
      seenDoers[doerKey] = true;
      doers.push({
        name: name,
        department: department,
        label: name + " - " + department
      });
    }

    if (!seenEmails[email]) {
      seenEmails[email] = true;
      buddies.push({
        email: email,
        label: name + " - " + department + " (" + email + ")"
      });
    }
  });

  return { doers: doers, buddies: buddies };
}

function parseDateInput(dateText, fieldName) {
  if (!dateText) {
    throw new Error(fieldName + " is required.");
  }

  var parts = String(dateText).split("-");
  if (parts.length !== 3) {
    throw new Error(fieldName + " must be a valid date.");
  }

  var date = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  if (isNaN(date.valueOf())) {
    throw new Error(fieldName + " must be a valid date.");
  }

  date.setHours(0, 0, 0, 0);
  return date;
}

function assignBuddyToTasks(form) {
  var name = String(form.name || "").trim();
  var department = String(form.department || "").trim();
  var buddyEmail = String(form.buddyEmail || "").trim();
  var fromDate = parseDateInput(form.fromDate, "From date");
  var toDate = parseDateInput(form.toDate, "To date");

  if (!name || !department) {
    throw new Error("Choose a name and department.");
  }
  if (!buddyEmail) {
    throw new Error("Choose a buddy email.");
  }
  if (fromDate.valueOf() > toDate.valueOf()) {
    throw new Error("From date cannot be after To date.");
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("Master");
  if (!masterSheet) {
    throw new Error("Master sheet was not found.");
  }

  var lastRow = masterSheet.getLastRow();
  if (lastRow < 2) {
    return { updated: 0 };
  }

  var data = masterSheet.getRange(2, 1, lastRow - 1, 11).getValues();
  var buddyColumn = data.map(function(row) {
    return [row[10]];
  });
  var updated = 0;

  data.forEach(function(row, index) {
    var taskDoer = String(row[0] || "").trim();
    var taskDepartment = String(row[2] || "").trim();
    var plannedDate = row[6];

    if (!(plannedDate instanceof Date)) {
      return;
    }

    var normalizedPlannedDate = new Date(plannedDate);
    normalizedPlannedDate.setHours(0, 0, 0, 0);

    if (
      taskDoer === name &&
      taskDepartment === department &&
      normalizedPlannedDate.valueOf() >= fromDate.valueOf() &&
      normalizedPlannedDate.valueOf() <= toDate.valueOf()
    ) {
      buddyColumn[index][0] = buddyEmail;
      updated = updated + 1;
    }
  });

  if (updated > 0) {
    masterSheet.getRange(2, 11, buddyColumn.length, 1).setValues(buddyColumn);
  }

  return { updated: updated };
}

function setTrigger() {
  removeTrigger()
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = ss.getSheetByName("Setup Sheet")
  var time = sheet.getRange("C15").getValue()
  createTrigger(time)
}
function createTrigger(time) {
  if (!time) {
    time = 10
  }
  ScriptApp.newTrigger('sendReminder')
    .timeBased()
    .everyDays(1)
    .atHour(time)
    .create();
}


function removeTrigger() {
  // Loop over all triggers.
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    // If the current trigger is the correct one, delete it.
    if (allTriggers[i].getUniqueId() == allTriggers[i].getUniqueId()) {
      ScriptApp.deleteTrigger(allTriggers[i]);
      break;
    }
  }
}


function sendReminder() {
  var forB = SpreadsheetApp.getActive();
  var forBsh = forB.getSheetByName("Doer List")
  var lastB = forBsh.getLastRow()
  var dataB = forBsh.getRange(1, 3, lastB, 1).getValues()
  var today = new Date();
  var tomorrow = new Date((new Date().setHours(0, 0, 0, 0)).valueOf() + 1000 * 3600 * 24);
  var ss = SpreadsheetApp.getActive()
  var sheet = ss.getSheetByName("Master")
  var columnToCheck = sheet.getRange("B:B").getValues()
  var lastrow = getLastRowSpecial(columnToCheck) - 1
  var data = sheet.getRange(2, 1, lastrow, 8).getValues()
  let emails = dataB.filter(function(r){
    let reminderTasks = []
    let pendingTomorrow = data.filter(function(pending){
      let doeremail = pending[1]
      let lastDate = pending[6]
      let actual = pending[7]
      let name = pending[0]
      let task = pending[5]
      if (r[0] === doeremail && lastDate.valueOf() === tomorrow.valueOf() && !actual){
        reminderTasks.push([name,task])
      }
    })
    if (reminderTasks.length > 0){
    let joinedTasks = reminderTasks.map(function(tasks){
            return "Task : "+tasks[1]
            }
            ).join("\n")
    let dataToSend = "Hello "+reminderTasks[0][0]+",\n\nYou have planned tasks pending for tomorrow.\n\n"+joinedTasks+"\n\nPlease ignore this message if you have already completed the tasks."
    GmailApp.sendEmail(r[0], "You have a Pending Tasks for tomorrow", dataToSend);
    }
  })
}



function archive() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("Scores")
  var data = sheet.getRange(4, 1, 6, 4).getValues()
  var name = sheet.getRange("A2").getValue()
  var week = sheet.getRange("D2").getValue()
  var arch = ss.getSheetByName("Archive")
  var archLast = arch.getLastRow() - 1
  var ared = data[3][1]
  var ayellow = data[3][2]
  var pred = data[5][1]
  var pyellow = data[5][2]
  arch.appendRow([name, week, pred, pyellow, "", ared, ayellow, ""])
  var spreadsheet = ss.getSheetByName("Scores")
  spreadsheet.getRange('B9:C9').activate();
  spreadsheet.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });
}
function sendtodoer() {
  var forB = SpreadsheetApp.getActive();
  var forBsh = forB.getSheetByName("Doer List")
  var lastB = forBsh.getLastRow()
  var dataB = forBsh.getRange(1, 1, lastB, 2).getValues()
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Send Tasks To Doer")
  var lastrow = sheet.getLastRow() - 3
  var range = sheet.getRange(4, 1, lastrow, 2)
  var htmlTable = SheetConverter.convertRange2html(range);
  var status = sheet.getRange("D2").getValue()
  var weeknum = sheet.getRange("E2").getValue().toString()
  var name = sheet.getRange("A2").getValue()
  var email = findinB(name, dataB)
  var body = "Here are your " + status + " tasks of Week " + weeknum + ".<br/><br/>" + htmlTable
  GmailApp.sendEmail(email, 'Details of Delegated Tasks', body, { htmlBody: body });
}


function createChecklist() {
  let ss = SpreadsheetApp.getActiveSpreadsheet()
  let taskListSheet = ss.getSheetByName("Task List")
  let taskListData = taskListSheet.getRange(2, 1, taskListSheet.getLastRow() - 1, 6).getValues()
  let skipSundays = ss.getSheetByName("Setup Sheet").getRange("B32").getValue()
  let todaysdate = new Date()

  if (!skipSundays) {
    skipSundays = "Yes"
  }

  let masterSheet = ss.getSheetByName("Master")
  let cell = masterSheet.getRange("H2") //actual Cell
  let masterTaskIDs = masterSheet.getRange("D:D").getValues()
  let masterLastRow = masterTaskIDs.filter(r => String(r)).length

  let lastMasterTaskID = Number(masterTaskIDs[masterLastRow - 1])
  if (!lastMasterTaskID) {
    lastMasterTaskID = 0
  }

  let masterArray = []

  let calendarSheet = ss.getSheetByName("Working Day Calender")
  let calendarDates = calendarSheet.getRange("A:A").getValues()
  let calendarLast = calendarDates.filter(String).length
  if (calendarLast <= 1) {
    throw new Error("Working Day Calender must contain at least one working date below the header.")
  }
  let allCalendarDates = calendarSheet.getRange(2, 1, calendarLast - 1, 1).getValues()
  let calendarFirstDate = allCalendarDates[0][0]
  let calendarLastDate = calendarDates[calendarLast - 1][0]
  let workingDates = allCalendarDates.flat()
  let workingDatesStr = workingDates.map(x => Utilities.formatDate(x, Session.getScriptTimeZone(), "yyyy-MM-dd"));

  taskListData.forEach(function (task, index) {
    let theTask = task[0]
    let theDoer = task[1]
    let theDepartment = task[2]
    let theFreq = task[3]
    let theDate = task[4]
    let theStatus = task[5]
    let theEmail = getEmail(theDoer,theDepartment)

    if (theEmail === "FAILED") {
      taskListSheet.getRange(index + 2, 6).setValue("Skipped Due to doer - department mismatch")
      return
    }


    if (theStatus != "Sent") {
      //weekly tasks 
      if (theFreq === "W") {
        while (Date.compare(calendarLastDate, theDate) === 1) {
          lastMasterTaskID = lastMasterTaskID + 1
          let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          let frozenDate = theDate.clone()

          theDate = getWorkingDateOnOrBefore(theDate, workingDatesStr, calendarFirstDate)
          endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          masterArray.push([theDoer, theEmail, theDepartment, lastMasterTaskID, theFreq, theTask, new Date(theDate),"","",""])

          theDate = Date.parse(frozenDate).addWeeks(1)
        }
      } else if (theFreq === "M") {
        while (Date.compare(calendarLastDate, theDate) === 1) {
          let norun = 0
          if(theDate.is().jan() && theDate.getDate() > 28){
            norun = 1
            lastMasterTaskID = lastMasterTaskID + 1
          let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          let frozenDate = theDate.clone()

          theDate = getWorkingDateOnOrBefore(theDate, workingDatesStr, calendarFirstDate)
          endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          if (skipSundays === "Yes") {
            if (theDate.is().sunday()) {
              theDate = Date.parse(theDate).addDays(-1)
            }
          }

          masterArray.push([theDoer, theEmail, theDepartment, lastMasterTaskID, theFreq, theTask, new Date(theDate),"","",""])

          theDate = Date.parse(frozenDate).addMonths(2)
          }

          if (norun === 0){
          lastMasterTaskID = lastMasterTaskID + 1
          let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          let frozenDate = theDate.clone()

          theDate = getWorkingDateOnOrBefore(theDate, workingDatesStr, calendarFirstDate)
          endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          if (skipSundays === "Yes") {
            if (theDate.is().sunday()) {
              theDate = Date.parse(theDate).addDays(-1)
            }
          }

          masterArray.push([theDoer, theEmail, theDepartment, lastMasterTaskID, theFreq, theTask, new Date(theDate),"","",""])

          theDate = Date.parse(frozenDate).addMonths(1)
          }
        }
      } else if (theFreq === "Y") {
        if (theDate.valueOf() > todaysdate.valueOf()){
          lastMasterTaskID = lastMasterTaskID + 1
          masterArray.push([theDoer, theEmail, theDepartment, lastMasterTaskID, theFreq, theTask, new Date(theDate),"","",""])
        }
        lastMasterTaskID = lastMasterTaskID + 1
        theDate = Date.parse(theDate).addYears(1)
        masterArray.push([theDoer, theEmail, theDepartment, lastMasterTaskID, theFreq, theTask, new Date(theDate),"","",""])
        lastMasterTaskID = lastMasterTaskID + 1
        theDate = Date.parse(theDate).addYears(1)
        masterArray.push([theDoer, theEmail, theDepartment, lastMasterTaskID, theFreq, theTask, new Date(theDate),"","",""])
        lastMasterTaskID = lastMasterTaskID + 1
        theDate = Date.parse(theDate).addYears(1)
        masterArray.push([theDoer, theEmail, theDepartment, lastMasterTaskID, theFreq, theTask, new Date(theDate),"","",""])

      } else if (theFreq === "D") {
        while (Date.compare(calendarLastDate, theDate) === 1) {
          lastMasterTaskID = lastMasterTaskID + 1
          let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          let frozenDate = theDate.clone()

          if (skipSundays === "Yes") {
            if (workingDatesStr.includes(endDateStr) && !theDate.is().sunday()) {
              masterArray.push([theDoer, theEmail, theDepartment, lastMasterTaskID, theFreq, theTask, new Date(theDate),"","",""])
            }
          } else {
            if (workingDatesStr.includes(endDateStr)) {
              masterArray.push([theDoer, theEmail, theDepartment, lastMasterTaskID, theFreq, theTask, new Date(theDate),"","",""])
            }
          }

          theDate = Date.parse(frozenDate).addDays(1)
        }
      } else if (theFreq === "Q") {
        while (Date.compare(calendarLastDate, theDate) === 1) {
          lastMasterTaskID = lastMasterTaskID + 1
          let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          let frozenDate = theDate.clone()

          theDate = getWorkingDateOnOrBefore(theDate, workingDatesStr, calendarFirstDate)
          endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          if (skipSundays === "Yes") {
            if (theDate.is().sunday()) {
              theDate = Date.parse(theDate).addDays(-1)
            }
          }

          masterArray.push([theDoer, theEmail, theDepartment, lastMasterTaskID, theFreq, theTask, new Date(theDate),"","",""])

          theDate = Date.parse(frozenDate).addMonths(3)
        }
      } else if (theFreq === "H" || theFreq === "HY" || theFreq === "Half Yearly") {
        while (Date.compare(calendarLastDate, theDate) === 1) {
          lastMasterTaskID = lastMasterTaskID + 1
          let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          let frozenDate = theDate.clone()

          theDate = getWorkingDateOnOrBefore(theDate, workingDatesStr, calendarFirstDate)
          endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          if (skipSundays === "Yes") {
            if (theDate.is().sunday()) {
              theDate = Date.parse(theDate).addDays(-1)
            }
          }

          masterArray.push([theDoer, theEmail, theDepartment, lastMasterTaskID, theFreq, theTask, new Date(theDate),"","",""])

          theDate = Date.parse(frozenDate).addMonths(6)
        }
      } else if (theFreq === "F") {
        while (Date.compare(calendarLastDate, theDate) === 1) {
          lastMasterTaskID = lastMasterTaskID + 1
          let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          let frozenDate = theDate.clone()

          theDate = getWorkingDateOnOrBefore(theDate, workingDatesStr, calendarFirstDate)
          endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          masterArray.push([theDoer, theEmail, theDepartment, lastMasterTaskID, theFreq, theTask, new Date(theDate),"","",""])

          theDate = Date.parse(frozenDate).addWeeks(2)
        }
      }else if (theFreq === "E1st") {
        while (Date.compare(calendarLastDate, theDate) === 1) {
          lastMasterTaskID = lastMasterTaskID + 1
          let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          let day = getDayInfo(theDate.getDay())

          let frozenDate = theDate.clone()

          theDate = getWorkingDateOnOrBefore(theDate, workingDatesStr, calendarFirstDate)
          endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          masterArray.push([theDoer, theEmail, theDepartment, lastMasterTaskID, theFreq, theTask, new Date(theDate),"","",""])

          theDate = eval("Date.parse(frozenDate).next().month().first()."+day+"()")
        }
      } else if (theFreq === "E2nd") {
        while (Date.compare(calendarLastDate, theDate) === 1) {
          lastMasterTaskID = lastMasterTaskID + 1
          let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          let day = getDayInfo(theDate.getDay())

          let frozenDate = theDate.clone()

          theDate = getWorkingDateOnOrBefore(theDate, workingDatesStr, calendarFirstDate)
          endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          masterArray.push([theDoer, theEmail, theDepartment, lastMasterTaskID, theFreq, theTask, new Date(theDate),"","",""])

          theDate = eval("Date.parse(frozenDate).next().month().second()."+day+"()")
        }
      } else if (theFreq === "E3rd") {
        while (Date.compare(calendarLastDate, theDate) === 1) {
          lastMasterTaskID = lastMasterTaskID + 1
          let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          let day = getDayInfo(theDate.getDay())

          let frozenDate = theDate.clone()

          theDate = getWorkingDateOnOrBefore(theDate, workingDatesStr, calendarFirstDate)
          endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          masterArray.push([theDoer, theEmail, theDepartment, lastMasterTaskID, theFreq, theTask, new Date(theDate),"","",""])

          theDate = eval("Date.parse(frozenDate).next().month().third()."+day+"()")
        }
      } else if (theFreq === "E4th") {
        while (Date.compare(calendarLastDate, theDate) === 1) {
          lastMasterTaskID = lastMasterTaskID + 1
          let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          let day = getDayInfo(theDate.getDay())

          let frozenDate = theDate.clone()

          theDate = getWorkingDateOnOrBefore(theDate, workingDatesStr, calendarFirstDate)
          endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          masterArray.push([theDoer, theEmail, theDepartment, lastMasterTaskID, theFreq, theTask, new Date(theDate),"","",""])

          theDate = eval("Date.parse(frozenDate).next().month().fourth()."+day+"()")
        }
      }else if (theFreq === "ELast") {
        while (Date.compare(calendarLastDate, theDate) === 1) {
          lastMasterTaskID = lastMasterTaskID + 1
          let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          let day = getDayInfo(theDate.getDay())

          let frozenDate = theDate.clone()

          theDate = getWorkingDateOnOrBefore(theDate, workingDatesStr, calendarFirstDate)
          endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          masterArray.push([theDoer, theEmail, theDepartment, lastMasterTaskID, theFreq, theTask, new Date(theDate),"","",""])

          theDate = eval("Date.parse(frozenDate).next().month().final()."+day+"()")
        }
      }

      taskListSheet.getRange(index + 2, 6).setValue("Sent")
    }

  })
  if (masterArray.length > 0) {
    masterSheet.getRange(masterLastRow + 1, 1, masterArray.length, 10).setValues(masterArray)
  }
}


function onChange() 
{
  var ss=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var active=ss.getActiveCell();
  var val=active.getValue();
  
  if (val === "Done") {
    
    if (ss.getSheetName() === "Master") {
      var row = active.getRow();
      var col = active.getColumn();
 
      var date = new Date();
 
      ts = ss.getRange(row, col - 1).getValue();//Actual Time column
      if (!ts) {
        ss.getRange(row, col - 1).setValue(date);
      }
    }
  }
}

function createTriggerOnChange() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onChange')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onChange()
      .create();
}
