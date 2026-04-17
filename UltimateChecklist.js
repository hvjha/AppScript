// // ============================================================
// //  ULTIMATE CHECKLIST — REWRITTEN
// //  Key fixes:
// //  1. Timeout → Batch processing with execution-time guard
// //  2. SubTaskID → Global sequential counter across ALL tasks
// //  3. Resume   → On re-run, continues SubTaskID from last row in Master
// // ============================================================


// // ─── ENTRY POINT (called by your Submit button) ─────────────
// function createChecklist() {
//   Logger.log("=== createChecklist START ===");

//   var ss = SpreadsheetApp.getActiveSpreadsheet();

//   // ── Sheets ──────────────────────────────────────────────
//   var taskListSheet  = ss.getSheetByName("Task_List");
//   var masterSheet    = ss.getSheetByName("Master");
//   var setupSheet     = ss.getSheetByName("Setup Sheet");
//   var calendarSheet  = ss.getSheetByName("Working Day Calender");

//   if (!taskListSheet || !masterSheet || !setupSheet || !calendarSheet) {
//     Logger.log("ERROR: One or more required sheets not found. Aborting.");
//     SpreadsheetApp.getUi().alert("Required sheet missing (Task_List / Master / Setup Sheet / Working Day Calender).");
//     return;
//   }

//   // ── Task_List data ───────────────────────────────────────
//   var taskLastRow = taskListSheet.getLastRow();
//   if (taskLastRow < 2) {
//     Logger.log("No data in Task_List. Aborting.");
//     return;
//   }
//   var taskListData = taskListSheet.getRange(2, 1, taskLastRow - 1, 11).getValues();

//   // ── Setup: Skip Sundays ──────────────────────────────────
//   var skipSundays = setupSheet.getRange("B32").getValue();
//   if (!skipSundays) skipSundays = "Yes";

//   // ── Working Day Calendar ─────────────────────────────────
//   var calDates    = calendarSheet.getRange("A:A").getValues();
//   var calLastIdx  = calDates.filter(String).length;          // count non-empty
//   if (calLastIdx < 2) {
//     Logger.log("ERROR: Working Day Calender is empty.");
//     return;
//   }
//   var allCalDates     = calendarSheet.getRange(2, 1, calLastIdx - 1, 1).getValues().flat();
//   var calendarLastDate = calDates[calLastIdx - 1][0];        // last working date
//   var workingDatesStr = allCalDates.map(function(d) {
//     return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
//   });

//   // ── SubTaskID: resume from last row in Master ────────────
//   //    SubTaskID lives in column E (index 4)
//   var masterLastRow   = getLastDataRow(masterSheet, "B");    // col B as anchor
//   var lastSubTaskID   = 0;
//   if (masterLastRow >= 2) {
//     var existingIDs = masterSheet.getRange(2, 5, masterLastRow - 1, 1).getValues();
//     existingIDs.forEach(function(r) {
//       var n = Number(r[0]);
//       if (!isNaN(n) && n > lastSubTaskID) lastSubTaskID = n;
//     });
//   }
//   Logger.log("Resuming SubTaskID from: " + lastSubTaskID);

//   // ── Doer email lookup (from Doer List sheet) ─────────────
//   var doerSheet    = ss.getSheetByName("Doer List");
//   var doerEmailMap = {};
//   if (doerSheet) {
//     var doerData = doerSheet.getRange(1, 1, doerSheet.getLastRow(), 3).getValues();
//     doerData.forEach(function(row) {
//       if (row[0]) doerEmailMap[String(row[0]).trim()] = String(row[2]).trim(); // col A=name, col C=email
//     });
//   }

//   // ── Main processing ──────────────────────────────────────
//   var masterArray = [];
//   var statusUpdates = [];           // track which Task_List rows to mark "Sent"
//   var START_TIME   = Date.now();
//   var TIME_LIMIT_MS = 5 * 60 * 1000; // 5 min safety limit (script limit = 6 min)

//   for (var i = 0; i < taskListData.length; i++) {

//     // ── Time-guard: if approaching limit, stop gracefully ──
//     if (Date.now() - START_TIME > TIME_LIMIT_MS) {
//       Logger.log("WARNING: Approaching time limit. Processed " + i + " of " + taskListData.length + " tasks. Re-run to continue.");
//       SpreadsheetApp.getUi().alert(
//         "Script is approaching the time limit.\n" +
//         "Processed " + i + " of " + taskListData.length + " tasks.\n" +
//         "Please run again to continue — it will resume from where it stopped."
//       );
//       break;
//     }

//     var task = taskListData[i];

//     var theTask       = task[0];
//     var theDoer       = task[1];
//     var theDepartment = task[2];
//     var theHow        = task[3];
//     var theDetails    = task[4];
//     var theTime       = task[5];
//     var theMobile     = task[6];
//     var theFreq       = task[7];
//     var theDate       = task[8];   // must be a Date object from Sheets
//     var theTasklistID = task[9];
//     var theStatus     = task[10];

//     if (theStatus === "Sent") continue;  // already processed

//     // Look up email
//     var theEmail = doerEmailMap[String(theDoer).trim()] || "";
//     if (!theEmail) {
//       Logger.log("WARN: No email for doer '" + theDoer + "' — row " + (i + 2) + " skipped.");
//       taskListSheet.getRange(i + 2, 11).setValue("Skipped - No email found for doer");
//       continue;
//     }

//     // Ensure theDate is a real Date
//     if (!(theDate instanceof Date) || isNaN(theDate.getTime())) {
//       Logger.log("WARN: Invalid date in row " + (i + 2) + " — skipped.");
//       taskListSheet.getRange(i + 2, 11).setValue("Skipped - Invalid date");
//       continue;
//     }

//     var rowsForThisTask = buildRowsForFrequency(
//       theFreq, theDate, calendarLastDate, workingDatesStr, skipSundays,
//       theDoer, theEmail, theDepartment, theTasklistID,
//       theTask, theHow, theDetails, theTime, theMobile,
//       lastSubTaskID
//     );

//     if (rowsForThisTask.length > 0) {
//       // rowsForThisTask already has sequential SubTaskIDs starting from lastSubTaskID+1
//       lastSubTaskID = rowsForThisTask[rowsForThisTask.length - 1][4]; // last SubTaskID used
//       masterArray = masterArray.concat(rowsForThisTask);
//     }

//     statusUpdates.push(i + 2); // row number in sheet (1-indexed, header is row 1)
//   }

//   // ── Write to Master ───────────────────────────────────────
//   if (masterArray.length > 0) {
//     var writeStartRow = masterLastRow >= 2 ? masterLastRow + 1 : 2;
//     masterSheet.getRange(writeStartRow, 1, masterArray.length, 15).setValues(masterArray);
//     Logger.log("Written " + masterArray.length + " rows to Master starting at row " + writeStartRow);
//   } else {
//     Logger.log("No new rows to add to Master.");
//   }

//   // ── Mark Task_List rows as Sent ───────────────────────────
//   statusUpdates.forEach(function(rowNum) {
//     taskListSheet.getRange(rowNum, 11).setValue("Sent");
//   });

//   Logger.log("=== createChecklist DONE. Total rows added: " + masterArray.length + " ===");
// }


// // ─────────────────────────────────────────────────────────────
// //  BUILD ROWS FOR A GIVEN FREQUENCY
// //  Returns array of 15-column rows, each with its own SubTaskID
// // ─────────────────────────────────────────────────────────────
// function buildRowsForFrequency(
//   freq, startDate, calendarLastDate, workingDatesStr, skipSundays,
//   doer, email, dept, tasklistID,
//   task, how, details, time, mobile,
//   currentLastSubTaskID
// ) {
//   var rows = [];
//   var counter = currentLastSubTaskID;

//   // Helper: build one master row
//   function makeRow(date) {
//     counter++;
//     return [
//       doer,           // A: Doer
//       email,          // B: Email
//       dept,           // C: Department
//       tasklistID,     // D: TaskList ID
//       counter,        // E: SubTask ID  ← global sequential
//       freq,           // F: Frequency
//       task,           // G: Task Name
//       how,            // H: How
//       details,        // I: Details
//       time,           // J: Time
//       mobile,         // K: Mobile
//       new Date(date), // L: Planned Date
//       "",             // M: Actual Date
//       "",             // N: Status
//       email           // O: Contact
//     ];
//   }

//   // Helper: format date as yyyy-MM-dd
//   function fmt(d) {
//     return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
//   }

//   // Helper: is this a working day?
//   function isWorkingDay(d) {
//     return workingDatesStr.indexOf(fmt(d)) !== -1;
//   }

//   // Helper: is Sunday?
//   function isSunday(d) {
//     return d.getDay() === 0;
//   }

//   // Helper: shift date backward to nearest working day (skips non-working & optionally Sundays)
//   function shiftToWorkingDay(d) {
//     var safety = 0;
//     while (!isWorkingDay(d) || (skipSundays === "Yes" && isSunday(d))) {
//       d = addDays(d, -1);
//       if (++safety > 60) break; // prevent infinite loop
//     }
//     return d;
//   }

//   // ── Date arithmetic helpers (no library dependency) ──────
//   function addDays(d, n) {
//     var r = new Date(d.getTime());
//     r.setDate(r.getDate() + n);
//     return r;
//   }
//   function addWeeks(d, n)  { return addDays(d, n * 7); }
//   function addMonths(d, n) {
//     var r = new Date(d.getTime());
//     r.setMonth(r.getMonth() + n);
//     return r;
//   }
//   function addYears(d, n) {
//     var r = new Date(d.getTime());
//     r.setFullYear(r.getFullYear() + n);
//     return r;
//   }
//   function dateCompare(a, b) {
//     // returns 1 if a > b, -1 if a < b, 0 if equal
//     if (a.getTime() > b.getTime()) return 1;
//     if (a.getTime() < b.getTime()) return -1;
//     return 0;
//   }

//   // Current date (for Y frequency check)
//   var today = new Date();

//   var theDate = new Date(startDate.getTime()); // working copy

//   // ─── DAILY ───────────────────────────────────────────────
//   if (freq === "D") {
//     while (dateCompare(calendarLastDate, theDate) === 1) {
//       if (skipSundays === "Yes") {
//         if (isWorkingDay(theDate) && !isSunday(theDate)) rows.push(makeRow(theDate));
//       } else {
//         if (isWorkingDay(theDate)) rows.push(makeRow(theDate));
//       }
//       theDate = addDays(theDate, 1);
//     }

//   // ─── WEEKLY ──────────────────────────────────────────────
//   } else if (freq === "W") {
//     while (dateCompare(calendarLastDate, theDate) === 1) {
//       var frozen = new Date(theDate.getTime());
//       theDate = shiftToWorkingDay(theDate);
//       rows.push(makeRow(theDate));
//       theDate = addWeeks(frozen, 1);
//     }

//   // ─── FORTNIGHTLY ─────────────────────────────────────────
//   } else if (freq === "F") {
//     while (dateCompare(calendarLastDate, theDate) === 1) {
//       var frozen = new Date(theDate.getTime());
//       theDate = shiftToWorkingDay(theDate);
//       rows.push(makeRow(theDate));
//       theDate = addWeeks(frozen, 2);
//     }

//   // ─── 26-DAY ──────────────────────────────────────────────
//   } else if (freq === "26D") {
//     while (dateCompare(calendarLastDate, theDate) === 1) {
//       var frozen = new Date(theDate.getTime());
//       if (skipSundays === "Yes") {
//         if (isWorkingDay(theDate) && !isSunday(theDate)) rows.push(makeRow(theDate));
//       } else {
//         if (isWorkingDay(theDate)) rows.push(makeRow(theDate));
//       }
//       theDate = addDays(frozen, 26);
//     }

//   // ─── MONTHLY ─────────────────────────────────────────────
//   } else if (freq === "M") {
//     while (dateCompare(calendarLastDate, theDate) === 1) {
//       var frozen = new Date(theDate.getTime());

//       // Special case: January dates > 28 skip to March (original logic preserved)
//       if (theDate.getMonth() === 0 && theDate.getDate() > 28) {
//         var adjusted = shiftToWorkingDay(new Date(theDate.getTime()));
//         rows.push(makeRow(adjusted));
//         theDate = addMonths(frozen, 2);
//       } else {
//         var adjusted = shiftToWorkingDay(new Date(theDate.getTime()));
//         rows.push(makeRow(adjusted));
//         theDate = addMonths(frozen, 1);
//       }
//     }

//   // ─── QUARTERLY ───────────────────────────────────────────
//   } else if (freq === "Q") {
//     while (dateCompare(calendarLastDate, theDate) === 1) {
//       var frozen = new Date(theDate.getTime());
//       theDate = shiftToWorkingDay(theDate);
//       rows.push(makeRow(theDate));
//       theDate = addMonths(frozen, 3);
//     }

//   // ─── HALF-YEARLY ─────────────────────────────────────────
//   } else if (freq === "HY") {
//     while (dateCompare(calendarLastDate, theDate) === 1) {
//       var frozen = new Date(theDate.getTime());
//       theDate = shiftToWorkingDay(theDate);
//       rows.push(makeRow(theDate));
//       theDate = addMonths(frozen, 6);
//     }

//   // ─── YEARLY ──────────────────────────────────────────────
//   } else if (freq === "Y") {
//     // Push current year if in future, then 3 more years ahead
//     if (theDate.getTime() > today.getTime()) {
//       rows.push(makeRow(theDate));
//     }
//     for (var y = 1; y <= 3; y++) {
//       theDate = addYears(new Date(startDate.getTime()), y);
//       rows.push(makeRow(theDate));
//     }

//   // ─── ONE-TIME ─────────────────────────────────────────────
//   } else if (freq === "O") {
//     if (skipSundays === "Yes") {
//       if (isWorkingDay(theDate) && !isSunday(theDate)) rows.push(makeRow(theDate));
//     } else {
//       if (isWorkingDay(theDate)) rows.push(makeRow(theDate));
//     }

//   // ─── EVERY 1st/2nd/3rd/4th/Last WEEKDAY OF MONTH ─────────
//   } else if (freq === "E1st" || freq === "E2nd" || freq === "E3rd" || freq === "E4th" || freq === "ELast") {
//     while (dateCompare(calendarLastDate, theDate) === 1) {
//       var frozen = new Date(theDate.getTime());
//       var adjusted = shiftToWorkingDay(new Date(theDate.getTime()));
//       rows.push(makeRow(adjusted));
//       theDate = getNextNthWeekday(frozen, freq);
//     }

//   } else {
//     Logger.log("WARN: Unknown frequency '" + freq + "' — skipped.");
//   }

//   return rows;
// }


// // ─────────────────────────────────────────────────────────────
// //  GET NEXT Nth WEEKDAY OF FOLLOWING MONTH
// //  Replaces eval("Date.parse(frozenDate).next().month().first().monday()")
// // ─────────────────────────────────────────────────────────────
// function getNextNthWeekday(fromDate, freq) {
//   var dayOfWeek = fromDate.getDay(); // 0=Sun, 1=Mon ... 6=Sat

//   // Move to 1st of next month
//   var nextMonth = new Date(fromDate.getFullYear(), fromDate.getMonth() + 1, 1);

//   var occurrences = [];
//   var d = new Date(nextMonth.getTime());
//   var daysInMonth = new Date(nextMonth.getFullYear(), nextMonth.getMonth() + 1, 0).getDate();

//   for (var day = 1; day <= daysInMonth; day++) {
//     d = new Date(nextMonth.getFullYear(), nextMonth.getMonth(), day);
//     if (d.getDay() === dayOfWeek) occurrences.push(new Date(d.getTime()));
//   }

//   if (occurrences.length === 0) return fromDate; // fallback

//   if (freq === "E1st")  return occurrences[0];
//   if (freq === "E2nd")  return occurrences[1] || occurrences[0];
//   if (freq === "E3rd")  return occurrences[2] || occurrences[0];
//   if (freq === "E4th")  return occurrences[3] || occurrences[0];
//   if (freq === "ELast") return occurrences[occurrences.length - 1];

//   return occurrences[0];
// }


// // ─────────────────────────────────────────────────────────────
// //  UTILITY: Get last row with data in a given column
// // ─────────────────────────────────────────────────────────────
// function getLastDataRow(sheet, col) {
//   var colValues = sheet.getRange(col + ":" + col).getValues();
//   for (var i = colValues.length - 1; i >= 0; i--) {
//     if (colValues[i][0] !== "") return i + 1;
//   }
//   return 1;
// }


// // ─────────────────────────────────────────────────────────────
// //  ALL OTHER ORIGINAL FUNCTIONS — UNCHANGED
// // ─────────────────────────────────────────────────────────────

// function onOpen() {
//   Logger.log("Onopen trigger");
//   var ui = SpreadsheetApp.getUi();
//   ui.createMenu('Ultimate Checklist')
//     .addItem('Setup Sheet', 'onOpen')
//     .addToUi();
//   Logger.log('Menu Created');
// }

// function checkMasterSize() {
//   var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master_New");
//   Logger.log("Last Row: " + sheet.getLastRow());
//   Logger.log("Last Column: " + sheet.getLastColumn());
// }

// Date.prototype.addTheDays = function(days) {
//   var date = new Date(this.valueOf());
//   date.setDate(date.getDate() + days);
//   return date;
// };

// function setTrigger() {
//   removeTrigger();
//   var ss    = SpreadsheetApp.getActiveSpreadsheet();
//   var sheet = ss.getSheetByName("Setup Sheet");
//   var time  = sheet.getRange("C15").getValue();
//   createTrigger(time);
// }

// function createTrigger(time) {
//   if (!time) time = 10;
//   ScriptApp.newTrigger('sendReminder')
//     .timeBased()
//     .everyDays(1)
//     .atHour(time)
//     .create();
// }

// function removeTrigger() {
//   var allTriggers = ScriptApp.getProjectTriggers();
//   for (var i = 0; i < allTriggers.length; i++) {
//     if (allTriggers[i].getUniqueId() == allTriggers[i].getUniqueId()) {
//       ScriptApp.deleteTrigger(allTriggers[i]);
//       break;
//     }
//   }
// }

// function sendReminder() {
//   var forB   = SpreadsheetApp.getActive();
//   var forBsh = forB.getSheetByName("Doer List");
//   var lastB  = forBsh.getLastRow();
//   var dataB  = forBsh.getRange(1, 3, lastB, 1).getValues();
//   var today  = new Date();
//   var tomorrow = new Date((new Date().setHours(0, 0, 0, 0)).valueOf() + 1000 * 3600 * 24);
//   var ss     = SpreadsheetApp.getActive();
//   var sheet  = ss.getSheetByName("Master");
//   var columnToCheck = sheet.getRange("B:B").getValues();
//   var lastrow = getLastRowSpecial(columnToCheck) - 1;
//   var data   = sheet.getRange(2, 1, lastrow, 8).getValues();

//   dataB.filter(function(r) {
//     var reminderTasks = [];
//     data.filter(function(pending) {
//       var doeremail = pending[1];
//       var lastDate  = pending[6];
//       var actual    = pending[7];
//       var name      = pending[0];
//       var task      = pending[5];
//       if (r[0] === doeremail && lastDate.valueOf() === tomorrow.valueOf() && !actual) {
//         reminderTasks.push([name, task]);
//       }
//     });
//     if (reminderTasks.length > 0) {
//       var joinedTasks = reminderTasks.map(function(t) { return "Task : " + t[1]; }).join("\n");
//       var dataToSend  = "Hello " + reminderTasks[0][0] + ",\n\nYou have planned tasks pending for tomorrow.\n\n" + joinedTasks + "\n\nPlease ignore this message if you have already completed the tasks.";
//       GmailApp.sendEmail(r[0], "You have a Pending Tasks for tomorrow", dataToSend);
//     }
//   });
// }

// function archive() {
//   var ss   = SpreadsheetApp.getActive();
//   var sheet = ss.getSheetByName("Dashboard");
//   var data  = sheet.getRange(4, 1, 6, 4).getValues();
//   var name  = sheet.getRange("A2").getValue();
//   var week  = sheet.getRange("D2").getValue();
//   var arch  = ss.getSheetByName("Archive");
//   var ared    = data[3][1];
//   var ayellow = data[3][2];
//   var pred    = data[5][1];
//   var pyellow = data[5][2];
//   arch.appendRow([name, week, pred, pyellow, "", ared, ayellow, ""]);
//   var spreadsheet = ss.getSheetByName("Dashboard");
//   spreadsheet.getRange('B9:C9').activate();
//   spreadsheet.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });
// }

// function sendtodoer() {
//   var forB   = SpreadsheetApp.getActive();
//   var forBsh = forB.getSheetByName("Doer List");
//   var lastB  = forBsh.getLastRow();
//   var dataB  = forBsh.getRange(1, 1, lastB, 2).getValues();
//   var ss     = SpreadsheetApp.getActiveSpreadsheet();
//   var sheet  = ss.getSheetByName("Send Tasks To Doer");
//   var lastrow = sheet.getLastRow() - 3;
//   var range  = sheet.getRange(4, 1, lastrow, 2);
//   var htmlTable = SheetConverter.convertRange2html(range);
//   var status  = sheet.getRange("D2").getValue();
//   var weeknum = sheet.getRange("E2").getValue().toString();
//   var name    = sheet.getRange("A2").getValue();
//   var email   = findinB(name, dataB);
//   var body    = "Here are your " + status + " tasks of Week " + weeknum + ".<br/><br/>" + htmlTable;
//   GmailApp.sendEmail(email, 'Details of Delegated Tasks', body, { htmlBody: body });
// }

// function onChange() {
//   var ss     = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
//   var active = ss.getActiveCell();
//   var val    = active.getValue();
//   if (val === "Done") {
//     if (ss.getSheetName() === "Master") {
//       var row  = active.getRow();
//       var col  = active.getColumn();
//       var date = new Date();
//       var ts   = ss.getRange(row, col - 1).getValue();
//       if (!ts) {
//         ss.getRange(row, col - 1).setValue(date);
//       }
//     }
//   }
// }

// function createTriggerOnChange() {
//   ScriptApp.newTrigger('onChange')
//     .forSpreadsheet(SpreadsheetApp.getActive())
//     .onChange()
//     .create();
// }

// function copyTasksToMasterNew() {
//   Logger.log(">>> STARTING BASIC COPY");
//   var startTime = Date.now();
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var taskListSheet  = ss.getSheetByName("Task_List");
//   var masterNewSheet = ss.getSheetByName("Master_New") || ss.insertSheet("Master_New");
//   var lastRow = taskListSheet.getLastRow();
//   Logger.log("LOG: Last row in Task_List: " + lastRow);
//   if (lastRow <= 1) { Logger.log("ABORTING: No data found."); return; }
//   var taskListData = taskListSheet.getRange(2, 1, lastRow - 1, 10).getValues();
//   Logger.log("LOG: Data loaded into memory.");
//   var masterArray = taskListData.map(function(row) {
//     return [row[1],"",row[2],row[9],"",row[7],row[0],row[3],row[4],row[5],row[6],row[8],"","",""];
//   });
//   Logger.log("LOG: Writing to Master_New...");
//   masterNewSheet.clearContents();
//   var headers = [["Doer","Email","Dept","TaskID","SubTaskID","Freq","Task","How","Details","Time","Mobile","Planned Date","Actual Date","Status","Contact"]];
//   masterNewSheet.getRange(1, 1, 1, 15).setValues(headers);
//   if (masterArray.length > 0) {
//     masterNewSheet.getRange(2, 1, masterArray.length, 15).setValues(masterArray);
//   }
//   var duration = (Date.now() - startTime) / 1000;
//   Logger.log(">>> SUCCESS: Finished in " + duration + " seconds.");
// }


// function onOpen() {
//   Logger.log("Onopen trigger")
//   var ui = SpreadsheetApp.getUi();
//   ui.createMenu('Ultimate Checklist')
//     .addItem('Setup Sheet', 'onOpen')
//     .addToUi();
//     Logger.log('Menu Created')
// }

// Date.prototype.addTheDays = function (days) {
//   var date = new Date(this.valueOf());
//   date.setDate(date.getDate() + days);
//   return date;
// }

// function setTrigger() {
//   removeTrigger()
//   var ss = SpreadsheetApp.getActiveSpreadsheet()
//   var sheet = ss.getSheetByName("Setup Sheet")
//   var time = sheet.getRange("C15").getValue()
//   createTrigger(time)
// }
// function createTrigger(time) {
//   if (!time) {
//     time = 10
//   }
//   ScriptApp.newTrigger('sendReminder')
//     .timeBased()
//     .everyDays(1)
//     .atHour(time)
//     .create();
// }


// function removeTrigger() {
//   // Loop over all triggers.
//   var allTriggers = ScriptApp.getProjectTriggers();
//   for (var i = 0; i < allTriggers.length; i++) {
//     // If the current trigger is the correct one, delete it.
//     if (allTriggers[i].getUniqueId() == allTriggers[i].getUniqueId()) {
//       ScriptApp.deleteTrigger(allTriggers[i]);
//       break;
//     }
//   }
// }


// function sendReminder() {
//   var forB = SpreadsheetApp.getActive();
//   var forBsh = forB.getSheetByName("Doer List")
//   var lastB = forBsh.getLastRow()
//   var dataB = forBsh.getRange(1, 3, lastB, 1).getValues()
//   var today = new Date();
//   var tomorrow = new Date((new Date().setHours(0, 0, 0, 0)).valueOf() + 1000 * 3600 * 24);
//   var ss = SpreadsheetApp.getActive()
//   var sheet = ss.getSheetByName("Master")
//   var columnToCheck = sheet.getRange("B:B").getValues()
//   var lastrow = getLastRowSpecial(columnToCheck) - 1
//   var data = sheet.getRange(2, 1, lastrow, 8).getValues()
//   let emails = dataB.filter(function(r){
//     let reminderTasks = []
//     let pendingTomorrow = data.filter(function(pending){
//       let doeremail = pending[1]
//       let lastDate = pending[6]
//       let actual = pending[7]
//       let name = pending[0]
//       let task = pending[5]
//       if (r[0] === doeremail && lastDate.valueOf() === tomorrow.valueOf() && !actual){
//         reminderTasks.push([name,task])
//       }
//     })
//     if (reminderTasks.length > 0){
//     let joinedTasks = reminderTasks.map(function(tasks){
//             return "Task : "+tasks[1]
//             }
//             ).join("\n")
//     let dataToSend = "Hello "+reminderTasks[0][0]+",\n\nYou have planned tasks pending for tomorrow.\n\n"+joinedTasks+"\n\nPlease ignore this message if you have already completed the tasks."
//     GmailApp.sendEmail(r[0], "You have a Pending Tasks for tomorrow", dataToSend);
//     }
//   })
// }



// function archive() {
//   var ss = SpreadsheetApp.getActive();
//   var sheet = ss.getSheetByName("Dashboard")
//   var data = sheet.getRange(4, 1, 6, 4).getValues()
//   var name = sheet.getRange("A2").getValue()
//   var week = sheet.getRange("D2").getValue()
//   var arch = ss.getSheetByName("Archive")
//   var archLast = arch.getLastRow() - 1
//   var ared = data[3][1]
//   var ayellow = data[3][2]
//   var pred = data[5][1]
//   var pyellow = data[5][2]
//   arch.appendRow([name, week, pred, pyellow, "", ared, ayellow, ""])
//   var spreadsheet = ss.getSheetByName("Dashboard")
//   spreadsheet.getRange('B9:C9').activate();
//   spreadsheet.getActiveRangeList().clear({ contentsOnly: true, skipFilteredRows: true });
// }
// function sendtodoer() {
//   var forB = SpreadsheetApp.getActive();
//   var forBsh = forB.getSheetByName("Doer List")
//   var lastB = forBsh.getLastRow()
//   var dataB = forBsh.getRange(1, 1, lastB, 2).getValues()
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var sheet = ss.getSheetByName("Send Tasks To Doer")
//   var lastrow = sheet.getLastRow() - 3
//   var range = sheet.getRange(4, 1, lastrow, 2)
//   var htmlTable = SheetConverter.convertRange2html(range);
//   var status = sheet.getRange("D2").getValue()
//   var weeknum = sheet.getRange("E2").getValue().toString()
//   var name = sheet.getRange("A2").getValue()
//   var email = findinB(name, dataB)
//   var body = "Here are your " + status + " tasks of Week " + weeknum + ".<br/><br/>" + htmlTable
//   GmailApp.sendEmail(email, 'Details of Delegated Tasks', body, { htmlBody: body });
// }


// function createChecklist() {
//   let ss = SpreadsheetApp.openById("1yRgq27UU_FDVdoGrEEXGa8xF0YtajBG1CNgGOcycPvg");
//   let taskListSheet = ss.getSheetByName("Task_List");
//   let taskListData = taskListSheet.getRange(2, 1, taskListSheet.getLastRow() - 1, 11).getValues()
//   let skipSundays = ss.getSheetByName("Setup Sheet").getRange("B32").getValue()
//   let todaysdate = new Date()

//   if (!skipSundays) {
//     skipSundays = "Yes"
//   }

//   let masterSheet = ss.getSheetByName("Master")
//   let cell = masterSheet.getRange("M2") //actual Cell
//   let masterTaskIDs = masterSheet.getRange("e:e").getValues()
//   let masterLastRow = masterTaskIDs.filter(r => String(r)).length

//   let lastMasterTaskID = Number(masterTaskIDs[masterLastRow - 1])
//   if (!lastMasterTaskID) {
//     lastMasterTaskID = 0
//   }

//   let masterArray = []

//   let calendarSheet = ss.getSheetByName("Working Day Calender")
//   let calendarDates = calendarSheet.getRange("A:A").getValues()
//   let calendarLast = calendarDates.filter(String).length
//   let allCalendarDates = calendarSheet.getRange(2, 1, calendarLast - 1, 1).getValues()
//   let calendarLastDate = calendarDates[calendarLast - 1][0]
//   let workingDates = allCalendarDates.flat()
//   let workingDatesStr = workingDates.map(x => Utilities.formatDate(x, Session.getScriptTimeZone(), "yyyy-MM-dd"));

//   taskListData.forEach(function (task, index) {
//     let theTask = task[0]
//     let theDoer = task[1]
//     let theDepartment = task[2]
//     let theHow = task[3]
//     let theDetails = task[4]
//     let theTime = task[5]
//     let theMobile = task[6]
//     let theFreq = task[7]
//     let theDate = task[8]
//     let theTasklistID = task[9]
//     let theStatus = task[10]
//     let theEmail = getEmail(theDoer,theDepartment,theHow,theDetails,theTime,theMobile,theFreq,theDate, theTasklistID)

//     if (theEmail === "FAILED") {
//       taskListSheet.getRange(index + 2, 11).setValue("Skipped Due to doer - department - how - details - time - mobile - freq - date - tasklistID mismatch")
//       return
//     }


//     if (theStatus != "Sent") {
//       //weekly tasks 
//       if (theFreq === "W") {
//         while (Date.compare(calendarLastDate, theDate) === 1) {
//           lastMasterTaskID = lastMasterTaskID + 1
//           let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

//           let frozenDate = theDate.clone()

//           while (!workingDatesStr.includes(endDateStr)) {
//             theDate = Date.parse(theDate).addDays(-1)
//             endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
//           }

//           masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate),"","",theEmail])

//           theDate = Date.parse(frozenDate).addWeeks(1)
//         }
//       } else if (theFreq === "M") {
//         while (Date.compare(calendarLastDate, theDate) === 1) {
//           let norun = 0
//           if(theDate.is().jan() && theDate.getDate() > 28){
//             norun = 1
//             lastMasterTaskID = lastMasterTaskID + 1
//           let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

//           let frozenDate = theDate.clone()

//           while (!workingDatesStr.includes(endDateStr)) {
//             theDate = Date.parse(theDate).addDays(-1)
//             endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
//           }

//           if (skipSundays === "Yes") {
//             if (theDate.is().sunday()) {
//               theDate = Date.parse(theDate).addDays(-1)
//             }
//           }

//           masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate),"","",theEmail])

//           theDate = Date.parse(frozenDate).addMonths(2)
//           }

//           if (norun === 0){
//           lastMasterTaskID = lastMasterTaskID + 1
//           let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

//           let frozenDate = theDate.clone()

//           while (!workingDatesStr.includes(endDateStr)) {
//             theDate = Date.parse(theDate).addDays(-1)
//             endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
//           }

//           if (skipSundays === "Yes") {
//             if (theDate.is().sunday()) {
//               theDate = Date.parse(theDate).addDays(-1)
//             }
//           }

//           masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate),"","",theEmail])

//           theDate = Date.parse(frozenDate).addMonths(1)
//           }
//         }
//       } else if (theFreq === "Y") {
//         if (theDate.valueOf() > todaysdate.valueOf()){
//           lastMasterTaskID = lastMasterTaskID + 1
//           masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate),"","",theEmail])
//         }
//         lastMasterTaskID = lastMasterTaskID + 1
//         theDate = Date.parse(theDate).addYears(1)
//         masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate),"","",theEmail])
//         lastMasterTaskID = lastMasterTaskID + 1
//         theDate = Date.parse(theDate).addYears(1)
//         masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate),"","",theEmail])
//         lastMasterTaskID = lastMasterTaskID + 1
//         theDate = Date.parse(theDate).addYears(1)
//         masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate),"","",theEmail])

//       } else if (theFreq === "D") {
//         while (Date.compare(calendarLastDate, theDate) === 1) {
//           lastMasterTaskID = lastMasterTaskID + 1
//           let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

//           let frozenDate = theDate.clone()

//           if (skipSundays === "Yes") {
//             if (workingDatesStr.includes(endDateStr) && !theDate.is().sunday()) {
//               masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate),"","",theEmail])
//             }
//           } else {
//             if (workingDatesStr.includes(endDateStr)) {
//               masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate),"","",theEmail])
//             }
//           }

//           theDate = Date.parse(frozenDate).addDays(1)
//         }
//       } else if (theFreq === "O") {
//   let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

//   if (skipSundays === "Yes") {
//     if (workingDatesStr.includes(endDateStr) && !theDate.is().sunday()) {
//       lastMasterTaskID += 1;
//       masterArray.push([
//         theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID,
//         theFreq, theTask, theHow, theDetails, theTime, theMobile,
//         new Date(theDate), "", "", theEmail
//       ]);
//     }
//   } else {
//     if (workingDatesStr.includes(endDateStr)) {
//       lastMasterTaskID += 1;
//       masterArray.push([
//         theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID,
//         theFreq, theTask, theHow, theDetails, theTime, theMobile,
//         new Date(theDate), "", "", theEmail
//       ]);
//     }
//   }
// } else if (theFreq === "26D") {
//         while (Date.compare(calendarLastDate, theDate) === 1) {
//           lastMasterTaskID = lastMasterTaskID + 1
//           let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

//           let frozenDate = theDate.clone()

//           if (skipSundays === "Yes") {
//             if (workingDatesStr.includes(endDateStr) && !theDate.is().sunday()) {
//               masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate),"","",theEmail])
//             }
//           } else {
//             if (workingDatesStr.includes(endDateStr)) {
//               masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate),"","",theEmail])
//             }
//           }

//           theDate = Date.parse(frozenDate).addDays(26)
//         }
//       } else if (theFreq === "Q") {
//         while (Date.compare(calendarLastDate, theDate) === 1) {
//           lastMasterTaskID = lastMasterTaskID + 1
//           let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

//           let frozenDate = theDate.clone()

//           while (!workingDatesStr.includes(endDateStr)) {
//             theDate = Date.parse(theDate).addDays(-1)
//             endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
//           }

//           if (skipSundays === "Yes") {
//             if (theDate.is().sunday()) {
//               theDate = Date.parse(theDate).addDays(-1)
//             }
//           }

//           masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate),"","",theEmail])

//           theDate = Date.parse(frozenDate).addMonths(3)
//         }
//       } else if (theFreq === "HY") {
//         while (Date.compare(calendarLastDate, theDate) === 1) {
//           lastMasterTaskID = lastMasterTaskID + 1
//           let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

//           let frozenDate = theDate.clone()

//           while (!workingDatesStr.includes(endDateStr)) {
//             theDate = Date.parse(theDate).addDays(-1)
//             endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
//           }

//           if (skipSundays === "Yes") {
//             if (theDate.is().sunday()) {
//               theDate = Date.parse(theDate).addDays(-1)
//             }
//           }

//           masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate),"","",theEmail])

//           theDate = Date.parse(frozenDate).addMonths(6)
//         }
//       } else if (theFreq === "F") {
//         while (Date.compare(calendarLastDate, theDate) === 1) {
//           lastMasterTaskID = lastMasterTaskID + 1
//           let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

//           let frozenDate = theDate.clone()

//           while (!workingDatesStr.includes(endDateStr)) {
//             theDate = Date.parse(theDate).addDays(-1)
//             endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
//           }

//           masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate),"","",theEmail])

//           theDate = Date.parse(frozenDate).addWeeks(2)
//         }
//       }else if (theFreq === "E1st") {
//         while (Date.compare(calendarLastDate, theDate) === 1) {
//           lastMasterTaskID = lastMasterTaskID + 1
//           let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

//           let day = getDayInfo(theDate.getDay())

//           let frozenDate = theDate.clone()

//           while (!workingDatesStr.includes(endDateStr)) {
//             theDate = Date.parse(theDate).addDays(-1)
//             endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
//           }

//           masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate),"","",theEmail])

//           theDate = eval("Date.parse(frozenDate).next().month().first()."+day+"()")
//         }
//       } else if (theFreq === "E2nd") {
//         while (Date.compare(calendarLastDate, theDate) === 1) {
//           lastMasterTaskID = lastMasterTaskID + 1
//           let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

//           let day = getDayInfo(theDate.getDay())

//           let frozenDate = theDate.clone()

//           while (!workingDatesStr.includes(endDateStr)) {
//             theDate = Date.parse(theDate).addDays(-1)
//             endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
//           }

//           masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate),"","",theEmail])

//           theDate = eval("Date.parse(frozenDate).next().month().second()."+day+"()")
//         }
//       } else if (theFreq === "E3rd") {
//         while (Date.compare(calendarLastDate, theDate) === 1) {
//           lastMasterTaskID = lastMasterTaskID + 1
//           let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

//           let day = getDayInfo(theDate.getDay())

//           let frozenDate = theDate.clone()

//           while (!workingDatesStr.includes(endDateStr)) {
//             theDate = Date.parse(theDate).addDays(-1)
//             endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
//           }

//           masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate),"","",theEmail])

//           theDate = eval("Date.parse(frozenDate).next().month().third()."+day+"()")
//         }
//       } else if (theFreq === "E4th") {
//         while (Date.compare(calendarLastDate, theDate) === 1) {
//           lastMasterTaskID = lastMasterTaskID + 1
//           let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

//           let day = getDayInfo(theDate.getDay())

//           let frozenDate = theDate.clone()

//           while (!workingDatesStr.includes(endDateStr)) {
//             theDate = Date.parse(theDate).addDays(-1)
//             endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
//           }

//           masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate),"","",theEmail])

//           theDate = eval("Date.parse(frozenDate).next().month().fourth()."+day+"()")
//         }
//       }else if (theFreq === "ELast") {
//         while (Date.compare(calendarLastDate, theDate) === 1) {
//           lastMasterTaskID = lastMasterTaskID + 1
//           let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

//           let day = getDayInfo(theDate.getDay())

//           let frozenDate = theDate.clone()

//           while (!workingDatesStr.includes(endDateStr)) {
//             theDate = Date.parse(theDate).addDays(-1)
//             endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
//           }

//           masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate),"","",theEmail])

//           theDate = eval("Date.parse(frozenDate).next().month().final()."+day+"()")
//         }
//       }

//       taskListSheet.getRange(index + 2, 11).setValue("Sent")
//     }

//   })
//   if (masterArray.length > 0) {
//     masterSheet.getRange(masterLastRow + 1, 1, masterArray.length, 15).setValues(masterArray)
//   }
// }

// function onChange() 
// {
//   var ss=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
//   var active=ss.getActiveCell();
//   var val=active.getValue();
  
//   if (val === "Done") {
    
//     if (ss.getSheetName() === "Master") {
//       var row = active.getRow();
//       var col = active.getColumn();
 
//       var date = new Date();
 
//       ts = ss.getRange(row, col - 1).getValue();//Actual Time column
//       if (!ts) {
//         ss.getRange(row, col - 1).setValue(date);
//       }
//     }
//   }
// }

// function createTriggerOnChange() {
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   ScriptApp.newTrigger('onChange')
//       .forSpreadsheet(SpreadsheetApp.getActive())
//       .onChange()
//       .create();
// }

// function getEmail(theName,theDepartment) {
//   try {
//   let ss = SpreadsheetApp.getActiveSpreadsheet()
//   let sheet = ss.getSheetByName("Doer List")
//   let data = sheet.getRange(1,1,sheet.getLastRow(),3).getValues()
//   let email = data.filter(r => r[0] === theName && r[1] === theDepartment)[0][2]
//   return email
//   } catch (e) {
//     Browser.msgBox("There is a mismatch in Doer Details in Task list: Doer Name - " + theName + ", Doer Department - " + theDepartment)
//     return "FAILED"
//   }
// }
