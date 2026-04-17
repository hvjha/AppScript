function generateTaskIDs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Task_List");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();  // Column J = index 9
  var idArr = [];
  // Get last existing ID (for global continuity)
  var lastID = 0;
  data.forEach(function(row) {
    var existing = row[9]; // Column J
    if (existing && existing.toString().startsWith("TL")) {
      var num = parseInt(existing.replace("TL", ""));
      if (!isNaN(num) && num > lastID) lastID = num;
    }
  });
  // Generate IDs
  for (var i = 0; i < data.length; i++) {
    var task = data[i][0];
    var existingID = data[i][9];
    if (!task) {
      idArr.push([""]);
      continue;
    }
    if (existingID) {
      idArr.push([existingID]); // keep existing
    } else {
      lastID++;
      idArr.push(["TL" + lastID]); // new global ID
    }
  }
  // Write back to Task_List
  sheet.getRange(2, 10, idArr.length, 1).setValues(idArr);
}







/*ULTIMATE CHECKLIST — COMPLETE FINAL SCRIPT 
Removed Setup Sheet dependency entirely
Removed Date.js library calls (addDays, addWeeks, addMonths, addYears, clone, is().sunday(), compare() — all replaced with 
pure vanilla JS)
Removed eval() calls for E1st/E2nd/E3rd/E4th/ELast
All sheet.getRange().setValue() calls moved OUT of the loop (single batch write at the end — this is the main timeout fix)
Global sequential SubTaskID with resume support
5-minute time guard with graceful save + alert to re-run
Uses getActiveSpreadsheet() — no hardcoded spreadsheet ID
*/


//CONFIG 
var CONFIG = {
  SKIP_SUNDAYS: true,          // replaces Setup Sheet B32
  TIME_LIMIT_MS: 5 * 60 * 1000 // 5 min guard (Apps Script limit = 6 min)
};

//  createChecklist — REWRITTEN
function createChecklistV1() {
  Logger.log("=== createChecklist START ===");
  var startTime = Date.now();

  var props = PropertiesService.getScriptProperties();
  var startIndex = parseInt(props.getProperty("CHECKLIST_INDEX") || "0", 10);

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── 1. Load all sheets ───────────────────────────────────
  var taskListSheet = ss.getSheetByName("Task_List");
  var masterSheet = ss.getSheetByName("Master_New");
  var calendarSheet = ss.getSheetByName("Working Day Calender");

  if (!taskListSheet) { alertAndLog("Sheet not found: Task_List"); return; }
  if (!masterSheet) { alertAndLog("Sheet not found: Master_New"); return; }
  if (!calendarSheet) { alertAndLog("Sheet not found: Working Day Calender"); return; }

  // ── 2. Load Task_List data ───────────────────────────────
  var taskLastRow = taskListSheet.getLastRow();
  if (taskLastRow < 2) {
    Logger.log("No data in Task_List. Aborting.");
    return;
  }
  // Columns: A=What, B=DoerName, C=Dept, D=How, E=Details,
  //          F=Time, G=Mobile, H=Freq, I=Date, J=TasklistID, K=Status
  var taskData = taskListSheet.getRange(2, 1, taskLastRow - 1, 11).getValues();

  Logger.log("TaskData length: " + taskData.length);
  Logger.log("First row: " + JSON.stringify(taskData[0]));

  // ── 3. Load Working Day Calendar ────────────────────────

  // var calColValues = calendarSheet.getRange("A:A").getValues();
  var calLastRow = calendarSheet.getLastRow();
  var calColValues = calendarSheet.getRange(1, 1, calLastRow, 1).getValues();
  var calNonEmpty = calColValues.filter(function (r) { return r[0] !== "" && r[0] instanceof Date; });

  if (calNonEmpty.length < 2) {
    alertAndLog("Working Day Calender sheet is empty or has no valid dates!");
    return;
  }

  var workingDatesStr = calNonEmpty.slice(1).map(function (r) { return fmtDate(r[0]); });
  var calendarLastDate = calNonEmpty[calNonEmpty.length - 1][0];
  Logger.log("Calendar loaded. Working days: " + workingDatesStr.length + " | Last: " + fmtDate(calendarLastDate));

  // ── 4. Get last SubTaskID from Master (for resume) ───────
  //    SubTaskID is in column E (index 4 = col 5)
  // var masterLastRow = getLastRowSpecial(masterSheet.getRange("B:B").getValues());
  var masterLastRow = masterSheet.getLastRow();
  var lastMasterTaskID = 0;

  if (masterLastRow >= 2) {
    var existingIDs = masterSheet.getRange(2, 5, masterLastRow - 1, 1).getValues();
    existingIDs.forEach(function (r) {
      var n = parseInt(r[0], 10);
      if (!isNaN(n) && n > lastMasterTaskID) lastMasterTaskID = n;
    });
  }
  Logger.log("Resuming SubTaskID from: " + lastMasterTaskID);

  // ── 5. Process each task row ─────────────────────────────
  var masterArray = [];   // all rows to write to Master at the end
  var statusMap = {};   // { rowIndex: "Sent" | "Skipped - reason" }
  var timedOut = false;
  var today = new Date();

  // for (var i = 0; i < taskData.length; i++) {
  var BATCH_SIZE = 35;
  // by this and do one thing directly check 
  for (var i = startIndex; i < taskData.length && i < startIndex + BATCH_SIZE; i++) {
    Logger.log("---- Row " + (i + 2) + " ----");

    // Time guard — save progress and stop if approaching limit
    if (Date.now() - startTime > CONFIG.TIME_LIMIT_MS) {
      Logger.log("WARNING: Time limit at task row " + (i + 2) + ". Will resume next run.");
      timedOut = true;
      break;
    }
    var row = taskData[i];
    var theTask = String(row[0] || "").trim();   // A - What
    var theDoer = String(row[1] || "").trim();   // B - Doer Name
    var theDept = String(row[2] || "").trim();   // C - Department
    var theHow = String(row[3] || "").trim();   // D - How
    var theDetails = String(row[4] || "").trim();   // E - Details
    var theTime = row[5];                        // F - Time
    var theMobile = row[6];                        // G - Mobile
    var theFreq = String(row[7] || "").trim();   // H - Frequencyvar result
    var theDate = row[8];                        // I - Day/Time (Date)
    var theTaskID = row[9];
    var theStatus = String(row[10] || "").trim();  // K - Status

    Logger.log("Task: " + theTask);
    Logger.log("Doer: " + theDoer);
    Logger.log("Freq: " + theFreq);
    Logger.log("Date: " + theDate);
        if (theStatus === "Sent") {
      Logger.log("Already Sent row " + (i + 2));
      continue;
    }

    // Skip completely blank rows
    if (!theTask && !theDoer) continue;

    // Validate date
    if (!(theDate instanceof Date) || isNaN(theDate.getTime())) {
      Logger.log("Invalid Date at row " + (i + 2));
      Logger.log("SKIP row " + (i + 2) + ": invalid date");
      statusMap[i] = "Skipped - Invalid Date";
      continue;
    }

    // Get email via existing helper (Name + Dept lookup in Doer List)
    var emailObj = getEmail(theDoer, theDept);

if (emailObj.status === "FAILED") {
  Logger.log("Failed" + emailObj.reason + " at row " + (i + 2));
  statusMap[i] = "Skipped - " + emailObj.reason;
  continue;
}

var theEmail = emailObj.email;

    // Build all occurrence rows for this task
    var result = buildOccurrences(
      theFreq, new Date(theDate.getTime()), calendarLastDate,
      workingDatesStr, lastMasterTaskID, today,
      theTask, theDoer, theEmail, theDept, theTaskID,
      theHow, theDetails, theTime, theMobile
    );



    if (result.rows.length > 0) {
      lastMasterTaskID = result.lastMasterTaskID;
      masterArray = masterArray.concat(result.rows);
    }

    statusMap[i] = "Sent";
  }

  if (i < taskData.length) {
    props.setProperty("CHECKLIST_INDEX", i);

    safeAlert(
      "Partial run complete.\n\n" +
      "Processed up to row " + (i + 1) + ".\n" +
      "Run again to continue."
    );
  }

  Logger.log("FINAL masterArray length: " + masterArray.length);

  //   masterArray.push([
  //   "Test Name",
  //   "test@email.com",
  //   "Dept",
  //   "T001",
  //   1,
  //   "D",
  //   "Test Task",
  //   "How",
  //   "Details",
  //   "10:00",
  //   "9999999999",
  //   new Date(),
  //   "",
  //   "",
  //   "test@email.com"
  // ]);

  // 6. BATCH write all rows to Master (single call)
  if (masterArray.length > 0) {
    var writeAt = masterLastRow + 1;
    if (writeAt < 2) writeAt = 2;
    var CHUNK = 5000;
    for (var j = 0; j < masterArray.length; j += CHUNK) {
      var chunkData = masterArray.slice(j, j + CHUNK);
      masterSheet.getRange(writeAt + j, 1, chunkData.length, 15).setValues(chunkData);
    }
    Logger.log("Written " + masterArray.length + " rows to Master starting at row " + writeAt);
  } else {
    Logger.log("No new rows to write to Master.");
  }

  // 7. BATCH update Task_List status column ───────────────
  var statusColumnArray = [];
  for (var si = 0; si < taskData.length; si++) {
    if (statusMap.hasOwnProperty(si)) {
      statusColumnArray.push([statusMap[si]]);
    } else {
      statusColumnArray.push([taskData[si][10]]);
    }
  }
  // ONE write call for the entire status column
  taskListSheet.getRange(2, 11, statusColumnArray.length, 1).setValues(statusColumnArray); 
  if (i < taskData.length) {
    safeAlert(
      "Partial run complete.\n\nProcessed up to row " + i + ".\nRun again to continue."
    );
  } else {
    safeAlert("Checklist fully created!");
  }
  if (timedOut) {
    safeAlert(
      "Progress saved!\n\n" +
      "The script hit the 5-minute safety limit.\n" +
      "Click Submit again to continue — SubTaskID will resume from where it stopped."
    );
  } else {
    Logger.log("createChecklist COMPLETE. Total rows added: " + masterArray.length + " ===");
  }
  if (masterArray.length > 0) {
  masterSheet.getRange(masterLastRow + 1, 1, masterArray.length, 15)
             .setValues(masterArray);
  return masterArray.length;
} else {
  return 0;
}
}


function buildOccurrences(
  freq, startDate, calendarLastDate, workingDatesStr,
  currentLastID, today,
  taskName, doerName, email, dept, taskID,
  how, details, time, mobile
) {
  var rows = [];
  var counter = currentLastID;

  function pushRow(d) {
    counter++;
    rows.push([
      doerName, email, dept, taskID, counter, freq,
      taskName, how, details, time, mobile,
      new Date(d), "", "", email
    ]);
  }

  // Date helpers
  function addDays(d, n) {
    var r = new Date(d.getTime());
    r.setDate(r.getDate() + n);
    return r;
  }
  function addWeeks(d, n) { return addDays(d, n * 7); }
  function addMonths(d, n) {
    var r = new Date(d.getTime());
    var day = r.getDate();
    r.setMonth(r.getMonth() + n);
    var maxDay = new Date(r.getFullYear(), r.getMonth() + 1, 0).getDate();
    r.setDate(Math.min(day, maxDay));
    return r;
  }
  function addYears(d, n) {
    var r = new Date(d.getTime());
    r.setFullYear(r.getFullYear() + n);
    return r;
  }
  function cloneDate(d) { return new Date(d.getTime()); }
  function isBefore(a, b) {
    return new Date(a.getFullYear(), a.getMonth(), a.getDate()).getTime() <
      new Date(b.getFullYear(), b.getMonth(), b.getDate()).getTime();
  }
  function isWorkingDay(d) {
    return workingDatesStr.indexOf(fmtDate(d)) !== -1;
  }
  function isSunday(d) { return d.getDay() === 0; }
  function shiftBack(d) {
    var cur = cloneDate(d);
    while (true) {
      if (isWorkingDay(cur) && !(CONFIG.SKIP_SUNDAYS && isSunday(cur))) return cur;
      cur = addDays(cur, -1);
    }
  }

  var d = cloneDate(startDate);

  if (freq === "D") {
    while (isBefore(d, calendarLastDate)) {
      if (isWorkingDay(d) && !(CONFIG.SKIP_SUNDAYS && isSunday(d))) pushRow(d);
      d = addDays(d, 1);
    }
  } else if (freq === "W") {
    while (isBefore(d, calendarLastDate)) {
      pushRow(shiftBack(d));
      d = addWeeks(cloneDate(d), 1);
    }
  } else if (freq === "F") {
    while (isBefore(d, calendarLastDate)) {
      pushRow(shiftBack(d));
      d = addWeeks(cloneDate(d), 2);
    }
  } else if (freq === "26D") {
    while (isBefore(d, calendarLastDate)) {
      if (isWorkingDay(d) && !(CONFIG.SKIP_SUNDAYS && isSunday(d))) pushRow(d);
      d = addDays(cloneDate(d), 26);
    }
  } else if (freq === "M") {
    while (isBefore(d, calendarLastDate)) {
      var frozen = cloneDate(d);
      if (d.getMonth() === 0 && d.getDate() > 28) {
        pushRow(shiftBack(d));
        d = addMonths(frozen, 2);
      } else {
        pushRow(shiftBack(d));
        d = addMonths(frozen, 1);
      }
    }
  } else if (freq === "Q") {
    while (isBefore(d, calendarLastDate)) {
      pushRow(shiftBack(d));
      d = addMonths(cloneDate(d), 3);
    }
  } else if (freq === "HY") {
    while (isBefore(d, calendarLastDate)) {
      pushRow(shiftBack(d));
      d = addMonths(cloneDate(d), 6);
    }
  } else if (freq === "Y") {
    if (d.getTime() > today.getTime()) pushRow(d);
    pushRow(addYears(cloneDate(startDate), 1));
    pushRow(addYears(cloneDate(startDate), 2));
    pushRow(addYears(cloneDate(startDate), 3));
  } else if (freq === "O") {
    if (isWorkingDay(d) && !(CONFIG.SKIP_SUNDAYS && isSunday(d))) pushRow(d);
  } else if (["E1st", "E2nd", "E3rd", "E4th", "ELast"].includes(freq)) {
    while (isBefore(d, calendarLastDate)) {
      pushRow(shiftBack(d));
      d = getNextNthWeekday(cloneDate(d), freq);
    }
  } else {
    Logger.log("WARN: Unknown frequency " + freq);
  }
  return {
    rows: rows,
    lastMasterTaskID: counter
  };
}


// ============================================================
//  getNextNthWeekday
//  Replaces: eval("Date.parse(frozenDate).next().month().first().monday()")
//  Finds the Nth occurrence of the same weekday in the NEXT calendar month
// ============================================================
function getNextNthWeekday(fromDate, freq) {
  var targetDay = fromDate.getDay();  // 0=Sun ... 6=Sat
  var year = fromDate.getFullYear();
  var nextMonth = fromDate.getMonth() + 1;

  // Handle December → January rollover
  if (nextMonth > 11) { nextMonth = 0; year++; }

  var daysInMonth = new Date(year, nextMonth + 1, 0).getDate();
  var occurrences = [];

  for (var day = 1; day <= daysInMonth; day++) {
    var candidate = new Date(year, nextMonth, day);
    if (candidate.getDay() === targetDay) {
      occurrences.push(candidate);
    }
  }

  if (occurrences.length === 0) return fromDate; // safety fallback

  if (freq === "E1st") return occurrences[0];
  if (freq === "E2nd") return occurrences.length > 1 ? occurrences[1] : occurrences[0];
  if (freq === "E3rd") return occurrences.length > 2 ? occurrences[2] : occurrences[0];
  if (freq === "E4th") return occurrences.length > 3 ? occurrences[3] : occurrences[0];
  if (freq === "ELast") return occurrences[occurrences.length - 1];

  return occurrences[0];
}

//  fmtDate — format Date as yyyy-MM-dd (used for calendar lookup)
function fmtDate(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return "";
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

//  alertAndLog — show UI alert and log the message
function alertAndLog(msg) {
  Logger.log("ERROR: " + msg);
  safeAlert(msg);
}

function getEmail(theName, theDepartment) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Doer List");
  let data = sheet.getRange(1, 1, sheet.getLastRow(), 3).getValues();

  // Check Name match
  let nameMatch = data.find(r => r[0] === theName);

  if (!nameMatch) {
    return { status: "FAILED", reason: "Doer Name not found" };
  }

  //Check Department match
  let deptMatch = data.find(r => r[0] === theName && r[1] === theDepartment);

  if (!deptMatch) {
    return { status: "FAILED", reason: "Department mismatch" };
  }

  //Check Email
  if (!deptMatch[2]) {
    return { status: "FAILED", reason: "Email missing" };
  }

  return { status: "OK", email: deptMatch[2] };
}

function safeAlert(msg) {
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    Logger.log("ALERT: " + msg);
  }
}


function isAuthorizedUser() {
  var allowedUsers = [
    "mis@jjprintindia.com",   //add your allowed emails
    "admin@company.com"
  ];

  var currentUser = Session.getActiveUser().getEmail();
   Logger.log("Current User: " + currentUser);
  return allowedUsers.includes(currentUser);
}
// function submitTask(){

//   // USER RESTRICTION
//   if (!isAuthorizedUser()) {
//     safeAlert("You are not authorized to run this.");
//     return;
//   }
//   // RUN
//   generateTaskIDs();
//   createChecklistV1();
// }

function submitTask() {
  try {
    // USER CHECK
    if (!isAuthorizedUser()) {
      return {
        status: "error",
        message: "You are not authorized to run this"
      };
    }

    // RUN LOGIC
    generateTaskIDs();
    let count = createChecklistV1();  // must return value

    if (!count || count === 0) {
      return {
        status: "no_data",
        message: "No new tasks to process"
      };
    }

    return {
      status: "success",
      message: count + " tasks created successfully"
    };

  } catch (err) {
    return {
      status: "error",
      message: err.message
    };
  }
}

const SECRET_TOKEN = "#@Harsh#@Aman@#$JJ@#%";

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);

    // Token check
    if (body.token !== SECRET_TOKEN) {
      return ContentService
        .createTextOutput(JSON.stringify({
          status: "error",
          message: "Unauthorized"
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Action handler
    if (body.action === "submitTask") {
      submitTask();

      return ContentService
        .createTextOutput(JSON.stringify({
          status: "success",
          message: "Task executed successfully"
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Default response
    return ContentService
      .createTextOutput(JSON.stringify({
        status: "error",
        message: "Invalid action"
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: "error",
        message: err.message
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}