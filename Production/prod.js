// PRODUCTION CHECKLIST — OPTIMIZED
// CONFIG
var CONFIG = {
  SKIP_SUNDAYS: true,
  TIME_LIMIT_MS: 5 * 60 * 1000
};
function generateTaskIDs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Task_List");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var data = sheet.getRange(2, 1, lastRow - 1, 10).getValues(); // cols A–J
  var idArr = [];
  var lastID = 0;

  data.forEach(function (row) {
    var existing = row[9]; // col J
    if (existing && existing.toString().indexOf("TL") === 0) {
      var num = parseInt(existing.toString().replace("TL", ""), 10);
      if (!isNaN(num) && num > lastID) lastID = num;
    }
  });

  for (var i = 0; i < data.length; i++) {
    var task = data[i][0];
    var existingID = data[i][9];
    if (!task) { idArr.push([""]); continue; }
    if (existingID) { idArr.push([existingID]); continue; }
    lastID++;
    idArr.push(["TL" + lastID]);
  }

  sheet.getRange(2, 10, idArr.length, 1).setValues(idArr);
}
function createChecklistV1() {
  Logger.log("=== createChecklistV1 START ===");
  var startTime = Date.now();

  var props = PropertiesService.getScriptProperties();
  var startIndex = parseInt(props.getProperty("CHECKLIST_INDEX") || "0", 10);

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  //Load sheets 
  var taskListSheet = ss.getSheetByName("Task_List");
  var masterSheet = ss.getSheetByName("Master_New");
  var calendarSheet = ss.getSheetByName("Working Day Calender");
  var doerSheet = ss.getSheetByName("Doer List");  // ← load once here

  if (!taskListSheet) { alertAndLog("Sheet not found: Task_List"); return; }
  if (!masterSheet) { alertAndLog("Sheet not found: Master_New"); return; }
  if (!calendarSheet) { alertAndLog("Sheet not found: Working Day Calender"); return; }
  if (!doerSheet) { alertAndLog("Sheet not found: Doer List"); return; }

  //Load Task_List — 1 read
  var taskLastRow = taskListSheet.getLastRow();
  if (taskLastRow < 2) { Logger.log("No data in Task_List. Aborting."); return; }

  var taskData = taskListSheet.getRange(2, 1, taskLastRow - 1, 11).getValues();
  Logger.log("TaskData length: " + taskData.length);
  Logger.log("First row: " + JSON.stringify(taskData[0]));

  // Load Working Day Calendar — 1 read
  var calLastRow = calendarSheet.getLastRow();
  var calColValues = calendarSheet.getRange(1, 1, calLastRow, 1).getValues();
  var calNonEmpty = calColValues.filter(function (r) { return r[0] !== "" && r[0] instanceof Date; });

  if (calNonEmpty.length < 2) {
    alertAndLog("Working Day Calender sheet is empty or has no valid dates!");
    return;
  }

  // FIX: Build O(1) object instead of array + indexOf() per lookup
  var workingSet = {};
  calNonEmpty.slice(1).forEach(function (r) { workingSet[fmtDate(r[0])] = true; });
  var calendarLastDate = calNonEmpty[calNonEmpty.length - 1][0];
  Logger.log("Calendar loaded. Working days: " + Object.keys(workingSet).length + " | Last: " + fmtDate(calendarLastDate));

  //Build Doer Map — 1 read, replaces per-row sheet reads 
  //reading Doer List sheet inside getEmail() on every task row
  var doerRaw = doerSheet.getRange(1, 1, doerSheet.getLastRow(), 3).getValues();
  var doerMap = {};  // key: "Name||Dept" → email
  doerRaw.forEach(function (r) {
    var name = String(r[0] || "").trim();
    var dept = String(r[1] || "").trim();
    var email = String(r[2] || "").trim();
    if (name) doerMap[name + "||" + dept] = email || "__MISSING__";
  });
  Logger.log("Doer map built. Entries: " + Object.keys(doerMap).length);

  //Get last SubTaskID from Master — 1 read
  var masterLastRow = masterSheet.getLastRow();
  var lastMasterTaskID = 0;

  if (masterLastRow >= 2) {
    masterSheet.getRange(2, 5, masterLastRow - 1, 1).getValues().forEach(function (r) {
      var n = parseInt(r[0], 10);
      if (!isNaN(n) && n > lastMasterTaskID) lastMasterTaskID = n;
    });
  }
  Logger.log("Resuming SubTaskID from: " + lastMasterTaskID);
  //Process task rows
  var masterArray = [];
  var statusMap = {};
  var timedOut = false;
  var today = new Date();
  var i;

  //Removed BATCH_SIZE=35 cap — time guard handles stopping gracefully
  for (i = startIndex; i < taskData.length; i++) {
    Logger.log("---- Row " + (i + 2) + " ----");

    if (Date.now() - startTime > CONFIG.TIME_LIMIT_MS) {
      Logger.log("WARNING: Time limit at task row " + (i + 2) + ". Will resume next run.");
      timedOut = true;
      break;
    }

    var row = taskData[i];

    // Column mapping (your original layout preserved exactly)
    var theTask = String(row[0] || "").trim();   // A - What
    var theDoer = String(row[1] || "").trim();   // B - Doer Name
    var theDept = String(row[2] || "").trim();   // C - Department
    var theHow = String(row[3] || "").trim();   // D - How
    var theDetails = String(row[4] || "").trim();   // E - Details
    var theTime = row[5];                        // F - Time
    var theMobile = row[6];                        // G - Mobile
    var theFreq = String(row[7] || "").trim();   // H - Frequency
    var theDate = row[8];                        // I - Date
    var theTaskID = row[9];                        // J - TasklistID
    var theStatus = String(row[10] || "").trim();  // K - Status

    Logger.log("Task: " + theTask + " | Doer: " + theDoer + " | Freq: " + theFreq + " | Date: " + theDate);

    if (theStatus === "Sent") { Logger.log("Already Sent — skip"); continue; }
    if (!theTask && !theDoer) continue;

    // if (!(theDate instanceof Date) || isNaN(theDate.getTime())) {
    //   Logger.log("SKIP row " + (i + 2) + ": invalid date");
    //   statusMap[i] = "Skipped - Invalid Date";
    //   continue;
    // }

    if (!(theDate instanceof Date) || isNaN(theDate.getTime())) {
      Logger.log("SKIP row " + (i + 2) + ": invalid date");
      statusMap[i] = "Skipped - Invalid Date";
      continue;
    }

    //NEW VALIDATION: Skip past dates
    var checkDate = new Date(theDate);
    checkDate.setHours(0, 0, 0, 0);

    var todayCheck = new Date();
    todayCheck.setHours(0, 0, 0, 0);

    if (checkDate < todayCheck) {
      Logger.log("SKIP row " + (i + 2) + ": past date");
      statusMap[i] = "Skipped - Past Date";
      continue;
    }

    //in-memory lookup — zero sheet reads per row
    var emailResult = getEmailFromMap(doerMap, theDoer, theDept);
    if (emailResult.status === "FAILED") {
      Logger.log("SKIP row " + (i + 2) + ": " + emailResult.reason);
      statusMap[i] = "Skipped - " + emailResult.reason;
      continue;
    }
    var theEmail = emailResult.email;

    var result = buildOccurrences(
      theFreq, new Date(theDate.getTime()), calendarLastDate,
      workingSet, lastMasterTaskID, today,
      theTask, theDoer, theEmail, theDept, theTaskID,
      theHow, theDetails, theTime, theMobile
    );

    if (result.rows.length > 0) {
      lastMasterTaskID = result.lastMasterTaskID;
      masterArray = masterArray.concat(result.rows);
    }

    statusMap[i] = "Sent";
  }

  Logger.log("FINAL masterArray length: " + masterArray.length);

  //Batch write to Master
  if (masterArray.length > 0) {
    var writeAt = masterLastRow + 1;
    if (writeAt < 2) writeAt = 2;
    var CHUNK = 5000;
    for (var j = 0; j < masterArray.length; j += CHUNK) {
      var chunkData = masterArray.slice(j, j + CHUNK);
      masterSheet.getRange(writeAt + j, 1, chunkData.length, 14).setValues(chunkData);
    }
    Logger.log("Written " + masterArray.length + " rows to Master from row " + writeAt);
  } else {
    Logger.log("No new rows to write to Master.");
  }

  /* Write status — only changed rows rewriting ALL status rows on every batch run
  Now only writes rows that were actually processed this run*/

  var statusKeys = Object.keys(statusMap);
  if (statusKeys.length > 0) {
    statusKeys.sort(function (a, b) { return parseInt(a) - parseInt(b); });
    statusKeys.forEach(function (idxStr) {
      var idx = parseInt(idxStr, 10);
      taskListSheet.getRange(idx + 2, 11).setValue(statusMap[idxStr]);
    });
  }

  //Single final alert — no duplicates 
  if (timedOut) {
    props.setProperty("CHECKLIST_INDEX", String(i));
    safeAlert(
      "⏱ Progress saved!\n\n" +
      "Processed up to row " + (i + 1) + " of " + (taskData.length + 1) + ".\n" +
      "Click Submit again to continue from where it stopped."
    );
  } else if (i < taskData.length) {
    props.setProperty("CHECKLIST_INDEX", String(i));
    safeAlert(
      "Partial run complete.\n\n" +
      "Processed up to row " + (i + 1) + ".\n" +
      "Run again to continue."
    );
  } else {
    props.deleteProperty("CHECKLIST_INDEX"); // reset resume pointer on full completion
    Logger.log("=== createChecklistV1 COMPLETE. Total rows added: " + masterArray.length + " ===");
    safeAlert("Checklist fully created!\n\n" + masterArray.length + " rows written to Master_New.");
  }

  return masterArray.length;
}
function getEmailFromMap(doerMap, theName, theDepartment) {
  if (!theName) return { status: "FAILED", reason: "Doer name is blank" };

  // Try exact Name + Dept key
  var key = theName + "||" + theDepartment;
  var email = doerMap[key];

  // Fallback: name only (dept blank in Doer List)
  if (email === undefined) {
    email = doerMap[theName + "||"];
  }

  if (email === undefined) return { status: "FAILED", reason: "Doer Name not found: " + theName };
  if (email === "__MISSING__") return { status: "FAILED", reason: "Email missing for: " + theName };

  return { status: "OK", email: email };
}
//  buildOccurrences — same logic, workingSet replaces indexOf()
function buildOccurrences(
  freq, startDate, calendarLastDate, workingSet,
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
      new Date(d), "", email
    ]);
  }

  function addDays(d, n) {
    var r = new Date(d.getTime()); r.setDate(r.getDate() + n); return r;
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
    var r = new Date(d.getTime()); r.setFullYear(r.getFullYear() + n); return r;
  }
  function cloneDate(d) { return new Date(d.getTime()); }
  function isBefore(a, b) {
    return new Date(a.getFullYear(), a.getMonth(), a.getDate()).getTime() <
      new Date(b.getFullYear(), b.getMonth(), b.getDate()).getTime();
  }

  // FIX: O(1) object lookup instead of array.indexOf() called thousands of times
  function isWorkingDay(d) { return !!workingSet[fmtDate(d)]; }
  function isSunday(d) { return d.getDay() === 0; }
  function isValid(d) { return isWorkingDay(d) && !(CONFIG.SKIP_SUNDAYS && isSunday(d)); }

  function shiftBack(d) {
    var cur = cloneDate(d);
    var safe = 0;
    while (!isValid(cur)) {
      cur = addDays(cur, -1);
      if (++safe > 90) { Logger.log("WARN: shiftBack > 90 days"); return cloneDate(d); }
    }
    return cur;
  }

  var d = cloneDate(startDate);

  if (freq === "D") {
    while (isBefore(d, calendarLastDate)) {
      if (isValid(d)) pushRow(d);
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
      if (isValid(d)) pushRow(d);
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
    if (isValid(d)) pushRow(d);
  } else if (freq === "E1st" || freq === "E2nd" || freq === "E3rd" ||
    freq === "E4th" || freq === "ELast") {
    while (isBefore(d, calendarLastDate)) {
      pushRow(shiftBack(d));
      d = getNextNthWeekday(cloneDate(d), freq);
    }
  } else {
    Logger.log("WARN: Unknown frequency '" + freq + "' for task '" + taskName + "' — skipped.");
  }

  return { rows: rows, lastMasterTaskID: counter };
}
//  getNextNthWeekday
function getNextNthWeekday(fromDate, freq) {
  var targetDay = fromDate.getDay();
  var year = fromDate.getFullYear();
  var nextMonth = fromDate.getMonth() + 1;
  if (nextMonth > 11) { nextMonth = 0; year++; }

  var daysInMonth = new Date(year, nextMonth + 1, 0).getDate();
  var occurrences = [];

  for (var day = 1; day <= daysInMonth; day++) {
    var candidate = new Date(year, nextMonth, day);
    if (candidate.getDay() === targetDay) occurrences.push(candidate);
  }

  if (!occurrences.length) return fromDate;
  if (freq === "E1st") return occurrences[0];
  if (freq === "E2nd") return occurrences.length > 1 ? occurrences[1] : occurrences[0];
  if (freq === "E3rd") return occurrences.length > 2 ? occurrences[2] : occurrences[0];
  if (freq === "E4th") return occurrences.length > 3 ? occurrences[3] : occurrences[0];
  if (freq === "ELast") return occurrences[occurrences.length - 1];
  return occurrences[0];
}
// fmtDate 
function fmtDate(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return "";
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}
//  alertAndLog
function alertAndLog(msg) {
  Logger.log("ERROR: " + msg);
  safeAlert(msg);
}
//  safeAlert 
function safeAlert(msg) {
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    Logger.log("ALERT: " + msg);
  }
}
//  isAuthorizedUser
function isAuthorizedUser() {
  var allowedUsers = [
    "mis@jjprintindia.com",
    "admin@company.com"
  ];
  var currentUser = Session.getActiveUser().getEmail();
  Logger.log("Current User: " + currentUser);
  return allowedUsers.indexOf(currentUser) !== -1;
}
//  submitTask
function submitTask() {
  try {
    if (!isAuthorizedUser()) {
      safeAlert("You are not authorized to run this script.");
      return { status: "error", message: "You are not authorized to run this" };
    }

    generateTaskIDs();
    var count = createChecklistV1();

    if (!count || count === 0) {
      return { status: "no_data", message: "No new tasks to process" };
    }

    return { status: "success", message: count + " tasks created successfully" };

  } catch (err) {
    Logger.log("submitTask ERROR: " + err.message);
    return { status: "error", message: err.message };
  }
}
//  resetChecklist — clears resume pointer (run manually if needed)
function resetChecklist() {
  PropertiesService.getScriptProperties().deleteProperty("CHECKLIST_INDEX");
  safeAlert("Resume pointer cleared. Next run will start from row 1.");
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
      var result = submitTask();

      return ContentService
        .createTextOutput(JSON.stringify(result)) 
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

function deleteFutureTasks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("Master_New");
  var deletedSheet = ss.getSheetByName("Deleted_Records");

  if (!deletedSheet) {
    deletedSheet = ss.insertSheet("Deleted_Records");
  }

  var ui = SpreadsheetApp.getUi();

  // Ask Task ID
  var taskPrompt = ui.prompt("Enter Task ID (e.g. TL252)");
  if (taskPrompt.getSelectedButton() !== ui.Button.OK) return;
  var taskID = taskPrompt.getResponseText().trim();

  // Ask SubTaskID (optional)
  var subPrompt = ui.prompt("Enter SubTaskID (optional, press OK to skip)");
  if (subPrompt.getSelectedButton() !== ui.Button.OK) return;
  var subTaskInput = subPrompt.getResponseText().trim();
  var subTaskID = subTaskInput ? parseInt(subTaskInput, 10) : null;

  var data = masterSheet.getDataRange().getValues();
  var header = data[0];
  var rows = data.slice(1);

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var remaining = [];
  var deleted = [];

  rows.forEach(function (row) {
    var rowTaskID = row[3];      // Column D = TaskID
    var rowSubID = row[4];      // Column E = SubTaskID
    var planned = row[11];     // Column L = Planned Date

    var isFuture = planned instanceof Date && planned >= today;
    var taskMatch = rowTaskID === taskID;
    var subMatch = subTaskID ? rowSubID >= subTaskID : true;

    if (taskMatch && isFuture && subMatch) {
      deleted.push(row);
    } else {
      remaining.push(row);
    }
  });

  //Write remaining back
  masterSheet.clearContents();
  masterSheet.getRange(1, 1, 1, header.length).setValues([header]);

  if (remaining.length > 0) {
    masterSheet.getRange(2, 1, remaining.length, header.length).setValues(remaining);
  }

  //Store deleted records
  if (deleted.length > 0) {
    var lastRow = deletedSheet.getLastRow();
    if (lastRow === 0) {
      deletedSheet.appendRow(header.concat(["Deleted On"]));
      lastRow = 1;
    }

    var deletedWithStamp = deleted.map(function (r) {
      return r.concat([new Date()]);
    });

    deletedSheet.getRange(lastRow + 1, 1, deletedWithStamp.length, header.length + 1)
      .setValues(deletedWithStamp);
  }

  ui.alert(deleted.length + " future tasks deleted successfully.");
}

// updation start 
function onEdit(e) {
  var sheet = e.source.getActiveSheet();

  if (sheet.getName() !== "Task_List") return;

  var row = e.range.getRow();
  var col = e.range.getColumn();

  //Ignore header
  if (row === 1) return;

  //Ignore multiple cells (paste / row insert / bulk edit)
  if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;

  //Ignore if no old value (means NEW row / first time entry)
  if (typeof e.oldValue === "undefined") return;

  //Only specific columns
  if (![2, 3, 7].includes(col)) return;

  //Now only REAL edit on existing row

  // Update Last Updated
  sheet.getRange(row, 13).setValue(new Date());

  // Highlight
  sheet.getRange(row, 1, 1, sheet.getLastColumn())
    .setBackground("#fff3cd");

  // Update Master
  updateMasterFromTaskList(row);
}

function updateMasterFromTaskList(rowIndex) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var taskSheet = ss.getSheetByName("Task_List");
  var masterSheet = ss.getSheetByName("Master_New");
  var doerSheet = ss.getSheetByName("Doer List");

  var row = taskSheet.getRange(rowIndex, 1, 1, 13).getValues()[0];

  var taskID = row[9]; // J
  var doer = row[1]; // B
  var dept = row[2]; // C
  var mobile = row[6]; // G

  if (!taskID) return;

  //Build Doer Map (FAST)
  var doerData = doerSheet.getDataRange().getValues();
  var email = "";

  for (var i = 1; i < doerData.length; i++) {
    if (doerData[i][0] === doer) {
      email = doerData[i][2];
      break;
    }
  }

  if (!email) {
    Logger.log("Email not found for: " + doer);
    return; // safer: don’t overwrite with blank
  }

  var data = masterSheet.getDataRange().getValues();

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var updated = 0;

  for (var i = 1; i < data.length; i++) {
    var rowTaskID = data[i][3];   // D
    var planned = data[i][11];  // L

    if (
      rowTaskID === taskID &&
      planned instanceof Date &&
      planned >= today
    ) {
      // Only update if changed (avoids unnecessary writes)
      if (
        data[i][0] !== doer ||
        data[i][1] !== email ||
        data[i][2] !== dept ||
        data[i][10] !== mobile
      ) {
        data[i][0] = doer;
        data[i][1] = email;
        data[i][2] = dept;
        data[i][10] = mobile;
        updated++;
      }
    }
  }

  if (updated > 0) {
    masterSheet
      .getRange(2, 1, data.length - 1, data[0].length)
      .setValues(data.slice(1));
  }

  Logger.log("Updated rows: " + updated);
}


function markExpiryColumn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var taskSheet = ss.getSheetByName("Task_List");
  var masterSheet = ss.getSheetByName("Master_New");

  var taskLastRow = taskSheet.getLastRow();
  var masterLastRow = masterSheet.getLastRow();

  if (taskLastRow < 2 || masterLastRow < 2) return;

  var taskData = taskSheet.getRange(2, 1, taskLastRow - 1, 14).getValues();
  var masterData = masterSheet.getRange(2, 1, masterLastRow - 1, 12).getValues();

  //Build map → TaskID → MAX Planned Date
  var expiryMap = {};

  masterData.forEach(function (r) {
    var taskID = r[3];   // D
    var date = r[11];  // L

    if (taskID && date instanceof Date) {
      if (!expiryMap[taskID] || expiryMap[taskID] < date) {
        expiryMap[taskID] = date;
      }
    }
  });

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var updates = [];

  for (var i = 0; i < taskData.length; i++) {

    var taskID = taskData[i][9];   // J
    var currentVal = taskData[i][13];  // N (Expiry column)

    var expiryDate = expiryMap[taskID];

    // Only ADD "Expired" if condition met
    if ((!expiryDate || expiryDate < today) && currentVal !== "Expired") {
      updates.push({ row: i + 2, value: "Expired" });
    }
  }

  //Update ONLY required rows (no overwrite)
  updates.forEach(function (u) {
    taskSheet.getRange(u.row, 14).setValue(u.value);
  });

  Logger.log("Expired rows updated: " + updates.length);
}