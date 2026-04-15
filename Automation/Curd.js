
function onEdit(e) {

  var sheet = e.range.getSheet();

  // Only run for Master sheet
  if (sheet.getName() !== "Master") return;

  var row = e.range.getRow();
  var col = e.range.getColumn();

  // Status column = Column 12 (L)
  if (col !== 12 || row === 1) return;

  var newValue = e.value; // value from dropdown

  if (!newValue) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var taskSheet = ss.getSheetByName("Task_List");

  var masterRow = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  var taskId = masterRow[3]; // Task ID (Column D)

  if (!taskId) return;

  var taskData = taskSheet.getDataRange().getValues();

  for (var i = 1; i < taskData.length; i++) {

    var tId = taskData[i][1]; // Task_List Task ID

    if (tId == taskId) {

      // Mark Updated when ANY change happens
      taskSheet.getRange(i + 1, 11).setValue("Updated"); 
    }
  }
}

//new Form Submit with automatic taskid generation

function onFormSubmit(e) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var taskSheet = ss.getSheetByName("Task_List");

  var headers = taskSheet.getRange(1, 1, 1, taskSheet.getLastColumn()).getValues()[0];
  var formData = e.namedValues;

  //Get Start Date
  var startDateValue = formData["Start Date"] ? formData["Start Date"][0] : null;

  if (startDateValue) {

    var startDate = new Date(startDateValue);
    var today = new Date();

    // Normalize
    startDate.setHours(0,0,0,0);
    today.setHours(0,0,0,0);

    //Backdate check
    if (startDate < today) {

      SpreadsheetApp.getActive().toast(
        "Back date is not allowed!",
        "Invalid Entry",
        5
      );

      return;
    }
  }

  //Generate Task ID
  var taskId = "T-" + Math.floor(1000 + Math.random() * 9000);

  var newRow = [];

  headers.forEach(function(header) {

    if (header === "Task ID") {
      newRow.push(taskId);
    } 
    else if (header === "Sent_Status") {
      newRow.push("");
    } 
    else if (formData[header]) {
      newRow.push(formData[header][0]);
    } 
    else {
      newRow.push("");
    }

  });

  taskSheet.appendRow(newRow);
}


function submitTasks() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var taskSheet   = ss.getSheetByName("Task_List");
  var masterSheet = ss.getSheetByName("Master");
  var doerSheet   = ss.getSheetByName("Doer List");
  var calSheet    = ss.getSheetByName("Working Day Calender");

  var taskData = taskSheet.getDataRange().getValues();
  var doerData = doerSheet.getDataRange().getValues();

  //Working Days (FIXED)
  var calData = calSheet.getRange(2, 3, calSheet.getLastRow(), 1).getValues();
  var workingDays = calData
    .flat()
    .filter(d => d)
    .map(d => {
      var dt = new Date(d);
      dt.setHours(0,0,0,0);
      return dt;
    });

  //Email lookup
  function getEmail(name) {
    for (var i = 0; i < doerData.length; i++) {
      if (doerData[i][0] == name) return doerData[i][1];
    }
    return "";
  }

  //Header (NO Master Task ID)
  var header = [
    "Doer","Email","Department","Task ID",
    "Frequency","Task","How","Details","Time","Mobile",
    "Start Date","Status","Subtask ID"
  ];

  if (masterSheet.getLastRow() === 0) {
    masterSheet.getRange(1, 1, 1, header.length).setValues([header]);
  }

  var masterData = masterSheet.getDataRange().getValues();

  //Find last Subtask ID
  var lastSubId = 0;

  for (var i = 1; i < masterData.length; i++) {
    var subId = masterData[i][12]; // Subtask ID column

    if (subId && subId.toString().startsWith("S-")) {
      var num = parseInt(subId.toString().replace("S-", ""));
      if (num > lastSubId) lastSubId = num;
    }
  }

  var output = [];
  var rowsToUpdate = [];

  for (var i = 1; i < taskData.length; i++) {

    var row = taskData[i];

    var task       = row[0];
    var taskId     = row[1];
    var doer       = row[2];
    var dept       = row[3];
    var how        = row[4];
    var details    = row[5];
    var time       = row[6];
    var mobile     = row[7];
    var freq       = row[8];
    var startDate  = new Date(row[9]);
    var status     = "Pending";
    var sentStatus = row[11];

    startDate.setHours(0,0,0,0);

    if (!task || sentStatus === "Sent") continue;

    var email = getEmail(doer);

    //Working days loop
    workingDays.forEach(function(currentDate) {

      if (currentDate < startDate) return;

      var include = false;
      var freqVal = freq ? freq.toString().trim() : "";

      if (freqVal === "D(Daily)") {
        include = true;
      }
      else if (freqVal === "W(Weekly)") {
        include = currentDate.getDay() === startDate.getDay();
      }
      else if (freqVal === "M(Monthly)") {
        include = currentDate.getDate() === startDate.getDate();
      }

      if (!include) return;

      //Increment Subtask ID
      lastSubId++;
      var subtaskId = "S-" + lastSubId;

      output.push([
        doer,
        email,
        dept,
        taskId,
        freq,
        task,
        how,
        details,
        time,
        mobile,
        new Date(currentDate),
        status,
        subtaskId
      ]);

    });

    rowsToUpdate.push(i + 1);
  }

  //Write to Master
  if (output.length > 0) {
    masterSheet
      .getRange(masterSheet.getLastRow() + 1, 1, output.length, output[0].length)
      .setValues(output);
  }

  // Mark Sent
  rowsToUpdate.forEach(function(r) {
    taskSheet.getRange(r, 11).setValue("Sent");
  });

}

function updateTasks() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var taskSheet   = ss.getSheetByName("Task_List");
  var masterSheet = ss.getSheetByName("Master");

  var taskData   = taskSheet.getDataRange().getValues();
  var masterData = masterSheet.getDataRange().getValues();

  var resetRows = [];

  for (var i = 1; i < taskData.length; i++) {

    var taskId     = taskData[i][1];   // Column B
    var sentStatus = taskData[i][10];  // Column K

    if (!taskId || sentStatus !== "Updated") continue;

    // Loop Master and update all matching Task IDs
    for (var j = 1; j < masterData.length; j++) {

      if (masterData[j][3] == taskId) {

        var status = masterData[j][12];

        if (status && status.toString().toLowerCase() === "complete") {
          masterSheet.getRange(j + 1, 14).setValue("Complete");
        } else {
          masterSheet.getRange(j + 1, 14).setValue("Updated");
        }
      }
    }

    resetRows.push(i + 1);
  }

  // Reset Sent_Status
  resetRows.forEach(function(r) {
    taskSheet.getRange(r, 11).setValue("Sent");
  });
}

function syncCompletedTasks() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var masterSheet    = ss.getSheetByName("Master");
  var completedSheet = ss.getSheetByName("Completed");

  if (!completedSheet) {
    completedSheet = ss.insertSheet("Completed");
  }

  var masterData = masterSheet.getDataRange().getValues();

  // Add header if empty
  if (completedSheet.getLastRow() === 0) {
    completedSheet.getRange(1, 1, 1, masterData[0].length)
      .setValues([masterData[0]]);
  }

  var rowsToMove = [];
  var rowsIndex  = [];

  // Find completed rows
  for (var i = 1; i < masterData.length; i++) {

    var status = masterData[i][11]; // Status column

    if (status && status.toString().toLowerCase() === "complete") {
      rowsToMove.push(masterData[i]);
      rowsIndex.push(i + 1); // store actual row number
    }
  }

  // Move to Completed
  if (rowsToMove.length > 0) {
    completedSheet
      .getRange(completedSheet.getLastRow() + 1, 1, rowsToMove.length, rowsToMove[0].length)
      .setValues(rowsToMove);
  }

  // DELETE from Master (bottom to top)
  for (var j = rowsIndex.length - 1; j >= 0; j--) {
    masterSheet.deleteRow(rowsIndex[j]);
  }

  Logger.log("Moved " + rowsToMove.length + " completed tasks");
}

