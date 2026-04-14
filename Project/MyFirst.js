function onEdit(e) {

  var sheet = e.range.getSheet();

  if (sheet.getName() !== "Task_List") return;

  var row = e.range.getRow();
  var col = e.range.getColumn();

  if (col === 12 && row > 1) {

    var sentStatusCell = sheet.getRange(row, 13); // Column M
    var currentSentStatus = sentStatusCell.getValue();

    // Only mark Updated if already Sent
    if (currentSentStatus === "Sent") {
      sentStatusCell.setValue("Updated");
    }
  }
}

// function copyExistingData() {

//   var ss = SpreadsheetApp.getActiveSpreadsheet();

//   var formSheet = ss.getSheetByName("Form Data");
//   var taskSheet = ss.getSheetByName("Task_List");

//   var data = formSheet.getDataRange().getValues();
//   if (data.length === 0) return;

//   // Remove Timestamp
//   var cleanedData = data.map(row => row.slice(1));

//   var header = cleanedData[0];
//   var rows = cleanedData.slice(1);

//   // Header check
//   var firstCell = taskSheet.getRange(1,1).getValue();

//   if (firstCell !== header[0]) {
//     taskSheet.getRange(1, 1, 1, header.length).setValues([header]);

//     taskSheet.getRange(1, header.length + 1).setValue("Status");
//     taskSheet.getRange(1, header.length + 2).setValue("Sent_Status");
//   }

//   if (rows.length > 0) {

//     var finalData = rows.map(function(row) {
//       row.push(""); 
//       row.push(""); 
//       return row;
//     });

//     var startRow = taskSheet.getLastRow() + 1;

//     taskSheet.getRange(startRow, 1, finalData.length, finalData[0].length)
//       .setValues(finalData);
//   }
// }

function onFormSubmit(e) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var taskSheet = ss.getSheetByName("Task_List");

  var headers = taskSheet.getRange(1, 1, 1, taskSheet.getLastColumn()).getValues()[0];

  var formData = e.namedValues;

  var newRow = [];

  headers.forEach(function(header) {

    if (header === "Sent_Status") {
      newRow.push(""); 
    } else if (formData[header]) {
      newRow.push(formData[header][0]);
    } else {
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

  var taskData = taskSheet.getDataRange().getValues();
  var doerData = doerSheet.getDataRange().getValues();

  //Email lookup
  function getEmail(name) {
    for (var i = 0; i < doerData.length; i++) {
      if (doerData[i][0] == name) return doerData[i][1];
    }
    return "";
  }

  var header = [
    "Doer","Email","Department","Task ID","Master Task ID",
    "Frequency","Task","How","Details","Time","Mobile",
    "Start Date","Status","Remark","Email_Copy","Subtask"
  ];

  if (masterSheet.getLastRow() === 0) {
    masterSheet.getRange(1, 1, 1, header.length).setValues([header]);
  }

  // Group rows by Task ID
  var groupedTasks = {};

  for (var i = 1; i < taskData.length; i++) {

    var row = taskData[i];

    var task       = row[0];
    var taskId     = row[1];
    var subtask    = row[2];
    var doer       = row[3];
    var dept       = row[4];
    var how        = row[5];
    var details    = row[6];
    var time       = row[7];
    var mobile     = row[8];
    var freq       = row[9];
    var startDate  = row[10];
    var status     = row[11] || "Pending";
    var sentStatus = row[12];

    //Skip empty or already sent
    if (!task || sentStatus === "Sent") continue;

    if (!groupedTasks[taskId]) {
      groupedTasks[taskId] = [];
    }

    groupedTasks[taskId].push({
      rowIndex: i + 1,
      doer, dept, how, details, time, mobile,
      freq, startDate, status, subtask, task
    });
  }

  var masterLastRow = masterSheet.getLastRow();
  var output = [];
  var rowsToUpdate = [];

  //Process grouped tasks
  for (var taskId in groupedTasks) {

    var group = groupedTasks[taskId];

    group.forEach(function(item) {

      var email = getEmail(item.doer);

      var masterTaskId = "M-" + (masterLastRow + output.length + 1);

      output.push([
        item.doer,
        email,
        item.dept,
        taskId,
        masterTaskId,
        item.freq,
        item.task,
        item.how,
        item.details,
        item.time,
        item.mobile,
        item.startDate,
        item.status,
        "",           // Remark
        email,
        item.subtask
      ]);

      rowsToUpdate.push(item.rowIndex);
    });
  }

  //Write to Master
  if (output.length > 0) {
    masterSheet
      .getRange(masterLastRow + 1, 1, output.length, output[0].length)
      .setValues(output);
  }

  // Update Sent_Status in Task_List
  rowsToUpdate.forEach(function(r) {
    taskSheet.getRange(r, 13).setValue("Sent"); 
  });

}

function updateTasks() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var taskSheet   = ss.getSheetByName("Task_List");
  var masterSheet = ss.getSheetByName("Master");

  var taskData   = taskSheet.getDataRange().getValues();
  var masterData = masterSheet.getDataRange().getValues();

  //Create map → TaskID + Subtask
  var masterMap = {};

  for (var j = 1; j < masterData.length; j++) {

    var taskId  = masterData[j][3];   
    var subtask = masterData[j][15];  

    if (!taskId || !subtask) continue;

    var key = taskId.toString().trim().toLowerCase() + "||" +
              subtask.toString().trim().toLowerCase();

    masterMap[key] = j + 1;
  }

  var updates = [];
  var remarks = [];
  var resetRows = [];

  for (var i = 1; i < taskData.length; i++) {

    var row = taskData[i];

    var taskId     = row[1];   // Column B
    var subtask    = row[2];   // Column C
    var newStatus  = row[11];  // Column L
    var sentStatus = row[12];  // Column M

    //Only process rows marked as Updated
    if (!sentStatus || sentStatus.toString().trim().toLowerCase() !== "updated") continue;

    if (!taskId || !subtask) continue;

    var key = taskId.toString().trim().toLowerCase() + "||" +
              subtask.toString().trim().toLowerCase();

    var masterRow = masterMap[key];

    if (!masterRow) {
      Logger.log("Not found: " + key);
      continue;
    }

    //Store updates
    updates.push({ row: masterRow, value: newStatus });

    //Remark logic
    if (newStatus && newStatus.toString().toLowerCase() === "complete") {
      remarks.push({ row: masterRow, value: "Complete" });
    } else {
      remarks.push({ row: masterRow, value: "Updated" });
    }

    //Mark row to reset Sent_Status
    resetRows.push(i + 1);
  }

  //Apply updates to Master
  updates.forEach(function(u) {
    masterSheet.getRange(u.row, 13).setValue(u.value); // Status (Col M)
  });

  remarks.forEach(function(r) {
    masterSheet.getRange(r.row, 14).setValue(r.value); // Remark (Col N)
  });

  //Reset Sent_Status → "Sent"
  resetRows.forEach(function(r) {
    taskSheet.getRange(r, 13).setValue("Sent"); 
  });
}

function syncCompletedTasks() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var masterSheet    = ss.getSheetByName("Master");
  var completedSheet = ss.getSheetByName("Completed");

  //Create sheet if not exists
  if (!completedSheet) {
    completedSheet = ss.insertSheet("Completed");
  }

  var masterData    = masterSheet.getDataRange().getValues();
  var completedData = completedSheet.getDataRange().getValues();

  //HEADER (copy from Master if empty)
  if (completedSheet.getLastRow() === 0) {
    var header = masterData[0];
    completedSheet.getRange(1, 1, 1, header.length).setValues([header]);
  }

  // Track existing TaskID + Subtask
  var existingKeys = new Set();

  for (var i = 1; i < completedData.length; i++) {

    var taskId  = completedData[i][3]; 
    var subtask = completedData[i][15];  

    if (!taskId || !subtask) continue;

    var key = taskId.toString().trim().toLowerCase() + "||" +
              subtask.toString().trim().toLowerCase();

    existingKeys.add(key);
  }

  var newRows = [];

  for (var i = 1; i < masterData.length; i++) {

    var row = masterData[i];

    var status  = row[12];  // Column M
    var taskId  = row[3];   // Column D
    var subtask = row[15];  // Column P

    if (!status || !taskId || !subtask) continue;

    //Only completed
    if (status.toString().trim().toLowerCase() === "complete") {

      var key = taskId.toString().trim().toLowerCase() + "||" +
                subtask.toString().trim().toLowerCase();

      //Avoid duplicates
      if (!existingKeys.has(key)) {
        newRows.push(row);
        existingKeys.add(key);
      }
    }
  }

  // Insert data
  if (newRows.length > 0) {
    completedSheet.getRange(
      completedSheet.getLastRow() + 1,
      1,
      newRows.length,
      newRows[0].length
    ).setValues(newRows);
  }

  Logger.log("Synced Completed Tasks: " + newRows.length);
}

function deleteSubtaskFromTaskList() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var taskSheet = ss.getSheetByName("Task_List");

  var ui = SpreadsheetApp.getUi();

  //Ask Task ID
  var taskResponse = ui.prompt(
    "Delete Subtask",
    "Enter Task ID (e.g., T-001):",
    ui.ButtonSet.OK_CANCEL
  );

  if (taskResponse.getSelectedButton() !== ui.Button.OK) return;

  var taskIdInput = taskResponse.getResponseText().toString().trim().toLowerCase();

  //Ask Subtask
  var subtaskResponse = ui.prompt(
    "Delete Subtask",
    "Enter Subtask (e.g., UI):",
    ui.ButtonSet.OK_CANCEL
  );

  if (subtaskResponse.getSelectedButton() !== ui.Button.OK) return;

  var subtaskInput = subtaskResponse.getResponseText().toString().trim().toLowerCase();

  if (!taskIdInput || !subtaskInput) {
    ui.alert("Invalid input");
    return;
  }

  var data = taskSheet.getDataRange().getValues();
  var rowsToDelete = [];

  //Find matching rows
  for (var i = 1; i < data.length; i++) {

    var taskId  = data[i][1]; // Column B
    var subtask = data[i][2]; // Column C

    if (!taskId || !subtask) continue;

    var tId = taskId.toString().trim().toLowerCase();
    var sub = subtask.toString().trim().toLowerCase();

    if (tId === taskIdInput && sub === subtaskInput) {
      rowsToDelete.push(i + 1);
    }
  }

  if (rowsToDelete.length === 0) {
    ui.alert("No matching subtask found!");
    return;
  }

  //Confirm delete
  var confirm = ui.alert(
    "Confirm Delete",
    "Delete " + rowsToDelete.length + " row(s) from Task_List?",
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  //Delete bottom → top
  for (var j = rowsToDelete.length - 1; j >= 0; j--) {
    taskSheet.deleteRow(rowsToDelete[j]);
  }

  ui.alert("Subtask deleted from Task_List!");
}
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Task Manager")
    .addItem("Delete Task", "deleteTaskById")
    .addItem("Delete Subtask (Master)", "deleteSubtask")
    .addItem("Delete Subtask (Task_List)", "deleteSubtaskFromTaskList")
    .addToUi();
}