function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Ultimate Checklist')
    .addItem('Create Checklist', 'createChecklist')
    .addItem('Setup Sheet', 'setTrigger')
    .addItem('Archive', 'archive')
    .addItem('Send to Doer', 'sendtodoer')
    .addToUi();
}


function getLastRowSpecial(columnToCheck) {
  for (var i = 0; i < columnToCheck.length; i++) {
    if (columnToCheck[i][0] === "") {
      return i;
    }
  }
  return columnToCheck.length;
}

if (typeof SheetConverter === "undefined") {
  var SheetConverter = {
    convertRange2html: function(range) {
      var values = range.getDisplayValues();
      var html = "<table border='1' cellpadding='5' style='border-collapse: collapse;'>";
      for (var i = 0; i < values.length; i++) {
        html += "<tr>" + values[i].map(function(c){ return "<td>" + c + "</td>" }).join("") + "</tr>";
      }
      html += "</table>";
      return html;
    }
  };
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
  var sheet = ss.getSheetByName("Dashboard")
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
  var spreadsheet = ss.getSheetByName("Dashboard")
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