Date.prototype.addTheDays = function (days) {
  var date = new Date(this.valueOf());
  date.setDate(date.getDate() + days);
  return date;
}

Number.prototype.addDays = function (days) { var d = new Date(this.valueOf()); d.setDate(d.getDate() + days); return d; }
Number.prototype.addWeeks = function (weeks) { var d = new Date(this.valueOf()); d.setDate(d.getDate() + weeks * 7); return d; }
Number.prototype.addMonths = function (months) { var d = new Date(this.valueOf()); d.setMonth(d.getMonth() + months); return d; }
Number.prototype.addYears = function (years) { var d = new Date(this.valueOf()); d.setFullYear(d.getFullYear() + years); return d; }
Number.prototype.next = function () { return new Date(this.valueOf()).next(); }

Date.prototype.clone = function () { return new Date(this.valueOf()); }
Date.prototype.is = function () {
  return {
    jan: () => this.getMonth() === 0,
    sunday: () => this.getDay() === 0
  };
};

Date.compare = function (d1, d2) {
  return new Date(d1).valueOf() > new Date(d2).valueOf() ? 1 :
    (new Date(d1).valueOf() < new Date(d2).valueOf() ? -1 : 0);
};

Date.prototype.next = function () {
  let ctx = this;
  let createDayMethods = function (nth, dateObj) {
    let days = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
    let obj = {};
    days.forEach((dayStr, index) => {
      obj[dayStr] = function () {
        let d = new Date(dateObj.getFullYear(), dateObj.getMonth() + 1, 1);
        if (nth === 'last') {
          d.setMonth(d.getMonth() + 1, 0);
          let diff = (d.getDay() - index + 7) % 7;
          d.setDate(d.getDate() - diff);
          return d;
        }
        let diff = (index - d.getDay() + 7) % 7;
        d.setDate(1 + diff + (nth - 1) * 7);
        return d;
      }
    });
    return obj;
  };

  return {
    month: function () {
      return {
        first: function () { return createDayMethods(1, ctx); },
        second: function () { return createDayMethods(2, ctx); },
        third: function () { return createDayMethods(3, ctx); },
        fourth: function () { return createDayMethods(4, ctx); },
        final: function () { return createDayMethods('last', ctx); }
      }
    }
  }
};

function getDayInfo(dayIdx) {
  return ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'][dayIdx];
}

function findinB(name, dataB) {
  for (var i = 0; i < dataB.length; i++) {
    if (dataB[i][0] === name) return dataB[i][1];
  }
  return "";
}

function getEmail(theDoer) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById("19nTuvRPpMvnGNgbKojibhCd0z31j4Euj26a5eKy7xyQ");
    var forBsh = ss.getSheetByName("Doer List");
    if (!forBsh) return "";
    var lastB = forBsh.getLastRow();
    if (lastB < 1) return "";
    var dataB = forBsh.getRange(1, 1, lastB, 2).getValues();
    return findinB(theDoer, dataB);
  } catch (e) { }
  return "";
}

function createChecklist() {
  let ss = SpreadsheetApp.openById("19nTuvRPpMvnGNgbKojibhCd0z31j4Euj26a5eKy7xyQ");
  let taskListSheet = ss.getSheetByName("Task_List");
  let maxTaskRows = taskListSheet.getLastRow() - 1;
  if (maxTaskRows < 1) return;
  let taskListData = taskListSheet.getRange(2, 1, maxTaskRows, 11).getValues()
  let skipSundays = ss.getSheetByName("Setup Sheet").getRange("B32").getValue()
  let todaysdate = new Date()

  if (!skipSundays) {
    skipSundays = "Yes"
  }

  let masterSheet = ss.getSheetByName("Master")
  let cell = masterSheet.getRange("M2") //actual Cell
  let masterTaskIDs = masterSheet.getRange("e:e").getValues()
  let masterLastRow = masterTaskIDs.filter(r => String(r)).length

  let lastMasterTaskID = 0;
  if (masterLastRow > 0) {
    lastMasterTaskID = Number(masterTaskIDs[masterLastRow - 1]);
    if (isNaN(lastMasterTaskID) || !lastMasterTaskID) {
      lastMasterTaskID = 0;
    }
  }

  let masterArray = []

  let calendarSheet = ss.getSheetByName("Working Day Calender")
  let calendarDates = calendarSheet.getRange("A:A").getValues()
  let calendarLast = calendarDates.filter(String).length
  let allCalendarDates = calendarSheet.getRange(2, 1, calendarLast - 1, 1).getValues()
  let calendarLastDate = calendarDates[calendarLast - 1][0]
  let workingDates = allCalendarDates.flat()
  let workingDatesStr = workingDates.map(x => Utilities.formatDate(x, Session.getScriptTimeZone(), "yyyy-MM-dd"));

  taskListData.forEach(function (task, index) {
    let theTask = task[0]
    let theDoer = task[1]
    let theDepartment = task[2]
    let theHow = task[3]
    let theDetails = task[4]
    let theTime = task[5]
    let theMobile = task[6]
    let theFreq = task[7]
    let theDate = task[8]
    let theTasklistID = task[9]
    let theStatus = task[10]
    let theEmail = getEmail(theDoer, theDepartment, theHow, theDetails, theTime, theMobile, theFreq, theDate, theTasklistID)

    if (theEmail === "FAILED") {
      taskListSheet.getRange(index + 2, 11).setValue("Skipped Due to doer - department - how - details - time - mobile - freq - date - tasklistID mismatch")
      return
    }

    if (theStatus != "Sent") {
      //weekly tasks 
      if (theFreq === "W") {
        while (Date.compare(calendarLastDate, theDate) === 1) {
          lastMasterTaskID = lastMasterTaskID + 1
          let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          let frozenDate = theDate.clone()

          while (!workingDatesStr.includes(endDateStr)) {
            theDate = Date.parse(theDate).addDays(-1)
            endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }

          masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate), "", "", theEmail])

          theDate = Date.parse(frozenDate).addWeeks(1)
        }
      } else if (theFreq === "M") {
        while (Date.compare(calendarLastDate, theDate) === 1) {
          let norun = 0
          if (theDate.is().jan() && theDate.getDate() > 28) {
            norun = 1
            lastMasterTaskID = lastMasterTaskID + 1
            let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

            let frozenDate = theDate.clone()

            while (!workingDatesStr.includes(endDateStr)) {
              theDate = Date.parse(theDate).addDays(-1)
              endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
            }

            if (skipSundays === "Yes") {
              if (theDate.is().sunday()) {
                theDate = Date.parse(theDate).addDays(-1)
              }
            }

            masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate), "", "", theEmail])

            theDate = Date.parse(frozenDate).addMonths(2)
          }

          if (norun === 0) {
            lastMasterTaskID = lastMasterTaskID + 1
            let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

            let frozenDate = theDate.clone()

            while (!workingDatesStr.includes(endDateStr)) {
              theDate = Date.parse(theDate).addDays(-1)
              endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
            }

            if (skipSundays === "Yes") {
              if (theDate.is().sunday()) {
                theDate = Date.parse(theDate).addDays(-1)
              }
            }

            masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate), "", "", theEmail])

            theDate = Date.parse(frozenDate).addMonths(1)
          }
        }
      } else if (theFreq === "Y") {
        if (theDate.valueOf() > todaysdate.valueOf()) {
          lastMasterTaskID = lastMasterTaskID + 1
          masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate), "", "", theEmail])
        }
        lastMasterTaskID = lastMasterTaskID + 1
        theDate = Date.parse(theDate).addYears(1)
        masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate), "", "", theEmail])
        lastMasterTaskID = lastMasterTaskID + 1
        theDate = Date.parse(theDate).addYears(1)
        masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate), "", "", theEmail])
        lastMasterTaskID = lastMasterTaskID + 1
        theDate = Date.parse(theDate).addYears(1)
        masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate), "", "", theEmail])

      } else if (theFreq === "D") {
        while (Date.compare(calendarLastDate, theDate) === 1) {
          lastMasterTaskID = lastMasterTaskID + 1
          let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          let frozenDate = theDate.clone()

          if (skipSundays === "Yes") {
            if (workingDatesStr.includes(endDateStr) && !theDate.is().sunday()) {
              masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate), "", "", theEmail])
            }
          } else {
            if (workingDatesStr.includes(endDateStr)) {
              masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate), "", "", theEmail])
            }
          }

          theDate = Date.parse(frozenDate).addDays(1)
        }
      } else if (theFreq === "O") {
        let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

        if (skipSundays === "Yes") {
          if (workingDatesStr.includes(endDateStr) && !theDate.is().sunday()) {
            lastMasterTaskID += 1;
            masterArray.push([
              theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID,
              theFreq, theTask, theHow, theDetails, theTime, theMobile,
              new Date(theDate), "", "", theEmail
            ]);
          }
        } else {
          if (workingDatesStr.includes(endDateStr)) {
            lastMasterTaskID += 1;
            masterArray.push([
              theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID,
              theFreq, theTask, theHow, theDetails, theTime, theMobile,
              new Date(theDate), "", "", theEmail
            ]);
          }
        }
      } else if (theFreq === "26D") {
        while (Date.compare(calendarLastDate, theDate) === 1) {
          lastMasterTaskID = lastMasterTaskID + 1
          let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          let frozenDate = theDate.clone()

          if (skipSundays === "Yes") {
            if (workingDatesStr.includes(endDateStr) && !theDate.is().sunday()) {
              masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate), "", "", theEmail])
            }
          } else {
            if (workingDatesStr.includes(endDateStr)) {
              masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate), "", "", theEmail])
            }
          }

          theDate = Date.parse(frozenDate).addDays(26)
        }
      } else if (theFreq === "Q") {
        while (Date.compare(calendarLastDate, theDate) === 1) {
          lastMasterTaskID = lastMasterTaskID + 1
          let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          let frozenDate = theDate.clone()

          while (!workingDatesStr.includes(endDateStr)) {
            theDate = Date.parse(theDate).addDays(-1)
            endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }

          if (skipSundays === "Yes") {
            if (theDate.is().sunday()) {
              theDate = Date.parse(theDate).addDays(-1)
            }
          }

          masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate), "", "", theEmail])

          theDate = Date.parse(frozenDate).addMonths(3)
        }
      } else if (theFreq === "HY") {
        while (Date.compare(calendarLastDate, theDate) === 1) {
          lastMasterTaskID = lastMasterTaskID + 1
          let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          let frozenDate = theDate.clone()

          while (!workingDatesStr.includes(endDateStr)) {
            theDate = Date.parse(theDate).addDays(-1)
            endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }

          if (skipSundays === "Yes") {
            if (theDate.is().sunday()) {
              theDate = Date.parse(theDate).addDays(-1)
            }
          }

          masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate), "", "", theEmail])

          theDate = Date.parse(frozenDate).addMonths(6)
        }
      } else if (theFreq === "F") {
        while (Date.compare(calendarLastDate, theDate) === 1) {
          lastMasterTaskID = lastMasterTaskID + 1
          let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          let frozenDate = theDate.clone()

          while (!workingDatesStr.includes(endDateStr)) {
            theDate = Date.parse(theDate).addDays(-1)
            endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }

          masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate), "", "", theEmail])

          theDate = Date.parse(frozenDate).addWeeks(2)
        }
      } else if (theFreq === "E1st") {
        while (Date.compare(calendarLastDate, theDate) === 1) {
          lastMasterTaskID = lastMasterTaskID + 1
          let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          let day = getDayInfo(theDate.getDay())

          let frozenDate = theDate.clone()

          while (!workingDatesStr.includes(endDateStr)) {
            theDate = Date.parse(theDate).addDays(-1)
            endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }

          masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate), "", "", theEmail])

          theDate = eval("Date.parse(frozenDate).next().month().first()." + day + "()")
        }
      } else if (theFreq === "E2nd") {
        while (Date.compare(calendarLastDate, theDate) === 1) {
          lastMasterTaskID = lastMasterTaskID + 1
          let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          let day = getDayInfo(theDate.getDay())

          let frozenDate = theDate.clone()

          while (!workingDatesStr.includes(endDateStr)) {
            theDate = Date.parse(theDate).addDays(-1)
            endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }

          masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate), "", "", theEmail])

          theDate = eval("Date.parse(frozenDate).next().month().second()." + day + "()")
        }
      } else if (theFreq === "E3rd") {
        while (Date.compare(calendarLastDate, theDate) === 1) {
          lastMasterTaskID = lastMasterTaskID + 1
          let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          let day = getDayInfo(theDate.getDay())

          let frozenDate = theDate.clone()

          while (!workingDatesStr.includes(endDateStr)) {
            theDate = Date.parse(theDate).addDays(-1)
            endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }

          masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate), "", "", theEmail])

          theDate = eval("Date.parse(frozenDate).next().month().third()." + day + "()")
        }
      } else if (theFreq === "E4th") {
        while (Date.compare(calendarLastDate, theDate) === 1) {
          lastMasterTaskID = lastMasterTaskID + 1
          let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          let day = getDayInfo(theDate.getDay())

          let frozenDate = theDate.clone()

          while (!workingDatesStr.includes(endDateStr)) {
            theDate = Date.parse(theDate).addDays(-1)
            endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }

          masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate), "", "", theEmail])

          theDate = eval("Date.parse(frozenDate).next().month().fourth()." + day + "()")
        }
      } else if (theFreq === "ELast") {
        while (Date.compare(calendarLastDate, theDate) === 1) {
          lastMasterTaskID = lastMasterTaskID + 1
          let endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

          let day = getDayInfo(theDate.getDay())

          let frozenDate = theDate.clone()

          while (!workingDatesStr.includes(endDateStr)) {
            theDate = Date.parse(theDate).addDays(-1)
            endDateStr = Utilities.formatDate(theDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }

          masterArray.push([theDoer, theEmail, theDepartment, theTasklistID, lastMasterTaskID, theFreq, theTask, theHow, theDetails, theTime, theMobile, new Date(theDate), "", "", theEmail])

          theDate = eval("Date.parse(frozenDate).next().month().final()." + day + "()")
        }
      }

      taskListSheet.getRange(index + 2, 11).setValue("Sent")
    }

  })
  if (masterArray.length > 0) {
    masterSheet.getRange(masterLastRow + 1, 1, masterArray.length, 15).setValues(masterArray)
  }
}
