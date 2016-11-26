/** * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Export as Calendar...', functionName: 'exportCalendar'}
  ];
  spreadsheet.addMenu('Radio Program', menuItems);
}

function exportCalendar() {
  var sheet = SpreadsheetApp.getActiveSheet();

  // Recreate the calendar
  var calendarName = "Radio Program";
  var calendars = CalendarApp.getCalendarsByName(calendarName);
  if (calendars.length > 0) {
    Logger.log("Deleting calendar");
    calendars[0].deleteCalendar();
  }
  var calendar = CalendarApp.createCalendar(calendarName, {});

  // Assume Calendar begins next Monday
  var d = new Date();
  var firstMonday = new Date();
  var n = d.getDay();
  if (n == 0) {
    firstMonday.setDate(d.getDate() + 1);
  }
  else if (n > 1) {
    // Set to the *following* Monday
    firstMonday.setDate(d.getDate() + (8 - n));
  }
  Logger.log(firstMonday.toString());
  
  // For each row
  var prevName = "";
  var prevTime;
  for (var j = 0; j < 7; j++) {
    for (var i = 2; i < 50; i++) {
      // For each day of the week
      var startTime = sheet.getRange(i, 1).getValue();
      if (startTime.toString() == "") {
        startTime = prevTime;
      }
      prevTime = startTime; 
      var programName = sheet.getRange(i, 2 + j).getValue();
      if (programName == "") 
        continue;
      
      try {
        var startDate = new Date(firstMonday);
        startDate.setDate(firstMonday.getDate() + j);
        startDate.setHours(startTime.getHours() - 1);
        startDate.setMinutes(startTime.getMinutes());
        startDate.setSeconds(startTime.getSeconds());
        
        var counter = 1;
        var nextProg = "";
        do {
          endTime = sheet.getRange(i + counter, 1).getValue();
          nextProg = sheet.getRange(i + counter, 2 + j).getValue();
          counter++;
          // If more than 3 cells without an end time, break
          if (counter > 3)
            break;
        } while (nextProg == "");
        if (endTime.toString() == "") {
          endTime = new Date(startDate);
          endTime.setHours(24);
          endTime.setMinutes(59);
        }
        
        var endDate = new Date(startDate);
        endDate.setHours(endTime.getHours() - 1);
        endDate.setMinutes(endTime.getMinutes());
        endDate.setSeconds(endTime.getSeconds());
        createProgramEvent(calendar, startDate, endDate, programName);
        prevName = programName;
      }
      catch (e) {}
    }
  }
}

function createProgramEvent(calendar, startTime, endTime, programName) {
  // Creates an event.
  var options = {}; //      {location: 'The Moon'}
  var event = calendar.createEvent(programName,
     startTime,
     endTime, 
     options);
  Logger.log('Event ID: ' + event.getId());
}
