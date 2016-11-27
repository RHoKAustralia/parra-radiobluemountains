/**
* A special function that runs when the spreadsheet is open, used to add a
* custom menu to the spreadsheet.
*/
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Export as Calendar', functionName: 'exportCalendar'},
    {name: 'Merge CRN Calendar', functionName: 'goThroughTheWeek' }    
  ];
  spreadsheet.addMenu("Radio program", menuItems);  
}


function doGet(){
  exportCalendar();
  return ContentService.createTextOutput('Yay! Calendar exported!');
}

function initialize(){
  
  if(!consolidatedCalendar)
    consolidatedCalendar = CalendarApp.getCalendarById("lvgac2dkit0n96ttb2a24dttrs@group.calendar.google.com");
  if(!rbmCalendar)
    rbmCalendar = CalendarApp.getCalendarById("7l3qtj8h2dud0jthrmm4cp6hss@group.calendar.google.com");
}

var consolidatedCalendar = undefined;
var rbmCalendar = undefined;

var apiCallCount = 0;

var calendarName = "Imported CRN program guide";

var DAYLIGHT_SAVING = 1;
var DEBUG = false;

function getCrnCalendar(){      
  // DROP AND RECREATE THE CALENDAR  
  var calendars = CalendarApp.getCalendarsByName(calendarName);
  if (calendars.length > 0) {
    Logger.log("Deleting calendar");
    calendars[0].deleteCalendar();
  }
  var calendarOptions = {
    summary: 'Community Radio Program',
    color: '#92e1c0',
    timeZone: 'Australia/Sydney'
  };
  return CalendarApp.createCalendar(calendarName, calendarOptions);
  
}

function showLink() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  var range = sheet.getRange(3, 2);
  var formulas = range.getFormulas();
  var output = [];
  for (var i = 0; i < formulas.length; i++) {
    var row = [];
    for (var j = 0; j < formulas[0].length; j++) {
      var url = formulas[i][j].match(/=hyperlink\("([^"]+)"/i);
      row.push(url ? url[1] : '');
    }
    output.push(row);
    Logger.log(row);
  }
}

function exportCalendar() {
  initialize();
  var sheet = SpreadsheetApp.getActiveSheet();
  var programEvents = [];
  
  
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
  
  
  // For each row
  var prevName = "", prevDayName = "";
  var prevTime;
  var timePrograms = {};
  var firstDayPrograms = {};
  for (var j = 0; j < 7; j++) {
    for (var i = 2; i < 50; i++) {
      // For each day of the week
      var startTime = sheet.getRange(i, 1).getValue();
      startTime = convertDate(startTime);
      var programName = sheet.getRange(i, 2 + j).getValue();
      try {
        if (programName == "" && startTime.toString() != "") {
          // Check if the first day of the week has a program that extends
          if (firstDayPrograms[i] === undefined) {
            prevTime = startTime; 
            continue;
          }
          programName = firstDayPrograms[i];
        }
        else {
          delete(firstDayPrograms[i]);
        }
      }
      catch (e) {
        continue;
      }
      if (startTime.toString() == "") {
        startTime = prevTime;
      }
      prevTime = startTime; 
      
      
      try {
        var startDate = new Date(firstMonday);
        startDate.setDate(firstMonday.getDate() + j);
        startDate.setHours(startTime.getHours() - DAYLIGHT_SAVING);
        startDate.setMinutes(startTime.getMinutes());
        startDate.setSeconds(startTime.getSeconds());
        
        var counter = 1;
        var nextProg = "";
        do {
          nextProg = sheet.getRange(i + counter, 2 + j).getValue();
          endTime = sheet.getRange(i + counter, 1).getValue();
          // If the end time value is empty, check the preceding value
          if (nextProg != "" && endTime == "") {
            endTime = sheet.getRange(i + counter - 1, 1).getValue();
          }
          endTime = convertDate(endTime);
          
          // If the first day of the week has a program at this time, 
          // we've hit the end of the current program
          if (firstDayPrograms[i + counter] !== undefined) {
            nextProg = firstDayPrograms[i + counter];
          }
          
          // If more than 6 cells without an end time and no name has been found, break
        } while (nextProg == "" && (++counter) < 6);
        
        var endDate = new Date(firstMonday);
        endDate.setYear(startDate.getYear());
        endDate.setMonth(startDate.getMonth());
        endDate.setDate(startDate.getDate());
        if (endTime.toString() == "") {
          startDate.setDate(startDate.getDate() + 1);
          endDate.setMonth(startDate.getMonth());
          endDate.setDate(startDate.getDate());
          endDate.setHours(23);
          endDate.setMinutes(59);
          endDate.setSeconds(00);
        }
        else {
          // More strangeness related to daylight saving
          endDate.setHours(endTime.getHours() - DAYLIGHT_SAVING);
          endDate.setMinutes(endTime.getMinutes());
          endDate.setSeconds(endTime.getSeconds());
          if (endDate.getHours() == 23) {
            endDate.setDate(startDate.getDate());
            endDate.setMonth(startDate.getMonth());
            endDate.getYear(startDate.getYear());
          }
        }
        // Add the program name for the first day
        if (j == 0) {
          firstDayPrograms[i] = programName;
        }
        
        
        // If the program name has changed, this is indeed a new event. So create it!
        if (prevName != programName && programName != "") {
          programEvents.push({ startDate: startDate, endDate: endDate, programName: programName});
        }
        prevName = programName;
      }
      catch (e) {
        Logger.log('Exception: ' + e + ":"+programName + ":"+ startDate + ":"+ endDate);
      }
    }
  }
  
  // Recreate the calendar
  var calendar;
  if (!DEBUG) {
    
    calendar = getCrnCalendar();
    
    
    
    // DO THIS TO REMOVE ALL EVENTS FROM A CURRENT CALENDAR
    /*
    if (crnCalendar === undefined) {
    crnCalendar = CalendarApp.getCalendarsByName(calendarName)[0];
    }
    var nextMonday = new Date(firstMonday);
    nextMonday.setDate(firstMonday.getDate() + 7);
    var events = crnCalendar.getEvents(firstMonday, nextMonday);
    var l = events.length;
    for (var i = 0; i < l; i++) {
    var event = events[i];
    event.deleteEvent();
    }
    */
    
  }
  
  // Create the events
  for (var i = 0; i < programEvents.length; i++) {
    var event = programEvents[i];
    try{
      if (!DEBUG) {
        createProgramEvent(calendar, event.startDate, event.endDate, event.programName);
      }
      else {
        Logger.log(i +": " + event.programName + ":"+ event.startDate + ":"+ event.endDate);
      }
    }
    catch (e) {
      Logger.log('Exception: ' + e + ":"+event.programName + ":"+ event.startDate + ":"+ event.endDate);
    }
  }
  
  // Merge the new CRN calendar
  //goThroughTheWeek();
}

function createProgramEvent(calendar, startTime, endTime, programName) {
  registerApiCall();
  // Creates an event.
  var options = {}; //      {location: 'The Moon'}
  var event = calendar.createEvent(programName,
                                   startTime,
                                   endTime, 
                                   options);  
}

function convertDate(d) {
  var s = d.toString();
  if (s.indexOf('\n') > 0) {
    s = s.substring(0, s.indexOf('\n'));
    var h = s.substring(0, s.indexOf(':'));
    var m = s.substring(s.indexOf(':') + 1);
    d = new Date();
    // Handle daylight saving issue
    d.setHours(h + DAYLIGHT_SAVING);
    if (d.getHours() < 12)
      d.setHours(d.getHours() + 12);
    //d.setHours(h);
    d.setMinutes(m);
  }
  return d;
}


function addDay(date, days){  
  var newDate = new Date( date.getTime() + days * 86400000 );
  return newDate;
}

function goThroughTheWeek(){
  initialize();
  var agendaDay = new Date();
  agendaDay.setHours(0);
  agendaDay.setMinutes(0);
  agendaDay.setSeconds(0);
  agendaDay.setMilliseconds(0);
  
  for(var i = 0;i<7;i++){
    mergeCalendars(addDay(agendaDay,i));    
  }
}

function registerApiCall(){
  apiCallCount++;
  if(apiCallCount >= 20){
    Logger.log("Sleeping a bit to give Google a rest");
    Utilities.sleep(10000); // avoid google's api restrictions
    apiCallCount = 0;
  }
}

function copyFromCrn(startTime, endTime){
  registerApiCall();
  if(startTime > endTime){
    Logger.log("Invalid interval for "+startTime+" and "+endTime);
    return;
  }
  
  var crnEvents = CalendarApp.getCalendarsByName(calendarName)[0].getEvents(startTime, endTime);
  for(i = 0;i<crnEvents.length;i++){
    var crnEvent = crnEvents[i];
    var finishTime  = crnEvent.getEndTime();
    var eventStartTime = crnEvent.getStartTime();
    
    if(eventStartTime < startTime){
      eventStartTime = startTime;
    }
    if(finishTime > endTime){
      finishTime = endTime;
    }    
    consolidatedCalendar.createEvent(crnEvent.getTitle(), eventStartTime , finishTime);
    
  }
}

// Copies everything from RMB, find the free gaps and then fill them with CRN programs (sometimes partialy)
function mergeCalendars(agendaDay) {
  Logger.log("Cloning events from RMB for "+agendaDay);
  
  // cleaning up consolidated calendar  
  var consolidatedEvents = consolidatedCalendar.getEventsForDay(agendaDay);
  for (var i = 0; i < consolidatedEvents.length; i++) {
    registerApiCall();
    consolidatedEvents[i].deleteEvent();    
  }
  
  var rbmEvents = rbmCalendar.getEventsForDay(agendaDay);
  
  if(rbmEvents.length == 0){
    copyFromCrn(agendaDay,addDay(agendaDay, 1));
  }
  
  for (var i = 0; i < rbmEvents.length; i++) {
    var event = rbmEvents[i];
    var startFreeTime;
    var finishTime;
    if(i == 0){
      copyFromCrn(agendaDay, event.getStartTime());
      
    }     
    startFreeTime = event.getEndTime();
    if(i == rbmEvents.length-1){
      finishTime = addDay(agendaDay, 1);
    } else { 
      finishTime = rbmEvents[i+1].getStartTime();               
    }
    
    
    copyFromCrn(startFreeTime, finishTime);    
    registerApiCall();
    consolidatedCalendar.createEvent(event.getTitle(), event.getStartTime(), event.getEndTime());
  }
  
}
