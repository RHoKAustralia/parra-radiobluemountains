// RHoK 26/11/2016

var consolidatedCalendar = CalendarApp.getCalendarById("0s6n1i40gorq5m4uad57ec0d5c@group.calendar.google.com");
var crnCalendar = CalendarApp.getCalendarById("1cs29j3n6ebb3n8d4pa98flhoc@group.calendar.google.com");
var rbmCalendar = CalendarApp.getCalendarById("7l3qtj8h2dud0jthrmm4cp6hss@group.calendar.google.com");



function copyFromCrn(description, startTime, endTime){
  var crnEvents = crnCalendar.getEvents(startTime, endTime);
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
    consolidatedCalendar.createEvent("CRN: "+crnEvent.getTitle(), eventStartTime , finishTime);
  }
}

// Copies everything from RMB, find the free gaps and then fill them with CRN programs (sometimes partialy)
function mergeCalendars() {
  var rbmCalendar = CalendarApp.getCalendarById("7l3qtj8h2dud0jthrmm4cp6hss@group.calendar.google.com");
    
  var agendaDay = new Date('Mon Nov 20 2016 00:00:00 GMT+1100 (AEDT)');
  
  
  // cleaning up consolidated calendar  
  var consolidatedEvents = consolidatedCalendar.getEventsForDay(agendaDay);
  for (var i = 0; i < consolidatedEvents.length; i++) {
    consolidatedEvents[i].deleteEvent();    
  }
  
  
  var rbmEvents = rbmCalendar.getEventsForDay(agendaDay);
  
  Logger.log("Cloning events from RMB...");
  for (var i = 0; i < rbmEvents.length; i++) {
    var event = rbmEvents[i];
    var startFreeTime;
    var finishTime;
    if(i == 0){
      copyFromCrn("First CRN Program", agendaDay, event.getStartTime());
      
    }     
    startFreeTime = event.getEndTime();
    if(i == rbmEvents.length-1){
      finishTime = new Date('Mon Nov 21 2016 00:00:00 GMT+1100 (AEDT)')
    } else { 
      finishTime = rbmEvents[i+1].getStartTime();               
    }
    
    copyFromCrn("CRN Program "+i, startFreeTime, finishTime);
    
    consolidatedCalendar.createEvent(event.getTitle(), event.getStartTime(), event.getEndTime());
  }
  
  
  
  
  
}
