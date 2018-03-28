function createEventForSheet(event) {
  var eventId = event.getId();
  
  var eventName = event.getTitle();
  var eventDescription = event.getDescription();
  
  var eventStartTime = event.getStartTime();
  var eventEndTime = event.getEndTime();

  var eventStartDateFormatted = Utilities.formatDate(eventStartTime, "EST", "M/d");
  var eventEndDateFormatted = 
      (eventStartTime.getDate() != eventEndTime.getDate()) ?
        Utilities.formatDate(eventEndTime, "EST", "M/d") :
        null;
  
  var eventStartTimeFormatted = Utilities.formatDate(eventStartTime, "EST", "h:mm a");
  var eventEndTimeFormatted = 
      (eventStartTime.getTime() != eventEndTime.getTime()) ?
        Utilities.formatDate(eventEndTime, "EST", "h:mm a") :
        null;
  
  var eventLocation = event.getLocation();
  
  return [eventName, eventDescription, eventStartDateFormatted, eventEndDateFormatted, eventStartTimeFormatted, eventEndTimeFormatted, eventLocation, null, null, eventId, eventStartTime];
}

function eventInSheet(event, eidToSheetRow, sheetEvents) {
  var eid = event.getId();
  if (eid in eidToSheetRow) {
    if (!event.isRecurringEvent()) {
      return true;
    }
    
    for (var row in eidToSheetRow[eid]) {
      if(new Date(sheetEvents[eidToSheetRow[eid][row]][10]).getTime() === new Date(event.getStartTime()).getTime()) {
        return true;
      }
    }
  }
  return false;
}


function syncCalToSheet() {
  
  var ss = SpreadsheetApp.openByUrl(MASTER_ANNOUNCEMENTS_SHEET).getSheets()[0];
  var sheetEvents = ss.getRange(2, 1, ss.getLastRow(), ss.getLastColumn()).sort(3).getValues();
  var eidToSheetRow = {};
  var eventsToAdd = [];
  
  // row number to event
  for (var x in sheetEvents) {
    var eid = sheetEvents[x][9];
    if (!(eid in eidToSheetRow)) eidToSheetRow[eid] = [];
    eidToSheetRow[eid].push(x);
  }
  
  var calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  var calEvents = calendar.getEvents(new Date('2018', '2', '1'), new Date('2018', '2', '31'));

  for (var x in calEvents) {
    var event = calEvents[x];
    var eid = event.getId();
    // case for events on cal but not on sheet
    if (!(eventInSheet(event, eidToSheetRow, sheetEvents))) {
      var newEvent = createEventForSheet(event);
      eventsToAdd.push(newEvent);
    }
  }
  
  for (var x in eventsToAdd) {
    ss.appendRow(eventsToAdd[x]);
  }
}
