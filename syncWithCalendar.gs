// This is the url of the Google Spreadsheet with the year's events on it
var MASTER_ANNOUNCEMENTS_SHEET = "https://docs.google.com/spreadsheets/d/1wUz2hj1q_FwQssLyyvtOtQEjedGlfQeoGD3lRN9ZrNI/edit";

// This is the id of the Google Calendar with the year's events on it
var CALENDAR_ID = "52l95kpefa66kqpi0th0nnjbtk@group.calendar.google.com";//Beta thing

function syncCalToSheet() {
  var ss = SpreadsheetApp.openByUrl(MASTER_ANNOUNCEMENTS_SHEET).getSheets()[0];
  
  var todaysDate = new Date();
  
  var calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  var allEvents = calendar.getEvents(new Date('2017', '1'), todaysDate);
  
  for (var event in allEvents) {
    var eventName = allEvents[event].getTitle();
    var eventDescription = allEvents[event].getDescription();
    
    var eventStartDate = Utilities.formatDate(allEvents[event].getStartTime(), "EST", "M/d");
    if (allEvents[event].getStartTime().getDate() != allEvents[event].getEndTime().getDate()) {
      var eventEndDate = Utilities.formatDate(allEvents[event].getEndTime(), "EST", "M/d");
    } else {
      eventEndDate = null;
    }
    
    var eventStartTime = Utilities.formatDate(allEvents[event].getStartTime(), "EST", "h:mm a");
    if (allEvents[event].getStartTime().getTime() != allEvents[event].getEndTime().getTime()) {
      var eventEndTime = Utilities.formatDate(allEvents[event].getEndTime(), "EST", "h:mm a");
    } else {
      eventEndTime = null;
    }
    
    var eventLocation = allEvents[event].getLocation();

    ss.appendRow([eventName, eventDescription, eventStartDate, eventEndDate, eventStartTime, eventEndTime, eventLocation]);
  }
}
  
