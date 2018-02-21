// This is the url of the Google Spreadsheet with the year's events on it
var MASTER_ANNOUNCEMENTS_SHEET = "https://docs.google.com/spreadsheets/d/1wUz2hj1q_FwQssLyyvtOtQEjedGlfQeoGD3lRN9ZrNI/edit";

// This is the IDs of the folder the announcement sheets go in
var FOLDER_IDS = {
  'Monthly Announcement Sheet':'1-OpW_6CEBTuMuYfxoESbwsBzB2kZEqRx',
  'Weekly Announcement Sheet':'1bVLGMP3INZPQQXhFjDDODvqPHYYIWYuS',
  'Sunday Celebration Announcement Sheet':'1eNmUdvGUUnEDA0K3pQumZFNO7TAIPGic',
  'Access Announcement Sheet':'13Vmvurv0ePhoYv9Um3VzPwe50VzR013j'
};

// This is the row number of the headers (Event Name, Location, etc)
var HEADER_ROW_NUM = 2;

// Returns dictionary of dates of requested day of week in next month
function getAnnouncementDays(dayOfWeek){
  var dayNames = ["Sunday", "Monday", "Tuesday", "Wednesday", 
    "Thursday", "Friday", "Saturday"
  ];
  dayOfWeek = dayNames.indexOf(dayOfWeek);
  
  var allDays = {};
  
  var date = new Date();
  var nextMonth = (date.getMonth() + 1) % 12;
  date.setMonth(nextMonth);
  date.setDate(1);
  
  // find first day date
  while (date.getDay() != dayOfWeek) {
    date.setDate(date.getDate()+1);
  }

  // get all days in month
  while (date.getMonth() == nextMonth) {
    allDays[date] = [];
    date.setDate(date.getDate()+7);
  }
  
  return allDays;
}


// Determines which events to announce this month
function getEventsToAnnounce() {
  
  // Gets all the events from MASTER_ANNOUNCEMENTS_SHEET
  var ss = SpreadsheetApp.openByUrl(MASTER_ANNOUNCEMENTS_SHEET).getSheets()[0];
  var allEvents = ss.getRange(HEADER_ROW_NUM+1, 1, ss.getLastRow(), ss.getLastColumn()).sort(3).getValues();
  var headers = ss.getRange(HEADER_ROW_NUM, 1, 1, ss.getLastColumn()).getValues()[0];
  
  var data = {
    'Monthly Announcement Sheet': getAnnouncementDays("Sunday"),
    'Weekly Announcement Sheet': getAnnouncementDays("Thursday"),
    'Sunday Celebration Announcement Sheet': getAnnouncementDays("Sunday"),
    'Access Announcement Sheet': getAnnouncementDays("Friday")
  };
  
  for (var event in allEvents) {
    var eventDate = new Date(allEvents[event][2]);
    var eventYear = Utilities.formatDate(eventDate, "EST", "Y");
    var eventDateInt = parseInt(Utilities.formatDate(eventDate, "EST", "D"));
    
    for (var type in data) {
      for (var date in data[type]) {
        var announcementDateInt = Utilities.formatDate(new Date(date), "EST", "D");
        var announcementYear = Utilities.formatDate(new Date(date), "EST", "Y");
        
        var diff = eventDateInt - announcementDateInt;
        if (diff <= 14 && diff >= 0 && eventYear >= announcementYear) {
          if (type == 'Monthly Announcement Sheet') {
            data[type][date].push(allEvents[event]);
          }
          else if (type == 'Sunday Celebration Announcement Sheet' 
              && allEvents[event][7]) {
            data[type][date].push(allEvents[event]);
          }
          else if (type == 'Access Announcement Sheet' 
              && allEvents[event][8]) {
            data[type][date].push(allEvents[event]);
          }
          else if (type == 'Weekly Announcement Sheet'
              && diff <= 7) {
            data[type][date].push(allEvents[event]);
          }
        }
          
      }
    }
    
  }
  return data;
}

function getEventDetails(data) {
  var date = ""; 
  if (data[2]) {
    date += Utilities.formatDate(new Date(data[2]), "EST", "EEE M/dd");
  }
  if (data[3]) {
    date += (" - " + Utilities.formatDate(new Date(data[3]), "EST", "EEE M/dd"));
  }

  var time = ""; 
  if (data[4]) {
    time += Utilities.formatDate(new Date(data[4]), "EST", "h:mm a");
  }
  if (data[5]) {
    time += (" - " + Utilities.formatDate(new Date(data[5]), "EST", "h:mm a"));
  }
  
  var location = "";
  if (data[6]) {
    location += data[6];
  }
  
  return date + '\t' + time + '\t' + location;
}

// Formats monthlyAnntSheet
function genMonthlyAnntSheet(type, doc, data) {
  
  var docBody = doc.getBody();
  // clear existing
  docBody.clear();
  
  var title = docBody.appendParagraph(type).editAsText();
  title.setBold(true);
  title.setFontSize(16);
  
  for (var i in data[type]) {
    
    // Sunday Date
    var sundayDateFormatted = Utilities.formatDate(new Date(i), "EST", "EEE M/dd");
    var sundayDate = docBody.appendParagraph(sundayDateFormatted).editAsText();
    sundayDate.setBold(true);
    sundayDate.setFontSize(14);
    
    for (var event in data[type][i]) {
    
      // Event Name
      var eventName = docBody.appendParagraph('\r' + data[type][i][event][0]).editAsText();
      eventName.setBold(true);
      eventName.setFontSize(11);
      
      // Event Details (Dates/Time/Location)
      var details = getEventDetails(data[type][i][event]);
      var eventDetails = docBody.appendParagraph(details).editAsText();
      eventDetails.setBold(false);
      eventDetails.setItalic(true);
      
      // Event Description
      var eventData = data[type][i][event][1].split('\n');
      for (var line in eventData) {
        var eventDescription = docBody.appendListItem(eventData[line]).setGlyphType(DocumentApp.GlyphType.BULLET).editAsText();
        eventDescription.setBold(false);
        eventDescription.setItalic(false);
      }
      
      docBody.appendHorizontalRule();
    }
    
   docBody.appendPageBreak();  
  }
}

// retrieve existing AnntSheet to update or create new Google Doc for this month
function getAnntDocument(dateTitle, type) {
  var docFileIterator = DriveApp.getFilesByName(dateTitle + ' ' + type);
  if (docFileIterator.hasNext()) {
    var doc = DocumentApp.openById(docFileIterator.next().getId()); 
  } else {
    var doc = DocumentApp.create(dateTitle + ' ' + type);
  }
  var docFile = DriveApp.getFileById(doc.getId());
  
  return doc;
}

// Makes new announcements sheet for next month
function createNewAnnouncements() {
  var date = new Date();
  var nextMonth = (date.getMonth() + 1) % 12;
  date.setMonth(nextMonth);
  
  var dateTitle = Utilities.formatDate(date, 'EST', 'Y MMMMMMMMM');
  
  var data = getEventsToAnnounce();
  
  for (var type in FOLDER_IDS) {
    var folder = DriveApp.getFolderById(FOLDER_IDS[type]);
    var doc = getAnntDocument(dateTitle, type);

    genMonthlyAnntSheet(type, doc, data);
    
    folder.addFile(DriveApp.getFileById(doc.getId()));
  }
}
 