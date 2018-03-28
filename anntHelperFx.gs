function formatDate(month, year) {
  var monthToNum = {'January':0, 'February':1, 'March':2, 'April':3, 'May':4, 'June':5, 'July':6, 'August':7, 'September':8, 'October':9, 'November':10, 'December':11};
  
  var date = new Date();
  date.setDate(1);
  date.setMonth(monthToNum[month]);
  date.setYear(year);
  return date;
}


function getEventDetails(data) {
  var date = ""; 
  if (data[2]) {
    Logger.log(data[2]);
    date += Utilities.formatDate(new Date(data[2] + "EST"), "EST", "EEE M/dd");
    Logger.log(date);    
  }
  if (data[3]) {
    date += (" - " + Utilities.formatDate(new Date(data[3] + "EST"), "EST", "EEE M/dd"));
  }

  var time = ""; 
  if (data[4]) {
    time += Utilities.formatDate(new Date(data[4] + "EST"), "EST", "h:mm a");
  }
  if (data[5]) {
    time += (" - " + Utilities.formatDate(new Date(data[5] + "EST"), "EST", "h:mm a"));
  }
  
  var location = "";
  if (data[6]) {
    location += data[6];
  }
  
  return date + ', ' + time + ' @ ' + location;
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
    var sundayDateFormatted = Utilities.formatDate(new Date(i + "EST"), "EST", "EEE M/dd");
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

// Returns dictionary of dates of requested day of week in next month
function getAnnouncementDays(dayOfWeek, date){
  var dayNames = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  dayOfWeek = dayNames.indexOf(dayOfWeek);
  
  var allDays = {};
  
  date = new Date(date.getYear(), date.getMonth(), 1);
  var month = date.getMonth();
  
  // find first day date
  while (date.getDay() != dayOfWeek) {
    date.setDate(date.getDate()+1);
  }

  // get all days in month
  while (date.getMonth() == month) {
    allDays[date] = [];
    date.setDate(date.getDate()+7);
  }
  
  return allDays;
}

function getAccessAnnouncementDays(date) {
  
  var accessFolderId = FOLDER_IDS['Access Announcement Sheet'];
  var accessFolder = DriveApp.getFolderById(accessFolderId);
  var accessDatesFile = accessFolder.getFilesByName('Access Dates').next();
  
  var accessDatesSpreadsheet = SpreadsheetApp.open(accessDatesFile);
  var ss = accessDatesSpreadsheet.getSheets()[0];
  var accessDates = ss.getRange(2, 1, ss.getLastRow()).getValues();

  var allDays = {};
  
  date = new Date(date.getYear(), date.getMonth(), 1);
  var month = date.getMonth();
  
  for (var x in accessDates) {
    var date = new Date(accessDates[x]);
    if (date.getMonth() == month) {
      allDays[accessDates[x]] = [];
    }
  }
  
  return allDays
}


// Determines which events to announce this month
function getEventsToAnnounce(date) {
  
  // Gets all the events from MASTER_ANNOUNCEMENTS_SHEET
  var ss = SpreadsheetApp.openByUrl(MASTER_ANNOUNCEMENTS_SHEET).getSheets()[0];
  var allEvents = ss.getRange(2, 1, ss.getLastRow(), ss.getLastColumn()).sort(3).getValues();
  var headers = ss.getRange(1, 1, 1, ss.getLastColumn()).getValues()[0];
  
  var data = {
    'Monthly Announcement Sheet': getAnnouncementDays("Sunday", date),
    'Weekly Announcement Sheet': getAnnouncementDays("Thursday", date),
    'Sunday Celebration Announcement Sheet': getAnnouncementDays("Sunday", date),
    'Access Announcement Sheet': getAccessAnnouncementDays(date)
  };
  
  for (var event in allEvents) {
    var eventDate = new Date(allEvents[event][2] + "EST");
    var eventYear = Utilities.formatDate(eventDate, "EST", "Y");
    var eventDateInt = parseInt(Utilities.formatDate(eventDate, "EST", "D"));
    
    for (var type in data) {
      for (var date in data[type]) {
        var announcementDateInt = Utilities.formatDate(new Date(date + "EST"), "EST", "D");
        var announcementYear = Utilities.formatDate(new Date(date + "EST"), "EST", "Y");
        
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

// Makes new announcements sheet for specified month
function createNewAnnouncements(date) {

  var dateTitle = Utilities.formatDate(date, 'EST', 'Y MMMMMMMMM');

  var data = getEventsToAnnounce(date);

  var rootFolder = DriveApp.getRootFolder();
  
  for (var type in FOLDER_IDS) {
    var folder = DriveApp.getFolderById(FOLDER_IDS[type]);
    var doc = getAnntDocument(dateTitle, type);

    genMonthlyAnntSheet(type, doc, data);
    
    var file = DriveApp.getFileById(doc.getId());
    folder.addFile(file);
    rootFolder.removeFile(file);
  }

}
 