// This is the url of the Google Spreadsheet with the year's events on it
var MASTER_ANNOUNCEMENTS_SHEET = "https://docs.google.com/spreadsheets/d/1wUz2hj1q_FwQssLyyvtOtQEjedGlfQeoGD3lRN9ZrNI/edit";
// This is the row number of the headers (Event Name, Location, etc)
var HEADER_ROW_NUM = 2;
// This is the ID of the folder the monthly announcement sheets go in
var FOLDER_ID = '1eNmUdvGUUnEDA0K3pQumZFNO7TAIPGic';

//TODO: check date formats or something idk. make sure the events are going in the right data. make sure key is right. make sure value is in good format for outputting.

// Returns dictionary of dates of Sundays in next month
function getSundaysInMonth(){
  var allSundays = {};
  
  var sundayDate = new Date();
  var nextMonth = (sundayDate.getMonth() + 1) % 12;
  sundayDate.setMonth(nextMonth);
  sundayDate.setDate(1);
  
  while (true) {
    var diff = sundayDate.getDate() + (7 - sundayDate.getDay());
    sundayDate = new Date(sundayDate.setDate(diff));
    if (sundayDate.getMonth() != nextMonth) {
      break;
    }
    allSundays[sundayDate] = []; 
  }
  return allSundays;
}

// Determines which events to announce this month
function getEventsToAnnounce(allEvents) {
  var data = getSundaysInMonth();
  
  for (var date in data) {    
    var sundayDateInt = Utilities.formatDate(new Date(date), "EST", "D");
    var sundayYear = Utilities.formatDate(new Date(date), "EST", "Y");
    
    for (var event in allEvents) {
      var eventDate = new Date(allEvents[event][7]);
      var eventYear = Utilities.formatDate(eventDate, "EST", "Y");
      var eventDateInt = parseInt(Utilities.formatDate(eventDate, "EST", "D"));
      
      var diff = eventDateInt - sundayDateInt;
      if (diff < 20 && diff > 0 && eventYear >= sundayYear) {
        data[date].push(allEvents[event]);
      }
    }
  }
  return data;
}

// Formats monthlyAnntSheet
function genMonthlyAnntSheet(doc) {
  // Gets all the events from MASTER_ANNOUNCEMENTS_SHEET
  var ss = SpreadsheetApp.openByUrl(MASTER_ANNOUNCEMENTS_SHEET).getSheets()[0];
  var allEvents = ss.getRange(HEADER_ROW_NUM+1, 1, ss.getLastRow(), ss.getLastColumn()).getValues();
  var headers = ss.getRange(HEADER_ROW_NUM, 1, 1, ss.getLastColumn()).getValues()[0];
  
  var data = getEventsToAnnounce(allEvents);

  var docBody = doc.getBody();
  
  for (var i in data) {
    
    // Sunday Date
    var sundayDateFormatted = Utilities.formatDate(new Date(i), "EST", "EEE M/dd");
    var sundayDate = docBody.appendParagraph(sundayDateFormatted).editAsText();
    sundayDate.setBold(true);
    sundayDate.setFontSize(14);
    
    for (var event in data[i]) {
    
      // Event Name
      var eventName = docBody.appendParagraph('\r' + data[i][event][0]).editAsText();
      eventName.setBold(true);
      eventName.setFontSize(11);
      
      // Event Details (Dates/Time/Location)
      var eventDetails = docBody.appendParagraph(data[i][event][3] + '\t' + data[i][event][2] + '\t' + data[i][event][4]).editAsText();
      eventDetails.setBold(false);
      eventDetails.setItalic(true);
      
      // Event Description
      var eventDescription = docBody.appendParagraph(data[i][event][1]).editAsText();
      eventDescription.setBold(false);
      eventDescription.setItalic(false);
    }
    
   docBody.appendHorizontalRule();    
  }
}

// Makes new announcements sheets for the 5 categories
function createNewAnnouncements() {
  var folder = DriveApp.getFolderById(FOLDER_ID);
  
  var date = new Date();
  var dateTitle = Utilities.formatDate(date, 'EST', 'M/Y');
  var type = 'Monthly Annt Sheet';
  
  // retrieve existing AnntSheet to update or create new Google Doc for this month
  var docFileIterator = DriveApp.getFilesByName(dateTitle + ' ' + type);
  if (docFileIterator.hasNext()) {
    var doc = DocumentApp.openById(docFileIterator.next().getId()); 
  } else {
    var doc = DocumentApp.create(dateTitle + ' ' + type);
  }
  var docFile = DriveApp.getFileById(doc.getId());
  
  // populate Annt Sheet
  genMonthlyAnntSheet(doc);
  
  // add new Annt Sheet to folder
  folder.addFile(docFile);
}
 