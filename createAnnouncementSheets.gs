// This is the url of the Google Spreadsheet with the year's events on it
var MASTER_ANNOUNCEMENTS_SHEET = "https://docs.google.com/spreadsheets/d/1wUz2hj1q_FwQssLyyvtOtQEjedGlfQeoGD3lRN9ZrNI/edit";

// This is the id of the Google Calendar with the year's events on it
var CALENDAR_ID = "52l95kpefa66kqpi0th0nnjbtk@group.calendar.google.com";

// This is the IDs of the folder the announcement sheets go in
var FOLDER_IDS = {
  'Monthly Announcement Sheet':'1-OpW_6CEBTuMuYfxoESbwsBzB2kZEqRx',
  'Weekly Announcement Sheet':'1bVLGMP3INZPQQXhFjDDODvqPHYYIWYuS',
  'Sunday Celebration Announcement Sheet':'1eNmUdvGUUnEDA0K3pQumZFNO7TAIPGic',
  'Access Announcement Sheet':'13Vmvurv0ePhoYv9Um3VzPwe50VzR013j'
};

function main() { 
  // WHAT MONTH DO YOU WANT?
  var month = 'May';
  var year = 2018;
  
  var date = formatDate(month, year);
  createNewAnnouncements(date);
}