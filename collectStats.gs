function collectStats() {
  // Source: https://github.com/jknoll/inbox-depth-and-age/blob/main/collectStats.gs
  // Edit Google Sheet URL below and copy/paste this into a project at https://script.google.com/home
  // Set a trigger for your preferred execution interval here: https://script.google.com/home/triggers
  // Replace this with your Google Sheet URL 
  var googleSheetUrl = 'https://docs.google.com/spreadsheets/d/1OdTqudTUrKnAzz1TZPLkiGIcYV2T3jFsGubk51SRZwc/';

  // Maximum pagesize for GmailApp is 500; if execution time exceeds six minutes script will be killed.
  var pageSize = 500;
  var start = 0;

  var userName = Session.getEffectiveUser().getEmail();
  var googleSheetApp = SpreadsheetApp.openByUrl(googleSheetUrl);
  
  var now = new Date();
  var threads;
  var ages = Array();
  var threadsCount = 0;
  do {
    threads = GmailApp.getInboxThreads(start, pageSize);

    threads.forEach(function(thread) {
      ages.push(dateDiffInDays(thread.getLastMessageDate(), now))
    });

    start += pageSize;
    threadsCount += threads.length;
  } while(threads.length > 0);
  
  var unreadThreadsCount = GmailApp.getInboxUnreadCount();
  
  // Add a row to the first sheet in the Google Sheet.
  var logSheet = googleSheetApp.getSheets()[0];
  var row = [now,userName,unreadThreadsCount,threadsCount, ...ages];
  logSheet.appendRow(row);
};

function dateDiffInDays(d1, d2) {
  return Math.round((datetimeToDate(d2) - datetimeToDate(d1)) / (1000 * 60 * 60 * 24));
}

// Create a date from a date-time
function datetimeToDate(d) {
  return new Date(d.getYear(), d.getMonth(), d.getDate());
}
