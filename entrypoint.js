//
// This is a Google app script for Google Spreadsheet
//

var _Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var _SheetSchedule = _Spreadsheet.getSheetByName("Schedule");

// Get all calendar event for next day
var GoogleCalendarID = "";
const calendar = CalendarApp.getCalendarById(GoogleCalendarID);

const today = new Date();
const nextday = new Date();
nextday.setDate(today.getDate() + 1);
// const Start = new Date(nextday.getFullYear(), nextday.getMonth(), nextday.getDate(), 0, 0, 0, 0);
// const End = new Date(nextday.getFullYear(), nextday.getMonth(), nextday.getDate(), 23, 59, 59, 999);
// const events = calendar.getEvents(Start, End);
// const events = calendar.getEventsForDay(nextday);

Logger.log('The calendar is named "%s".', calendar.getName());

// Lisent to calendar event
function on_week_duty() {
  var range = SpreadsheetApp.getActive().getRangeByName("CALENDAR");
  var values = range.getValues();
  // Logger.log(JSON.stringify(values));

  // get row length in values
  // skip first row, loop 7 days
  var msg = "本週行事曆: \n";
  for (var i = 1; i < 8; i++) {
    msg += format_on_duty_message(values[i]) + "\n";
  }

  // Logger.log(msg);
  if (msg != "") {
    sendLineNotification(msg);
  }
}

function on_day_duty() {
  var range = SpreadsheetApp.getActive().getRangeByName("CALENDAR");
  var values = range.getValues();
  // Logger.log(JSON.stringify(values));

  // get row length in values
  // find next day in row of values
  var rowLength = values.length;
  var msg = "";
  for (var i = 1; i < rowLength; i++) {
    const date_of_nextday = Utilities.formatDate(nextday, "GMT+8", "MM/dd");
    const row_of_nextday = Utilities.formatDate(values[i][0], "GMT+8", "MM/dd");
    if (date_of_nextday == row_of_nextday) {
      msg += format_on_duty_message(values[i]) + "\n";
      msg += "謝謝值班家長";
      break;
    }
  }

  // Logger.log(msg);
  if (msg != "") {
    sendLineNotification(msg);
  }
}

function on_update_calendar() {
  var range = SpreadsheetApp.getActive().getRangeByName("CALENDAR");
  var values = range.getValues();
  // Logger.log(JSON.stringify(values));

  // get row length in values
  var rowLength = values.length;
  for (var i = 1; i < rowLength; i++) {
    var row = values[i];
    update_calendar_on_duty_event(calendar, row);
  }
}

function onOpen() {
  var entries = [];
  entries.push( { name : "LINE - 傳送日值班", functionName : "on_day_duty"} );
  entries.push( { name : "LINE - 傳送週值班表", functionName : "on_week_duty"} );
  entries.push( { name : "Calendar - 更新泳隊行事曆", functionName : "on_update_calendar"} );
  _Spreadsheet.addMenu("[Scripts]", entries);

  _Spreadsheet.setActiveSheet(_SheetSchedule)
};
