//
// This is a Google app script for Google Spreadsheet
//

var _Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

// Sheets: Schedule, Calendar, 行事曆版本, 文字版, Contact
var _SheetSchedule = _Spreadsheet.getSheetByName("Schedule");
// var _SheetCalendar = _Spreadsheet.getSheetByName("Calendar");
// var _SheetCalendarText = _Spreadsheet.getSheetByName("文字版");
// var _SheetCalendarVersion = _Spreadsheet.getSheetByName("行事曆版本");
// var _SheetContact = _Spreadsheet.getSheetByName("Contact");
var _SheetCheckIn = _Spreadsheet.getSheetByName("CheckIn");

// Named Ranges:
const _rangeCALENDAR = SpreadsheetApp.getActive().getRangeByName("CALENDAR");
// const _rangeDUTY = SpreadsheetApp.getActive().getRangeByName("DUTY");
// const _rangeSCHEDULE = SpreadsheetApp.getActive().getRangeByName("SCHEDULE");
// const _rangeDUTYMAP = SpreadsheetApp.getActive().getRangeByName("DUTYMAP");
// const _rangeScheduleHeader = SpreadsheetApp.getActive().getRangeByName("ScheduleHeader");
// const _rangeScheduleName = SpreadsheetApp.getActive().getRangeByName("ScheduleName");
// const _rangeSDATE = SpreadsheetApp.getActive().getRangeByName("SDATE");
// const _rangeWEEKMAP = SpreadsheetApp.getActive().getRangeByName("WEEKMAP");

const _today = new Date();
const _nextday = new Date();
_nextday.setDate(_today.getDate() + 1);

function onOpen() {
    var entries = [];
    entries.push({ name: "LINE - 傳送日值班", functionName: "on_day_duty" });
    entries.push({ name: "LINE - 傳送週值班表", functionName: "on_week_duty" });
    entries.push({ name: "Google Calendar - 更新泳隊行事曆", functionName: "on_update_calendar" });
    _Spreadsheet.addMenu("[Scripts]", entries);

    _Spreadsheet.setActiveSheet(_SheetSchedule)
};

function on_duty() {
    // call on_week_duty() on Saturday, and on_day_duty() on other days
    var weekday = _today.getDay();
    if (weekday == 6) {
        on_week_duty();
    } else {
        on_day_duty();
    }
}

function on_week_duty() {
    send_week_duty_notification(prod_LineNotifyAccessToken);
}

function on_day_duty() {
    send_day_duty_notification(prod_LineNotifyAccessToken);
}

function on_update_calendar() {
    const calendar = CalendarApp.getCalendarById(prod_GoogleCalendarID);
    update_calendar_by_sheet(calendar);
}
