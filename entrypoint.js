//
// This is a Google app script for Google Spreadsheet
//

var _Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var _SheetSchedule = _Spreadsheet.getSheetByName("Schedule");
var _CalendarRange = SpreadsheetApp.getActive().getRangeByName("CALENDAR");

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
