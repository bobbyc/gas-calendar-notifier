//
// This is a Google app script for Google Spreadsheet
//

function format_on_duty_message(values) {
    var datetime = Utilities.formatDate(values[0], "GMT+8", "MM/dd (E) ");
    var time = values[2];
    var note = values[3];
    var location = values[4];
    var duty1 = values[5];
    var duty2 = values[6];
    var onduty = Utilities.formatString("@%s @%s", duty1, duty2);

    if (time != "")
    datetime += Utilities.formatDate(new Date(values[2]), "GMT+8", "HH:mm");

    // return formatted message string with date, time, location, and onduty
    var msg = "";
    if (note != "") {
        msg += Utilities.formatString("日期: %s (%s)\n", datetime, note);
    } else {
        msg += Utilities.formatString("日期: %s\n", datetime);
    }

    if (location != "") {
        msg += Utilities.formatString("地點: %s\n", location);
    }

    if (duty1 != "" && duty2 != "") {
        msg += Utilities.formatString("值班家長:  %s\n", onduty);
    }

    return msg;
}

function send_day_duty_notification(token) {
    var values = _CalendarRange.getValues();
    // Logger.log(JSON.stringify(values));
    // get row length in values
    // find next day in row of values
    var rowLength = values.length;
    var index = find_nextday_index(values);
    var msg = "\n\n ***** 泳訓通知 ***** \n";
    msg += format_on_duty_message(values[index]) + "\n";
    msg += "謝謝值班家長";

    // Logger.log(msg);
    if (msg != "") {
        sendLineNotification(token, msg);
    }
}

function send_week_duty_notification(token) {
    var values = _CalendarRange.getValues();
    // Logger.log(JSON.stringify(values));
    // get row length in values
    // skip first row, loop 7 days

    // 1. find next day in row of values
    // 2. collect all onduty in next 7 days
    var index = find_nextday_index(values);
    var msg = "\n\n ***** 下週行事曆 ***** \n\n";
    for (var i = index; i < index + 7; i++) {
        msg += format_on_duty_message(values[i]) + "\n";
    }

    // Logger.log(msg);
    if (msg != "") {
        sendLineNotification(token, msg);
    }
}

function find_nextday_index(values) {
    var index = 0;
    var rowLength = values.length;
    for (var i = 1; i < rowLength; i++) {
        const date_of_nextday = Utilities.formatDate(_nextday, "GMT+8", "MM/dd");
        const row_of_nextday = Utilities.formatDate(values[i][0], "GMT+8", "MM/dd");
        if (date_of_nextday == row_of_nextday) {
            index = i;
            break;
        }
    }
    return index;
}
