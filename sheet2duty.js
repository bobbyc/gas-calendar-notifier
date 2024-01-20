//
// This is a Google app script for Google Spreadsheet
//

function format_on_duty_message(values) {
    var datetime = Utilities.formatDate(values["日期"], "GMT+8", "MM/dd (E) ");
    var time = values["時間"];
    var timeend = values["結束"];
    var note = values["備註"];
    var location = values["地點"];
    var duty1 = values["值1"];
    var duty2 = values["值2"];
    var onduty = Utilities.formatString("@%s @%s", duty1, duty2);

    if (time != "")
        datetime += Utilities.formatDate(new Date(time), "GMT+8", "HH:mm");

    if (timeend != "")
        datetime += " - " + Utilities.formatDate(new Date(timeend), "GMT+8", "HH:mm");

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
    var values = _rangeCALENDAR.getValues();
    // Logger.log(JSON.stringify(values));
    // get row length in values
    // find next day in row of values
    var index = find_nextday_index(values);
    var valueMap = map_calendar_values(values[0], values[index]);
    var msg_template =
    '\n\n' +
    ' ***** 泳訓通知 ***** \n' +
    '$on_duty_message\n' +
    '謝謝值班家長\n' +
    '\n' +
    '$extra_message\n' +
    '\n' +
    '$weather_message';

    var duty_msg = format_on_duty_message(valueMap);
    var weather_msg = weather_description(valueMap);

    // if dayofweek is weekday, add extra message
    var dayofweek = valueMap["週"];
    var extra_msg = "";
    if (dayofweek == "二" ||
        dayofweek == "三" ||
        dayofweek == "四") {
        extra_msg += "\n\n ***** ！注意！ ***** \n";
        extra_msg += "低年級同學請 07:50 入班晨光時間！\n";
        extra_msg += "中高年級隊員 08:40 進教室，準備上課！\n";
    } else if (dayofweek == "一" || dayofweek == "五") {
        extra_msg += "\n\n ***** ！注意！ ***** \n";
        extra_msg += "全體隊員請於 07:50 前進教室！\n";
    } else {
    }

    // Format message
    msg_template = msg_template.replace('$on_duty_message', duty_msg);
    msg_template = msg_template.replace('$extra_message', extra_msg);
    msg_template = msg_template.replace('$weather_message', weather_msg);

    // Logger.log(msg);
    if (msg_template != "") {
        sendLineNotification(token, msg_template);
    }
}

function send_week_duty_notification(token) {
    var values = _rangeCALENDAR.getValues();
    // Logger.log(JSON.stringify(values));
    // get row length in values
    // skip first row, loop 7 days

    // 1. find next day in row of values
    // 2. collect all onduty in next 7 days
    var index = find_nextday_index(values);
    var msg = "\n\n ***** 下週行事曆 ***** \n\n";
    for (var i = index; i < index + 7; i++) {
        var valueMap = map_calendar_values(values[0], values[i]);
        msg += format_on_duty_message(valueMap) + "\n";
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
