//
// This is a Google app script for Google Spreadsheet
//

// const Start = new Date(nextday.getFullYear(), nextday.getMonth(), nextday.getDate(), 0, 0, 0, 0);
// const End = new Date(nextday.getFullYear(), nextday.getMonth(), nextday.getDate(), 23, 59, 59, 999);

function map_calendar_values(columnName, row) {
    // map column name and row value
    var valueMap = {};
    var columnLength = columnName.length;
    for (var col = 0; col < columnLength; col++) {
        valueMap[columnName[col]] = row[col];
    }
    return valueMap;
}

function update_calendar_by_sheet(calendar) {
    var values = _rangeCALENDAR.getValues();
    // Logger.log(JSON.stringify(values));

    // get row length in values
    var rowLength = values.length;

    // Update calendar event by each row
    var columnName = values[0];
    for (var i = 1; i < rowLength; i++) {
        var row = values[i];

        var valueMap = map_calendar_values(columnName, row);
        update_calendar_on_duty_event(calendar, valueMap);
    }
}

function update_calendar_on_duty_event(calendar, values) {
    Logger.log('Google Calendar is "%s".', calendar.getName());

    var date = values["日期"];
    var time = values["時間"];
    var timeend = values["結束"];
    var note = values["備註"];
    var location = values["地點"];
    var onduty = Utilities.formatString("@%s @%s", values["值1"], values["值2"]);
    var title = location + " " + onduty;
    var email1 = values["Email1"];
    var email2 = values["Email2"];

    email1 = email1 ? email1 : "";
    email2 = email2 ? email2 : "";

    // search event tagged by 'onDuty'
    // if event tagged by 'onDuty' = "Yes" exist, update event info
    const search_date = new Date(date);
    const found_events = calendar.getEventsForDay(search_date);
    var onduty_event = found_events.find(e => e.getTag("onDuty") == "Yes");

    // update onduty event
    var onduty_date = Utilities.formatDate(date, "GMT+8", "yyyy/MM/dd ");
    var onduty_datetime = onduty_date + Utilities.formatDate(new Date(time), "GMT+8", "HH:mm");
    var endduty_datetime = onduty_datetime;
    if (timeend != "") endduty_datetime = onduty_date + Utilities.formatDate(new Date(timeend), "GMT+8", "HH:mm");

    Logger.log("%s, %s, %s, %s, %s, %s", onduty_datetime, endduty_datetime, title, location, note, onduty);
    var onduty_event_time = new Date(onduty_datetime);
    var endduty_event_time = new Date(endduty_datetime);
    if (onduty_event == null) {
        // create new google caldendar event
        // set title, location and time
        onduty_event = calendar.createEvent(title, onduty_event_time, onduty_event_time, {
            description: note,
            location: location,
        });
    }
    onduty_event.setTag("onDuty", "Yes");
    onduty_event.setLocation(location);
    onduty_event.setDescription(note);
    if (time == "") {
        onduty_event.setTitle(note);
        onduty_event.setAllDayDate(onduty_event_time);
    } else {
        onduty_event.setTitle(title);
        onduty_event.setTime(onduty_event_time, endduty_event_time);
    }

    // remove guests if guest is not email1 nor email2 from guest list
    var remove_guests = onduty_event.getGuestList();
    remove_guests.forEach(g => {
        if (g.getEmail() != email1 && g.getEmail() != email2) {
            onduty_event.removeGuest(g.getEmail());
        }
    });

    // check email1 and email2 is in guest list or not
    // if not exist, add them to guest list
    var guest_list = onduty_event.getGuestList();
    var guest1 = guest_list.find(g => g.getEmail() == email1);
    var guest2 = guest_list.find(g => g.getEmail() == email2);
    if (guest1 == null && email1 != "") onduty_event.addGuest(email1);
    if (guest2 == null && email2 != "") onduty_event.addGuest(email2);
}
