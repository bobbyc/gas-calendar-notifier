//
// This is a Google app script for Google Spreadsheet
//

// const Start = new Date(nextday.getFullYear(), nextday.getMonth(), nextday.getDate(), 0, 0, 0, 0);
// const End = new Date(nextday.getFullYear(), nextday.getMonth(), nextday.getDate(), 23, 59, 59, 999);

function update_calendar_by_sheet(calendar) {
    var values = _CalendarRange.getValues();
    // Logger.log(JSON.stringify(values));

    // get row length in values
    var rowLength = values.length;
    for (var i = 1; i < rowLength; i++) {
        var row = values[i];
        update_calendar_on_duty_event(calendar, row);
    }
}

function update_calendar_on_duty_event(calendar, values) {
    Logger.log('Google Calendar is "%s".', calendar.getName());

    var date = values[0];
    var time = values[2];
    var note = values[3];
    var location = values[4];
    var onduty = Utilities.formatString("@%s @%s", values[5], values[6]);
    var title = location + " " + onduty;
    var email1 = values[7] ? values[7] : "";
    var email2 = values[8] ? values[8] : "";

    // search event tagged by 'onDuty'
    // if event tagged by 'onDuty' = "Yes" exist, update event info
    const search_date = new Date(date);
    const found_events = calendar.getEventsForDay(search_date);
    var onduty_event = found_events.find(e => e.getTag("onDuty") == "Yes");

    // update onduty event
    var onduty_datetime = Utilities.formatDate(date, "GMT+8", "yyyy/MM/dd ");
    onduty_datetime += Utilities.formatDate(new Date(time), "GMT+8", "HH:mm");
    Logger.log("%s, %s, %s, %s, %s", onduty_datetime, title, location, note, onduty);
    var onduty_event_time = new Date(onduty_datetime);
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
        onduty_event.setTime(onduty_event_time, onduty_event_time);
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
