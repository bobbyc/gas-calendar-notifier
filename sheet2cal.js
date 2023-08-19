//
// This is a Google app script for Google Spreadsheet
//

function update_calendar_on_duty_event(calendar, values) {
  var date = Utilities.formatDate(values[0], "GMT+8", "yyyy/MM/dd ");
  date += Utilities.formatDate(new Date(values[2]), "GMT+8", "HH:mm");
  Logger.log(date);

  var note = values[3];
  var location = values[4];
  var onduty = "@" + values[5] + " @" + values[6];
  var title = location + " " + onduty;

  // search event tagged by 'onDuty'
  // if event tagged by 'onDuty' = "Yes" exist, update event info
  const search_date = new Date(values[0]);
  const events = calendar.getEventsForDay(search_date);
  var event = events.find(event => event.getTag("onDuty") == "Yes");
  var event_date = new Date(date);
  if (event) {
    event.setTime(event_date, event_date);
  } else {
    // create new google caldendar event
    // set title, location and time
    event = calendar.createEvent(title, event_date, event_date, {
      description: note,
      location: location,
    });
  }
  event.setTitle(title);
  event.setTag("onDuty", "Yes");
  event.setLocation(location);
  event.setDescription(note);
}

function format_on_duty_message(values) {
  var date = Utilities.formatDate(values[0], "GMT+8", "MM/dd (E) ");
  date += Utilities.formatDate(new Date(values[2]), "GMT+8", "HH:mm");
  var note = values[3];
  var location = values[4];
  var onduty = "@" + values[5] + " @" + values[6];

  // return formatted message string with date, time, location, and onduty
  var msg = "日期: " + date;
  if (note != "") msg += " (" + note + ")";
  msg += "\n";
  msg += "地點: " + location + "\n";
  msg += "值班家長: " + onduty + "\n";
  return msg;
}
