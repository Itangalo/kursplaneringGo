// Set some global parameters.
VERSION_NUMBER = '1.6 beta EXPERIMENTAL'
START_TIME_COLUMN = 1;
END_TIME_COLUMN = 3;
EVENT_ID_COLUMN = 4;
EVENT_FIRST_ROW = 5;
EVENT_NAME_COLUMN = 5;
EVENT_LOCATION_COLUMN = 6;
ANNOUNCEMENT_TEXT_COLUMN = 13;
ANNOUNCEMENT_TIME_COLUMN = 12;
DESCRIPTION_COLUMN_FIRST_CELL = "D1";
DESCRIPTION_COLUMN_LAST_CELL = "D2";
DESCRIPTION_DELIMINATOR = "\n---\n";
USE_DESCRIPTION_HEADERS_CELL = "F1";
COURSE_ID_CELL = "L1";
CALENDAR_ID_CELL = "B2";
CALENDAR_NAME_CELL = "B1";

// Get the current sheet and load some values. It will be needed.
SHEET = SpreadsheetApp.getActiveSheet();
DESCRIPTION_COLUMN_FIRST = SHEET.getRange(DESCRIPTION_COLUMN_FIRST_CELL).getValue();
DESCRIPTION_COLUMN_LAST = SHEET.getRange(DESCRIPTION_COLUMN_LAST_CELL).getValue();
USE_DESCRIPTION_HEADERS = SHEET.getRange(USE_DESCRIPTION_HEADERS_CELL).getValue();
LAST_COLUMN = Math.max(START_TIME_COLUMN, END_TIME_COLUMN, EVENT_ID_COLUMN, EVENT_NAME_COLUMN, EVENT_LOCATION_COLUMN, DESCRIPTION_COLUMN_LAST, ANNOUNCEMENT_TEXT_COLUMN, ANNOUNCEMENT_TIME_COLUMN);
COURSE_ID = SHEET.getRange(COURSE_ID_CELL).getValue();

// Adds a menu when the spreadsheet is opened.
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('Kursplanering Go')
      .addItem('Skapa/uppdatera valda kalenderhändelser', 'event_update')
      .addItem('Radera valda kalenderhändelser', 'event_delete')
      .addSeparator()
      .addItem('Läs in kalenderhändelser', 'calendar_read')
      .addSeparator()
      .addItem('Radera hela kalendern', 'calendar_delete')
      .addSeparator()
      .addItem('Om Kursplanering Go ' + VERSION_NUMBER, 'about')
      .addSeparator()
      .addItem('Schemalägg Classroom-meddelanden (EXPERIMENTAL)', 'create_announcements')
      .addItem('Lista ID:n för Classroom-klasser (EXPERIMENTAL)', 'find_course_id')
      .addToUi();
}

// Updates or creates calendar events for the selected rows.
// (Creates a calendar if necessary.)
function event_update() {
  // Abort if the selected range is invalid.
  if (!event_range_validate()) {
    return false;
  }
  // Verify that the user may create/update events.
  var cal = calendar_get();
  if (!cal.isOwnedByMe() && !calendar_editable(cal)) {
    throw 'You do not have permission to update events in this calendar.';
  }
  // Get the data for the selected rows.
  var start_row = SpreadsheetApp.getActiveRange().getRow();
  var row_span = SpreadsheetApp.getActiveRange().getNumRows();
  var data = SHEET.getRange(start_row, 1, row_span, LAST_COLUMN).getValues();
  // Create or update events based on the selected rows.
  for (e in data) {
    var id = data[e][EVENT_ID_COLUMN - 1];
    if (id) {
      var event = cal.getEventById(id);
      event.setTime(data[e][START_TIME_COLUMN - 1], data[e][END_TIME_COLUMN - 1]);
      event.setTitle(data[e][EVENT_NAME_COLUMN - 1]);
      event.setDescription(event_build_description(data[e]));
      event.setLocation(data[e][EVENT_LOCATION_COLUMN - 1]);
    }
    else {
      var event = cal.createEvent(
        data[e][EVENT_NAME_COLUMN - 1],
        new Date(data[e][START_TIME_COLUMN - 1]),
        new Date(data[e][END_TIME_COLUMN - 1]),
        {
          description: event_build_description(data[e]),
          location: data[e][EVENT_LOCATION_COLUMN - 1]
        }
      );
      SHEET.getRange(start_row + parseInt(e), EVENT_ID_COLUMN).setValue(event.getId());
    }
  }
}

// Deletes calendar events for the selected rows.
function event_delete() {
  // Abort if the selected range is invalid.
  if (!event_range_validate()) {
    return false;
  }
  // Verify that the user may create events.
  var cal = calendar_get();
  if (!cal.isOwnedByMe() && !calendar_editable(cal)) {
    throw 'You do not have permission to delete events in this calendar.';
  }
  // Get the data for the selected rows.
  var start_row = SpreadsheetApp.getActiveRange().getRow();
  var row_span = SpreadsheetApp.getActiveRange().getNumRows();
  var data = SHEET.getRange(start_row, 1, row_span, LAST_COLUMN).getValues();
  // Delete all events that have IDs.
  for (e in data) {
    var id = data[e][EVENT_ID_COLUMN - 1];
    if (id) {
      var event = cal.getEventById(id);
      event.deleteEvent();
      SHEET.getRange(start_row + parseInt(e), EVENT_ID_COLUMN).clear();
    }
  }
}

// Reads calendar events from -1 to +1 year and populates the spreadsheet.
function calendar_read() {
  var cal = calendar_get();
  // Get all event IDs currently listed.
  var event_ids = SHEET.getRange(EVENT_FIRST_ROW, EVENT_ID_COLUMN, SHEET.getLastRow() - EVENT_FIRST_ROW + 1, 1).getValues();
  event_ids = event_ids.map(i => i[0]);
  // Get all events from -1 to +1 year from now.
  var start_date = new Date();
  start_date.setFullYear(start_date.getFullYear() - 1);
  var end_date = new Date();
  end_date.setFullYear(end_date.getFullYear() + 1);
  var events = cal.getEvents(start_date, end_date);

  // Create or update events, as appropriate.
  for (e in events) {
    // Read current data from sheet, if the event ID exists. Otherwise pick new line.
    var row = event_ids.indexOf(events[e].getId());
    if (row < 0) {
      row = 1 + SHEET.getLastRow();
    }
    else {
      row = row + EVENT_FIRST_ROW;
    }
    var selection = SHEET.getRange(row, 1, 1, LAST_COLUMN);
    var data = selection.getValues()[0];

    // Populate the sheet with data from the event.
    data[START_TIME_COLUMN - 1] = events[e].getStartTime();
    data[END_TIME_COLUMN - 1] = events[e].getEndTime();
    data[EVENT_NAME_COLUMN - 1] = events[e].getTitle();
    data[EVENT_ID_COLUMN - 1] = events[e].getId();
    data[EVENT_LOCATION_COLUMN - 1] = events[e].getLocation();
    var description = event_parse_description(events[e].getDescription());
    for (i in description) {
      if (!description[i]) {
        description[i] = '';
      }
      data[DESCRIPTION_COLUMN_FIRST - 1 + parseInt(i)] = description[i];
    }
    selection.setValues([data]);
  }
}

// Helper function getting data from the selected rows.
function get_selected_row_data() {
  var start_row = SpreadsheetApp.getActiveRange().getRow();
  var row_span = SpreadsheetApp.getActiveRange().getNumRows();
  return SHEET.getRange(start_row, 1, row_span, LAST_COLUMN).getValues();
}

// Verifies that the selected range only includes event rows.
function event_range_validate() {
  if (SpreadsheetApp.getActiveRange().getRow() < EVENT_FIRST_ROW || SpreadsheetApp.getActiveRange().getLastRow() > SHEET.getLastRow()) {
    alert("Markera (endast) rader med kalenderhändelser innan du utför den här åtgärden.")
    return false;
  }
  else return true;
}

// Helper function for merging content from description cells.
function event_build_description(data_row) {
  var description = [];
  for (c in data_row) {
    if (parseInt(c) + 1 >= DESCRIPTION_COLUMN_FIRST && parseInt(c) + 1 <= DESCRIPTION_COLUMN_LAST && data_row[c].length > 0) {
      // Add the header if the option for this is enabled.
      if (USE_DESCRIPTION_HEADERS) {
        data_row[c] = SHEET.getRange(EVENT_FIRST_ROW - 1, parseInt(c) + 1).getValue() + "\r\n" + data_row[c];
      }
      description.push(data_row[c]);
    }
  }
  return description.join(DESCRIPTION_DELIMINATOR);
}

// Attempts to build a description in separate cells from a plain-text event description.
// Returns an array of the same size as the number of description headers.
function event_parse_description(description_string) {
  var parts = description_string.split(DESCRIPTION_DELIMINATOR);
  var num_headers = DESCRIPTION_COLUMN_LAST - DESCRIPTION_COLUMN_FIRST + 1;
  var headers = SHEET.getRange(EVENT_FIRST_ROW - 1, DESCRIPTION_COLUMN_FIRST, 1, num_headers).getValues()[0];
  var output = [];
  output[num_headers - 1] = null; // Ensure that the output array has sufficient length.
  // First try to find description parts starting with the headers in the spreadsheet.
  var match = false;
  for (var c in headers) {
    for (var p in parts) {
      // If the start of the descrption part matches the header, add it to the output and remove that line.
      if (parts[p].substr(0, headers[c].length) == headers[c]) {
        output[c] = parts[p].substr(headers[c].length).trim();
        parts.splice(p, 1);
        match = true;
      }
    }
  }
  // If there is no header match, assume that the description parts map trivialy to the description columns.
  if (!match) {
    for (var c in headers) {
      output[c] = parts.shift();
    }
  }
  // If there are any description parts left, append them to the last description column.
  if (parts.length) {
    output[num_headers - 1] = output[num_headers - 1] + DESCRIPTION_DELIMINATOR + parts.join(DESCRIPTION_DELIMINATOR);
  }
  return output;
}

// Helper function for loading the relevant calendar. Creates one if none exists.
function calendar_get(skip_create) {
  // Create a new calendar if no calendar ID exists. Unless it should be skipped.
  if (SHEET.getRange(CALENDAR_ID_CELL).isBlank()) {
    if (skip_create) {
      return false;
    }
    var cal = CalendarApp.createCalendar(SHEET.getRange(CALENDAR_NAME_CELL).getValue());
    SHEET.getRange(CALENDAR_ID_CELL).setValue(cal.getId());
    alert('Ny kalender skapad och tillagd i Google Calendar.');
  }
  else {
    var cal = CalendarApp.getCalendarById(SHEET.getRange(CALENDAR_ID_CELL).getValue());
    if (cal == null) {
      throw 'Invalid calendar ID.';
    }
  }
  return cal;
}

// Returns true if the user may edit the calendar, otherwise false.
function calendar_editable(cal) {
  try {
    var e = cal.createEvent('Temporary event',
    new Date(),
    new Date());
  }
  catch(error) {
    return false;
  }
  e.deleteEvent();
  return true;
}

// Deletes the calendar for the sheet (along with any events).
function calendar_delete() {
  var cal = calendar_get(true);
  if (cal == false) {
    alert('Det finns ingen kalender att radera.');
    return;
  }
  if (!cal.isOwnedByMe()) {
    throw 'You do not have permission to delete the calendar.';
  }

  cal.deleteCalendar();
  // Clear calendar ID and all event IDs.
  SHEET.getRange(CALENDAR_ID_CELL).clear();
  SHEET.getRange(EVENT_FIRST_ROW, EVENT_ID_COLUMN, SHEET.getLastRow() - EVENT_FIRST_ROW + 1, 1).clear();
  alert('Kalender raderad.');
  return;
}

// Displays a popup for finding help and more.
function about() {
  var htmlOutput = HtmlService
    .createHtmlOutput("<p>För mer information, se <a href='https://github.com/Itangalo/kursplaneringGo' target='_blank'>github.com/Itangalo/kursplaneringGo</a>.</p>")
    .setWidth(350)
    .setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Om Kursplaning Go ' + VERSION_NUMBER);
}

// Helper function for displaying messages to the user.
function alert(message) {
  SpreadsheetApp.getUi().alert(message);
}

// Schedules announcements in Google Classroom, for selected rows.
function create_announcements() {
  if (!event_range_validate()) {
    return false;
  }
  var data = get_selected_row_data();
  for (d in data) {
    announcement = {
      'text': data[d][ANNOUNCEMENT_TEXT_COLUMN - 1],
    };

    // Scheduling can only be done for future times.
    announce_time = new Date(data[d][ANNOUNCEMENT_TIME_COLUMN - 1]);
    var now = new Date();
    if (announce_time > now) {
      announcement.state = 'DRAFT';
      announcement.scheduledTime = announce_time.toISOString();
    }
    Classroom.Courses.Announcements.create(announcement, COURSE_ID);
  }
}

// Displayes a pop-up with all course names and IDs.
function find_course_id() {
  var courses = Classroom.Courses.list().courses;
  var output = [];
  if (courses && courses.length > 0) {
    for (c in courses) {
      output.push(courses[c].name + ': ' + courses[c].id);
    }
  } else {
    output.push('No courses found.');
  }
  alert(output.join("\n"));
}
