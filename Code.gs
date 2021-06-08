// Set some global parameters.
VERSION_NUMBER = '1.0 beta'
START_TIME_COLUMN = 1;
END_TIME_COLUMN = 3;
EVENT_ID_COLUMN = 4;
EVENT_FIRST_ROW = 5;
EVENT_NAME_COLUMN = 5;
DESCRIPTION_COLUMN_FIRST = 6;
DESCRIPTION_COLUMN_LAST = 7;
DESCRIPTION_DELIMINATOR = "\r\n---\r\n";
USE_DESCRIPTION_HEADERS = true;
CALENDAR_ID_CELL = "B2";
CALENDAR_NAME_CELL = "B1";

// Get the current sheet. It will be needed.
SHEET = SpreadsheetApp.getActiveSheet();

// Adds a menu when the spreadsheet is opened.
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('Kursplanering Go')
      .addItem('Skapa/uppdatera valda kalenderhändelser', 'event_update')
      .addItem('Radera valda kalenderhändelser', 'event_delete')
      .addSeparator()
      .addItem('Radera hela kalendern', 'calendar_delete')
      .addSeparator()
      .addItem('Om Kursplanering Go ' + VERSION_NUMBER, 'about')
      .addToUi();
}

// Updates or creates calendar events for the selected rows.
// (Creates a calendar if necessary.)
function event_update() {
  // Abort if the selected range is invalid.
  if (!event_range_validate()) {
    return false;
  }
  // Get the data for the selected rows.
  var start_row = SpreadsheetApp.getActiveRange().getRow();
  var row_span = SpreadsheetApp.getActiveRange().getNumRows();
  var last_column = Math.max(START_TIME_COLUMN, END_TIME_COLUMN, EVENT_ID_COLUMN, EVENT_NAME_COLUMN, DESCRIPTION_COLUMN_LAST);
  var data = SHEET.getRange(start_row, 1, row_span, last_column).getValues();
  // Create or update events based on the selected rows.
  var cal = calendar_get();
  for (e in data) {
    var id = data[e][EVENT_ID_COLUMN - 1];
    if (id) {
      var event = cal.getEventById(id);
      event.setTime(data[e][START_TIME_COLUMN - 1], data[e][END_TIME_COLUMN - 1]);
      event.setTitle(data[e][EVENT_NAME_COLUMN - 1]);
      event.setDescription(event_build_description(data[e]));
    }
    else {
      var event = cal.createEvent(
        data[e][EVENT_NAME_COLUMN - 1],
        new Date(data[e][START_TIME_COLUMN - 1]),
        new Date(data[e][END_TIME_COLUMN - 1]),
        {description: event_build_description(data[e])}
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
  // Get the data for the selected rows.
  var start_row = SpreadsheetApp.getActiveRange().getRow();
  var row_span = SpreadsheetApp.getActiveRange().getNumRows();
  var last_column = Math.max(START_TIME_COLUMN, END_TIME_COLUMN, EVENT_ID_COLUMN, EVENT_NAME_COLUMN, DESCRIPTION_COLUMN_LAST);
  var data = SHEET.getRange(start_row, 1, row_span, last_column).getValues();
  // Delete all events that have IDs.
  var cal = calendar_get();
  for (e in data) {
    var id = data[e][EVENT_ID_COLUMN - 1];
    if (id) {
      var event = cal.getEventById(id);
      event.deleteEvent();
      SHEET.getRange(start_row + parseInt(e), EVENT_ID_COLUMN).clear();
    }
  }
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
    if (parseInt(c) + 1 >= DESCRIPTION_COLUMN_FIRST && parseInt(c) + 1 <= DESCRIPTION_COLUMN_LAST) {
      // Add the header if the option for this is enabled.
      if (USE_DESCRIPTION_HEADERS) {
        data_row[c] = SHEET.getRange(EVENT_FIRST_ROW - 1, parseInt(c) + 1).getValue() + "\r\n" + data_row[c];
      }
      description.push(data_row[c]);
    }
  }
  return description.join(DESCRIPTION_DELIMINATOR);
}

// Helper function for loading the relevant calendar. Creates one if none exists.
function calendar_get(skip_create) {
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
  }
  return cal;
}

// Deletes the calendar for the sheet (along with any events).
function calendar_delete() {
  var cal = calendar_get(true);
  if (cal == false) {
    alert('Det finns ingen kalender att radera.');
    return;
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
  alert("För mer information, se https://github.com/Itangalo/kursplaneringGo.");
}

// Helper function for displaying messages to the user.
function alert(message) {
  SpreadsheetApp.getUi().alert(message);
}
