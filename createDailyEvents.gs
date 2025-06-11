// A = 1, B = 2, C = 3, D = 4, E = 5, F = 6, G = 7,
// H = 8, I = 9, J = 10, K = 11, L = 12, M = 13, N = 14, 
// O = 15, P = 16, Q = 17, R = 18, S = 19, T = 20, 
// U = 21, V = 22, W = 23, X = 24, Y = 25, and Z = 26

/**
 * CREATE EVENT CALENDAR - DAILY EVENT
 */

function createCalendarEvents() {
  // const sheetSchedule = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Meeting Schedule");
  const sheetWrite = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(" ✒️WRITE HERE");
  const calendar = CalendarApp.getCalendarById("acb8821a06898e03212fdb61694c40e9b5b82d5f390ae15632d808c04507cfaa@group.calendar.google.com");
  const rows = sheetWrite.getDataRange().getValues();


  for (let i = 1; i < rows.length; i++) {

    //Array for all columns in the sheet
    const [title, location, action, lastAction, meetingStartDate, startTime, meetingEndDate, endTime, contactEmail, phoneNumber, livestream, tableArrangement, foodArrangements, amountPaid, paymentStatus, numberOfPeople, groupType, description, timestamp, eventId, start, end, titleLocation, modifiedBy, createdBy] = rows[i];

    // Check if the row has been processed already
    if (rows[i][2] == "Create") {
      // Check if the cell value is "Create"

      //if (!endTime) {
      // UI Alert for event created
      //SpreadsheetApp.getUi().alert(`END TIME MISSING : 
      //Please enter event END TIME to complete.`)

      // } else {

      const event = calendar.createEvent(titleLocation, new Date(start), new Date(end), {
        description: description,
        location: location
      });

      // Mark the row as "Added"
      sheetWrite.getRange(i + 1, 4).setValue("Created");  // Set the status cell to "Created"
      sheetWrite.getRange(i + 1, 3).setValue("Select Action");  // Set the cell back to "Select Action"
      sheetWrite.getRange(i + 1, 20).setValue(event.getId()); // Set Event ID
      sheetWrite.getRange(i + 1, 19).setValue(new Date()); // Set timestamp
      sheetWrite.getRange(i + 1, 24).setValue(Session.getActiveUser().getEmail()); // Set Modified By Email
      sheetWrite.getRange(i + 1, 25).setValue(Session.getActiveUser().getEmail()); // Set Created By Email

      // UI Alert for event created
      SpreadsheetApp.getUi().alert(`EVENT CREATED SUCCESSFULLY
       ${description}
       `)

    } if (rows[i][2] == 'Update' && eventId) {

      const event = calendar.getEventById(eventId);

      if (event) {
        event.setTitle(titleLocation);
        event.setDescription(description);
        event.setLocation(location);
        event.setTime(new Date(start), new Date(end));

        sheetWrite.getRange(i + 1, 4).setValue("Updated");
        sheetWrite.getRange(i + 1, 3).setValue("Select Action");  // Column F - Update action on WRITE Sheet
        sheetWrite.getRange(i + 1, 19).setValue(new Date()); // Set timestamp
        sheetWrite.getRange(i + 1, 24).setValue(Session.getActiveUser().getEmail()); // Set Modified By Email

      }
      //UI Alert for event updated
      SpreadsheetApp.getUi().alert(`EVENT UPDATED SUCCESSFULLY
        ${description} `);

      //}
    } else if (rows[i][2] == 'Delete' && eventId) {

      const event = calendar.getEventById(eventId);

      if (event) {
        event.deleteEvent();

        sheetWrite.getRange(i + 1, 4).setValue("Deleted");
        sheetWrite.getRange(i + 1, 3).setValue("Select Action");  // Column F - Update action on WRITE Sheet
        sheetWrite.getRange(i + 1, 19).setValue(new Date()); // Set timestamp
        sheetWrite.getRange(i + 1, 24).setValue(Session.getActiveUser().getEmail()); // Set Modified By Email

      }
      //UI Alert for event deleted
      SpreadsheetApp.getUi().alert(`EVENT DELETED SUCCESSFULLY
      ${description}`);

    }
  }
}
