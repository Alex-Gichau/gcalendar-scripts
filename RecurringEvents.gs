// A = 1, B = 2, C = 3, D = 4, E = 5, F = 6, G = 7,
// H = 8, I = 9, J = 10, K = 11, L = 12, M = 13, N = 14, 
// O = 15, P = 16, Q = 17, R = 18, S = 19, T = 20, 
// U = 21, V = 22, W = 23, X = 24, Y = 25, and Z = 26

//Recurring Events function

/**
 * WEEKLY RECURRENCE
 */

function recurringEvents() {
  const sheetRecurring = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(" ✒️WRITE HERE ( RECURRING )");
  if (!sheetRecurring) {
    SpreadApp.getUi().alert('Sheet named "-" not found.');
    return;
  }
  const calendar = CalendarApp.getCalendarById("acb8821a06898e03212fdb61694c40e9b5b82d5f390ae15632d808c04507cfaa@group.calendar.google.com");
  const dataRange = sheetRecurring.getDataRange();
  const rows = dataRange.getValues();

  // Start from row 1 to skip the header
  for (var i = 1; i < rows.length; i++) {
    try {
      // A = 1, B = 2, C = 3, D = 4, E = 5, F = 6, G = 7,
      // H = 8, I = 9, J = 10, K = 11, L = 12, M = 13, N = 14, 
      // O = 15, P = 16, Q = 17, R = 18, S = 19, T = 20, 
      // U = 21, V = 22, W = 23, X = 24, Y = 25, and Z = 26

      const
        [title,                       // A
          meetingRoom,                // B
          action,                     // C
          lastAction,                 // D
          meetingStartDate,           // E
          startTime,                  // F
          meetingEndDate,             // G
          endTime,                    // H
          recurring,                  // I
          recurringUntil,             // J
          livestream,                 // K
          tableArrangement,           // L
          foodArrangements,           // M
          amountPaid,                 // N
          paymentStatus,              // O
          numberOfPeople,             // P
          description,                // Q
          timestamp,                  // R
          eventId,                    // S
          start,                      // T
          end,                        // U
          titleLocation,              // V
          modifiedBy,                 // W
          createdBy,                  // X
          recurrenceMode,             // Y
          weekNumber                  // Z
          ]
          = rows[i];

      // Process only if Action is "Create", recurrence is "Weekly", and it hasn't been created yet (eventId is empty)
      //=======================//
      //  CREATE OPERATION   //
      //=======================//
      if (action === "Create" && recurring === "Weekly" && !eventId) {

        if (!(meetingStartDate instanceof Date) || !(startTime instanceof Date) || !(endTime instanceof Date) || !(recurringUntil instanceof Date)) {
          console.log(`Skipping row ${i + 1} due to invalid date/time format.`);
          // UI Alert for the created event
          SpreadsheetApp.getUi().alert(`Missing start date on the events: 
        ${titleLocation}
        `);
          continue; // Skip to the next row if dates/times are not valid
        }

        // Create the recurrence rule
        const recurrence = CalendarApp.newRecurrence()
          .addWeeklyRule().until(new Date(recurringUntil));

        // Set event options
        const eventOptions = {
          description: description,
          location: meetingRoom
        };

        // Create the event series
        const newEventSeries = calendar.createEventSeries(titleLocation, new Date(start), new Date(end), recurrence, eventOptions);
        const newEventId = newEventSeries.getId();

        // Log and save the new event ID back to the sheet in the 'eventId' column (column 19)
        console.log(`Created event series with ID: ${newEventId} for row ${i + 1}`);
        sheetRecurring.getRange(i + 1, 19).setValue(newEventId); // Column S is the 19th column
        sheetRecurring.getRange(i + 1, 23).setValue(Session.getActiveUser().getEmail()); // Set Modified By Email
        sheetRecurring.getRange(i + 1, 24).setValue(Session.getActiveUser().getEmail()); // Set Created By Email

        // Optional: Update the 'Action' status to "Created"
        sheetRecurring.getRange(i + 1, 4).setValue("Created"); // Column C is the 3rd column
        sheetRecurring.getRange(i + 1, 3).setValue("Select Action"); // Column C is the 3rd column

        // UI Alert for the created event
        SpreadsheetApp.getUi().alert(`RECURRING EVENT CREATED: 
        ${description}
        `);
      } else if (action === "Create" && recurring === "Daily" && !eventId) {

        if (!(meetingStartDate instanceof Date) || !(startTime instanceof Date) || !(endTime instanceof Date) || !(recurringUntil instanceof Date)) {
          console.log(`Skipping row ${i + 1} due to invalid date/time format.`);
          // UI Alert for the created event
          SpreadsheetApp.getUi().alert(`Missing start date / end date / recurring until date on the events: 
        ${titleLocation}
        `);
          continue; // Skip to the next row if dates/times are not valid
        }

        // Create the recurrence rule
        const recurrence = CalendarApp.newRecurrence()
          .addDailyRule().until(new Date(recurringUntil));

        // Set event options
        const eventOptions = {
          description: description,
          location: meetingRoom
        };

        // Create the event series
        const newEventSeries = calendar.createEventSeries(titleLocation, new Date(start), new Date(end), recurrence, eventOptions);
        const newEventId = newEventSeries.getId();

        // Log and save the new event ID back to the sheet in the 'eventId' column (column 19)
        console.log(`Created event series with ID: ${newEventId} for row ${i + 1}`);
        sheetRecurring.getRange(i + 1, 19).setValue(newEventId); // Column S is the 19th column
        sheetRecurring.getRange(i + 1, 23).setValue(Session.getActiveUser().getEmail()); // Set Modified By Email
        sheetRecurring.getRange(i + 1, 24).setValue(Session.getActiveUser().getEmail()); // Set Created By Email

        // Optional: Update the 'Action' status to "Created"
        sheetRecurring.getRange(i + 1, 4).setValue("Created"); // Column C is the 3rd column
        sheetRecurring.getRange(i + 1, 3).setValue("Select Action"); // Column C is the 3rd column

        // UI Alert for the created event
        SpreadsheetApp.getUi().alert(`RECURRING EVENT CREATED: 
        ${description}
        `);

      } else if (action === "Create" && recurring === "Monthly" && !eventId) {

        if (!(meetingStartDate instanceof Date) || !(startTime instanceof Date) || !(endTime instanceof Date) || !(recurringUntil instanceof Date)) {
          console.log(`Skipping row ${i + 1} due to invalid date/time format.`);
          // UI Alert for the created event
          SpreadsheetApp.getUi().alert(`Missing start date / end date / recurring until date on the events: 
        ${titleLocation}
        `);
          continue; // Skip to the next row if dates/times are not valid
        }

        // Create the recurrence rule
        const recurrence = CalendarApp.newRecurrence()
          .addMonthlyRule().until(new Date(recurringUntil));

        // Set event options
        const eventOptions = {
          description: description,
          location: meetingRoom
        };

        // Create the event series
        const newEventSeries = calendar.createEventSeries(titleLocation, new Date(start), new Date(end), recurrence, eventOptions);
        const newEventId = newEventSeries.getId();

        // Log and save the new event ID back to the sheet in the 'eventId' column (column 19)
        console.log(`Created event series with ID: ${newEventId} for row ${i + 1}`);
        sheetRecurring.getRange(i + 1, 19).setValue(newEventId); // Column S is the 19th column
        sheetRecurring.getRange(i + 1, 23).setValue(Session.getActiveUser().getEmail()); // Set Modified By Email
        sheetRecurring.getRange(i + 1, 24).setValue(Session.getActiveUser().getEmail()); // Set Created By Email

        // Optional: Update the 'Action' status to "Created"
        sheetRecurring.getRange(i + 1, 4).setValue("Created"); // Column C is the 3rd column
        sheetRecurring.getRange(i + 1, 3).setValue("Select Action"); // Column C is the 3rd column

        // UI Alert for the created event
        SpreadsheetApp.getUi().alert(`RECURRING EVENT CREATED: 
        ${description}
        `);
      } else if (action === "Create" && recurring === "This Day Every Month" && !eventId) {

        if (!(meetingStartDate instanceof Date) || !(startTime instanceof Date) || !(endTime instanceof Date) || !(recurringUntil instanceof Date)) {
          console.log(`Skipping row ${i + 1} due to invalid date/time format.`);
          // UI Alert for the created event
          SpreadsheetApp.getUi().alert(`Missing start date / end date / recurring until date on the events: 
        ${titleLocation}
        `);
          continue; // Skip to the next row if dates/times are not valid
        }

        // Determine the day of the week selected and store as dynamic week
        const dynamicWeekDay = sheetRecurring.getRange(i + 1, 25).getValue();
        console.log("dynamicWeekDay", dynamicWeekDay);
        const weekNum = sheetRecurring.getRange(i + 1, 26).getValue();
        console.log("weekNum", weekNum);
       
        // Create the recurrence rule
        const recurrence = CalendarApp.newRecurrence()
          .addMonthlyRule().onlyOnWeekday(CalendarApp.Weekday.TUESDAY).interval(1).until(new Date(recurringUntil));

        // Set event options
        const eventOptions = {
          description: description,
          location: meetingRoom
        };

        // Create the event series
        const newEventSeries = calendar.createEventSeries(titleLocation, new Date(start), new Date(end), recurrence, eventOptions);
        const newEventId = newEventSeries.getId();

        // Log and save the new event ID back to the sheet in the 'eventId' column (column 19)
        console.log(`Created event series with ID: ${newEventId} for row ${i + 1}`);
        sheetRecurring.getRange(i + 1, 19).setValue(newEventId); // Column S is the 19th column
        sheetRecurring.getRange(i + 1, 23).setValue(Session.getActiveUser().getEmail()); // Set Modified By Email
        sheetRecurring.getRange(i + 1, 24).setValue(Session.getActiveUser().getEmail()); // Set Created By Email

        // Optional: Update the 'Action' status to "Created"
        sheetRecurring.getRange(i + 1, 4).setValue("Created"); // Column C is the 3rd column
        sheetRecurring.getRange(i + 1, 3).setValue("Select Action"); // Column C is the 3rd column

        // UI Alert for the created event
        SpreadsheetApp.getUi().alert(`RECURRING EVENT CREATED: 
        ${description}
        `);
      }

      //=======================//
      //  DELETE OPERATION   //
      //=======================//
      else if (action === "Delete" && eventId) {
        const event = calendar.getEventById(eventId);
        if (event) {
          event.deleteEvent(); // Delete the entire series
        } else {
          // If event is not found, it's already gone. Just clean up the sheet.
          SpreadsheetApp.getUi().alert(`Event titled ${titleLocation} not found on calendar for deletion. Already delete`);
          console.log(`Event titled ${titleLocation} not found on calendar for deletion. Already deleted.`);
        }

        // Update the sheet to show the operation is complete
        sheetRecurring.getRange(i + 1, 3).setValue("Select Action"); // Reset Action (Col C)
        sheetRecurring.getRange(i + 1, 4).setValue("Deleted");      // Set Last Action (Col D)
        sheetRecurring.getRange(i + 1, 19).clear();
        sheetRecurring.getRange(i + 1, 23).setValue(Session.getActiveUser().getEmail()); // Set Modified By Email

        SpreadsheetApp.getUi().alert(`EVENT DELETED: 
        ${description}`);
      }

      //=======================//
      //  UPDATE OPERATION   //
      //=======================//
      else if (action === "Update" && eventId) {
        const event = calendar.getEventById(eventId);
        if (!event) {
          SpreadsheetApp.getUi().alert(`Event not found on calendar. It may have been deleted manually.`);
          throw new Error("Event not found on calendar. It may have been deleted manually.");
        }

        const updateRecurrence = CalendarApp.newRecurrence().addWeeklyRule().until(new Date(recurringUntil));
        const eventOptions = { description: description, location: meetingRoom };

        const newEventSeries = calendar.createEventSeries(titleLocation, new Date(start), new Date(end), updateRecurrence, eventOptions);

        // Update the sheet with the new event's ID
        sheetRecurring.getRange(i + 1, 19).setValue(newEventSeries.getId()); // Set Event ID (Col S)
        sheetRecurring.getRange(i + 1, 3).setValue("Select Action");         // Reset Action (Col C)
        sheetRecurring.getRange(i + 1, 4).setValue("Updated");              // Set Last Action (Col D)
        sheetRecurring.getRange(i + 1, 23).setValue(Session.getActiveUser().getEmail()); // Set Modified By Email

        SpreadsheetApp.getUi().alert(`EVENT UPDATED: 
        ${description}`);
      }

    } catch (e) {
      // Log any errors and continue with the next row
      console.error(`Error processing row ${i + 1}: ${e.toString()}`);
      SpreadsheetApp.getUi().alert(`An error occurred on row ${i + 1}. Please check the logs.`);
    }
  }
}
