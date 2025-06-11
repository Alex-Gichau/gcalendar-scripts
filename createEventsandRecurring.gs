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
          recurrenceMode]             // Y
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

        // Set event options
        const eventOptions = {
          description: description,
          location: meetingRoom
        };

        // Create the event series
        const newEventSeries = calendar.createEventSeries(titleLocation, new Date(start), new Date(end), recurrenceMode, eventOptions);
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
      } else if (action === "Create" && recurring === "Monthly On This Day" && !eventId) {

        if (!(meetingStartDate instanceof Date) || !(startTime instanceof Date) || !(endTime instanceof Date) || !(recurringUntil instanceof Date)) {
          console.log(`Skipping row ${i + 1} due to invalid date/time format.`);
          // UI Alert for the created event
          SpreadsheetApp.getUi().alert(`Missing start date / end date / recurring until date on the events: 
        ${titleLocation}
        `);
          continue; // Skip to the next row if dates/times are not valid
        }

        // Convert meeting start day to dynamic week day format
        const startDayNum = meetingStartDate.getDay();
        console.log(parseInt(startDayNum));

        // Determine the day of the week selected and store as dynamic week
        let dynamicWeekDay;
        switch (parseInt(startDayNum)) {
          case 1: dynamicWeekDay = CalendarApp.Weekday.MONDAY; break;
          case 2: dynamicWeekDay = CalendarApp.Weekday.TUESDAY; break;
          case 3: dynamicWeekDay = CalendarApp.Weekday.WEDNESDAY; break;
          case 4: dynamicWeekDay = CalendarApp.Weekday.THURSDAY; break;
          case 5: dynamicWeekDay = CalendarApp.Weekday.FRIDAY; break;
          case 6: dynamicWeekDay = CalendarApp.Weekday.SATURDAY; break;
          case 7: dynamicWeekDay = CalendarApp.Weekday.SUNDAY; break;
          default:
            throw new Error("Could not determine weekday from start day.");
        }

        // Create the recurrence rule
        const recurrence = CalendarApp.newRecurrence()
          .addMonthlyRule().onlyOnWeekday(dynamicWeekDay).onlyOnWeek().until(new Date(recurringUntil));

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
