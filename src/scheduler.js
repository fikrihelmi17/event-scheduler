// Adjust the value of the calendarId variable to the desired Google Calendar ID.
let calendarId = Session.getActiveUser().getEmail();

// This function adds a menu in Google Sheets named "Scheduler" with an item called "Sync to Calendar."
function onOpen() {
    let ui = SpreadsheetApp.getUi();
    ui.createMenu("Scheduler")
      .addItem("Sync to Calendar", "syncCalendar")
      .addToUi();
}

/* This function will read data from Google Sheets and determine the action to be performed:  
- Add (add an event)  
- Update (update an event)  
- Delete (delete an event) 
*/  
function syncCalendar() {
    let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName("Scheduler");
    let range = sheet.getActiveRange();
    let firstRow = range.getRow();
    let values = range.getValues();

    for (i = 0; i < values.length; i++) {
        let event = values[i];
        let rowNumber = firstRow + i;
        let returnedData = [];

        if (event[0] != '') {
            switch (event[0]) {
                case 'Add':
                    returnedData = addEvent(event)
                    break;
                case 'Update':
                    returnedData = updateEvent(event);
                    break;
                case 'Delete':
                    returnedData = deleteEvent(event);
                    break;
            }
            // Ensure returnedData has exactly 2 values before setting
            if (returnedData.length === 2) {
                sheet.getRange(rowNumber, 9, 1, 2).setValues([returnedData]);
            }
        }
    }
    sheet.getRange(firstRow, 1, values.length, 1).clearContent();
}

/*
  Adds an event to Google Calendar from a spreadsheet row.
  This function checks if the Event ID column is empty before adding a new event.
  If an Event ID already exists, it prevents duplication by returning an error message.
 */
function addEvent(event) {
    // Check if event ID is already set
    if (event[9]) {
        let existingStatus = event[8];
        let existingEventId = event[9];
        SpreadsheetApp.getUi().alert("Error: Event already exists!"); // Show popup alert
        return [existingStatus, existingEventId];
    }

    let title = event[4];
    let description = event[5];
    let emails = event[6];
    let notify = event[7];
    let startTime = new Date(event[10]);
    let endTime = new Date(event[11]);

    // Ensure Proper Jakarta Time Conversion
    let startTimeFormatted = Utilities.formatDate(startTime, "Asia/Jakarta", "yyyy-MM-dd'T'HH:mm:ss");
    let endTimeFormatted = Utilities.formatDate(endTime, "Asia/Jakarta", "yyyy-MM-dd'T'HH:mm:ss");

    let attendees = [];
    // Add multiple guests if emails are provided
    if (emails) {
        attendees = emails.split(/\r?\n/).map(email => ({
          email: email.trim(),
          responseStatus: "needsAction"
        }));
    }

    let eventDetails = {
      summary: title,
      description: description || "",
      start: { 
        dateTime: startTimeFormatted, 
        timeZone: "Asia/Jakarta"
      },
      end: { 
        dateTime: endTimeFormatted, 
        timeZone: "Asia/Jakarta"
      },
      attendees: attendees,
      guestsCanSeeOtherGuests: false,
      sendUpdates: notify ? "all" : "none"
    };

    let createdEvent = Calendar.Events.insert(eventDetails, calendarId);
    return ["Added", createdEvent.getId()];
}

/*
  Updates an existing event in Google Calendar based on spreadsheet data.
  
  This function retrieves event details from the spreadsheet, finds the corresponding
  event in the Google Calendar, and updates its title, description, 
  time, and guest list. If the event is not found, it displays an error message 
  and returns the existing status and event ID.
*/
function updateEvent(event) {
    let title = event[4];
    let description = event[5];
    let emails = event[6];
    let eventId = event[9];
    let startTime = new Date(event[10]);
    let endTime = new Date(event[11]);

    // Format the start and end times to ISO 8601 format
    let formattedStartTime = Utilities.formatDate(startTime, "Asia/Jakarta", "yyyy-MM-dd'T'HH:mm:ss");
    let formattedEndTime = Utilities.formatDate(endTime, "Asia/Jakarta", "yyyy-MM-dd'T'HH:mm:ss");

    let attendees = [];
    if (emails) {
        attendees = emails.split(/\r?\n/).map(email => ({
          email: email.trim(),
          responseStatus: "needsAction"
        }));
    }

    let eventDetails = {
      summary: title,
      description: description || "",
      start: { 
        dateTime: formattedStartTime, 
        timeZone: "Asia/Jakarta"
      },
      end: { 
        dateTime: formattedEndTime,
        timeZone: "Asia/Jakarta"
      },
      attendees: attendees,
      guestsCanSeeOtherGuests: false,
    };

    try {
        // Update the event using Google Calendar API
        let updatedEvent = Calendar.Events.update(eventDetails, calendarId, eventId);
        return ["Updated", updatedEvent.id];
    } catch (e) {
        SpreadsheetApp.getUi().alert("Error: Event not found or failed to update!"); // Show popup alert
        return [event[8], eventId]; // Return previous status
    }
}

/*
  Deletes an existing event from Google Calendar based on spreadsheet data.
  
  This function retrieves the event ID from the spreadsheet and attempts to find 
  the corresponding event in Google Calendar. If found, it deletes the event. 
  If the event is not found, it displays an error message and returns the existing 
  status and event ID.
*/
function deleteEvent(event) {
    let eventId = event[9];

    try {
        // Use Calendar API to remove the event
        Calendar.Events.remove(calendarId, eventId);
        return ["Deleted", ""]; // Event successfully deleted
    } catch (e) {
        SpreadsheetApp.getUi().alert("Error: Event not found or failed to delete!"); // Show popup alert
        return [event[8], eventId]; // Return previous status
    }
}