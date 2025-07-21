# Event Scheduler

Event Scheduler is a Google Apps Script project designed to seamlessly synchronize events between a Google Sheets spreadsheet and Google Calendar. This tool allows users to add, update, or delete calendar events directly from a spreadsheet, making event management more efficient and collaborative.

## Features

- **Sync to Calendar:** Add, update, or delete Google Calendar events from the "Scheduler" sheet.
- **Custom Menu:** Easily access synchronization via a custom "Scheduler" menu in Google Sheets.
- **Bulk Actions:** Perform actions on multiple events at once using the active range.
- **Guest Management:** Add multiple guests to events by entering their emails.
- **Time Zone Support:** Automatically formats event times to Asia/Jakarta time zone.
- **Status Tracking:** Tracks the status and event ID for each row to prevent duplication and manage updates.

## How It Works

1. **Prepare the Spreadsheet:**
   - Use a sheet named `Scheduler`.
   - Each row represents an event with columns for action (`Add`, `Update`, `Delete`), event details, guests, and status.

2. **Trigger the Script:**
   - Open the Google Sheet.
   - Use the "Scheduler" menu and select "Sync to Calendar" to process the selected rows.

3. **Actions:**
   - **Add:** Creates a new event if no Event ID exists.
   - **Update:** Modifies an existing event using the Event ID.
   - **Delete:** Removes the event from Google Calendar using the Event ID.

4. **Status & Event ID:**
   - After each action, the script updates the status and event ID columns for tracking.

## Usage

1. **Install the Script:**
   - Open your Google Sheet.
   - Go to `Extensions > Apps Script` and paste the contents of `scheduler.js`.

2. **Set Up the Sheet:**
   - Ensure your sheet is named `Scheduler`.
   - Structure your columns as follows (example):
     1. Action (`Add`, `Update`, `Delete`)
     2. ... (other columns as needed)
     9. Status
     10. Event ID
     11. Start Time
     12. End Time
   - Here’s the sample spreadsheet: https://docs.google.com/spreadsheets/d/1_8SaymsbrK59rSl_nPaHpctmXR9KBKe2oeBcjfRkqDM/edit?usp=sharing

3. **Authorize the Script:**
   - On first run, grant the necessary permissions for Calendar and Sheets access.

4. **Sync Events:**
   - Select the rows you want to process.
   - Click `Scheduler > Sync to Calendar`.

## Notes

- The script uses the active user's email as the default calendar. You can change the `calendarId` variable if needed.
- Ensure date and time columns are formatted correctly in the sheet.
- Guest emails should be separated by new lines.

## License

MIT License

---

*Developed for DevCoach by Dicoding.*
