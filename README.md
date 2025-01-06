# Google Calendar ICS Importer

A Google Apps Script to automatically process `.ics` files from Gmail, import events into Google Calendar, update existing events, or delete canceled events.

## Features

- Automatically processes `.ics` files attached to Gmail messages.
- Imports new events into Google Calendar.
- Updates existing events based on changes in `.ics` files.
- Deletes canceled events when detected in `.ics` files.
- Prevents duplicate events.
- Labels processed Gmail threads with a `Processed` label.

## Installation

1. Open [Google Apps Script](https://script.google.com).
2. Create a new project and paste the code from `calendar_importer.js`.
3. Save the script and give it a meaningful name (e.g., `ICS Calendar Importer`).
4. Add a trigger to run the `processICSFiles` function periodically (e.g., every hour).

## Usage

1. Attach `.ics` files to Gmail messages with relevant event details.
2. Run the script manually or wait for the scheduled trigger to execute.
3. Events will be added, updated, or deleted in your Google Calendar.

## License

This project is licensed under the MIT License.
