/**
 * Google Apps Script: Calendar Importer
 * 
 * Description:
 * A script to automatically import, update, or delete events from `.ics` files
 * attached to Gmail messages. The script processes messages, checks for duplicates,
 * updates existing events, and handles event cancellations.
 * 
 * Usage:
 * Add this script to your Google Apps Script project and set up a trigger
 * to run the `processICSFiles` function at your preferred interval.
 * 
 * Author: [Your Name or GitHub Username]
 * License: MIT
 */

// Global array to store deleted events to prevent re-adding them
var deletedEvents = [];

/**
 * Main function to process `.ics` files from Gmail messages.
 */
function processICSFiles() {
  createLabelIfNotExists("Processed");

  var threads = GmailApp.search('has:attachment filename:ics newer_than:1d -label:Processed');
  threads.forEach(function (thread) {
    var messages = thread.getMessages();

    messages.forEach(function (message) {
      var subject = message.getSubject();
      var attachments = message.getAttachments();

      attachments.forEach(function (attachment) {
        if (attachment.getContentType() === "text/calendar") {
          var icsContent = attachment.getDataAsString();

          if (subject.includes("Событие") && subject.includes("отменено")) {
            deleteEvent(icsContent);
          } else if (subject.startsWith("Изменение события")) {
            updateEvent(icsContent);
          } else {
            importToCalendar(icsContent);
          }
        }
      });

      markThreadAsProcessed(thread);
    });
  });
}

/**
 * Create Gmail label if it does not exist.
 * @param {string} labelName - Name of the label to create.
 */
function createLabelIfNotExists(labelName) {
  var label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    GmailApp.createLabel(labelName);
  }
}

/**
 * Mark Gmail thread as processed by adding the "Processed" label.
 * @param {GmailThread} thread - The Gmail thread to label.
 */
function markThreadAsProcessed(thread) {
  var label = GmailApp.getUserLabelByName("Processed");
  if (label) {
    thread.addLabel(label);
  }
}

/**
 * Import a new event into the calendar.
 * @param {string} icsContent - The content of the `.ics` file.
 */
function importToCalendar(icsContent) {
  var eventDetails = parseICSContent(icsContent);

  if (deletedEvents.includes(eventDetails.title)) {
    console.log("Event was previously deleted and will not be added: " + eventDetails.title);
    return;
  }

  if (isDuplicateEvent(eventDetails)) {
    console.log("Duplicate event detected: " + eventDetails.title);
    return;
  }

  var validAttendees = eventDetails.attendees.filter(isValidEmail);

  var calendar = CalendarApp.getDefaultCalendar();
  calendar.createEvent(eventDetails.title, eventDetails.start, eventDetails.end, {
    description: eventDetails.description + "\n\nImported automatically by script",
    location: eventDetails.location,
    guests: validAttendees.join(","),
    sendInvites: false
  });

  console.log("Event added: " + eventDetails.title);
}

/**
 * Update an existing event in the calendar.
 * @param {string} icsContent - The content of the `.ics` file.
 */
function updateEvent(icsContent) {
  var eventDetails = parseICSContent(icsContent);

  if (deletedEvents.includes(eventDetails.title)) {
    console.log("Event was previously deleted and will not be updated: " + eventDetails.title);
    return;
  }

  if (!(eventDetails.start instanceof Date) || isNaN(eventDetails.start.getTime())) {
    console.log("Invalid start time: " + eventDetails.start);
    return;
  }

  var calendar = CalendarApp.getDefaultCalendar();
  var events = calendar.getEvents(eventDetails.start, eventDetails.end);

  for (var i = 0; i < events.length; i++) {
    var existingEvent = events[i];

    if (existingEvent.getTitle() === eventDetails.title) {
      console.log("Updating event: " + eventDetails.title);

      existingEvent.deleteEvent();
      console.log("Old event deleted: " + eventDetails.title);

      var validAttendees = eventDetails.attendees.filter(isValidEmail);

      calendar.createEvent(eventDetails.title, eventDetails.start, eventDetails.end, {
        description: eventDetails.description + "\n\nImported automatically by script",
        location: eventDetails.location,
        guests: validAttendees.join(","),
        sendInvites: false
      });

      console.log("New event added: " + eventDetails.title);
      return;
    }
  }

  console.log("No matching event found to update: " + eventDetails.title);
}

/**
 * Delete an event from the calendar.
 * @param {string} icsContent - The content of the `.ics` file.
 */
function deleteEvent(icsContent) {
  var eventDetails = parseICSContent(icsContent);

  if (!deletedEvents.includes(eventDetails.title)) {
    deletedEvents.push(eventDetails.title);
  }

  var calendar = CalendarApp.getDefaultCalendar();
  var events = calendar.getEvents(eventDetails.start, eventDetails.end);

  for (var i = 0; i < events.length; i++) {
    var existingEvent = events[i];

    if (existingEvent.getTitle() === eventDetails.title) {
      existingEvent.deleteEvent();
      console.log("Event deleted: " + eventDetails.title);
      return;
    }
  }

  console.log("No matching event found to delete: " + eventDetails.title);
}

/**
 * Check for duplicate events in the calendar.
 * @param {Object} eventDetails - The details of the event to check.
 * @returns {boolean} True if duplicate is found, otherwise false.
 */
function isDuplicateEvent(eventDetails) {
  var calendar = CalendarApp.getDefaultCalendar();
  var events = calendar.getEvents(eventDetails.start, eventDetails.end);

  for (var i = 0; i < events.length; i++) {
    var existingEvent = events[i];

    if (existingEvent.getTitle() === eventDetails.title &&
        existingEvent.getStartTime().getTime() === eventDetails.start.getTime() &&
        existingEvent.getEndTime().getTime() === eventDetails.end.getTime()) {
      return true;
    }
  }

  return false;
}

/**
 * Parse `.ics` content into event details.
 * @param {string} icsContent - The content of the `.ics` file.
 * @returns {Object} Event details.
 */
function parseICSContent(icsContent) {
  var eventLines = icsContent.split("\n");
  var title = "";
  var description = "";
  var location = "";
  var startTime = "";
  var endTime = "";
  var attendees = [];
  var isDescriptionLine = false;

  eventLines.forEach(function (line) {
    if (line.startsWith("SUMMARY:")) {
      title = line.replace("SUMMARY:", "").trim();
      isDescriptionLine = false;
    } else if (line.startsWith("DESCRIPTION:")) {
      description += line.replace("DESCRIPTION:", "").trim();
      isDescriptionLine = true;
    } else if (isDescriptionLine && (line.startsWith(" ") || line.startsWith("\t"))) {
      description += " " + line.trim();
    } else if (line.startsWith("LOCATION:")) {
      location = line.replace("LOCATION:", "").trim();
      isDescriptionLine = false;
    } else if (line.startsWith("DTSTART;")) {
      startTime = line.split(":")[1].trim();
      isDescriptionLine = false;
    } else if (line.startsWith("DTEND;")) {
      endTime = line.split(":")[1].trim();
      isDescriptionLine = false;
    } else if (line.startsWith("ATTENDEE;") || line.startsWith("ATTENDEE:")) {
      var emailMatch = line.match(/mailto:([^>\s]+)/i);
      if (emailMatch) {
        var email = emailMatch[1].trim();
        if (isValidEmail(email)) {
          attendees.push(email);
        }
      }
      isDescriptionLine = false;
    }
  });

  return {
    title: title,
    description: description.replace(/\\n/g, "\n").trim(),
    location: location,
    start: convertICSToDate(startTime),
    end: convertICSToDate(endTime),
    attendees: attendees
  };
}

/**
 * Convert `.ics` date to JavaScript Date object.
 * @param {string} icsDate - The date string in `.ics` format.
 * @returns {Date} JavaScript Date object.
 */
function convertICSToDate(icsDate) {
  if (!icsDate) {
    console.error("Empty date in `.ics` file");
    return null;
  }

  try {
    var year = parseInt(icsDate.slice(0, 4), 10);
    var month = parseInt(icsDate.slice(4, 6), 10) - 1;
    var day = parseInt(icsDate.slice(6, 8), 10);
    var hours = parseInt(icsDate.slice(9, 11), 10) || 0;
    var minutes = parseInt(icsDate.slice(11, 13), 10) || 0;

    return new Date(year, month, day, hours, minutes);
  } catch (error) {
    console.error("Error converting `.ics` date: " + error);
    return null;
  }
}

/**
 * Validate email address format.
 * @param {string} email - The email address to validate.
 * @returns {boolean} True if valid, otherwise false.
 */
function isValidEmail(email) {
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}
