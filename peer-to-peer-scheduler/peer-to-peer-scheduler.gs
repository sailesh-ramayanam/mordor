// Whenever you publish a new version of the add-on, you must increment the below version number. This helps for book-keeping and comparing with previous versions.
const LAST_PUBLISHED_VERSION = 0;

/**
 * These are 0-based indices. Do not confuse them with column numbers.
 * First column will have index as 0.
 */
const GROUP_NUMBER_INDEX = 0;
const STUDENT_EMAILS_INDEX = 1;
const SSM_EMAILS_INDEX = 2;
const OPS_REQUEST_INDEX = 3;
const WEEKDAYS_WEEKENDS_INDEX = 4;
const OCCURANCE_COUNT_INDEX = 5;
const DATE_INDEX = 6;
const START_TIME_INDEX = 7;
const END_TIME_INDEX = 8;
const MEETING_ID_INDEX = 9;
const MEETING_LINK_INDEX = 10;
const CANCELLATION_DATES_INDEX = 11;
const RECORDING_LINKS_INDEX = 12;
const CHAT_LINKS_INDEX = 13;
/**
 * If you add/remove/modify columns, you must also change
 * 1. DATA_END_COLUMN
 * 2. MSG_HELP
 */

const DATA_BEGIN_ROW = 2;
const DATA_END_COLUMN = "N";

const FLAG_IGNORE_ROW = 0;
const FLAG_SEND_INVITE = 1;
const FLAG_CANCEL_INVITE = 2;

const FLAG_WEEKDAY = 1;
const FLAG_WEEKEND = 2;

const NOT_FOUND = "NOT FOUND";
const RECORDING_MIME_TYPE = "video/mp4";
const CHAT_MIME_TYPE = "text/plain";

const MENU_TITLE = "10x P2P";
const MENU_INVITES_TO_STUDENTS = "Send invites to students";
const MENU_GET_RECORDING_AND_CHAT_LINKS = "Get recording and chat links"
const MENU_CANCEL_MEETINGS = "Cancel meetings";
const MENU_HELP = "Help";

const TITLE_ERROR = "Error";
const TITLE_SUCCESS = "Success";
const TITLE_HELP = "Help";

const MSG_CALENDAR_REJECTED = "You chose not to use the current calendar.";
const MSG_CALENDAR_NOT_FOUND = "Calendar not found";
const MSG_COULD_NOT_FETCH_DATA = "Unable to fetch data from the sheet";
const MSG_HELP = `
Version: ${LAST_PUBLISHED_VERSION + 1}
- Your sheet should be in the following format
- 1st row must contain column names
- Actual data must start from 2nd row

- Your sheet must have the following columns in the given order

1. Group Number - To be filled by the team.
2. Student Emails - Email ids of the students - Comma seperated values. To be filled by the team.
3. SSM emails - Additional emails you want to include in the meeting - Comma separated values.
4. Ops Request - (0 - ignore the row),(1 - for meet invites),(2 - for Cancel Meeting) To be filled by the team.
5. Weekday / Weekend - (1 - Weekday), (2 - Weekend). To be filled by the team.
6. Occurance Count - No of days the meet should occur. To be filled by the team.
7. Date - Date of P2P in yyyy-mm-dd format. To be filled by the team.
8. Start time - Start time of P2P in hh:mm format (24 hours). To be filled by the team.
9. End time - End time of P2P in hh:mm format (24 hours). To be filled by the team.
10. Meeting ID - Auto populated by the script.
11. Meeting link - Auto populated by the script.
12. Cancellation Dates - Dates in yyyy-mm-dd format - Comma seperated values. To be filled by the team.
13. Recording Links - Auto populated by the script.
14. Chat Links - Auto populated by the script.

NOTE: If you want to cancel full invite for a row, Give 2 in Cancellation Dates (12th Column)
`;

const DATE_DELIMITER = "-";

const MEETING_TITLE_PREFIX = "P2P | Group ";
const MEETING_TITLE_SUFFIX = " | Placement readiness - 10x Academy"
const MEETING_FIRST_REMINDER_MINUTES = 60;
const MEETING_SECOND_REMINDER_MINUTES = 30;

const ERROR_CODE_CALENDAR_NOT_FOUND = "10x-error-calendar-not-found";
const ERROR_CODE_COULD_NOT_GET_DATA = "10x-error-could-not-get-data";
const ERROR_CODE_COULD_NOT_GET_EVENT = "10x-error-could-not-get-event";
const ERROR_CODE_COULD_NOT_CREATE_EVENT = "10x-error-could-not-create-event";
const ERROR_CODE_COULD_NOT_CANCEL_EVENT = "10x-error-could-not-cancel-event";

const INFO_CODE_CALENDAR_REJECTED = "10x-info-calendar-rejected";

function onInstall(e) {
  // Without this, user will have to reload the sheet after installing.
  onOpen(e);
}

function onOpen(e) {
  SpreadsheetApp.getUi()
  .createMenu(MENU_TITLE)
  .addItem(MENU_INVITES_TO_STUDENTS, "sendInvitesToStudents")
  .addItem(MENU_GET_RECORDING_AND_CHAT_LINKS, "getRecordingAndChatLinks")
  .addItem(MENU_CANCEL_MEETINGS, "cancelEvents")
  .addItem(MENU_HELP, "help")
  .addToUi();
}

function logBeinDelimiter() {
  console.log("BEGIN: 10x---------------10x");
}

function logEndDelimiter() {
  console.log("END: 10x---------------10x");
}

function getCalendar() {
  let calendar = null;
  try {
    calendar = CalendarApp.getDefaultCalendar();
    if (calendar) {
      calendar.setTimeZone("Asia/Kolkata");
    }
  } catch (error) {
    logBeinDelimiter();
    console.log(ERROR_CODE_CALENDAR_NOT_FOUND);
    console.log(error);
    logEndDelimiter();
  }
  return calendar;
}

/**
 * Ask the user if they want to use the given calendar.
 */
function confirmCalendar(calendar) {
  if (!calendar) {
    return false;
  }

  let calendarId = calendar.getId();
  let calendarName = calendar.getName();
  let question = `Current calendar belongs to the following account. Do you want to use this account?

  Name: ${calendarName}
  Id: ${calendarId}
  `;

  let ui = SpreadsheetApp.getUi();
  let userResponse = ui.alert("Please confirm", question, ui.ButtonSet.YES_NO);
  if (userResponse === ui.Button.YES) {
    return true;
  }
  logBeinDelimiter();
  console.log(INFO_CODE_CALENDAR_REJECTED);
  logEndDelimiter();
  return false;
}

function showInfo(title, body) {
  let ui = SpreadsheetApp.getUi();
  ui.alert(title, body, ui.ButtonSet.OK);
}

/**
 * dateString in yyyy-mm-dd format
 * e.g. 2022-10-05 (represents 5th October, 2022)
 */
function parseDate(dateString) {
  try {
    if (!dateString) {
      return null;
    }

    let parts = dateString.trim().split(DATE_DELIMITER);
    if (parts.length !== 3) {
      return null;
    }

    let year = parseInt(parts[0]);
    if (year !== 2022 && year !== 2023) {
      return null;
    }

    let month = parseInt(parts[1]);
    if (month < 1 || month > 12) {
      return null;
    }

    let day = parseInt(parts[2]);
    if (day < 1 || day > 31) {
      return null;
    }

    return {year: year, month: month, day: day};
  } catch (error) {
    return null;
  }
}

/**
 * timeString in hh:mm format (24 hours)
 * e.g. 14:30 (represents 2:30 pm)
 */
function parseTime(timeString) {
  try {
    let parts = timeString.trim().split(':');
    if (parts.length !== 2) {
      return null;
    }

    let hours = parseInt(parts[0]);
    if (hours < 0 || hours > 23) {
      return null;
    }

    let minutes = parseInt(parts[1]);
    if (minutes < 0 || minutes > 59) {
      return null;
    }

    return {hours: hours, minutes: minutes};
  } catch (error) {
    return null;
  }
}

function validateEmail(email) {
  try {
    return email.endsWith("@gmail.com") || email.endsWith("@the10xacademy.com");
  } catch (error) {
    return false;
  }
}

function getEvent(calendar, meetId) {
  try {
    return calendar.getEventById(meetId);
  } catch (error) {
    logBeinDelimiter();
    console.log(ERROR_CODE_COULD_NOT_GET_EVENT);
    console.log(error);
    logEndDelimiter();
    return null;
  }
}

function get1to1Data(sheet) {
  try {
    let dataRange = sheet.getRange(`A${DATA_BEGIN_ROW}:${DATA_END_COLUMN}${sheet.getLastRow()}`);
    return dataRange.getDisplayValues();
  } catch (error) {
    logBeinDelimiter();
    console.log(ERROR_CODE_COULD_NOT_GET_DATA);
    console.log(error);
    logEndDelimiter();
    return null;
  }
}

function initSetup() {
  let result = {
    success: false,
    calendar: null,
    activeSheet: null,
    oneToOneData: null
  };
  result.calendar = getCalendar();
  if (!result.calendar) {
    showInfo(TITLE_ERROR, MSG_CALENDAR_NOT_FOUND);
    return result;
  }

  if (!confirmCalendar(result.calendar)) {
    showInfo(TITLE_ERROR, MSG_CALENDAR_REJECTED);
    return result;
  }

  result.activeSheet = SpreadsheetApp.getActiveSheet();
  result.oneToOneData = get1to1Data(result.activeSheet);
  if (!result.oneToOneData) {
    showInfo(TITLE_ERROR, MSG_COULD_NOT_FETCH_DATA);
    return result;
  }
  result.success = true;
  return result;
}

function createMeeting(calendar, details) {
  let result = {errorMsg: "", eventId: null, eventLink: null};

  if (!calendar) {
    result.errorMsg = MSG_CALENDAR_NOT_FOUND;
    return result;
  }

  let guestList = "";

  let studentEmails = details[STUDENT_EMAILS_INDEX].split(",");
  for (let eIndex = 0; eIndex < studentEmails.length; ++eIndex) {
    let currEmail = studentEmails[eIndex].trim();
    if (!validateEmail(currEmail)) {
      result.errorMsg = "Invalid email: " + currEmail;
      return result;
    }
    studentEmails[eIndex] = currEmail;
  }

  guestList = studentEmails.join(",");

  let ssmEmails = details[SSM_EMAILS_INDEX].split(",");
  for (let eIndex = 0; eIndex < ssmEmails.length; ++eIndex) {
    let currEmail = ssmEmails[eIndex].trim();
    if (validateEmail(currEmail)) {
      guestList += ("," + currEmail);
    }
  }

  let groupNumber = details[GROUP_NUMBER_INDEX].trim();
  if (groupNumber.length === 0) {
    // If Group number is not given, we will take Group number as empty
    groupNumber = "";
  }

  let title = MEETING_TITLE_PREFIX + groupNumber + MEETING_TITLE_SUFFIX;

  let meetDate = parseDate(details[DATE_INDEX]);
  if (!meetDate) {
    result.errorMsg = "Invalid date: " + details[DATE_INDEX];
    return result;
  }

  let startTime = parseTime(details[START_TIME_INDEX]);
  if (!startTime) {
    result.errorMsg = "Invalid start time: " + details[START_TIME_INDEX];
    return result;
  }

  let endTime = parseTime(details[END_TIME_INDEX]);
  if (!endTime) {
    result.errorMsg = "Invalid end time: " + details[END_TIME_INDEX];
    return result;
  }

  if (endTime.hours < startTime.hours || (endTime.hours === startTime.hours && endTime.minutes < startTime.minutes)) {
    result.errorMsg = "Invalid end time: " + details[END_TIME_INDEX];
    return result;
  }

  let meetStart = new Date(meetDate.year, meetDate.month - 1, meetDate.day, startTime.hours, startTime.minutes);
  let meetEnd = new Date(meetDate.year, meetDate.month - 1, meetDate.day, endTime.hours, endTime.minutes);

  let dayType = parseInt(details[WEEKDAYS_WEEKENDS_INDEX].trim());
  let occuranceCount = parseInt(details[OCCURANCE_COUNT_INDEX].trim());

  if (dayType === FLAG_WEEKDAY) {
    try {
      let event = calendar.createEventSeries(title, meetStart, meetEnd,
        CalendarApp.newRecurrence().addWeeklyRule()
        .onlyOnWeekdays([CalendarApp.Weekday.MONDAY, CalendarApp.Weekday.TUESDAY, CalendarApp.Weekday.WEDNESDAY, CalendarApp.Weekday.THURSDAY]).times(occuranceCount),
        {guests: guestList, sendInvites: true});
      event.setGuestsCanInviteOthers(false)
      .setGuestsCanModify(false)
      .setGuestsCanSeeGuests(true);

      result.eventId = event.getId().replace("@google.com", "");
      let eventResource = Calendar.Events.get(calendar.getId(), result.eventId);
      result.eventLink = eventResource.hangoutLink;
    } catch (error) {
      result.errorMsg = "Event creation failed. " + error.toString();
    }
  } else if (dayType === FLAG_WEEKEND) {
    try {
      let event = calendar.createEventSeries(title, meetStart, meetEnd,
        CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekdays([CalendarApp.Weekday.SATURDAY, CalendarApp.Weekday.SUNDAY]).times(occuranceCount),
        {guests: guestList, sendInvites: true});
      event.setGuestsCanInviteOthers(false)
      .setGuestsCanModify(false)
      .setGuestsCanSeeGuests(true);

      result.eventId = event.getId().replace("@google.com", "");
      let eventResource = Calendar.Events.get(calendar.getId(), result.eventId);
      result.eventLink = eventResource.hangoutLink;
    } catch (error) {
      result.errorMsg = "Event creation failed. " + error.toString();
    }
  }
  return result;
}

function sendInvitesToStudents() {
  let setupResult = initSetup();
  if (!setupResult.success) {
    return;
  }
  let { calendar, activeSheet, oneToOneData } = setupResult;

  let errorRows = [];
  let skippedRows = [];
  let successRows = [];
  for (let rowIndex = 0; rowIndex < oneToOneData.length; ++rowIndex) {
    let rowNumber = rowIndex + DATA_BEGIN_ROW;
    let opsRequestFlag = parseInt(activeSheet.getRange(rowNumber, OPS_REQUEST_INDEX + 1).getValue());
    if (opsRequestFlag !== FLAG_SEND_INVITE) {
      // The row should be skipped
      skippedRows.push(rowNumber);
      continue;
    }

    let meetIdCell = activeSheet.getRange(rowNumber, MEETING_ID_INDEX + 1);
    let meetLinkCell = activeSheet.getRange(rowNumber, MEETING_LINK_INDEX + 1);
    if (meetIdCell.getValue().trim() !== "") {
      // Meeting is already set
      skippedRows.push(rowNumber);
      continue;
    }

    let result = createMeeting(calendar, oneToOneData[rowIndex]);
    if (!result.eventId) {
      errorRows.push({rowNumber: rowNumber, error: result.errorMsg});
      logBeinDelimiter();
      console.log(ERROR_CODE_COULD_NOT_CREATE_EVENT);
      console.log(result.errorMsg);
      logEndDelimiter();
      continue;
    }
    meetIdCell.setValue(result.eventId);
    meetLinkCell.setValue(result.eventLink);
    successRows.push(rowNumber);
  }

  let summary = `Number of meetings created: ${successRows.length}
  Number of skipped rows: ${skippedRows.length} (meeting was already created previously or skipped intentionally)
  Number of rows with errors: ${errorRows.length}
  ---
  Following are the error rows.
  `
  for (let i = 0; i < errorRows.length; ++i) {
    summary += `\nRow ${errorRows[i].rowNumber}: ${errorRows[i].error}`;
  }
  summary += `\n---
  Following are skipped rows.
  `
  for (let i = 0; i < skippedRows.length; ++i) {
    summary += `\nRow ${skippedRows[i]}`;
  }
  showInfo(TITLE_SUCCESS, summary);
}

function getRecordingAndChatLinks() {
  let setupResult = initSetup();
  if (!setupResult.success) {
    return;
  }
  let { calendar, activeSheet, oneToOneData } = setupResult;

  let noMeetingRows = [];
  let invalidMeetings = [];
  let noRecordingLinks = [];
  let noChatLinks = [];
  let numSuccess = 0;
  let numSkipped = 0;
  let calendarId = calendar.getId();

  for (let rowIndex = 0; rowIndex < oneToOneData.length; ++rowIndex) {
    let rowNumber = rowIndex + DATA_BEGIN_ROW;
    let opsRequestFlag = parseInt(activeSheet.getRange(rowNumber, OPS_REQUEST_INDEX + 1).getValue());
    if (opsRequestFlag !== FLAG_SEND_INVITE) {
      // The row should be skipped
      ++numSkipped;
      continue;
    }

    let meetIdCell = activeSheet.getRange(rowNumber, MEETING_ID_INDEX + 1);
    let meetId = meetIdCell.getValue().trim();
    if (meetId === "") {
      // Meeting is not set
      noMeetingRows.push(rowNumber);
      continue;
    }

    let event = getEvent(calendar, meetId);
    if (!event) {
      invalidMeetings.push(rowNumber);
      continue;
    }

    let instancesIds = [];
    Calendar.Events.instances(calendarId, meetId).items.forEach(
      function (instance) {
        instancesIds.push(instance.id);
      }
    );

    let recordingLinkCell = activeSheet.getRange(rowNumber, RECORDING_LINKS_INDEX + 1);
    let chatLinkCell = activeSheet.getRange(rowNumber, CHAT_LINKS_INDEX + 1);
    let videoFound = false;
    let chatFound = false;

    let recordingLinksArray = [];
    let chatLinksArray = [];

    for (let currentInstance = 0; currentInstance < instancesIds.length; ++currentInstance) {
      let eventResource = Calendar.Events.get(calendarId, instancesIds[currentInstance]);
      if (!eventResource) {
        console.log("Event could not be retrieved - " + calendarId + ", " + instancesIds[currentInstance]);
        continue;
      }
      let eventAttachments = eventResource.attachments;
      
      if (!eventAttachments) {
        continue;
      }

      for (let i = 0; i < eventAttachments.length; ++i) {
        let currentAttachment = eventAttachments[i];
        let linkUrl = currentAttachment.fileUrl;
        if (!currentAttachment || !linkUrl) {
          continue;
        }
        if (currentAttachment.mimeType === RECORDING_MIME_TYPE) {
          videoFound = true;
          recordingLinksArray.push(linkUrl);
        } else if (currentAttachment.mimeType === CHAT_MIME_TYPE) {
          chatFound = true;
          chatLinksArray.push(linkUrl);
        }
      }
    }
    recordingLinkCell.setValue(recordingLinksArray.join("\n"));
    chatLinkCell.setValue(chatLinksArray.join("\n"));
    if (videoFound && chatFound) {
      ++numSuccess;
      continue;
    }
    if (!videoFound) {
      recordingLinkCell.setValue(NOT_FOUND);
      noRecordingLinks.push(rowNumber);
    }
    if (!chatFound) {
      chatLinkCell.setValue(NOT_FOUND);
      noChatLinks.push(rowNumber);
    }
  }

  let summary = `Number of rows for which recording and chat links were generated: ${numSuccess}
  Number of rows for which no recording links were found: ${noRecordingLinks.length}
  Number of rows for which no chat links were found: ${noChatLinks.length}
  Number of rows skipped: ${numSkipped} (Links were already present or intentionally skipped)
  Number of rows with invalid meeting id: ${invalidMeetings.length}
  Number of rows with no meeting: ${noMeetingRows.length}
  `;
  
  summary += `\n---

  Following rows have invalid meeting id.
  `;
  for (let i = 0; i < invalidMeetings.length; ++i) {
    summary += `\nRow ${invalidMeetings[i]}`;
  }
  
  summary += `\n---

  Following rows have no recording links.
  `;
  for (let i = 0; i < noRecordingLinks.length; ++i) {
    summary += `\nRow ${noRecordingLinks[i]}`;
  }

  summary += `\n---

  Following rows have no chat links.
  `;
  for (let i = 0; i < noChatLinks.length; ++i) {
    summary += `\nRow ${noChatLinks[i]}`;
  }

  summary += `\n---

  Following rows have no meeting.
  `;
  for (let i = 0; i < noMeetingRows.length; ++i) {
    summary += `\nRow ${noMeetingRows[i]}`;
  }
  showInfo(TITLE_SUCCESS, summary);
}

function cancelEvents() {
  let setupResult = initSetup();
  if (!setupResult.success) {
    return;
  }
  let { calendar, activeSheet, oneToOneData } = setupResult;
  let calendarId = calendar.getId();

  let canceledRows = [];
  let invalidMeetings = [];
  let partialCanceledRows = [];
  for (let rowIndex = 0; rowIndex < oneToOneData.length; ++rowIndex) {
    let rowNumber = rowIndex + DATA_BEGIN_ROW;
    let opsRequestFlag = parseInt(activeSheet.getRange(rowNumber, OPS_REQUEST_INDEX + 1).getValue());
    if (opsRequestFlag !== FLAG_CANCEL_INVITE) {
      // This row doesn't need to be canceled
      continue;
    }

    let meetIdCell = activeSheet.getRange(rowNumber, MEETING_ID_INDEX + 1);
    let meetLinkCell = activeSheet.getRange(rowNumber, MEETING_LINK_INDEX + 1);
    let meetId = meetIdCell.getValue().trim();
    if (meetId === "") {
      // Meeting is not set
      continue;
    }

    let event = getEvent(calendar, meetId);
    if (!event) {
      continue;
    }
    
    let cancelDatesCell = activeSheet.getRange(rowNumber, CANCELLATION_DATES_INDEX + 1);
    let cancelDates = cancelDatesCell.getValue();
    if (cancelDates === "") {
      // No canellation dates are given
      continue;
    }

    if (parseInt(cancelDates) === FLAG_CANCEL_INVITE) { //Full Invite cancellation
      try {
        // If the event was canceled previously (e.g. someone manually canceled), then deleteEvent() throws exception.
        event.deleteEvent();
      } catch (error) {
        logBeinDelimiter();
        console.log(ERROR_CODE_COULD_NOT_CANCEL_EVENT);
        console.log(error);
        logEndDelimiter();
        invalidMeetings.push(rowNumber);
        continue;
      }
      canceledRows.push(rowNumber);
      meetIdCell.clearContent();
      meetLinkCell.clearContent();
      cancelDatesCell.clearContent();
      continue;
    }

    let cancelDatesArray = cancelDates.split(",");
    for (let i = 0; i < cancelDatesArray.length; ++i) {
      if(!parseDate(cancelDatesArray[i].trim())) {
        continue;
      }
    }
    let cancelSuccessDates = [];
    try {
      // If the event was canceled previously (e.g. someone manually canceled), then deleteEvent() throws exception.
      let instancesDates = {} // Each instance id format is "meetId_yyyymmddThhmmssZ"
      Calendar.Events.instances(calendarId, meetId).items.forEach(
        function (instance) {
          currentDate = instance.id.split("_")[1].substring(0,8)
          instancesDates[currentDate] = instance.id;
        }
      );
      for (let i = 0; i < cancelDatesArray.length; ++i) {
        let currentDate = cancelDatesArray[i].trim().split(DATE_DELIMITER).join("");
        if (currentDate in instancesDates) {
          let event = getEvent(calendar, instancesDates[currentDate]);
          if (!event) {
            continue;
          }
          event.deleteEvent();
          cancelSuccessDates.push(cancelDatesArray[i]);
        }
      }
      let instancesCount = Calendar.Events.instances(calendarId, meetId).items.length;
      if (!instancesCount) {
        canceledRows.push(rowNumber);
        meetIdCell.clearContent();
        meetLinkCell.clearContent();
        cancelDatesCell.clearContent();
        continue;
      }
    } catch (error) {
      logBeinDelimiter();
      console.log(ERROR_CODE_COULD_NOT_CANCEL_EVENT);
      console.log(error);
      logEndDelimiter();
      invalidMeetings.push(rowNumber);
      continue;
    }
    partialCanceledRows.push([rowNumber, cancelSuccessDates.join(" , ")]);
    cancelDatesCell.clearContent();
  }
  let summary = `Number of meetings fully canceled: ${canceledRows.length}
  Number of meetings partially cancelled: ${partialCanceledRows.length}
  Number of rows with invalid meeting id: ${invalidMeetings.length} (The calendar event does not exist, or it has already been deleted.)
  ---
  Meetings in the following rows have been fully canceled.
  `;
  for (let i = 0; i < canceledRows.length; ++i) {
    summary += `\nRow ${canceledRows[i]}`;
  }
  summary += `\n---
  Meetings in the following rows have been partially canceled.
  `;
  for (let i = 0; i < partialCanceledRows.length; ++i) {
    summary += `\nRow ${partialCanceledRows[i][0]}`;
    summary += `: ${partialCanceledRows[i][1]}`
  }
  summary += `\n---
  Following rows have invalid meeting id.
  `;
  for (let i = 0; i < invalidMeetings.length; ++i) {
    summary += `\nRow ${invalidMeetings[i]}`;
  }
  showInfo(TITLE_SUCCESS, summary);
}

function help() {
  showInfo(TITLE_HELP, MSG_HELP);
}
