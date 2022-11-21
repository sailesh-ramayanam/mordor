// Whenever you publish a new version of the add-on, you must increment the below version number. This helps for book-keeping and comparing with previous versions.
const LAST_PUBLISHED_VERSION = 6;

/**
 * These are 0-based indices. Do not confuse them with column numbers.
 * First column will have index as 0.
 */
const MODULE_NAME_INDEX = 0;
const STUDENT_NAME_INDEX = 1;
const STUDENT_EMAIL_INDEX = 2;
const DATE_INDEX = 3;
const START_TIME_INDEX = 4;
const END_TIME_INDEX = 5;
const MEETING_ID_INDEX = 6;
const STUDENT_RESPONSE_INDEX = 7;
const OPS_REQUEST_INDEX = 8;
const SSM_EMAILS_INDEX = 9;
const MENTOR_EMAIL_INDEX = 10;
const MENTOR_RESPONSE_INDEX = 11;
const RECORDING_LINK_INDEX = 12;
const CHAT_LINK_INDEX = 13;
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

const NOT_INVITED = "NOT INVITED";
const NOT_FOUND = "NOT FOUND";
const RECORDING_MIME_TYPE = "video/mp4";
const CHAT_MIME_TYPE = "text/plain";

const MENU_TITLE = "10x 1:1";
const MENU_INVITES_TO_STUDENTS = "Send invites to students";
const MENU_GET_RESPONSES = "Get responses";
const MENU_INVITES_TO_MENTORS = "Send invites to mentors";
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

1. Module Name - Name of the Module. To be filled by the team.
2. Student Name - Name of the student. To be filled by the team.
3. Student Email - Email id of the student. To be filled by the team.
4. Date - Date of 1:1 in yyyy-mm-dd format. To be filled by the team.
5. Start time - Start time of 1:1 in hh:mm format (24 hours). To be filled by the team.
6. End time - End time of 1:1 in hh:mm format (24 hours). To be filled by the team.
7. Meeting link - Auto populated by the script.
8. Student response - Auto populated by the script.
9. Ops Request - (0 - ignore the row),(1 - for meet invites),(2 - for Cancel Meeting) To be filled by the team.
10. SSM emails - Additional emails you want to include in the meeting - Comma separated values.
11. Mentor email - Email id of the mentor. To be filled by the team.
12. Mentor response - Auto populated by the script.
13. Recording Link - Auto populated by the script.
14. Chat Link - Auto populated by the script.
`;

const DATE_DELIMITER = "-";

const MEETING_TITLE_PREFIX = " - 1:1 (";
const MEETING_TITLE_SUFFIX = ") - 10x Academy"
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
  .addItem(MENU_GET_RESPONSES, "getResponses")
  .addItem(MENU_INVITES_TO_MENTORS, "sendInvitesToMentors")
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
  let result = {errorMsg: "", eventId: null};

  if (!calendar) {
    result.errorMsg = MSG_CALENDAR_NOT_FOUND;
    return result;
  }

  let guestList = "";

  let studentEmail = details[STUDENT_EMAIL_INDEX].trim();
  if (!validateEmail(studentEmail)) {
    result.errorMsg = "Invalid email: " + studentEmail;
    return result;
  }
  guestList = studentEmail;

  let ssmEmails = details[SSM_EMAILS_INDEX].split(",");
  for (let eIndex = 0; eIndex < ssmEmails.length; ++eIndex) {
    let currEmail = ssmEmails[eIndex].trim();
    if (validateEmail(currEmail)) {
      guestList += ("," + currEmail);
    }
  }

  let studentName = details[STUDENT_NAME_INDEX].trim();
  if (studentName.length === 0) {
    // If name is not given, we will take email as name
    studentName = studentEmail;
  }

  let moduleName = details[MODULE_NAME_INDEX].trim();
  if (moduleName.length === 0) {
    // If Module name is not given, we will take module name as Session
    moduleName = "Session";
  }

  let title = studentName + MEETING_TITLE_PREFIX + moduleName + MEETING_TITLE_SUFFIX;

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

  try {
    let event = calendar.createEvent(title, meetStart, meetEnd, {guests: guestList, sendInvites: true});
    event.setGuestsCanInviteOthers(false)
    .setGuestsCanModify(false)
    .setGuestsCanSeeGuests(true);

    result.eventId = event.getId().replace("@google.com", "");
  } catch (error) {
    result.errorMsg = "Event creation failed. " + error.toString();
  }
  return result;
}

function addAttendees(calendarId, meetId, emails) {
  try {
    // Do not use addGuest method. It won't send emails to new guests.
    let eventResource = Calendar.Events.get(calendarId, meetId);
    if (!eventResource) {
      console.log("Event could not be retrieved - " + calendarId + ", " + meetId);
      return false;
    }
    if (!eventResource.attendees) {
      eventResource.attendees = [];
    }
    for (let index = 0; index < emails.length; ++index) {
      eventResource.attendees.push({email: emails[index]});
    }
    let requestBody = {attendees: eventResource.attendees};
    let queryParams = {sendUpdates: "all", conferenceDataVersion: 1};
    Calendar.Events.patch(requestBody, calendarId, meetId, queryParams);
  } catch (error) {
    console.log("Event could not be patched - " + calendarId + ", " + meetId);
    console.log(error.toString());
    return false;
  }
  return true;
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

function getResponsesHelper(calendar, sheet, oneToOneData) {
  let noMeetingRows = [];
  let invalidRows = [];
  let skippedRows = [];
  let numStudentsUpdated = 0;
  let numMentorsUpdated = 0;
  for (let rowIndex = 0; rowIndex < oneToOneData.length; ++rowIndex) {
    let rowNumber = rowIndex + DATA_BEGIN_ROW;
    let opsRequestFlag = parseInt(sheet.getRange(rowNumber, OPS_REQUEST_INDEX + 1).getValue());
    if (opsRequestFlag !== FLAG_SEND_INVITE) {
      // The row should be skipped
      skippedRows.push(rowNumber);
      continue;
    }
    let meetIdCell = sheet.getRange(rowNumber, MEETING_ID_INDEX + 1);
    let meetId = meetIdCell.getValue().trim();
    if (meetId === "") {
      // Meeting is not set
      noMeetingRows.push(rowNumber);
      continue;
    }

    let event = getEvent(calendar, meetId);
    if (!event) {
      invalidRows.push(rowNumber);
      continue;
    }
    
    let guestList = event.getGuestList();
    let studentEmail = sheet.getRange(rowNumber, STUDENT_EMAIL_INDEX + 1).getValue().trim();
    let studentResponseCell = sheet.getRange(rowNumber, STUDENT_RESPONSE_INDEX + 1);
    studentResponseCell.setValue(NOT_INVITED);

    let mentorEmail = sheet.getRange(rowNumber, MENTOR_EMAIL_INDEX + 1).getValue().trim();
    let mentorResponseCell = sheet.getRange(rowNumber, MENTOR_RESPONSE_INDEX + 1);
    mentorResponseCell.setValue(NOT_INVITED);

    for (let i = 0; i < guestList.length; ++i) {
      let currGuest = guestList[i];
      let guestEmail = currGuest.getEmail().trim();
      if (guestEmail === studentEmail) {
        studentResponseCell.setValue(currGuest.getGuestStatus());
        ++numStudentsUpdated;
      } else if (guestEmail === mentorEmail) {
        mentorResponseCell.setValue(currGuest.getGuestStatus());
        ++numMentorsUpdated;
      }
    }
  }
  return {noMeetingRows, skippedRows, invalidRows, numStudentsUpdated, numMentorsUpdated};
}

function getResponses() {
  let setupResult = initSetup();
  if (!setupResult.success) {
    return;
  }
  let { calendar, activeSheet, oneToOneData } = setupResult;

  let responseStats = getResponsesHelper(calendar, activeSheet, oneToOneData);
  let summary = `Number of student responses updated: ${responseStats.numStudentsUpdated}
  Number of mentor responses updated: ${responseStats.numMentorsUpdated}
  Number of skipped rows: ${responseStats.skippedRows.length} (meeting was already created previously or skipped intentionally)
  Number of rows with no meeting: ${responseStats.noMeetingRows.length}
  Number of rows with invalid meeting id: ${responseStats.invalidRows.length}
  ---
  Following are skipped rows.
  `
  for (let i = 0; i < responseStats.skippedRows.length; ++i) {
    summary += `\nRow ${responseStats.skippedRows[i]}`;
  }
  summary += `\n---
  Following rows have no meeting.
  `;
  for (let i = 0; i < responseStats.noMeetingRows.length; ++i) {
    summary += `\nRow ${responseStats.noMeetingRows[i]}`;
  }
  summary += `\n---
  Following rows have invalid meeting id.
  `;
  for (let i = 0; i < responseStats.invalidRows.length; ++i) {
    summary += `\nRow ${responseStats.invalidRows[i]}`;
  }
  showInfo(TITLE_SUCCESS, summary);
}

function sendInvitesToMentors() {
  let setupResult = initSetup();
  if (!setupResult.success) {
    return;
  }
  let { calendar, activeSheet, oneToOneData } = setupResult;

  let noMeetingRows = [];
  let invalidMeetings = [];
  let invalidEmails = [];
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
    
    let mentorEmail = activeSheet.getRange(rowNumber, MENTOR_EMAIL_INDEX + 1).getValue().trim();
    if (!validateEmail(mentorEmail)) {
      invalidEmails.push({row: rowNumber, email: mentorEmail});
      continue;
    }

    let mentorInvited = event.getGuestByEmail(mentorEmail);
    if (!mentorInvited) {
      // Do not use addGuest method. It won't send emails to new guests.
      if (addAttendees(calendarId, meetId, [mentorEmail])) {
        ++numSuccess;
      } else {
        invalidMeetings.push(rowNumber);
      }
    } else {
      ++numSkipped;
    }
  }

  // Update the responses so that there is no confusion with older responses
  getResponsesHelper(calendar, activeSheet, oneToOneData);

  let summary = `Number of rows for which invites have been sent: ${numSuccess}
  Number of rows skipped: ${numSkipped} (invite was already sent previously or intentionally skipped)
  Number of rows with invalid emails: ${invalidEmails.length}
  Number of rows with invalid meeting id: ${invalidMeetings.length}
  Number of rows with no meeting: ${noMeetingRows.length}
  `;
  
  summary += `\n---

  Following rows have invalid email id.
  `;
  for (let i = 0; i < invalidEmails.length; ++i) {
    summary += `\nRow ${invalidEmails[i].row}: ${invalidEmails[i]}.email`;
  }

  summary += `\n---

  Following rows have invalid meeting id.
  `;
  for (let i = 0; i < invalidMeetings.length; ++i) {
    summary += `\nRow ${invalidMeetings[i]}`;
  }

  summary += `\n---

  Following rows have no meeting.
  `;
  for (let i = 0; i < noMeetingRows.length; ++i) {
    summary += `\nRow ${noMeetingRows[i]}`;
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

    let eventResource = Calendar.Events.get(calendarId, meetId);
    if (!eventResource) {
      console.log("Event could not be retrieved - " + calendarId + ", " + meetId);
      invalidMeetings.push(rowNumber);
      continue;
    }

    let eventAttachments = eventResource.attachments;
    
    let recordingLinkCell = activeSheet.getRange(rowNumber, RECORDING_LINK_INDEX + 1);
    let chatLinkCell = activeSheet.getRange(rowNumber, CHAT_LINK_INDEX + 1);

    if (!eventAttachments) {
      recordingLinkCell.setValue(NOT_FOUND);
      chatLinkCell.setValue(NOT_FOUND);
      noRecordingLinks.push(rowNumber);
      noChatLinks.push(rowNumber);
      continue;
    }
    let videoFound = false;
    let chatFound = false;
    for (let i = 0; i < eventAttachments.length; ++i) {
      let currentAttachment = eventAttachments[i];
      let linkUrl = currentAttachment.fileUrl;
      if (!currentAttachment || !linkUrl) {
        continue;
      }
      if (currentAttachment.mimeType === RECORDING_MIME_TYPE) {
        videoFound = true;
        recordingLinkCell.setValue(linkUrl);
      } else if (currentAttachment.mimeType === CHAT_MIME_TYPE) {
        chatFound = true;
        chatLinkCell.setValue(linkUrl);
      }
    }
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

  let canceledRows = [];
  let invalidMeetings = [];
  for (let rowIndex = 0; rowIndex < oneToOneData.length; ++rowIndex) {
    let rowNumber = rowIndex + DATA_BEGIN_ROW;
    let opsRequestFlag = parseInt(activeSheet.getRange(rowNumber, OPS_REQUEST_INDEX + 1).getValue());
    if (opsRequestFlag !== FLAG_CANCEL_INVITE) {
      // This row doesn't need to be canceled
      continue;
    }

    let meetIdCell = activeSheet.getRange(rowNumber, MEETING_ID_INDEX + 1);
    let meetId = meetIdCell.getValue().trim();
    if (meetId === "") {
      // Meeting is not set
      continue;
    }

    let event = getEvent(calendar, meetId);
    if (!event) {
      continue;
    }
    
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
    let studentResponseCell = activeSheet.getRange(rowNumber, STUDENT_RESPONSE_INDEX + 1);
    studentResponseCell.clearContent();
    let mentorResponseCell = activeSheet.getRange(rowNumber, MENTOR_RESPONSE_INDEX + 1);
    mentorResponseCell.clearContent();
  }
  let summary = `Number of meetings canceled: ${canceledRows.length}
  Number of rows with invalid meeting id: ${invalidMeetings.length} (The calendar event does not exist, or it has already been deleted.)
  ---
  Meetings in the following rows have been canceled.
  `;
  for (let i = 0; i < canceledRows.length; ++i) {
    summary += `\nRow ${canceledRows[i]}`;
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
