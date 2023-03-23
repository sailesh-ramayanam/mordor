// Whenever you publish a new version of the add-on, you must increment the below version number. This helps for book-keeping and comparing with previous versions.
const LAST_PUBLISHED_VERSION = 0;

/*
Current namedValues object is
nameValues = {
  'Timestamp': ['6/7/2015 20:54:13'], // Form Submitted time
  'Email Address' : ["ssm@the10xacademy.com"], // Auto Captured Email of form submitter
  'Student Name' : ["Student A"],
  'Student Email ID' : ["student_a@gmail.com"],
  'SSM Email ID' : ["ssm@the10xacademy.com"],
  'Which project is allotted to the student?' : ["Contacts manager"],
  'Which project did the student discuss?' : ["Contacts manager"],
  'Did the student switch on the video?' : ["Yes"],
  'Duration of the video' : ["1 - 3 min"],
  'Do you think student is reading from somewhere (like screen, paper etc.) or memorized?' : ["Memorized"],
  'How is the flow of the speech?' : ["Good"],
  'Rate the student's explanation on project general features' : ["Good"],
  'Rate the student's explanation on tech stack' : ["Good"],
  'Rate the student's explanation on his/her contribution' : ["Good"]
}
*/

/**
 * These are 0-based indices. Do not confuse them with column numbers.
 * First column will have index as 0.
 * 0 - 13 indices are covered by Form Questions
 * If you add/remove/modify columns, you must also change below items and MSG_HELP
 */

const TIMESTAMP_INDEX = 0;
const EMAIL_ADDRESS_INDEX = 1;
const STUDENT_NAME_INDEX = 2;
const STUDENT_EMAIL_INDEX = 3;
const SSM_EMAIL_INDEX = 4;
const ALLOTTED_PROJECT_INDEX = 5;
const DISCUSSED_PROJECT_INDEX = 6;
const VIDEO_STATUS_INDEX = 7;
const VIDEO_DURATION_INDEX = 8;
const PROJECT_EXPLANATION_SCRIPT_INDEX = 9;
const SPEECH_FLOW_INDEX = 10;
const GENERAL_FEATURES_EXPLANATION_INDEX = 11;
const TECH_STACK_EXPLANATION_INDEX = 12;
const CONTRIBUTION_EXPLANATION_INDEX = 13;

const SCORE_INDEX = 14;
const MAIL_STATUS_INDEX = 15;
const MAIL_BODY_INDEX = 16;

const MAIL_STATUS_SENT = "SENT";
const MAIL_STATUS_FAILED = "FAILED";

const FORM_STUDENT_NAME = "Student Name";
const FORM_STUDENT_EMAIL = "Student Email ID";

const FORM_ALLOTTED_PROJECT = "Which project is allotted to the student?";
const FORM_DISCUSSED_PROJECT = "Which project did the student discuss?";

const FORM_VIDEO_STATUS = "Did the student switch on the video?";
const FORM_VIDEO_DURATION = "Duration of the video";
const FORM_PROJECT_EXPLANATION_SCRIPT = "Do you think student is reading from somewhere (like screen, paper etc.) or memorized?";
const FORM_SPEECH_FLOW = "How is the flow of the speech?";
const FORM_GENERAL_FEATURES_EXPLANATION = "Rate the student's explanation on project general features";
const FORM_TECH_STACK_EXPLANATION = "Rate the student's explanation on tech stack";
const FORM_CONTRIBUTION_EXPLANATION = "Rate the student's explanation on his/her contribution";

/**
 * IDEAL Answers
 */
const VIDEO_STATUS_IDEAL_ANSWER = "Yes";
const VIDEO_DURATION_IDEAL_ANSWER = "1 - 3 min";
const PROJECT_EXPLANATION_SCRIPT_IDEAL_ANSWER = "Memorized";
const EXPLANATION_IDEAL_ANSWER = "Good";
const IDEAL_SCORE = 7;

/**
 * Other Answers
 */
const VIDEO_DURATION_ANSWER_SHORT = "< 1 min";
const VIDEO_DURATION_ANSWER_LONG = "> 3 min";
const EXPLANATION_ANSWER_DETAILED = "Too detailed";
const EXPLANATION_ANSWER_INSUFFICIENT = "Insufficient";
const EXPLANATION_ANSWER_NO_TOUCH = "Did not touch this area";


/**
 * Template Default texts
 */
const OWNER_NAME = "The 10x Academy";
const MAIL_SUBJECT = "Feedback - Project Explanation Video";
const MAIL_GREETING = "Dear ";

const MAIL_BODY_MAIN_HEADING_IDEAL_TEXT = "The project explanation video submitted by you meets expectations. Please provide similar explanation in interviews. Wish you the best.";
const MAIL_BODY_MAIN_HEADING_NON_IDEAL_TEXT = "The project explanation video submitted by you does not meet expectations. Detailed feedback is provided below.";

const MAIL_SIGNATURE = `
Thanks and Regards
10x Academy.`;

const LINE_BREAK = "\n";

const PROJECT_MISMATCH_PRE_TEXT = "The project allocated to you was ";
const PROJECT_MISMATCH_MID_TEXT = ". But you discussed ";
const PROJECT_MISMATCH_POST_TEXT = ". Please discuss the correct project.";
const VIDEO_STATUS_NON_IDEAL_TEXT = ". You have not switched on your video. Please switch on the video and speak to the camera.";
const VIDEO_DURATION_NON_IDEAL_SHORT_TEXT = ". Your explanation is too short. Please speak for 1 to 3 minutes.";
const VIDEO_DURATION_NON_IDEAL_LONG_TEXT = ". Your explanation is too long. Please speak for 1 to 3 minutes.";
const PROJECT_EXPLANATION_SCRIPT_NON_IDEAL_TEXT = ". It feels like you are reading from somewhere (like screen, paper etc.). Practice your script multiple times and record only after you feel confident, so that it feels more authentic.";
const SPEECH_FLOW_NON_IDEAL_TEXT = ". We have observed that you are getting stuck often. Write down what you want to speak. Practice it multiple times. You can even recite it to your family members or in front of a mirror. Record after you feel confident about the content.";
const GENERAL_FEATURES_EXPLANATION_NON_IDEAL_TOO_LONG_TEXT = ". Explanation about general features is too long. Interviewer may not give you that much time. Please restrict it to 2 - 3 lines.";
const GENERAL_FEATURES_EXPLANATION_NON_IDEAL_INSUFFICIENT_TEXT = ". Explanation about general features is too short. Interviewer may not understand what the project is about. Please speak around 2 - 3 lines.";
const GENERAL_FEATURES_EXPLANATION_NON_IDEAL_NO_TOUCH_TEXT = ". You have not explained what your application does. Please speak 2 - 3 lines on this.";
const TECH_STACK_EXPLANATION_NON_IDEAL_TOO_LONG_TEXT = ". Explanation about tech stack is too long. Interviewer may not give you that much time. Please restrict it to 2 - 4 lines.";
const TECH_STACK_EXPLANATION_NON_IDEAL_INSUFFICIENT_TEXT = ". Explanation about tech stack is too short. Interviewer may feel that you are technically weak. Please speak around 2 - 4 lines.";
const TECH_STACK_EXPLANATION_NON_IDEAL_NO_TOUCH_TEXT = ". You have not talked about the tech stack. Please speak 2 - 4 lines on this.";
const CONTRIBUTION_EXPLANATION_NON_IDEAL_TOO_LONG_TEXT = ". Explanation about your contribution is too long. Interviewer may not give you that much time. Please restrict it to 3 - 6 lines. Focus on your technical strengths.";
const CONTRIBUTION_EXPLANATION_NON_IDEAL_INSUFFICIENT_TEXT = ". Explanation about your contribution is too short. Interviewer may feel that you are technically weak. Please speak around 3 - 6 lines. Highlight the areas where you are confident.";
const CONTRIBUTION_EXPLANATION_NON_IDEAL_NO_TOUCH_TEXT = ". You have not talked about your contribution. Please speak 3 - 6 lines on this.";

const PARTIAL_SCORE_TEXT = "Please create the video again, improving the point(s) discussed above and re-submit the updated video.";

/**
 * Spreadsheet UI
 */

const DATA_BEGIN_ROW = 2;
const DATA_END_COLUMN = "N";

const ERROR_CODE_COULD_NOT_GET_DATA = "10x-error-could-not-get-data";

const FORM_TIMESTAMP = "Timestamp";
const FORM_EMAIL_ADDRESS = "Email Address";
const FORM_SSM_EMAIL = "SSM Email ID";

const MENU_TITLE = "10x Feedback Mailer";
const MENU_HELP = "Help";

const SCORE_HEADER_TITLE = "Score";
const MAIL_STATUS_HEADER_TITLE = "Mail Status";
const MAIL_BODY_HEADER_TITLE = "Mail Body Sent";
const GENERATE_HEADERS = "Generate Headers";
const SEND_MAILS = "Send Mails to Students";
const MSG_MAIL_REJECTED = "You chose not to use the current email for sending mails to students.";
const MSG_HEADERS_GENERATED = "Headers generated!";
const MAIL_SENDING_CONFIRMATION_QUESTION = "Current email belongs to the following account. Do you want to use this account?"

const TITLE_ERROR = "Error";
const TITLE_SUCCESS = "Success";
const TITLE_HELP = "Help";

const MSG_HELP = `
${MENU_TITLE}
Version: ${LAST_PUBLISHED_VERSION + 1}
This extension helps in sending feedback emails to the students.
- Your response (current) sheet should be linked to Feedback form.
- (1 to 14) columns of response sheet are covered by Form Questions.

1. ${FORM_TIMESTAMP} - Auto generated by Form.
2. ${FORM_EMAIL_ADDRESS} - "Collect email addresses" should be enabled in Form "Responses" settings .

- Your Form must have the following Questions in the order (Case & Format Sensitive).

3. ${FORM_STUDENT_NAME}
4. ${FORM_STUDENT_EMAIL}
5. ${FORM_SSM_EMAIL}
6. ${FORM_ALLOTTED_PROJECT}
7. ${FORM_DISCUSSED_PROJECT}
8. ${FORM_VIDEO_STATUS}
9. ${FORM_VIDEO_DURATION}
10. ${FORM_PROJECT_EXPLANATION_SCRIPT}
11. ${FORM_SPEECH_FLOW}
12. ${FORM_GENERAL_FEATURES_EXPLANATION}
13. ${FORM_TECH_STACK_EXPLANATION}
14. ${FORM_CONTRIBUTION_EXPLANATION}

- Using "Generate Headers" Button, following columns will be generated:

15. ${SCORE_HEADER_TITLE} - Ideal Score is 7 (Auto Generated using "${SEND_MAILS}" Button)
16. ${MAIL_STATUS_HEADER_TITLE} - Mail Status (SENT or FAILED) (Auto Generated using "${SEND_MAILS}" Button)
17. ${MAIL_BODY_HEADER_TITLE} - Mail bosy sent to students (Auto Generated using "${SEND_MAILS}" Button)
`;

function onInstall(e) {
  // Without this, user will have to reload the sheet after installing.
  onOpen(e);
}

function onOpen(e) {
  SpreadsheetApp.getUi()
  .createMenu(MENU_TITLE)
  .addItem(SEND_MAILS, "sendMailsToStudents")
  .addItem(GENERATE_HEADERS, "generateHeaders")
  .addItem(MENU_HELP, "help")
  .addToUi();
}

function help() {
  showInfo(TITLE_HELP, MSG_HELP);
}

function showInfo(title, body) {
  let ui = SpreadsheetApp.getUi();
  ui.alert(title, body, ui.ButtonSet.OK);
}

function logBeinDelimiter() {
  console.log("BEGIN: 10x---------------10x");
}

function logEndDelimiter() {
  console.log("END: 10x---------------10x");
}

function generateHeaders() {
  const workSheet = SpreadsheetApp.getActiveSheet();
  workSheet.getRange(1, SCORE_INDEX + 1).setValue(SCORE_HEADER_TITLE);
  workSheet.getRange(1, MAIL_STATUS_INDEX + 1).setValue(MAIL_STATUS_HEADER_TITLE);
  workSheet.getRange(1, MAIL_BODY_INDEX + 1).setValue(MAIL_BODY_HEADER_TITLE);
  showInfo(TITLE_SUCCESS, MSG_HEADERS_GENERATED);
}

/**
 * Ask the user if they want to use their current email to sent the emails.
 */
function confirmSendingMail(email) {

  let emailOwner = email.substring(0, email.lastIndexOf("@"));
  let question = `${MAIL_SENDING_CONFIRMATION_QUESTION}

  Name: ${emailOwner}
  Id: ${email}
  `;

  let ui = SpreadsheetApp.getUi();
  let userResponse = ui.alert("Please confirm", question, ui.ButtonSet.YES_NO);
  if (userResponse === ui.Button.YES) {
    return true;
  }
  logBeinDelimiter();
  console.log(MSG_MAIL_REJECTED);
  logEndDelimiter();
  return false;
}

function initSetup() {
  let result = {
    success: false,
    activeSheet: null,
    feedbackData: null
  };

  result.activeSheet = SpreadsheetApp.getActiveSheet();
  result.feedbackData = sheetData(result.activeSheet);
  if (!result.feedbackData) {
    showInfo(TITLE_ERROR, ERROR_CODE_COULD_NOT_GET_DATA);
    return result;
  }
  result.success = true;
  return result;
}

function sheetData(sheet) {
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

function validateEmail(email) {
  try {
    return email.endsWith("@gmail.com") || email.endsWith("@the10xacademy.com") || email.endsWith("@yahoo.com");
  } catch (error) {
    return false;
  }
}

function sendEmail(studentEmail, textBody) {
  let status = "";
  let errorMsg = "";
  if (!validateEmail(studentEmail)) {
    status = MAIL_STATUS_FAILED;
    errorMsg = "Invalid email: " + studentEmail;
    return { status , errorMsg };
  }
  try {
    let htmlRawText = textBody.replaceAll(LINE_BREAK,`<br>`);
    let htmlFile = HtmlService.createTemplate(`<p> ${htmlRawText} </p>`);
    let htmlText = htmlFile.evaluate().getContent();
    GmailApp.sendEmail(studentEmail, MAIL_SUBJECT, textBody, { name: OWNER_NAME , htmlBody : htmlText } );
    status = MAIL_STATUS_SENT;
    return { status , errorMsg };
  } catch (error) {
    status = MAIL_STATUS_FAILED;
    errorMsg = "Mail sending failed. " + error.toString();
    return { status , errorMsg };
  }
}

function sendMailsToStudents() {

  let setupResult = initSetup();
  if (!setupResult.success) {
    return;
  }

  if (!confirmSendingMail(Session.getActiveUser().getEmail())) {
    showInfo(TITLE_ERROR, MSG_MAIL_REJECTED);
    return;
  }

  let { activeSheet, feedbackData } = setupResult;

  let errorRows = [];
  let skippedRows = [];
  let successRows = [];

  for (let rowIndex = 0; rowIndex < feedbackData.length; ++rowIndex) {
    let rowNumber = rowIndex + DATA_BEGIN_ROW;

    let mailSentStatusCell = activeSheet.getRange(rowNumber, MAIL_STATUS_INDEX + 1);
    if (mailSentStatusCell.getValue().trim() === MAIL_STATUS_SENT) {
      // Mail is already sent
      skippedRows.push(rowNumber);
      continue;
    }

    let studentFeedback = {}
    studentFeedback[FORM_TIMESTAMP] = [feedbackData[rowIndex][TIMESTAMP_INDEX]];
    studentFeedback[FORM_EMAIL_ADDRESS] = [feedbackData[rowIndex][EMAIL_ADDRESS_INDEX]];
    studentFeedback[FORM_STUDENT_NAME] = [feedbackData[rowIndex][STUDENT_NAME_INDEX]];
    studentFeedback[FORM_STUDENT_EMAIL] = [feedbackData[rowIndex][STUDENT_EMAIL_INDEX]];
    studentFeedback[FORM_SSM_EMAIL] = [feedbackData[rowIndex][SSM_EMAIL_INDEX]];
    studentFeedback[FORM_ALLOTTED_PROJECT] = [feedbackData[rowIndex][ALLOTTED_PROJECT_INDEX]];
    studentFeedback[FORM_DISCUSSED_PROJECT] = [feedbackData[rowIndex][DISCUSSED_PROJECT_INDEX]];
    studentFeedback[FORM_VIDEO_STATUS] = [feedbackData[rowIndex][VIDEO_STATUS_INDEX]];
    studentFeedback[FORM_VIDEO_DURATION] = [feedbackData[rowIndex][VIDEO_DURATION_INDEX]];
    studentFeedback[FORM_PROJECT_EXPLANATION_SCRIPT] = [feedbackData[rowIndex][PROJECT_EXPLANATION_SCRIPT_INDEX]];
    studentFeedback[FORM_SPEECH_FLOW] = [feedbackData[rowIndex][SPEECH_FLOW_INDEX]];
    studentFeedback[FORM_GENERAL_FEATURES_EXPLANATION] = [feedbackData[rowIndex][GENERAL_FEATURES_EXPLANATION_INDEX]];
    studentFeedback[FORM_TECH_STACK_EXPLANATION] = [feedbackData[rowIndex][TECH_STACK_EXPLANATION_INDEX]];
    studentFeedback[FORM_CONTRIBUTION_EXPLANATION] = [feedbackData[rowIndex][CONTRIBUTION_EXPLANATION_INDEX]];

    let { body , score } = createMailBody(studentFeedback);

    let scoreCell = activeSheet.getRange(rowNumber, SCORE_INDEX + 1);
    let mailBodyCell = activeSheet.getRange(rowNumber, MAIL_BODY_INDEX + 1);
    scoreCell.setValue(score);
    mailBodyCell.setValue(body);
    let { status , errorMsg } = sendEmail(studentFeedback[FORM_STUDENT_EMAIL][0], body);
    mailSentStatusCell.setValue(status);

    if (status === MAIL_STATUS_FAILED) {
      // Mail sending failure
      errorRows.push({rowNumber: rowNumber, error: errorMsg});
      continue;
    }
    successRows.push(rowNumber);
  }

  let summary = `Number of Mails Sent: ${successRows.length}
  Number of skipped rows: ${skippedRows.length} (Mail was already sent previously)
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

function createMailBody(info) {

  let score = IDEAL_SCORE;
  let feedbackCount = 0; // for Tracking number of feedbacks
  
  let body = MAIL_GREETING + info[FORM_STUDENT_NAME][0] + LINE_BREAK + LINE_BREAK;

  body += MAIL_BODY_MAIN_HEADING_NON_IDEAL_TEXT + LINE_BREAK + LINE_BREAK;

  if (info[FORM_ALLOTTED_PROJECT][0] !== info[FORM_DISCUSSED_PROJECT][0]) {
    body += PROJECT_MISMATCH_PRE_TEXT + info[FORM_ALLOTTED_PROJECT][0] + PROJECT_MISMATCH_MID_TEXT + info[FORM_DISCUSSED_PROJECT][0] + PROJECT_MISMATCH_POST_TEXT + LINE_BREAK;
    body += LINE_BREAK + PARTIAL_SCORE_TEXT + LINE_BREAK;
    body += MAIL_SIGNATURE
    score = 0;
    return { body , score };
  }

  if (info[FORM_VIDEO_STATUS][0] !== VIDEO_STATUS_IDEAL_ANSWER) {
    feedbackCount += 1;
    body += feedbackCount + VIDEO_STATUS_NON_IDEAL_TEXT + LINE_BREAK;
  }

  switch (info[FORM_VIDEO_DURATION][0]) {
    case VIDEO_DURATION_ANSWER_SHORT:
      feedbackCount += 1;
      body += feedbackCount + VIDEO_DURATION_NON_IDEAL_SHORT_TEXT + LINE_BREAK;
      break;
    case VIDEO_DURATION_ANSWER_LONG:
      feedbackCount += 1;
      body += feedbackCount + VIDEO_DURATION_NON_IDEAL_LONG_TEXT + LINE_BREAK;
  }

  if (info[FORM_PROJECT_EXPLANATION_SCRIPT][0] !== PROJECT_EXPLANATION_SCRIPT_IDEAL_ANSWER) {
    feedbackCount += 1;
    body += feedbackCount + PROJECT_EXPLANATION_SCRIPT_NON_IDEAL_TEXT + LINE_BREAK;
  }

  if (info[FORM_SPEECH_FLOW][0] !== EXPLANATION_IDEAL_ANSWER) {
    feedbackCount += 1;
    body += feedbackCount + SPEECH_FLOW_NON_IDEAL_TEXT + LINE_BREAK;
  }

  switch (info[FORM_GENERAL_FEATURES_EXPLANATION][0]) {
    case EXPLANATION_ANSWER_DETAILED:
      feedbackCount += 1;
      body += feedbackCount + GENERAL_FEATURES_EXPLANATION_NON_IDEAL_TOO_LONG_TEXT + LINE_BREAK;
      break;
    case EXPLANATION_ANSWER_INSUFFICIENT:
      feedbackCount += 1;
      body += feedbackCount + GENERAL_FEATURES_EXPLANATION_NON_IDEAL_INSUFFICIENT_TEXT + LINE_BREAK;
      break;
    case EXPLANATION_ANSWER_NO_TOUCH:
      feedbackCount += 1;
      body += feedbackCount + GENERAL_FEATURES_EXPLANATION_NON_IDEAL_NO_TOUCH_TEXT + LINE_BREAK;
  }

  switch (info[FORM_TECH_STACK_EXPLANATION][0]) {
    case EXPLANATION_ANSWER_DETAILED:
      feedbackCount += 1;
      body += feedbackCount + TECH_STACK_EXPLANATION_NON_IDEAL_TOO_LONG_TEXT + LINE_BREAK;
      break;
    case EXPLANATION_ANSWER_INSUFFICIENT:
      feedbackCount += 1;
      body += feedbackCount + TECH_STACK_EXPLANATION_NON_IDEAL_INSUFFICIENT_TEXT + LINE_BREAK;
      break;
    case EXPLANATION_ANSWER_NO_TOUCH:
      feedbackCount += 1;
      body += feedbackCount + TECH_STACK_EXPLANATION_NON_IDEAL_NO_TOUCH_TEXT + LINE_BREAK;
  }

  switch (info[FORM_CONTRIBUTION_EXPLANATION][0]) {
    case EXPLANATION_ANSWER_DETAILED:
      feedbackCount += 1;
      body += feedbackCount + CONTRIBUTION_EXPLANATION_NON_IDEAL_TOO_LONG_TEXT + LINE_BREAK;
      break;
    case EXPLANATION_ANSWER_INSUFFICIENT:
      feedbackCount += 1;
      body += feedbackCount + CONTRIBUTION_EXPLANATION_NON_IDEAL_INSUFFICIENT_TEXT + LINE_BREAK;
      break;
    case EXPLANATION_ANSWER_NO_TOUCH:
      feedbackCount += 1;
      body += feedbackCount + CONTRIBUTION_EXPLANATION_NON_IDEAL_NO_TOUCH_TEXT + LINE_BREAK;
  }

  score -= feedbackCount;

  if (score === IDEAL_SCORE) {
    body = body.replace(MAIL_BODY_MAIN_HEADING_NON_IDEAL_TEXT,MAIL_BODY_MAIN_HEADING_IDEAL_TEXT);
  } else {
    body += LINE_BREAK + PARTIAL_SCORE_TEXT + LINE_BREAK;
  }

  body += MAIL_SIGNATURE;

  return { body , score };
}