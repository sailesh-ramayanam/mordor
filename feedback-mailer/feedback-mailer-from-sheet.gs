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
const VIDEO_STATUS_SWITCHED_ON = "Yes";
const VIDEO_DURATION_IDEAL = "1 - 3 min";
const PROJECT_EXPLANATION_SCRIPT_MEMORIZED = "Memorized";
const EXPLANATION_GOOD = "Good";
const IDEAL_SCORE = 7;

/**
 * Other Answers
 */
const VIDEO_DURATION_SHORT = "< 1 min";
const VIDEO_DURATION_TOO_LONG = "> 3 min";
const EXPLANATION_TOO_LONG = "Too detailed";
const EXPLANATION_INSUFFICIENT = "Insufficient";
const EXPLANATION_NOT_TOUCHED = "Did not touch this area";


/**
 * Template Default texts
 */
const OWNER_NAME = "The 10x Academy";
const MAIL_SUBJECT = "Feedback - Project Explanation Video";

const VIDEO_STATUS_SWITCH_ON_TEXT = "You have not switched on your video. Please switch on the video and speak to the camera.";
const VIDEO_DURATION_SHORT_TEXT = "Your explanation is too short. Please speak for 1 to 3 minutes.";
const VIDEO_DURATION_TOO_LONG_TEXT = "Your explanation is too long. Please speak for 1 to 3 minutes.";
const PROJECT_EXPLANATION_SCRIPT_NOT_MEMORIZED_TEXT = "It feels like you are reading from somewhere (like screen, paper etc.). Practice your script multiple times and record only after you feel confident, so that it feels more authentic.";
const SPEECH_FLOW_NOT_GOOD_TEXT = "We have observed that you are getting stuck often. Write down what you want to speak. Practice it multiple times. You can even recite it to your family members or in front of a mirror. Record after you feel confident about the content.";
const GENERAL_FEATURES_EXPLANATION_TOO_LONG_TEXT = "Explanation about general features is too long. Interviewer may not give you that much time. Please restrict it to 2 - 3 lines.";
const GENERAL_FEATURES_EXPLANATION_INSUFFICIENT_TEXT = "Explanation about general features is too short. Interviewer may not understand what the project is about. Please speak around 2 - 3 lines.";
const GENERAL_FEATURES_EXPLANATION_NOT_TOUCHED_TEXT = "You have not explained what your application does. Please speak 2 - 3 lines on this.";
const TECH_STACK_EXPLANATION_TOO_LONG_TEXT = "Explanation about tech stack is too long. Interviewer may not give you that much time. Please restrict it to 2 - 4 lines.";
const TECH_STACK_EXPLANATION_INSUFFICIENT_TEXT = "Explanation about tech stack is too short. Interviewer may feel that you are technically weak. Please speak around 2 - 4 lines.";
const TECH_STACK_EXPLANATION_NOT_TOUCHED_TEXT = "You have not talked about the tech stack. Please speak 2 - 4 lines on this.";
const CONTRIBUTION_EXPLANATION_TOO_LONG_TEXT = "Explanation about your contribution is too long. Interviewer may not give you that much time. Please restrict it to 3 - 6 lines. Focus on your technical strengths.";
const CONTRIBUTION_EXPLANATION_INSUFFICIENT_TEXT = "Explanation about your contribution is too short. Interviewer may feel that you are technically weak. Please speak around 3 - 6 lines. Highlight the areas where you are confident.";
const CONTRIBUTION_EXPLANATION_NOT_TOUCHED_TEXT = "You have not talked about your contribution. Please speak 3 - 6 lines on this.";

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

function logBeginDelimiter() {
  console.log("BEGIN: 10x---------------10x");
}

function logEndDelimiter() {
  console.log("END: 10x---------------10x");
}

function generateHeaders() {
  const workSheet = SpreadsheetApp.getActiveSheet();
  workSheet.getRange(1, SCORE_INDEX + 1).setValue(SCORE_HEADER_TITLE);
  workSheet.getRange(1, MAIL_STATUS_INDEX + 1).setValue(MAIL_STATUS_HEADER_TITLE);
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
  logBeginDelimiter();
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
    logBeginDelimiter();
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

function sendEmail(studentEmail, mailBody) {
  let status = "";
  let errorMsg = "";
  if (!validateEmail(studentEmail)) {
    status = MAIL_STATUS_FAILED;
    errorMsg = "Invalid email: " + studentEmail;
    return { status , errorMsg };
  }
  try {
    let htmlFile = HtmlService.createTemplate(mailBody);
    let htmlText = htmlFile.evaluate().getContent();
    GmailApp.sendEmail(studentEmail, MAIL_SUBJECT, mailBody, { name: OWNER_NAME , htmlBody : htmlText } );
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

    let studentFormFeedback = {}
    studentFormFeedback[FORM_TIMESTAMP] = feedbackData[rowIndex][TIMESTAMP_INDEX];
    studentFormFeedback[FORM_EMAIL_ADDRESS] = feedbackData[rowIndex][EMAIL_ADDRESS_INDEX];
    studentFormFeedback[FORM_STUDENT_NAME] = feedbackData[rowIndex][STUDENT_NAME_INDEX];
    studentFormFeedback[FORM_STUDENT_EMAIL] = feedbackData[rowIndex][STUDENT_EMAIL_INDEX];
    studentFormFeedback[FORM_SSM_EMAIL] = feedbackData[rowIndex][SSM_EMAIL_INDEX];
    studentFormFeedback[FORM_ALLOTTED_PROJECT] = feedbackData[rowIndex][ALLOTTED_PROJECT_INDEX];
    studentFormFeedback[FORM_DISCUSSED_PROJECT] = feedbackData[rowIndex][DISCUSSED_PROJECT_INDEX];
    studentFormFeedback[FORM_VIDEO_STATUS] = feedbackData[rowIndex][VIDEO_STATUS_INDEX];
    studentFormFeedback[FORM_VIDEO_DURATION] = feedbackData[rowIndex][VIDEO_DURATION_INDEX];
    studentFormFeedback[FORM_PROJECT_EXPLANATION_SCRIPT] = feedbackData[rowIndex][PROJECT_EXPLANATION_SCRIPT_INDEX];
    studentFormFeedback[FORM_SPEECH_FLOW] = feedbackData[rowIndex][SPEECH_FLOW_INDEX];
    studentFormFeedback[FORM_GENERAL_FEATURES_EXPLANATION] = feedbackData[rowIndex][GENERAL_FEATURES_EXPLANATION_INDEX];
    studentFormFeedback[FORM_TECH_STACK_EXPLANATION] = feedbackData[rowIndex][TECH_STACK_EXPLANATION_INDEX];
    studentFormFeedback[FORM_CONTRIBUTION_EXPLANATION] = feedbackData[rowIndex][CONTRIBUTION_EXPLANATION_INDEX];

    let { mailBody , score } = createMailBody(studentFormFeedback);

    let scoreCell = activeSheet.getRange(rowNumber, SCORE_INDEX + 1);
    scoreCell.setValue(score);
    let { status , errorMsg } = sendEmail(studentFormFeedback[FORM_STUDENT_EMAIL], mailBody);
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

function generateStudentFeedbacksArray(studentFormFeedback) {
  let studentFeedbacksArray = [];

  if (studentFormFeedback[FORM_ALLOTTED_PROJECT] !== studentFormFeedback[FORM_DISCUSSED_PROJECT]) {
    let feedback = `The project allocated to you was <b>${studentFormFeedback[FORM_ALLOTTED_PROJECT]}</b>. But you discussed <b>${studentFormFeedback[FORM_DISCUSSED_PROJECT]}</b>. Please discuss the correct project.`;

    studentFeedbacksArray.push(feedback);

    let score = 0;

    return { studentFeedbacksArray, score };
  }

  if (studentFormFeedback[FORM_VIDEO_STATUS] !== VIDEO_STATUS_SWITCHED_ON) {
    studentFeedbacksArray.push(VIDEO_STATUS_SWITCH_ON_TEXT);
  }

  switch (studentFormFeedback[FORM_VIDEO_DURATION]) {
    case VIDEO_DURATION_SHORT:
      studentFeedbacksArray.push(VIDEO_DURATION_SHORT_TEXT);
      break;
    case VIDEO_DURATION_TOO_LONG:
      studentFeedbacksArray.push(VIDEO_DURATION_TOO_LONG_TEXT);
  }

  if (studentFormFeedback[FORM_PROJECT_EXPLANATION_SCRIPT] !== PROJECT_EXPLANATION_SCRIPT_MEMORIZED) {
    studentFeedbacksArray.push(PROJECT_EXPLANATION_SCRIPT_NOT_MEMORIZED_TEXT);
  }

  if (studentFormFeedback[FORM_SPEECH_FLOW] !== EXPLANATION_GOOD) {
    studentFeedbacksArray.push(SPEECH_FLOW_NOT_GOOD_TEXT);
  }

  switch (studentFormFeedback[FORM_GENERAL_FEATURES_EXPLANATION]) {
    case EXPLANATION_TOO_LONG:
      studentFeedbacksArray.push(GENERAL_FEATURES_EXPLANATION_TOO_LONG_TEXT);
      break;
    case EXPLANATION_INSUFFICIENT:
      studentFeedbacksArray.push(GENERAL_FEATURES_EXPLANATION_INSUFFICIENT_TEXT);
      break;
    case EXPLANATION_NOT_TOUCHED:
      studentFeedbacksArray.push(GENERAL_FEATURES_EXPLANATION_NOT_TOUCHED_TEXT);
  }

  switch (studentFormFeedback[FORM_TECH_STACK_EXPLANATION]) {
    case EXPLANATION_TOO_LONG:
      studentFeedbacksArray.push(TECH_STACK_EXPLANATION_TOO_LONG_TEXT);
      break;
    case EXPLANATION_INSUFFICIENT:
      studentFeedbacksArray.push(TECH_STACK_EXPLANATION_INSUFFICIENT_TEXT);
      break;
    case EXPLANATION_NOT_TOUCHED:
      studentFeedbacksArray.push(TECH_STACK_EXPLANATION_NOT_TOUCHED_TEXT);
  }

  switch (studentFormFeedback[FORM_CONTRIBUTION_EXPLANATION]) {
    case EXPLANATION_TOO_LONG:
      studentFeedbacksArray.push(CONTRIBUTION_EXPLANATION_TOO_LONG_TEXT);
      break;
    case EXPLANATION_INSUFFICIENT:
      studentFeedbacksArray.push(CONTRIBUTION_EXPLANATION_INSUFFICIENT_TEXT);
      break;
    case EXPLANATION_NOT_TOUCHED:
      studentFeedbacksArray.push(CONTRIBUTION_EXPLANATION_NOT_TOUCHED_TEXT);
  }

  let score = IDEAL_SCORE - studentFeedbacksArray.length;

  return { studentFeedbacksArray, score }

}

function createMailBody(studentFormFeedback) {

  let { studentFeedbacksArray, score } = generateStudentFeedbacksArray(studentFormFeedback)

  let mailBody = ``;

  if (studentFeedbacksArray.length === 0) {
    mailBody = `
      Dear <b>${studentFormFeedback[FORM_STUDENT_NAME]}</b>,<br />
      <br />
      The project explanation video submitted by you <b>meets expectations</b>. Please provide similar explanation in interviews. Wish you the best.<br />
      <br />
      <b>Thanks and Regards,<br />
      10x Academy</b>.
      `
    ;

    return { mailBody , score };
  }

  mailBody = `
    Dear <b>${studentFormFeedback[FORM_STUDENT_NAME]}</b>,<br />
    <br />
    The project explanation video submitted by you does not meet expectations. Detailed feedback is provided below.<br />
    
    <ol>`
  ;
  
  for (let feedback = 0; feedback <studentFeedbacksArray.length; feedback++) {
    mailBody += `<li> ${studentFeedbacksArray[feedback]} </li>`;
  }

  mailBody += `</ol>

    Please create the video again, improving the point(s) discussed above and re-submit the updated video.<br />
    <br />
    <b>Thanks and Regards,<br />
    10x Academy</b>.
    `
  ;

  return { mailBody , score };
}