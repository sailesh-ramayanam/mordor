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

const INDEX_TIMESTAMP = 0;
const INDEX_EMAIL_ADDRESS = 1;
const INDEX_STUDENT_NAME = 2;
const INDEX_STUDENT_EMAIL = 3;
const INDEX_SSM_EMAIL = 4;
const INDEX_ALLOTTED_PROJECT = 5;
const INDEX_DISCUSSED_PROJECT = 6;
const INDEX_VIDEO_STATUS = 7;
const INDEX_VIDEO_DURATION = 8;
const INDEX_PROJECT_EXPLANATION_SCRIPT = 9;
const INDEX_SPEECH_FLOW = 10;
const INDEX_FEATURES = 11;
const INDEX_TECH_STACK = 12;
const INDEX_SELF_CONTRIBUTION = 13;

const INDEX_SCORE = 14;
const INDEX_MAIL_STATUS = 15;

const MAIL_STATUS_SENT = "SENT";
const MAIL_STATUS_FAILED = "FAILED";

const FORM_TIMESTAMP = "Timestamp";
const FORM_EMAIL_ADDRESS = "Email Address";
const FORM_SSM_EMAIL = "SSM Email ID";
const FORM_STUDENT_NAME = "Student Name";
const FORM_STUDENT_EMAIL = "Student Email ID";

const FORM_ALLOTTED_PROJECT = "Which project is allotted to the student?";
const FORM_DISCUSSED_PROJECT = "Which project did the student discuss?";

const FORM_VIDEO_STATUS = "Did the student switch on the video?";
const FORM_VIDEO_DURATION = "Duration of the video";
const FORM_SPEECH_APPEARANCE = "Do you think student is reading from somewhere (like screen, paper etc.) or memorized?";
const FORM_SPEECH_FLOW = "How is the flow of the speech?";
const FORM_GENERAL_FEATURES = "Rate the student's explanation on project general features";
const FORM_TECH_STACK = "Rate the student's explanation on tech stack";
const FORM_SELF_CONTRIBUTION = "Rate the student's explanation on his/her contribution";

/**
 * IDEAL Answers
 */
const VIDEO_SWITCHED_ON = "Yes";
const VIDEO_DURATION_IDEAL = "1 - 3 min";
const SPEECH_MEMORIZED = "Memorized";
const EXPLANATION_GOOD = "Good";
const IDEAL_SCORE = 7;

/**
 * Other Answers
 */
const VIDEO_SHORT = "< 1 min";
const VIDEO_LONG = "> 3 min";
const EXPLANATION_LONG = "Too detailed";
const EXPLANATION_INSUFFICIENT = "Insufficient";
const EXPLANATION_NOT_TOUCHED = "Did not touch this area";


/**
 * Template Default texts
 */
const OWNER_NAME = "The 10x Academy";
const MAIL_SUBJECT = "Feedback - Project Explanation Video";
const BCC_EMAIL = "operations@the10xacademy.com";

// Following are feedback texts
const FB_VIDEO_OFF = "You have not switched on your video. Please switch on the video and speak to the camera.";
const FB_VIDEO_SHORT = "Your explanation is too short. Please speak for 1 to 3 minutes.";
const FB_VIDEO_LONG = "Your explanation is too long. Please speak for 1 to 3 minutes.";
const FB_SPEECH_NOT_MEMORIZED = "It feels like you are reading from somewhere (like screen, paper etc.). Practice your script multiple times and record only after you feel confident, so that it feels more authentic.";
const FB_SPEECH_STUCK = "We have observed that you are getting stuck often. Write down what you want to speak. Practice it multiple times. You can even recite it to your family members or in front of a mirror. Record after you feel confident about the content.";
const FB_FEATURES_TOO_LONG = "Explanation about general features is too long. Interviewer may not give you that much time. Please restrict it to 2 - 3 lines.";
const FB_FEATURES_INSUFFICIENT = "Explanation about general features is too short. Interviewer may not understand what the project is about. Please speak around 2 - 3 lines.";
const FB_FEATURES_NOT_TOUCHED = "You have not explained what your application does. Please speak 2 - 3 lines on this.";
const FB_TECH_STACK_TOO_LONG = "Explanation about tech stack is too long. Interviewer may not give you that much time. Please restrict it to 2 - 4 lines.";
const FB_TECH_STACK_INSUFFICIENT = "Explanation about tech stack is too short. Interviewer may feel that you are technically weak. Please speak around 2 - 4 lines.";
const FB_TECH_STACK_NOT_TOUCHED = "You have not talked about the tech stack. Please speak 2 - 4 lines on this.";
const FB_CONTRIBUTION_TOO_LONG = "Explanation about your contribution is too long. Interviewer may not give you that much time. Please restrict it to 3 - 6 lines. Focus on your technical strengths.";
const FB_CONTRIBUTION_INSUFFICIENT = "Explanation about your contribution is too short. Interviewer may feel that you are technically weak. Please speak around 3 - 6 lines. Highlight the areas where you are confident.";
const FB_CONTRIBUTION_NOT_TOUCHED = "You have not talked about your contribution. Please speak 3 - 6 lines on this.";

/**
 * Spreadsheet UI
 */

const DATA_BEGIN_ROW = 2;
const DATA_END_COLUMN = "N";

const ERROR_CODE_COULD_NOT_GET_DATA = "10x-error-could-not-get-data";

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
10. ${FORM_SPEECH_APPEARANCE}
11. ${FORM_SPEECH_FLOW}
12. ${FORM_GENERAL_FEATURES}
13. ${FORM_TECH_STACK}
14. ${FORM_SELF_CONTRIBUTION}

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
  workSheet.getRange(1, INDEX_SCORE + 1).setValue(SCORE_HEADER_TITLE);
  workSheet.getRange(1, INDEX_MAIL_STATUS + 1).setValue(MAIL_STATUS_HEADER_TITLE);
  showInfo(TITLE_SUCCESS, MSG_HEADERS_GENERATED);
}

/**
 * Ask the user if they want to use their current email to sent the emails.
 */
function confirmSendingMail(email) {
  let question = `Current email belongs to the following account. Do you want to use this account?

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
    return email.endsWith("@gmail.com") || email.endsWith("@the10xacademy.com");
  } catch (error) {
    return false;
  }
}

function sendEmail(studentEmail, mailBody) {
  let status = "";
  let errorMsg = "";
  if (!validateEmail(studentEmail)) {
    status = MAIL_STATUS_FAILED;
    errorMsg = `Invalid email: ${studentEmail}`;
    return { status , errorMsg };
  }
  try {
    GmailApp.sendEmail(studentEmail, MAIL_SUBJECT, mailBody, { name: OWNER_NAME , htmlBody : mailBody , bcc : BCC_EMAIL } );
    status = MAIL_STATUS_SENT;
  } catch (error) {
    status = MAIL_STATUS_FAILED;
    errorMsg = "Mail sending failed. " + error.toString();
  }
  return { status , errorMsg };
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

    let mailSentStatusCell = activeSheet.getRange(rowNumber, INDEX_MAIL_STATUS + 1);
    if (mailSentStatusCell.getValue().trim() === MAIL_STATUS_SENT) {
      // Mail is already sent
      skippedRows.push(rowNumber);
      continue;
    }

    let studentFormFeedback = {}
    studentFormFeedback[FORM_TIMESTAMP] = feedbackData[rowIndex][INDEX_TIMESTAMP];
    studentFormFeedback[FORM_EMAIL_ADDRESS] = feedbackData[rowIndex][INDEX_EMAIL_ADDRESS];
    studentFormFeedback[FORM_STUDENT_NAME] = feedbackData[rowIndex][INDEX_STUDENT_NAME];
    studentFormFeedback[FORM_STUDENT_EMAIL] = feedbackData[rowIndex][INDEX_STUDENT_EMAIL];
    studentFormFeedback[FORM_SSM_EMAIL] = feedbackData[rowIndex][INDEX_SSM_EMAIL];
    studentFormFeedback[FORM_ALLOTTED_PROJECT] = feedbackData[rowIndex][INDEX_ALLOTTED_PROJECT];
    studentFormFeedback[FORM_DISCUSSED_PROJECT] = feedbackData[rowIndex][INDEX_DISCUSSED_PROJECT];
    studentFormFeedback[FORM_VIDEO_STATUS] = feedbackData[rowIndex][INDEX_VIDEO_STATUS];
    studentFormFeedback[FORM_VIDEO_DURATION] = feedbackData[rowIndex][INDEX_VIDEO_DURATION];
    studentFormFeedback[FORM_SPEECH_APPEARANCE] = feedbackData[rowIndex][INDEX_PROJECT_EXPLANATION_SCRIPT];
    studentFormFeedback[FORM_SPEECH_FLOW] = feedbackData[rowIndex][INDEX_SPEECH_FLOW];
    studentFormFeedback[FORM_GENERAL_FEATURES] = feedbackData[rowIndex][INDEX_FEATURES];
    studentFormFeedback[FORM_TECH_STACK] = feedbackData[rowIndex][INDEX_TECH_STACK];
    studentFormFeedback[FORM_SELF_CONTRIBUTION] = feedbackData[rowIndex][INDEX_SELF_CONTRIBUTION];

    let { mailBody , score } = createMailBody(studentFormFeedback);

    let scoreCell = activeSheet.getRange(rowNumber, INDEX_SCORE + 1);
    scoreCell.setValue(score);
    let { status , errorMsg } = sendEmail(studentFormFeedback[FORM_STUDENT_EMAIL].trim(), mailBody);
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

  if (studentFormFeedback[FORM_VIDEO_STATUS] !== VIDEO_SWITCHED_ON) {
    studentFeedbacksArray.push(FB_VIDEO_OFF);
  }

  switch (studentFormFeedback[FORM_VIDEO_DURATION]) {
    case VIDEO_SHORT:
      studentFeedbacksArray.push(FB_VIDEO_SHORT);
      break;
    case VIDEO_LONG:
      studentFeedbacksArray.push(FB_VIDEO_LONG);
  }

  if (studentFormFeedback[FORM_SPEECH_APPEARANCE] !== SPEECH_MEMORIZED) {
    studentFeedbacksArray.push(FB_SPEECH_NOT_MEMORIZED);
  }

  if (studentFormFeedback[FORM_SPEECH_FLOW] !== EXPLANATION_GOOD) {
    studentFeedbacksArray.push(FB_SPEECH_STUCK);
  }

  switch (studentFormFeedback[FORM_GENERAL_FEATURES]) {
    case EXPLANATION_LONG:
      studentFeedbacksArray.push(FB_FEATURES_TOO_LONG);
      break;
    case EXPLANATION_INSUFFICIENT:
      studentFeedbacksArray.push(FB_FEATURES_INSUFFICIENT);
      break;
    case EXPLANATION_NOT_TOUCHED:
      studentFeedbacksArray.push(FB_FEATURES_NOT_TOUCHED);
  }

  switch (studentFormFeedback[FORM_TECH_STACK]) {
    case EXPLANATION_LONG:
      studentFeedbacksArray.push(FB_TECH_STACK_TOO_LONG);
      break;
    case EXPLANATION_INSUFFICIENT:
      studentFeedbacksArray.push(FB_TECH_STACK_INSUFFICIENT);
      break;
    case EXPLANATION_NOT_TOUCHED:
      studentFeedbacksArray.push(FB_TECH_STACK_NOT_TOUCHED);
  }

  switch (studentFormFeedback[FORM_SELF_CONTRIBUTION]) {
    case EXPLANATION_LONG:
      studentFeedbacksArray.push(FB_CONTRIBUTION_TOO_LONG);
      break;
    case EXPLANATION_INSUFFICIENT:
      studentFeedbacksArray.push(FB_CONTRIBUTION_INSUFFICIENT);
      break;
    case EXPLANATION_NOT_TOUCHED:
      studentFeedbacksArray.push(FB_CONTRIBUTION_NOT_TOUCHED);
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
  
  let feedbackLines = "";
  for (let fbIndex = 0; fbIndex < studentFeedbacksArray.length; fbIndex++) {
    feedbackLines += `<li> ${studentFeedbacksArray[fbIndex]} </li>`;
  }

  mailBody = `
    Dear <b>${studentFormFeedback[FORM_STUDENT_NAME]}</b>,<br />
    <br />
    The project explanation video submitted by you does not meet expectations. Detailed feedback is provided below.<br />
    
    <ol>
    ${feedbackLines}
    </ol>

    Please create the video again, improving the point(s) discussed above and re-submit the updated video.<br />
    <br />
    <b>Thanks and Regards,<br />
    10x Academy</b>.
    `
  ;

  return { mailBody , score };
}