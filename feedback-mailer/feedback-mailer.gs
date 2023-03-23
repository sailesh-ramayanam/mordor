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

function afterFormSubmit(e) {
  const info = e.namedValues;
  const { body , score } = createMailBody(info);
  const entryRow = e.range.getRow();
  const workSheet = SpreadsheetApp.getActiveSheet();
  workSheet.getRange(entryRow, SCORE_INDEX + 1).setValue(score);
  workSheet.getRange(entryRow, MAIL_BODY_INDEX + 1).setValue(body);
  const { status , errorMsg } = sendEmail(e.namedValues[FORM_STUDENT_EMAIL][0], body);
  workSheet.getRange(entryRow, MAIL_STATUS_INDEX + 1).setValue(status);
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