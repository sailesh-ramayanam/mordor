// Whenever you publish a new version of the add-on, you must increment the below version number. This helps for book-keeping and comparing with previous versions.
const LAST_PUBLISHED_VERSION = 0;

/*
Current namedValues object is
nameValues = {
  'Timestamp': ['6/7/2015 20:54:13'],
  'Email Address' : ["efgh@the10xacademy.com"],
  'Student Name' : ["abc d"],
  'Student Email ID' : ["abcd@gmail.com"],
  'SSM Email ID' : ["efgh@the10xacademy.com"],
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
const MAIL_BODY_INDEX = 15;

const STUDENT_NAME = "Student Name";
const STUDENT_EMAIL = "Student Email ID";

const ALLOTTED_PROJECT = "Which project is allotted to the student?";
const DISCUSSED_PROJECT = "Which project did the student discuss?";

const VIDEO_STATUS = "Did the student switch on the video?";
const VIDEO_DURATION = "Duration of the video";
const PROJECT_EXPLANATION_SCRIPT = "Do you think student is reading from somewhere (like screen, paper etc.) or memorized?";
const SPEECH_FLOW = "How is the flow of the speech?";
const GENERAL_FEATURES_EXPLANATION = "Rate the student's explanation on project general features";
const TECH_STACK_EXPLANATION = "Rate the student's explanation on tech stack";
const CONTRIBUTION_EXPLANATION = "Rate the student's explanation on his/her contribution";

/**
 * IDEAL Answers
 */
const VIDEO_STATUS_IDEAL = "Yes";
const VIDEO_DURATION_IDEAL = "1 - 3 min";
const PROJECT_EXPLANATION_SCRIPT_IDEAL = "Memorized";
const SPEECH_FLOW_IDEAL = "Good";
const GENERAL_FEATURES_EXPLANATION_IDEAL = "Good";
const TECH_STACK_EXPLANATION_IDEAL = "Good";
const CONTRIBUTION_EXPLANATION_IDEAL = "Good";
const IDEAL_SCORE = 7;


/**
 * Template Default texts
 */

const FEEDBACK_STUDENT_NAME = "Student Name";
const FEEDBACK_TEXT = "Feedback Text";
const FINAL_TEXT = "Final Text";

const MESSAGE_TEMPLATE = `
Dear ${FEEDBACK_STUDENT_NAME},

We have reviewed your Project Explanation Video and below is the Feedback:

${FEEDBACK_TEXT}

${FINAL_TEXT}

Thanks & Regards,
The 10x Academy`;

const LINE_BREAK = ".\n";
const PROJECT_MISMATCH_PRE_TEXT = "Please discuss about the allotted project while explaining. Your allotted project is ";
const PROJECT_MISMATCH_POST_TEXT = ", but you have discussed about ";
const VIDEO_STATUS_NON_IDEAL = ". You have not switched on your video, Please switch on your video while project explanation";
const VIDEO_DURATION_NON_IDEAL_PRE_TEXT = ". Your video duration is ";
const VIDEO_DURATION_NON_IDEAL_MID_TEXT = ", The ideal video duration is ";
const PROJECT_EXPLANATION_SCRIPT_NON_IDEAL = ". It feels like you are reading from somewhere (like screen, paper etc.). Practice memorizing your script so, it feels more authentic";
const SPEECH_FLOW_NON_IDEAL = ". We have observed that your are getting stuck often with your flow of speech. Take necessary pauses to have a free flow of speech";
const GENERAL_FEATURES_EXPLANATION_NON_IDEAL_PRE_TEXT = ". Explanation on Project General Features: ";
const GENERAL_FEATURES_EXPLANATION_NON_IDEAL_POST_TEXT = " Expected: 2 - 3 Lines. Please improve on this";
const TECH_STACK_EXPLANATION_NON_IDEAL_PRE_TEXT = ". Explanation on Project Tech Stack: ";
const TECH_STACK_EXPLANATION_NON_IDEAL_POST_TEXT = " Expected: 2 - 4 Lines. Please improve on this";
const CONTRIBUTION_EXPLANATION_NON_IDEAL_PRE_TEXT = ". Explanation on your contribution in the Project: ";
const CONTRIBUTION_EXPLANATION_NON_IDEAL_POST_TEXT = " Expected: 3 - 6 Lines. Please improve on this";

const FULL_SCORE_TEXT_MAIN = `
In your project explanation video, you have discussed:

1. Project General Features,
2. Tech Stack &
3. Your contribution towards the Project

With a good flow of speech.`;

const FULL_SCORE_TEXT_END = "Thanks for Submitting!"
const PARTIAL_SCORE_TEXT = "Please create the video again, improving the point(s) discussed above and re-submit the updated video.";

const MAIL_SUBJECT = "Project Explanation Video - Feedback";
const OWNER_NAME = "The 10x Academy";


function afterFormSubmit(e) {
  const info = e.namedValues;
  const { body , score } = createMailBody(info);
  const entryRow = e.range.getRow();
  const workSheet = SpreadsheetApp.getActiveSheet();
  workSheet.getRange(entryRow, SCORE_INDEX + 1).setValue(score);
  workSheet.getRange(entryRow, MAIL_BODY_INDEX + 1).setValue(body);
  sendEmail(e.namedValues[STUDENT_EMAIL][0], body);
}

function validateEmail(email) {
  try {
    return email.endsWith("@gmail.com") || email.endsWith("@the10xacademy.com") || email.endsWith("@yahoo.com");
  } catch (error) {
    return false;
  }
}

function sendEmail(studentEmail, textBody) {
  if (!validateEmail(studentEmail)) {
    let errorMsg = "Invalid email: " + studentEmail;
    return errorMsg;
  }
  try {
    GmailApp.sendEmail(studentEmail, MAIL_SUBJECT, textBody, {
    attachments: [],
    name: OWNER_NAME
    });
  } catch (error) {
    let errorMsg = "Mail sending failed. " + error.toString();
    return errorMsg;
  }
}

function createMailBody(info) {

  let score = IDEAL_SCORE;
  let feedbackCount = 0; // for Tracking number of feedbacks

  let body = MESSAGE_TEMPLATE;

  body = body.replace(FEEDBACK_STUDENT_NAME, info[STUDENT_NAME][0]);

  if (info[ALLOTTED_PROJECT][0] !== info[DISCUSSED_PROJECT][0]) {
    body = body.replace(FEEDBACK_TEXT, PROJECT_MISMATCH_PRE_TEXT + info[ALLOTTED_PROJECT][0] + PROJECT_MISMATCH_POST_TEXT + info[DISCUSSED_PROJECT][0]);
    body = body.replace(FINAL_TEXT, PARTIAL_SCORE_TEXT);
    score = 0;
    return { body , score };
  }

  let CURRENT_FEEDBACK_TEXT = "";

  if (info[VIDEO_STATUS][0] !== VIDEO_STATUS_IDEAL) {
    feedbackCount += 1;
    CURRENT_FEEDBACK_TEXT += feedbackCount + VIDEO_STATUS_NON_IDEAL + LINE_BREAK;
  }

  if (info[VIDEO_DURATION][0] !== VIDEO_DURATION_IDEAL) {
    feedbackCount += 1;
    CURRENT_FEEDBACK_TEXT += feedbackCount + VIDEO_DURATION_NON_IDEAL_PRE_TEXT + info[VIDEO_DURATION][0] + VIDEO_DURATION_NON_IDEAL_MID_TEXT + VIDEO_DURATION_IDEAL + LINE_BREAK;
  }

  if (info[PROJECT_EXPLANATION_SCRIPT][0] !== PROJECT_EXPLANATION_SCRIPT_IDEAL) {
    feedbackCount += 1;
    CURRENT_FEEDBACK_TEXT += feedbackCount + PROJECT_EXPLANATION_SCRIPT_NON_IDEAL + LINE_BREAK;
  }

  if (info[SPEECH_FLOW][0] !== SPEECH_FLOW_IDEAL) {
    feedbackCount += 1;
    CURRENT_FEEDBACK_TEXT += feedbackCount + SPEECH_FLOW_NON_IDEAL + LINE_BREAK;
  }

  if (info[GENERAL_FEATURES_EXPLANATION][0] !== GENERAL_FEATURES_EXPLANATION_IDEAL) {
    feedbackCount += 1;
    CURRENT_FEEDBACK_TEXT += feedbackCount + GENERAL_FEATURES_EXPLANATION_NON_IDEAL_PRE_TEXT + info[GENERAL_FEATURES_EXPLANATION][0] + GENERAL_FEATURES_EXPLANATION_NON_IDEAL_POST_TEXT + LINE_BREAK;
  }

  if (info[TECH_STACK_EXPLANATION][0] !== TECH_STACK_EXPLANATION_IDEAL) {
    feedbackCount += 1;
    CURRENT_FEEDBACK_TEXT += feedbackCount + TECH_STACK_EXPLANATION_NON_IDEAL_PRE_TEXT + info[TECH_STACK_EXPLANATION][0] + TECH_STACK_EXPLANATION_NON_IDEAL_POST_TEXT + LINE_BREAK;
  }

  if (info[CONTRIBUTION_EXPLANATION][0] !== CONTRIBUTION_EXPLANATION_IDEAL) {
    feedbackCount += 1;
    CURRENT_FEEDBACK_TEXT += feedbackCount + CONTRIBUTION_EXPLANATION_NON_IDEAL_PRE_TEXT + info[CONTRIBUTION_EXPLANATION][0] + CONTRIBUTION_EXPLANATION_NON_IDEAL_POST_TEXT + LINE_BREAK;
  }

  score -= feedbackCount;

  if (score === IDEAL_SCORE) {
    body = body.replace(FEEDBACK_TEXT,FULL_SCORE_TEXT_MAIN);
    body = body.replace(FINAL_TEXT, FULL_SCORE_TEXT_END);
  } else {
    body = body.replace(FEEDBACK_TEXT,CURRENT_FEEDBACK_TEXT);
    body = body.replace(FINAL_TEXT, PARTIAL_SCORE_TEXT);
  }

  return { body , score };
}