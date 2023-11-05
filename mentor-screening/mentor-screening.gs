const RESPONSES_TAB_NAME = "Form Responses 1";

const Q_FIRST_COL = 6; // i.e. column F
const Q_LAST_COL = 16; // i.e. column P
const EMAIL_COL = 2; // i.e. column B
const CORRECT_FB_COL = 18; // i.e. column R
const MISSED_FB_COL = CORRECT_FB_COL + 1;
const INCORRECT_FB_COL = MISSED_FB_COL + 1;
const ADDITIONAL_FB_COL = INCORRECT_FB_COL + 1;

const Q_BASIC_IMPL =
{
  "question": "Is there anything missing from the basic implementation?",
  "options": [
    {
      "text": "Given submission looks like a completely different assignment",
      "isCorrect": false
    },
    {
      "text": "\"Learn to paint by watching others\" is missing",
      "isCorrect": false
    },
    {
      "text": "\"See how experienced ...\" paragraph is missing",
      "isCorrect": false
    },
    {
      "text": "\"Try it free for 7 days ...\" banner is missing",
      "isCorrect": false
    },
    {
      "text": "Form background (white) is missing",
      "isCorrect": false
    },
    {
      "text": "\"First Name\" field is missing",
      "isCorrect": false
    },
    {
      "text": "\"Last Name\" field is missing",
      "isCorrect": false
    },
    {
      "text": "\"Email Address\" field is missing",
      "isCorrect": false
    },
    {
      "text": "\"Password\" field is missing",
      "isCorrect": false
    },
    {
      "text": "Button is missing",
      "isCorrect": false
    },
    {
      "text": "\"By clicking the button ...\" line is missing",
      "isCorrect": false
    },
    {
      "text": "All are present",
      "isCorrect": true
    }
  ]
};

const Q_BANNER_APPEAR =
{
  "question": "Banner - Appearance: What is wrong with the appearance of \"Try it free for 7 days ...\" banner in the implementation?",
  "options": [
    {
      "text": "Given submission looks like a completely different assignment",
      "isCorrect": false
    },
    {
      "text": "\"Try it free for 7 days ...\" banner is completely absent.",
      "isCorrect": false
    },
    {
      "text": "Size of the banner is not correct. It must have 558px width and 57px height.",
      "isCorrect": false
    },
    {
      "text": "Color of the banner is not correct. It must be #5D55A1.",
      "isCorrect": false
    },
    {
      "text": "Position of the banner is not correct. It must be 136px from the top and 756px from the left.",
      "isCorrect": false
    },
    {
      "text": "Border radius of the banner is not correct. It must have 4px border radius.",
      "isCorrect": false
    },
    {
      "text": "\"Try it free for 7 days ...\" banner looks correct.",
      "isCorrect": true
    }
  ]
};

const Q_BANNER_TEXT =
{
  "question": "Banner - Text: What is wrong with the text \"Try it free for 7 days ...\" in the implementation?",
  "options": [
    {
      "text": "Given submission looks like a completely different assignment",
      "isCorrect": false
    },
    {
      "text": "\"Try it free for 7 days ...\" banner is completely absent.",
      "isCorrect": false
    },
    {
      "text": "There is a spelling mistake in \"Try it free for 7 days ...\" text.",
      "isCorrect": false
    },
    {
      "text": "The case of some letters is not correct in \"Try it free for 7 days ...\" text.",
      "isCorrect": false
    },
    {
      "text": "\"Try it free for 7 days ...\" text must be perfectly centered within the banner. But it is not centered.",
      "isCorrect": false
    },
    {
      "text": "Font size is not correct. It must be 14px.",
      "isCorrect": false
    },
    {
      "text": "Font color is not correct. It must be #FFFFFF.",
      "isCorrect": false
    },
    {
      "text": "\"Try it free for 7 days ...\" text should be in a single line.",
      "isCorrect": false
    },
    {
      "text": "\"Try it free for 7 days then\" has incorrect font. It must be Open Sans, Bold.",
      "isCorrect": false
    },
    {
      "text": "\"$20/mo there after\" has incorrect font. It must be Open Sans, Regular.",
      "isCorrect": false
    },
    {
      "text": "\"Try it free for 7 days ...\" text looks correct.",
      "isCorrect": false
    }
  ]
};

const Q_FORM_APPEAR =
{
  "question": "Form - Appearance: What is wrong with the Form background in the implementation?",
  "options": [
    {
      "text": "Given submission looks like a completely different assignment",
      "isCorrect": false
    },
    {
      "text": "Form background is completely absent.",
      "isCorrect": false
    },
    {
      "text": "Size of the Form background is not correct. It must have 558px width and 487px height.",
      "isCorrect": false
    },
    {
      "text": "Color of the Form background is not correct. It must be #FFFFFF.",
      "isCorrect": false
    },
    {
      "text": "Position of the Form background is not correct. It must be 22px below the banner and perfectly aligned with it..",
      "isCorrect": false
    },
    {
      "text": "Border radius of the Form background is not correct. It must have 10px border radius.",
      "isCorrect": false
    },
    {
      "text": "Form background looks correct.",
      "isCorrect": true
    }
  ]
};

const Q_EMAIL_APPEAR =
{
  "question": "Email - Appearance: What is wrong with the appearance of \"Email Address\" field in the implementation?",
  "options": [
    {
      "text": "Given submission looks like a completely different assignment",
      "isCorrect": false
    },
    {
      "text": "\"Email Address\" field is completely absent.",
      "isCorrect": false
    },
    {
      "text": "Size of the \"Email Address\" field is not correct. It must have 475px width and 56px height.",
      "isCorrect": true
    },
    {
      "text": "Border radius of the \"Email Address\" field is not correct. It must be 4px.",
      "isCorrect": false
    },
    {
      "text": "Border thickness of the \"Email Address\" field is not correct. It must be 1px.",
      "isCorrect": false
    },
    {
      "text": "Placeholder is missing. It must have a placeholder \"Email Address\".",
      "isCorrect": false
    },
    {
      "text": "There are spelling mistakes in the placeholder (\"Email Address\").",
      "isCorrect": false
    },
    {
      "text": "Placeholder (\"Email Address\") must be 35px from the left border and vertically center aligned within the input field. But it is not.",
      "isCorrect": true
    },
    {
      "text": "Position of the \"Email Address\" field is not correct. It must be horizontally center aligned within the form background and 23px below the bottom border of the \"Last Name\" field.",
      "isCorrect": true
    },
    {
      "text": "The \"Email Address\" field looks correct.",
      "isCorrect": false
    }
  ]
};

const Q_EMAIL_FUNC =
{
  "question": "Email - Functionality: What is wrong with the functionality of \"Email Address\" field in the implementation?",
  "options": [
    {
      "text": "Given submission looks like a completely different assignment",
      "isCorrect": false
    },
    {
      "text": "\"Email Address\" field is completely absent.",
      "isCorrect": false
    },
    {
      "text": "When the \"Email Address\" field becomes active (e.g. when user clicks inside the field or tabs into the field), the border color must change to #5C54A1. But it is not happening.",
      "isCorrect": true
    },
    {
      "text": "When the \"Email Address\" field is not active, its border color must be #C9C9C9. But it has a different color.",
      "isCorrect": false
    },
    {
      "text": "By clicking inside the \"Email Address\" field, I should be able to enter some text. But it is not allowing me to type.",
      "isCorrect": false
    },
    {
      "text": "If the entered email address is not valid, the border should become red in color.",
      "isCorrect": true
    },
    {
      "text": "If the entered email address is not valid, a message \"Please enter a valid email\" should appear below the email address field.",
      "isCorrect": true
    },
    {
      "text": "The \"Email Address\" field works fine.",
      "isCorrect": false
    }
  ]
};

const Q_PASSWORD_APPEAR =
{
  "question": "Password - Appearance: What is wrong with the appearance of \"Password\" field in the implementation?",
  "options": [
    {
      "text": "Given submission looks like a completely different assignment",
      "isCorrect": false
    },
    {
      "text": "\"Password\" field is completely absent.",
      "isCorrect": false
    },
    {
      "text": "Size of the \"Password\" field is not correct. It must have 475px width and 56px height.",
      "isCorrect": true
    },
    {
      "text": "Border radius of the \"Password\" field is not correct. It must be 4px.",
      "isCorrect": false
    },
    {
      "text": "Border thickness of the \"Password\" field is not correct. It must be 1px.",
      "isCorrect": false
    },
    {
      "text": "Placeholder is missing. It must have a placeholder \"Password\".",
      "isCorrect": false
    },
    {
      "text": "There are spelling mistakes in the placeholder (\"Password\").",
      "isCorrect": false
    },
    {
      "text": "Placeholder (\"Password\") must be 35px from the left border and vertically center aligned within the input field. But it is not.",
      "isCorrect": true
    },
    {
      "text": "Position of the \"Password\" field is not correct. It must be horizontally center aligned within the form background and 23px below the bottom border of the \"Email Address\" field.",
      "isCorrect": true
    },
    {
      "text": "The \"Password\" field looks correct.",
      "isCorrect": false
    }
  ]
};

const Q_PASSWORD_FUNC =
{
  "question": "Password - Functionality: What is wrong with the functionality of \"Password\" field in the implementation?",
  "options": [
    {
      "text": "Given submission looks like a completely different assignment",
      "isCorrect": false
    },
    {
      "text": "\"Password\" field is completely absent.",
      "isCorrect": false
    },
    {
      "text": "When the \"Password\" field becomes active (e.g. when user clicks inside the field or tabs into the field), the border color must change to #5C54A1. But it is not happening.",
      "isCorrect": true
    },
    {
      "text": "When the \"Password\" field is not active, its border color must be #C9C9C9. But it has a different color.",
      "isCorrect": false
    },
    {
      "text": "By clicking inside the \"Password\" field, I should be able to enter some text. But it is not allowing me to type.",
      "isCorrect": false
    },
    {
      "text": "\"Password\" field should have a show/hide toggle button that will show/mask the password. The show/hide button is missing.",
      "isCorrect": true
    },
    {
      "text": "By default, the password must be masked. But the password is shown as clear text.",
      "isCorrect": false
    },
    {
      "text": "If the entered password is not valid (viz. less than 8 characters), the border should become red in color.",
      "isCorrect": true
    },
    {
      "text": "If the entered password is not valid (viz. less than 8 characters), a message \"Password must have at least 8 characters\" should appear below the email address field.",
      "isCorrect": true
    },
    {
      "text": "The \"Password\" field works fine.",
      "isCorrect": false
    }
  ]
};

const Q_BUTTON_APPEAR =
{
  "question": "Button - Appearance: What is wrong with the appearance of the button in the implementation?",
  "options": [
    {
      "text": "Given submission looks like a completely different assignment",
      "isCorrect": false
    },
    {
      "text": "The button is completely absent.",
      "isCorrect": false
    },
    {
      "text": "Size of the button is not correct. It must have 475px width and 56px height.",
      "isCorrect": true
    },
    {
      "text": "Border radius of the button is not correct. It must be 4px.",
      "isCorrect": false
    },
    {
      "text": "Fill color of the button is not correct. It must be #37CC89.",
      "isCorrect": false
    },
    {
      "text": "Title of the button is not correct. It must be \"CLAIM YOUR FREE TRAIL\".",
      "isCorrect": false
    },
    {
      "text": "There are spelling mistakes in the title (\"CLAIM YOUR FREE TRAIL\").",
      "isCorrect": false
    },
    {
      "text": "Title (\"CLAIM YOUR FREE TRAIL\") must be all uppercase. But it is not.",
      "isCorrect": false
    },
    {
      "text": "Title (\"CLAIM YOUR FREE TRAIL\") must be horizontally and vertically center aligned within the button. But it is not.",
      "isCorrect": true
    },
    {
      "text": "Shadow is missing (or incorrect) for the button. It must have \"box-shadow: 0px 3px 6px #00000029;\"",
      "isCorrect": false
    },
    {
      "text": "Position of the button is not correct. It must be horizontally center aligned within the form background and 25px below the bottom border of the \"Password\" field.",
      "isCorrect": true
    },
    {
      "text": "The button looks correct.",
      "isCorrect": false
    }
  ]
};

const Q_BUTTON_FUNC =
{
  "question": "Button - Functionality: What is wrong with the functionality of the button in the implementation?",
  "options": [
    {
      "text": "Given submission looks like a completely different assignment",
      "isCorrect": false
    },
    {
      "text": "The button is completely absent.",
      "isCorrect": false
    },
    {
      "text": "When any of the First/Last Name fields is empty and we click the button, the empty fields must be highlighted using red border.",
      "isCorrect": true
    },
    {
      "text": "When any of the First/Last Name fields is empty and we click the button, a message must be shown below the respective fields that that field must not be empty.",
      "isCorrect": true
    },
    {
      "text": "When any of the Email/Password fields is not valid and we click the button, the invalid fields must be highlighted using red border.",
      "isCorrect": true
    },
    {
      "text": "When any of the Email/Password fields is not valid and we click the button, a message must be shown below the respective fields that a valid value must be entered.",
      "isCorrect": true
    },
    {
      "text": "The button works fine.",
      "isCorrect": false
    }
  ]
};

const Q_TnC =
{
  "question": "T&C: What is wrong with the \"By clicking the button ...\" line in the implementation?",
  "options": [
    {
      "text": "Given submission looks like a completely different assignment",
      "isCorrect": false
    },
    {
      "text": "\"By clicking the button ...\" line is completely absent.",
      "isCorrect": false
    },
    {
      "text": "Position of the \"By clicking the button ...\" line is not correct. It must be horizontally center aligned within the form background and 13px below the button.",
      "isCorrect": false
    },
    {
      "text": "\"By clicking the button ...\" line has incorrect font size. It must be 13px.",
      "isCorrect": false
    },
    {
      "text": "There are spelling mistakes in \"By clicking the button ...\" line.",
      "isCorrect": false
    },
    {
      "text": "\"By clicking the button you are accepting\" has incorrect font. It must be Open Sans, Regular.",
      "isCorrect": false
    },
    {
      "text": "\"Terms & Conditions\" has incorrect font. It must be Open Sans, Semibold.",
      "isCorrect": false
    },
    {
      "text": "\"By clicking the button you are accepting\" has incorrect color. It must be #A8A8A8.",
      "isCorrect": false
    },
    {
      "text": "\"Terms & Conditions\" has incorrect color. It must be #FF7978.",
      "isCorrect": false
    },
    {
      "text": "\"Terms & Conditions\" must be a hyperlink. But it is normal text.",
      "isCorrect": true
    },
    {
      "text": "\"By clicking the button ...\" line looks and works correctly.",
      "isCorrect": false
    }
  ]
};

const POINTS_MAP = 
{
  "questions": [
    Q_BASIC_IMPL,
    Q_BANNER_APPEAR,
    Q_BANNER_TEXT,
    Q_FORM_APPEAR,
    Q_EMAIL_APPEAR,
    Q_EMAIL_FUNC,
    Q_PASSWORD_APPEAR,
    Q_PASSWORD_FUNC,
    Q_BUTTON_APPEAR,
    Q_BUTTON_FUNC,
    Q_TnC
  ]
};

const TOTAL_QUESTIONS = POINTS_MAP.questions.length;
let NUM_ANSWERS = 0,
    NUM_NON_ANSWERS = 0;
for (const q of POINTS_MAP.questions) {
  for (const option of q.options) {
    if (option.isCorrect) {
      ++NUM_ANSWERS;
    } else {
      ++NUM_NON_ANSWERS;
    }
  }
}

const MENU_TITLE = "10x score calculator";
const MENU_HELP = "Help";
const MENU_SCORES = "Generate scores";

const TITLE_HELP = "Help";

const MSG_HELP = `
${MENU_TITLE}
Some common mistakes. 
1. Did you modify the responses sheet?
  -- Did you rename the tab? Expected name is "Form Responses 1".
  -- Did you add any new columns? Add new columns only at the end.
  -- Did you delete any columns? Do not delete any existing columns.
  -- Did you re-order the columns? Do not re-order the columns.
2. Did you modify the Google form?
  -- Did you add any new questions? Do not add new questions to the form.
  -- Did you delete any existing questions? Do not delete any existing questions.
  -- Did you modify any options (even like adding a space)? Do not modify the options at all.
  -- Did you add any new options? Do not add any new options.
  -- Did you change the order of the questions? Do not change order of the questions.

Contact Sailesh.
`;

function onInstall(e) {
  // Without this, user will have to reload the sheet after installing.
  onOpen(e);
}

function onOpen(e) {
  SpreadsheetApp.getUi()
  .createMenu(MENU_TITLE)
  .addItem(MENU_SCORES, "generateScores")
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

function getMetrics(response, qIndex) {
  if (qIndex < 0) {
    console.error(`Received negative qIndex: ${qIndex}`);
  } else if (qIndex >= POINTS_MAP.questions.length) {
    console.error(`qIndex is out of bounds: ${qIndex}`);
  }
  const options = POINTS_MAP.questions[qIndex].options;
  let answersMarked = 0,
      answersNotMarked = 0,
      nonAnswersMarked = 0,
      nonAnswersNotMarked = 0;
  for (const option of options) {
    const searchText = option.text.trim();
    if (response.includes(searchText)) {
      if (option.isCorrect) {
        ++answersMarked;
      } else {
        ++nonAnswersMarked;
      }
      response = response.replace(searchText, "").trim();
    } else {
      if (option.isCorrect) {
        ++answersNotMarked;
      } else {
        ++nonAnswersNotMarked;
      }
    }
  }
  response = response.replaceAll(",", "").trim();
  const additionalFeedback = response.length > 0;
  return {
    answersMarked,
    answersNotMarked,
    nonAnswersMarked,
    nonAnswersNotMarked,
    additionalFeedback
  };
}

function generateScoreForRow(sheet, rowNumber) {
  const responderId = sheet.getRange(rowNumber, EMAIL_COL).getValue().trim();
  console.log(`Computing for ${responderId}`);

  let answersMarked = 0,
      answersNotMarked = 0,
      nonAnswersMarked = 0,
      additionalFeedback = 0;
  for (let colNumber = Q_FIRST_COL; colNumber <= Q_LAST_COL; ++colNumber) {
    const response = sheet.getRange(rowNumber, colNumber).getValue().trim();
    const metrics = getMetrics(response, colNumber - Q_FIRST_COL);
    answersMarked += metrics.answersMarked;
    answersNotMarked += metrics.answersNotMarked;
    nonAnswersMarked += metrics.nonAnswersMarked;
    if (metrics.additionalFeedback) {
      ++additionalFeedback;
    }
  }
  
  const correctFeedback = parseInt((answersMarked*100)/(answersMarked + nonAnswersMarked));
  const missedFeedback = parseInt((answersNotMarked*100)/(answersMarked + answersNotMarked));
  const incorrectFeedback = parseInt((nonAnswersMarked*100)/(answersMarked + nonAnswersMarked));
  additionalFeedback = parseInt((additionalFeedback * 100)/TOTAL_QUESTIONS);

  if (rowNumber === 2) {
    let errorMsg = null;
    if (correctFeedback < 100) {
      errorMsg += `\n${responderId} should get 100% correct feedback`;
    }
    if (missedFeedback !== 0) {
      errorMsg += `\n${responderId} should get 0% missed feedback`;
    }
    if (incorrectFeedback !== 0) {
      errorMsg += `${responderId} should get 0% incorrect feedback`;
    }
    if (errorMsg) {
      console.log(errorMsg);
      throw(errorMsg);
    }
  }

  try {
    sheet.getRange(rowNumber, CORRECT_FB_COL).setValue(correctFeedback);
    sheet.getRange(rowNumber, MISSED_FB_COL).setValue(missedFeedback);
    sheet.getRange(rowNumber, INCORRECT_FB_COL).setValue(incorrectFeedback);
    sheet.getRange(rowNumber, ADDITIONAL_FB_COL).setValue(additionalFeedback);
  } catch(err) {
    console.error(`Unable to write the metrics to the row ${rowNumber} for ${responderId}`);
    console.error(err);
  }
  return;
}

function getResponseSheet() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESPONSES_TAB_NAME);
  } catch (err) {
    console.error(`Unable to retrieve the sheet - ${RESPONSES_TAB_NAME}`);
  }
  return null;
}

function generateScores() {
  const sheet = getResponseSheet();
  const lastRow = sheet.getLastRow();
  for (let i = 2; i <= lastRow; ++i) {
    generateScoreForRow(sheet, i);
  }
}