function showAuthorizationDialog() {
  FormApp;
  SpreadsheetApp;
  GmailApp;
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Show Authorization Dialog', 'showAuthorizationDialog')
    .addSeparator()
    .addItem('Update Course Lineup', 'updateCourseLineup')
    .addSeparator()
    .addItem('Update Instructor Lineup', 'updateInstructorLine')
    .addToUi();
}

function updateCourseLineup() {
  let formId = extractFormIdFromUrl_(FORM_URL);
  let targetForm = FormApp.openById(formId);
  let items = targetForm.getItems(FormApp.ItemType.CHECKBOX);

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let coursesSheet = ss.getSheetByName(COURSES_SHEET_NAME);

  for (let item of items) {
    let checkboxItemQuestion = item.asCheckboxItem();
    let itemTitle = item.getTitle();

    let column;
    if (AA_COURSES_ITEMS.includes(itemTitle)) {
      column = 'A';
    } else if (BB_COURSES_ITEMS.includes(itemTitle)) {
      column = 'C';
    } else if (CC_COURSES_ITEMS.includes(itemTitle)) {
      column = 'E';
    }

    if (column) {
      // Get all values in the column and filter out empty ones
      let allValues = coursesSheet.getRange(column + '2:' + column).getValues();
      let filteredValues = allValues.filter((value) => value[0] !== '');

      // Create choices from the non-empty values
      let choices = filteredValues.map((course) =>
        checkboxItemQuestion.createChoice(course[0])
      );

      checkboxItemQuestion.setChoices(choices);
    }
  }
}

function extractFormIdFromUrl_(formUrl) {
  let regex = /\/d\/(.*?)(\/|$)/;
  let matches = formUrl.match(regex);
  return matches ? matches[1] : null;
}

function updateInstructorLine() {
  let formId = extractFormIdFromUrl_(FORM_URL);
  let targetForm = FormApp.openById(formId);
  let items = targetForm.getItems(FormApp.ItemType.LIST);

  for (item of items) {
    let listItemQuestion = item.asListItem();
    let eachItem = item;
    let itemTitle = eachItem.getTitle();
    // console.log(eachItemTitle);

    let column;
    if (AA_INSTRUCTORS_ITEMS.includes(itemTitle)) {
      column = 'C';
    } else if (BB_INSTRUCTORS_ITEMS.includes(itemTitle)) {
      column = 'E';
    } else if (CC_INSTRUCTORS_ITEMS.includes(itemTitle)) {
      column = 'G';
    }

    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let instructorSheet = ss.getSheetByName(INSTRUCTORS_SHEET_NAME);

    if (column) {
      // Get all values in the column and filter out empty ones
      let allValues = instructorSheet
        .getRange(column + '2:' + column)
        .getValues();
      console.log(allValues);
      let filteredValues = allValues.filter((value) => value[0] !== '');
      console.log(filteredValues);
      let choices = filteredValues.map((instructor) =>
        listItemQuestion.createChoice(instructor[0])
      );

      listItemQuestion.setChoices(choices);
    }
  }
}

// This function should be triggered when a form is submitted
function onFormSubmit(event) {
  // Use 'namedValues' for a Spreadsheet trigger
  const responseObj = event.namedValues;
  // Initialize variables for the responses
  let firstOptionResponse = null;
  let secondOptionResponse = null;
  let email = null;
  let name = null;

  // Iterate over the response object to find the relevant items
  for (let title in responseObj) {
    if (title.includes(FIRST_OPTION_ITEM_PHRASE)) {
      firstOptionResponse = escapeHtml_(responseObj[title][0]); // Use the first element of the array
    }
    if (title.includes(SECOND_OPTION_ITEM_PHRASE)) {
      secondOptionResponse = escapeHtml_(responseObj[title][0]); // Use the first element of the array
    }
    if (title === 'Email Address') {
      email = escapeHtml_(responseObj[title][0]);
    }
    if (title === '2. Name') {
      name = escapeHtml_(responseObj[title][0]);
    }
  }

  // Check if both responses exist and if they are the same
  if (
    firstOptionResponse &&
    secondOptionResponse &&
    firstOptionResponse === secondOptionResponse
  ) {
    const targetResponses = {
      email: email,
      name: name,
      firstReponse: firstOptionResponse,
    };
    console.log(
      'This response is subject to the resubmission. Thus, notification email will be sent.'
    );

    sendNotificationEmail_(targetResponses);
  }
}

// Function to send an email notification
function sendNotificationEmail_(responseObj) {
  let recipient = responseObj.email;
  let subject = `Please Resubmit your Application`;

  // Load the HTML template for the email body
  let htmlTemplate = HtmlService.createTemplateFromFile('email-template');
  htmlTemplate.name = responseObj['name']; // Assuming 'Name' is the title for the name field
  htmlTemplate.firstChoice = responseObj['firstReponse'];
  htmlTemplate.googleFormURL = FORM_URL;
  let body = htmlTemplate.evaluate().getContent();

  let options = {
    htmlBody: body,
    cc: CC,
  };

  GmailApp.sendEmail(recipient, subject, '', options);
}
function escapeHtml_(unsafe) {
  return unsafe
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
