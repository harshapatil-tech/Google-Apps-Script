// Email ID used for sending notifications
const SENDER_EMAIL_ID = Session.getActiveUser().getEmail();

const SUPPORT_CC = "support@upthink.com";

// Predefined headers used for CRUD operations in the spreadsheet
const CRUD_HEADERS = [
  "Ticket No.",
  "Timestamp",
  "Employee Name",
  "Contact Number",
  "Device Type",
  "Hardware Issue",
  "Software Issue",
  "IT Support Diagnosis",
  "Estimated Time",
  "Issue Status",
  "Update?",
  "Email Address",
  "Location",
  "Date",
  "Time",
  "Issue Type",
  "Department",
  "Detailed Address",
  "Laptop Brand",
  "Laptop Number",
  "Laptop Number Picture",
  "Issue Description",
  "Remarks"
];



/**
 * This function is triggered when the spreadsheet is opened.
 * - It fetches the "CRUD" and "Backend" sheets from the spreadsheet.
 * - Retrieves the 'Current Issue Status' column from the 'Backend' sheet to use as data validation options.
 * - Applies the data validation to cell E4 in the "CRUD" sheet.
 * - Inserts the current date into specified cells.
 * - Adds a custom menu to the Google Sheets UI for launching a form dialog.
 */
function onOpen() {
  const spreadsheetId = "1fL71mTIzO8gw1quS-lgCiLnYyYHRGpESoiksUQTpbHQ";
  const crudSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const crudSheet = crudSpreadsheet.getSheetByName('CRUD');
  const crudBackendSheet = crudSpreadsheet.getSheetByName('Backend');

  // Get header from Backend sheet
  const backendHeader = crudBackendSheet
    .getRange(1, 1, 1, crudBackendSheet.getLastColumn())
    .getValues()[0];
  const statusColIndex = backendHeader.indexOf('Current Issue Status') + 1;
  const numRows = crudBackendSheet.getLastRow() - 1;
  const statusData =
    numRows > 0
      ? crudBackendSheet
          .getRange(2, statusColIndex, numRows, 1)
          .getValues()
          .flat()
      : [];

  // Set data validation on cell E4 in CRUD sheet using statusData
  const statusCell = crudSheet.getRange(4, 5);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statusData, true)
    .setAllowInvalid(false)
    .build();
  statusCell.setDataValidation(rule);

  // Insert the current date into specific cells
  insertDate(crudSheet);

  // Add custom menu to launch the dialog
  SpreadsheetApp.getUi()
    .createMenu('Custom Dialog')
    .addItem('Open Form', 'showForm')
    .addToUi();

}

/**
 * Displays a custom modal dialog for form input.
 * - Ensures the user selects exactly one cell.
 * - Ensures the selected cell is in column H (column 8) and row 9 or later.
 * - Passes the target cell information to the HTML template.
 */
function showForm() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();

  // Ensure that exactly one cell is selected
  if (!range || range.getNumRows() !== 1 || range.getNumColumns() !== 1) {
    SpreadsheetApp.getUi().alert(
      "Please select exactly ONE cell (in column H, row 9 or below) and try again."
    );
    return;
  }

  // Ensure the selected cell is in column H (column 8) and row is 9 or later
  if (range.getColumn() !== 8 || range.getRow() < 9) {
    SpreadsheetApp.getUi().alert(
      "Please select a cell in column H and try again."
    );
    return;
  }

  // Pass target cell info (A1 notation and sheet name) into the HTML template
  const template = HtmlService.createTemplateFromFile("Dialogue_box");
  template.targetCell = range.getA1Notation();
  template.sheetName = sheet.getName();
  const htmlOutput = template.evaluate().setWidth(400).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Select Options");
}

// Processes form data submitted from the dialog.
function processForm(choices, targetCell, sheetName) {
  // choices: array of selected option names (strings)
  const spreadsheetId = "1fL71mTIzO8gw1quS-lgCiLnYyYHRGpESoiksUQTpbHQ";
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);

  // Update the target cell with the joined choices (using setValue for a single cell)
  sheet.getRange(targetCell).setValue(choices.join(', '));
}


/**
 * Inserts the current date into specific cells (B2 and B4) in the provided sheet.
 * - Formats the date as "dd-MM-yyyy".
 *
 * @param {Sheet} sheet - Google Sheet where dates will be inserted.
 */
function insertDate(sheet) {
  const currentDate = new Date();
  const startDateCell = sheet.getRange(2, 2);
  const endDateCell = sheet.getRange(4, 2);
  startDateCell.setValue(currentDate);
  endDateCell.setValue(currentDate);
  // Format cells as "dd-MM-yyyy"
  startDateCell.setNumberFormat("dd/MM/yyyy");
  endDateCell.setNumberFormat("dd/MM/yyyy");
}


/**
 * Includes the content of an HTML file for use in templates.
 * - Facilitates reusability of HTML files in modal dialogs.
 *
 * @param {string} filename - Name of the HTML file to include.
 * @returns {string} - Raw content of the HTML file.
 */
function include(filename) {
  return HtmlService.createTemplateFromFile(filename).getRawContent();
}

//cc email lists
const ccEmailLists = {
  "Mathematics" : "tejas.jagtap@upthink.com",
  "Statistics" : "asmita.sane@upthink.com",
  "Physics" : "hanumant.lahane@upthink.com",
  "Chemistry" : "shruti.sardesai@upthink.com",
  "Biology" : "shruti.sardesai@upthink.com",
  "Accounts" : "tushar.jangale@upthink.com",
  "Finance" : "tushar.jangale@upthink.com",
  "Economics" : "tushar.jangale@upthink.com",
  "Computer Science" : "kuber.deokar@upthink.com",
  "Instructional Design" : "apurva.yadav@upthink.com",
  "Finance & Accounts" : "pranav.deshpande@upthink.com",
  "Human Resource" : "tanaya.adulkar@upthink.com",
  "Operations" : "apurva.yadav@upthink.com",
  "Administration" : "tejas.jagtap@upthink.com",
  "Marketing" : "apurva.yadav@upthink.com",
  "Technology" : "sreenjay.sen@upthink.com",
}

const siliconRentalTeam = {
  "Silicon Rental" : ["aman@silicongroup1.com","sudhir@silicongroup1.com","satish@silicongroup1.com","bhaskar@silicongroup1.com","kumar@silicongroup1.com"],
}

function temp(){
  console.log(siliconRentalTeam["Silicon Rental"].join(", ") )
}



