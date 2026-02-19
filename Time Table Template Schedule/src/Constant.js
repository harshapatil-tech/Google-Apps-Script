//const SUBJECTS = ["Biology", "Business", "Chemistry", "Computer Science", "Mathematics", "Physics", "Statistics"];
const SUBJECTS=["Physics"]
// Time table folder for every subject from 2022
//const GOOGLE_DRIVE_FOLDER_ID = "1GsRppc-WVzGYJJ9BSLazbCP4OS42R5ud";
const GOOGLE_DRIVE_FOLDER_ID = "1Sp4xDezGYW8XgIJkMRJ9DG9FSogeeHjx";


function onOpen(e) {
  const sheet = e.source.getSheetByName("Create Spreadsheet");  // Correctly referencing the sheet
  if (sheet) {
    const range1 = sheet.getRange("B4");  // Specify the cell where the dropdown will appear
    const rule1 = SpreadsheetApp
                    .newDataValidation()
                    .requireValueInList(["All", ...SUBJECTS])
                    .build();
    range1.setDataValidation(rule1);  // Applying the data validation to the cell
    // UPDATE USER IN NEW TIMETABLE
    const range2 = sheet.getRange("B12:B16");
    const rule2 = SpreadsheetApp
                    .newDataValidation()
                    .requireValueInList(SUBJECTS)
                    .build();
    range2.setDataValidation(rule2);
    // UPDATE CLIENT DROPDOWN
    const range3 = sheet.getRange("C4");
    const rule3 = SpreadsheetApp
                    .newDataValidation()
                    .requireValueInList(["BF", "NT"])
                    .build();
    range3.setDataValidation(rule3);
    // UPDATE SPREADSHEET CREATION TYPE
    const range4 = sheet.getRange("D4");
    const rule4 = SpreadsheetApp
                    .newDataValidation()
                    .requireValueInList(["Timetable", "QC"])
                    .build();
    range4.setDataValidation(rule4);
  }
}