// function onOpen() {
//   SpreadsheetApp.getUi()
//   .createMenu("Combined Data")
//   .addItem("Update Sheet", "combine")
//   .addToUi();
// }

// function combine() {
//   const currentSS = SpreadsheetApp.getActiveSpreadsheet();
//   const currentBackend = currentSS.getSheetByName("Backend_Data");
//   const lastColumn = 15;

//   const labGradingSS = SpreadsheetApp.openById("1XtERJDXnGyY5HIHisqsxPmXlz1Aqw13lhJb2ej_okXw");
//   const labGradingBackendSheet = labGradingSS.getSheetByName("Backend_Data");
//   const labGradingBackend = labGradingBackendSheet
//         .getRange(2, 1, labGradingBackendSheet.getLastRow() - 1, lastColumn)
//         .getValues();

//   const assignmentGradingSS = SpreadsheetApp.openById("1gSdxxH3E8j4RCspIqlE7sk4cUZlw8JZNb5bQQ6gou2w");
//   const assignmentGradingBackendSheet = assignmentGradingSS.getSheetByName("Backend_Data");
//   const assignmentGradingBackend = assignmentGradingBackendSheet
//         .getRange(2, 1, assignmentGradingBackendSheet.getLastRow() - 1, lastColumn)
//         .getValues();

//   // Combine both datasets
//   const combinedData = [...labGradingBackend, ...assignmentGradingBackend];

//   // 👉 Sort by date (first column) in ascending order
//   combinedData.sort((a, b) => {
//     const da = a[0];
//     const db = b[0];

//     // If they are already date objects or serial numbers, this handles both
//     if (da instanceof Date && db instanceof Date) {
//       return da - db;
//     }
//     if (typeof da === 'number' && typeof db === 'number') {
//       return da - db; // Google Sheets date serials
//     }

//     // Fallback: try to parse as strings
//     return new Date(da) - new Date(db);
//   });

//   console.log("Combined rows:", combinedData.length);

//   if (combinedData.length === 0) return;

//   // Clear old data (from row 2 onwards, including formula columns)
//   const lastRow = currentBackend.getLastRow();
//   if (lastRow > 1) {
//     currentBackend.getRange(2, 1, lastRow - 1, lastColumn + 2).clearContent();
//   }

//   // Write sorted combined data
//   currentBackend
//     .getRange(2, 1, combinedData.length, combinedData[0].length)
//     .setValues(combinedData);

//   // Add formulas
//   currentBackend
//     .getRange(2, lastColumn + 1) // first row for formula
//     .setFormula('=ARRAYFORMULA(IF(E2:E="","",VLOOKUP(E2:E, Course_Details!$A$3:$F, 2, FALSE)))');

//   const monthFormula =
//     '=ARRAYFORMULA(IF(LEN(B2:B)=0,"",IFERROR(MATCH(B2:B,{"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"},0)&" "&B2:B&" "&RIGHT(C2:C,2),"")))';

//   currentBackend
//     .getRange(2, lastColumn + 2)
//     .setFormula(monthFormula);
// }



// function combine() {
//   const currentSS = SpreadsheetApp.getActiveSpreadsheet();
//   const currentBackend = currentSS.getSheetByName("Backend_Data");
//   const lastColumn = 15;

//   const labGradingSS = SpreadsheetApp.openById("1XtERJDXnGyY5HIHisqsxPmXlz1Aqw13lhJb2ej_okXw");
//   const labGradingBackendSheet = labGradingSS.getSheetByName("Backend_Data");
//   const labGradingBackend = labGradingBackendSheet
//                     .getRange(2, 1, labGradingBackendSheet.getLastRow(), lastColumn).getValues();
//   const assignmentGradingSS = SpreadsheetApp.openById("1gSdxxH3E8j4RCspIqlE7sk4cUZlw8JZNb5bQQ6gou2w");
//   const assignmentGradingBackendSheet = assignmentGradingSS.getSheetByName("Backend_Data");
//   const assignmentGradingBackend = assignmentGradingBackendSheet
//                         .getRange(2, 1, assignmentGradingBackendSheet.getLastRow(), lastColumn).getValues();

//   const combinedData = [...labGradingBackend, ...assignmentGradingBackend];

//   console.log(combinedData.length)

//   if (combinedData.length === 0) return;

//   // Clear from second row down
//   const lastRow = currentBackend.getLastRow();
//   if (lastRow > 1) {
//     currentBackend.getRange(2, 1, lastRow - 1, lastColumn+2).clearContent();
//   }

//   // Write combined data starting at row 2
//   currentBackend.getRange(2, 1, combinedData.length, combinedData[0].length).setValues(combinedData);

//   // Get last column number
//   // const lastColumn = currentBackend.getLastColumn();

//   // // Add VLOOKUP formula in the next column
//   // const formulaRange = currentBackend.getRange(2, lastColumn, combinedData.length);
//   currentBackend
//   .getRange(2, lastColumn+1) // first row of data
//   .setFormula('=ARRAYFORMULA(IF(E2:E="","",VLOOKUP(E2:E, Course_Details!$A$3:$F, 2, FALSE)))');

//   const monthFormula =
//   '=ARRAYFORMULA(IF(LEN(B2:B)=0,"",IFERROR(MATCH(B2:B,{"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"},0)&" "&B2:B&" "&RIGHT(C2:C,2),"")))';

//   currentBackend
//   .getRange(2, lastColumn+2)
//   .setFormula(monthFormula);

// }























// function combine() {
//   const currentSS = SpreadsheetApp.getActiveSpreadsheet();
//   const currentBackend = currentSS.getSheetByName("Backend_Data");

//   const labGradingSS = SpreadsheetApp.openById("1XtERJDXnGyY5HIHisqsxPmXlz1Aqw13lhJb2ej_okXw");
//   const labGradingBackend = labGradingSS.getSheetByName("Backend_Data").getDataRange().getValues();

//   const assignmentGradingSS = SpreadsheetApp.openById("1gSdxxH3E8j4RCspIqlE7sk4cUZlw8JZNb5bQQ6gou2w");
//   const assignmentGradingBackend = assignmentGradingSS.getSheetByName("Backend_Data").getDataRange().getValues();

//   const combinedData = [...labGradingBackend.slice(1,), ...assignmentGradingBackend.slice(1,)];
  

//   // Clear from second row down
//   const lastRow = currentBackend.getLastRow();
//   if (lastRow > 1) {
//     currentBackend.getRange(2, 1, lastRow - 1, currentBackend.getLastColumn()).clearContent();
//   }

//   // Write combined data starting at row 2
//   currentBackend.getRange(2, 1, combinedData.length, combinedData[0].length).setValues(combinedData);

//   // Get last column number
//   const lastColumn = currentBackend.getLastColumn();

//   // Add ARRAYFORMULA in the next column starting from row 3
//   currentBackend
//     .getRange(2, lastColumn) // only the first cell of the new column
//     .setFormula('=ARRAYFORMULA(IF(E2:E="","",VLOOKUP(E3:E, Course_Details!$A$2:$F$104, 2, FALSE)))');
// }
