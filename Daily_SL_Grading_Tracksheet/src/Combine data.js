function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Combined Data")
    .addItem("Update Sheet", "combine")
    .addToUi();
}

function combine() {
  const LAST_COLUMN = 15;      // Data from columns A–O
  const DATE_COLUMN_INDEX = 0; // Column A (0-based in arrays)

  const currentSS = SpreadsheetApp.getActiveSpreadsheet();
  const currentBackend = currentSS.getSheetByName("Backend_Data");

  // ---------- 1. Define date window: last 2 months ----------
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // Start from the 1st of (currentMonth - 2)
  const refreshFromDate = new Date(today.getFullYear(), today.getMonth() - 2, 1);
  refreshFromDate.setHours(0, 0, 0, 0);
  // Example: if today = 28 Dec → refreshFromDate = 1 Oct

  // ---------- 2. Read source data ----------
  const labGradingSS = SpreadsheetApp.openById("1XtERJDXnGyY5HIHisqsxPmXlz1Aqw13lhJb2ej_okXw");
  const labGradingBackendSheet = labGradingSS.getSheetByName("Backend_Data");
  const labLastRow = labGradingBackendSheet.getLastRow();
  const labGradingBackend = labLastRow > 1
    ? labGradingBackendSheet.getRange(2, 1, labLastRow - 1, LAST_COLUMN).getValues()
    : [];

  const assignmentGradingSS = SpreadsheetApp.openById("1gSdxxH3E8j4RCspIqlE7sk4cUZlw8JZNb5bQQ6gou2w");
  const assignmentGradingBackendSheet = assignmentGradingSS.getSheetByName("Backend_Data");
  const assignLastRow = assignmentGradingBackendSheet.getLastRow();
  const assignmentGradingBackend = assignLastRow > 1
    ? assignmentGradingBackendSheet.getRange(2, 1, assignLastRow - 1, LAST_COLUMN).getValues()
    : [];

  // ---------- 3. Filter rows in sources between refreshFromDate and today ----------
  const filterByDateWindow = (row) => {
    const cell = row[DATE_COLUMN_INDEX];
    if (!cell) return false;

    let d = cell;
    if (!(d instanceof Date)) {
      d = new Date(cell);
    }
    if (isNaN(d)) return false;

    d.setHours(0, 0, 0, 0);
    return d >= refreshFromDate && d <= today;
  };

  const filteredLab = labGradingBackend.filter(filterByDateWindow);
  const filteredAssignment = assignmentGradingBackend.filter(filterByDateWindow);

  const combinedData = [...filteredLab, ...filteredAssignment];

  console.log("Rows in window:", combinedData.length);

  // If there is no data in that window, do nothing
  if (combinedData.length === 0) return;

  // ---------- 4. In current backend, find first row with date >= refreshFromDate ----------
  const lastRow = currentBackend.getLastRow(); // includes header
  let writeStartRow = 2; // default if only header exists

  if (lastRow > 1) {
    const existingDataRange = currentBackend.getRange(2, 1, lastRow - 1, LAST_COLUMN + 2);
    const existingData = existingDataRange.getValues();

    let deleteFromRow = null; // sheet row index (1-based)

    for (let i = 0; i < existingData.length; i++) {
      const row = existingData[i];
      const cell = row[DATE_COLUMN_INDEX];

      if (!cell) continue;

      let d = cell;
      if (!(d instanceof Date)) {
        d = new Date(cell);
      }
      if (isNaN(d)) continue;

      d.setHours(0, 0, 0, 0);

      // First row whose date is >= refreshFromDate (e.g., 1 Oct)
      if (d >= refreshFromDate) {
        deleteFromRow = 2 + i; // because data starts at row 2
        break;
      }
    }

    if (deleteFromRow !== null) {
      // Clear ONLY rows from that date onwards
      const numRowsToClear = lastRow - deleteFromRow + 1;
      if (numRowsToClear > 0) {
        currentBackend
          .getRange(deleteFromRow, 1, numRowsToClear, LAST_COLUMN + 2)
          .clearContent();
      }
      writeStartRow = deleteFromRow;
    } else {
      // No rows with date >= refreshFromDate; append at bottom
      writeStartRow = lastRow + 1;
    }
  }

  // ---------- 5. Write combined (filtered) data ----------
  currentBackend
    .getRange(writeStartRow, 1, combinedData.length, combinedData[0].length)
    .setValues(combinedData);

  // ---------- 6. Re-apply formulas ----------
  currentBackend
    .getRange(2, LAST_COLUMN + 1)
    .setFormula('=ARRAYFORMULA(IF(E2:E="","",VLOOKUP(E2:E, Course_Details!$A$3:$F, 2, FALSE)))');

  const monthFormula =
    '=ARRAYFORMULA(IF(LEN(B2:B)=0,"",IFERROR(MATCH(B2:B,{"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"},0)&" "&B2:B&" "&RIGHT(C2:C,2),"")))';

  currentBackend
    .getRange(2, LAST_COLUMN + 2)
    .setFormula(monthFormula);
}
