/**
 * Updates six “OnlineWk” sheets in the same spreadsheet:
 *   • unmerges rows 7→end in cols N→R
 *   • repopulates Sr. No., Name, per-row =SUM(P:Q)
 *   • appends a Total row with column SUMs
 */
class OnlineWkSummaryUpdater {
  /**
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss  the spreadsheet
   * @param {string[]}               onlineSheetNames   e.g. ["OnlineWk1",…]
   */
  constructor(ss, date, onlineSheetNames = ["Online_Wk_1", "Online_Wk_2","Online_Wk_3", "Online_Wk_4","Online_Wk_5", "Online_Wk_6"]) {
    this.ss               = ss;
    this.onlineSheetNames = onlineSheetNames;
    this.tutorSheetName   = "Tutor_Names";
    this.runAll();
  }

  runAll() {
    // grab the up-to-date list of tutors (excludes X/Y)
    const tutorNames = this._getCleanTutorNames();

    // update each OnlineWk sheet
    this.onlineSheetNames.forEach(name => {
      const sh = this.ss.getSheetByName(name);
      if (sh) this._updateOneSummary(sh, tutorNames);
    });
  }

  _getCleanTutorNames() {
    const sheet = this.ss.getSheetByName(this.tutorSheetName);
    const [headers, rows] = CentralLibrary.get_Data_Indices_From_Sheet(sheet);
    const rawNames = rows
      .map(r => r[headers["Name of Tutors"]])
      .filter(v => v && v !== "Y");
    return rawNames;
  }

  _updateOneSummary(sheet, tutorNames) {
    const startRow = 7;
    const srCol    = 14;  // N
    const nameCol  = 15;  // O
    const dayCol   = 16;  // P
    const nightCol = 17;  // Q
    const totCol   = 18;  // R
    const numCols  = totCol - srCol + 1;

    // 1) Find the existing “Total” row by scanning column O
    const lastPossibleRow = sheet.getLastRow();
    let lastExistingRow = startRow;
    const srColRange = sheet.getRange(startRow, srCol, lastPossibleRow - startRow + 1, 1);
    const names     = srColRange.getValues().flat();
    for (let i = 0; i < names.length; i++) {
      if (names[i] === "Total") {
        lastExistingRow = startRow + i;
        break;
      }
    }
    // if “Total” wasn’t found, default to clearing just the header+1 row
    if (lastExistingRow < startRow) {
      lastExistingRow = startRow;
    }

    // 2) Un-merge and clear everything from row 7 through the old Total
    const oldBlock = sheet.getRange(startRow, srCol,
                                    lastExistingRow - startRow + 1,
                                    numCols);
    oldBlock.breakApart();
    oldBlock.clear();

    const N        = tutorNames.length;
    const totalRow = startRow + N;

    // 4) Write the data rows
    for (let i = 0; i < N; i++) {
      const r = startRow + i;
      sheet.getRange(r, srCol)   .setValue(i + 1);
      sheet.getRange(r, nameCol) .setValue(tutorNames[i]);
      const pA1 = sheet.getRange(r, dayCol)  .getA1Notation();
      const qA1 = sheet.getRange(r, nightCol).getA1Notation();
      sheet.getRange(r, totCol)
          .setFormula(`=SUM(${pA1}:${qA1})`);
    }

    // 5) Write the new Total row (merge N+O, no serial)
    sheet.getRange(totalRow, srCol, 1, 2).merge();
    sheet.getRange(totalRow, srCol).setValue("Total");

    // 6) Column-sums
    const pRange = `${sheet.getRange(startRow,   dayCol).getA1Notation()}:${sheet.getRange(startRow + N - 1, dayCol).getA1Notation()}`;
    const qRange = `${sheet.getRange(startRow,   nightCol).getA1Notation()}:${sheet.getRange(startRow + N - 1, nightCol).getA1Notation()}`;
    const rRange = `${sheet.getRange(startRow,   totCol).getA1Notation()}:${sheet.getRange(startRow + N - 1, totCol).getA1Notation()}`;

    sheet.getRange(totalRow, dayCol)  .setFormula(`=SUM(${pRange})`);
    sheet.getRange(totalRow, nightCol).setFormula(`=SUM(${qRange})`);
    sheet.getRange(totalRow, totCol)  .setFormula(`=SUM(${rRange})`);

    // 7) Draw thin-grid borders around the new block
    const newBlock = sheet.getRange(startRow, srCol,
                                    totalRow - startRow + 1,
                                    numCols);
    newBlock
    .setFontFamily("Roboto")
    .setFontSize(10)
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true);

    // 8) Bold + medium border on the Total row
    sheet.getRange(totalRow, srCol, 1, numCols)
        .setFontFamily("Roboto")
        .setFontSize(10)
        .setHorizontalAlignment("center")
        .setFontWeight("bold")
        .setBorder(true, true, true, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }


}