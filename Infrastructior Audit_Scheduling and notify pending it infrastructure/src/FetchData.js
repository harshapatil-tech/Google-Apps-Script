
function fetchAuditData() {
  const fetch = new AuditFetcher();
  fetch.loadData();
  fetch.processData();
  fetch.updateReccruiterTab();
}

class AuditFetcher {
  constructor() {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.auditSheet = this.ss.getSheetByName("Audit_Requests");
    this.recruiterSheet = this.ss.getSheetByName("Recruiter");
    this.auditData = [];
    this.auditHeaders = [];
    this.latestEntries = new Map();
    this.rowToInsert = [];
  }

  loadData() {
    const dataRange = this.auditSheet.getDataRange().getValues();
    this.auditHeaders = dataRange[0];
    this.auditData = dataRange.slice(1);
    this.recruiterHeaders = this.recruiterSheet.getRange(4, 1, 1, this.recruiterSheet.getLastColumn()).getValues()[0];
  }

  processData() {
    const emailCol = this.auditHeaders.indexOf("Candidate's Email ID");
    const timestampCol = this.auditHeaders.indexOf("Timestamp");
    const auditStatusCol = this.auditHeaders.indexOf("Audit Status");
    const remarkCol = this.auditHeaders.indexOf("Remark");
    const followupCol = this.auditHeaders.indexOf("Followup Status");

    const currentTimestamp = new Date();

    for (let row of this.auditData) {
      if (!row || row.length === 0 || !row[emailCol]) continue;
      // const currentTimestamp=new Date(row[timestampCol]);
      if (isNaN(currentTimestamp)) continue;
      const existing = this.latestEntries.get(row[emailCol]);
      if (!existing || currentTimestamp > new Date(row[timestampCol])) {
        this.latestEntries[row[emailCol]] = row;
      }
    }


    for (const [key, row] of Object.entries(this.latestEntries)) {
      const auditStatus = row[auditStatusCol];
      const remark = row[remarkCol];
      const followup = row[followupCol];
      const shouldinclude = (auditStatus === "Fail" && remark === "Audit failed due to insufficient or unavailable power backup.") ||
        auditStatus === "Postponed" ||
        auditStatus === "Not Connected";
      const shouldExclude = auditStatus === "Pass" && followup === "Closed";
      if (shouldinclude && !shouldExclude) {
        const mappedRow = this.recruiterHeaders.map((header) => {
          const idx = this.auditHeaders.lastIndexOf(header);
          return idx > -1 ? row[idx] : "";
        });
        this.rowToInsert.push(mappedRow);

      }
    };
    console.log(this.rowToInsert);
  }


  updateReccruiterTab() {
    const startRow = 5;
    const numRows = this.rowToInsert.length;
    if (numRows > 0) {
      const numCols = this.rowToInsert[0].length;
      // this.recruiterSheet.getRange(startRow,1,numRows,this.rowToInsert[0].length).setValues(this.rowToInsert);
      const range = this.recruiterSheet.getRange(startRow, 1, numRows, numCols);
      range.setValues(this.rowToInsert);
    }
    else {
      //SpreadsheetApp.getUi().alert("No eligible data found");
      console.log("No eligible data found");
    }


    //DOJ
    const dojCol = this.recruiterHeaders.indexOf("DOJ") + 1;
    if (dojCol > 0) {
      const dojRange = this.recruiterSheet.getRange(5, dojCol, 15, 1); // Rows 5–19
      const rule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
      dojRange.setDataValidation(rule).setNumberFormat("dd-mmm-yyyy");

      const totalRows = this.recruiterSheet.getMaxRows();
      if (totalRows > 19) {
        this.recruiterSheet.getRange(20, dojCol, totalRows - 19, 1)
          .clearDataValidations()
          .setNumberFormat("@");
      }
    }

    //Reminder Days
    const backendSheet = this.ss.getSheetByName("Backend");
    const reminderCol = this.recruiterHeaders.indexOf("Days before reminder") + 1;

    if (reminderCol > 0) {
      const optionsRange = backendSheet.getRange("R2:R5");
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(optionsRange, true)
        .setAllowInvalid(false)
        .build();
      const limitedRange = this.recruiterSheet.getRange(5, reminderCol, 15, 1);
      limitedRange.setDataValidation(rule);
    }

    //Followup Status
    const followupCol = this.recruiterHeaders.indexOf("Followup Status") + 1;

    if (followupCol > 0) {
      const optionsRange = backendSheet.getRange("S2:S6");
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(optionsRange, true)
        .setAllowInvalid(false)
        .build();
      const limitedRange = this.recruiterSheet.getRange(5, followupCol, 15, 1);
      limitedRange.setDataValidation(rule);
    }

    //Email Candidate
    const emailCandidateColIndex = this.recruiterHeaders.indexOf("Email Candidate") + 1;

    if (emailCandidateColIndex > 0) {
      const checkboxRange = this.recruiterSheet.getRange(5, emailCandidateColIndex, 15, 1);
      const checkboxRule = SpreadsheetApp.newDataValidation()
        .requireCheckbox()
        .build();
      checkboxRange.setDataValidation(checkboxRule);
    }

    //Update
    const updateColIndex = this.recruiterHeaders.indexOf("Update") + 1;

    if (updateColIndex > 0) {
      const checkboxRange = this.recruiterSheet.getRange(5, updateColIndex, 15, 1);
      const checkboxRule = SpreadsheetApp.newDataValidation()
        .requireCheckbox()
        .build();
      checkboxRange.setDataValidation(checkboxRule);
    }
  }
}




















































































































































