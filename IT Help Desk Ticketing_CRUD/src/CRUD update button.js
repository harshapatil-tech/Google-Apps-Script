function updateButton() {

  // Instantiate the UpdateUI class and call the update method
  const updateUI = new UpdateUI();
  updateUI.update()

}

class UpdateUI{
  constructor(){
    // Open the spreadsheets using their IDs
    this.spreadsheet = SpreadsheetApp.openById("1fL71mTIzO8gw1quS-lgCiLnYyYHRGpESoiksUQTpbHQ");
    this.masterDBSpreadsheet = SpreadsheetApp.openById("1iByitSy5R35cu13rupuppctpzV0X8dTzqDSJJP2ilAk");

    // Get references to the CRUD and Master DB sheets
    this.crudSheet = this.spreadsheet.getSheetByName('CRUD');
    this.masterDBSheet = this.masterDBSpreadsheet.getSheetByName('Master DB');
    
    // Define the starting row for data in CRUD sheet
    this.crudSheetStartRow = 9;
    this.dbSheetStartRow = 2;

    // Fetch headers and data from both sheets
    [this.crudHeaders, this.crudData] = CentralLibrary.get_Data_Indices_From_Sheet(this.crudSheet, this.crudSheetStartRow - 2);
    [this.dbHeaders, this.dbData] = CentralLibrary.get_Data_Indices_From_Sheet(this.masterDBSheet);

    this.dbStatus = {
      "1.1 Open" : ["3.1 Inform Silicon Rental", "4.1 Close Ticket", "2.1 Ticket Raised [A]", "1.1 Open"],
      "2.1 Ticket Raised [A]" : ["4.1 Close Ticket", "3.1 Inform Silicon Rental", "2.1 Ticket Raised [A]"],
      "3.1 Inform Silicon Rental" : ["4.1 Close Ticket", "3.2 Silicon Rental Informed [A]", "3.3 Ticket Raised [Silicon] [A]", "2.1 Ticket Raised [A]", "3.1 Inform Silicon Rental"],
      "2.2 Silicon Rental Informed [A]" : ["4.1 Close Ticket", "2.1 Ticket Raised [A]","2.2 Silicon Rental Informed [A]"],
      "2.3 Ticket Raised [Silicon] [A]" : ["4.1 Close Ticket", "2.3 Ticket Raised [Silicon] [A]"],
      "4.1 Close Ticket" : ["5.1 Re-open ticket", "4.2 Ticket closed [A]", "4.1 Close Ticket"],
      "4.2 Ticket closed [A]" : ["5.1 Re-open ticket", "4.2 Ticket closed [A]"],
      "5.1 Re-open ticket" : ["6.1 Final closure", "5.2 Ticket reopened [A]", "5.1 Re-open ticket"],
      "5.2 Ticket reopened [A]" : ["6.1 Final closure", "5.2 Ticket reopened [A]"],
      "6.1 Final closure" : ["6.2 Final closure [A]", "6.1 Final closure"],
      "3.3 Ticket Raised [Silicon] [A]" : ["3.3 Ticket Raised [Silicon] [A]", "4.1 Close Ticket"]
    }

    // console.log(this.dbStatus["1.1 Open"].includes());

    // Retrieve the column index for Ticket No. from the Master DB headers
    this.ticketnoColumnInDb = this.dbHeaders["Ticket No."];
    // Create an empty object to store a mapping of Ticket No. to its corresponding row data
    this.dbMapByTicketNo = {};
    for (let i = 0; i < this.dbData.length; i++) {
      const row = this.dbData[i];
      const ticketno = row[this.ticketnoColumnInDb];
      this.dbMapByTicketNo[ticketno] = row;
    }
    this.dbRowToIndexMap = this.getTicketNoToDBRowIndex();

    // Indices for referencing columns in the CRUD
    this.ticketNoColumnInCrud = this.crudHeaders["Ticket No."];
    this.updateColumnInCrud = this.crudHeaders["Update?"];
    this.issueStatusColumnInCrud = this.crudHeaders["Issue Status"];
    
  }

  update(){
    const rowsToDelete = []; 
    // 1) Gather the row numbers in the CRUD that have "Update?" = TRUE
    const updateTrueCRUDRows = this.crudData
      .map((row, idx) => {
        if (row[this.updateColumnInCrud] === true) {
          return idx + this.crudSheetStartRow; // Convert 0-based index to sheet row
        }
        return -1;
      })
      .filter(rowNum => rowNum !== -1);
    if (updateTrueCRUDRows.length === 0) {
      // No rows to update
      SpreadsheetApp.getUi().alert("No rows selected for update.");
      return;
    }

    const rowsThatFailValidation = [];
    const rowsThatPassValidation = [];
    // Process each row marked for update
    updateTrueCRUDRows.forEach(crudRowNumber =>{
      const crudDataIndex = crudRowNumber - this.crudSheetStartRow;  // Convert back to crudData index
      const crudRow = this.crudData[crudDataIndex];
      const ticketNo = crudRow[this.crudHeaders["Ticket No."]];
      const timestamp = crudRow[this.crudHeaders["Timestamp"]];
      const employeeName = crudRow[this.crudHeaders["Employee Name"]];
      const contactNumber = crudRow[this.crudHeaders["Contact Number"]];
      const deviceType = crudRow[this.crudHeaders["Device Type"]];
      const hardwareIssue = crudRow[this.crudHeaders["Hardware Issue"]];
      const softwareIssue = crudRow[this.crudHeaders["Software Issue"]];
      const itSupportDiagnosis = crudRow[this.crudHeaders["IT Support Diagnosis"]];
      const estimatedTime = crudRow[this.crudHeaders["Estimated Time"]];
      const issueStatus = crudRow[this.crudHeaders["Issue Status"]];
      const emailAddress = crudRow[this.crudHeaders["Email Address"]];
      const location = crudRow[this.crudHeaders["Location"]];
      const date = crudRow[this.crudHeaders["Date"]];
      const time = crudRow[this.crudHeaders["Time"]];
      const issueType = crudRow[this.crudHeaders["Issue Type"]];
      const department = crudRow[this.crudHeaders["Department"]];
      const detailedAddress = crudRow[this.crudHeaders["Detailed Address"]];
      const laptopBrand = crudRow[this.crudHeaders["Laptop Brand"]];
      const laptopNumber = crudRow[this.crudHeaders["Laptop Number"]];
      const laptopNumberPicture = crudRow[this.crudHeaders["Laptop Number Picture"]];
      const issueDescription = crudRow[this.crudHeaders["Issue Description"]];
      const remarks = crudRow[this.crudHeaders["Remarks"]];

      //find the index of the row in the master db by ticket no
      const dbRowIndex = this.dbData.findIndex(dbRow => dbRow[this.dbHeaders["Ticket No."]] === ticketNo);
      // console.log("dbRow",dbRowIndex);
      if (dbRowIndex === -1)return; //skip if ticket no not found in master db

      const dbRow = this.dbData[dbRowIndex];


      // const finalClosure = dbRow[this.dbHeaders["Final Closure"]];

      // // If "Final Closure" is true, show an error dialog and do not update
      // if (finalClosure === true) {
      //   this.showErrorDialog(`Ticket ${ticketNo} is in Final Closure and cannot be updated.`);
      //   return;
      // }

      let hasChanges = false;
      
      // Compare each column and update Master DB where necessary
      Object.keys(this.crudHeaders).forEach(colName => {
        if (this.dbHeaders[colName] !== undefined) {
          if (crudRow[this.crudHeaders[colName]] !== dbRow[this.dbHeaders[colName]]) {
            // Update the Master DB with the new value
            // this.masterDBSheet.getRange(dbRowIndex + 2, this.dbHeaders[colName] + 1).setValue(crudRow[this.crudHeaders[colName]]);
            hasChanges = true;
          }
        }
      });

      // If changes were made, mark row for deletion from CRUD
      if (hasChanges) {
        rowsToDelete.push(crudRowNumber);
      }
      
    })

    // 2) Validate each row in turn
    for (const crudRowNumber of updateTrueCRUDRows) {
      const crudRowIndex = crudRowNumber - this.crudSheetStartRow; 
      const crudRow = this.crudData[crudRowIndex];
      // Logger.log("CrudRow " + crudRow);
      const ticketno = crudRow[this.ticketNoColumnInCrud];
      const dbRow = this.dbMapByTicketNo[ticketno]; // old data from Master DB
      // Logger.log("dbRow from db " + dbRow);
      // If there's no matching TicketNo in DB, skip / fail
      if (!dbRow) {
        rowsThatFailValidation.push(crudRowNumber);
        continue;
      }
      // Perform all validations
      const passesAll = this.validateRow(crudRow, dbRow);
      if (passesAll) {
        rowsThatPassValidation.push(crudRowNumber);
        // console.log("IF Block")
      } else {
        rowsThatFailValidation.push(crudRowNumber);
        // console.log("Else Block");
      }
    }

    // 3) Update the Master DB for rows that pass validation
    //Then mark them for deletion from the CRUD
    if (rowsThatPassValidation.length > 0) {
      this.updateMasterDB(rowsThatPassValidation);
    }

    // 4) Delete rows from CRUD in descending order (to avoid index shifting)
    const allRowsToDelete = rowsThatPassValidation;
    allRowsToDelete.sort((a, b) => b - a); // descending
    for (const rowNum of allRowsToDelete) {
      this.crudSheet.deleteRow(rowNum);
    }
    SpreadsheetApp.getUi().alert(`Rows not updated: ${rowsThatFailValidation.length}\n` + 
                                 `Rows updated: ${allRowsToDelete.length}`);
  
  }

  // Update master db - For each validated row, update Master DB data
  updateMasterDB(validRowNumbers){
    for(const crudRowNumber of validRowNumbers){
      const crudRowIndex = crudRowNumber - this.crudSheetStartRow;
      const crudRow = this.crudData[crudRowIndex];
      const dbDataIndex = this.dbRowToIndexMap[crudRow[this.ticketNoColumnInCrud]];
      const dbSheetRowNumber = this.dbSheetStartRow + dbDataIndex;
      const dbRow = this.dbData[dbDataIndex];

      this.updateDBRow(dbRow, crudRow);
      // console.log("Db Row", dbRow);
      this.masterDBSheet.getRange(dbSheetRowNumber, 1, 1, dbRow.length).setValues([dbRow]);

      // emailTemplate1(this.masterDBSheet, this.dbHeaders, dbRow, 0);
      emailTemplate4(this.masterDBSheet, this.dbHeaders, dbRow, dbDataIndex);
      emailTemplate5(this.masterDBSheet, this.dbHeaders, dbRow, dbDataIndex);
      emailTemplate6(this.masterDBSheet, this.dbHeaders, dbRow, dbDataIndex);
      // emailTemplate3(this.masterDBSheet, this.dbHeaders, dbRow, dbRowIndex);
      // emailTemplate2(this.masterDBSheet, this.dbHeaders, dbRow, dbRowIndex);
      emailTemplate3(this.masterDBSheet, this.dbHeaders, dbRow, dbDataIndex);
      emailTemplate2(this.masterDBSheet, this.dbHeaders, dbRow, dbDataIndex);
      
    }
  }

  getTicketNoToDBRowIndex() {
    // Create a mapping of ticket numbers to their corresponding row indices in the Master DB
    const ticketNoToDBRowIndex = {};
    for (let i = 0; i < this.dbData.length; i++) {
      // Retrieve the row from the Master DB sheet data
      const row = this.dbData[i];
      // Extract the ticket number from the row
      const ticketno = row[this.ticketnoColumnInDb];
      // Map the ticket number to its row index in the Master DB
      ticketNoToDBRowIndex[ticketno] = i;
    }
    // Return the mapping of ticket numbers to row indices
    return ticketNoToDBRowIndex;
  }

  // Validate a single row
  // Returns true if row passes all checks; false otherwise
  validateRow(crudRow, dbRow){
    const oldTicketNumber = dbRow[this.dbHeaders["Ticket No."]].trim();
    const newTicketNumber = crudRow[this.ticketNoColumnInCrud].trim();
    const oldIssueStatus = dbRow[this.dbHeaders["Issue Status"]].trim();
    
    const newIssueStatus = crudRow[this.crudHeaders["Issue Status"]].trim();
    const validate = this.issueStatusValidation(oldIssueStatus, newIssueStatus)

    if (!validate) {
      SpreadsheetApp.getUi().alert("Automation Error: Please verify if specified criteria are met before proceeding.");
      return false;
  }

    // if(validate == false){
     
    //   return false;
    // }
    return true;
  }

  // Check allowed dbSatus
  issueStatusValidation(oldIssueStatus, newIssueStatus){
    // If oldTicketNo not in dictionary, let’s assume it can’t change
    // if (!this.dbStatus[newTicketNo]) return false;
    // Check if newTicketNo is allowed from oldTicketNo
    // return this.dbStatus[oldTicketNo].includes(newTicketNo);
    // console.log("Old St", oldIssueStatus, newIssueStatus);
    // console.log("Valid Status", this.dbStatus[oldIssueStatus].includes(newIssueStatus));
    if(oldIssueStatus.trim() ==="6.2 Final closure [A]")
      return false;
    return this.dbStatus[oldIssueStatus].includes(newIssueStatus);

  }

  updateDBRow(dbRow, crudRow){
    // Update the Ticket No. field in the Master DB with the value from the CRUD sheet
    dbRow[this.dbHeaders["Ticket No."]] = crudRow[this.ticketNoColumnInCrud];

    // Only update IT Support diagnosis if it's empty in db
    if(dbRow[this.dbHeaders["IT Support diagnosis"]] === ""){
      dbRow[this.dbHeaders["IT Support diagnosis"]] = crudRow[this.crudHeaders["IT Support Diagnosis"]];
    }

    // Only update Estimated Time if it's empty in db
    if(dbRow[this.dbHeaders["Estimated Time"]] === ""){
      dbRow[this.dbHeaders["Estimated Time"]] = crudRow[this.crudHeaders["Estimated Time"]];
    }
    // Update the Issue Status field in the Master DB with the value from the CRUD sheet
    dbRow[this.dbHeaders["Issue Status"]] = crudRow[this.crudHeaders["Issue Status"]];

    // Update the Remarks field in the Master DB with the value from the CRUD sheet
    dbRow[this.dbHeaders["Remarks"]] = crudRow[this.crudHeaders["Remarks"]];

    // Update the Last Update Date field in the Master DB with the current date
    dbRow[this.dbHeaders["Last Update Date"]] = Utilities.formatDate(new Date(),Session.getScriptTimeZone, "dd-MMM-yyyy");

  }


}