function smeDBNew() {

  const updater = new SMEUpdater(
    EMPLOYEE_DETAIL_SHEET_ID,
    EMPLOYEE_TAB_ID,
    MASTER_DB_SPREADSHEET_ID,
    SME_DB_TAB_ID
  );
  updater.makeNewRow();
  //console.log("New rows to append:", updater.updatedOutputData);
  updater.appendNewRows();
}


//sme DB sheet update function 
class SMEUpdater {
  constructor(employeeeSheetId, employeeTabId, smeSheetId, smeTabId) {

    this.employeeeSheetId = employeeeSheetId;
    this.employeeTabId = employeeTabId;

    this.smeSheetId = smeSheetId;
    this.smeTabId = smeTabId;

    //employee sheet
    this.employeeSheetObj = CentralLibrary.DataAndHeaders(this.employeeeSheetId);
    const empWrapper = this.employeeSheetObj.getSheetById(this.employeeTabId);
    this.employeeSheet = empWrapper.sheet;
    const [empHeaders, empData] = empWrapper.getDataIndicesFromSheet();
    this.empHeaders = empHeaders;
    this.empData = empData;

    //sme sheet
    this.smeSheetObj = CentralLibrary.DataAndHeaders(this.smeSheetId);
    const smeWrapper = this.smeSheetObj.getSheetById(this.smeTabId);
    this.smeSheet = smeWrapper.sheet;
    const [smeHeaders, smeData] = smeWrapper.getDataIndicesFromSheet();
    this.smeHeaders = smeHeaders;
    this.smeData = smeData;

    //map department data
    this.hrDepartmentData = departmentMapping();
    //console.log(this.hrDepartmentData);

    this.updatedOutputData = [];
    //console.log(this.updatedOutputData);

    this.lastSrNo = this.smeSheet.getRange(this.smeSheet.getLastRow(), 1).getValue() || 0;
    //console.log( this.lastSrNo);
  }

  makeNewRow() {
    const resultArray = [];

    for (let r of this.empData) {
      const deptKey = r[this.empHeaders["Department"]];
      const officialEmail = r[this.empHeaders["Official Email ID"]];
      const uniqueId = r[this.empHeaders["Unique ID"]];

      if (this.hrDepartmentData.hasOwnProperty(deptKey) && officialEmail !== '-' && uniqueId) {

        const department = this.hrDepartmentData[deptKey];
        //console.log("Exists in hrDepartmentData?", this.hrDepartmentData.hasOwnProperty(deptKey));

        resultArray.push([
          0,
          uniqueId,
          officialEmail,
          department,
          //r[this.empHeaders["Employee Name"]],
          `${r[this.empHeaders["Employee Name"]]} (${r[this.empHeaders["Emp Id"]]})`,
          "", "",
          r[this.empHeaders["Grade"]],
          r[this.empHeaders["Designation"]],
          r[this.empHeaders["Reporting Manager"]]
        ]);
      }
    }

    for (let row of resultArray) {
      this.checkAndUpdate(row);
    }
  }


  //check employee already exist or not 
  checkAndUpdate(newRow) {
    const uniqueId = newRow[1];
    let found = false;
    //console.log("Checking Unique ID:", uniqueId);


    for (let i = 0; i < this.smeData.length; i++) {

      const existingRow = this.smeData[i];
      const existingId = existingRow[this.smeHeaders["Unique ID"]];
      //console.log("Comparing with existing:", existingId);

      if (existingRow[this.smeHeaders["Unique ID"]] === uniqueId) {
        found = true;
        //if exist then update the rows
        this.updateIfChanged(i, "Email ID", existingRow, newRow[2]);
        this.updateIfChanged(i, "SME Name", existingRow, newRow[4]);
        this.updateIfChanged(i, "Department", existingRow, newRow[3]);
        this.updateIfChanged(i, "Grade", existingRow, newRow[7]);
        this.updateIfChanged(i, "Designation", existingRow, newRow[8]);
        this.updateIfChanged(i, "Reporting Manager", existingRow, newRow[9]);

        break;
      }
    }


    //if not exist then increment the sr.no. add new rows,
    if (!found) {
      console.log(`New row to be added: ${uniqueId}`);
      this.lastSrNo++;
      newRow[0] = this.lastSrNo;
      this.updatedOutputData.push(newRow);
    } else {
      console.log(`Already exists: ${uniqueId}`);
    }

  }


  updateIfChanged(rowIndex, fieldName, existingRow, newValue) {
    const colIndex = this.smeHeaders[fieldName];
    const oldValue = existingRow[colIndex];

    //console.log(`Comparing for field "${fieldName}": OLD=[${oldValue}], NEW=[${newValue}]`);

    if (oldValue !== newValue) {
      //console.log(`--> Updating at row ${rowIndex + 2}, col ${colIndex + 1}`);
      this.smeSheet.getRange(rowIndex + 2, colIndex + 1).setValue(newValue);
    }
  }


  appendNewRows() {
    if (this.updatedOutputData.length > 0) {
      const lastRow = this.smeSheet.getLastRow();
      this.smeSheet.getRange(lastRow + 1, 1, this.updatedOutputData.length, this.updatedOutputData[0].length).setValues(this.updatedOutputData);
      // console.log("Appending rows:", this.updatedOutputData.length);

      const activeCol = this.smeHeaders["Active?"] + 1;

      this.smeSheet.getRange(lastRow + 1, activeCol, this.updatedOutputData.length, 1).insertCheckboxes().setValue(false);
    }
  }
}









//__________***** OLD Function *****_____________________
// function smeDBNew() {
//   const employeeDetailsSpreadsheetID = "11b5A_ZDi9DskTiKJ0j4fBDkRbY-0I2h7rVRxr6KVK1o";
//   const employeeDetailsSpreadsheet = CentralLibrary.DataAndHeaders(employeeDetailsSpreadsheetID);
//   const employeeDetailsSheet = employeeDetailsSpreadsheet.getSheetById(1020906063);
//   const [ inputheaders, inputData ] = employeeDetailsSheet.getDataIndicesFromSheet();
  
//   const smeDBSpreadsheetID = "1eJ-7nKeICwYXEBh4oEzLwC6SpkYzCw5ntYxarbD2jxA";
//   const smeDBSpreadsheet = CentralLibrary.DataAndHeaders(smeDBSpreadsheetID);
//   const smeDBSheet = smeDBSpreadsheet.getSheetById(1777024406).sheet;
//   const [ outputheaders, outputData ] = smeDBSpreadsheet.getDataIndicesFromSheet();

//   const hrDepartmentData = departmentMapping();

//   let index = 1;
//   let lastSrNo = smeDBSheet.getRange(smeDBSheet.getLastRow(), 1).getValue(); 

//   const resultArray = [];
//   inputData.forEach(r => {

//     if(hrDepartmentData.hasOwnProperty(r[inputheaders["Department"]]) && r[inputheaders["Official Email ID"]] !== '-'){

//       const department = hrDepartmentData[r[inputheaders["Department"]]];
//       resultArray.push([index++, r[inputheaders["Official Email ID"]], department, r[inputheaders["Employee Name"]], "", "", r[inputheaders["Grade"]], r[inputheaders["Designation"]], r[inputheaders["Reporting Manager"]]]);
//     }
    
//   });

//   const updatedOutputData = [];
//   resultArray.forEach((row, index) => {
//       const email = row[1];
//       const resultDepartment = row[2];
//       const resultPyramidCategory = row[6];
//       const resultDesignation = row[7];
//       const resultReportingManager = row[8]
//       let found = false;
      
//       // Find the row in outputData with this email
//       for (let r of outputData) {
        
//           if (r[outputheaders["Email ID"]] === email) {

//               found = true;
//               // Check if department has changed
//               if (r[outputheaders["Department"]] !== resultDepartment) {

//                   console.log(`Department change for ${email}: ${r[outputheaders["Department"]]} to ${resultDepartment}`);
//                   smeDBSheet.getRange(outputData.indexOf(r) + 2, outputheaders["Department"] + 1).setValue(resultDepartment);

//               }

//               // Check if pyramidCategory has changed
//               if (r[outputheaders["Grade"]] !== resultPyramidCategory) {

//                   console.log(`Pyramid Category change for ${email}: ${r[outputheaders["Grade"]]} to ${resultPyramidCategory}`);
//                   smeDBSheet.getRange(outputData.indexOf(r) + 2, outputheaders["Grade"] + 1).setValue(resultPyramidCategory);

//               }

//               // Check if pyramidCategory has changed
//               if (r[outputheaders["Designation"]] !== resultDesignation) {

//                   console.log(`Pyramid Category change for ${email}: ${r[outputheaders["Designation"]]} to ${resultDesignation}`);
//                   smeDBSheet.getRange(outputData.indexOf(r) + 2, outputheaders["Designation"] + 1).setValue(resultDesignation);

//               }

//               if (r[outputheaders["Reporting Manager"]] !== resultReportingManager) {

//                   console.log(`Reporting Manager Change for ${email}: ${r[outputheaders["Reporting Manager"]]} to ${resultReportingManager}`);
//                   smeDBSheet.getRange(outputData.indexOf(r) + 2, outputheaders["Reporting Manager"] + 1).setValue(resultReportingManager);

//               }

//               break;
//           }
//       }

//       // If email not found in outputData, add this row
//       if (!found) {
//           lastSrNo ++;
//           row[0] = lastSrNo
//           updatedOutputData.push(row);

//       }
//   });
//   // Append the new data (if any) to outputSheet
//   if (updatedOutputData.length > 0) {
//     const lastRow = smeDBSheet.getLastRow();
//     smeDBSheet.getRange(lastRow + 1, 1, updatedOutputData.length, updatedOutputData[0].length).setValues(updatedOutputData);
//     smeDBSheet.getRange(lastRow + 1, outputheaders["Active?"]+1, updatedOutputData.length, 1).insertCheckboxes().setValue(false);
//   }
// }




