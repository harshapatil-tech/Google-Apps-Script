const SOURCE_SPREADSHEET_ID = "1PKj4kWmuHs9_76_aZUwR0YVk7ifW_3jF5P-XJMOGKGw";
const TARGET_SPREADSHEET_ID = "1yAVztZBtGYPugT62jef9jbQS5FyjnRdkWQy-dKeIwSg";
const SOURCE_SHEET_NAME = "Headcount";
const TARGET_SHEET_NAME = "Employee Data";

// function updateSheetA() {
//   const sourceSS = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
//   const targetSS = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
//   const sourceSheet = sourceSS.getSheetByName(SOURCE_SHEET_NAME);
//   const targetSheet = targetSS.getSheetByName(TARGET_SHEET_NAME);
  
//   const sourceData = sourceSheet.getDataRange().getValues();
//   const targetHeaders = targetSheet.getRange(4, 1, 1, targetSheet.getLastColumn()).getValues()[0];

//   if (sourceData.length < 2) return;

//   // Column mappings between source and target sheet
//   const columnMappings = {
//     "Unique ID": "Unique ID",
//     "New Emp Id": "New Emp ID",
//     "Employee Name": "Employee Name",
//     "Function": "Function",
//     "Reporting Manager": "Reporting Manager",
//     "Grade": "Grade",
//     "Designation": "Designation",
//     "Department": "Department",
//     "Gender": "Gender",
//     "DOJ": "DOJ",
//     "DOL": "DOL",
//     "Tenure": "Tenure",
//     "Date of Resignation": "Date of Resignation",
//     "Status at the time of leaving": "Status at the time of leaving",
//     "Exit Type": "Exit Type",
//     "Status": "Status",
//     "Personal email ID": "Personal email ID",
//     "Official email ID": "Official Email ID",
//     "Current Address": "Current Address",
//     "Permanent Address": "Permanent Address",
//     "Phone Number": "Phone Number",
//     "Location": "Location",
//     "Employee Identifier": "Employee Identifier"
//   };

//   // Get headers from the source sheet
//   const sourceHeaders = sourceData[0];
//   let headerMap = {};

//   // Mapping source columns to target columns
//   Object.keys(columnMappings).forEach(targetHeader => {
//     const sourceHeader = columnMappings[targetHeader];
//     const sourceIndex = sourceHeaders.indexOf(sourceHeader);
//     const targetIndex = targetHeaders.indexOf(targetHeader);
//     if (sourceIndex !== -1 && targetIndex !== -1) {
//       headerMap[targetIndex] = sourceIndex;
//     }
//   });

  
//   // Identify the "Emp Id" column in the source sheet
//   const empIdIndexSource = sourceHeaders.indexOf("New Emp Id");
//   console.log(sourceHeaders.indexOf("New Emp Id"))
//   if (empIdIndexSource === -1) return;

//   // Create a map for existing records in the target sheet by "Emp Id"
//   const targetData = targetSheet.getRange(5, 1, targetSheet.getLastRow() - 4, targetSheet.getLastColumn()).getValues();
//   let targetMap = {};
//   console.log(targetData)
//   targetData.forEach((row, index) => {
//     const empId = row[targetHeaders.indexOf("New Emp Id")];
//     if (empId) targetMap[empId.toString().trim()] = index + 5; // Store row index in target sheet
//   });

//   // Now we will loop through each row in the source sheet and update or append data
//   let updatedData = [];
//   for (let i = 1; i < sourceData.length; i++) {
//     const sourceRow = sourceData[i];
//     const empId = sourceRow[empIdIndexSource] ? sourceRow[empIdIndexSource].toString().trim() : "";

//     // If empId exists in targetMap, update the row, else append new row
//     if (empId && targetMap[empId]) {
//       const targetRowIndex = targetMap[empId];
//       let updatedRow = new Array(targetHeaders.length).fill("");
      
//       Object.keys(headerMap).forEach(targetColIndex => {
//         const sourceColIndex = headerMap[targetColIndex];
//         updatedRow[targetColIndex] = sourceRow[sourceColIndex];
//       });
      
//       targetSheet.getRange(targetRowIndex, 1, 1, updatedRow.length).setValues([updatedRow]);
//       Logger.log('Updated row for Emp Id: ' + empId);
//     } else {
//       // Append a new row if empId does not exist in the target sheet
//       let newRow = new Array(targetHeaders.length).fill("");
      
//       Object.keys(headerMap).forEach(targetColIndex => {
//         const sourceColIndex = headerMap[targetColIndex];
//         newRow[targetColIndex] = sourceRow[sourceColIndex];
//       });
      
//       updatedData.push(newRow);
//       Logger.log('Appended new row for Emp Id: ' + empId);
//     }
//   }

//   // Append the new rows to the target sheet
//   if (updatedData.length > 0) {
//     targetSheet.getRange(targetSheet.getLastRow() + 1, 1, updatedData.length, updatedData[0].length).setValues(updatedData);
//   }

//   // Auto adjust columns and row heights
//   autoAdjustColumns(targetSheet);
// }

// const SOURCE_SPREADSHEET_ID = "1PKj4kWmuHs9_76_aZUwR0YVk7ifW_3jF5P-XJMOGKGw";
// const TARGET_SPREADSHEET_ID = "1yAVztZBtGYPugT62jef9jbQS5FyjnRdkWQy-dKeIwSg";
// const SOURCE_SHEET_NAME    = "Headcount";
// const TARGET_SHEET_NAME    = "Employee Data";

function updateSheetA() {
  const sourceSS   = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  const targetSS   = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  const sourceSheet = sourceSS.getSheetByName(SOURCE_SHEET_NAME);
  const targetSheet = targetSS.getSheetByName(TARGET_SHEET_NAME);

  const sourceData   = sourceSheet.getDataRange().getValues();
  const targetHeaders = targetSheet
    .getRange(4, 1, 1, targetSheet.getLastColumn())
    .getValues()[0];

  // nothing to do if there's no data below the header in source
  if (sourceData.length < 2) return;

  // map: target‑sheet‑header → source‑sheet‑header
  const columnMappings = {
    "Unique ID":                   "UUID",
    "New Emp Id":                  "New Emp Id",
    "Employee Name":               "Employee Name",
    "Function":                    "Function",
    "Reporting Manager":           "Reporting Manager",
    "Grade":                       "Grade",
    "Designation":                 "Designation",
    "Department":                  "Department",
    "Gender":                      "Gender",
    "DOJ":                         "DOJ",
    "DOL":                         "DOL",
    "Tenure":                      "Tenure",
    "Date of Resignation":         "Date of Resignation",
    "Status at the time of leaving":"Status at the time of leaving",
    "Exit Type":                   "Exit Type",
    "Status":                      "Status",
    "Personal email ID":           "Personal email ID",
    "Official email ID":           "Official email ID",
    "Current Address":             "Current Address",
    "Permanent Address":           "Permanent Address",
    "Phone Number":                "Phone Number",
    "Location":                    "Location",
    "Employee Identifier":         "Employee Identifier"
  };

  // build a map of target‑colIndex → source‑colIndex
  const sourceHeaders = sourceData[0];
  let headerMap = {};
  Object.entries(columnMappings).forEach(([tgtHdr, srcHdr]) => {
    const srcIdx = sourceHeaders.indexOf(srcHdr);
    const tgtIdx = targetHeaders.indexOf(tgtHdr);
    if (srcIdx >= 0 && tgtIdx >= 0) {
      headerMap[tgtIdx] = srcIdx;
    }
  });

  // which column in source holds the Emp ID?
  const empIdSourceIndex = sourceHeaders.indexOf("New Emp Id");
  if (empIdSourceIndex < 0) {
    throw new Error("Source header 'New Emp Id' not found");
  }

  // prepare to read any existing rows in the target (below the headers)
  const dataStartRow = 5;
  const numTargetRows = Math.max(0, targetSheet.getLastRow() - (dataStartRow - 1));
  const targetData = numTargetRows
    ? targetSheet.getRange(dataStartRow, 1, numTargetRows, targetSheet.getLastColumn()).getValues()
    : [];

  const empIdTargetIndex = targetHeaders.indexOf("New Emp Id");
  if (empIdTargetIndex < 0) {
    throw new Error("Target header 'New Emp ID' not found");
  }

  // do we have any non‑empty Emp IDs in the block?
  const hasTargetData = targetData.some(
    row => row[empIdTargetIndex] && row[empIdTargetIndex].toString().trim()
  );

  // build a lookup map EmpID → row number in target sheet
  let targetMap = {};
  if (hasTargetData) {
    targetData.forEach((row, i) => {
      const id = row[empIdTargetIndex].toString().trim();
      if (id) {
        targetMap[id] = dataStartRow + i;
      }
    });
  }

  // now process each source row (skip header at index 0)
  let rowsToAppend = [];
  for (let i = 1; i < sourceData.length; i++) {
    const srcRow = sourceData[i];
    const empId = srcRow[empIdSourceIndex]?.toString().trim() || "";

    if (hasTargetData && empId && targetMap[empId]) {
      // update existing row
      const tgtRowNum = targetMap[empId];
      let newRow = new Array(targetHeaders.length).fill("");
      Object.entries(headerMap).forEach(([tgtCol, srcCol]) => {
        newRow[tgtCol] = srcRow[srcCol];
      });
      targetSheet.getRange(tgtRowNum, 1, 1, newRow.length).setValues([newRow]);
      Logger.log("Updated row for Emp ID: " + empId);
    } else {
      // queue up for append (either sheet was blank, or ID not found)
      let newRow = new Array(targetHeaders.length).fill("");
      Object.entries(headerMap).forEach(([tgtCol, srcCol]) => {
        newRow[tgtCol] = srcRow[srcCol];
      });
      rowsToAppend.push(newRow);
      Logger.log("Appended new row for Emp ID: " + empId);
    }
  }

  // finally, append any new rows
  if (rowsToAppend.length) {
    const startRow = targetSheet.getLastRow() + 1;
    targetSheet
      .getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length)
      .setValues(rowsToAppend);
  }

  // Add checkboxes
  // const totalRows = targetSheet.getLastRow();
  // const checkboxColumns = ["Create Relieving Letter?", "Email Relieving Letter?"];
  // const statusColumns = ["Letter Created", "Email Sent"];

  // checkboxColumns.forEach(header => {
  //   const colIndex = targetHeaders.indexOf(header) + 1;
  //   if (colIndex > 0) {
  //     const range = targetSheet.getRange(5, colIndex, totalRows - 4);
  //     range.insertCheckboxes();
  //    // Logger.log(`Inserted checkboxes: '${header}' (Col ${colIndex}) from row 5 to ${totalRows}`);
  //   }
  //   else {
  //   //Logger.log(`Column header not found for checkbox: '${header}'`);
  // }
  // });


  // statusColumns.forEach(header => {
  //   const colIndex = targetHeaders.indexOf(header) + 1;
  //   if (colIndex > 0) {
  //     const range = targetSheet.getRange(5, colIndex, totalRows - 4);
  //     range.clearFormat();
  //     //Logger.log(`Cleared formatting : '${header}' (Col ${colIndex}) from row 5 to ${totalRows}`);
  //   }
  //   else{
  //   //Logger.log(`Column header not found for status column: '${header}'`);
  //   }
  // });

  autoAdjustColumns(targetSheet);
}

function autoAdjustColumns(sheet) {
  const lastRow    = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow > 4) {
    sheet.autoResizeColumns(1, lastColumn);
    sheet.setRowHeightsForced(5, lastRow - 4, 21);
    sheet.setFrozenRows(4);
  }
}


function autoAdjustColumns(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow > 4) {
    sheet.autoResizeColumns(1, lastColumn);
    sheet.setRowHeightsForced(5, lastRow - 4, 21);
    sheet.setFrozenRows(4);
  }
}
