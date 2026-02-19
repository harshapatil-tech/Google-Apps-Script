function runOnOpen(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const reviewerManagementSheet = spreadsheet.getSheetByName("Reviewer Management");
  const accountManagementSheet = spreadsheet.getSheetByName("Account Management");
  const departmentSheet = spreadsheet.getSheetByName("Department Management");
  const smeManagementSheet = spreadsheet.getSheetByName("SME Management");

  // const dataRange = outputSheet.getRange(1, 2, outputSheet.getLastRow(), outputSheet.getLastColumn() - 1).getValues();
  // const dataRange = outputSheet.getRange(1, 2, outputSheet.getLastRow(), outputSheet.getLastColumn() - 1).getValues();

  const reviewerAddDataRange = reviewerManagementSheet.getRange(3, 1, 4, 4).getValues();
  const reviewerAddHeaders = reviewerAddDataRange[0], reviewerAddData = reviewerAddDataRange.slice(1);

  const reviewerRemoveDataRange = reviewerManagementSheet.getRange(9, 1, 4, 4).getValues();
  const reviewerRemoveHeaders = reviewerRemoveDataRange[0], reviewerRemoveData = reviewerRemoveDataRange.slice(1);

  const reviewerAddIndices = {
    emailIdx : reviewerAddHeaders.indexOf("Email ID"),
    departmentIdx : reviewerAddHeaders.indexOf("Department"),
    addIdx : reviewerAddHeaders.indexOf("Add?"),
  }

  const reviewerRemoveIndices = {
    emailIdx : reviewerRemoveHeaders.indexOf("Email ID"),
    departmentIdx : reviewerRemoveHeaders.indexOf("Department"),
    removeIdx : reviewerRemoveHeaders.indexOf("Remove?"),
  }

  const smeAddDataRange = smeManagementSheet.getRange(4, 1, 11, 4).getValues();
  const smeAddHeaders = smeAddDataRange[0], smeAddData = smeAddDataRange.slice(1);

  const smeRemoveDataRange = smeManagementSheet.getRange(17, 1, 6, 4).getValues();
  const smeRemoveHeaders = smeRemoveDataRange[0], smeRemoveData = smeRemoveDataRange.slice(1);

  const smeAddIndices = {
    emailIdx : smeAddHeaders.indexOf("Email ID"),
    departmentIdx : smeAddHeaders.indexOf("Department"),
    addIdx : smeAddHeaders.indexOf("Add?"),
  }

  const smeRemoveIndices = {
    emailIdx : smeRemoveHeaders.indexOf("Email ID"),
    departmentIdx : smeRemoveHeaders.indexOf("Department"),
    removeIdx : smeRemoveHeaders.indexOf("Remove?"),
  }  

  clearCellsBelow(accountManagementSheet, 3, 2);
  const departments = getDepartmentList();
  const validation = SpreadsheetApp.newDataValidation().requireValueInList(departments).setAllowInvalid(false).build();

  smeAddData.forEach((r, index) => {
    smeManagementSheet.getRange(index+5, smeAddIndices.emailIdx + 1).clearContent();
    smeManagementSheet.getRange(index+5, smeAddIndices.departmentIdx + 1).setDataValidation(validation);
    smeManagementSheet.getRange(index+5, smeAddIndices.addIdx + 1).setValue(false);
  });

  smeRemoveData.forEach((r, index) => {
    smeManagementSheet.getRange(index+18, smeRemoveIndices.emailIdx + 1).clearContent();
    smeManagementSheet.getRange(index+18, smeRemoveIndices.departmentIdx + 1).setDataValidation(validation);
    smeManagementSheet.getRange(index+18, smeRemoveIndices.removeIdx + 1).setValue(false);
  });

  reviewerAddData.forEach((r, index) => {
    reviewerManagementSheet.getRange(index+4, reviewerAddIndices.emailIdx + 1).clearContent();
    reviewerManagementSheet.getRange(index+4, reviewerAddIndices.departmentIdx + 1).setDataValidation(validation);
    reviewerManagementSheet.getRange(index+4, reviewerAddIndices.addIdx + 1).setValue(false);
  });

  reviewerRemoveData.forEach((r, index) => {
    reviewerManagementSheet.getRange(index+10, reviewerRemoveIndices.emailIdx + 1).clearContent();
    reviewerManagementSheet.getRange(index+10, reviewerRemoveIndices.departmentIdx + 1).setDataValidation(validation);
    reviewerManagementSheet.getRange(index+10, reviewerRemoveIndices.removeIdx + 1).setValue(false);
  });

  reviewerManagementSheet.getRange(15, 3).setDataValidation(validation)
  departmentSheet.getRange(1, 3).setDataValidation(validation)
  clearRowsBelow(departmentSheet, 4)
}


function pushChangesRest(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const reviewerIndexSheet = spreadsheet.getSheetByName("Index");
  
  const dataRangeReviewer = reviewerIndexSheet.getRange(2, 1, reviewerIndexSheet.getLastRow(), reviewerIndexSheet.getLastColumn()).getValues();
  const headerReviewer = dataRangeReviewer[0]
  let dataReviewer = dataRangeReviewer.slice(1);


  const reviewerIndices = {
    srNoIdx : headerReviewer.indexOf("#"),
    emailIdx : headerReviewer.indexOf("QA Reviewer Email"),
    departmentIdx : headerReviewer.indexOf("Department"),
    sheetLinkIdx : headerReviewer.indexOf("Sheet Link"),
  }

  // dataReviewer = dataReviewer.filter(r => r[reviewerIndices.departmentIdx] === "Biology")

  const filteredData = dataReviewer.filter(r => r[reviewerIndices.srNoIdx]!=='' 
                                              && (r[reviewerIndices.emailIdx]!=='' && r[reviewerIndices.emailIdx]!=='sreenjay.sen@upthink.com')
                                              && r[reviewerIndices.departmentIdx]!=='' && r[reviewerIndices.sheetLinkIdx]!=='');

  
  filteredData.slice(30, ).forEach(r => {
    try {
      const link = r[reviewerIndices.sheetLinkIdx];
      const department = r[reviewerIndices.departmentIdx];
      //createBackendReviwerSheetByDepartment(link, department);   
      createBackendReviewerSheet(link,department)
      console.log(link)
    } catch (e) {
      console.error(`Error processing file: ${r[reviewerIndices.sheetLinkIdx]}`);
      console.error(e);
      // Log the error or take other actions if needed
    }
  });           
}


function pushChanges(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const reviewerIndexSheet = spreadsheet.getSheetByName("Index");
  const dataRangeReviewer = reviewerIndexSheet.getRange(2, 1, reviewerIndexSheet.getLastRow(), reviewerIndexSheet.getLastColumn()).getValues();
  const headerReviewer = dataRangeReviewer[0]
  let dataReviewer = dataRangeReviewer.slice(1);


  const reviewerIndices = {
    srNoIdx : headerReviewer.indexOf("#"),
    emailIdx : headerReviewer.indexOf("QA Reviewer Email"),
    departmentIdx : headerReviewer.indexOf("Department"),
    sheetLinkIdx : headerReviewer.indexOf("Sheet Link"),
  }

  // dataReviewer = dataReviewer.filter(r => r[reviewerIndices.departmentIdx] === "Biology")

  const filteredData = dataReviewer.filter(r => r[reviewerIndices.srNoIdx]!=='' 
                                              && (r[reviewerIndices.emailIdx]!=='' && r[reviewerIndices.emailIdx]!=='sreenjay.sen@upthink.com')
                                              && r[reviewerIndices.departmentIdx]!=='' && r[reviewerIndices.sheetLinkIdx]!=='');

  
  filteredData.slice(0, 30).forEach(r => {
    try {
      const link = r[reviewerIndices.sheetLinkIdx];
      const department = r[reviewerIndices.departmentIdx];
      //createBackendReviwerSheetByDepartment(link,department);
      createBackendReviewerSheet(link,department);

      console.log(link)
    } catch (e) {
      console.error(`Error processing file: ${r[reviewerIndices.sheetLinkIdx]}`);
      console.error(e);
      // Log the error or take other actions if needed
    }
  });           
}


function pushChangesSME(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const smeIndexSheet = spreadsheet.getSheetByName("Index - SME");

  const dataRangeSME = smeIndexSheet.getRange(2, 1, smeIndexSheet.getLastRow(), smeIndexSheet.getLastColumn()).getValues();
  const headerSME = dataRangeSME[0], dataSME = dataRangeSME.slice(1);


  const smeIndices = {
    srNoIdx : headerSME.indexOf("#"),
    emailIdx : headerSME.indexOf("SME Email"),
    departmentIdx : headerSME.indexOf("Department"),
    sheetLinkIdx : headerSME.indexOf("Sheet Link"),
  }


  const filteredDataSME = dataSME.filter(r => r[smeIndices.srNoIdx]!=='' 
                                              && r[smeIndices.emailIdx]!=='' 
                                              && r[smeIndices.emailIdx] !== 'sreenjay.sen@upthink.com'
                                              && r[smeIndices.departmentIdx]!=='' && r[smeIndices.sheetLinkIdx]!=='');


  filteredDataSME.forEach(r => {
    try {
      const link = r[smeIndices.sheetLinkIdx];
      const smeEmail = r[smeIndices.emailIdx];
      // console.log(smeEmail);
      // copySheetToAnotherSpreadsheet(link)
      // const fileId = SpreadsheetApp.openByUrl(link).getId()
      // DriveApp.getFileById(fileId).setOwner("automation@upthink.com")
      smeBackend(link, smeEmail);
      copySheetToAnotherSpreadsheet(link)
    } catch(e) {
      console.error(`Error processing file: ${r[smeIndices.sheetLinkIdx]}`);
      console.error(e);
    }  
  });                             
}



function copySheetToAnotherSpreadsheet(destinationSpreadsheeturl) {
  // IDs of the source and destination spreadsheets
  var sourceSpreadsheetId = '17t_fKuXAkXtoVnB4FmCRDIHyJzdiAzIYGbQR0eEdeQM';//'1WDmG0nebB6NgCLUOs41wwutbcbpB9tFlFF4f1t7Xld4';
  
  // Name of the sheet/tab to copy from the source spreadsheet
  var sourceSheetName = 'Rubric'; // Example: 'Sheet1'
  
  // Name for the new sheet in the destination spreadsheet
  var newSheetName = 'Rubric'; // Example: 'Copied Sheet', can be null or omitted
  
  // Open the source spreadsheet and get the sheet by name
  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
  
  // Open the destination spreadsheet
  var destinationSpreadsheet = SpreadsheetApp.openByUrl(destinationSpreadsheeturl);
  
  // Get all the sheets in the destination spreadsheet
  var sheets = destinationSpreadsheet.getSheets();
  
  // Check if a sheet named "Rubric" already exists
  var rubricExists = sheets.some(function(sheet) {
    return sheet.getName() === 'Rubric';
  });
  
  // If a sheet contains "Copy of Rubric" in its name, delete it
  sheets.forEach(function(sheet) {
    if (sheet.getName().includes('Copy of Rubric')) {
      destinationSpreadsheet.deleteSheet(sheet);
    }
  });
  
  // If "Rubric" sheet does not exist, copy it from the source spreadsheet
  if (!rubricExists) {
    var copiedSheet = sourceSheet.copyTo(destinationSpreadsheet);
    
    // If a new sheet name is provided, rename the copied sheet
    if (newSheetName) {
      copiedSheet.setName(newSheetName);
    }
  }
}


function createSheetWithScriptCopy(link) {
  var templateSpreadsheetId = "17t_fKuXAkXtoVnB4FmCRDIHyJzdiAzIYGbQR0eEdeQM";//"1WDmG0nebB6NgCLUOs41wwutbcbpB9tFlFF4f1t7Xld4"; // Replace with the ID of your template Google Sheet

  var templateSheet = SpreadsheetApp.openById(templateSpreadsheetId).getSheetByName("Rubric");
  var destination = SpreadsheetApp.openByUrl(link)
  var destinationSheet = destination.getSheetByName("Rubric");

  destination.deleteSheet(destinationSheet)
  let newSheet = templateSheet.copyTo(destination);
  newSheet.setName("Rubric")
  destination.setActiveSheet(newSheet);
  destination.moveActiveSheet(1);
}


// function copySheetToAnotherSpreadsheet(destinationSpreadsheeturl) {
//   // IDs of the source and destination spreadsheets
//   var sourceSpreadsheetId = '1WDmG0nebB6NgCLUOs41wwutbcbpB9tFlFF4f1t7Xld4';
  
//   // Name of the sheet/tab to copy from the source spreadsheet
//   var sourceSheetName = 'Rubric'; // Example: 'Sheet1'
  
//   // Optional: Name for the new sheet in the destination spreadsheet. If not provided or set to null,
//   // the copied sheet will have the same name as the source sheet appended with " Copy"
//   var newSheetName = 'Rubric'; // Example: 'Copied Sheet', can be null or omitted
  
//   // Open the source spreadsheet and get the sheet by name
//   var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
//   var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
  
//   // Open the destination spreadsheet
//   var destinationSpreadsheet = SpreadsheetApp.openByUrl(destinationSpreadsheeturl);
  
//   if (!destinationSpreadsheet.getSheets().includes(newSheetName)){
//     // Copy the source sheet to the destination spreadsheet
//     var copiedSheet = sourceSheet.copyTo(destinationSpreadsheet);
    
//     // If a new sheet name is provided, rename the copied sheet
//     if (newSheetName) {
//       copiedSheet.setName(newSheetName);
//     }
//   } else{

//   }
// }






