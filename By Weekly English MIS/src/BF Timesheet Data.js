/**
 * Updates the "BF Timesheet Data" sheet by filtering and mapping data from the "Summary" sheet
 * located in a separate input spreadsheet. Clears existing data in "BF Timesheet Data"
 * except for headers and appends the filtered data.
 */
function updateBFTimesheetData() {
  // ====== Configuration ======
  // Replace the below string with your external input spreadsheet's ID
  var externalSpreadsheetId = '1XH5eGI4thuz-yAsHDeVbI90KFX2AWee3w556yUcGe6I';
  // ============================

  // Open the active spreadsheet (destination)
  var destinationSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get the destination sheet "BF Timesheet Data"
  var bfTimesheetSheet = destinationSpreadsheet.getSheetByName("BF Timesheet Data");
  
  // Validate the existence of the destination sheet
  if (!bfTimesheetSheet) {
    console.log('Sheet "BF Timesheet Data" not found in the active spreadsheet.');
    return;
  }
  
  // Step 1: Clear "BF Timesheet Data" Sheet Except Headers
  clearSheetData(bfTimesheetSheet);
  
  // Step 2: Access the External Spreadsheet and Get the "Summary" Sheet
  var externalSpreadsheet;
  try {
    externalSpreadsheet = SpreadsheetApp.openById(externalSpreadsheetId);
  } catch (e) {
    console.log('Unable to open external spreadsheet. Please check the Spreadsheet ID.');
    return;
  }
  
  var summarySheet = externalSpreadsheet.getSheetByName("Summary");
  
  if (!summarySheet) {
    console.log('Sheet "Summary" not found in the external spreadsheet.');
    return;
  }
  
  // Step 3: Get all data from "Summary" sheet
  var summaryData = summarySheet.getDataRange().getValues();
  
  if (summaryData.length < 2) { // No data beyond headers
    console.log('No data available in "Summary" sheet to update.');
    return;
  }
  
  // Get headers from "Summary" and "BF Timesheet Data"
  var summaryHeaders = summaryData[0];
  var bfTimesheetHeaders = bfTimesheetSheet.getDataRange().getValues()[0];
  
  // Define the header mapping
  var headerMap = {
    "Comments": "BF Essay ID",
    "Account No.": "BF Account No.",
    "Start Date": "Timesheet Date",
    "Start Time": "BF EST Time"
  };
  
  // Create a mapping from Summary headers to their column indices
  var summaryHeaderIndex = {};
  summaryHeaders.forEach(function(header, index) {
    summaryHeaderIndex[header] = index;
  });
  
  // Create a mapping from BF Timesheet headers to their column indices
  var bfTimesheetHeaderIndex = {};
  bfTimesheetHeaders.forEach(function(header, index) {
    bfTimesheetHeaderIndex[header] = index;
  });
  
  // Prepare an array to hold the rows to be appended to "BF Timesheet Data"
  var rowsToAppend = [];
  
  // Get date range for the previous two full weeks
  var dateRange = calculatePreviousTwoWeeksDateRange();
  var startDate = dateRange.startDate;
  var endDate = dateRange.endDate;
  console.log(startDate, endDate)
  // Iterate through each row in "Summary" starting from row 2
  for (var i = 1; i < summaryData.length; i++) {
    var row = summaryData[i];
    
    // Apply filters:
    // Department == "English" AND Comments != Blank AND Start Date is in the previous two weeks
    var department = row[summaryHeaderIndex["Department"]];
    var comments = row[summaryHeaderIndex["Comments"]];
    var timesheetDate = parseDate(row[summaryHeaderIndex["Start Date"]]);

    // Check if timesheetDate is within the previous two weeks
    var isDateInRangeFlag = isDateInRange(timesheetDate, startDate, endDate);

    if (department !== "English" || !comments || comments.toString().trim() === "" || !isDateInRangeFlag) {
      continue; // Skip this row
    }
    
    // Extract the last 7 characters of the "Comments" string for BF Essay ID
    var commentsStr = comments.toString();
    var bfEssayId = commentsStr.length >= 7 ? commentsStr.slice(-7) : commentsStr;
    
    // Map the required fields
    var bfAccountNo = row[summaryHeaderIndex["Account No."]];
    var bfESTTime = row[summaryHeaderIndex["Start Time"]];
    
    // Initialize a mapped row with empty strings
    var mappedRow = new Array(bfTimesheetHeaders.length).fill("");
    
    // Assign values based on header mapping
    // BF Essay ID
    var bfEssayIdCol = bfTimesheetHeaderIndex["BF Essay ID"];
    if (bfEssayIdCol !== -1) {
      mappedRow[bfEssayIdCol] = bfEssayId;
    }
    
    // BF Account No.
    var bfAccountNoCol = bfTimesheetHeaderIndex["BF Account No."];
    if (bfAccountNoCol !== -1) {
      mappedRow[bfAccountNoCol] = bfAccountNo;
    }
    
    // Timesheet Date
    var timesheetDateCol = bfTimesheetHeaderIndex["Timesheet Date"];
    if (timesheetDateCol !== -1) {
      mappedRow[timesheetDateCol] = timesheetDate;
    }
    
    // BF EST Time
    var bfESTTimeCol = bfTimesheetHeaderIndex["BF EST Time"];
    if (bfESTTimeCol !== -1) {
      mappedRow[bfESTTimeCol] = bfESTTime;
    }
    
    // Add the mapped row to the array
    rowsToAppend.push(mappedRow);
  }
  
  if (rowsToAppend.length === 0) {
    console.log('No data matched the specified filters in the "Summary" sheet.');
    return;
  }
  
  // Append the rows to "BF Timesheet Data" sheet
  bfTimesheetSheet.getRange(bfTimesheetSheet.getLastRow() + 1, 1, rowsToAppend.length, bfTimesheetHeaders.length).setValues(rowsToAppend);
  
  // Optionally, notify the user
  console.log(rowsToAppend.length + ' rows have been successfully updated to "BF Timesheet Data".');

  // Call subsequent functions if needed
  collateRAWDataBFQMS();  // calling the next function since it also has to run in the same biweekly trigger
  

  // Following rows are to update the MIS Output sheet tabs. This ensures that the update only happens once the input sheet codes are completed.
  updateMISData();

  filterAndCopyDataToOutputSheet();

  processManualEntryData();
}
