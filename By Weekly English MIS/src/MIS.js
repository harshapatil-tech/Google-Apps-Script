/**
 * Updates the "MIS" sheet in the Output spreadsheet by calculating counts and sums
 * from the Input spreadsheet's sheets within the previous two full weeks.
 * Uses EST Date for comparison with BF Timesheet Data and IST Date for comparison with Manual Entry data.
 * Excludes entries where "Person ID" equals "a8bf99b1-79bf-47a9-87f5-69340111e861" in the QMS Data sheet.
 */
function updateMISData() {
  // Open the Input and Output spreadsheets by their IDs
  var inputSpreadsheet = SpreadsheetApp.openById('1m1x5OxS2_P80Hr8tH-Rab41bNYWE2VOk56NJTlNLbKg');
  var outputSpreadsheet = SpreadsheetApp.openById('1JyaXCYgePPzYRpoh6PMYYiu7MFxr-khPeVw6WU2jzhc');
  
  // Select the 'MIS' sheet in the Output spreadsheet
  var outputSheet = outputSpreadsheet.getSheetByName('MIS');
  if (!outputSheet) {
    console.log('Sheet "MIS" not found in the Output spreadsheet.');
    return;
  }
  
  // Get date range for the previous two full weeks
  var dateRange = calculatePreviousTwoWeeksDateRange();
  var startDate = dateRange.startDate;
  var endDate = dateRange.endDate;
  
  // Variables to hold counts and sums
  var estDateCount = 0; // For QMS Data using EST Date
  var istDateCount = 0; // For QMS Data using IST Date
  var timesheetCount = 0; // For BF Timesheet Data
  var essaysSum = 0; // For Manual Entry data
  
  // Open necessary sheets in the Input spreadsheet
  var qmsSheet = inputSpreadsheet.getSheetByName('QMS Data');
  var timesheetSheet = inputSpreadsheet.getSheetByName('BF Timesheet Data');
  var manualEntrySheet = inputSpreadsheet.getSheetByName('Manual Entry data');
  
  if (!qmsSheet || !timesheetSheet || !manualEntrySheet) {
    console.log('One or more required sheets not found in the Input spreadsheet.');
    return;
  }
  
  // Process 'QMS Data' sheet
  var qmsDataRange = qmsSheet.getDataRange();
  var qmsData = qmsDataRange.getValues(); // Get all data including headers
  
  if (qmsData.length < 2) {
    console.log('No data available in "QMS Data" to process.');
    return;
  }
  
  // Get headers and map column indices
  var qmsHeaders = qmsData[0];
  var qmsHeaderIndex = {};
  qmsHeaders.forEach(function(header, index) {
    qmsHeaderIndex[header.trim()] = index;
  });
  
  // Ensure required headers exist
  if (!("EST Date" in qmsHeaderIndex) || !("IST Date" in qmsHeaderIndex) || !("Person ID" in qmsHeaderIndex)) {
    console.log('Missing required headers in "QMS Data" sheet.');
    return;
  }
  
  // Get indices of required columns
  var estDateIndex = qmsHeaderIndex["EST Date"];
  var istDateIndex = qmsHeaderIndex["IST Date"];
  var personIdIndex = qmsHeaderIndex["Person ID"];
  
  // Loop through the data rows
  for (var i = 1; i < qmsData.length; i++) {
    var row = qmsData[i];
    var personId = row[personIdIndex];
    
    // Exclude entries with the specified Person ID
    if (personId === 'a8bf99b1-79bf-47a9-87f5-69340111e861') {
      continue; // Skip this entry
    }
    if (personId === 'd1933dda-3071-7048-8bf1-343304792b13') {
      continue; // Skip this entry
    }
    
    // Process EST Date for comparison with BF Timesheet Data
    var estDateValue = parseDate(row[estDateIndex]);
    if (estDateValue && isDateInRange(estDateValue, startDate, endDate)) {
      estDateCount++;
    }
    
    // Process IST Date for comparison with Manual Entry data
    var istDateValue = parseDate(row[istDateIndex]);
    if (istDateValue && isDateInRange(istDateValue, startDate, endDate)) {
      istDateCount++;
    }
  }
  
  // Process 'BF Timesheet Data' sheet
  var timesheetDataRange = timesheetSheet.getDataRange();
  var timesheetData = timesheetDataRange.getValues();
  
  if (timesheetData.length < 2) {
    console.log('No data available in "BF Timesheet Data" to process.');
    return;
  }
  
  // Get headers and map column indices
  var timesheetHeaders = timesheetData[0];
  var timesheetHeaderIndex = {};
  timesheetHeaders.forEach(function(header, index) {
    timesheetHeaderIndex[header.trim()] = index;
  });
  
  // Ensure required headers exist
  if (!("Timesheet Date" in timesheetHeaderIndex)) {
    console.log('Missing required header "Timesheet Date" in "BF Timesheet Data" sheet.');
    return;
  }
  
  // Get index of 'Timesheet Date' column
  var timesheetDateIndex = timesheetHeaderIndex["Timesheet Date"];
  
  // Loop through the data rows
  for (var i = 1; i < timesheetData.length; i++) {
    var row = timesheetData[i];
    var dateValue = parseDate(row[timesheetDateIndex]);
    if (dateValue && isDateInRange(dateValue, startDate, endDate)) {
      timesheetCount++;
    }
  }
  
  // Process 'Manual Entry data' sheet
  var manualDataRange = manualEntrySheet.getDataRange();
  var manualData = manualDataRange.getValues();
  
  if (manualData.length < 2) {
    console.log('No data available in "Manual Entry data" to process.');
    return;
  }
  
  // Get headers and map column indices
  var manualHeaders = manualData[0];
  var manualHeaderIndex = {};
  manualHeaders.forEach(function(header, index) {
    manualHeaderIndex[header.trim()] = index;
  });
  
  // Ensure required headers exist
  if (!("# Essays" in manualHeaderIndex)) {
    console.log('Missing required header "# Essays" in "Manual Entry data" sheet.');
    return;
  }
  
  // Get index of '# Essays' column
  var essaysIndex = manualHeaderIndex["# Essays"];
  
  // Loop through the data rows
  for (var i = 1; i < manualData.length; i++) {
    var row = manualData[i];
    var essayValue = parseFloat(row[essaysIndex]);
    if (!isNaN(essayValue)) {
      essaysSum += essayValue;
    }
  }
  
  // Calculate differences
  var differenceQMS_BF = estDateCount - timesheetCount;
  var differenceQMS_Manual = istDateCount - essaysSum;
  
  // Update the 'MIS' sheet
  // Prepare the ranges and values to update
  var rangesToUpdate = [
    {range: 'C7', value: estDateCount}, // #QMS for comparison with # BF (EST Date)
    {range: 'C15', value: istDateCount}, // #QMS for comparison with # Manual (IST Date)
    {range: 'H7', value: timesheetCount}, // # BF
    {range: 'H15', value: essaysSum}, // # Manual
    {range: 'M7', value: differenceQMS_BF},
    {range: 'M15', value: differenceQMS_Manual}
  ];
  
  // Update the cells
  for (var i = 0; i < rangesToUpdate.length; i++) {
    outputSheet.getRange(rangesToUpdate[i].range).setValue(rangesToUpdate[i].value);
  }
  
  console.log('MIS sheet has been updated with the latest counts and sums.');
}
