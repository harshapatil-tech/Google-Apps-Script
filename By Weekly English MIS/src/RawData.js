function ClearData() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var rawDataSheet = spreadsheet.getSheetByName("Raw Data");
  clearSheetData(rawDataSheet);
}



/**
 * Updates the "QMS Data" sheet by filtering and mapping data from the "Raw Data" sheet.
 * Clears existing data in "QMS Data" except headers and adds "Person Name" from "Backend" sheet.
 */
function updateData() {
  // Open the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get the relevant sheets
  var rawDataSheet = spreadsheet.getSheetByName("Raw Data");
  var qmsDataSheet = spreadsheet.getSheetByName("QMS Data");
  var backendSheet = spreadsheet.getSheetByName("Backend");
  
  // Validate sheet existence
  if (!rawDataSheet || !qmsDataSheet || !backendSheet) {
    SpreadsheetApp.getUi().alert('One or more required sheets are missing.');
    return;
  }
  
  // Step 1: Clear "QMS Data" Sheet Except Headers
  clearSheetData(qmsDataSheet);
  
  // Step 2: Retrieve "Backend" Data and Create a Mapping
  var backendData = backendSheet.getDataRange().getValues();
  if (backendData.length < 2) { // No data beyond headers
    SpreadsheetApp.getUi().alert('No data available in "Backend" sheet.');
    return;
  }
  
  // Assuming "Backend" sheet has headers "person ID" and "Name"
  var backendHeaders = backendData[0];
  var backendPersonIdIndex = backendHeaders.indexOf("person ID");
  var backendNameIndex = backendHeaders.indexOf("Name");
  
  if (backendPersonIdIndex === -1 || backendNameIndex === -1) {
    SpreadsheetApp.getUi().alert('Please ensure "Backend" sheet has "person ID" and "Name" headers.');
    return;
  }
  
  // Create a map of person ID to Name
  var personIdToNameMap = {};
  for (var i = 1; i < backendData.length; i++) {
    var personId = backendData[i][backendPersonIdIndex];
    var name = backendData[i][backendNameIndex];
    if (personId) { // Ensure personId is not empty
      personIdToNameMap[personId] = name;
    }
  }
  
  // Step 3: Get all data from "Raw Data" sheet
  var rawData = rawDataSheet.getDataRange().getValues();
  
  if (rawData.length < 2) { // No data beyond headers
    SpreadsheetApp.getUi().alert('No data available in "Raw Data" to update.');
    return;
  }
  
  // Get headers from "Raw Data" and "QMS Data"
  var rawHeaders = rawData[0];
  var qmsHeaders = qmsDataSheet.getDataRange().getValues()[0];
  
  // Define the header mapping
  var headerMap = {
    "personId": "Person ID",
    "essayId": "Essay ID",
    "credName": "Account No",
    "createdAt_IST": ["IST Time", "IST Date"],
    "createdAt_EST": ["EST Time", "EST Date"]
  };
  
  // Ensure "Person Name" exists in "QMS Data" headers
  if (!qmsHeaders.includes("Person Name")) {
    SpreadsheetApp.getUi().alert('"QMS Data" sheet must contain a "Person Name" column.');
    return;
  }
  
  // Create a mapping from QMS Data headers to their column indices
  var qmsHeaderIndex = {};
  qmsHeaders.forEach(function(header, index) {
    qmsHeaderIndex[header] = index;
  });
  
  // Create a mapping from Raw Data headers to their column indices
  var rawHeaderIndex = {};
  rawHeaders.forEach(function(header, index) {
    rawHeaderIndex[header] = index;
  });
  
  // Prepare an array to hold the rows to be appended to "QMS Data"
  var rowsToAppend = [];
  
  // Get date range for the previous two full weeks
  var dateRange = calculatePreviousTwoWeeksDateRange();
  console.log(dateRange)
  var startDate = dateRange.startDate;
  var endDate = dateRange.endDate;
  
  // Iterate through each row in "Raw Data" starting from row 2
  for (var i = 1; i < rawData.length; i++) {
    var row = rawData[i];
    
    // Apply filters:
    // userAction == "checked_in" AND clientName == "BrainFuse" AND (createdAt_IST OR createdAt_EST is in the previous two weeks)
    var userAction = row[rawHeaderIndex["userAction"]];
    var clientName = row[rawHeaderIndex["clientName"]];
    var createdAtIST = parseDate(row[rawHeaderIndex["createdAt_IST"]]);
    var createdAtEST = parseDate(row[rawHeaderIndex["createdAt_EST"]]);
    var personId = row[rawHeaderIndex["personId"]];
    
    // Check userAction and clientName
    if (userAction !== "checked_in" || clientName !== "BrainFuse") {
      continue; // Skip this row
    }
    
    // Check if either createdAt_IST or createdAt_EST is in the previous two weeks
    var isISTInRange = isDateInRange(createdAtIST, startDate, endDate);
    var isESTInRange = isDateInRange(createdAtEST, startDate, endDate);
    
    if (!isISTInRange && !isESTInRange) {
      continue; // Skip this row
    }
    
    // Map the fields to "QMS Data" headers
    var mappedRow = new Array(qmsHeaders.length).fill(""); // Initialize with empty strings
    
    // Iterate through "QMS Data" headers and populate the mappedRow accordingly
    qmsHeaders.forEach(function(qmsHeader) {
      // Find the corresponding Raw Data header
      var value = "";
      
      // Find which Raw Data field maps to this QMS header
      for (var key in headerMap) {
        if (headerMap.hasOwnProperty(key)) {
          var mappedHeaders = headerMap[key];
          
          if (Array.isArray(mappedHeaders)) {
            // Handle split fields (Date and Time)
            if (mappedHeaders.includes(qmsHeader)) {
              var dateValue = parseDate(row[rawHeaderIndex[key]]);
              if (qmsHeader.includes("Time")) {
                // Extract time portion
                value = dateValue ? Utilities.formatDate(dateValue, spreadsheet.getSpreadsheetTimeZone(), "HH:mm:ss") : "";
              } else if (qmsHeader.includes("Date")) {
                // Extract date portion
                value = dateValue ? Utilities.formatDate(dateValue, spreadsheet.getSpreadsheetTimeZone(), "yyyy-MM-dd") : "";
              }
              break;
            }
          } else {
            if (mappedHeaders === qmsHeader) {
              rawHeader = key;
              value = row[rawHeaderIndex[key]];
              break;
            }
          }
        }
      }
      
      // Assign the value to the appropriate column in mappedRow
      var qmsColIndex = qmsHeaderIndex[qmsHeader];
      if (qmsColIndex !== undefined) {
        mappedRow[qmsColIndex] = value;
      }
    });
    
    // After mapping, add "Person Name" based on "Person ID"
    var personName = personIdToNameMap[personId] || ""; // Default to empty string if not found
    var personNameColIndex = qmsHeaderIndex["Person Name"];
    if (personNameColIndex !== undefined) {
      mappedRow[personNameColIndex] = personName;
    }
    
    // Add the mapped row to the array
    rowsToAppend.push(mappedRow);
  }
  
  if (rowsToAppend.length === 0) {
    SpreadsheetApp.getUi().alert('No data matched the specified filters.');
    return;
  }
  
  // Append the rows to "QMS Data" sheet
  qmsDataSheet.getRange(qmsDataSheet.getLastRow() + 1, 1, rowsToAppend.length, qmsHeaders.length).setValues(rowsToAppend);
  
  // Optionally, notify the user
  SpreadsheetApp.getUi().alert(rowsToAppend.length + ' rows have been successfully updated to "QMS Data".');
}
