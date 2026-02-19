/**
 * Processes data from "Manual Entry data" and "QMS Data" tabs,
 * writes to "Further Analysis sheet (QMS-Manual)" in the Output Sheet,
 * and sorts the data based on the "Difference" column.
 * Incorporates date filtering for the previous two full weeks and uses helper functions.
 */
function processManualEntryData() {
  // Replace with the actual ID of the Output Sheet
  var outputSpreadsheetId = '1JyaXCYgePPzYRpoh6PMYYiu7MFxr-khPeVw6WU2jzhc'; // Set the actual ID of the Output Sheet

  // Open the input (current) spreadsheet
  var inputSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Open the Output Sheet (another spreadsheet)
  var outputSpreadsheet = SpreadsheetApp.openById(outputSpreadsheetId);

  // Get the relevant sheets
  var manualEntrySheet = inputSpreadsheet.getSheetByName("Manual Entry data");
  var qmsDataSheet = inputSpreadsheet.getSheetByName("QMS Data");
  var outputSheet = outputSpreadsheet.getSheetByName("Further Analysis sheet (QMS-Manual)");

  // Validate input sheets
  if (!manualEntrySheet) {
    console.log('Sheet "Manual Entry data" not found in the current spreadsheet.');
    return;
  }
  if (!qmsDataSheet) {
    console.log('Sheet "QMS Data" not found in the current spreadsheet.');
    return;
  }

  // Check if the output sheet exists in the output spreadsheet
  if (!outputSheet) {
    // Create the sheet if it doesn't exist in the output spreadsheet
    outputSheet = outputSpreadsheet.insertSheet("Further Analysis sheet (QMS-Manual)");
  } else {
    // Clear existing data except headers in the output spreadsheet
    clearSheetData(outputSheet);
  }

  // Get date range for the previous two full weeks
  var dateRange = calculatePreviousTwoWeeksDateRange();
  var startDate = dateRange.startDate;
  var endDate = dateRange.endDate;

  // Read data from "Manual Entry data" sheet
  var manualData = manualEntrySheet.getDataRange().getValues();
  if (manualData.length < 2) {
    console.log('No data available in "Manual Entry data" to process.');
    return;
  }

  // Get headers from "Manual Entry data"
  var manualHeaders = manualData[0];
  var manualHeaderIndex = {};
  manualHeaders.forEach(function(header, index) {
    manualHeaderIndex[header.trim()] = index;
  });

  // Ensure required headers exist in "Manual Entry data"
  if (!("Name" in manualHeaderIndex) || !("# Essays" in manualHeaderIndex)) {
    console.log('Missing required headers in "Manual Entry data" sheet.');
    return;
  }

  // var dateIndexManual = manualHeaderIndex["Date"];

  // Build a map of manual entries: { Person Name: # Essays - Manual entry }
  var manualEntries = {};
  for (var i = 1; i < manualData.length; i++) {
    var row = manualData[i];
    var name = row[manualHeaderIndex["Name"]] || "";
    var numEssays = parseFloat(row[manualHeaderIndex["# Essays"]]) || 0;
    // var dateValue = parseDate(row[dateIndexManual]);
    // if (dateValue && isDateInRange(dateValue, startDate, endDate)) {
      if (name) {
        if (manualEntries[name]) {
          manualEntries[name] += numEssays;
        } else {
          manualEntries[name] = numEssays;
        }
      }
    // }
  }

  // Read data from "QMS Data" sheet
  var qmsData = qmsDataSheet.getDataRange().getValues();
  if (qmsData.length < 2) {
    console.log('No data available in "QMS Data" to process.');
    return;
  }

  // Get headers from "QMS Data"
  var qmsHeaders = qmsData[0];
  var qmsHeaderIndex = {};
  qmsHeaders.forEach(function(header, index) {
    qmsHeaderIndex[header.trim()] = index;
  });

  // Ensure required headers exist in "QMS Data"
  if (!("Person Name" in qmsHeaderIndex) || !("IST Date" in qmsHeaderIndex)) {
    console.log('Missing required headers in "QMS Data" sheet.');
    return;
  }

  var personNameIndex = qmsHeaderIndex["Person Name"];
  var istDateIndex = qmsHeaderIndex["IST Date"];

  // Build a map of counts for each person within the date range
  var countsByPerson = {};

  // Process "QMS Data" rows
  for (var i = 1; i < qmsData.length; i++) {
    var row = qmsData[i];
    var personName = row[personNameIndex];
    var istDateValue = parseDate(row[istDateIndex]);

    if (istDateValue && isDateInRange(istDateValue, startDate, endDate)) {
      if (personName) {
        if (!countsByPerson[personName]) {
          countsByPerson[personName] = 1;
        } else {
          countsByPerson[personName]++;
        }
      }
    }
  }

  // Build the output data array
  // Assume that the output sheet has headers: "Person Name", "# Essays - Manual entry", "# Essays - QMS", "Difference"
  // We'll read the headers from the output sheet to ensure alignment

  var outputHeaders = outputSheet.getRange(1, 1, 1, outputSheet.getLastColumn()).getValues()[0];
  var outputHeaderIndex = {};
  outputHeaders.forEach(function(header, index) {
    outputHeaderIndex[header.trim()] = index;
  });

  // Ensure required headers exist in output sheet
  var requiredOutputHeaders = ["Person Name", "# Essays - Manual entry", "# Essays - QMS", "Difference"];
  for (var i = 0; i < requiredOutputHeaders.length; i++) {
    if (!(requiredOutputHeaders[i] in outputHeaderIndex)) {
      console.log('Missing required header "' + requiredOutputHeaders[i] + '" in output sheet.');
      return;
    }
  }

  var outputData = [];

  // Build data rows
  // Include all names from manualEntries and countsByPerson
  var allNames = new Set(Object.keys(manualEntries).concat(Object.keys(countsByPerson)));

  allNames.forEach(function(name) {
    var numEssaysManual = manualEntries[name] || 0;
    var numEssaysQMS = countsByPerson[name] || 0;
    var difference = numEssaysQMS - numEssaysManual;

    var row = [];
    row[outputHeaderIndex["Person Name"]] = name;
    row[outputHeaderIndex["# Essays - Manual entry"]] = numEssaysManual;
    row[outputHeaderIndex["# Essays - QMS"]] = numEssaysQMS;
    row[outputHeaderIndex["Difference"]] = difference;

    outputData.push(row);
  });

  // Sort outputData based on "Difference" (ascending order)
  outputData.sort(function(a, b) {
    return a[outputHeaderIndex["Difference"]] - b[outputHeaderIndex["Difference"]];
  });

  // Ensure each row has values for all columns
  var numColumns = outputHeaders.length;
  for (var i = 0; i < outputData.length; i++) {
    var row = outputData[i];
    for (var j = 0; j < numColumns; j++) {
      if (row[j] === undefined) {
        row[j] = '';
      }
    }
  }

  // Write outputData to outputSheet starting from row 2 (since headers are in row 1)
  if (outputData.length > 0) {
    // Clear existing data from row 2 onwards
    var lastRow = outputSheet.getLastRow();
    if (lastRow > 1) {
      outputSheet.getRange(2, 1, lastRow - 1, outputSheet.getLastColumn()).clearContent();
    }
    outputSheet.getRange(2, 1, outputData.length, numColumns).setValues(outputData);
  } else {
    // No data to write
    console.log('No data to write to output sheet.');
  }

  // Optionally, notify the user
  console.log('Data has been processed and written to "Further Analysis sheet (QMS-Manual)" in the Output Sheet.');
}
