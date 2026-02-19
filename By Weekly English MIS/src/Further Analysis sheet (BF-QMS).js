/**
 * Filters data from "RAW Collated data (BF-QMS)" in the current spreadsheet
 * and copies it to "Further Analysis sheet (BF-QMS)" in the Output Sheet spreadsheet.
 * Excludes rows where "Person Name" equals "Operations".
 * Ensures that the "EST Time difference" field is copied properly with Duration format.
 * Uses only the provided helper functions.
 */
function filterAndCopyDataToOutputSheet() {
  // ====== Configuration ======
  var outputSpreadsheetId = '1JyaXCYgePPzYRpoh6PMYYiu7MFxr-khPeVw6WU2jzhc';
  // ===========================

  // Open the current spreadsheet (input spreadsheet)
  var inputSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Open the output spreadsheet
  var outputSpreadsheet = SpreadsheetApp.openById(outputSpreadsheetId);

  // Get the relevant sheets
  var rawCollatedSheet = inputSpreadsheet.getSheetByName("RAW Collated data (BF-QMS)");
  var furtherAnalysisSheet = outputSpreadsheet.getSheetByName("Further Analysis sheet (BF-QMS)");

  // Validate sheet existence
  if (!rawCollatedSheet) {
    console.log('Sheet "RAW Collated data (BF-QMS)" not found in the input spreadsheet.');
    return;
  }

  if (!furtherAnalysisSheet) {
    // Create the sheet if it doesn't exist
    furtherAnalysisSheet = outputSpreadsheet.insertSheet("Further Analysis sheet (BF-QMS)");
  } else {
    // Clear existing data except headers
    clearSheetData(furtherAnalysisSheet);
  }

  // Get data from "RAW Collated data (BF-QMS)"
  var dataRange = rawCollatedSheet.getDataRange();
  var dataValues = dataRange.getValues();

  if (dataValues.length < 2) {
    console.log('No data available in "RAW Collated data (BF-QMS)" to process.');
    return;
  }

  var headers = dataValues[0];
  var headerIndex = {};
  headers.forEach(function(header, index) {
    headerIndex[header.trim()] = index;
  });

  // Ensure required headers exist
  var requiredHeaders = [
    "Essay ID",
    "QMS Account no",
    "QMS EST Date",
    "QMS EST Time",
    "BF Account No",
    "BF EST Date",
    "BF EST Time",
    "Account no Match?",
    "EST Date Match",
    "Absolute Time difference less than 3?",
    "Person ID",
    "Person Name",
    "EST Time difference"
  ];
  for (var i = 0; i < requiredHeaders.length; i++) {
    if (!(requiredHeaders[i] in headerIndex)) {
      console.log('Missing required header "' + requiredHeaders[i] + '" in "RAW Collated data (BF-QMS)" sheet.');
      return;
    }
  }

  // Get date range for the previous two full weeks
  var dateRangeObj = calculatePreviousTwoWeeksDateRange();
  var startDate = dateRangeObj.startDate;
  var endDate = dateRangeObj.endDate;

  // Get indices of relevant columns
  var personNameIndex = headerIndex["Person Name"];

  // Prepare filtered data
  var filteredData = [];
  filteredData.push(headers); // Include headers

  // Loop through data rows
  for (var i = 1; i < dataValues.length; i++) {
    var row = dataValues[i];

    var accountNoMatch = (row[headerIndex["Account no Match?"]] || "").toString().trim();
    var estDateMatch = (row[headerIndex["EST Date Match"]] || "").toString().trim();
    var absTimeDiffLessThan3 = (row[headerIndex["Absolute Time difference less than 3?"]] || "").toString().trim();
    var personName = (row[personNameIndex] || "").toString().trim();

    // Exclude rows where "Person Name" equals "Operations"
    if (personName === "Operations") {
      continue;
    }

    // Parse QMS EST Date and BF EST Date
    var qmsEstDateStr = row[headerIndex["QMS EST Date"]];
    var bfEstDateStr = row[headerIndex["BF EST Date"]];
    var qmsEstDate = parseDate(qmsEstDateStr);
    var bfEstDate = parseDate(bfEstDateStr);

    // Check if either QMS EST Date or BF EST Date is within the date range
    var isQmsDateInRange = qmsEstDate && isDateInRange(qmsEstDate, startDate, endDate);
    var isBfDateInRange = bfEstDate && isDateInRange(bfEstDate, startDate, endDate);

    // If neither date is in range, skip this row
    if (!isQmsDateInRange && !isBfDateInRange) {
      continue;
    }

    // Check if any of the conditions are met (logical OR)
    if (accountNoMatch === "N" || estDateMatch === "N" || absTimeDiffLessThan3 === "N") {
      filteredData.push(row);
    }
  }

  // Write the filtered data to "Further Analysis sheet (BF-QMS)" in the output spreadsheet
  if (filteredData.length > 1) {
    var numRows = filteredData.length;
    var numCols = filteredData[0].length;
    furtherAnalysisSheet.getRange(1, 1, numRows, numCols).setValues(filteredData);

    // Set number formats for time, date, and duration columns
    // Time format: "HH:mm:ss"
    var timeFormat = "HH:mm:ss";
    // Date format: "yyyy-MM-dd"
    var dateFormat = "yyyy-MM-dd";
    // Duration format handling negative durations
    var durationFormat = "[h]:mm:ss;-[h]:mm:ss";

    // Get indices of time and date columns (1-based indexing for Sheets)
    var qmsEstTimeColumn = headerIndex["QMS EST Time"] + 1;
    var bfEstTimeColumn = headerIndex["BF EST Time"] + 1;
    var estTimeDifferenceColumn = headerIndex["EST Time difference"] + 1;

    var qmsEstDateColumn = headerIndex["QMS EST Date"] + 1;
    var bfEstDateColumn = headerIndex["BF EST Date"] + 1;

    // Apply number formats (excluding header row)
    var dataRowCount = numRows - 1;
    if (dataRowCount > 0) {
      // Apply time format to time columns
      furtherAnalysisSheet.getRange(2, qmsEstTimeColumn, dataRowCount, 1).setNumberFormat(timeFormat);
      furtherAnalysisSheet.getRange(2, bfEstTimeColumn, dataRowCount, 1).setNumberFormat(timeFormat);

      // Apply duration format to "EST Time difference" column
      furtherAnalysisSheet.getRange(2, estTimeDifferenceColumn, dataRowCount, 1).setNumberFormat(durationFormat);

      // Apply date format to date columns
      furtherAnalysisSheet.getRange(2, qmsEstDateColumn, dataRowCount, 1).setNumberFormat(dateFormat);
      furtherAnalysisSheet.getRange(2, bfEstDateColumn, dataRowCount, 1).setNumberFormat(dateFormat);
    }
  } else {
    // No data to write, write headers only
    furtherAnalysisSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    console.log('No rows matched the criteria.');
    return;
  }

  // Optionally, notify the user
  console.log('Data has been successfully filtered and copied to "Further Analysis sheet (BF-QMS)" in the output spreadsheet.');
}
