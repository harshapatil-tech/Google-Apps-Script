/**
 * Collates data from "QMS Data" and "BF Timesheet Data" into "RAW Collated data (BF-QMS)"
 * based on specified filters and performs necessary calculations.
 * Replaces "Present in QMS?" with "Person ID" and adds "Person Name" from "Backend" sheet.
 */
function collateRAWDataBFQMS() {
  // Open the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Get the relevant sheets
  var qmsDataSheet = spreadsheet.getSheetByName("QMS Data");
  var bfTimesheetSheet = spreadsheet.getSheetByName("BF Timesheet Data");
  var rawCollatedSheet = spreadsheet.getSheetByName("RAW Collated data (BF-QMS)");
  var backendSheet = spreadsheet.getSheetByName("Backend"); // New sheet

  // Validate sheet existence
  if (!qmsDataSheet || !bfTimesheetSheet || !rawCollatedSheet || !backendSheet) {
    console.log('One or more required sheets are missing.');
    return;
  }

  // Step 1: Clear "RAW Collated data (BF-QMS)" Sheet Except Headers
  clearSheetData(rawCollatedSheet);

  // Get date range for the previous two full weeks
  var dateRange = calculatePreviousTwoWeeksDateRange();
  var startDate = dateRange.startDate;
  var endDate = dateRange.endDate;
 

  // Step 2: Get data from "QMS Data" with the specified filter
  var qmsData = qmsDataSheet.getDataRange().getValues();

  if (qmsData.length < 2) {
    console.log('No data available in "QMS Data" to process.');
    return;
  }

  var qmsHeaders = qmsData[0];
  var qmsHeaderIndex = {};
  qmsHeaders.forEach(function(header, index) {
    qmsHeaderIndex[header.trim()] = index;
  });

  // Ensure required headers exist
  var requiredQMSHeaders = ["Essay ID", "Account No", "EST Date", "EST Time", "Person ID"];
  for (var i = 0; i < requiredQMSHeaders.length; i++) {
    if (!(requiredQMSHeaders[i] in qmsHeaderIndex)) {
      console.log('Missing required header "' + requiredQMSHeaders[i] + '" in "QMS Data" sheet.');
      return;
    }
  }

  // Filter QMS Data for dates in the previous two full weeks
  var qmsFilteredData = [];

  for (var i = 1; i < qmsData.length; i++) {
    var row = qmsData[i];
    var estDate = parseDate(row[qmsHeaderIndex["EST Date"]]);
    if (estDate && isDateInRange(estDate, startDate, endDate)) {
      qmsFilteredData.push({
        "Essay ID": row[qmsHeaderIndex["Essay ID"]],
        "QMS Account no": row[qmsHeaderIndex["Account No"]],
        "QMS EST Date": estDate,
        "QMS EST Time": row[qmsHeaderIndex["EST Time"]],
        "Person ID": row[qmsHeaderIndex["Person ID"]]
      });
    }
  }

  // Step 3: Get data from "BF Timesheet Data" with the specified filter
  var bfData = bfTimesheetSheet.getDataRange().getValues();

  if (bfData.length < 2) {
    console.log('No data available in "BF Timesheet Data" to process.');
    return;
  }

  var bfHeaders = bfData[0];
  var bfHeaderIndex = {};
  bfHeaders.forEach(function(header, index) {
    bfHeaderIndex[header.trim()] = index;
  });

  // Ensure required headers exist
  var requiredBFHeaders = ["BF Essay ID", "BF Account No.", "Timesheet Date", "BF EST Time"];
  for (var i = 0; i < requiredBFHeaders.length; i++) {
    if (!(requiredBFHeaders[i] in bfHeaderIndex)) {
      console.log('Missing required header "' + requiredBFHeaders[i] + '" in "BF Timesheet Data" sheet.');
      return;
    }
  }

  // Filter BF Timesheet Data for dates in the previous two full weeks and create a map based on BF Essay ID
  var bfDataMap = {};

  for (var i = 1; i < bfData.length; i++) {
    var row = bfData[i];
    var bfEssayId = row[bfHeaderIndex["BF Essay ID"]];
    var timesheetDate = parseDate(row[bfHeaderIndex["Timesheet Date"]]);

    if (bfEssayId && timesheetDate && isDateInRange(timesheetDate, startDate, endDate)) {
      bfDataMap[bfEssayId] = {
        "Essay ID": bfEssayId, // Aligning "BF Essay ID" with "Essay ID"
        "BF Account No": row[bfHeaderIndex["BF Account No."]],
        "BF EST Date": timesheetDate,
        "BF EST Time": row[bfHeaderIndex["BF EST Time"]]
      };
    }
  }

  // Step 4: Read "Backend" Data and create a mapping of person ID to Name
  var backendData = backendSheet.getDataRange().getValues();
  if (backendData.length < 2) {
    console.log('No data available in "Backend" sheet.');
    return;
  }

  var backendHeaders = backendData[0];
  var backendHeaderIndex = {};
  backendHeaders.forEach(function(header, index) {
    backendHeaderIndex[header.trim()] = index;
  });

  // Ensure required headers exist
  var requiredBackendHeaders = ["person ID", "Name"];
  for (var i = 0; i < requiredBackendHeaders.length; i++) {
    if (!(requiredBackendHeaders[i] in backendHeaderIndex)) {
      console.log('Missing required header "' + requiredBackendHeaders[i] + '" in "Backend" sheet.');
      return;
    }
  }

  // Create mapping from person ID to Name
  var personIdToNameMap = {};
  for (var i = 1; i < backendData.length; i++) {
    var row = backendData[i];
    var personId = row[backendHeaderIndex["person ID"]];
    var personName = row[backendHeaderIndex["Name"]];
    if (personId) {
      personIdToNameMap[personId] = personName;
    }
  }

  // Step 5: Merge QMS Data and BF Timesheet Data
  var collatedDataMap = {}; // Keyed by Essay ID

  // Add QMS Data to collatedDataMap
  qmsFilteredData.forEach(function(item) {
    var essayId = item["Essay ID"];
    collatedDataMap[essayId] = item;
  });

  // Add BF Timesheet Data to collatedDataMap
  for (var bfEssayId in bfDataMap) {
    if (collatedDataMap.hasOwnProperty(bfEssayId)) {
      // Update existing entry
      var existingEntry = collatedDataMap[bfEssayId];
      existingEntry["BF Account No"] = bfDataMap[bfEssayId]["BF Account No"];
      existingEntry["BF EST Date"] = bfDataMap[bfEssayId]["BF EST Date"];
      existingEntry["BF EST Time"] = bfDataMap[bfEssayId]["BF EST Time"];
    } else {
      // Add new entry
      collatedDataMap[bfEssayId] = bfDataMap[bfEssayId];
    }
  }

  // Step 6: Perform Calculations and Prepare Data for Output
  var outputData = [];

  // Prepare headers for output sheet
  var outputHeaders = [
    "Essay ID",
    "QMS Account no",
    "QMS EST Date",
    "QMS EST Time",
    "BF Account No",
    "BF EST Date",
    "BF EST Time",
    "Account no Match?",
    "EST Date Match",
    "EST Time difference",
    "Absolute Time difference less than 3?",
    "Person ID",
    "Person Name" // New header
  ];

  // Add headers to outputData
  outputData.push(outputHeaders);

  for (var essayId in collatedDataMap) {
    var item = collatedDataMap[essayId];
    var row = [];

    // Get values or empty strings if undefined
    var qmsAccountNo = item["QMS Account no"] || "";
    var qmsEstDate = item["QMS EST Date"] || "";
    var qmsEstTimeRaw = item["QMS EST Time"];
    var bfAccountNo = item["BF Account No"] || "";
    var bfEstDate = item["BF EST Date"] || "";
    var bfEstTimeRaw = item["BF EST Time"];
    var personId = item["Person ID"] || "";

    // Get Person Name from mapping
    var personName = "";
    if (personId) {
      personName = personIdToNameMap[personId] || "";
    }

    // Parse and format times
    var qmsEstTime = parseTime(qmsEstTimeRaw);
    var bfEstTime = parseTime(bfEstTimeRaw);

    var qmsEstTimeStr = qmsEstTime ? formatTime(qmsEstTime) : "";
    var bfEstTimeStr = bfEstTime ? formatTime(bfEstTime) : "";

    // Account no Match?
    var accountNoMatch = "N";
    if (qmsAccountNo && bfAccountNo && qmsAccountNo === bfAccountNo) {
      accountNoMatch = "Y";
    }

    // EST Date Match
    var estDateMatch = "N";
    if (qmsEstDate && bfEstDate && isSameDate(qmsEstDate, bfEstDate)) {
      estDateMatch = "Y";
    }

    // EST Time difference
    var estTimeDifference = null; // Will be set to a numeric value representing the duration in days
    var absTimeDifferenceLessThan3 = "";
    if (qmsEstTime && bfEstTime) {
      var timeDiffMillis = bfEstTime.getTime() - qmsEstTime.getTime();
      estTimeDifference = timeDiffMillis / (24 * 60 * 60 * 1000); // Convert milliseconds to days
      var absTimeDiffMinutes = Math.abs(timeDiffMillis) / (1000 * 60);
      absTimeDifferenceLessThan3 = (absTimeDiffMinutes < 3) ? "Y" : "N";
    } else {
      estTimeDifference = ""; // Keep it empty if times are not available
    }

    // Build row
    row.push(
      essayId || "",
      qmsAccountNo,
      formatDate(qmsEstDate),
      qmsEstTimeStr,
      bfAccountNo,
      formatDate(bfEstDate),
      bfEstTimeStr,
      accountNoMatch,
      estDateMatch,
      estTimeDifference, // Numeric value representing duration in days
      absTimeDifferenceLessThan3,
      personId, // Person ID
      personName // Person Name
    );

    outputData.push(row);
  }

  // Step 7: Write Data to "RAW Collated data (BF-QMS)" Sheet
  rawCollatedSheet.getRange(1, 1, outputData.length, outputHeaders.length).setValues(outputData);

  // Step 8: Set the number format for "EST Time difference" column to duration format
  var estTimeDifferenceColIndex = outputHeaders.indexOf("EST Time difference") + 1; // +1 because columns are 1-based
  var numRows = outputData.length;
  if (numRows > 1) { // More than just headers
    // Apply duration format to "EST Time difference" column (excluding header row)
    rawCollatedSheet.getRange(2, estTimeDifferenceColIndex, numRows - 1, 1).setNumberFormat('[h]:mm:ss');
  }

  // Optionally, notify the user
  console.log('Data has been successfully collated into "RAW Collated data (BF-QMS)".');
}
