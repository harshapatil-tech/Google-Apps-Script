/**
 * ============================================================================
 * BRAINFUSE DATA SYNCHRONIZATION - COMPLETE FIXED VERSION
 * ============================================================================
 * 
 * This script processes raw tutoring data and syncs it to Master_Data sheet.
 * 
 * KEY FIXES:
 * 1. Fixed date matching to check BOTH dates in the SAME row
 * 2. Proper data organization by account type (Single first, then Multiple)
 * 3. Correct total hours calculation and placement
 * 4. Better error handling and logging
 * 5. Removed early "return" statement that prevented execution
 * 
 * AUTHOR: Claude (Anthropic)
 * DATE: February 2026
 */

// ============================================================================
// MAIN PROCESSING FUNCTIONS
// ============================================================================

/**
 * Split a date range into periods (default 14 days), respecting month boundaries
 */
function splitIntoPeriods(startDate, endDate, periodDays = 14) {
  const periods = [];
  let currentStart = new Date(startDate);
  let remainingDaysInPeriod = periodDays;

  while (currentStart <= endDate) {
    const endOfMonth = new Date(currentStart.getFullYear(), currentStart.getMonth() + 1, 0);
    let currentEnd;

    // Calculate days left in current month
    const daysLeftInMonth = Math.ceil((endOfMonth - currentStart) / (1000 * 60 * 60 * 24)) + 1;

    if (daysLeftInMonth >= remainingDaysInPeriod) {
      // Period fits within current month
      currentEnd = new Date(currentStart);
      currentEnd.setDate(currentStart.getDate() + remainingDaysInPeriod - 1);
      periods.push({ from: new Date(currentStart), to: new Date(currentEnd) });

      // Move to next period
      currentStart = new Date(currentEnd);
      currentStart.setDate(currentStart.getDate() + 1);
      remainingDaysInPeriod = periodDays;
    } else {
      // Period spans month boundary
      currentEnd = endOfMonth;
      periods.push({ from: new Date(currentStart), to: new Date(currentEnd) });

      // Carry remaining days to next month
      remainingDaysInPeriod -= daysLeftInMonth;
      currentStart = new Date(currentEnd);
      currentStart.setDate(currentStart.getDate() + 1);
    }
  }

  return periods;
}

/**
 * Process raw tutoring data and organize by account and period
 */
function brainfuseAccountWise(spreadSheet) {
  const ss = spreadSheet.getSheetByName("Summary");
  if (!ss) throw new Error("Summary sheet not found! Please create a 'Summary' sheet with raw transaction data.");

  const data = ss.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  if (rows.length === 0) {
    throw new Error("Summary sheet is empty!");
  }

  const accounts = {}; // Group by subject + account

  // Step 1: Group all rows by subject and account
  rows.forEach(row => {
    let department = row[0]; // Subject column
    const account = row[1];
    const type = row[2] ? row[2].toString().trim().toLowerCase() : "";
    const activityType = row[3] || "";
    const startDate = row[4];
    const hours = parseFloat(row[6]) || 0;

    // Normalize department names
    if (department.toString().trim().toLowerCase() === "intro accounting") {
      department = "Accounting";
    }
    if (department.toString().trim().toLowerCase() === "mathematics") {
      department = "Calculus";
    }

    if (!accounts[department]) accounts[department] = {};
    if (!accounts[department][account]) accounts[department][account] = [];

    accounts[department][account].push({
      date: new Date(startDate),
      occupancy: activityType.toString().includes("Tutored") ? hours : 0,
      nonOccupancy: activityType.toString().includes("Waited") ? hours : 0,
      accountType: (type === "single" ? "Single" : "Multiple"),
      subject: department
    });
  });

  // Step 2: Create global periods based on all dates
  const allDates = rows.map(r => new Date(r[4]));
  const minDate = new Date(Math.min(...allDates));
  const maxDate = new Date(Math.max(...allDates));
  const globalPeriods = splitIntoPeriods(minDate, maxDate, 14);

  Logger.log(`Processing date range: ${minDate.toDateString()} to ${maxDate.toDateString()}`);
  Logger.log(`Total periods: ${globalPeriods.length}`);

  const resultArray = [];

  // Step 3: Process each subject and account
  for (const subject in accounts) {
    for (const account in accounts[subject]) {
      const accountData = accounts[subject][account];

      globalPeriods.forEach(period => {
        const periodOccupancy = accountData
          .filter(d => d.date >= period.from && d.date <= period.to)
          .reduce((sum, d) => sum + d.occupancy, 0);

        const periodNonOccupancy = accountData
          .filter(d => d.date >= period.from && d.date <= period.to)
          .reduce((sum, d) => sum + d.nonOccupancy, 0);

        const monthName = Utilities.formatDate(period.from, Session.getScriptTimeZone(), "MMMM");
        const year = Utilities.formatDate(period.from, Session.getScriptTimeZone(), "yyyy");

        resultArray.push({
          [monthName]: {
            "Occupancy": {
              hours: periodOccupancy,
              subject,
              account,
              firstDay: Utilities.formatDate(period.from, Session.getScriptTimeZone(), "dd-MMM-yyyy"),
              lastDay: Utilities.formatDate(period.to, Session.getScriptTimeZone(), "dd-MMM-yyyy"),
              month: monthName,
              finYear: getFinancialYear(monthName, year),
              accountType: accountData[0].accountType
            },
            "Non-Occupancy": {
              hours: periodNonOccupancy,
              subject,
              account,
              firstDay: Utilities.formatDate(period.from, Session.getScriptTimeZone(), "dd-MMM-yyyy"),
              lastDay: Utilities.formatDate(period.to, Session.getScriptTimeZone(), "dd-MMM-yyyy"),
              month: monthName,
              finYear: getFinancialYear(monthName, year),
              accountType: accountData[0].accountType
            }
          }
        });
      });
    }
  }

  return resultArray;
}

/**
 * MAIN FUNCTION: Sync data to Master_Data sheet
 * 
 * CRITICAL FIX: Removed the early "return" statement that was preventing execution!
 */
function settingBrainFuse_Occupancy_non_Occupanacy() {
  try {
    Logger.log("=== Starting Brainfuse Data Sync ===");
    
    // Get data from Summary sheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let dataToCopy = brainfuseAccountWise(spreadsheet);

    // Access Master_Data sheet
    const MASTER_SHEET_ID = "1WTAwoyzkLtzKOEPqejOVNs7_3xzRMBEHKTumlNwMd-M";
    const ss = SpreadsheetApp.openById(MASTER_SHEET_ID).getSheetByName("Master_Data");
    
    if (!ss) throw new Error("Master_Data sheet not found!");

    // Get existing data
    const dataRange = ss.getRange(1, 1, ss.getLastRow(), ss.getLastColumn()).getDisplayValues();
    let headers = dataRange[0];
    let data = dataRange.slice(1);

    // Map column indices
    const colMap = {
      srNo: headers.indexOf("Invoice No"),
      finYear: headers.indexOf("Financial Year"),
      semester: headers.indexOf("Semester"),
      month: headers.indexOf("Month"),
      from: headers.indexOf("From"),
      to: headers.indexOf("To"),
      subject: headers.indexOf("Subject"),
      accountType: headers.indexOf("Single/Multiple"),
      account: headers.indexOf("Account No"),
      occupancyType: headers.indexOf("Occupancy"),
      hours: headers.indexOf("Hours"),
      totalHours: headers.indexOf("Total Hours")
    };

    // Validate all required columns exist
    for (const [key, index] of Object.entries(colMap)) {
      if (index === -1) {
        throw new Error(`Column "${key}" not found in Master_Data sheet!`);
      }
    }

    // Sort and reorganize data
    dataToCopy = sortDataByEarliestMonth(dataToCopy);
    dataToCopy = reorganizeData(dataToCopy);

    Logger.log(`Total months to process: ${Object.keys(dataToCopy).length}`);

    // ========================================================================
    // CRITICAL FIX #1: Proper date matching using findIndex
    // ========================================================================
    
    let deletionStartRow = null;
    let lastSerialNum = ss.getLastRow() > 1 ? ss.getRange(ss.getLastRow(), colMap.srNo + 1).getValue() : 0;
    
    // Build a lookup of existing date ranges
    const existingRanges = data.map((row, idx) => ({
      from: row[colMap.from],
      to: row[colMap.to],
      rowNum: idx + 2 // +2 because: +1 for header, +1 for 1-based indexing
    }));

    // Find the first matching date range
    outerLoop:
    for (const [monthKey, monthData] of Object.entries(dataToCopy)) {
      for (const accountType of ["Single", "Multiple"]) {
        if (!monthData[accountType]) continue;

        for (const entry of monthData[accountType]) {
          const firstDayToCheck = entry.Occupancy?.firstDay || entry["Non-Occupancy"]?.firstDay;
          const lastDayToCheck = entry.Occupancy?.lastDay || entry["Non-Occupancy"]?.lastDay;
          
          if (!firstDayToCheck || !lastDayToCheck) continue;

          // Find exact match where BOTH dates match in SAME row
          const matchIdx = existingRanges.findIndex(
            range => range.from === firstDayToCheck && range.to === lastDayToCheck
          );

          if (matchIdx !== -1) {
            deletionStartRow = existingRanges[matchIdx].rowNum;
            Logger.log(`✓ Found matching date range: ${firstDayToCheck} to ${lastDayToCheck}`);
            Logger.log(`✓ Will delete from row ${deletionStartRow} onwards`);
            break outerLoop;
          }
        }
      }
    }

    // Delete old data if match found
    if (deletionStartRow !== null) {
      const rowsToDelete = ss.getLastRow() - deletionStartRow + 1;
      Logger.log(`Deleting ${rowsToDelete} rows (rows ${deletionStartRow} to ${ss.getLastRow()})`);
      ss.deleteRows(deletionStartRow, rowsToDelete);
      Logger.log("✓ Old data deleted successfully");
    } else {
      Logger.log("No matching date ranges found - will append new data");
      deletionStartRow = ss.getLastRow() + 1;
    }

    // ========================================================================
    // CRITICAL FIX #2: Proper iteration order and group tracking
    // ========================================================================
    
    let currentRow = deletionStartRow;
    let currentGroup = { subject: "", accountType: "", month: "" };
    let groupTotalHours = 0;
    let totalRowsInserted = 0;

    for (const [monthKey, monthData] of Object.entries(dataToCopy)) {
      Logger.log(`\nProcessing month: ${monthKey}`);
      
      // Process Single accounts first, then Multiple
      for (const accountType of ["Single", "Multiple"]) {
        if (!monthData[accountType]) continue;

        Logger.log(`  Processing ${accountType} accounts...`);

        for (const entry of monthData[accountType]) {
          // Process both Occupancy and Non-Occupancy
          for (const [occKey, val] of Object.entries(entry)) {
            if (!val) continue;

            const subject = val.subject;
            const accType = val.accountType;

            // ================================================================
            // CRITICAL FIX #3: Proper total hours calculation and placement
            // ================================================================
            
            // Check if we're starting a new group
            if (currentGroup.subject !== subject || currentGroup.accountType !== accType) {
              // Write total for PREVIOUS group
              if (currentGroup.subject !== "" && groupTotalHours > 0) {
                const totalRow = currentRow - 1; // Previous row!
                applyCustomFormatting(
                  ss.getRange(totalRow, colMap.totalHours + 1),
                  { bgColor: "#fbbc04", fontWeight: "bold" }
                ).setValue(groupTotalHours);
                Logger.log(`    ✓ Total for ${currentGroup.subject} (${currentGroup.accountType}): ${groupTotalHours} hours`);
              }

              // Reset for new group
              currentGroup = { subject, accountType: accType, month: monthKey };
              groupTotalHours = 0;
            }

            // Add to group total
            groupTotalHours += val.hours;

            // ================================================================
            // Insert row data
            // ================================================================
            
            lastSerialNum++;
            
            // Invoice No
            applyCustomFormatting(ss.getRange(currentRow, colMap.srNo + 1))
              .setValue(lastSerialNum);
            
            // Financial Year
            applyCustomFormatting(ss.getRange(currentRow, colMap.finYear + 1))
              .setValue(val.finYear);
            
            // Semester
            applyCustomFormatting(ss.getRange(currentRow, colMap.semester + 1))
              .setValue(getSeason(val.month));
            
            // Month
            applyCustomFormatting(ss.getRange(currentRow, colMap.month + 1))
              .setValue(val.month);
            
            // From Date
            applyCustomFormatting(dateValidation(ss.getRange(currentRow, colMap.from + 1)))
              .setValue(val.firstDay);
            
            // To Date
            applyCustomFormatting(dateValidation(ss.getRange(currentRow, colMap.to + 1)))
              .setValue(val.lastDay);
            
            // Subject
            applyCustomFormatting(ss.getRange(currentRow, colMap.subject + 1))
              .setValue(val.subject);
            
            // Account Type (Single/Multiple) with color coding
            const accountTypeCell = ss.getRange(currentRow, colMap.accountType + 1);
            if (val.accountType === "Single") {
              applyCustomFormatting(accountTypeCell, { bgColor: "#93c47d" }).setValue("Single");
            } else {
              applyCustomFormatting(accountTypeCell, { bgColor: "#e380e3" }).setValue("Multiple");
            }
            
            // Account Number (extract numeric part only)
            applyCustomFormatting(
              ss.getRange(currentRow, colMap.account + 1),
              { fontWeight: "bold" }
            ).setValue(extractNumberFromString(val.account.toString()));
            
            // Hours
            applyCustomFormatting(ss.getRange(currentRow, colMap.hours + 1))
              .setValue(val.hours);
            
            // Occupancy Type
            applyCustomFormatting(ss.getRange(currentRow, colMap.occupancyType + 1))
              .setValue(occKey);

            currentRow++;
            totalRowsInserted++;
          }
        }
      }
    }

    // ========================================================================
    // CRITICAL FIX #4: Don't forget the last group's total!
    // ========================================================================
    
    if (currentGroup.subject !== "" && groupTotalHours > 0) {
      const totalRow = currentRow - 1;
      applyCustomFormatting(
        ss.getRange(totalRow, colMap.totalHours + 1),
        { bgColor: "#fbbc04", fontWeight: "bold" }
      ).setValue(groupTotalHours);
      Logger.log(`    ✓ Total for ${currentGroup.subject} (${currentGroup.accountType}): ${groupTotalHours} hours`);
    }

    Logger.log(`\n=== Sync Complete ===`);
    Logger.log(`Total rows inserted: ${totalRowsInserted}`);
    Logger.log(`Final row number: ${currentRow - 1}`);
    
    SpreadsheetApp.getUi().alert(
      'Success!',
      `Data synchronization completed successfully!\n\n` +
      `• Rows inserted: ${totalRowsInserted}\n` +
      `• Processing completed at row ${currentRow - 1}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {
    Logger.log("ERROR: " + error.toString());
    Logger.log(error.stack);
    SpreadsheetApp.getUi().alert(
      'Error',
      'An error occurred during synchronization:\n\n' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw error;
  }
}

// ============================================================================
// DATA ORGANIZATION FUNCTIONS
// ============================================================================

/**
 * Reorganize data by account type (Single/Multiple)
 */
function reorganizeData(dataToCopy) {
  const modifiedData = {};

  for (const [monthKey, values] of Object.entries(dataToCopy)) {
    const monthData = { Single: [], Multiple: [] };

    values.forEach(entry => {
      ['Single', 'Multiple'].forEach(accountType => {
        const occupancy = entry.Occupancy && entry.Occupancy.accountType === accountType 
          ? entry.Occupancy : null;
        const nonOccupancy = entry['Non-Occupancy'] && entry['Non-Occupancy'].accountType === accountType 
          ? entry['Non-Occupancy'] : null;

        if (occupancy || nonOccupancy) {
          const dataObject = {};
          if (occupancy) dataObject.Occupancy = occupancy;
          if (nonOccupancy) dataObject['Non-Occupancy'] = nonOccupancy;
          monthData[accountType].push(dataObject);
        }
      });
    });

    // Remove empty categories
    if (monthData.Single.length === 0) delete monthData.Single;
    if (monthData.Multiple.length === 0) delete monthData.Multiple;

    // Only add month if it has data
    if (Object.keys(monthData).length > 0) {
      modifiedData[monthKey] = monthData;
    }
  }

  return modifiedData;
}

/**
 * Sort data by calendar month order
 */
function sortDataByEarliestMonth(data) {
  const monthOrder = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];
  
  let sortedData = {};

  // Initialize with empty arrays
  monthOrder.forEach(month => {
    sortedData[month] = [];
  });

  // Group by month
  data.forEach(item => {
    monthOrder.forEach(month => {
      if (item[month]) {
        sortedData[month].push(item[month]);
      }
    });
  });

  // Filter out empty months
  let finalSortedData = {};
  Object.keys(sortedData).forEach(month => {
    if (sortedData[month].length > 0) {
      finalSortedData[month] = sortedData[month];
    }
  });

  return finalSortedData;
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Get financial year (assumes fiscal year starts in July)
 */
function getFinancialYear(monthName, year) {
  const monthIndex = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ].indexOf(monthName);
  
  // Fiscal year starts in July (month index 6)
  if (monthIndex >= 6) {
    return `${year}-${String(parseInt(year) + 1).slice(-2)}`;
  } else {
    return `${parseInt(year) - 1}-${String(year).slice(-2)}`;
  }
}

/**
 * Get season/semester based on month
 */
function getSeason(monthName) {
  const seasonMap = {
    "December": "Winter",
    "January": "Spring",
    "February": "Spring",
    "March": "Spring",
    "April": "Spring",
    "May": "Spring",
    "June": "Summer",
    "July": "Summer",
    "August": "Summer",
    "September": "Fall",
    "October": "Fall",
    "November": "Fall"
  };
  return seasonMap[monthName] || "Unknown";
}

/**
 * Extract numbers from a string (e.g., "Account 12345" → "12345")
 */
function extractNumberFromString(str) {
  const match = str.toString().match(/\d+/);
  return match ? match[0] : str;
}

/**
 * Apply custom formatting to a cell range
 */
function applyCustomFormatting(range, options = {}) {
  // Apply optional formatting
  if (options.bgColor) range.setBackground(options.bgColor);
  if (options.fontWeight) range.setFontWeight(options.fontWeight);
  if (options.fontSize) range.setFontSize(options.fontSize);
  if (options.fontColor) range.setFontColor(options.fontColor);
  
  // Apply default formatting
  range.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  range.setHorizontalAlignment("center");
  range.setVerticalAlignment("middle");
  
  return range;
}

/**
 * Apply date validation to a range
 */
function dateValidation(range) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
  return range;
}




// function splitIntoPeriods(startDate, endDate, periodDays = 14) {
//   const periods = [];
//   let currentStart = new Date(startDate); // clone to avoid mutation
//   let remainingDaysInPeriod = periodDays;

//   while (currentStart <= endDate) {
//     const endOfMonth = new Date(currentStart.getFullYear(), currentStart.getMonth() + 1, 0);
//     let currentEnd;

//     // Days left in this month
//     const daysLeftInMonth = Math.ceil((endOfMonth - currentStart) / (1000 * 60 * 60 * 24)) + 1;

//     if (daysLeftInMonth >= remainingDaysInPeriod) {
//       // Period fits within current month
//       currentEnd = new Date(currentStart);
//       currentEnd.setDate(currentStart.getDate() + remainingDaysInPeriod - 1);
//       periods.push({ from: new Date(currentStart), to: new Date(currentEnd) });

//       // Prepare for next period
//       currentStart = new Date(currentEnd);
//       currentStart.setDate(currentStart.getDate() + 1);
//       remainingDaysInPeriod = periodDays;
//     } else {
//       // Period spans month boundary
//       currentEnd = endOfMonth;
//       periods.push({ from: new Date(currentStart), to: new Date(currentEnd) });

//       // Remaining days carry to next month
//       remainingDaysInPeriod -= daysLeftInMonth;

//       currentStart = new Date(currentEnd);
//       currentStart.setDate(currentStart.getDate() + 1);
//     }
//   }

//   return periods;
// }


// function brainfuseAccountWise(spreadSheet) {
//   const ss = spreadSheet.getSheetByName("Summary");
//   if (!ss) throw new Error("Summary sheet not found");

//   const data = ss.getDataRange().getValues();
//   const headers = data[0];
//   const rows = data.slice(1);

//   const accounts = {}; // group by subject + account

//   // Step 1: Group all rows by subject (department) and account
//   rows.forEach(row => {
//     let department = row[0]; // Subject
//     const account = row[1];
//     const type = row[2].trim().toLowerCase();
//     const activityType = row[3];
//     const startDate = row[4];
//     const hours = row[6] || 0;

//     if (department.trim().toLowerCase() === "intro accounting") department = "Accounting";
//     if (department.trim().toLowerCase() === "mathematics") department = "Calculus";

//     if (!accounts[department]) accounts[department] = {};
//     if (!accounts[department][account]) accounts[department][account] = [];

//     accounts[department][account].push({
//       date: new Date(startDate),
//       occupancy: activityType.includes("Tutored") ? hours : 0,
//       nonOccupancy: activityType.includes("Waited") ? hours : 0,
//       accountType: (type === "single" ? "Single" : "Multiple"),
//       subject: department
//     });
//   });

//   // Step 2: Create global periods based on all dates in sheet
//   const allDates = rows.map(r => new Date(r[4]));
//   const minDate = new Date(Math.min(...allDates));
//   const maxDate = new Date(Math.max(...allDates));
//   const globalPeriods = splitIntoPeriods(minDate, maxDate, 14); // 14-day period logic

//   const resultArray = [];

//   // Step 3: Iterate subjects and accounts
//   for (const subject in accounts) {
//     for (const account in accounts[subject]) {
//       const accountData = accounts[subject][account];

//       globalPeriods.forEach(period => {
//         const periodOccupancy = accountData
//           .filter(d => d.date >= period.from && d.date <= period.to)
//           .reduce((sum, d) => sum + d.occupancy, 0);

//         const periodNonOccupancy = accountData
//           .filter(d => d.date >= period.from && d.date <= period.to)
//           .reduce((sum, d) => sum + d.nonOccupancy, 0);

//         // Include even if zero hours (optional: remove if you want only non-zero)
//         const monthName = Utilities.formatDate(period.from, Session.getScriptTimeZone(), "MMMM");
//         const year = Utilities.formatDate(period.from, Session.getScriptTimeZone(), "yyyy");

//         resultArray.push({
//           [monthName]: {
//             "Occupancy": {
//               hours: periodOccupancy,
//               subject,
//               account,
//               firstDay: Utilities.formatDate(period.from, Session.getScriptTimeZone(), "dd-MMM-yyyy"),
//               lastDay: Utilities.formatDate(period.to, Session.getScriptTimeZone(), "dd-MMM-yyyy"),
//               month: monthName,
//               finYear: getFinancialYear(monthName, year),
//               accountType: accountData[0].accountType
//             },
//             "Non-Occupancy": {
//               hours: periodNonOccupancy,
//               subject,
//               account,
//               firstDay: Utilities.formatDate(period.from, Session.getScriptTimeZone(), "dd-MMM-yyyy"),
//               lastDay: Utilities.formatDate(period.to, Session.getScriptTimeZone(), "dd-MMM-yyyy"),
//               month: monthName,
//               finYear: getFinancialYear(monthName, year),
//               accountType: accountData[0].accountType
//             }
//           }
//         });
//       });
//     }
//   }
//   return resultArray;
// }


// // function settingBrainFuse_Occupancy_non_Occupanacy() {
// //   // Retrieve data from brainfuse using the provided id
// //   const spreadsheet = SpreadsheetApp.openById("1XH5eGI4thuz-yAsHDeVbI90KFX2AWee3w556yUcGe6I");
// //   let dataToCopy = brainfuseAccountWise(spreadsheet, "Single").flat();

// //   // Access the "MasterData" sheet
// //   const ss = SpreadsheetApp.openById("1WTAwoyzkLtzKOEPqejOVNs7_3xzRMBEHKTumlNwMd-M").getSheetByName("Master_Data");
  
// //   const dataRange = ss.getRange(1, 1, ss.getLastRow(), ss.getLastColumn()).getDisplayValues();
// //   let headers = dataRange[0], data = dataRange.slice(1);

// //   // Column indices
// //   const srNoIdx = headers.indexOf("Invoice No");
// //   const finYearIdx = headers.indexOf("Financial Year");
// //   const semesterIdx = headers.indexOf("Semester");
// //   const monthIdx = headers.indexOf("Month");
// //   const fromIdx = headers.indexOf("From");
// //   const toIdx = headers.indexOf("To");
// //   const subIdx = headers.indexOf("Subject");
// //   const accountTypeIdx = headers.indexOf("Single/Multiple");
// //   const accountIdx = headers.indexOf("Account No");
// //   const occupancyTypeIdx = headers.indexOf("Occupancy");
// //   const numHrsIdx = headers.indexOf("Hours");
// //   const totalHrsIdx = headers.indexOf("Total Hours");

// //   const firstDates = data.map(r => r[fromIdx]);
// //   const lastDates = data.map(r => r[toIdx]);
  
// //   dataToCopy = sortDataByEarliestMonth(dataToCopy);

// //   console.log(dataToCopy)

// //   return;

// //   let lastRowIndex = ss.getLastRow();
// //   let lastSerialNum = ss.getRange(lastRowIndex, srNoIdx+1).getValue();
  
// //   let isDeletionIndexFound = false;
// //   let rowDeletionStart = ss.getLastRow();

// //   let currentGroup = { subject: "", accountType: "" };
// //   let groupTotalHours = 0;

// //   for (const [monthKey, monthValuesObject] of Object.entries(dataToCopy)) {
// //     for (const object of monthValuesObject) {
// //       const subject = object["Occupancy"].subject;
// //       const accountType = object["Occupancy"].accountType;

// //       // New group check when subject ya account type change then put the total of old  group
// //       if (currentGroup.subject !== subject || currentGroup.accountType !== accountType) {
// //         // Set Total Hours for previous group
// //         if (currentGroup.subject !== "") {
// //           applyCustomFormatting(ss.getRange(lastRowIndex, totalHrsIdx+1), {"bgColor":"#fbbc04","fontWeight":"bold"})
// //             .setValue(groupTotalHours);
// //         }
// //         // Reset for new group
// //         currentGroup.subject = subject;
// //         currentGroup.accountType = accountType;
// //         groupTotalHours = 0;
// //       }

// //       for (const [key, val] of Object.entries(object)) {
// //         if (!isDeletionIndexFound && firstDates.includes(val["firstDay"]) && lastDates.includes(val["lastDay"])) {
// //           rowDeletionStart = firstDates.indexOf(val["firstDay"]) + 2;
// //           const rowsToDelete = ss.getLastRow() - rowDeletionStart + 1;
// //           console.log(`Deleting rows from ${val["firstDay"]} to ${val["lastDay"]} (rows ${rowDeletionStart} to ${ss.getLastRow()})`);
// //           ss.deleteRows(rowDeletionStart, rowsToDelete);
// //           isDeletionIndexFound = true;
// //           lastRowIndex = rowDeletionStart - 1;
// //         }

// //         const hours = val["hours"];
// //         groupTotalHours += hours;
// //         console.log(`Inserting row for Account:- From: ${val["firstDay"]}, To: ${val["lastDay"]}`);
// //         lastRowIndex += 1;
// //         applyCustomFormatting(ss.getRange(lastRowIndex, srNoIdx+1)).setValue(lastSerialNum + 1);
// //         applyCustomFormatting(ss.getRange(lastRowIndex, finYearIdx+1)).setValue(val["finYear"]);
// //         applyCustomFormatting(ss.getRange(lastRowIndex, semesterIdx+1)).setValue(getSeason(val["month"]));
// //         applyCustomFormatting(ss.getRange(lastRowIndex, monthIdx+1)).setValue(val["month"]);
// //         applyCustomFormatting(dateValidation(ss.getRange(lastRowIndex, fromIdx+1))).setValue(val["firstDay"]);
// //         applyCustomFormatting(dateValidation(ss.getRange(lastRowIndex, toIdx+1))).setValue(val["lastDay"]);
// //         applyCustomFormatting(ss.getRange(lastRowIndex, subIdx+1)).setValue(val["subject"]);

// //         if (val["accountType"] == "Single")
// //           applyCustomFormatting(ss.getRange(lastRowIndex, accountTypeIdx+1), {"bgColor": "#93c47d"}).setValue("Single");
// //         else
// //           applyCustomFormatting(ss.getRange(lastRowIndex, accountTypeIdx+1), {"bgColor":"#e380e3"}).setValue("Multiple");

// //         applyCustomFormatting(ss.getRange(lastRowIndex, accountIdx+1), {"fontWeight":"bold"})
// //                               .setValue(extractNumberFromString(val["account"]));
// //         applyCustomFormatting(ss.getRange(lastRowIndex, numHrsIdx+1)).setValue(hours);

// //         if (key === "Occupancy")
// //           applyCustomFormatting(ss.getRange(lastRowIndex, occupancyTypeIdx+1)).setValue("Occupancy");
// //         else if (key === "Non-Occupancy")
// //           applyCustomFormatting(ss.getRange(lastRowIndex, occupancyTypeIdx+1)).setValue("Non-Occupancy");
// //       }
// //     }
// //   }

// //   // Set Total Hours for last group
// //   if (currentGroup.subject !== "") {
// //     applyCustomFormatting(ss.getRange(lastRowIndex, totalHrsIdx+1), {"bgColor":"#fbbc04","fontWeight":"bold"})
// //       .setValue(groupTotalHours);
// //   }
// // }

 
// // function reorganizeData(dataToCopy) {
// //   const modifiedData = {};

// //   for (const [monthKey, values] of Object.entries(dataToCopy)) {
// //     // Initialize the month data with empty arrays for Single and Multiple account types
// //     const monthData = { Single: [], Multiple: [] };

// //     values.forEach(entry => {
// //       // Iterate through each entry and organize by accountType
// //       ['Single', 'Multiple'].forEach(accountType => {
// //         const occupancy = entry.Occupancy && entry.Occupancy.accountType === accountType ? entry.Occupancy : null;
// //         const nonOccupancy = entry['Non-Occupancy'] && entry['Non-Occupancy'].accountType === accountType ? entry['Non-Occupancy'] : null;

// //         // If either occupancy or non-occupancy data exists for the accountType, push a new object
// //         if (occupancy || nonOccupancy) {
// //           const dataObject = {};
// //           if (occupancy) dataObject.Occupancy = occupancy;
// //           if (nonOccupancy) dataObject['Non-Occupancy'] = nonOccupancy;
          
// //           monthData[accountType].push(dataObject);
// //         }
// //       });
// //     });

// //     // Remove empty categories (Single or Multiple) if they have no data
// //     if (monthData.Single.length === 0) delete monthData.Single;
// //     if (monthData.Multiple.length === 0) delete monthData.Multiple;

// //     // Only add the month data to the modifiedData if there's meaningful content
// //     if (Object.keys(monthData).length > 0) {
// //       modifiedData[monthKey] = monthData;
// //     }
// //   }

// //   return modifiedData;
// // }


// // function sortDataByEarliestMonth(data) {
// //   const monthOrder = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
// //   let sortedData = {};

// //   // Initialize the sortedData object with months to ensure the order
// //   monthOrder.forEach(month => {
// //     sortedData[month] = [];
// //   });

// //   // Iterate over the data to group them by months
// //   data.forEach(item => {
// //     monthOrder.forEach(month => {
// //       if (item[month]) {
// //         sortedData[month].push(item[month]);
// //       }
// //     });
// //   });

// //   // Filter out months that have no data
// //   let finalSortedData = {};
// //   Object.keys(sortedData).forEach(month => {
// //     if (sortedData[month].length > 0) {
// //       finalSortedData[month] = sortedData[month];
// //     }
// //   });

// //   return finalSortedData;
// // }


