// Constant to hold the Employee Master Sheet ID for accessing Google Sheets
const EMPLOYEE_MASTER_SHEET_ID = '1hDqudXKOjCZ5FV1gv_n1HBcOI9UqnUYruFSDnGBY3sg';
//"1hDqudXKOjCZ5FV1gv_n1HBcOI9UqnUYruFSDnGBY3sg"                 
// "1PKj4kWmuHs9_76_aZUwR0YVk7ifW_3jF5P-XJMOGKGw";

// Set the current date to November 3, 2024, with hours set to 0
const TODAY = new Date(2024, 9, 3); // Note: Month is 0-indexed (0 = January, 9 = October)
TODAY.setHours(0, 0, 0, 0); // Set time to midnight to compare only dates


/**
 * Retrieves data from the Employee Master Sheet and updates the HR Input Sheet.
 */
function getDataFromEmployeeSheet() {
  // Open the Employee Master Sheet and get the 'Headcount' sheet
  const employeeMasterSheet = SpreadsheetApp.openById(EMPLOYEE_MASTER_SHEET_ID).getSheetByName("Headcount");

  // Retrieve headers and data from the Employee Master Sheet
  const [employeeMasterHeaders, employeeMasterData] = CentralLibrary.get_Data_Indices_From_Sheet(employeeMasterSheet);

  // Get the active spreadsheet's 'Input Sheet' and retrieve headers and data
  const hrInputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Input Sheet");
  const [hrInputHeaders, hrInputData] = CentralLibrary.get_Data_Indices_From_Sheet(hrInputSheet, 2);
  

  // Get dropdown data for HR names and Signing Authority names
  const dropdowns = generateNameLists();

  // Collect all email IDs from the HR Input Data to avoid duplicates
  const hrInputDataEmailIds = hrInputData.map(employeeRow => employeeRow[hrInputHeaders["New Emp Id"]]);
  // const alreadyPresentUUIDs = hrInputData.map(employeeRow => employeeRow[hrInputHeaders["Unique ID"]]);

  // Initialize an array to hold new employee data to be written to the HR Input Sheet
  const array = [];

  // Iterate over the Employee Master Data
  for (let i = 0; i < employeeMasterData.length; i++) {
    const employee = employeeMasterData[i];
    console.log(new Date(employee[employeeMasterHeaders["DOJ"]]))
    // Check if the Date of Joining (DOJ) is greater than today and if the employee is not already in HR Input
    if (new Date(employee[employeeMasterHeaders["DOJ"]]).getTime() > TODAY.getTime() &&
      !hrInputDataEmailIds.includes(employee[employeeMasterHeaders["New Emp Id"]])) {
      // Create an inner array to hold the employee's information
      const innerArray = [];
      innerArray.push(
        // "", // Placeholder for potential future data
        employee[employeeMasterHeaders["UUID"]],
        employee[employeeMasterHeaders["New Emp Id"]],
        employee[employeeMasterHeaders["Employee Name"]],
        employee[employeeMasterHeaders["Function"]],
        employee[employeeMasterHeaders["Reporting Manager"]],
        employee[employeeMasterHeaders["Grade"]],
        employee[employeeMasterHeaders["Designation"]],
        employee[employeeMasterHeaders["Department"]],
        employee[employeeMasterHeaders["Gender"]],
        employee[employeeMasterHeaders["DOB"]],
        employee[employeeMasterHeaders["DOJ"]],
        employee[employeeMasterHeaders["Date of Leaving"]],
        employee[employeeMasterHeaders["Tenure"]],
        employee[employeeMasterHeaders["Date of Resignation"]],
        employee[employeeMasterHeaders["Reason"]],
        employee[employeeMasterHeaders["Resignation Category"]],
        employee[employeeMasterHeaders["status at the time of leaving (routine, absconding, termination)"]],
        employee[employeeMasterHeaders["Exit Type"]],
        employee[employeeMasterHeaders["Status"]],
        employee[employeeMasterHeaders["PAN Card No."]],
        employee[employeeMasterHeaders["Name as per PAN"]],
        employee[employeeMasterHeaders["Aadhar Card"]],
        employee[employeeMasterHeaders["Location"]],
        employee[employeeMasterHeaders["Personal email ID"]],
        employee[employeeMasterHeaders["Official email ID"]],
        employee[employeeMasterHeaders["Current Address"]],
        employee[employeeMasterHeaders["Permanent Address"]],
        employee[employeeMasterHeaders["Appraisal Cycle"]],
        employee[employeeMasterHeaders["Days"]],
        employee[employeeMasterHeaders["Hours"]]
      );
      array.push(innerArray); // Add the inner array to the main array
    }
  }

  // If there are new employees to add to the HR Input Sheet
  if (array.length > 0) {
    // Write the new employee data to the HR Input Sheet
    const startRow = hrInputSheet.getLastRow() + 1; // Determine the starting row for writing new data
    const numRows = array.length; // Number of new rows to write
    //Logger.log("Start Row: " + startRow + ", Num Rows: " + numRows);
    const numColumns = array[0].length; // Number of columns in the new data
    hrInputSheet.getRange(startRow, 1, numRows, numColumns).setValues(array); // Set the values in the sheet
   
    // // Set data validation for the last column to be a checkbox
    const checkboxColumn = hrInputHeaders["Appointment & NDA Letter Trigger"] + 1; // Convert to 1-based index
   // Logger.log(checkboxColumn);
    const checkboxRange = hrInputSheet.getRange(startRow, checkboxColumn, numRows, 1); // Define range for checkboxes
    const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build(); // Create checkbox rule
    checkboxRange.setDataValidation(checkboxRule); // Apply checkbox validation


    //Create checkboxes for "Create Relieving Letter?" and "Email Relieving Letter?"

    const createRelievingCol = hrInputHeaders["Create Relieving Letter?"] + 1; // 1-based index
    //Logger.log("Create Relieving Col: " + createRelievingCol);
    const createRelievingRange = hrInputSheet.getRange(startRow, createRelievingCol, numRows, 1);

    createRelievingRange.setDataValidation(checkboxRule);

    // const emailRelievingCol = hrInputHeaders["Create Relieving Letter?"] + 1;   // 1-based index
    // //Logger.log("Email Relieving Col: " + emailRelievingCol);
    // const emailRelievingRange = hrInputSheet.getRange(startRow, emailRelievingCol, numRows, 1);

    // emailRelievingRange.setDataValidation(checkboxRule);



    // Set data validation for the HR Name column
    const hrColumnName = "HR Name"; // Adjust to the exact header name in your sheet
    const hrColumnIndex = hrInputHeaders[hrColumnName] + 1; // Convert to 1-based index
    const hrRange = hrInputSheet.getRange(startRow, hrColumnIndex, numRows, 1); // Define range for HR names
    const hrRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(dropdowns.hrNames, true) // Create rule requiring value in the dropdown list
      .build();
    hrRange.setDataValidation(hrRule); // Apply HR name validation

    // Set data validation for the Signing Authority column
    const saColumnName = "Signing Authority"; // Adjust to the exact header name in your sheet
    const saColumnIndex = hrInputHeaders[saColumnName] + 1; // Convert to 1-based index
    const saRange = hrInputSheet.getRange(startRow, saColumnIndex, numRows, 1); // Define range for Signing Authority
    const saRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(dropdowns.signingAuthorityNames, true) // Create rule requiring value in the dropdown list
      .build();
    saRange.setDataValidation(saRule); // Apply Signing Authority validation
  }
}


/**
 * Retrieves dropdown data for HR and Signing Authority from the 'Drop Downs' sheet.
 * @returns {Object} An object containing arrays of HR names and Signing Authority names.
 */
function getDropDowns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Drop Downs"); // Access the 'Drop Downs' sheet
  const object = {};
  const lastRow = sheet.getLastRow(); // Get the last row to define the range

  // Get HR data starting from row 3, columns B and C (columns 2 and 3)
  const hrDataRange = sheet.getRange(3, 2, lastRow - 2, 2); // Define range for HR data
  const hrData = hrDataRange.getValues(); // Get HR data values
  const hr = hrData
    .filter(row => row[0] != "" && row[1] != "") // Filter out empty rows
    .map(row => ({
      "Name": row[0], // HR Name
      "Designation": row[1] // HR Designation
    }));

  // Get Signing Authority data starting from row 3, columns E and F (columns 5 and 6)
  const saDataRange = sheet.getRange(3, 5, lastRow - 2, 2); // Define range for Signing Authority data
  const saData = saDataRange.getValues(); // Get Signing Authority data values
  const signingAuthority = saData
    .filter(row => row[0] != "" && row[1] != "") // Filter out empty rows
    .map(row => ({
      "Name": row[0], // Signing Authority Name
      "Designation": row[1] // Signing Authority Designation
    }));

  object["HR"] = hr; // Store HR data in the object
  object["SigningAuthority"] = signingAuthority; // Store Signing Authority data in the object
  return object; // Return the object containing both HR and Signing Authority data
}


/**
 * Generates lists of HR names and Signing Authority names for dropdowns.
 * @returns {Object} An object containing arrays of HR names and Signing Authority names.
 */
function generateNameLists() {
  const dropdowns = getDropDowns(); // Get the dropdown data
  const hrList = dropdowns.HR.map(hr => hr.Name); // Extract HR names
  const signingAuthorityList = dropdowns.SigningAuthority.map(sa => sa.Name); // Extract Signing Authority names

  return {
    hrNames: hrList, // Store HR names
    signingAuthorityNames: signingAuthorityList // Store Signing Authority names
  };
}


/**
 * Retrieves the designation of a signing authority based on their name.
 * @param {string} name - The name of the signing authority.
 * @returns {string|null} The designation of the signing authority, or null if not found.
 */
function getSigningAuthority(name) {
  const dropdowns = getDropDowns(); // Get the dropdown data
  const signingAuthorities = dropdowns.SigningAuthority; // Access Signing Authority data

  // Use the Array.prototype.find() method to locate the object with the matching name
  const authority = signingAuthorities.find(authority => authority.Name === name);

  if (authority) {
    return authority.Designation; // Return the designation if found
  } else {
    console.log("Name not found in SigningAuthority"); // Log if name not found
    return null; // Return null if not found
  }
}


/**
 * Retrieves the designation of an HR based on their name.
 * @param {string} name - The name of the HR personnel.
 * @returns {string|null} The designation of the HR, or null if not found.
 */
function getHrDesignation(name) {
  const dropdowns = getDropDowns(); // Get the dropdown data
  const signingAuthorities = dropdowns.HR; // Access HR data

  // Use the Array.prototype.find() method to locate the object with the matching name
  const authority = signingAuthorities.find(authority => authority.Name === name);

  if (authority) {
    return authority.Designation; // Return the designation if found
  } else {
    console.log("Name not found in SigningAuthority"); // Log if name not found
    return null; // Return null if not found
  }
}


/**
 * Determines the financial year based on a given date.
 * @param {Date} date - A Date object for which to determine the financial year.
 * @returns {string} The financial year in 'YY-YY' format.
 * @throws Will throw an error if the input is not a valid Date object.
 */
function getFinancialYear(date) {
  // Validate that the input is a Date object
  if (!(date instanceof Date)) {
    throw new Error("Invalid input. Please provide a valid Date object.");
  }

  const year = date.getFullYear(); // Get the full year from the date
  const month = date.getMonth() + 1; // Get the month from the date (0-indexed, so +1)

  // If the month is April (4) or later, it's the current year - next year
  if (month >= 4) {
    return `${String(year).slice(-2)}-${String(year + 1).slice(-2)}`; // Return format 'YY-YY'
  }
  // If it's before April, it's the previous year - current year
  else {
    return `${String(year - 1).slice(-2)}-${String(year).slice(-2)}`; // Return format 'YY-YY'
  }
}
