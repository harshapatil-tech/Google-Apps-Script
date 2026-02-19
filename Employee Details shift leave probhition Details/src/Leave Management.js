function runleaveManagement() {
  const x = new LeaveManagement();
  x.fillUpEmployeeIndicatorForEverySubject();
}


class LeaveManagement {
  constructor() {
    this.employeeDetails = new EmployeeDetails();

    // Get the array [headers, data] for active employees
    this.activeEmployeeList = this.employeeDetails.getAllActiveEmployees();

    // Deconstruct if needed or just store as-is
    const [employeeSheetHeaders, employeeData] = this.activeEmployeeList;
    this.activeEmployeeList = [employeeSheetHeaders, employeeData];

    // Map each subject to its spreadsheet ID
    this.sheetLinks = {
      "Biology":           "1MlBc2V4VZgj3AfJQGdAberbqZAK_gbgwJ37WuRNv7w8",
      "Business":          "1Scfv9aGqJrWCYaaZmKlgBVZIUES6r8e8FK9W1KMk3PY",
      "Chemistry":         "1s6T_ua5gns4sC7h_5ljT8lcONRLKZX62w4LeQTgFeJ4",
      "Computer Science":  "1g-dfSEG6VbzMDe2gG-l_ir5gjLyNaME11iWlmReMleI",
      "English": "1wcNARFXwACIwsFQrf6qfhuwbGoZ_iVU00ps4EfpIGTk",
      "Mathematics":       "1kSeGBDP12Yscjbk1ZjAkq0RlIt5MXCOhGUJ3bNLw5K4",
      "Physics":           "1I1CvwEAEj1zfipvqxSoR1TgMMWgsYdhrZDaST7GhDHA",
      "Statistics": "1_aELI5paQ4OxVo9CuUSGAIn-A9UByRXjP3y8TEDYNaU"
    };
  }

  /**
   * Loops over every subject and fills up indicator data
   */
  fillUpEmployeeIndicatorForEverySubject() {
    Object.keys(this.sheetLinks).forEach(subject => {
      console.log("Subject is", subject);
      this.fillUpEmployeeIndicatorForOneSubject(subject)
    });

  }

  /**
   * Fills employee data for the given subject into the subject's spreadsheet
   */
  fillUpEmployeeIndicatorForOneSubject(subject) {
    // Use the EmployeeDetails method to get employees for a subject
    const employeeList = this.employeeDetails.getEmployeesForASubject(
      this.activeEmployeeList,
      subject
    );
    console.log("Employee list for subject", subject, employeeList);

    const nameList = []
    // Open the subject-specific spreadsheet
    const spreadsheet = SpreadsheetApp.openById(this.sheetLinks[subject]);
    const sheetName = "Employee List";
    let employeeListSheet = spreadsheet.getSheetByName(sheetName);
    const databaseSheet = spreadsheet.getSheetByName("Leave_data");
    //const databaseSheet = spreadsheet.getSheetByName("Copy Leave_data");

    const alreadyPresentUUIDs = employeeListSheet.getRange(2, 1, employeeListSheet.getLastRow() + 1).getValues().flat().filter(Boolean);

    employeeList.forEach(row => {
      if (!alreadyPresentUUIDs.includes(row[0])) {
        nameList.push(row);
      }
    })

    console.log("Name list is:-", nameList)

    //if (nameList.length <= 0) return;

    if (nameList.length > 0) {

      let nameRange;
      // Create the Employee List sheet if it doesn't exist
      if (!employeeListSheet) {
        employeeListSheet = spreadsheet.insertSheet();
        employeeListSheet.setName(sheetName);
        // Insert header row at the top
        nameList.unshift(["UniqueID", "Employee Identifier", "Names"]);
        nameRange = employeeListSheet.getRange(1, 1, nameList.length, 3);
      } else {
        nameRange = employeeListSheet.getRange(employeeListSheet.getLastRow() + 1, 1, nameList.length, 3);
      }
      nameRange.setValues(nameList);
    }

    // this.updateNamesWithEmployeeIdentifier(employeeListSheet, databaseSheet);
    const lastRow = employeeListSheet.getLastRow();
    const smeList = employeeListSheet.getRange(2, 2, lastRow, 1).getValues().flat().filter(Boolean);

    this.updateNames(employeeListSheet, databaseSheet);
    this.updateUUID(employeeListSheet, databaseSheet);
    this.updateNamesWithEmployeeIdentifier(employeeListSheet, databaseSheet);

    smeList.unshift("HR");
    this.enableDropdowns(databaseSheet, ["Name of SME", "Name of The Reporting Person", "Approver"], smeList);
  }


  updateNamesWithEmployeeIdentifier(employeeListSheet, leaveDataSheet) {
    const [leaveHeaders, leaveData] = CentralLibrary.get_Data_Indices_From_Sheet(leaveDataSheet);
    const [empHeaders, empData] = CentralLibrary.get_Data_Indices_From_Sheet(employeeListSheet);

    const targetColumns = ["Name of SME", "Name of The Reporting Person", "Approver"];

    // Clean name: remove nested brackets and any existing (E###)
    const cleanNameOnly = (str) => {
      if (!str) return "";
      let s = str.toString().trim();

      // 1) If there is nested form like "(Name (E30))", remove the whole bracketed nested part first
      s = s.replace(/\([^()]*\([^()]*E\d+[^()]*\)[^()]*\)/gi, "");

      // 2) Remove any remaining (E###)
      s = s.replace(/\(E\d+\)/gi, "");

      // 3) Remove extra parentheses that may remain like "(Name)" and extra spaces
      s = s.replace(/^\(|\)$/g, "").trim();
      s = s.replace(/\u00A0/g, " ");
      s = s.replace(/\s+/g, " ");

      return s;
    };

    // Extract code like E30 from identifier cell (handles "E30", "(E30)", "Name (E30)", "Name (Name (E30))")
    const extractEcode = (identifierStr) => {
      if (!identifierStr) return "";
      const s = identifierStr.toString();
      const m = s.match(/E\d+/i);
      if (m) return m[0].toUpperCase();
      // fallback: try digits only
      const digits = s.match(/\d+/);
      return digits ? `E${digits[0]}` : s.trim();
    };

    // Build map: CLEAN NAME -> CODE (E###)
    const nameToCode = {};
    empData.forEach(row => {
      const rawNameCell = row[empHeaders["Names"]] ?? "";
      const rawIdCell = row[empHeaders["Employee Identifier"]] ?? "";

      const nameOnly = cleanNameOnly(rawNameCell);
      const code = extractEcode(rawIdCell);

      if (nameOnly && code) {
        nameToCode[nameOnly] = code; 
      }
    });

    console.log("NAME→CODE MAP:", nameToCode);

    // Now replace values in leaveData with "Name (E###)"
    leaveData.forEach((row, rIdx) => {
      targetColumns.forEach(colName => {
        const colIndex = leaveHeaders[colName];
        if (colIndex === undefined) return;

        const cellVal = row[colIndex];
        if (!cellVal) return;

        const nameOnly = cleanNameOnly(cellVal);
        const code = nameToCode[nameOnly];
        if (!code) return; // no mapping found

        const finalValue = `${nameOnly} (${code})`;

        if (finalValue !== cellVal) {
          leaveDataSheet.getRange(rIdx + 2, colIndex + 1).setValue(finalValue);
          console.log(`Updated → ${finalValue}`);
        }
      });
    });
  }

  updateUUID(backendSheet, leaveDataSheet) {
    const [backendHeaders, backendData] = CentralLibrary.get_Data_Indices_From_Sheet(backendSheet);
    const obj = Object.fromEntries(
      backendData.map(row => [row[2], row[0]])
    );

    const [leaveHeaders, leaveData] = CentralLibrary.get_Data_Indices_From_Sheet(leaveDataSheet);
    leaveData.forEach((row, idx) => {
      const smeName = row[leaveHeaders["Name of SME"]];
      const reportingName = row[leaveHeaders["Name of The Reporting Person"]];
      const approverName = row[leaveHeaders["Approver"]]
      if (obj[smeName]) {
        leaveDataSheet.getRange(idx + 2, leaveHeaders["UUID SME"] + 1).setValue(obj[smeName]);
      }
      if (obj[reportingName]) {
        leaveDataSheet.getRange(idx + 2, leaveHeaders["UUID Reporting Person"] + 1).setValue(obj[reportingName]);
      }
      if (obj[approverName]) {
        leaveDataSheet.getRange(idx + 2, leaveHeaders["UUID Approver"] + 1).setValue(obj[approverName]);
      }
    })
  }


  updateNames(backendSheet, leaveDataSheet) {
    const [backendHeaders, backendData] = CentralLibrary.get_Data_Indices_From_Sheet(backendSheet);
    const obj = Object.fromEntries(
      backendData.map(row => [row[0], row[1]])
    );

    const [leaveHeaders, leaveData] = CentralLibrary.get_Data_Indices_From_Sheet(leaveDataSheet);
    leaveData.forEach((row, idx) => {
      const smeUUID = row[leaveHeaders["UUID SME"]];
      const reportingUUID = row[leaveHeaders["UUID Reporting Person"]];
      const approverUUID = row[leaveHeaders["UUID Approver"]]

      if (obj[smeUUID]) {
        leaveDataSheet.getRange(idx + 2, leaveHeaders["Name of SME"] + 1).setValue(obj[smeUUID]);
      }
      if (obj[reportingUUID]) {
        leaveDataSheet.getRange(idx + 2, leaveHeaders["Name of The Reporting Person"] + 1).setValue(obj[reportingUUID]);
      }
      if (obj[approverUUID]) {
        leaveDataSheet.getRange(idx + 2, leaveHeaders["Approver"] + 1).setValue(obj[approverUUID]);
      }
    })
  }


  enableDropdowns(sheet, smeHeadersColumnArray, smeList) {
    // Employee Identifier dropdowns
    const lastRow = sheet.getLastRow();

    const expandedList = [];

    smeList.forEach(entry => {
      if (!entry) return;

      expandedList.push(entry);

      // If entry contains "(E...)" — extract clean name also
      const match = entry.match(/^(.+?)\s*\((E\d+)\)$/i);
      if (match) {
        expandedList.push(match[1].trim()); // clean version
      }
    });

    // Remove duplicates
    const finalList = [...new Set(expandedList)];


    const smeNamesIndices = CentralLibrary.createIndexMap(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues().flat().filter(Boolean));
    for (const [headerName, headerIndex] of Object.entries(smeNamesIndices)) {
      if (smeHeadersColumnArray.includes(headerName)) {
        console.log("Header names are:-", headerName)
        // 1) Build a DataValidation rule for a dropdown
        const rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(finalList, true) // true = show dropdown, disallow free-form
          .setAllowInvalid(false)
          .build();

        // 2) Apply this rule to the correct range
        //    Range starts at row 2 (below header), column = headerIndex+1
        //    Goes down through lastRow
        sheet
          .getRange(2, headerIndex + 1, lastRow - 1, 1)
          // "-1" if row 1 is header
          .setDataValidation(rule).set;
      }
      else
        continue;
    }
  }


  lockAndProtectEmployeeList() {
    const sheet = SpreadsheetApp.openById("1kSeGBDP12Yscjbk1ZjAkq0RlIt5MXCOhGUJ3bNLw5K4").getSheetByName("Employee List");
    sheet.hideSheet();
    const protect = sheet.protect();
    protect.addEditor("sreenjay.sen@upthink.com");
  }

}






//old code
// function runleaveManagement() {
//   const x = new LeaveManagement();
//   x.fillUpEmployeeIndicatorForEverySubject();
// }


// class LeaveManagement {
//   constructor() {
//     this.employeeDetails = new EmployeeDetails();

//     // Get the array [headers, data] for active employees
//     this.activeEmployeeList = this.employeeDetails.getAllActiveEmployees();

//     // Deconstruct if needed or just store as-is
//     const [employeeSheetHeaders, employeeData] = this.activeEmployeeList;
//     this.activeEmployeeList = [employeeSheetHeaders, employeeData];

//     // Map each subject to its spreadsheet ID
//     this.sheetLinks = {
//       // "Biology":           "1MlBc2V4VZgj3AfJQGdAberbqZAK_gbgwJ37WuRNv7w8",
//       // "Business":          "1Scfv9aGqJrWCYaaZmKlgBVZIUES6r8e8FK9W1KMk3PY",
//       // // "Chemistry":         "1s6T_ua5gns4sC7h_5ljT8lcONRLKZX62w4LeQTgFeJ4",
//       // "Computer Science":  "1g-dfSEG6VbzMDe2gG-l_ir5gjLyNaME11iWlmReMleI",
//       // "English":           "1wcNARFXwACIwsFQrf6qfhuwbGoZ_iVU00ps4EfpIGTk",
//       // "Mathematics":       "1kSeGBDP12Yscjbk1ZjAkq0RlIt5MXCOhGUJ3bNLw5K4",
//       // "Physics":           "1I1CvwEAEj1zfipvqxSoR1TgMMWgsYdhrZDaST7GhDHA",
//       "Statistics": "1_aELI5paQ4OxVo9CuUSGAIn-A9UByRXjP3y8TEDYNaU"
//     };
//   }

//   /**
//    * Loops over every subject and fills up indicator data
//    */
//   fillUpEmployeeIndicatorForEverySubject() {
//     Object.keys(this.sheetLinks).forEach(subject => {
//       console.log("Subject is", subject);
//       this.fillUpEmployeeIndicatorForOneSubject(subject)
//     });

//   }

//   /**
//    * Fills employee data for the given subject into the subject's spreadsheet
//    */
//   fillUpEmployeeIndicatorForOneSubject(subject) {
//     // Use the EmployeeDetails method to get employees for a subject
//     const employeeList = this.employeeDetails.getEmployeesForASubject(
//       this.activeEmployeeList,
//       subject
//     );


//     const nameList = []
//     // Open the subject-specific spreadsheet
//     const spreadsheet = SpreadsheetApp.openById(this.sheetLinks[subject]);
//     const sheetName = "Employee List";
//     let employeeListSheet = spreadsheet.getSheetByName(sheetName);
//     //const databaseSheet = spreadsheet.getSheetByName("Leave_data");
//     const databaseSheet = spreadsheet.getSheetByName("Copy of Leave_data");

//     const alreadyPresentUUIDs = employeeListSheet.getRange(2, 1, employeeListSheet.getLastRow() + 1).getValues().flat().filter(Boolean);

//     employeeList.forEach(row => {
//       if (!alreadyPresentUUIDs.includes(row[0])) {
//         nameList.push(row);
//       }
//     })

//     console.log("Name list is:-", nameList)

//     if (nameList.length <= 0) return;


    

//     let nameRange;
//     // Create the Employee List sheet if it doesn't exist
//     if (!employeeListSheet) {
//       employeeListSheet = spreadsheet.insertSheet();
//       employeeListSheet.setName(sheetName);
//       // Insert header row at the top
//       nameList.unshift(["UniqueID", "Employee Identifier", "Names"]);
//       nameRange = employeeListSheet.getRange(1, 1, nameList.length, 3);
//     } else {
//       nameRange = employeeListSheet.getRange(employeeListSheet.getLastRow() + 1, 1, nameList.length, 3);
//     }

//     nameRange.setValues(nameList);

//     const lastRow = employeeListSheet.getLastRow();
//     const smeList = employeeListSheet.getRange(2, 2, lastRow, 1).getValues().flat().filter(Boolean);
//     this.updateNames(employeeListSheet, databaseSheet);
//     this.updateUUID(employeeListSheet, databaseSheet);

//     smeList.unshift("HR");
//     this.enableDropdowns(databaseSheet, ["Name of SME", "Name of The Reporting Person", "Approver"], smeList);
//}
//   updateUUID(backendSheet, leaveDataSheet) {
//     const [backendHeaders, backendData] = CentralLibrary.get_Data_Indices_From_Sheet(backendSheet);
//     const obj = Object.fromEntries(
//       backendData.map(row => [row[2], row[0]])
//     );

//     const [leaveHeaders, leaveData] = CentralLibrary.get_Data_Indices_From_Sheet(leaveDataSheet);
//     leaveData.forEach((row, idx) => {
//       const smeName = row[leaveHeaders["Name of SME"]];
//       const reportingName = row[leaveHeaders["Name of The Reporting Person"]];
//       const approverName = row[leaveHeaders["Approver"]]
//       if (obj[smeName]) {
//         leaveDataSheet.getRange(idx + 2, leaveHeaders["UUID SME"] + 1).setValue(obj[smeName]);
//       }
//       if (obj[reportingName]) {
//         leaveDataSheet.getRange(idx + 2, leaveHeaders["UUID Reporting Person"] + 1).setValue(obj[reportingName]);
//       }
//       if (obj[approverName]) {
//         leaveDataSheet.getRange(idx + 2, leaveHeaders["UUID Approver"] + 1).setValue(obj[approverName]);
//       }
//     })
//   }


//   updateNames(backendSheet, leaveDataSheet) {
//     const [backendHeaders, backendData] = CentralLibrary.get_Data_Indices_From_Sheet(backendSheet);
//     const obj = Object.fromEntries(
//       backendData.map(row => [row[0], row[1]])
//     );

//     const [leaveHeaders, leaveData] = CentralLibrary.get_Data_Indices_From_Sheet(leaveDataSheet);
//     leaveData.forEach((row, idx) => {
//       const smeUUID = row[leaveHeaders["UUID SME"]];
//       const reportingUUID = row[leaveHeaders["UUID Reporting Person"]];
//       const approverUUID = row[leaveHeaders["UUID Approver"]]

//       if (obj[smeUUID]) {
//         leaveDataSheet.getRange(idx + 2, leaveHeaders["Name of SME"] + 1).setValue(obj[smeUUID]);
//       }
//       if (obj[reportingUUID]) {
//         leaveDataSheet.getRange(idx + 2, leaveHeaders["Name of The Reporting Person"] + 1).setValue(obj[reportingUUID]);
//       }
//       if (obj[approverUUID]) {
//         leaveDataSheet.getRange(idx + 2, leaveHeaders["Approver"] + 1).setValue(obj[approverUUID]);
//       }
//     })
//   }

//   enableDropdowns(sheet, smeHeadersColumnArray, smeList) {
//     // Employee Identifier dropdowns
//     const lastRow = sheet.getLastRow();
//     const smeNamesIndices = CentralLibrary.createIndexMap(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues().flat().filter(Boolean));
//     for (const [headerName, headerIndex] of Object.entries(smeNamesIndices)) {
//       if (smeHeadersColumnArray.includes(headerName)) {
//         console.log("Header names are:-", headerName)
//         // 1) Build a DataValidation rule for a dropdown
//         const rule = SpreadsheetApp.newDataValidation()
//           .requireValueInList(smeList, true) // true = show dropdown, disallow free-form
//           .setAllowInvalid(false)
//           .build();

//         // 2) Apply this rule to the correct range
//         //    Range starts at row 2 (below header), column = headerIndex+1
//         //    Goes down through lastRow
//         sheet
//           .getRange(2, headerIndex + 1, lastRow - 1, 1)
//           // "-1" if row 1 is header
//           .setDataValidation(rule).set;
//       }
//       else
//         continue;
//     }
//   }

//   lockAndProtectEmployeeList() {
//     const sheet = SpreadsheetApp.openById("1kSeGBDP12Yscjbk1ZjAkq0RlIt5MXCOhGUJ3bNLw5K4").getSheetByName("Employee List");
//     sheet.hideSheet();
//     const protect = sheet.protect();
//     protect.addEditor("sreenjay.sen@upthink.com");
//   }
// }

