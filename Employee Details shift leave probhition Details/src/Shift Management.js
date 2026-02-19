function runShiftManagement() {
  const x = new ShiftManagement();
  x.fillUpEmployeeIndicatorForEverySubject();
}


class ShiftManagement {
  constructor() {
    this.employeeDetails = new EmployeeDetails();

    // Get the array [headers, data] for active employees
    this.activeEmployeeList = this.employeeDetails.getAllActiveEmployees();

    // Deconstruct if needed or just store as-is
    const [employeeSheetHeaders, employeeData] = this.activeEmployeeList;
    this.activeEmployeeList = [employeeSheetHeaders, employeeData];


    // Map each subject to its spreadsheet ID
    this.sheetLinks = {
      "Biology": "1ALYGGzJBaZ2zq9g6810cLmx18Jmxil3qcNJKflPoWrs",
      "Business": "1HQUmQiu0T0EA6TSROar_rP6Q120U6W4OFrNlZSBwC64",
      "Chemistry": "1Jtl6fFrk4mbizWXo0vM2q0Ck52TkqSl-wFkJjzjeKzU",
      "Computer Science": "16c_7uM8wndI5kPWAQ9qn_yY2Vw12X0j0hHgvBpJiLRQ",
      "Mathematics": "1ymgBWhpACz44mwfRlljtUE-hUV3jtsHRebf6ljXVgqQ",
      "Physics": "131BnxaOPzXjwhGGbduvoQZqXlmejen1hTtcCacUoX7Y",
      "Statistics": "1OcVc9epAGA7IjH8N0wclVnO2DA9-uzO2h6VkBIAY6nI"

    };
  }

  /**
   * Loops over every subject and fills up indicator data
   */
  fillUpEmployeeIndicatorForEverySubject() {
    Object.keys(this.sheetLinks).forEach(subject =>
      this.fillUpEmployeeIndicatorForOneSubject(subject)
    );
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
    console.log("Employee List is:-",employeeList);

    const nameList = [];
    console.log("name list", nameList);

    // Open the subject-specific spreadsheet
    const spreadsheet = SpreadsheetApp.openById(this.sheetLinks[subject]);
    const sheetName = "Employee List";
    let employeeListSheet = spreadsheet.getSheetByName(sheetName);
    //const databaseSheet = spreadsheet.getSheetByName("Copy of Data");
    const databaseSheet = spreadsheet.getSheetByName("Data");
    //const databaseSheet = spreadsheet.getSheetByName("Copy Data");

    const alreadyPresentUUIDs = employeeListSheet.getRange(2, 1, employeeListSheet.getLastRow() + 1).getValues().flat().filter(Boolean);
    //console.log("already present uuid",alreadyPresentUUIDs);

    employeeList.forEach(row => {
      const uuid = row[0];         // Unique ID
      const identifier = row[1];   // Employee Identifier (E-code)
      const empName = row[2];      // Employee Name

      if (!alreadyPresentUUIDs.includes(uuid)) {
        //nameList.push(row);
        nameList.push([uuid, identifier, empName]);
      }
    });
    //console.log("EMPLOYEES RETURNED FOR SUBJECT:", employeeList);
    //if (nameList.length <=0) return;

    if (nameList.length > 0) {

      let nameRange;
      // Create the Employee List sheet if it doesn't exist
      if (!employeeListSheet) {
        employeeListSheet = spreadsheet.insertSheet();
        employeeListSheet.setName(sheetName);
        // Insert header row at the top
        //nameList.unshift(["UniqueID","Names"]);
        nameList.unshift(["UniqueID", "Employee Identifier", "Names"]);
        nameRange = employeeListSheet.getRange(1, 1, nameList.length, 3);
      } else {
        nameRange = employeeListSheet.getRange(employeeListSheet.getLastRow() + 1, 1, nameList.length, 3);
      }
      nameRange.setValues(nameList);
    }

    
    const lastRow = employeeListSheet.getLastRow();
    const smeList = employeeListSheet.getRange(2, 2, lastRow, 1).getValues().flat().filter(Boolean);
    
    this.updateUUID(employeeListSheet, databaseSheet);
    this.updateNames(employeeListSheet, databaseSheet);
    this.updateNamesWithEmployeeIdentifier(employeeListSheet, databaseSheet);
    this.enableDropdowns(databaseSheet, ["Originally Assigned", "Adjusted by","Approved by"], smeList);
  }


  updateNamesWithEmployeeIdentifier(employeeListSheet, shiftDataSheet) {
    const [empHeaders, empData] = CentralLibrary.get_Data_Indices_From_Sheet(employeeListSheet);
    const [shiftHeaders, shiftData] = CentralLibrary.get_Data_Indices_From_Sheet(shiftDataSheet);

    const nameCol = empHeaders["Names"];
    const identifierCol = empHeaders["Employee Identifier"];

    // Build: plainName → "Name (E###)"
    const nameMap = {};
    empData.forEach(row => {
      const plain = row[nameCol]?.toString().trim();
      const full = row[identifierCol]?.toString().trim();
      if (plain && full) nameMap[plain] = full;
    });

    const targetColumns = ["Originally Assigned", "Adjusted by","Approved by"];
    

    shiftData.forEach((row, rIdx) => {
      targetColumns.forEach(colName => {
        const colIndex = shiftHeaders[colName];
        if (colIndex === undefined) return;

        const cellVal = row[colIndex];
        if (!cellVal) return;

        const plain = cellVal.toString().trim();
        const mapped = nameMap[plain];
        console.log(
          `Row ${rIdx + 2} | Column "${colName}" | Found: "${plain}" → Replace With: "${mapped}"`
        );

        if (mapped && mapped !== cellVal) {
          const cell = shiftDataSheet.getRange(rIdx + 2, colIndex + 1);
          const old = cell.getDataValidation();
          cell.setDataValidation(null);
          cell.setValue(mapped);
          if (old) cell.setDataValidation(old);
        }
      });
    });
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
        //expandedList.push(match[1].trim()); // clean version

      }
    });


    // Remove duplicates
    const finalList = [...new Set(expandedList)];

    const smeNamesIndices = CentralLibrary.createIndexMap(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues().flat().filter(Boolean));

    for (const [headerName, headerIndex] of Object.entries(smeNamesIndices)) {
      if (smeHeadersColumnArray.includes(headerName)) {
        // 1) Build a DataValidation rule for a dropdown
        const rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(finalList, true) // true = show dropdown, disallow free-form
          .setAllowInvalid(false)
          .build();

        // 2) Apply this rule to the correct range
        //    Range starts at row 2 (below header), column = headerIndex+1
        //    Goes down through lastRow
        // sheet
        //   .getRange(2, headerIndex + 1, lastRow - 1, 1)
        //   // "-1" if row 1 is header
        //   .setDataValidation(rule);

        sheet
          .getRange(2, headerIndex + 1, sheet.getMaxRows() - 1, 1)
          .setDataValidation(rule);

      }
      else
        continue;
    }
  }


  updateUUID(backendSheet, shiftDataSheet) {
    const [backendHeaders, backendData] = CentralLibrary.get_Data_Indices_From_Sheet(backendSheet);
    const obj = Object.fromEntries(
      backendData.map(row => [row[1], row[0]]) //1 //2
    );

    const [shiftHeaders, shiftData] = CentralLibrary.get_Data_Indices_From_Sheet(shiftDataSheet);
    shiftData.forEach((row, idx) => {
      const smeName = row[shiftHeaders["Originally Assigned"]];
      const reportingName = row[shiftHeaders["Adjusted by"]];
      const approvedname = row[shiftHeaders["Approved by"]];

      if (obj[smeName]) {
        shiftDataSheet.getRange(idx + 2, shiftHeaders["UUID Originally Assigned"] + 1).setValue(obj[smeName]);
        console.log(`UUID Originally Assigned set for ${smeName} → ${obj[smeName]}`);
      }
      if (obj[reportingName]) {
        shiftDataSheet.getRange(idx + 2, shiftHeaders["UUID Adjusted By"] + 1).setValue(obj[reportingName]);
        console.log(`UUID Adjusted By set for ${reportingName} → ${obj[reportingName]}`);
      }
      if (obj[approvedname]) {
         shiftDataSheet.getRange(idx + 2, shiftHeaders["UUID Approved By"] + 1).setValue(obj[approvedname]);
         console.log(`UUID Approved By set for ${approvedname} → ${obj[approvedname]}`);
       }
    })
  }


  updateNames(backendSheet, shiftDataSheet) {
    const [backendHeaders, backendData] = CentralLibrary.get_Data_Indices_From_Sheet(backendSheet);
    const obj = Object.fromEntries(
      backendData.map(row => [row[0], row[1]])
    );

    const [shiftHeaders, shiftData] = CentralLibrary.get_Data_Indices_From_Sheet(shiftDataSheet);
    shiftData.forEach((row, idx) => {
      const smeUUID = row[shiftHeaders["UUID Originally Assigned"]];
      const reportingUUID = row[shiftHeaders["UUID Adjusted By"]];
      const approverUUID = row[shiftHeaders["UUID Approved By"]];


      if (obj[smeUUID]) {
        shiftDataSheet.getRange(idx + 2, shiftHeaders["Originally Assigned"] + 1).setValue(obj[smeUUID]);
      }
      if (obj[reportingUUID]) {
        shiftDataSheet.getRange(idx + 2, shiftHeaders["Adjusted by"] + 1).setValue(obj[reportingUUID]);
      }
      if (obj[approverUUID]) {
        shiftDataSheet.getRange(idx + 2, shiftHeaders["Approved by"] + 1).setValue(obj[approverUUID]);
       }

    })
  }
}








//old code
//function runShiftManagement() {
//   const x = new ShiftManagement();
//   x.fillUpEmployeeIndicatorForEverySubject();
// }


// class ShiftManagement {
//   constructor() {
//     this.employeeDetails = new EmployeeDetails();
    
//     // Get the array [headers, data] for active employees
//     this.activeEmployeeList = this.employeeDetails.getAllActiveEmployees();

//     // Deconstruct if needed or just store as-is
//     const [employeeSheetHeaders, employeeData] = this.activeEmployeeList;
//     this.activeEmployeeList = [employeeSheetHeaders, employeeData];
//     // console.log("EMPLOYEE HEADERS:", employeeSheetHeaders);
//     // console.log("EMPLOYEE DATA SAMPLE (first row):", employeeData[0]);


//     // Map each subject to its spreadsheet ID
//     this.sheetLinks = {
//        "Biology":           "1ALYGGzJBaZ2zq9g6810cLmx18Jmxil3qcNJKflPoWrs",
//       // "Business":          "1HQUmQiu0T0EA6TSROar_rP6Q120U6W4OFrNlZSBwC64",
//       // "Chemistry":         "1Jtl6fFrk4mbizWXo0vM2q0Ck52TkqSl-wFkJjzjeKzU",
//       // "Computer Science":  "16c_7uM8wndI5kPWAQ9qn_yY2Vw12X0j0hHgvBpJiLRQ",
//       // "Mathematics":       "1ymgBWhpACz44mwfRlljtUE-hUV3jtsHRebf6ljXVgqQ",
//       // "Physics":           "131BnxaOPzXjwhGGbduvoQZqXlmejen1hTtcCacUoX7Y",
//       // "Statistics":        "1OcVc9epAGA7IjH8N0wclVnO2DA9-uzO2h6VkBIAY6nI"
//     };
//   }

//   /**
//    * Loops over every subject and fills up indicator data
//    */
//   fillUpEmployeeIndicatorForEverySubject() {
//     Object.keys(this.sheetLinks).forEach(subject =>
//       this.fillUpEmployeeIndicatorForOneSubject(subject)
//     );
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
//     const databaseSheet = spreadsheet.getSheetByName("Data");

//     const alreadyPresentUUIDs = employeeListSheet.getRange(2, 1, employeeListSheet.getLastRow() + 1).getValues().flat().filter(Boolean);

//     employeeList.forEach(row => {
//       if ( !alreadyPresentUUIDs.includes(row[0]) ){
//         nameList.push(row);
//   //       const uniqueId = row[0];
//   //       const employeeName = row[1];       // Employee Details sheet se
//   //       const empId = row[2];              // Emp Id column
//   //       const employeeIdentifier = `${employeeName} (${empId})`;

//   //       nameList.push([
//   //         uniqueId,
//   //         employeeIdentifier,
//   //         employeeName
//   //       ]);
//       }
//     });
//   //  console.log("EMPLOYEES RETURNED FOR SUBJECT:", employeeList);
//     if (nameList.length <=0) return;

//     let nameRange;
//     // Create the Employee List sheet if it doesn't exist
//     if (!employeeListSheet) {
//       employeeListSheet = spreadsheet.insertSheet();
//       employeeListSheet.setName(sheetName);
//       // Insert header row at the top
//       nameList.unshift(["UniqueID","Names"]);
//     // nameList.unshift(["UniqueID", "Employee Identifier","Employee Name"]);

//       nameRange = employeeListSheet.getRange(1, 1, nameList.length, 3);
//     } else {
//       nameRange = employeeListSheet.getRange(employeeListSheet.getLastRow()+1, 1, nameList.length, 3);
//     }

//     nameRange.setValues(nameList);

//     const lastRow = employeeListSheet.getLastRow();
//     const smeList = employeeListSheet.getRange(2, 2, lastRow, 1).getValues().flat().filter(Boolean);

//     this.updateUUID(employeeListSheet, databaseSheet);
//     this.updateNames(employeeListSheet, databaseSheet);
    
//     this.enableDropdowns(databaseSheet, ["Originally Assigned", "Adjusted by"], smeList);
//   }

//   enableDropdowns(sheet, smeHeadersColumnArray, smeList) {
//       // Employee Identifier dropdowns
//       const lastRow = sheet.getLastRow();
//       const smeNamesIndices = CentralLibrary.createIndexMap(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues().flat().filter(Boolean));
//       for(const [headerName, headerIndex] of Object.entries(smeNamesIndices)) {
//         if (smeHeadersColumnArray.includes(headerName)){
//           // 1) Build a DataValidation rule for a dropdown
//           const rule = SpreadsheetApp.newDataValidation()
//             .requireValueInList(smeList, true) // true = show dropdown, disallow free-form
//             .setAllowInvalid(false)
//             .build();
          
//           // 2) Apply this rule to the correct range
//           //    Range starts at row 2 (below header), column = headerIndex+1
//           //    Goes down through lastRow
//           sheet
//             .getRange(2, headerIndex + 1, lastRow - 1, 1)
//             // "-1" if row 1 is header
//             .setDataValidation(rule);
//         }
//         else  
//           continue;
//     }
//   }


//   updateUUID(backendSheet, leaveDataSheet) {
//     const [ backendHeaders, backendData] = CentralLibrary.get_Data_Indices_From_Sheet(backendSheet);
//     const obj = Object.fromEntries(
//       backendData.map(row => [ row[2], row[0] ])
//     );
    
//     const [ leaveHeaders, leaveData ] = CentralLibrary.get_Data_Indices_From_Sheet(leaveDataSheet);
//     leaveData.forEach((row, idx) => {
//       const smeName = row[leaveHeaders["Originally Assigned"]];
//       const reportingName = row[leaveHeaders["Adjusted by"]];

//       if (obj[smeName]) {
//         leaveDataSheet.getRange(idx+2, leaveHeaders["UUID Originally Assigned"]+1).setValue(obj[smeName]);
//       }
//       if (obj[reportingName]){
//         leaveDataSheet.getRange(idx+2, leaveHeaders["UUID Adjusted By"]+1).setValue(obj[reportingName]);
//       }
//     })
//   }


//   updateNames(backendSheet, leaveDataSheet) {
//     const [ backendHeaders, backendData] = CentralLibrary.get_Data_Indices_From_Sheet(backendSheet);
//     const obj = Object.fromEntries(
//       backendData.map(row => [ row[0], row[1] ])
//     );
    
//     const [ leaveHeaders, leaveData ] = CentralLibrary.get_Data_Indices_From_Sheet(leaveDataSheet);
//     leaveData.forEach((row, idx) => {
//       const smeUUID = row[leaveHeaders["UUID SME"]];
//       const reportingUUID = row[leaveHeaders["UUID Reporting Person"]];
//       const approverUUID = row[leaveHeaders["UUID Approver"]]

//       if (obj[smeUUID]) {
//         leaveDataSheet.getRange(idx+2, leaveHeaders["Originally Assigned"]+1).setValue(obj[smeUUID]);
//       }
//       if (obj[reportingUUID]){
//         leaveDataSheet.getRange(idx+2, leaveHeaders["Adjusted by"]+1).setValue(obj[reportingUUID]);
//       }

//     })
//   }
// }







