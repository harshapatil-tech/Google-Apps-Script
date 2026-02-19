// function checkForChanges () {
//   const referenceSheet = SpreadsheetApp.openById("11b5A_ZDi9DskTiKJ0j4fBDkRbY-0I2h7rVRxr6KVK1o").getSheetByName("Employee Info");
//   const [referenceHeaders, referenceData] = CentralLibrary.get_Data_Indices_From_Sheet(referenceSheet);
//   const thisSheet = SpreadsheetApp.openById("1p9SNUd6ud0gM0KEDdN4eilUE7vxxytAg67nAOOVm0dw").getSheetByName("Employee Info");
//   const [thisSheetHeaders, thisSheetData] = CentralLibrary.get_Data_Indices_From_Sheet(thisSheet);

//   // Compare the data arrays
//   if (areArraysEqual(referenceData, thisSheetData)) {
//     Logger.log("No changes detected. The data is identical.");
//     return 
//   } else {
//     Logger.log("Changes detected. The data differs.");

//     // You can add additional actions here, such as sending notifications or updating records
//   }
// }

function myfunction() {
  // const leaveManagement = new LeaveManagement();
  // leaveManagement.fillUpEmployeeIndicatorForEverySubject();

  const shiftManagement = new ShiftManagement();
  shiftManagement.fillUpEmployeeIndicatorForEverySubject();
}


class EmployeeDetails {
  constructor() {
    this.employeeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetById(1550264425);
  }

  /**
     * Returns a 2D array of employees for the specified subject or group of subjects.
     * Each returned row is: [ "UniqueID", "Employee Identifier" ]
  */
  getEmployeesForASubject(activeEmployeeSheetDetails, subject) {
    const [employeeSheetHeaders, employeeData] = activeEmployeeSheetDetails;

    // If subject is "Business", treat "Accounts" and "Finance" as part of the same group
    let subjectList = [subject];
    if (subject === "Business") {
      subjectList = ["Economics", "Accounts", "Finance"];
    }

    // Filter rows whose Department matches any value in subjectList
    const filteredEmployees = employeeData.filter(row => 
      subjectList.includes(row[employeeSheetHeaders["Department"]])
    );

    // Return only the columns [UniqueID, Employee Identifier]
    return filteredEmployees.map(row => [
      row[employeeSheetHeaders["Unique ID"]],
      row[employeeSheetHeaders["Employee Identifier"]],
      row[employeeSheetHeaders["Employee Name"]]
    ]);
  }

  /**
   * Returns a 2D array [ employeeSheetHeaders, filteredEmployeeData ]
   * for all active employees.
   */
  getAllActiveEmployees() {
    // Example usage of your central library function. Make sure it returns
    // [employeeSheetHeaders, employeeSheetData].
    const [employeeSheetHeaders, employeeSheetData] =
      CentralLibrary.get_Data_Indices_From_Sheet(this.employeeSheet);

    // // Filter only "Active" employees
    // const activeData = employeeSheetData.filter(
    //   row => row[employeeSheetHeaders["Status"]] === "Active"
    // );

    return [employeeSheetHeaders, employeeSheetData];
  }
}




