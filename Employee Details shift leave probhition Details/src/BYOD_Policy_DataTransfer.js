

// function copyData_BYOD() {
//   const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
//   const employeeDetailsSheet = sourceSpreadsheet.getSheetByName("Employee Info");

//   const destinationSpreadsheet = SpreadsheetApp.openById("1D0NOkjL5kOuzG3tzg4H4PdcTqIuaiZRFEn1DhoR3xJg");
//   const destinationSheet = destinationSpreadsheet.getSheetByName("Undertaking Status");

//   const dataRange = employeeDetailsSheet.getDataRange().getValues();
//   const headerIndices = createIndexMap(dataRange[0]);
//   const data = dataRange.slice(1).filter(row => row[headerIndices["Status"]] === 'Active');

//   const startDate = new Date("2023-12-01");
//   startDate.setHours(0, 0, 0, 0);

//   const destinationDataRange = destinationSheet.getDataRange().getValues();
//   const destinationHeaders = destinationDataRange[0];
//   const destinationData = destinationDataRange.slice(1);
//   const destinationHeaderIndices = createIndexMap(destinationHeaders);

//   const columnMapping = {
//     "Employee Name": "Employee Name",
//     "Grade": "Grade",
//     "Department": "Department",
//     "DOJ": "DOJ",
//     "Reporting Manager Email ID": "Reporting Manager Email ID",
//     "Official Email ID": "Official Email ID",
//     "Department Head Email ID": "Department Head Email ID",
//     "Department Head Name": "Department Head Name"
//   };

//   const employeeMap = employeeMapping(headerIndices, data);
//   const existingEmails = new Set(destinationData.map(row => row[destinationHeaderIndices["Official Email ID"]]).filter(Boolean));

//   const dataToWrite = data
//     .filter(row => row[headerIndices["DOJ"]] >= startDate)
//     .map(row => {
//       const employeeName = row[headerIndices["Employee Name"]];
//       if (!employeeMap[employeeName]) return null;
//       const {
//         head,
//         headEmail,
//         reportingManagerEmail
//       } = employeeMap[employeeName];
//       return {
//         "Employee Name": employeeName,
//         "Grade": row[headerIndices["Grade"]],
//         "Department": row[headerIndices["Department"]],
//         "DOJ": row[headerIndices["DOJ"]],
//         "Reporting Manager Email ID": reportingManagerEmail,
//         "Official Email ID": row[headerIndices["Official Email ID"]],
//         "Department Head Email ID": headEmail,
//         "Department Head Name": head
//       };
//     })
//     .filter(Boolean);

//   let lastRow = destinationSheet.getLastRow();
//   const statusRule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(["Agreed", "Pending", "Closed"]).build();
//   const dateValidation = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();

//   dataToWrite.forEach(record => {
//     if (!existingEmails.has(record["Official Email ID"])) {
//       Object.entries(columnMapping).forEach(([sourceKey, destKey]) => {
//         const colIndex = destinationHeaderIndices[destKey] + 1;
//         const cell = destinationSheet.getRange(lastRow + 1, colIndex);
//         if (record[sourceKey]) {
//           cell.setValue(record[sourceKey]);
//           if (sourceKey === "DOJ") {
//             cell.setDataValidation(dateValidation).setNumberFormat('dd-MMM-yyyy');
//           }
//         } else if (["BYOD Email", "Reminder Email 1", "Reminder Email 2", "Undertaking"].includes(destKey)) {
//           cell.insertCheckboxes();
//         } else if (destKey === "Status") {
//           cell.setDataValidation(statusRule);
//         }
//       });
//       lastRow++;
//     }
//   });

//   applyCustomFormatting(destinationSheet.getRange(2, 1, destinationSheet.getLastRow() - 1, destinationSheet.getLastColumn()));
// }

// // Utility functions remain the same
// function employeeMapping(inputIndices, inputData) {
//   const reportingManagersMap = {};
//   const emailMap = {};

//   function getHeadOfDepartment(employee) {
//     const manager = reportingManagersMap[employee];
//     return manager ? getHeadOfDepartment(manager) : { employee, email: emailMap[employee] };
//   }

//   inputData.forEach(row => {
//     const employee = row[inputIndices["Employee Name"]];
//     const manager = row[inputIndices["Reporting Manager"]];
//     const email = row[inputIndices["Official Email ID"]];
//     emailMap[employee] = email;
//     if (!reportingManagersMap[employee] && !['Apurva Yadav', 'Amogh Chaphalkar'].includes(manager)) {
//       reportingManagersMap[employee] = manager;
//     }
//   });

//   return Object.fromEntries(Object.keys(reportingManagersMap).map(employee => [
//     employee,
//     {
//       head: getHeadOfDepartment(employee).employee,
//       headEmail: getHeadOfDepartment(employee).email,
//       reportingManagerEmail: emailMap[reportingManagersMap[employee]]
//     }
//   ]));
// }

// function createIndexMap(headers) {
//   return headers.reduce((acc, header, index) => {
//     acc[header] = index;
//     return acc;
//   }, {});
// }


function copyData_BYOD() {
  const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const employeeDetailsSheet = sourceSpreadsheet.getSheetByName("Employee Info");

  const destinationSpreadsheet = SpreadsheetApp.openById("1D0NOkjL5kOuzG3tzg4H4PdcTqIuaiZRFEn1DhoR3xJg");
  const destinationSheet = destinationSpreadsheet.getSheetByName("Undertaking Status");

  const dataRange = employeeDetailsSheet.getDataRange().getValues();
  const headerIndices = createIndexMap(dataRange[0]);
  const data = dataRange.slice(1).filter(row => row[headerIndices["Status"]] === 'Active');

  const startDate = new Date("2023-12-01");
  startDate.setHours(0, 0, 0, 0);

  const destinationDataRange = destinationSheet.getDataRange().getValues();
  const destinationHeaders = destinationDataRange[0];
  const destinationData = destinationDataRange.slice(1);
  const destinationHeaderIndices = createIndexMap(destinationHeaders);

  const columnMapping = {
    "Employee Name": "Employee Name",
    "Grade": "Grade",
    "Department": "Department",
    "DOJ": "DOJ",
    "Reporting Manager Email ID": "Reporting Manager Email ID",
    "Official Email ID": "Official Email ID",
    "Department Head Email ID": "Department Head Email ID",
    "Department Head Name": "Department Head Name"
  };

  const employeeMap = employeeMapping(headerIndices, data);
  const existingEmails = new Set(destinationData.map(row => row[destinationHeaderIndices["Official Email ID"]]).filter(Boolean));

  const dataToWrite = data
    .filter(row => row[headerIndices["DOJ"]] >= startDate)
    .map(row => {
      const employeeName = row[headerIndices["Employee Name"]];
      if (!employeeMap[employeeName]) return null;
      const {
        head,
        headEmail,
        reportingManagerEmail
      } = employeeMap[employeeName];
      return {
        "Employee Name": employeeName,
        "Grade": row[headerIndices["Grade"]],
        "Department": row[headerIndices["Department"]],
        "DOJ": row[headerIndices["DOJ"]],
        "Reporting Manager Email ID": reportingManagerEmail,
        "Official Email ID": row[headerIndices["Official Email ID"]],
        "Department Head Email ID": headEmail,
        "Department Head Name": head
      };
    })
    .filter(Boolean);

  let lastRow = destinationSheet.getLastRow();
  const statusRule = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(["Agreed", "Pending", "Closed"]).build();
  const dateValidation = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();

  dataToWrite.forEach(record => {
    if (!existingEmails.has(record["Official Email ID"])) {
      Object.entries(columnMapping).forEach(([sourceKey, destKey]) => {
        const colIndex = destinationHeaderIndices[destKey] + 1;
        const cell = destinationSheet.getRange(lastRow + 1, colIndex);
        if (record[sourceKey]) {
          cell.setValue(record[sourceKey]);
          if (sourceKey === "DOJ") {
            cell.setDataValidation(dateValidation).setNumberFormat('dd-MMM-yyyy');
          }
        } else if (destKey === "Status") {
          cell.setDataValidation(statusRule);
        }
      });
      
      // Insert checkboxes in specified columns
      ["BYOD Email", "Reminder Email 1", "Reminder Email 2", "Undertaking"].forEach(column => {
        if (destinationHeaderIndices[column] !== undefined) {
          const colIndex = destinationHeaderIndices[column] + 1;
          destinationSheet.getRange(lastRow + 1, colIndex).insertCheckboxes();
        }
      });
      
      lastRow++;
    }
  });

  applyCustomFormatting(destinationSheet.getRange(2, 1, destinationSheet.getLastRow() - 1, destinationSheet.getLastColumn()));
}

// Utility functions remain the same
function employeeMapping(inputIndices, inputData) {
  const reportingManagersMap = {};
  const emailMap = {};

  function getHeadOfDepartment(employee) {
    const manager = reportingManagersMap[employee];
    return manager ? getHeadOfDepartment(manager) : { employee, email: emailMap[employee] };
  }

  inputData.forEach(row => {
    const employee = row[inputIndices["Employee Name"]];
    const manager = row[inputIndices["Reporting Manager"]];
    const email = row[inputIndices["Official Email ID"]];
    emailMap[employee] = email;
    if (!reportingManagersMap[employee] && !['Apurva Yadav', 'Amogh Chaphalkar'].includes(manager)) {
      reportingManagersMap[employee] = manager;
    }
  });

  return Object.fromEntries(Object.keys(reportingManagersMap).map(employee => [
    employee,
    {
      head: getHeadOfDepartment(employee).employee,
      headEmail: getHeadOfDepartment(employee).email,
      reportingManagerEmail: emailMap[reportingManagersMap[employee]]
    }
  ]));
}

function createIndexMap(headers) {
  return headers.reduce((acc, header, index) => {
    acc[header] = index;
    return acc;
  }, {});
}
