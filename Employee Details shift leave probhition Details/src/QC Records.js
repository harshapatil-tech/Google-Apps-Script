const QC_RECORDS = {
  "Computer Science": "16s0ATTjTUok5aRj7Zxpeouj85cV-QfAV7-FRuW56STQ",
  // "Statistics": "1gIxlyyx-YCY8QdMCjmI8brN2jC-xZfmBZ_jDpDxEhjw"
}


function run() {
  const qcRecords = new QCRecords();
  for (const[key, value] of Object.entries(QC_RECORDS)) {
    const empByDept = qcRecords.getEmployeesByDept(key);
    const individualSheet = new IndividualQCSheet(value);
    individualSheet.createOrUpdateBESheet(empByDept);
    individualSheet.updateMasterSheet();
  }
  // const bioEmps = qcRecords.getEmployeesByDept("Statistics");
  

  // const individualSheet = new IndividualQCSheet(spreadsheetId);
  // individualSheet.createOrUpdateBESheet(bioEmps);
  // individualSheet.updateMasterSheet();
}


class IndividualQCSheet {

  constructor (ssId) {
    this.ss = SpreadsheetApp.openById(ssId);
    this.backendSheetName = "Employee List"
    this.masterSheetName = "QC_Report"

  }

  createOrUpdateBESheet(data) {
    const allSheets = this.ss.getSheets();
    const found = allSheets.find(sheet => sheet.getName().trim() === this.backendSheetName);
    if (!found) {   //found === -1
      // create a sheet by the name of the backend sheet
    this.ss.insertSheet(this.backendSheetName.trim());
    }
    this._updateBESheet(data);
  }

  //new code
  updateMasterSheet() {
    const backendSheet = this.ss.getSheetByName(this.backendSheetName.trim());
    const [backendHeaders, backendData] = CentralLibrary.get_Data_Indices_From_Sheet(backendSheet);

    const masterSheet  = this.ss.getSheetByName(this.masterSheetName);
    const [masterHeaders, masterData] = CentralLibrary.get_Data_Indices_From_Sheet(masterSheet);

    // const existingKeys = new Set(masterData.map(row => row[masterHeaders["UUID"]]));

    const dataMap = Object.fromEntries( backendData.map(row => [ row[backendHeaders["UUID"]], row[backendHeaders["Employee Identifier"]] ]));

    console.log("Backend Map (UUID → Employee Identifier):", dataMap);


    // for (const [key, value] of Object.entries(dataMap)) {

    // } 

    masterData.forEach((row, idx) => {
      const existingUUID = row[masterHeaders["UUID"]];
      if (dataMap[existingUUID] && existingUUID != "") {
        const rowIdx = idx+2;
        console.log(
          `Row ${rowIdx}: UPDATED → UUID: "${existingUUID}" | SME Name set to: "${dataMap[existingUUID]}"`
        );
        masterSheet.getRange(rowIdx, masterHeaders["SME Name"]+1).setValue(dataMap[existingUUID]);
      }
    })
    const empIdentifierList = backendData
      .map(row => row[backendHeaders["Employee Identifier"]])
      .filter(v => v && String(v).trim() !== ""); // remove empty

    // Only call enableDropdowns if we have some options
    if (empIdentifierList.length) {
      this.enableDropdowns(masterSheet, ["SME Name"], empIdentifierList);
    }
   }






  
  
  
  
  
  // updateMasterSheet() {
  //   const masteSheet = this.ss.getSheetByName(this.masterSheetName);
  //   const backendSheet = this.ss.getSheetByName(this.backendSheetName);

  //   const [backendHeaders, backendData] = CentralLibrary.get_Data_Indices_From_Sheet(backendSheet);

  //   // Build lookup: cleanName → [UUID, Employee Identifier]
  //   const nameToLabel = Object.fromEntries(
  //     backendData.map(row => {
  //       const uuid = row[0].trim();
  //       const identifier = row[1].trim();

  //       // Clean name: remove (E###)
  //       const cleanName = identifier
  //         .replace(/\(E\d+\)/i, "")
  //         .trim()
  //         .toLowerCase();

  //       return [cleanName, [uuid, identifier]];
  //     })
  //   );

  //   const [headers, data] = CentralLibrary.get_Data_Indices_From_Sheet(masteSheet);
  //   const smeNameIdx = headers["SME Name"];
  //   const uuidIdx = headers["UUID"];

  //   data.forEach((row, idx) => {
  //     let name = row[smeNameIdx];

  //     if (!name) {
  //       console.log(`Row ${idx + 2}: EMPTY NAME — skipped`);
  //       return;
  //     }

  //     const originalName = name;
  //     const cleanName = name.trim().toLowerCase();

  //     const newLabel = nameToLabel[cleanName];

  //     if (newLabel) {
  //       const [uuid, employeeIdentifier] = newLabel;

  //       console.log(
  //         `Row ${idx + 2}: MATCH FOUND → SME Name: "${originalName}", ` +
  //         `Identifier: "${employeeIdentifier}", UUID: "${uuid}"`
  //       );

  //       masteSheet.getRange(idx + 2, smeNameIdx + 1).setValue(employeeIdentifier);
  //       masteSheet.getRange(idx + 2, uuidIdx + 1).setValue(uuid);
  //     } else {
  //       console.log(
  //         `Row ${idx + 2}: NO MATCH FOUND → SME Name: "${originalName}" (Clean: "${cleanName}")`
  //       );
  //     }
  //   });
  // }


  _updateBESheet(data) {
    const sheet = this.ss.getSheetByName(this.backendSheetName.trim());
    const [headers, backendData] = CentralLibrary.get_Data_Indices_From_Sheet(sheet);
    const dataMap = Object.fromEntries( data.map(row=> [ row[0], row[2] ] ) );

    const existingKeys = backendData.map(row => row[0]);
    const newRows = [];
    for (const [key, value] of Object.entries(dataMap)) {
      if (existingKeys.includes(key)) {
        const uuidIdx = backendData.findIndex(row => row[0] === key) + 2;
        sheet.getRange(uuidIdx, 2).setValue(value);
      } else {
        newRows.push([key, value]);
      }
    }

    if (newRows.length) {
      const startRow = sheet.getLastRow() + 1;
      const numCols = newRows[0].length;
      sheet.getRange(startRow, 1, newRows.length, numCols).setValues(newRows);
    }
  }


  enableDropdowns(sheet, columnArray, optionsList) {
    const finalList = [...new Set(optionsList.filter(v => v))];
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerIndices = {};
    headers.forEach((h, i) => { if(h) headerIndices[h] = i; });

    columnArray.forEach(headerName => {
      const colIndex = headerIndices[headerName];
      if (colIndex === undefined) return;

      const lastRow = sheet.getLastRow();
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(finalList, true)
        .setAllowInvalid(false)
        .build();

      sheet.getRange(2, colIndex + 1, lastRow - 1, 1).setDataValidation(rule);
    });
  }


}



// updateMasterSheet() {
//     const masteSheet = this.ss.getSheetByName(this.masterSheetName);
//     const backendSheet = this.ss.getSheetByName(this.backendSheetName);
//     const [backendHeaders, backendData] = CentralLibrary.get_Data_Indices_From_Sheet(backendSheet);
    
//     // Make a lookup object;
//     // build a lookup: cleanName → full label (second element)
//     const nameToLabel = Object.fromEntries(
//       backendData.map(row => [ row[row.length - 1].trim().toLowerCase(), [row[0].trim(), row[1].trim() ] ])
//     );

  
//     const [headers, data] = CentralLibrary.get_Data_Indices_From_Sheet(masteSheet);
//     const smeNameIdx = headers["SME Name"];
//     const uuidIdx = headers["UUID"];
//     const nameRange = masteSheet.getRange(2, smeNameIdx+1, masteSheet.getLastRow()-2+1, 1);
//     const oldNames = nameRange.getValues().flat().filter(Boolean);
//     data.forEach((row, idx) => {
//       let name = row[smeNameIdx]
//       if (name === "")
//         return;
//       name = name.trim().toLowerCase();
//       const newLabel = nameToLabel[name];

//       if (newLabel) {
//         const [uuid, employeeIdentifier] = newLabel;
//         masteSheet.getRange(idx+2, smeNameIdx+1).setValue(employeeIdentifier);
//         masteSheet.getRange(idx+2, uuidIdx+1).setValue(uuid);
//       }
//     })

//   }




class QCRecords {

  constructor() {
    const empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Employee Info");
    const [headers, data] = CentralLibrary.get_Data_Indices_From_Sheet(empSheet);
    this.employeeDetails = this._getEmployeeDetails(headers, data);
  }

  getEmployeesByDept(department) {
    return this.employeeDetails[department];
  }


  _getEmployeeDetails(headers, data) {

    const departmentObj = {};

    data.forEach(row => {
      let department = row[headers["Department"]];
      if (department === "Economics" || department === "Finance" || department === "Accounts") {
        department = "Business";
      }
      if (!departmentObj[department]) {
        departmentObj[department] = [];
      } else {
        departmentObj[department].push([row[headers["Unique ID"]], row[headers["Employee Name"]], row[headers["Employee Identifier"]]])
      }
    })
    return departmentObj;
  }
}