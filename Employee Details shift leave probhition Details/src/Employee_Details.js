function getEmployeeDetails() {
  
  const inputSpreadsheet = SpreadsheetApp.openById("1hDqudXKOjCZ5FV1gv_n1HBcOI9UqnUYruFSDnGBY3sg")
  const inputHeadCountSheet = inputSpreadsheet.getSheetByName("Headcount");
  const inputResignationSheet = inputSpreadsheet.getSheetByName("Resignations");
  const [inputHeaderMap, inputData] = CentralLibrary.get_Data_Indices_From_Sheet(inputHeadCountSheet);

  const outputSpreadsheet = SpreadsheetApp.openById("11b5A_ZDi9DskTiKJ0j4fBDkRbY-0I2h7rVRxr6KVK1o");
  const outputHeadCountSheet = outputSpreadsheet.getSheetByName("Employee Info");
  const [outputHeaderMap, outputData] = CentralLibrary.get_Data_Indices_From_Sheet(outputHeadCountSheet);

  const headerMapping = {
    "Employee Name" : "Employee Name",
    "New Emp Id" : "Emp Id",
    "Function" : "Function",
    "Reporting Manager" : "Reporting Manager",
    "Grade" : "Grade",
    "Designation" : "Designation",
    "Department" : "Department",
    "Gender" : "Gender",
    "DOJ" : "DOJ",
    "Official email ID" : "Official Email ID",
    "Personal contact Number" : "Phone Number",
    // "Location" : "Location",
    "UUID" : "Unique ID",
    "Days":	"Days",
    "Hours": "Hours"
  }

  const requiredHeaders = Object.keys(headerMapping);
  const requiredIndices = requiredHeaders
    .map(header => inputHeaderMap[header])
    .filter(index => index !== undefined); 
    
  const filteredData = inputData.filter(row => {
    return requiredIndices.every(index => {
      const cell = row[index];
      return cell !== "" && cell !== null && cell !== undefined;
    });
  });


  const outputSheetMap = new Map();
  const inputSheetUUIDSet = new Set();

  for (let i=0; i < outputData.length; i++) {
    const row = outputData[i];
    outputSheetMap.set(row[outputHeaderMap["Unique ID"]], i + 2);
  }

  filteredData.forEach( row => {
    const mainKey = row[inputHeaderMap["UUID"]]
    inputSheetUUIDSet.add(mainKey);
    if (outputSheetMap.has(mainKey)) {
      var rowToUpdate = outputSheetMap.get(mainKey);
    
      Object.keys(headerMapping).forEach(inputHeader => {
        const newValue = row[inputHeaderMap[inputHeader]];
        const outputHeader = headerMapping[inputHeader];
        const colIndex = outputHeaderMap[outputHeader] + 1; // +1 for 1-based index
        outputHeadCountSheet.getRange(rowToUpdate, colIndex).setValue(newValue);
      });
      outputHeadCountSheet.getRange(rowToUpdate, outputHeaderMap["Status"] + 1).setValue("Active");
      // outputHeadCountSheet.getRange(rowToUpdate, outputHeaderMap["Employee Identifier"] + 1).setValue(`${row[outputHeaderMap["Employee Name"]]} (${row[outputHeaderMap["Emp Id"]]})`);
      outputSheetMap.delete(mainKey);
    } else {
      console.log("executed")
      let lastRow = outputHeadCountSheet.getLastRow();
      outputHeadCountSheet.insertRowAfter(lastRow);
      Object.keys(headerMapping).forEach(inputHeader => {
        const value = row[inputHeaderMap[inputHeader]];
        const outputHeader = headerMapping[inputHeader];
        const colIndex = outputHeaderMap[outputHeader] + 1; // +1 for 1-based index
        outputHeadCountSheet.getRange(lastRow + 1, colIndex).setValue(value);
      });
      outputHeadCountSheet.getRange(lastRow + 1, outputHeaderMap["Status"] + 1).setValue("Active");
      const empName = outputHeadCountSheet.getRange(lastRow + 1, outputHeaderMap["Employee Name"] + 1).getValue();
      const empId = outputHeadCountSheet.getRange(lastRow + 1, outputHeaderMap["Emp Id"] + 1).getValue();
      outputHeadCountSheet.getRange(lastRow + 1, outputHeaderMap["Employee Identifier"] + 1)
      .setValue(`${empName} (${empId})`);
      outputHeadCountSheet.getRange(lastRow+1, 1, 1, outputHeadCountSheet.getLastColumn()).setBackground("white").setFontWeight("normal").setFontFamily("Roboto")
    }
  });


  // Mark rows that no longer exist in the input sheet as "Separated"
  outputSheetMap.forEach(function(row, key) {
    if (!inputSheetUUIDSet.has(key)) {

      outputHeadCountSheet.getRange(row, outputHeaderMap["Status"] + 1).setValue("Separated");
    }
  });

    // Resignation Code DOL
  const [ipResigHeaders, ipResignData] = CentralLibrary.get_Data_Indices_From_Sheet(inputResignationSheet);
  
  outputData.forEach((row, index) => {
    if (row[outputHeaderMap["Status"]].trim().toLowerCase() === "separated" && row[outputHeaderMap["DOL"]] === "") {
      // Find the "Official Email ID" in resignation sheet
      let dateOfLeaving;
      ipResignData.forEach((resigRow) => {
        if (resigRow[ipResigHeaders["Official email ID"]].trim().toLowerCase() === row[outputHeaderMap["Official Email ID"]].trim().toLowerCase()) {
          const dol = resigRow[ipResigHeaders["DOL"]];
          if (dol !== ""){
            dateOfLeaving = new Date(dol);
          }
          return;
        }
      })
      if (dateOfLeaving !== undefined)
        outputHeadCountSheet.getRange(index + 2, outputHeaderMap["DOL"]+1).setValue(Utilities.formatDate(dateOfLeaving, "IST", "dd-MMM-yyyy"))
    }
  })

  

}




function mapEmployeeDetails() {
  // Open the active spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Define the source and destination sheet names
  const sourceSheetName = 'Employee Info';
  const destinationSheetName = 'Copy of Employee Info';
  
  // Access the sheets
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  const destinationSheet = ss.getSheetByName(destinationSheetName);
  
  if (!sourceSheet || !destinationSheet) {
    Logger.log("One or both sheets do not exist.");
    return;
  }
  
  // Fetch the data from both sheets
  const sourceData = sourceSheet.getDataRange().getValues();
  const destinationData = destinationSheet.getDataRange().getValues();
  
  // Extract header row
  const sourceHeaders = sourceData[0];
  const destinationHeaders = destinationData[0];
  
  // Map columns by headers
  const headerMap = {};
  sourceHeaders.forEach(function(header, index) {
    if (destinationHeaders.includes(header)) {
      headerMap[index] = destinationHeaders.indexOf(header);
    }
  });
  
  // Prepare updated rows
  const updatedRows = [];
  
  for (let i = 1; i < sourceData.length; i++) {
    const sourceRow = sourceData[i];
    const updatedRow = new Array(destinationHeaders.length).fill("");
    
    for (let sourceIndex in headerMap) {
      const destinationIndex = headerMap[sourceIndex];
      updatedRow[destinationIndex] = sourceRow[sourceIndex];
    }
    updatedRows.push(updatedRow);
  }
  
  // Clear and update destination sheet
  destinationSheet.getRange(2, 1, destinationSheet.getLastRow() - 1, destinationSheet.getLastColumn()).clearContent();
  destinationSheet.getRange(2, 1, updatedRows.length, updatedRows[0].length).setValues(updatedRows);
  
  Logger.log("Data has been successfully mapped and updated.");
}























