function settingBrainFuseDataMIS() {
  // Retrieve data from brainfuse using the provided id
  const spreadsheet = SpreadsheetApp.openById("1XH5eGI4thuz-yAsHDeVbI90KFX2AWee3w556yUcGe6I");
  const dataToCopy = brainfuse(spreadsheet);

  // Access the "MasterData" sheet in the spreadsheet
  const ss = SpreadsheetApp.openById("1TwMAUV7-G93Rbh51KL4yaE_KOA9WkCHVhLovGXQ76_E").getSheetByName("MIS");

  
  const dataRange = ss.getRange(1, 1, ss.getLastRow(), ss.getLastColumn()).getValues();
  let headers = dataRange[0], data = dataRange.slice(1);

  // Find the indices of relevant columns in the headers
  const srNoIdx = headers.indexOf("Sr. No.");
  const subIdx = headers.indexOf("Subjects");
  const clientIdx = headers.indexOf("Clients");
  const modeIdx = headers.indexOf("Mode");
  const unitsIdx = headers.indexOf("Units");
  const categoryIdx = headers.indexOf("Category");
  const rateIdx = headers.indexOf("Rate/Hr (In $)");
  const semesterIdx = headers.indexOf("Semester");
  const finYearIdx = headers.indexOf("Financial Year");
  const monthIdx = headers.indexOf("Month");
  const yearIdx = headers.indexOf("Year");
  const numHrsIdx = headers.indexOf("Number of Hours/Essays");
  const revenueIdx = headers.indexOf("Revenue (In $)");

  let objList = []; // To store objects that don't have a match in the data

  let foundMatch = false; // Flag to check if a match is found while iterating

  // Iterate over each object in dataToCopy
  dataToCopy.forEach(obj => {
    // Iterate over each key-value pair in the object
    for (const [key, val] of Object.entries(obj)) {
      let value = val;
      foundMatch = false;

      if (data.length > 0){
        // Iterate over each row in the data
        data.forEach((row, idx) => {
          // Check if the row corresponds to Brainfuse
          if (row[clientIdx] === 'Brainfuse') {
            // Check if the key, subject, and year match in the row and the current object
            if (key === row[monthIdx] && row[subIdx] === value["subject"] && row[yearIdx] == value.year) {
              // Check the units and update the corresponding cell
              if (row[unitsIdx] === 'Offline') {
                const currentValue = ss.getRange(idx + 2, numHrsIdx + 1).getValue();
                ss.getRange(idx + 2, numHrsIdx + 1).setValue(currentValue + value["Offline"]);
              } else if (row[modeIdx] === "Online") {
                if (row[unitsIdx] === "Single_Subject_Session") {
                  const currentValue = ss.getRange(idx + 2, numHrsIdx + 1).getValue();
                  ss.getRange(idx + 2, numHrsIdx + 1).setValue(currentValue + value["Single_Subject_Session"]);
                } else if (row[unitsIdx] === "Single_Subject_Idle") {
                  const currentValue = ss.getRange(idx + 2, numHrsIdx + 1).getValue();
                  ss.getRange(idx + 2, numHrsIdx + 1).setValue(currentValue + value["Single_Subject_Idle"]);
                } else if (row[unitsIdx] === "Multiple_Subject_Session") {
                  const currentValue = ss.getRange(idx + 2, numHrsIdx + 1).getValue();
                  ss.getRange(idx + 2, numHrsIdx + 1).setValue(currentValue + value["Multiple_Subject_Session"]);
                } else if (row[unitsIdx] === "Multiple_Subject_Idle") {
                  const currentValue = ss.getRange(idx + 2, numHrsIdx + 1).getValue();
                  ss.getRange(idx + 2, numHrsIdx + 1).setValue(currentValue + value["Multiple_Subject_Idle"]);
                }
              }
              foundMatch = true;
            }
          }
        });
      }

      // If no match is found, add the key-value pair to the objList
      if (!foundMatch) {
        objList.push({ key, value });
      }
    }
  });

  let lastRowIdx = 2; // Get the index of the next available row
  let lastSerialNum = 0; // Get the last serial number from the sheet

  // Function to set the values for a new row based on unitType and category.
  const setNewRowValues = (obj, unitType, category) => {
    lastSerialNum += 1; // Increment the serial number      
    ss.getRange(lastRowIdx, srNoIdx + 1)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true)
      .setValue(lastSerialNum);

    ss.getRange(lastRowIdx, subIdx + 1)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true)
      .setValue(obj['subject']);

    ss.getRange(lastRowIdx, clientIdx + 1)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true)
      .setValue("Brainfuse");

    ss.getRange(lastRowIdx, modeIdx + 1)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true)
      .setValue("Online");

    ss.getRange(lastRowIdx, unitsIdx + 1)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true)
      .setValue(unitType);

    ss.getRange(lastRowIdx, categoryIdx + 1)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true)
      .setValue(category);
    ss.getRange(lastRowIdx, monthIdx + 1)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true)
      .setValue(obj["month"]);

    ss.getRange(lastRowIdx, semesterIdx + 1)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true)
      .setValue(getSeason(obj["month"]));

      ss.getRange(lastRowIdx, finYearIdx + 1)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true)
      .setValue(obj['finYear']);

    ss.getRange(lastRowIdx, yearIdx + 1)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true)
      .setValue(obj["year"]);

    ss.getRange(lastRowIdx, numHrsIdx + 1)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true)
      .setValue(obj[unitType]);    // Number of hours based on unit type and category.

    ss.getRange(lastRowIdx, revenueIdx + 1)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true)
      .setFormula(multiplyTwoColumns(rateIdx+1, numHrsIdx+1, lastRowIdx)); // Set the formula for revenue in $
    lastRowIdx += 1  // Increment the index of the last Row.
  }

  // Iterate over each object in the objList
  for (const { key, value } of objList) {
    setNewRowValues(value, "Single_Subject_Session", "Occupied");
    setNewRowValues(value, "Single_Subject_Idle", "Unoccupied");
    setNewRowValues(value, "Multiple_Subject_Session", "Occupied");
    setNewRowValues(value, "Multiple_Subject_Idle", "Unoccupied");
    setNewRowValues(value, "Offline", "");
  }
}


function brainfuse(spreadSheet) {

  const sheets = spreadSheet.getSheets();
  const array = [];

  let ss;
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    
    if (sheetName.trim() === 'Calculus' || sheetName.trim() === "Statistics" ||
          sheetName.trim() === 'English' || sheetName.trim() === "Chemistry" ||
          sheetName.trim() === 'Physics' || sheetName.trim() === "Biology" ||
          sheetName.trim() === 'Finance' || sheetName.trim() === "Economics" ||
          sheetName.trim() === 'Computer Science' || sheetName.trim() === "Intro Accounting") {
      // if (sheetName.trim() === 'Calculus'){
      ss = spreadSheet.getSheetByName(sheetName);
      Logger.log(sheetName)
      
      if (ss.getLastRow() > 0){
        const dateColumn = ss.getRange(1, 1, ss.getLastRow(), 1).getValues().flat();
        let headers = ss.getRange(2, 1, 1, ss.getLastColumn()).getValues().flat();
        let totalColumn = headers.indexOf("Total") + 1;
        const startRowSingleSubject = dateColumn.indexOf("Date") + 1;
        const totalHrsRowSingleSubject = dateColumn.indexOf("Total") + 1;
        const valuesSingleSub = ss.getRange(startRowSingleSubject+1,totalColumn,totalHrsRowSingleSubject-startRowSingleSubject, 3)
                                  .getValues();
        const datesSingleSub = dateColumn.slice(startRowSingleSubject, totalHrsRowSingleSubject - 1);
        
        //single subject are converted to formatted array
        const updatedValuesSingleSub = valuesSingleSub.map((element, index) => {
          if (index === 0)
            return [dateColumn[1], ...element];
          else
            return [datesSingleSub[index], ...element];
        }).filter(element => element[0] != "" && element[0] != undefined);


        const startRowMultSubject = dateColumn.lastIndexOf("Date") + 1;

        let updatedValuesMulSub = []
        if (startRowMultSubject !== startRowSingleSubject) {
          headers = ss.getRange(startRowMultSubject, 1, 1, ss.getLastColumn()).getValues().flat();
          totalColumn = headers.indexOf("Total") + 1;
          const totalHrsRowMulSubject = dateColumn.lastIndexOf("Total") + 1;
          const valuesMulSub = ss.getRange(startRowMultSubject+1, totalColumn, totalHrsRowMulSubject - startRowMultSubject, 3)
                                .getValues();

          const datesMulSub = dateColumn.slice(startRowMultSubject, totalHrsRowMulSubject - 1);
          
          updatedValuesMulSub = valuesMulSub.map((element, index) => {
            if (index === 0)
              return [dateColumn[1], ...element];
            else
              return [datesMulSub[index], ...element];
          }).filter(element => element[0] != "" && element[0] != undefined);

        }

        const combinedValuesSingleNMul = []
        updatedValuesSingleSub.forEach((row, i) =>{
          if (i === 0){
            combinedValuesSingleNMul.push(row)
          }else if (updatedValuesMulSub.length > 0){
            combinedValuesSingleNMul.push([row[0],[row[1], updatedValuesMulSub[i][1]], [row[2], updatedValuesMulSub[i][2]], 
            row[3] + updatedValuesMulSub[i][3]])
          }else {
            combinedValuesSingleNMul.push([row[0],[row[1], 0], [row[2], 0], 
            row[3] + 0])
          }
        })

        result = {};
        const keys = ["subject", "firstDay", "lastDay", "month", "year", "finYear"];

        combinedValuesSingleNMul.slice(1).forEach(([dateObj, ...values]) => {
          const modifiedDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd-MMMM-YYYY");
          const [date, month, year] = modifiedDate.split("-");
          if (!result.hasOwnProperty(month)){
            result[month] = {}
            result[month]["Single_Subject_Session"] = Number(values[0][0])
            result[month]["Single_Subject_Idle"] = Number(values[1][0])
            result[month]["Multiple_Subject_Session"] = Number(values[0][1])
            result[month]["Multiple_Subject_Idle"] = Number(values[1][1])
            result[month]["Offline"] = Number(values[2])
            keys.forEach((key, index) => {
            if (key === "year")
              result[month][key] = year;
            else if (key === "firstDay")
              result[month].firstDay = date;
            else if (key === "lastDay")
              result[month][key] = date;
            else {
              result[month][key] = 0;
              if (sheetName === "Calculus")
                result[month].subject = "Mathematics";
              else if(sheetName === "SEO")
                result[month].subject = "Search Engine Optimization";
              else
                result[month].subject = sheetName;
              result[month].finYear = getFinancialYear(month, year);
              result[month].month = month;
            }
        });
          }else {
            result[month]["Single_Subject_Session"] += Number(values[0][0]);
            result[month]["Single_Subject_Idle"] += Number(values[1][0]);
            result[month]["Multiple_Subject_Session"] += Number(values[0][1]);
            result[month]["Multiple_Subject_Idle"] += Number(values[1][1]);
            result[month]["Offline"] += Number(values[2]);
            result[month].lastDay = date;
          }
        })
        array.push(result);
      }
    }
  });
  // Logger.log(array)
  return array;
}