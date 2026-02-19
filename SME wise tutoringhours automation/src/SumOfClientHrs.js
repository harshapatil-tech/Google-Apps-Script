function smeWiseClientHrs() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = sheet.getSheetByName("Summary");
  const outputSheet = sheet.getSheetByName("SMEwise_ClientHours");

  const data = inputSheet.getRange(1, 1, inputSheet.getLastRow(), inputSheet.getLastColumn()).getValues()
  const inputHeaders = data[0]
  const inputData = data.slice(1,)

  const yearIdx = inputHeaders.indexOf("Year");
  const monthIdx = inputHeaders.indexOf("Month");
  const subjectIdx = inputHeaders.indexOf("Subject");
  const smeIdx = inputHeaders.indexOf("SME Name");
  const clientIdx = inputHeaders.indexOf("Client");
  const dayNightIdx = inputHeaders.indexOf("Day/Night");
  const hoursIdx = inputHeaders.indexOf("Hours");

  //Get the values of the dropdowns
  const yearDropdown = outputSheet.getRange('B3').getValue();
  const monthStartDropdown = outputSheet.getRange('D4').getValue();
  const monthEndDropdown = outputSheet.getRange('E4').getValue();
  const subjectDropdown = outputSheet.getRange('H3').getValue();

  // Clear Previous entries
  clearRowsContent(outputSheet, 8, 1, 10)
  let filteredData;
  if (getMonthNumber(monthStartDropdown) <= getMonthNumber(monthEndDropdown)){
    filteredData = inputData.filter(row => row[yearIdx] === yearDropdown)
                            .filter(row => getMonthNumber(row[monthIdx]) >= getMonthNumber(monthStartDropdown) && 
                              getMonthNumber(row[monthIdx]) <= getMonthNumber(monthEndDropdown)
                            )
                            .filter(row =>{ 
                              if (subjectDropdown === 'All')
                                return true
                              else 
                              return row[subjectIdx] === subjectDropdown
                            })
  }else if (getMonthNumber(monthStartDropdown) > getMonthNumber(monthEndDropdown)){
    filteredData = inputData.filter(row => row[yearIdx] == yearDropdown && 
                            getMonthNumber(row[monthIdx]) >= getMonthNumber(monthStartDropdown) ||
                            (row[yearIdx] == yearDropdown + 1 && 
                              getMonthNumber(row[monthIdx]) <= getMonthNumber(monthEndDropdown)
                            ))
                          .filter(row =>{ 
                            if (subjectDropdown === 'All')
                              return true
                            else 
                            return row[subjectIdx] === subjectDropdown
                          })
  }

  const clientNames = setData(inputData, clientIdx)
  const mapObject = {}

  filteredData.forEach(row => { 
    const smeName = row[smeIdx];
    const clientName = row[clientIdx];
    const dayNight = row[dayNightIdx];
    const hours = row[hoursIdx];

    if (!mapObject[smeName]) {
      mapObject[smeName] = {};
    }

    clientNames.forEach(name => {
      if (!mapObject[smeName][name]) {
        mapObject[smeName][name] = {
          "Day": 0,
          "Night": 0
        };
      }
    });
    if (!mapObject[smeName][clientName]) {
      mapObject[smeName][clientName] = {
        "Day": 0,
        "Night": 0
      };
    }
    mapObject[smeName][clientName][dayNight] += hours;
  });

  // Output Sheet Indices
  const outputSheetHeadersTop = outputSheet.getRange(6, 1, 1, outputSheet.getLastColumn()).getValues().flat();
  const outputSheetHeadersBottom = outputSheet.getRange(7, 1, 1, outputSheet.getLastColumn()).getValues().flat();
  
  const smarthinkingIdx = outputSheetHeadersTop.indexOf("Smarthinking")
  const brainfuseIdx = outputSheetHeadersTop.indexOf("Brainfuse");
  const netTutorIdx = outputSheetHeadersTop.indexOf("NetTutor");
  const outputSMEIdx = outputSheetHeadersTop.indexOf("SME Name");
  const totalIdx = outputSheetHeadersTop.indexOf("Total Hours");
  const srNoIdx = outputSheetHeadersTop.indexOf("Sr. No.");

  let currentRow = 7, srNo = 0
  for(const [nameKey, nameValue] of Object.entries(mapObject)){
    currentRow += 1, srNo += 1;
    outputSheet.getRange(currentRow, srNoIdx + 1)
            .setBorder(true, true, true, true, true, true)
            .setFontFamily('Roboto')
            .setVerticalAlignment("middle")
            .setHorizontalAlignment("center")
            .setWrap(true)
            .setValue(srNo)
    outputSheet.getRange(currentRow, outputSMEIdx+1)
               .setBorder(true, true, true, true, true, true)
               .setFontFamily('Roboto')
               .setVerticalAlignment("middle")
               .setHorizontalAlignment("center")
               .setWrap(true)
               .setValue(nameKey)
    for(const [clientKey, clientValue] of Object.entries(nameValue)){
      if (clientKey === 'ST'){
        outputSheet.getRange(currentRow, smarthinkingIdx+1)
                   .setBorder(true, true, true, true, true, true)
                   .setFontFamily('Roboto')
                   .setVerticalAlignment("middle")
                   .setHorizontalAlignment("center")
                   .setWrap(true)
                   .setValue(clientValue['Day'].toFixed(2))
        outputSheet.getRange(currentRow, smarthinkingIdx+2)               
                   .setBorder(true, true, true, true, true, true)
                   .setFontFamily('Roboto')
                   .setVerticalAlignment("middle")
                   .setWrap(true)
                   .setHorizontalAlignment("center")
                   .setValue(clientValue['Night'].toFixed(2))
      }if (clientKey === 'BF'){
        outputSheet.getRange(currentRow, brainfuseIdx+1)
                   .setBorder(true, true, true, true, true, true)
                   .setFontFamily('Roboto')
                   .setVerticalAlignment("middle")
                   .setHorizontalAlignment("center")
                   .setWrap(true)
                   .setValue(clientValue['Day'].toFixed(2))
        outputSheet.getRange(currentRow, brainfuseIdx+2)               
                   .setBorder(true, true, true, true, true, true)
                   .setFontFamily('Roboto')
                   .setVerticalAlignment("middle")
                   .setHorizontalAlignment("center")
                   .setWrap(true)
                   .setValue(clientValue['Night'].toFixed(2))
      }if (clientKey === 'NT'){
        outputSheet.getRange(currentRow, netTutorIdx+1)               
                   .setBorder(true, true, true, true, true, true)
                   .setFontFamily('Roboto')
                   .setVerticalAlignment("middle")
                   .setHorizontalAlignment("center")
                   .setWrap(true)
                   .setValue(clientValue['Day'].toFixed(2))
        outputSheet.getRange(currentRow, netTutorIdx+2)   
                   .setBorder(true, true, true, true, true, true)
                   .setFontFamily('Roboto')
                   .setVerticalAlignment("middle")
                   .setHorizontalAlignment("center")
                   .setWrap(true)
                   .setValue(clientValue['Night'].toFixed(2))
      }
    }
    // Calculate totals for each row
    outputSheet.getRange(currentRow, totalIdx + 1)
               .setBorder(true, true, true, true, true, true)
               .setFontFamily('Roboto')
               .setVerticalAlignment("middle")
               .setHorizontalAlignment("center")
               .setWrap(true)
               .setFormula(`=SUM(${numberToLetter(smarthinkingIdx+1)}${currentRow}, ${numberToLetter(brainfuseIdx+1)}${currentRow}, ${numberToLetter(netTutorIdx+1)}${currentRow})`);
    outputSheet.getRange(currentRow, totalIdx + 2)
               .setBorder(true, true, true, true, true, true)
               .setFontFamily('Roboto')
               .setVerticalAlignment("middle")
               .setHorizontalAlignment("center")
               .setWrap(true)
               .setFormula(`=SUM(${numberToLetter(smarthinkingIdx+2)}${currentRow}, ${numberToLetter(brainfuseIdx+2)}${currentRow}, ${numberToLetter(netTutorIdx+2)}${currentRow})`);
  }
}


function smeSubjectMapping(data) {
  subjectSMEMapping = {};
  const subjectSet = [... new Set(data.map(row=>row[4]))]
  subjectSet.forEach(subject => {
    const smeNames = data.filter(row => row[4] === subject).map(row => row[0])
    subjectSMEMapping[subject] = smeNames;
  })
  
  return subjectSMEMapping;
}


// function onEdit(e) {
//   const editedRange = e.range;
//   const editedSheet = editedRange.getSheet();
//   const editedCell = editedRange.getA1Notation();

//   if (editedSheet.getName() === 'SMEwise_ClientHours' && editedCell === 'H3') {
//     const sheet = SpreadsheetApp.getActiveSpreadsheet();
//     const inputSheet = sheet.getSheetByName("Summary");
//     const data = inputSheet.getRange(1, 1, inputSheet.getLastRow(), inputSheet.getLastColumn()).getValues();

//     const subjectSMEMapping = smeSubjectMapping(data);

//     const outputSheet = sheet.getSheetByName("SMEwise_ClientHours");
//     const subjectDropdown = outputSheet.getRange('H3');
//     const smeNames = subjectSMEMapping[subjectDropdown.getValue()] || ['All'];

//     const validationRule = SpreadsheetApp.newDataValidation()
//       .requireValueInList(smeNames, true)
//       .build();

//     const outputRange = outputSheet.getRange('K3');
//     outputRange.clearContent();
//     outputRange.setDataValidation(validationRule);
//   }
// }



Number.prototype.round = function(places) {
  return +(Math.round(this + "e+" + places)  + "e-" + places);
}















// Helper Functions for getting unique entries from the data
const setData = (data, idx) => [... new Set(data.map(row => row[idx]))];

function setValuesOutputSheet(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = sheet.getSheetByName("Summary");
  const outputSheet = sheet.getSheetByName("SMEwise_ClientHours");

  const data = inputSheet.getRange(1, 1, inputSheet.getLastRow(), inputSheet.getLastColumn()).getValues()
  const inputHeaders = data[0]
  const inputData = data.slice(1,)

  const yearIdx = inputHeaders.indexOf("Year");
  const monthIdx = inputHeaders.indexOf("Month");
  const subjectIdx = inputHeaders.indexOf("Subject");

  const yearSet = setData(inputData, yearIdx);
  const monthSet = setData(inputData, monthIdx);
  const subjectSet = [... setData(inputData, subjectIdx), 'All'];

  const yearValues = SpreadsheetApp.newDataValidation()
                                          .requireValueInList(yearSet)
                                          .setAllowInvalid(false)
                                          .build();

  const monthValues = SpreadsheetApp.newDataValidation()
                                          .requireValueInList(monthSet)
                                          .setAllowInvalid(false)
                                          .build();

  const subjectValues = SpreadsheetApp.newDataValidation()
                                          .requireValueInList(subjectSet)
                                          .setAllowInvalid(false)
                                          .build();

  outputSheet.getRange('B3')
        .clearContent()
        .clearDataValidations()
        .setDataValidation(yearValues)
        .setValue(yearValues[0])
        .setBackground("#dcdcd0")
        .setHorizontalAlignment("center")
        .setBorder(true, true, true, true, true, true)
        .setFontFamily("Roboto");

  outputSheet.getRange('D4')
        .clearContent()
        .clearDataValidations()
        .setDataValidation(monthValues)
        .setValue(monthValues[0])
        .setBackground("#dcdcd0")
        .setHorizontalAlignment("center")
        .setBorder(true, true, true, true, true, true)
        .setFontFamily("Roboto");

  outputSheet.getRange('E4')
        .clearContent()
        .clearDataValidations()
        .setDataValidation(monthValues)
        .setValue(monthValues[0])
        .setBackground("#dcdcd0")
        .setHorizontalAlignment("center")
        .setBorder(true, true, true, true, true, true)
        .setFontFamily("Roboto");
             
  outputSheet.getRange('H3')
          .clearContent()
          .clearDataValidations()
          .setDataValidation(subjectValues)
          .setValue(subjectValues[0])
          .setBackground("#dcdcd0")
          .setHorizontalAlignment("center")
          .setBorder(true, true, true, true, true, true)
          .setFontFamily("Roboto");
  
}






function getMonthNumber(month) {
  
  var monthMapping = {
    January: 1,
    February: 2,
    March: 3,
    April: 4,
    May: 5,
    June: 6,
    July: 7,
    August: 8,
    September: 9,
    October: 10,
    November: 11,
    December: 12
  };

  // Convert the month to title case (e.g., "january" to "January")
  var formattedMonth = month.charAt(0).toUpperCase() + month.slice(1).toLowerCase();

  // Check if the formatted month exists in the mapping
  if (monthMapping.hasOwnProperty(formattedMonth)) {
    return monthMapping[formattedMonth];
  } else {
    return 'Invalid month entered.';
  }
}




function clearRowsContent(sheet, startRow, startColumn, endColumn) {
  const endRow = sheet.getLastRow(); // Ending row to clear
  Logger.log("End row " + endRow)
  if ((endRow-startRow) > 0){
    sheet.getRange(startRow, startColumn, endRow - startRow + 1, endColumn).clear();
  }
}


function numberToLetter(columnNumber) {
  let columnName = '';
  while (columnNumber > 0) {
    let remainder = (columnNumber - 1) % 26;
    columnName = String.fromCharCode(65 + remainder) + columnName;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return columnName;
}

