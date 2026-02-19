// function smeWiseClientHrs() {
//   const smeHandler = new SMEWiseClientHours();
//   smeHandler.setValuesOutputSheet();
//   smeHandler.smeWiseClientHrs();
// }


class SMEWiseClientHours {
  constructor() {
    this.sheet = SpreadsheetApp.getActiveSpreadsheet();
    this.inputSheet = this.sheet.getSheetByName("Summary");
    this.outputSheet = this.sheet.getSheetByName("SMEwise_ClientHours");
    this.monthMapping = {
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
  }

  getMonthNumber(month) {
    const formattedMonth = month.charAt(0).toUpperCase() + month.slice(1).toLowerCase(); //format the month name to have the first letter capitalized
    return this.monthMapping[formattedMonth] || 'Invalid month entered.';
  }

  clearRowsContent(startRow, startColumn, endColumn) {
    const endRow = this.outputSheet.getLastRow();  //get the last row of the output sheet
    if ((endRow - startRow) > 0) {
      //clear the range from startRow to endRow for specified columns
      this.outputSheet.getRange(startRow, startColumn, endRow - startRow + 1, endColumn).clear();
    }
  }

  numberToLetter(columnNumber) {
  let columnName = '';
  while (columnNumber > 0) {
    const remainder = (columnNumber - 1) % 26;
    columnName = String.fromCharCode(65 + remainder) + columnName;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return columnName;
  }

  setData(data, idx) {
    return [...new Set(data.map(row => row[idx]))];  //return unique value from a specific column
  }

  //create a mapping of subjects to SME names
  smeSubjectMapping(data) {
    const subjectSMEMapping = {};
    const subjectSet = [...new Set(data.map(row => row[4]))];     //get unique subject
  subjectSet.forEach(subject => {
    const smeNames = data.filter(row => row[4] === subject).map(row => row[0]);   //get sme name for each subject
    subjectSMEMapping[subject] = smeNames;
  });
  return subjectSMEMapping;
  }

  // Helper Functions for getting unique entries from the data
  setValuesOutputSheet() {
    const data = this.inputSheet.getRange(1, 1, this.inputSheet.getLastRow(), this.inputSheet.getLastColumn()).getValues();
    // console.log(data);
    const inputHeaders = data[0];
    const inputData = data.slice(1);

    const yearIdx = inputHeaders.indexOf("Year");
    const monthIdx = inputHeaders.indexOf("Month");
    const subjectIdx = inputHeaders.indexOf("Subject");

    //get unique values for year, month and subject column
    const yearSet = this.setData(inputData, yearIdx);
    const monthSet = this.setData(inputData, monthIdx);
    const subjectSet = [...this.setData(inputData, subjectIdx), 'All'];

    const yearValues = SpreadsheetApp.newDataValidation().requireValueInList(yearSet).setAllowInvalid(false).build();
    const monthValues = SpreadsheetApp.newDataValidation().requireValueInList(monthSet).setAllowInvalid(false).build();
    const subjectValues = SpreadsheetApp.newDataValidation().requireValueInList(subjectSet).setAllowInvalid(false).build();

    const fields = [
      { range: 'B3', validation: yearValues, values: yearSet },
      { range: 'D4', validation: monthValues, values: monthSet },
      { range: 'E4', validation: monthValues, values: monthSet },
      { range: 'H3', validation: subjectValues, values: subjectSet }
    ];

    fields.forEach(field => {
      this.outputSheet.getRange(field.range)
        .clearContent()
        .clearDataValidations()
        .setDataValidation(field.validation)
        .setValue(field.values[0])
        .setBackground("#dcdcd0")
        .setHorizontalAlignment("center")
        .setBorder(true, true, true, true, true, true)
        .setFontFamily("Roboto");
    });
  }

  smeWiseClientHrs() {
    const data = this.inputSheet.getRange(1, 1, this.inputSheet.getLastRow(), this.inputSheet.getLastColumn()).getValues();
    const inputHeaders = data[0];
    const inputData = data.slice(1);

    const yearIdx = inputHeaders.indexOf("Year");
    const monthIdx = inputHeaders.indexOf("Month");
    const subjectIdx = inputHeaders.indexOf("Subject");
    const smeIdx = inputHeaders.indexOf("SME Name");
    const clientIdx = inputHeaders.indexOf("Client");
    const dayNightIdx = inputHeaders.indexOf("Day/Night");
    const hoursIdx = inputHeaders.indexOf("Hours");

    const yearDropdown = this.outputSheet.getRange('B3').getValue();
    const monthStartDropdown = this.outputSheet.getRange('D4').getValue();
    const monthEndDropdown = this.outputSheet.getRange('E4').getValue();
    const subjectDropdown = this.outputSheet.getRange('H3').getValue();

    //clear existing rows in the output sheet before adding new data
    this.clearRowsContent(8, 1, 10);

    //get unique client names
    const clientNames = this.setData(inputData, clientIdx);
    const mapObject = {};   //this object store sme wise client hours

    //filter data based on dropdown values
    const getFilteredData = (data, yearDropdown, monthStartDropdown, monthEndDropdown, subjectDropdown) => {
      const startMonth = this.getMonthNumber(monthStartDropdown);
      const endMonth = this.getMonthNumber(monthEndDropdown);

      return data.filter(row => {
        const rowYear = row[yearIdx];
        const rowMonth = this.getMonthNumber(row[monthIdx]);
        const validMonth = startMonth <= endMonth
          ? rowMonth >= startMonth && rowMonth <= endMonth
          : (rowYear === yearDropdown && rowMonth >= startMonth) ||
            (rowYear === yearDropdown + 1 && rowMonth <= endMonth);
        const validSubject = subjectDropdown === 'All' || row[subjectIdx] === subjectDropdown;
        return rowYear === yearDropdown && validMonth && validSubject;
      });
    };

    //filter data based on user selection
    const filteredData = getFilteredData(inputData, yearDropdown, monthStartDropdown, monthEndDropdown, subjectDropdown);

     //process the filtered data and organize hours by SME and client
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
          mapObject[smeName][name] = { Day: 0, Night: 0 };    //initialize hours for day and night
        }
      });

      if (!mapObject[smeName][clientName]) {
        mapObject[smeName][clientName] = { Day: 0, Night: 0 };  //initialize if client not present
      }

      //increment hours for day and night based on client
      mapObject[smeName][clientName][dayNight] += hours;
    });

    this.populateOutputSheet(mapObject, clientNames);
  } 

  populateOutputSheet(mapObject, clientNames) {
    //get headers from the output sheet
    const outputSheetHeadersTop = this.outputSheet.getRange(6, 1, 1, this.outputSheet.getLastColumn()).getValues().flat();
    console.log(outputSheetHeadersTop)
    const outputSheetHeadersBottom = this.outputSheet.getRange(7, 1, 1, this.outputSheet.getLastColumn()).getValues().flat();
   
    const smarthinkingIdx = outputSheetHeadersTop.indexOf("Smarthinking");
    const brainfuseIdx = outputSheetHeadersTop.indexOf("Brainfuse");
    const netTutorIdx = outputSheetHeadersTop.indexOf("NetTutor");
    const outputSMEIdx = outputSheetHeadersTop.indexOf("SME Name");
    const totalIdx = outputSheetHeadersTop.indexOf("Total Hours");
    const srNoIdx = outputSheetHeadersTop.indexOf("Sr. No.");

    let currentRow = 7;
    let srNo = 0;

    //loop through each sme and populate data
    for (const [nameKey, nameValue] of Object.entries(mapObject)) {
      currentRow += 1;
      srNo += 1;

    //set serial number in the output sheet 
    this.outputSheet.getRange(currentRow, srNoIdx + 1)
      .setBorder(true, true, true, true, true, true)
      .setFontFamily('Roboto')
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("center")
      .setWrap(true)
      .setValue(srNo);

    //set sme name in the output sheet
    this.outputSheet.getRange(currentRow, outputSMEIdx + 1)
      .setBorder(true, true, true, true, true, true)
      .setFontFamily('Roboto')
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("center")
      .setWrap(true)
      .setValue(nameKey);

    //loop through each client for the sme
    for (const [clientKey, clientValue] of Object.entries(nameValue)) {
      const dayIdx = clientKey === 'ST' ? smarthinkingIdx : clientKey === 'BF' ? brainfuseIdx : netTutorIdx;
      const nightIdx = dayIdx + 1;

    //set day hours for the client
    this.outputSheet.getRange(currentRow, dayIdx + 1)
        .setBorder(true, true, true, true, true, true)
        .setFontFamily('Roboto')
        .setVerticalAlignment("middle")
        .setHorizontalAlignment("center")
        .setWrap(true)
        .setValue(clientValue['Day'].toFixed(2));

    //set night hours for the client
    this.outputSheet.getRange(currentRow, nightIdx + 1)
        .setBorder(true, true, true, true, true, true)
        .setFontFamily('Roboto')
        .setVerticalAlignment("middle")
        .setHorizontalAlignment("center")
        .setWrap(true)
        .setValue(clientValue['Night'].toFixed(2));
    }

    // Calculate totals for each row
    this.outputSheet.getRange(currentRow, totalIdx + 1)
      .setBorder(true, true, true, true, true, true)
      .setFontFamily('Roboto')
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("center")
      .setWrap(true)
      .setFormula(`=SUM(${this.numberToLetter(smarthinkingIdx + 1)}${currentRow}, ${this.numberToLetter(brainfuseIdx + 1)}${currentRow}, ${this.numberToLetter(netTutorIdx + 1)}${currentRow})`);

    this.outputSheet.getRange(currentRow, totalIdx + 2)
      .setBorder(true, true, true, true, true, true)
      .setFontFamily('Roboto')
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("center")
      .setWrap(true)
      .setFormula(`=SUM(${this.numberToLetter(smarthinkingIdx + 2)}${currentRow}, ${this.numberToLetter(brainfuseIdx + 2)}${currentRow}, ${this.numberToLetter(netTutorIdx + 2)}${currentRow})`);
  
    }
  }

}