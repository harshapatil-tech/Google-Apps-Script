function finacialYearWiseClientHrs() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = sheet.getSheetByName("Summary");
  const outputSheet = sheet.getSheetByName("FinancialYear_SMEwise_ClientHours");

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
  const financialYearDropdown = outputSheet.getRange('B3').getValue();
  const startYear = financialYearDropdown.split("-")[0];
  const endYear = financialYearDropdown.split("-")[1];
  console.log(startYear, endYear)
  const startMonth = "April"
  const endMonth = "March"
  const subjectDropdown = outputSheet.getRange('I3').getValue();
  
  // Clear Previous entries
  clearRowsContent(outputSheet, 8, 1, 11)
  
  const filteredData = inputData.filter(row => row[yearIdx] == startYear && 
                            getMonthNumber(row[monthIdx]) >= getMonthNumber(startMonth) ||
                            (row[yearIdx] == endYear && 
                              getMonthNumber(row[monthIdx]) <= getMonthNumber(endMonth)
                            ))
                          .filter(row =>{ 
                            if (subjectDropdown === 'All')
                              return true
                            else 
                            return row[subjectIdx] === subjectDropdown
                          });

  const clientNames = setData(inputData, clientIdx)
  const mapObject = {}

  filteredData.forEach(row => { 
    const smeName = row[smeIdx];
    const subject = row[subjectIdx];
    const clientName = row[clientIdx];
    const dayNight = row[dayNightIdx];
    const hours = row[hoursIdx];

    if (!mapObject[smeName]) {
      mapObject[smeName] = {};
    }

    clientNames.forEach(name => {
      if (!mapObject[smeName][name]) {
        mapObject[smeName][name] = {
          "Subject" : subject,
          "Day": 0,
          "Night": 0
        };
      }
    });
    if (!mapObject[smeName][clientName]) {
      mapObject[smeName][clientName] = {
        "Subject": subject,
        "Day": 0,
        "Night": 0
      };
    }
    mapObject[smeName][clientName][dayNight] += hours;
  });

  // console.log(mapObject)
  // Output Sheet Indices
  const outputSheetHeadersTop = outputSheet.getRange(6, 1, 1, outputSheet.getLastColumn()).getValues().flat();
  const outputSheetHeadersBottom = outputSheet.getRange(7, 1, 1, outputSheet.getLastColumn()).getValues().flat();

  const smarthinkingIdx = outputSheetHeadersTop.indexOf("Smarthinking");
  const brainfuseIdx = outputSheetHeadersTop.indexOf("Brainfuse");
  const netTutorIdx = outputSheetHeadersTop.indexOf("NetTutor");
  const outputSMEIdx = outputSheetHeadersTop.indexOf("SME Name");
  const outputSubjectIdx = outputSMEIdx + 1;  // Assuming Subject comes right after SME Name
  const totalIdx = outputSheetHeadersTop.indexOf("Total Hours");
  const srNoIdx = outputSheetHeadersTop.indexOf("Sr. No.");

  let currentRow = 7, srNo = 0;
  for (const [nameKey, nameValue] of Object.entries(mapObject)) {
    currentRow += 1;
    srNo += 1;
    outputSheet.getRange(currentRow, srNoIdx + 1)
      .setBorder(true, true, true, true, true, true)
      .setFontFamily('Roboto')
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("center")
      .setWrap(true)
      .setValue(srNo);
    outputSheet.getRange(currentRow, outputSMEIdx + 1)
      .setBorder(true, true, true, true, true, true)
      .setFontFamily('Roboto')
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("center")
      .setWrap(true)
      .setValue(nameKey);
    outputSheet.getRange(currentRow, outputSubjectIdx + 1)
      .setBorder(true, true, true, true, true, true)
      .setFontFamily('Roboto')
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("center")
      .setWrap(true)
      .setValue(nameValue["ST"]["Subject"]);  // Assuming all subjects are the same for an SME Name

    for (const [clientKey, clientValue] of Object.entries(nameValue)) {
      if (clientKey === 'ST') {
        outputSheet.getRange(currentRow, smarthinkingIdx + 1)
          .setBorder(true, true, true, true, true, true)
          .setFontFamily('Roboto')
          .setVerticalAlignment("middle")
          .setHorizontalAlignment("center")
          .setWrap(true)
          .setValue(clientValue['Day'].toFixed(2));
        outputSheet.getRange(currentRow, smarthinkingIdx + 2)
          .setBorder(true, true, true, true, true, true)
          .setFontFamily('Roboto')
          .setVerticalAlignment("middle")
          .setWrap(true)
          .setHorizontalAlignment("center")
          .setValue(clientValue['Night'].toFixed(2));
      } if (clientKey === 'BF') {
        outputSheet.getRange(currentRow, brainfuseIdx + 1)
          .setBorder(true, true, true, true, true, true)
          .setFontFamily('Roboto')
          .setVerticalAlignment("middle")
          .setHorizontalAlignment("center")
          .setWrap(true)
          .setValue(clientValue['Day'].toFixed(2));
        outputSheet.getRange(currentRow, brainfuseIdx + 2)
          .setBorder(true, true, true, true, true, true)
          .setFontFamily('Roboto')
          .setVerticalAlignment("middle")
          .setHorizontalAlignment("center")
          .setWrap(true)
          .setValue(clientValue['Night'].toFixed(2));
      } if (clientKey === 'NT') {
        outputSheet.getRange(currentRow, netTutorIdx + 1)
          .setBorder(true, true, true, true, true, true)
          .setFontFamily('Roboto')
          .setVerticalAlignment("middle")
          .setHorizontalAlignment("center")
          .setWrap(true)
          .setValue(clientValue['Day'].toFixed(2));
        outputSheet.getRange(currentRow, netTutorIdx + 2)
          .setBorder(true, true, true, true, true, true)
          .setFontFamily('Roboto')
          .setVerticalAlignment("middle")
          .setHorizontalAlignment("center")
          .setWrap(true)
          .setValue(clientValue['Night'].toFixed(2));
      }
    }
    // Calculate totals for each row
    outputSheet.getRange(currentRow, totalIdx + 1)
      .setBorder(true, true, true, true, true, true)
      .setFontFamily('Roboto')
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("center")
      .setWrap(true)
      .setFormula(`=SUM(${numberToLetter(smarthinkingIdx + 1)}${currentRow}, ${numberToLetter(brainfuseIdx + 1)}${currentRow}, ${numberToLetter(netTutorIdx + 1)}${currentRow})`);
    outputSheet.getRange(currentRow, totalIdx + 2)
      .setBorder(true, true, true, true, true, true)
      .setFontFamily('Roboto')
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("center")
      .setWrap(true)
      .setFormula(`=SUM(${numberToLetter(smarthinkingIdx + 2)}${currentRow}, ${numberToLetter(brainfuseIdx + 2)}${currentRow}, ${numberToLetter(netTutorIdx + 2)}${currentRow})`);
  }


}