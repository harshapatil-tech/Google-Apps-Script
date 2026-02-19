// function financialYearWiseClientsHrs() {
//   const fyClientHrs = new FinancialYearWiseClientHrs();
//   fyClientHrs.execute();
// }


// class FinancialYearWiseClientHrs {
//   constructor() {
//     this.sheet = SpreadsheetApp.getActiveSpreadsheet();
//     this.inputSheet = this.sheet.getSheetByName("Summary");
//     this.outputSheet = this.sheet.getSheetByName("FinancialYear_SMEwise_ClientHours");
//   }

//   // this method to fetch headers and data from the input sheet
//   getHeadersAndData() {           
//     const data = this.inputSheet.getRange(1, 1, this.inputSheet.getLastRow(), this.inputSheet.getLastColumn()).getValues();
//     // console.log(data);
//     return {
//       headers: data[0],  //first row header
//       rows: data.slice(1)   //remaining rows as data
//     };
//   }

//   //this method get month number from the month name
//   getMonthNumber(month) {
//     const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
//     return months.indexOf(month) + 1;
//   }

//   //clear previous entries in the output sheet
//   clearPreviousEntries(rangeStartRow, rangeStartCol, numCols) {
//     this.outputSheet.getRange(rangeStartRow, rangeStartCol, this.outputSheet.getLastRow() - rangeStartRow + 1, numCols).clearContent();
//   }

//   //unique values from a specific column
//   setData(data, colIndex) {
//     return [...new Set(data.map(row => row[colIndex]))];
//   }

//   applyCellFormatting(range, value, wrap = true) {
//     range
//       .setBorder(true, true, true, true, true, true)
//       .setFontFamily('Roboto')
//       .setVerticalAlignment("middle")
//       .setHorizontalAlignment("center")
//       .setWrap(wrap)
//       .setValue(value);
//   }

//   //convert column number to column letter (e.g., 1 -> A)
//   numberToLetter(num) {
//     let letter = '';
//     while (num > 0) {
//       num -= 1;
//       letter = String.fromCharCode(65 + (num % 26)) + letter;
//       num = Math.floor(num / 26);
//     }
//     return letter;
//   }

//   populateOutputSheet(mapObject, clientNames, headers) {
//     const outputHeadersTop = this.outputSheet.getRange(6, 1, 1, this.outputSheet.getLastColumn()).getValues().flat();
//     // console.log(outputHeadersTop);
//     const smarthinkingIdx = outputHeadersTop.indexOf("Smarthinking");
//     const brainfuseIdx = outputHeadersTop.indexOf("Brainfuse");
//     const netTutorIdx = outputHeadersTop.indexOf("NetTutor");
//     const outputSMEIdx = outputHeadersTop.indexOf("SME Name");
//     const outputSubjectIdx = outputSMEIdx + 1; // Assuming Subject comes right after SME Name
//     const totalIdx = outputHeadersTop.indexOf("Total Hours");
//     const srNoIdx = outputHeadersTop.indexOf("Sr. No.");

//     let currentRow = 7;
//     let srNo = 0;

//     for (const [nameKey, nameValue] of Object.entries(mapObject)) {
//       currentRow += 1;   //move to next row
//       srNo += 1;      //increment serial number

//       //apply formatting and set values for Sr. No. and SME Name
//       this.applyCellFormatting(this.outputSheet.getRange(currentRow, srNoIdx + 1), srNo);
//       this.applyCellFormatting(this.outputSheet.getRange(currentRow, outputSMEIdx + 1), nameKey);
//       this.applyCellFormatting(this.outputSheet.getRange(currentRow, outputSubjectIdx + 1), nameValue["ST"]["Subject"]);

//       for (const [clientKey, clientValue] of Object.entries(nameValue)) {
//         if (clientKey === 'ST') {
//           this.applyCellFormatting(this.outputSheet.getRange(currentRow, smarthinkingIdx + 1), clientValue['Day'].toFixed(2));
//           this.applyCellFormatting(this.outputSheet.getRange(currentRow, smarthinkingIdx + 2), clientValue['Night'].toFixed(2));
//         } 
//         if (clientKey === 'BF') {
//           this.applyCellFormatting(this.outputSheet.getRange(currentRow, brainfuseIdx + 1), clientValue['Day'].toFixed(2));
//           this.applyCellFormatting(this.outputSheet.getRange(currentRow, brainfuseIdx + 2), clientValue['Night'].toFixed(2));
//         } 
//         if (clientKey === 'NT') {
//           this.applyCellFormatting(this.outputSheet.getRange(currentRow, netTutorIdx + 1), clientValue['Day'].toFixed(2));
//           this.applyCellFormatting(this.outputSheet.getRange(currentRow, netTutorIdx + 2), clientValue['Night'].toFixed(2));
//         }
//       }

//       // Calculate totals for each row (day and night)
//       this.outputSheet.getRange(currentRow, totalIdx + 1)
//         .setFormula(`=SUM(${this.numberToLetter(smarthinkingIdx + 1)}${currentRow}, ${this.numberToLetter(brainfuseIdx + 1)}${currentRow}, ${this.numberToLetter(netTutorIdx + 1)}${currentRow})`)
//         .setBorder(true, true, true, true, true, true)
//         .setFontFamily('Roboto')
//         .setVerticalAlignment("middle")
//         .setHorizontalAlignment("center")
//         .setWrap(true);

//       this.outputSheet.getRange(currentRow, totalIdx + 2)
//         .setFormula(`=SUM(${this.numberToLetter(smarthinkingIdx + 2)}${currentRow}, ${this.numberToLetter(brainfuseIdx + 2)}${currentRow}, ${this.numberToLetter(netTutorIdx + 2)}${currentRow})`)
//         .setBorder(true, true, true, true, true, true)
//         .setFontFamily('Roboto')
//         .setVerticalAlignment("middle")
//         .setHorizontalAlignment("center")
//         .setWrap(true);
//     }
//   }

//   execute() {
//     const { headers, rows } = this.getHeadersAndData();  //get headers and data from input sheet
//     const yearIdx = headers.indexOf("Year");
//     const monthIdx = headers.indexOf("Month");
//     const subjectIdx = headers.indexOf("Subject");
//     const smeIdx = headers.indexOf("SME Name");
//     const clientIdx = headers.indexOf("Client");
//     const dayNightIdx = headers.indexOf("Day/Night");
//     const hoursIdx = headers.indexOf("Hours");

//     //get financial year and subject filter dropdown
//     const financialYearDropdown = this.outputSheet.getRange('B3').getValue();
//     const startYear = financialYearDropdown.split("-")[0];
//     const endYear = financialYearDropdown.split("-")[1];
//     const startMonth = "April";
//     const endMonth = "March";
//     const subjectDropdown = this.outputSheet.getRange('I3').getValue();

//     this.clearPreviousEntries(8, 1, 11);   //clear old data from output sheet

//     //filter rows based on financial year and subject
//     const filteredData = rows.filter(row => 
//       (row[yearIdx] == startYear && this.getMonthNumber(row[monthIdx]) >= this.getMonthNumber(startMonth)) ||
//       (row[yearIdx] == endYear && this.getMonthNumber(row[monthIdx]) <= this.getMonthNumber(endMonth)))
//       .filter(row => subjectDropdown === "All" || row[subjectIdx] === subjectDropdown);

//     const clientNames = this.setData(rows, clientIdx);

//     const mapObject = {};
//     filteredData.forEach(row => {
//       const smeName = row[smeIdx];
//       const subject = row[subjectIdx];
//       const clientName = row[clientIdx];
//       const dayNight = row[dayNightIdx];
//       const hours = row[hoursIdx];

//       if (!mapObject[smeName]) {
//         mapObject[smeName] = {};
//       }

//       clientNames.forEach(name => {
//         if (!mapObject[smeName][name]) {
//           mapObject[smeName][name] = {
//             Subject: subject,
//             Day: 0,
//             Night: 0
//           };
//         }
//       });
//        if (!mapObject[smeName][clientName]) {
//       mapObject[smeName][clientName] = {
//       Subject: subject,
//       Day: 0,
//       Night: 0
//       };
//     }

      
//       //increment hours for day and night based on client
//       mapObject[smeName][clientName][dayNight] += hours;
//     });

//     this.populateOutputSheet(mapObject, clientNames, headers);
//   }
// }

