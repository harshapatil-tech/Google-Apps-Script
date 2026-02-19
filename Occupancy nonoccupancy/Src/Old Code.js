//old code
// function settingBrainFuse_Occupancy_non_Occupanacy() {
//   // Retrieve data from brainfuse using the provided id
//  // const spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); //.openById("1XH5eGI4thuz-yAsHDeVbI90KFX2AWee3w556yUcGe6I");
//  const spreadsheet=SpreadsheetApp.openById("1XF3AppF7FpItTH_pAcQSntr61hOVfBXKe_Q4mClVNzs");
//  let dataToCopy = brainfuseAccountWise(spreadsheet, "Single").flat();
// // let dataToCopy = brainfuseAccountWise(spreadsheet);
//  //let dataToCopy = brainfuseAccountWiseFromSummary(spreadsheet).flat();

//   // Access the "MasterData" sheet in the spreadsheet 1WTAwoyzkLtzKOEPqejOVNs7_3xzRMBEHKTumlNwMd-M 
//   //const ss = SpreadsheetApp.openById("1WTAwoyzkLtzKOEPqejOVNs7_3xzRMBEHKTumlNwMd-M").getSheetByName("Master_Data");
//   const ss = SpreadsheetApp.openById("1O_NKq3y2u0bMolJ52iEWKzD7_jbyIXR3UIrHsxkuEkc").getSheetByName("Master_Data");
  
//   const dataRange = ss.getRange(1, 1, ss.getLastRow(), ss.getLastColumn()).getDisplayValues();
//   let headers = dataRange[0], data = dataRange.slice(1);

//   // Find the indices of relevant columns in the headers
//   const srNoIdx = headers.indexOf("Invoice No");
//   const finYearIdx = headers.indexOf("Financial Year");
//   const semesterIdx = headers.indexOf("Semester");
//   const monthIdx = headers.indexOf("Month");
//   const fromIdx = headers.indexOf("From");
//   const toIdx = headers.indexOf("To");
//   const subIdx = headers.indexOf("Subject");
//   const accountTypeIdx = headers.indexOf("Single/Multiple");
//   const accountIdx = headers.indexOf("Account No");
//   const occupancyTypeIdx = headers.indexOf("Occupancy");
//   const numHrsIdx = headers.indexOf("Hours");
//   const totalHrsIdx = headers.indexOf("Total Hours");

//   const firstDates = data.map(r => r[fromIdx]);
//   const lastDates = data.map(r => r[toIdx]);
  
//   dataToCopy = sortDataByEarliestMonth(dataToCopy);

//   let objList = []; // To store objects that don't have a match in the data

//   let foundMatch = false; // Flag to check if a match is found while iterating

//   let lastRowIndex = ss.getLastRow();
//   console.log(lastRowIndex);
//   let lastSerialNum = ss.getRange(lastRowIndex, srNoIdx+1).getValue();
  
//   let isDeletionIndexFound = false
//   // Iterate over each object in dataToCopy
//   let rowDeletionStart = ss.getLastRow()
  
//   for (const [monthKey, monthValuesObject] of Object.entries(dataToCopy)) {
//     for (const object of monthValuesObject) {
//       let totalHours = 0
//       for (const [key, val] of Object.entries(object)) {
//         if (!isDeletionIndexFound && firstDates.includes(val["firstDay"]) && lastDates.includes(val["lastDay"])) {
//           rowDeletionStart = firstDates.indexOf(val["firstDay"]) + 2
//           const rowsToDelete = ss.getLastRow() - rowDeletionStart + 1;
//           console.log(`Deleting rows from ${val["firstDay"]} to ${val["lastDay"]} (rows ${rowDeletionStart} to ${ss.getLastRow()})`);
//           ss.deleteRows(rowDeletionStart, rowsToDelete)
//           isDeletionIndexFound = true;
//           lastRowIndex = rowDeletionStart - 1;
//         }
//         const hours = val["hours"]
//         totalHours += hours;
//         console.log(`Inserting row for Account:- From: ${val["firstDay"]}, To: ${val["lastDay"]}`);
//         applyCustomFormatting(ss.getRange(lastRowIndex+1, srNoIdx+1)).setValue(lastSerialNum + 1);
      
//         applyCustomFormatting(ss.getRange(lastRowIndex+1, finYearIdx+1)).setValue(val["finYear"]);
//         applyCustomFormatting(ss.getRange(lastRowIndex+1, semesterIdx+1)).setValue(getSeason(val["month"]));
//         applyCustomFormatting(ss.getRange(lastRowIndex+1, monthIdx+1)).setValue(val["month"]);
//         applyCustomFormatting(dateValidation(ss.getRange(lastRowIndex+1, fromIdx+1))).setValue(val["firstDay"]);
//         applyCustomFormatting(dateValidation(ss.getRange(lastRowIndex+1, toIdx+1))).setValue(val["lastDay"]);
//         applyCustomFormatting(ss.getRange(lastRowIndex+1, subIdx+1)).setValue(val["subject"]);
//         if (val["accountType"] == "Single")
//           applyCustomFormatting(ss.getRange(lastRowIndex+1, accountTypeIdx+1), {"bgColor": "#93c47d"}).setValue("Single");
//         else
//           applyCustomFormatting(ss.getRange(lastRowIndex+1, accountTypeIdx+1), {"bgColor":"#e380e3"}).setValue("Multiple");
//         applyCustomFormatting(ss.getRange(lastRowIndex+1, accountIdx+1), {"fontWeight":"bold"})
//                               .setValue(extractNumberFromString(val["account"]));  
//         applyCustomFormatting(ss.getRange(lastRowIndex+1, numHrsIdx+1)).setValue(hours);
//         if (key === "Occupancy")
//         applyCustomFormatting(ss.getRange(lastRowIndex+1, occupancyTypeIdx+1)).setValue("Occupancy");
//         else if (key === "Non-Occupancy")
//         applyCustomFormatting(ss.getRange(lastRowIndex+1, occupancyTypeIdx+1)).setValue("Non-Occupancy");
//         lastRowIndex += 1;
//         applyCustomFormatting(ss.getRange(lastRowIndex, totalHrsIdx+1)); 
//       }
//        applyCustomFormatting(ss.getRange(lastRowIndex, totalHrsIdx+1), {"bgColor":"#fbbc04", "fontWeight":"bold"}).setValue(totalHours);
//     }
//   }

// }



// function copyData() {

//   const SPREADSHEET = SpreadsheetApp.openById("1azmcGWS2os6jdsQXN1bOS6euctPi9DGuGuKbnFrn4UQ");
//   const archiveSheet = SPREADSHEET.getSheetByName("Brainfuse Timesheet Archive");

//   const scraperSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Summary");

//   const [scraperIndices, scraperData] = get_Data_Indices_From_Sheet(scraperSheet);
//   const [archiveIndices, archiveData] = get_Data_Indices_From_Sheet(archiveSheet);

//   let lastRowIndex = archiveSheet.getLastRow();

//   const scriptProperties = PropertiesService.getScriptProperties();
//   const lastExecutionDate = scriptProperties.getProperty('lastExecutionDate');

//   if (lastExecutionDate) {
//     const daysSinceLastExecution = Math.floor((new Date() - new Date(lastExecutionDate)) / (1000 * 60 * 60 * 24));

//     if (daysSinceLastExecution <= 1) {
//       SpreadsheetApp.getUi().alert("Function already executed in the last 14 days.");
//       return;
//     } else {

//       scraperData.forEach(function (scraperRow) {

//         archiveSheet.getRange(lastRowIndex+1, archiveIndices["Department"]+1).setValue(scraperRow[scraperIndices["Department"]]);
//         archiveSheet.getRange(lastRowIndex+1, archiveIndices["Account No."]+1).setValue(scraperRow[scraperIndices["Account No."]]);
//         archiveSheet.getRange(lastRowIndex+1, archiveIndices["Type"]+1).setValue(scraperRow[scraperIndices["Type"]]);
//         archiveSheet.getRange(lastRowIndex+1, archiveIndices["Activity Type"]+1).setValue(scraperRow[scraperIndices["Activity Type"]]);
//         archiveSheet.getRange(lastRowIndex+1, archiveIndices["Start Date"]+1).setValue(scraperRow[scraperIndices["Start Date"]]);
//         archiveSheet.getRange(lastRowIndex+1, archiveIndices["Start Time"]+1).setValue(scraperRow[scraperIndices["Start Time"]]);
//         archiveSheet.getRange(lastRowIndex+1, archiveIndices["Hours"]+1).setValue(scraperRow[scraperIndices["Hours"]]);
//         archiveSheet.getRange(lastRowIndex+1, archiveIndices["Comments"]+1).setValue(scraperRow[scraperIndices["Comments"]]);

//         lastRowIndex += 1;
//       });
//     }

//   }

//   scriptProperties.setProperty('lastExecutionDate', new Date().toISOString());

// }










// =======================================OLD CODE==============================================
// function settingbfOccupancy_NonOccupancy() {
// //   const masterSheet = SpreadsheetApp.openById('1WTAwoyzkLtzKOEPqejOVNs7_3xzRMBEHKTumlNwMd-M').getSheetByName("Master_Data");
// // //'1rzmKo_CkPp3_TdYNRsTRnnf0RB8IANwcrSnf2PDU8Bc'
// const masterSheet = SpreadsheetApp.openById('1spLkZValt_lt9HUS2QDKje9cTPEXChl39vbdeCNvMaQ').getSheetByName("Master_Data");
// //'1rzmKo_CkPp3_TdYNRsTRnnf0RB8IANwcrSnf2PDU8Bc'
//  const [dataObject,fromDate, toDate] = getBFOccupancyData();
//   const formattedFrom = formatDate(fromDate);
//   const formattedTo = formatDate(toDate);

//   const fullMonthName = fromDate.toLocaleDateString("en-US", { month: 'long' });
//   const year = fromDate.getFullYear();
//   const financialYear = getFinancialYear(fullMonthName, year);

//   const invoiceNo = getNextInvoiceNo(masterSheet);
//   const semester = getSemester(fromDate);
//   const monthName = fullMonthName;

//   // Logger.log("formatted date" + formattedFrom);
//   // Logger.log("To date" + formattedTo);
//   // Logger.log("month name" + monthName);
//   // Logger.log("Financial year" + financialYear);
//   // Logger.log("Invoice number" + invoiceNo);
//   // Logger.log("semester" + semester);

//   deleteOldRange(masterSheet, formattedFrom, formattedTo);

//   const output = [];
//   const hourGroup = {};

//   for (let department in dataObject) {
//     for (let singleDual in dataObject[department]) {

//       let groupKey = `${department}_${singleDual}`;
//       let totalHours = 0;

//       for (let accountNum in dataObject[department][singleDual]) {

//         const entry = dataObject[department][singleDual][accountNum];
//         totalHours += (entry["Occupancy"] || 0) + (entry["Non-Occupancy"] || 0);

//       }
//       hourGroup[groupKey] = totalHours;
//     }
//   }
//   for (let department in dataObject) {
//     for (let singleDual in dataObject[department]) {
//       const groupKey = `${department}_${singleDual}`;
//       const entries = Object.entries(dataObject[department][singleDual]);
//       const lastIndex = entries.length - 1;

//       entries.forEach(([accountNum, entry], index) => {
//         const occupancy = entry["Occupancy"] || 0;
//         const nonOccupancy = entry["Non-Occupancy"] || 0;

//         if (occupancy) {
//           output.push([
//             invoiceNo,
//             financialYear,
//             semester,
//             monthName,
//             formattedFrom,
//             formattedTo,
//             department,
//             singleDual,
//             accountNum,
//             "Occupancy",
//             occupancy,
//             null
//           ]);
//         }

//         if (entry) {
//           const isLast = index === lastIndex;
//           output.push([
//             invoiceNo,
//             financialYear,
//             semester,
//             monthName,
//             formattedFrom,
//             formattedTo,
//             department,
//             singleDual,
//             accountNum,
//             "Non-Occupancy",
//             nonOccupancy,
//             isLast ? hourGroup[groupKey] : null
//           ]);
//         }
//       });
//     }
//   }

//   if (output.length > 0) {

//     const startRow = masterSheet.getLastRow() + 1;
//     const range = masterSheet.getRange(startRow, 1, output.length, output[0].length);
//     range.setValues(output);

//     for (let i = 0; i < output.length; i++) {
//       const singleDual = output[i][7];
//       const isTotal = output[i][11] !== null;

//       if (singleDual === 'Single') {
//         masterSheet.getRange(startRow + i, 8).setBackground('#93c47d');
//       } else if (singleDual === 'Multiple') {
//         masterSheet.getRange(startRow + i, 8).setBackground('#e380e3');
//       }


//       if (isTotal) {
//         masterSheet.getRange(startRow + i, 12).setBackground('#fbbc04');
//       }
//     }
//     //Logger.log("Appending " + output.length + " rows to Master_Data");
//   }
//   else {
//     //Logger.log("No data to append.");
//   }
// }

// function deleteOldRange(sheet, from, to) {
//   const data = sheet.getDataRange().getValues();
//   const header = data[0];
//   const fromIndex = header.indexOf("From");
//   const toIndex = header.indexOf("To");
//   //Logger.log("Searching and deleting: " + from + " and To: " + to);

//   for (let i = data.length - 1; i > 0; i--) {
//     const rowFrom = formatDate(new Date(data[i][fromIndex]));
//     const rowTo = formatDate(new Date(data[i][toIndex]));

//     //Logger.log("Checking row " + (i + 1) + ": rowFrom = " + rowFrom + ", rowTo = " + rowTo);
//     if (rowFrom === from && rowTo === to) {
//       sheet.deleteRow(i + 1);
//     }
//   }
// }


// function getNextInvoiceNo(sheet) {

//   const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
//   const nums = data.filter(n => typeof n === "number");
//   //Logger.log("Existing invoice numbers: " + nums);

//   const nextNo = nums.length > 0 ? Math.max(...nums) + 1 : 1;
//   //Logger.log("Next Number: " + nextNo);
//   return nextNo;
// }

// function getSemester(date) {
//   const month = date.getMonth() + 1;
//   let semester;
//   if (month >= 1 && month <= 5) semester = "Spring";
//   else if (month >= 6 && month <= 8) semester = "Summer";
//   else semester = "Fall";

//   Logger.log("Month: " + month + ", Semester: " + semester);
//   return semester;
// }


// ***************************************************/
// function getNextDay(sheet) {
//   const data = sheet.getDataRange().getValues();
//   const header = data[0];
//   const toIndex = header.indexOf("To");
  

//   console.log("Headers"+header);
//   console.log("Index"+toIndex);

//   let latestToDate = null;
//   for (let i = data.length - 1; i > 0; i--) {
//     const toCell = data[i][toIndex];
//     if (toCell) {
//       latestToDate = new Date(toCell);
//       console.log(latestToDate);
//       break;
//     }
//   }

//   if (!latestToDate) throw new Error("No valid 'To' date found.");

//   let from = new Date(latestToDate);
//   from.setDate(from.getDate() + 1); // next day
  
//   console.log("from date"+from);

//   let to = new Date(from);
//   to.setDate(from.getDate() + 13); // 14-day window

//   console.log("To date"+to);

//   const endOfMonth = new Date(from.getFullYear(), from.getMonth() + 1, 0);
//   console.log("end of month"+endOfMonth);

//   if (to > endOfMonth) {
//     const firstWindowFrom = new Date(from);
//     const firstWindowTo = new Date(endOfMonth);
    
    

//     const daysUsed = Math.floor((formattedFrom - formattedTo) / (1000 * 60 * 60 * 24)) + 1;
//     const leftoverDays = 14 - daysUsed;

//     const secondWindowFrom = new Date(toDate);
//     secondWindowFrom.setDate(secondWindowFrom.getDate() + 1);

//     const secondWindowTo = new Date(fromDate);
//     secondWindowTo.setDate(secondWindowFrom.getDate() + leftoverDays - 1);
    

//     if (leftoverDays === 1) {
//       secondWindowTo.setDate(secondWindowFrom.getDate()); // same as from
//     }

//     return [
//   firstWindowFrom,   // 19-May-2025
//   firstWindowTo,     // 31-May-2025
//   secondWindowFrom,  // 01-Jun-2025
//   secondWindowTo     // 01-Jun-2025
// ];

//   }

//   return [from, to];
// }



//===========================================================================================
//old code
// function getBFOccupancyData() {
//   const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
//   const summarySheet = spreadSheet.getSheetByName("Summary")
//   const [headerRowMap, data] = get_Data_Indices_From_Sheet(summarySheet);
  
//   const leastDate = new Date(Math.min.apply(null, data.map(r => new Date(r[headerRowMap["Start Date"]]))));
//   console.log("least date",leastDate);
//   const maxDate = new Date(Math.max.apply(null, data.map(r => new Date(r[headerRowMap["Start Date"]]))));
//   console.log("maxdate:-",maxDate);
  
//   const dataObject = {};
  
//   data.forEach(singleRow => {

//       //const accountNum = singleRow[headerRowMap["Account No."]];
//      let accountNum = singleRow[headerRowMap["Account No."]];
//      accountNum = accountNum.toString().replace(/\D/g, ''); 


//       let singleDual = singleRow[headerRowMap["Type"]];
     
//       let department = singleRow[headerRowMap["Department"]];
     
//       if (department === "Mathematics") {
//         department = "Calculus"
//       }
//       if (department === "Intro Accounting") {
//         department = "Accounting"
//       }
//       let occupancyType = singleRow[headerRowMap["Activity Type"]];
//       const hours = singleRow[headerRowMap["Hours"]];

//       if (occupancyType === 'IA-Waited' || occupancyType === "IA-Tutored") {

//         if (singleDual === "Dual")
//           singleDual = "Multiple"

//         if (occupancyType === "IA-Tutored")
//           occupancyType = "Occupancy"

//         if (occupancyType === "IA-Waited")
//           occupancyType = "Non-Occupancy"

//         if (!dataObject.hasOwnProperty(department)) {

//           dataObject[department] = {}
//         }
//         if (!dataObject[department].hasOwnProperty(singleDual)) {
//           dataObject[department][singleDual] = {}
//         }

//         if (!dataObject[department][singleDual].hasOwnProperty(accountNum)) {
//           dataObject[department][singleDual][accountNum] = {}
//         }

//         if (!dataObject[department][singleDual][accountNum].hasOwnProperty(occupancyType)) {
//           dataObject[department][singleDual][accountNum][occupancyType] = 0
//         }
        
//         dataObject[department][singleDual][accountNum][occupancyType] += hours;
//       }
//     });
    
//      console.log("Final Data Object:", JSON.stringify(dataObject, null, 2));

//     return [dataObject, leastDate, maxDate];
// }


//old code 

// function brainfuseAccountWise(spreadSheet) {

//   const sheets = spreadSheet.getSheets();
//   const array = [];
//   // const sheet = sheets[5]
//   let ss;
//   for (const sheet of sheets) {
//       const sheetName = sheet.getName();
//       const innerArray = [];
      
//       if (
//           sheetName.trim() === 'Calculus' || sheetName.trim() === "Statistics" ||
//           sheetName.trim() === 'English' || sheetName.trim() === "Chemistry" ||
//           sheetName.trim() === 'Physics' || sheetName.trim() === "Biology" ||
//           sheetName.trim() === 'Finance' || sheetName.trim() === "Economics" ||
//           sheetName.trim() === "Intro Accounting" || sheetName.trim() === 'Computer Science'
//           ) {
//         // if (sheetName.trim() === 'Calculus'){
//         ss = spreadSheet.getSheetByName(sheetName);
//        // console.log(sheetName)
//         if (ss.getLastRow() === 0){
//           continue;
//         }else {
//           const dateColumn = ss.getRange(1, 1, ss.getLastRow(), 1).getValues().flat();
        
//           const startRowSingleSubject = dateColumn.indexOf("Date") + 1;
//           const totalHrsRowSingleSubject = dateColumn.indexOf("Total") + 1;

//           let headers = ss.getRange(startRowSingleSubject, 1, 1, ss.getLastColumn()).getValues().flat();
//           let totalColumn = headers.indexOf("Total") + 1;

//           innerArray.push(getAccountData(ss, sheetName, headers, startRowSingleSubject, totalHrsRowSingleSubject, totalColumn, dateColumn, "Single"));

//           const startRowMultipleSubject = dateColumn.lastIndexOf("Date") + 1;
//           const totalHrsRowMultipleSubject = dateColumn.lastIndexOf("Total") + 1;
//           headers = ss.getRange(startRowMultipleSubject, 1, 1, ss.getLastColumn()).getValues().flat();
//           totalColumn = headers.indexOf("Total") + 1;

//           if (startRowSingleSubject !== startRowMultipleSubject){
//             var data = getAccountData (ss, sheetName, headers, startRowMultipleSubject, totalHrsRowMultipleSubject, totalColumn, dateColumn, "Multiple");
//             innerArray.push(data)
//           }
//         }
//         array.push(...innerArray);
//       }
      
//   };

//   return array;
// }


// function getAccountData (sheet, sheetName, headers, startRow, totalHrsRow, totalColumn, dateColumn, accountType) {

//   const array = [];
  
//   for (let i = 2; i<totalColumn; i+=3) {
  
//       const valuesSingleSub = sheet.getRange(startRow+1, i, totalHrsRow-startRow, 2)
//                               .getValues();
  
//       const datesSingleSub = dateColumn.slice(startRow, totalHrsRow - 1);
   
//       const updatedValuesSingleSub = valuesSingleSub.map((element, index) => {
//         if (index === 0)
//           return [dateColumn[1], headers[i-1], ...element];
          
//         else
//           return [datesSingleSub[index], headers[i-1], ...element];
//       }).filter(element => element[0] != "" && element[0] != undefined);
      
//       result = {};
//       const keys = ["subject", "account", "firstDay", "lastDay", "month", "year", "finYear"];

//       updatedValuesSingleSub.slice(1).forEach(([dateObj, account, ...values]) => {
//         const modifiedDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd-MMMM-YYYY");
//         const modifiedDateShort = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd-MMM-YYYY")
//         const [date, month, year] = modifiedDate.split("-");
//         if (!result.hasOwnProperty(month)){
//           result[month] = {}
//           if (!result[month].hasOwnProperty("Occupancy")) {
//             result[month]["Occupancy"] = {};
//           }
//           if(!result[month].hasOwnProperty("Non-Occupancy")){
//             result[month]["Non-Occupancy"] = {}
//           }
//           result[month]["Occupancy"].hours = Number(values[0]);
//           result[month]["Non-Occupancy"].hours = Number(values[1]);
//           result[month]["Occupancy"].accountType = accountType;
//           result[month]["Non-Occupancy"].accountType = accountType;
//           keys.forEach((key, index) => {
//           if (key === "firstDay"){
//             result[month]["Occupancy"].firstDay = modifiedDateShort;
//             result[month]["Non-Occupancy"].firstDay = modifiedDateShort;
//           }
//           else if (key === "lastDay") {
//             result[month]["Occupancy"].lastDay = modifiedDateShort;
//             result[month]["Non-Occupancy"].lastDay = modifiedDateShort;
//           }
//           else {
//             // result[month]["Occupancy"][key] = 0;
//             // result[month]["Non-Occupancy"][key] = 0
//             // if (sheetName === "Calculus")
//             //   result[month].subject = "Mathematics";
//             // else{

//             result[month]["Occupancy"].subject = sheetName;
//             result[month]["Non-Occupancy"].subject = sheetName;
//             result[month]["Occupancy"].finYear = getFinancialYear(month, year);
//             result[month]["Occupancy"].month = month;
//             result[month]["Occupancy"].account = account;
//             result[month]["Occupancy"].accountType = accountType;
//             result[month]["Non-Occupancy"].finYear = getFinancialYear(month, year);
//             result[month]["Non-Occupancy"].month = month;
//             result[month]["Non-Occupancy"].account = account;
//             result[month]["Non-Occupancy"].accountType = accountType;
//           }
//       });
//         }else {
//           result[month]["Occupancy"].hours += Number(values[0]);
//           result[month]["Non-Occupancy"].hours += Number(values[1]);

//           result[month]["Occupancy"].lastDay = modifiedDateShort;
//           result[month]["Non-Occupancy"].lastDay = modifiedDate;
//         }

      
//       })
//     array.push(result);
//     }
//     console.log("result:-",result);
//     return array;
// }



//=============================================================
// function settingBFOccupancyNonOccupanacy(){
//   const [dataToCopy, minDate, maxDate] = getBFOccupancyData();
//   const spreadSheet = SpreadsheetApp.openById("1WTAwoyzkLtzKOEPqejOVNs7_3xzRMBEHKTumlNwMd-M")
 
//   const masterSheet = spreadSheet.getSheetByName("Master_Data");
//   const [outputHeaderIndices, data] = get_Data_Indices_From_Sheet(masterSheet);

//   const formattedMinDate = Utilities.formatDate(minDate, Session.getScriptTimeZone(), "dd-MMM-yyyy");
//   const formattedMaxDate = Utilities.formatDate(maxDate, Session.getScriptTimeZone(), "dd-MMM-yyyy");

//   let startingRowsToDelete = undefined;

//   const [firstDates, lastDates] = [
//     data.map(r => new Date(r[outputHeaderIndices["From"]])),
//     data.map(r => new Date(r[outputHeaderIndices["To"]]))
//   ];

//   const [startDateMap, endDateMap] = [new Map(), new Map()];
//   firstDates.forEach((date, index) =>{
//     const dateInMillis = date.getTime();
//     if(!startDateMap.has(dateInMillis))
//       startDateMap.set(dateInMillis, index)
//     });
//   lastDates.forEach((date, index) => {
//     const dateInMillis = date.getTime();
//     if(!endDateMap.has(dateInMillis))
//       endDateMap.set(date.getTime(), index)
//     });

  
//   if ((startDateMap.has(minDate.getTime()) && endDateMap.has(maxDate.getTime())) && 
//   (startDateMap.get(minDate.getTime()) == endDateMap.get(maxDate.getTime()))) {
//     startingRowsToDelete = startDateMap.get(minDate.getTime()) + 2;
//   }

//   if (startingRowsToDelete !== undefined) {
//     masterSheet.getRange(startingRowsToDelete, 1, masterSheet.getLastRow(), masterSheet.getLastColumn()).clear();
//   }

//   const modifiedDate = Utilities.formatDate(maxDate, Session.getScriptTimeZone(), "dd-MMMM-yyyy");
//   const [_, month, year] = modifiedDate.split("-")
//   const finYear = getFinancialYear(month, year);

//   let lastRow = masterSheet.getLastRow();
//   const lastInvoiceNum = masterSheet.getRange(lastRow, outputHeaderIndices["Invoice No"]+1).getValue()
  
//   for (const [subject, subjectValues] of Object.entries(dataToCopy)) {

//     for(let [singleDual, singleDualValues] of Object.entries(subjectValues)) {
//       let totalHours = 0;
//       for (const [accountNum, occupancyTypeValues] of Object.entries(singleDualValues)) {
        
//         for (let [occupancyType, hours] of Object.entries(occupancyTypeValues)){
//           totalHours += hours;
          
//           applyCustomFormatting(masterSheet.getRange(lastRow + 1, outputHeaderIndices["Invoice No"]+1).setValue(lastInvoiceNum+1));
//           applyCustomFormatting(masterSheet.getRange(lastRow + 1, outputHeaderIndices["Financial Year"]+1).setValue(finYear));
//           applyCustomFormatting(masterSheet.getRange(lastRow + 1, outputHeaderIndices["Semester"]+1).setValue(getSeason(month)));
//           applyCustomFormatting(masterSheet.getRange(lastRow + 1, outputHeaderIndices["Month"]+1).setValue(month));
//           applyCustomFormatting(dateValidation.call(this, masterSheet.getRange(lastRow + 1, outputHeaderIndices["From"]+1)))
//                   .setValue(formattedMinDate).setNumberFormat("dd-MMM-YYYY")
//           applyCustomFormatting(dateValidation.call(this, masterSheet.getRange(lastRow + 1, outputHeaderIndices["To"]+1)))
//                   .setValue(formattedMaxDate).setNumberFormat("dd-MMM-YYYY");
//           applyCustomFormatting(masterSheet.getRange(lastRow + 1, outputHeaderIndices["Subject"]+1)).setValue(subject)
//           if (singleDual === "Single")
//             applyCustomFormatting(masterSheet.getRange(lastRow + 1, outputHeaderIndices["Single/Multiple"]+1), {"bgColor": "#93c47d"})
//                       .setValue("Single");
//           if (singleDual === "Multiple")
//             applyCustomFormatting(masterSheet.getRange(lastRow + 1, outputHeaderIndices["Single/Multiple"]+1), {"bgColor": "#e380e3"})
//                       .setValue("Multiple");
//           applyCustomFormatting(masterSheet.getRange(lastRow + 1, outputHeaderIndices["Account No"]+1), {"fontWeight":"bold"})
//                       .setValue(extractNumberFromString(accountNum));
//           applyCustomFormatting(masterSheet.getRange(lastRow + 1, outputHeaderIndices["Occupancy"]+1).setValue(occupancyType));
//           applyCustomFormatting(masterSheet.getRange(lastRow + 1, outputHeaderIndices["Hours"]+1).setValue(hours));

//           lastRow += 1;
//           applyCustomFormatting(masterSheet.getRange(lastRow, outputHeaderIndices["Total Hours"]+1));
//         }
  
//       }
//       applyCustomFormatting(masterSheet.getRange(lastRow, outputHeaderIndices["Total Hours"]+1), 
//                               {"bgColor":"#fbbc04", "fontWeight":"bold"}).setValue(totalHours);
//     }
//   }
// }


