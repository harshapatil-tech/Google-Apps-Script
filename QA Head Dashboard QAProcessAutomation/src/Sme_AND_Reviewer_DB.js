function reviewerDBSheet() {
  const spreadSheet = SpreadsheetApp.openById("1u4i5rMCWM0mSgDWdf6EmGVi40kb9i2Qx51SE1cupIjU")
  const outputSheet = spreadSheet.getSheetByName('Reviewer DB');
  const outputHeaders = outputSheet.getRange(1, 1, 1, outputSheet.getLastColumn()).getValues().flat();
  const outputSrNoIdx = outputHeaders.indexOf("#") + 1;
  const outputEmailIdx = outputHeaders.indexOf("Email ID") + 1;
  const outputReviewerIdx = outputHeaders.indexOf("Reviewer Name") + 1;
  const outputDepartmentIdx = outputHeaders.indexOf("Department") + 1;
  const outputActiveIdx = outputHeaders.indexOf("Active?") + 1;

  const inputSheet = SpreadsheetApp.openById('11b5A_ZDi9DskTiKJ0j4fBDkRbY-0I2h7rVRxr6KVK1o').getSheetByName('Details');


  const anotherInputSheet = spreadSheet.getSheetByName('Copy of QA Team')
  let deptQATable = getTableData(anotherInputSheet, 1, anotherInputSheet.getLastRow() - 2, 8)
  deptQATable = inverseMapping(deptQATable);
  const masterData = inputSheet.getRange(1, 1, inputSheet.getLastRow(), inputSheet.getLastColumn()).getValues().filter(row => row.some(Boolean));
  const headers = masterData[0];
  let data = masterData.slice(1);
  const empNameIdx = headers.indexOf('Employee Name');
  const deptIdx = headers.indexOf('Department');
  const officialEmailIdx = headers.indexOf('Official email ID');

  const emailEmployeeMap = {}
  const departmentEmployeeMap = {}
  for (const row of data) {
    const employee = row[empNameIdx]
    const email = row[officialEmailIdx];
    emailEmployeeMap[employee] = email;
    departmentEmployeeMap[employee] = row[deptIdx]
  }

  let rowIndex = 2, srNo = 1;
  for (const [reviewer, department] of Object.entries(deptQATable)) {
    if (reviewer in emailEmployeeMap) {
      const email = emailEmployeeMap[reviewer];
      outputSheet.getRange(rowIndex, outputSrNoIdx).applyCustomFormatting().setValue(srNo);
      outputSheet.getRange(rowIndex, outputEmailIdx).applyCustomFormatting().setValue(email);
      outputSheet.getRange(rowIndex, outputReviewerIdx).applyCustomFormatting().setValue(reviewer);
      outputSheet.getRange(rowIndex, outputDepartmentIdx).applyCustomFormatting().setValue(department);
      outputSheet.getRange(rowIndex, outputActiveIdx).applyCustomFormatting().insertCheckboxes().setValue(false);
      rowIndex++, srNo++;

    }
  }
}


function departmentMapping() {
  const spreadsheet = SpreadsheetApp.openById("1u4i5rMCWM0mSgDWdf6EmGVi40kb9i2Qx51SE1cupIjU");
  const inputsheet = spreadsheet.getSheetByName("Backend - Dept Mapping");
  const inputDataRange = inputsheet.getRange(1, 1, inputsheet.getLastRow(), inputsheet.getLastColumn()).getValues();
  const inputHeaders = inputDataRange[0], inputData = inputDataRange.slice(1);

  const inputIndices = {
    hrDepartmentIndex: inputHeaders.indexOf("HR Department"),
    qaDepartmentIndex: inputHeaders.indexOf("QA Department"),
  }

  const departmentMap = {}
  inputData.forEach(r => {
    if (!departmentMap.hasOwnProperty(r[inputIndices.hrDepartmentIndex])) {
      departmentMap[r[inputIndices.hrDepartmentIndex]] = r[inputIndices.qaDepartmentIndex]
    }
  })

  return departmentMap;
}





// function smeDBSheet() {
//   const inputSheet = SpreadsheetApp.openById("11b5A_ZDi9DskTiKJ0j4fBDkRbY-0I2h7rVRxr6KVK1o").getSheetByName('Details');
//   const spreadSheet = SpreadsheetApp.openById("1eJ-7nKeICwYXEBh4oEzLwC6SpkYzCw5ntYxarbD2jxA");
//   const outputSheet = spreadSheet.getSheetByName("SME DB");

//   const masterData = inputSheet.getRange(1, 1, inputSheet.getLastRow(), inputSheet.getLastColumn()).getValues().filter(row => row.some(Boolean));
//   const headers = masterData[0];
//   let data = masterData.slice(1);

//   const hrDepartmentData = departmentMapping();

//   const empNameIdx = headers.indexOf('Employee Name');
//   const reportManagerIdx = headers.indexOf('Reporting Manager');
//   const pyramidCategoryIdx = headers.indexOf('Pyramid Category');
//   const deptIdx = headers.indexOf('Department');
//   const officialEmailIdx = headers.indexOf('Official email ID');


//   // CHECK
//   const resultArray = [];

//   const outputRange = outputSheet.getRange(1, 1, outputSheet.getLastRow(), outputSheet.getLastColumn()).getValues();
//   const outputHeaders = outputRange[0], outputData = outputRange.slice(1);

//   const outputEmailIdIdx = outputHeaders.indexOf('Email ID');
//   const outputDepartmentIdx = outputHeaders.indexOf('Department');
//   const outputPyramidCatIdx = outputHeaders.indexOf('Pyramid Category');
//   const outputReportMangIdx = outputHeaders.indexOf('Reporting Manager');
//   const outputActiveIdx = outputHeaders.indexOf('Active?');

//   let index = 1;
//   let lastSrNo = outputSheet.getRange(outputSheet.getLastRow(), 1).getValue(); 

//   data.forEach(r => {

//     if(hrDepartmentData.hasOwnProperty(r[deptIdx]) && r[officialEmailIdx] !== '-'){

//       const department = hrDepartmentData[r[deptIdx]];

//       if ( r[empNameIdx] !== 'Tushar Jangale' ) {

//         resultArray.push([index++, r[officialEmailIdx], department, r[empNameIdx], "", "", r[pyramidCategoryIdx], r[reportManagerIdx]]);
//       }

//     }
//     if(r[empNameIdx] === 'Tejas Jagtap'){

//       resultArray.push([index++, r[officialEmailIdx], 'Others', r[empNameIdx], "", "", r[pyramidCategoryIdx], r[reportManagerIdx]]);

//     }
//   })


//   const updatedOutputData = [];
//   resultArray.forEach((row, index) => {
//       const email = row[1];
//       const resultDepartment = row[2];
//       const resultPyramidCategory = row[6];
//       const resultReportingManager = row[7]
//       let found = false;

//       // Find the row in outputData with this email
//       for (let r of outputData) {

//           if (r[outputEmailIdIdx] === email) {

//               found = true;
//               // Check if department has changed
//               if (r[outputDepartmentIdx] !== resultDepartment) {

//                   console.log(`Department change for ${email}: ${r[outputDepartmentIdx]} to ${resultDepartment}`);
//                   outputSheet.getRange(outputData.indexOf(r) + 2, outputDepartmentIdx + 1).setValue(resultDepartment);

//               }

//               // Check if pyramidCategory has changed
//               if (r[outputPyramidCatIdx] !== resultPyramidCategory) {

//                   console.log(`Pyramid Category change for ${email}: ${r[outputPyramidCatIdx]} to ${resultPyramidCategory}`);
//                   outputSheet.getRange(outputData.indexOf(r) + 2, outputPyramidCatIdx + 1).setValue(resultPyramidCategory);

//               }

//               if (r[outputReportMangIdx] !== resultReportingManager) {

//                   console.log(`Reporting Manager Change for ${email}: ${r[outputReportMangIdx]} to ${resultReportingManager}`);
//                   outputSheet.getRange(outputData.indexOf(r) + 2, outputReportMangIdx + 1).setValue(resultReportingManager);

//               }

//               break;
//           }
//       }

//       // If email not found in outputData, add this row
//       if (!found) {
//           lastSrNo ++;
//           row[0] = lastSrNo
//           updatedOutputData.push(row);

//       }
//   });
//   // Append the new data (if any) to outputSheet
//   if (updatedOutputData.length > 0) {
//     const lastRow = outputSheet.getLastRow()
//     outputSheet.getRange(lastRow + 1, 1, updatedOutputData.length, updatedOutputData[0].length).setValues(updatedOutputData);
//     outputSheet.getRange(lastRow + 1, outputActiveIdx+1, updatedOutputData.length, 1).insertCheckboxes().setValue(false);
//   }
// }

















// function smeDBSheet() {
//   const inputSheet = SpreadsheetApp.openById("11b5A_ZDi9DskTiKJ0j4fBDkRbY-0I2h7rVRxr6KVK1o").getSheetByName('Details');
//   const spreadSheet = SpreadsheetApp.openById("1eJ-7nKeICwYXEBh4oEzLwC6SpkYzCw5ntYxarbD2jxA");
//   const outputSheet = spreadSheet.getSheetByName("SME DB");

//   const masterData = inputSheet.getRange(1, 1, inputSheet.getLastRow(), inputSheet.getLastColumn()).getValues().filter(row => row.some(Boolean));
//   const headers = masterData[0];
//   let data = masterData.slice(1);

//   const hrDepartmentData = departmentMapping();

//   const empNameIdx = headers.indexOf('Employee Name');
//   const reportManagerIdx = headers.indexOf('Reporting Manager');
//   const pyramidCategoryIdx = headers.indexOf('Pyramid Category');
//   // const desigIdx = headers.indexOf('Designation');
//   const deptIdx = headers.indexOf('Department');
//   const officialEmailIdx = headers.indexOf('Official email ID');
//   // const emplyoyeeCategoryIdx = headers.indexOf("Emp Category");

//   // const anotherInputSheet = spreadSheet.getSheetByName('Copy of Copy of QA Team');
//   // let qaSMETable = getTableData(anotherInputSheet, 1, anotherInputSheet.getLastRow()+1, anotherInputSheet.getLastColumn());
//   // // Logger.log(Object.values(qaSMETable))
//   // qaSMETable = inverseMapping(qaSMETable);


//   // CHECK
//   const resultArray = [];
//   // let srNo = 1;

//   const outputRange = outputSheet.getRange(1, 1, outputSheet.getLastRow(), outputSheet.getLastColumn()).getValues();
//   const outputHeaders = outputRange[0], outputData = outputRange.slice(1);

//   // const outputSNoIdx = outputHeaders.indexOf('Sr. No.');
//   const outputEmailIdIdx = outputHeaders.indexOf('Email ID');
//   const outputDepartmentIdx = outputHeaders.indexOf('Department');
//   // const outputSMENameIdx = outputHeaders.indexOf('SME Name');
//   // const outputQAReviewerIdx = outputHeaders.indexOf('QA Reviewer');
//   // const outputQAReviewer2Idx = outputHeaders.indexOf('QA Reviewer 2');
//   const outputPyramidCatIdx = outputHeaders.indexOf('Pyramid Category');
//   const outputReportMangIdx = outputHeaders.indexOf('Reporting Manager');
//   // const outputTLEmailIdx = outputHeaders.indexOf('TL');
//   // const outputSMECEmailIdx = outputHeaders.indexOf('SME C');
//   // const outputSMEBEmailIdx = outputHeaders.indexOf('SME B');
//   // const outputSMEAEmailIdx = outputHeaders.indexOf('SME A');
//   const outputActiveIdx = outputHeaders.indexOf('Active?');

//   let index = 1;


//   data.forEach(r => {

//     if(hrDepartmentData.hasOwnProperty(r[deptIdx]) && r[officialEmailIdx] !== '-'){

//       const department = hrDepartmentData[r[deptIdx]];

//       if ( r[empNameIdx] !== 'Tushar Jangale' ) {

//         resultArray.push([index++, r[officialEmailIdx], department, r[empNameIdx], "", "", r[pyramidCategoryIdx], r[reportManagerIdx]]);
//       }

//     }
//     if(r[empNameIdx] === 'Tejas Jagtap'){

//       resultArray.push([index++, r[officialEmailIdx], 'Others', r[empNameIdx], "", "", r[pyramidCategoryIdx], r[reportManagerIdx]]);

//     }
//   })


//   // data.forEach((r) =>{
//   //   if(hrDepartmentData.hasOwnProperty(r[deptIdx]) && r[officialEmailIdx] !== '-'){
//   //     const department = hrDepartmentData[r[deptIdx]]
//   //     if (r[empNameIdx] !== 'Tushar Jangale'){
//   //       if(qaSMETable.hasOwnProperty(r[empNameIdx])){
//   //         const reviewer = qaSMETable[r[empNameIdx]];
//   //         resultArray.push([index++, r[officialEmailIdx], department, r[empNameIdx], reviewer, "", r[pyramidCategoryIdx], r[reportManagerIdx]]);
//   //       }  
//   //       else
//   //       resultArray.push([index++, r[officialEmailIdx], department, r[empNameIdx], "", "", r[pyramidCategoryIdx], r[reportManagerIdx]]);
//   //     }
//   //   }
//   //   if(r[empNameIdx] === 'Tejas Jagtap'){
//   //     resultArray.push([index++, r[officialEmailIdx], 'Others', r[empNameIdx], "", "", r[pyramidCategoryIdx], r[reportManagerIdx]]);
//   //   }
//   // })

//   const updatedOutputData = [];
//   resultArray.forEach((row, index) => {
//       const email = row[1];
//       const resultDepartment = row[2];
//       const resultPyramidCategory = row[6];
//       const resultReportingManager = row[7]
//       let found = false;

//       // Find the row in outputData with this email
//       for (let r of outputData) {
//           if (r[outputEmailIdIdx] === email) {
//               found = true;
//               // Check if department has changed
//               if (r[outputDepartmentIdx] !== resultDepartment) {
//                   console.log(`Department change for ${email}: ${r[outputDepartmentIdx]} to ${resultDepartment}`);
//                   outputSheet.getRange(outputData.indexOf(r) + 2, outputDepartmentIdx + 1).setValue(resultDepartment);
//               }

//               // Check if pyramidCategory has changed
//               if (r[outputPyramidCatIdx] !== resultPyramidCategory) {
//                   console.log(`Pyramid Category change for ${email}: ${r[outputPyramidCatIdx]} to ${resultPyramidCategory}`);
//                   outputSheet.getRange(outputData.indexOf(r) + 2, outputPyramidCatIdx + 1).setValue(resultPyramidCategory);
//               }

//               if (r[outputReportMangIdx] !== resultReportingManager) {
//                   console.log(`Reporting Manager Change for ${email}: ${r[outputReportMangIdx]} to ${resultReportingManager}`);
//                   outputSheet.getRange(outputData.indexOf(r) + 2, outputReportMangIdx + 1).setValue(resultReportingManager);
//               }
//               break;
//           }
//       }

//       // If email not found in outputData, add this row
//       if (!found) {
//           updatedOutputData.push(row);
//       }
//   });
//   // Append the new data (if any) to outputSheet
//   if (updatedOutputData.length > 0) {
//     const lastRow = outputSheet.getLastRow()
//     outputSheet.getRange(lastRow + 1, 1, updatedOutputData.length, updatedOutputData[0].length).setValues(updatedOutputData);
//     outputSheet.getRange(lastRow + 1, outputActiveIdx+1, updatedOutputData.length, 1).insertCheckboxes().setValue(false);
//   }
// }




function getTableData(sheet, startRowIndex, numberRows, endColumn) {
  // numberRows = numberRows-startRowIndex+1
  const data = sheet.getRange(startRowIndex, 2, numberRows, endColumn).getValues().filter(r => r !== '');
  return createObject(data);
}


function inverseMapping(obj) {
  const inverseMap = {}
  for (const [key, values] of Object.entries(obj)) {
    for (const val of values) {
      inverseMap[val] = key;
    }
  }
  return inverseMap;
}

function createObject(data) {
  const keys = data[0];
  const values = data.slice(1);

  const result = {};

  values.forEach(valueRow => {
    valueRow.forEach((value, index) => {
      const key = keys[index];
      if (!result[key]) {
        result[key] = [];
      }
      if (value != '' && value != 'other' && value != 'Other')
        result[key].push(value);
    });
  });
  return result;
}

function getReportingManagers(employee, reportingManagerMap) {
  const reportingManagers = [];
  if (employee in reportingManagerMap) {
    const directManager = reportingManagerMap[employee];
    reportingManagers.push(directManager);
    // Recursively find reporting managers of the direct manager
    const indirectManagers = getReportingManagers(directManager, reportingManagerMap);

    reportingManagers.push(...indirectManagers);
  }
  return reportingManagers;
}

function applyCustomFormatting(range, options) {

  options = options || {};

  var fontSize = options.fontSize || 10;
  var fontColor = options.fontColor || 'black';
  var bgColor = options.bgColor || 'white';
  var fontWeight = options.fontWeight || 'normal'

  range.setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true)
    .setFontFamily("Roboto")
    .setFontSize(fontSize)
    .setFontColor(fontColor)
    .setFontWeight(fontWeight)
    .setBorder(true, true, true, true, true, true)
    .setBackground(bgColor);
  return range;
};


Object.prototype.applyCustomFormatting = function (options) {
  return applyCustomFormatting(this, options);
}