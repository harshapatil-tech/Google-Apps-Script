function departmentMapping(){
  const spreadsheet = SpreadsheetApp.openById("1u4i5rMCWM0mSgDWdf6EmGVi40kb9i2Qx51SE1cupIjU");
  const inputsheet = spreadsheet.getSheetByName("Backend - Dept Mapping");
  const inputDataRange = inputsheet.getRange(1, 1, inputsheet.getLastRow(), inputsheet.getLastColumn()).getValues();
  const inputHeaders = inputDataRange[0], inputData = inputDataRange.slice(1);

  const inputIndices = {
    hrDepartmentIndex : inputHeaders.indexOf("HR Department"),
    qaDepartmentIndex : inputHeaders.indexOf("QA Department"),
  }

  const departmentMap = {}
  inputData.forEach(r => {
    if (!departmentMap.hasOwnProperty(r[inputIndices.hrDepartmentIndex])){
      departmentMap[r[inputIndices.hrDepartmentIndex]] = r[inputIndices.qaDepartmentIndex]
    }
  })

  return departmentMap;
}





// function smeDBSheet() {
//   const inputSheet = SpreadsheetApp.openById("11b5A_ZDi9DskTiKJ0j4fBDkRbY-0I2h7rVRxr6KVK1o").getSheetByName('Details');
//   const spreadSheet = SpreadsheetApp.openById("1eJ-7nKeICwYXEBh4oEzLwC6SpkYzCw5ntYxarbD2jxA");
//   const outputSheet = spreadSheet.getSheetByName("Copy of SME DB");

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




