//view and update data of QA Reviwer & QA Reviwer2 
class TeamManagement {
  constructor() {
    // 1.Reviwer managment sheet
    this.inputSheet = CentralLibrary.DataAndHeaders(QA_HEAD_DASHBORD_ID)
      .getSheetById(REVIEWER_MANAGEMENT_TAB_ID).sheet;

    // smeSpredsheet
    const smeWrapper = CentralLibrary.DataAndHeaders(MASTER_DB_SPREADSHEET_ID);

    //2.SMEDB tab
    const smeObj = smeWrapper.getSheetById(SME_DB_TAB_ID);
    this.smeSheet = smeObj.sheet;
    const [smeHeaders, smeData] = smeObj.getDataIndicesFromSheet();
    this.smeHeaders = smeHeaders;
    //console.log("SME Headers:-",this.smeHeaders);
    this.smeData = smeData;
    //console.log("SME Data:-",this.smeData);

    //3.ReviwerDB tab
    const reviewerObj = smeWrapper.getSheetById(REVIEWER_DB_TAB_ID);
    this.reviewerSheet = reviewerObj.sheet;
    const [reviewerHeaders, reviewerData] = reviewerObj.getDataIndicesFromSheet();
    this.reviewerHeaders = reviewerHeaders;
    //console.log("Reviwer Headers are:-",this.reviewerHeaders);
    this.reviewerData = reviewerData;
    //console.log("Reviwer Data:-",this.reviewerData);
  }


  view() {
    const department = this.inputSheet.getRange(15, 3).getValue();
    //console.log("Department:-",department);
    const inputDataRange = this.inputSheet.getRange(17, 1, this.inputSheet.getLastRow() - 16, 5).getValues();
    //console.log("Input data range is:-",inputDataRange);
    const inputHeaders = inputDataRange[0];
    //console.log("Input Heders are:-",inputHeaders);
    const inputIdx = {
      srNoIdx: inputHeaders.indexOf("#"),
      smeIdx: inputHeaders.indexOf("SME Name"),
      reviewerIdx: inputHeaders.indexOf("QA Reviewer"),
      reviewer2Idx: inputHeaders.indexOf("QA Reviewer 2"),
      updateIdx: inputHeaders.indexOf("Update?")
    };
    //console.log("Input index are:-",inputIdx);

    const smeReviewer = this.smeData
      .filter(r => r[this.smeHeaders["Department"]]?.toLowerCase().trim() === department.toLowerCase().trim())
      .filter(r =>
        (r[this.smeHeaders["Active?"]] === false && r[this.smeHeaders["Removed Date"]] === "") ||
        (r[this.smeHeaders["Active?"]] === true && r[this.smeHeaders["Removed Date"]] !== "") ||
        (r[this.smeHeaders["Active?"]] === true && r[this.smeHeaders["Removed Date"]] === "")
      );
    //console.log("Sme Reviwers are:-",smeReviewer);

    const uniqueReviewer = this.reviewerData
      .filter(r => r[this.reviewerHeaders["Department"]]?.toLowerCase().trim() === department.toLowerCase().trim())
      .filter(r => r[this.reviewerHeaders["Active?"]] === true)
      .map(r => r[this.reviewerHeaders["Reviewer Name"]]);
    //console.log("Unique reviwers are:-",uniqueReviewer);

    const validation = SpreadsheetApp.newDataValidation()
      .requireValueInList([...uniqueReviewer, "-"])
      .setAllowInvalid(false)
      .build();
    //console.log("Validation",validation);

    const startRow = 18;
    //console.log("Start Row is:-",startRow);
    clearRange(this.inputSheet, startRow, 1, 5);

    const rows = [];
    //console.log("Rows are:-",rows);

    smeReviewer.forEach((row, index) => {
      const sme = row[this.smeHeaders["SME Name"]];
      //console.log("SME Name is:-",sme);
      const reviewer = row[this.smeHeaders["QA Reviewer"]];
      //console.log("QA Reviwer is:-",reviewer);
      const reviewer2 = row[this.smeHeaders["QA Reviewer 2"]];
      //console.log("QA Reviwer2 is:-",reviewer2);

      if (reviewer === '' || !uniqueReviewer.includes(reviewer)) {
        rows.push([index + 1, sme, "-", reviewer2 && uniqueReviewer.includes(reviewer2) ? reviewer2 : "-"]);
      } else {
        rows.push([index + 1, sme, reviewer, reviewer2 && uniqueReviewer.includes(reviewer2) ? reviewer2 : "-"]);
      }
    });

    this.inputSheet.getRange(startRow, inputIdx.updateIdx + 1, rows.length, 1).insertCheckboxes();
    this.inputSheet.getRange(startRow, inputIdx.reviewerIdx + 1, rows.length, 1).setDataValidation(validation);
    this.inputSheet.getRange(startRow, inputIdx.reviewer2Idx + 1, rows.length, 1).setDataValidation(validation);
    this.inputSheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);

    const currentRowNum = rows.length + startRow - 1;
    //console.log("current row num is:-",currentRowNum);

    let startDashboardRow = 19;
    //console.log("startDashbord is:-",startDashboardRow);

    clearRange(this.inputSheet, startDashboardRow, 7, 8);

    this.inputSheet.getRange(17, 8).setFormula(`=COUNTIF(C${startRow}:C${currentRowNum}, "-")`);

    for (const reviewer of uniqueReviewer) {
      const reviewerCell = this.inputSheet.getRange(startDashboardRow, 7);
      //console.log("reviwer cell is:-",reviewerCell);
      reviewerCell.setValue(reviewer);
      const countFormula = `=COUNTIF(C${startRow}:D${currentRowNum}, ${reviewerCell.getA1Notation()})`;
      //console.log("countformula is:-",countFormula);

      this.inputSheet.getRange(startDashboardRow, 8).setFormula(countFormula);
      startDashboardRow++;
    }

    if (uniqueReviewer.length > 0) {
      const colorRange = this.inputSheet.getRange(19, 8, uniqueReviewer.length, 1);
      //console.log("Color range is:-",colorRange);
      const rule = SpreadsheetApp.newConditionalFormatRule()
        .setGradientMinpointWithValue("#57BB8A", SpreadsheetApp.InterpolationType.NUMBER, 1)
        .setGradientMidpointWithValue("#fbbc04", SpreadsheetApp.InterpolationType.NUMBER, 5)
        .setGradientMaxpointWithValue("#F00000", SpreadsheetApp.InterpolationType.NUMBER, 10)
        .setRanges([colorRange])
        .build();

      const rules = this.inputSheet.getConditionalFormatRules();
      rules.push(rule);
      this.inputSheet.setConditionalFormatRules(rules);
    }
  }


  update() {
    const startRow = 18;
    console.log("Start Row:-", startRow);
    const numCols = 5;
    console.log("Number of column", numCols);

    const inputValues = this.inputSheet.getRange(startRow, 1, this.inputSheet.getLastRow() - startRow + 1, numCols).getValues();
    console.log("Input Values are:-", inputValues);

    const smeNameIdx = this.smeHeaders["SME Name"];
    console.log("SME Name:- ", smeNameIdx);
    const reviewerIdx = this.smeHeaders["QA Reviewer"];
    console.log("QA Reviwer", reviewerIdx);
    const reviewer2Idx = this.smeHeaders["QA Reviewer 2"];
    console.log("QA Reviwer2", reviewer2Idx);

    inputValues.forEach((row, i) => {
      const smeName = row[1];
      console.log("smename :-", smeName);
      const reviewer = row[2] === "-" ? "" : row[2];
      console.log("reviwer:-", reviewer);
      const reviewer2 = row[3] === "-" ? "" : row[3];
      console.log("reviwer2", reviewer2);
      const update = row[4];
      console.log("update:-", update);


      if (update === true && smeName) {
        const matchIndex = this.smeData.findIndex(
          r => (r[smeNameIdx] + "").toLowerCase().trim() === (smeName + "").toLowerCase().trim()
        );
        console.log("matchindex:-", matchIndex);

        if (matchIndex !== -1) {
          const sheetRow = matchIndex + 2;
          this.smeSheet.getRange(sheetRow, reviewerIdx + 1).setValue(reviewer);
          this.smeSheet.getRange(sheetRow, reviewer2Idx + 1).setValue(reviewer2);
        }
        this.inputSheet.getRange(startRow + i, 5).setValue(false);
      }
    });
  }
}


function viewTeamManagement() {
  const manager = new TeamManagement();
  manager.view();

}


function updateTeamManagement() {
  const manager2 = new TeamManagement();
  manager2.update();
}


function protectRange(range, description) {
  const protection = range.protect().setDescription(description);
  const me = Session.getEffectiveUser();
  // protection.addEditor(me);
  protection.removeEditor(me);

  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}



function unprotectRange(sheet) {
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  const me = Session.getEffectiveUser();
  for (const protection of protections) {
    protection.addEditor(me.getEmail());
    protection.remove();
  }
}


function clearRange(sheet, startRow, startCol, endCol) {
  var lastRow = sheet.getLastRow();
  var numRows = lastRow - startRow + 1;
  const numCols = endCol - startCol + 1;

  if (numRows > 0) {
    var range = sheet.getRange(startRow, startCol, numRows, numCols);
    range.setBorder(null, null, null, null, null, null);
    var dropdowns = range.getDataValidations();

    // Clear dropdowns
    for (var i = 0; i < numRows; i++) {
      for (var j = 0; j < numCols; j++) {
        if (dropdowns[i][j] != null) {
          dropdowns[i][j] = null;
        }
      }
    }

    // Clear content and remove checkboxes
    range.clearContent().clear().removeCheckboxes();

    // Set the modified dataValidations back to the range
    range.setDataValidations(dropdowns);
  }
}


function applyColorScale(start, end) {
  // Replace these values with your specific sheet and range
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange("G" + start + ":G" + end); // Change to your desired range
  console.log(range.getA1Notation())
  // Find the maximum and minimum values in the range
  const values = range.getValues().flat(); // Get the values as a flat array
  const max = Math.max(...values);
  const min = Math.min(...values);

  // Define the colors for the color scale (from red to green)
  const minColor = "#FF0000"; // Red
  const maxColor = "#00FF00"; // Green

  // Loop through each cell in the range
  for (let i = 1; i <= range.getHeight(); i++) {
    for (let j = 1; j <= range.getWidth(); j++) {
      const cell = range.getCell(i, j);
      const cellValue = cell.getValue();

      // Calculate the color based on the cell's value and the range's min and max
      const colorScale = (cellValue - min) / (max - min);
      const cellColor = getColorGradient(minColor, maxColor, colorScale);

      // Set the background color of the cell
      cell.setBackground(cellColor);
    }
  }
}


function getColorGradient(startColor, endColor, scale) {
  const startRGB = hexToRgb(startColor);
  const endRGB = hexToRgb(endColor);

  const r = Math.round(startRGB.r + (endRGB.r - startRGB.r) * scale);
  const g = Math.round(startRGB.g + (endRGB.g - startRGB.g) * scale);
  const b = Math.round(startRGB.b + (endRGB.b - startRGB.b) * scale);

  return rgbToHex(r, g, b);
}


function hexToRgb(hex) {
  const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return result
    ? {
      r: parseInt(result[1], 16),
      g: parseInt(result[2], 16),
      b: parseInt(result[3], 16),
    }
    : null;
}


function rgbToHex(r, g, b) {
  const red = r.toString(16).padStart(2, '0');
  const green = g.toString(16).padStart(2, '0');
  const blue = b.toString(16).padStart(2, '0');
  return `#${red}${green}${blue}`;
}

//_______________OLD viewTeamManagment code_____________________________________ 

// function viewTeamManagement(){
//   const spreadsheet = SpreadsheetApp.openById("1FrKuzuyN6uo1UP41MAvD9zLvPUMNTOzBYz4iwHl7VhY");
//   const inputSheet = spreadsheet.getSheetByName("Reviewer Management");
//   const inputDataRange = inputSheet.getRange(17, 1, inputSheet.getLastRow(), 5).getValues();
//   const inputHeaders = inputDataRange[0], inputData = inputDataRange.slice(1);
//   const department = inputSheet.getRange(15, 3).getValue();

//   const inputIndices = {
//     srNoIdx : inputHeaders.indexOf("#"),
//     smeIdx : inputHeaders.indexOf("SME Name"),
//     reviewerIdx : inputHeaders.indexOf("QA Reviewer"),
//     reviewer2Idx : inputHeaders.indexOf("QA Reviewer 2"),
//     updateIdx : inputHeaders.indexOf("Update?"),
//   }

//   const backendSheet = SpreadsheetApp.openById("1eJ-7nKeICwYXEBh4oEzLwC6SpkYzCw5ntYxarbD2jxA");
//   const smeDB = backendSheet.getSheetByName("SME DB");
//   const reviewerDB = backendSheet.getSheetByName("Reviewer DB");
//   const smeDBDataRange = smeDB.getRange(1, 1, smeDB.getLastRow(), smeDB.getLastColumn()).getValues();
//   const reviewerDBDataRange = reviewerDB.getRange(1, 1, reviewerDB.getLastRow(), reviewerDB.getLastColumn()).getValues();
//   const smeDBHeaders = smeDBDataRange[0], smeDBData = smeDBDataRange.slice(1)
//   const reviewerDBHeaders = reviewerDBDataRange[0], reviewerDBData = reviewerDBDataRange.slice(1);

//   const smeDBIndices = {
//     smeNameIdx : smeDBHeaders.indexOf("SME Name"),
//     reviewerNameIdx : smeDBHeaders.indexOf("QA Reviewer"),
//     rewviewerName2Idx : smeDBHeaders.indexOf("QA Reviewer 2"),
//     departmentIdx : smeDBHeaders.indexOf("Department"),
//     pyramidCategoryIdx : smeDBHeaders.indexOf('Pyramid Category'),
//     addedDateIdx : smeDBHeaders.indexOf("Added Date"),
//     activeIdx : smeDBHeaders.indexOf("Active?"),
//     removedDateIdx : smeDBHeaders.indexOf('Removed Date'),
//   }

//   const reviewerDBIndices = {
//     reviewName : reviewerDBHeaders.indexOf("Reviewer Name"),
//     departmentIdx :reviewerDBHeaders.indexOf("Department"),
//     activeStatusIdx : reviewerDBHeaders.indexOf("Active?"),
//   }


//   const smeReviewer = smeDBData.filter(r => r[smeDBIndices.departmentIdx].trim().toLowerCase() === department.trim().toLowerCase())
//                                .filter(r => r[smeDBIndices.pyramidCategoryIdx] !== 'Others')
//                                .filter(r=>  (r[smeDBIndices.activeIdx]===false && r[smeDBIndices.removedDateIdx]==='') || 
//                                             (r[smeDBIndices.activeIdx]===true && r[smeDBIndices.removedDateIdx]!=='') ||
//                                             (r[smeDBIndices.activeIdx]===true && r[smeDBIndices.removedDateIdx]===''));

//   const uniqueReviewer = reviewerDBData.filter(r => r[reviewerDBIndices.departmentIdx].trim().toLowerCase() === department.trim().toLowerCase())
//                                        .filter(r => r[reviewerDBIndices.activeStatusIdx] === true)
//                                        .map(r => r[reviewerDBIndices.reviewName]);

//   const validation = SpreadsheetApp.newDataValidation().requireValueInList([...uniqueReviewer, "-"]).setAllowInvalid(false).build();

//   const startRow = 18;
//   clearRange(inputSheet, startRow, 1, 5)

//   const rows = [];
//   smeReviewer.forEach((row, index) => {
//     const sme = row[smeDBIndices.smeNameIdx];
//     let reviewer = row[smeDBIndices.reviewerNameIdx];
//     let reviewer2 = row[smeDBIndices.rewviewerName2Idx];
//     if (reviewer === '' || !uniqueReviewer.includes(reviewer)) {

//       if (reviewer2 === '' || !uniqueReviewer.includes(reviewer2))
//         rows.push([index + 1, sme, "-", "-"]);
//       else if(reviewer2 !== '' || uniqueReviewer.includes(reviewer2))
//         rows.push([index + 1, sme, "-", reviewer2]);

//     }else if(reviewer !== '' || uniqueReviewer.includes(reviewer)){

//       if (reviewer2 === '' || !uniqueReviewer.includes(reviewer2))
//         rows.push([index + 1, sme, reviewer, "-"]);
//       else if(reviewer2 !== '' || uniqueReviewer.includes(reviewer2)){
//         rows.push([index + 1, sme, reviewer, reviewer2]);
//       }
//     }
//   });  

//   applyCustomFormatting(inputSheet.getRange(startRow, inputIndices.updateIdx + 1, rows.length, 1)).insertCheckboxes()
//   const reviewerRange = inputSheet.getRange(startRow, inputIndices.reviewerIdx + 1, rows.length, 1);
//   applyCustomFormatting(reviewerRange).setDataValidation(validation);
//   const reviewerRange2 = inputSheet.getRange(startRow, inputIndices.reviewer2Idx + 1, rows.length, 1);
//   applyCustomFormatting(reviewerRange2).setDataValidation(validation);
//   applyCustomFormatting(inputSheet.getRange(startRow, 1, rows.length, rows[0].length)).setValues(rows)

//   // Show the counts
//   const currentRowNum = rows.length + startRow -1;

//   let startDashboardRow = 19;

//   clearRange(inputSheet, startDashboardRow, 7, 8)

//   formula = `=COUNTIF(C${startRow}:C${currentRowNum}, "-")`;
//   // Adjust the target range based on startDashboardRow
//   applyCustomFormatting(inputSheet.getRange(17, 8)).setFormula(formula);


//   for (const reviewer of uniqueReviewer) {
//     let formula;
//     const reviewerRange = applyCustomFormatting(inputSheet.getRange(startDashboardRow, 7));
//     reviewerRange.setValue(reviewer);
//     formula = `=COUNTIF(C${startRow}:D${currentRowNum}, ${reviewerRange.getA1Notation()})`;
//     // Adjust the target range based on startDashboardRow
//     applyCustomFormatting(inputSheet.getRange(startDashboardRow, 8)).setFormula(formula);
//     startDashboardRow++;
//   }

//   if (uniqueReviewer.length>0){
//     const colorRange = inputSheet.getRange(19, 8, uniqueReviewer.length, 1)

//     const colorScaleMin = "#57BB8A";
//     const colorScaleMax = "#F00000";
//     const values = [... Array(10).keys()].map(i=>i+1)

//     const minValue = Math.min.apply(null, values);
//     const maxValue = Math.max.apply(null, values);

//     const rule = SpreadsheetApp.newConditionalFormatRule()
//                               // .setGradientMinpoint(colorScaleMin)
//                               .setGradientMaxpointWithValue(colorScaleMax, SpreadsheetApp.InterpolationType.NUMBER, maxValue)
//                               .setGradientMidpointWithValue("#fbbc04", SpreadsheetApp.InterpolationType.NUMBER, (maxValue + minValue)/2)
//                               .setGradientMinpointWithValue(colorScaleMin, SpreadsheetApp.InterpolationType.NUMBER, minValue)
//                               // .setGradientMaxpointConditionalFormatRules();(colorScaleMax)
//                               .setRanges([colorRange])
//                               .build();

//     const rules = inputSheet.getConditionalFormatRules();
//     rules.push(rule);
//     inputSheet.setConditionalFormatRules(rules);
//   }
// }



// function updateTeamManagement(){
//   const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
//   const inputSheet = spreadsheet.getSheetByName("Reviewer Management");
//   const inputDataRange = inputSheet.getRange(17, 2, inputSheet.getLastRow(), 4).getValues();
//   const inputHeaders = inputDataRange[0], inputData = inputDataRange.slice(1);

//   const outputSheet = SpreadsheetApp.openById("1eJ-7nKeICwYXEBh4oEzLwC6SpkYzCw5ntYxarbD2jxA").getSheetByName("SME DB");
//   const outputDataRange = outputSheet.getRange(1, 1, outputSheet.getLastRow(), outputSheet.getLastColumn()).getValues();
//   const outputHeaders = outputDataRange[0], outputData = outputDataRange.slice(1);

//   const inputIndices = {
//     smeIdx : inputHeaders.indexOf("SME Name"),
//     reviewerIdx : inputHeaders.indexOf("QA Reviewer"),
//     reviewer2Idx : inputHeaders.indexOf("QA Reviewer 2"),
//     updateIdx : inputHeaders.indexOf("Update?"),
//   }

//   const outputIndices = {
//     smeNameIdx : outputHeaders.indexOf("SME Name"),
//     reviewerNameIdx : outputHeaders.indexOf("QA Reviewer"),
//     reviewer2NameIdx : outputHeaders.indexOf("QA Reviewer 2"),
//   }


//   inputData.forEach((row, index) => {
//     // Find the sme name in the SME DB
//     if(row[inputIndices.updateIdx] === true){
//       const smeName = row[inputIndices.smeIdx];
//       const reviewer = row[inputIndices.reviewerIdx];
//       const reviewer2 = row[inputIndices.reviewer2Idx];
//       const foundIndex = outputData.findIndex(outputRow => outputRow[outputIndices.smeNameIdx] === smeName);
//       if (foundIndex !== -1){
//         if(reviewer === '-'){
//           outputSheet.getRange(foundIndex + 2, outputIndices.reviewerNameIdx + 1).setValue("");
//         }else{
//           outputSheet.getRange(foundIndex + 2, outputIndices.reviewerNameIdx + 1).setValue(reviewer);
//           if(reviewer2 === '')
//             outputSheet.getRange(foundIndex + 2, outputIndices.reviewer2NameIdx + 1).setValue("");
//           else
//             outputSheet.getRange(foundIndex + 2, outputIndices.reviewer2NameIdx + 1).setValue(reviewer2);
//         }
//     }
//     inputSheet.getRange(index+18, inputIndices.updateIdx+2).setValue(false)
//     }  
//   })
// }





