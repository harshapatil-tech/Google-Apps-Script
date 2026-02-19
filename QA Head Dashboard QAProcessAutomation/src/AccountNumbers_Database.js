//get All department from Backend_Topic sheet
function getDepartmentList() {
  const inputSheet = SpreadsheetApp.openById("1u4i5rMCWM0mSgDWdf6EmGVi40kb9i2Qx51SE1cupIjU").getSheetByName("Backend - Topic");
  const inputDataRange = inputSheet.getRange(2, 1, inputSheet.getLastRow(), inputSheet.getLastColumn()).getValues();
  const inputHeaders = inputDataRange[0], inputData = inputDataRange.slice(1);
  const departmentIdx = inputHeaders.indexOf('Department');
  return [... new Set(inputData.map(r => r[departmentIdx]))];

}

//From Backend-Account Number sheet View Account numbers data in Account Managment sheet 
function viewAccountNumbers() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const outputSheet = spreadsheet.getSheetByName("Account Management");
  const dataRange = outputSheet.getRange(1, 2, outputSheet.getLastRow(), outputSheet.getLastColumn() - 1).getValues();
  const subjects = dataRange[0];
  const update = dataRange[1];
  const accounts = dataRange.slice(2);

  const inputSheet = SpreadsheetApp.openById("1u4i5rMCWM0mSgDWdf6EmGVi40kb9i2Qx51SE1cupIjU").getSheetByName("Backend - Account Numbers"); //1eJ-7nKeICwYXEBh4oEzLwC6SpkYzCw5ntYxarbD2jxA
  const inputDataRange = inputSheet.getRange(1, 2, inputSheet.getLastRow(), inputSheet.getLastColumn() - 1).getValues();
  const inputData = inputDataRange.slice(1), inputHeader = inputDataRange[0];

  const inputIndices = {
    maths: inputHeader.indexOf("Mathematics"),
    stats: inputHeader.indexOf("Statistics"),
    chem: inputHeader.indexOf("Chemistry"),
    bio: inputHeader.indexOf("Biology"),
    engg: inputHeader.indexOf("Engineering"),
    phy: inputHeader.indexOf("Physics"),
    business: inputHeader.indexOf("Business"),
    compSc: inputHeader.indexOf("Computer Science"),
    adobe: inputHeader.indexOf("Adobe& IT"),
  }

  const outputIndices = {
    maths: subjects.indexOf("Mathematics"),
    stats: subjects.indexOf("Statistics"),
    chem: subjects.indexOf("Chemistry"),
    bio: subjects.indexOf("Biology"),
    engg: subjects.indexOf("Engineering"),
    phy: subjects.indexOf("Physics"),
    business: subjects.indexOf("Business"),
    compSc: subjects.indexOf("Computer Science"),
    adobe: subjects.indexOf("Adobe& IT"),
  }

  let startRowIndex = 3
  inputData.forEach(r => {
    outputSheet.getRange(startRowIndex, outputIndices.maths + 2).setValue(r[inputIndices.maths])
    outputSheet.getRange(startRowIndex, outputIndices.stats + 2).setValue(r[inputIndices.stats])
    outputSheet.getRange(startRowIndex, outputIndices.chem + 2).setValue(r[inputIndices.chem])
    outputSheet.getRange(startRowIndex, outputIndices.bio + 2).setValue(r[inputIndices.bio])
    outputSheet.getRange(startRowIndex, outputIndices.engg + 2).setValue(r[inputIndices.engg])
    outputSheet.getRange(startRowIndex, outputIndices.phy + 2).setValue(r[inputIndices.phy])
    outputSheet.getRange(startRowIndex, outputIndices.business + 2).setValue(r[inputIndices.business])
    outputSheet.getRange(startRowIndex, outputIndices.compSc + 2).setValue(r[inputIndices.compSc])
    outputSheet.getRange(startRowIndex, outputIndices.adobe + 2).setValue(r[inputIndices.adobe])
    startRowIndex++;
  })

}


//From Backend-Account Number sheet Update Account numbers data in Account Managment sheet 
function updateAccountNumbers() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = spreadsheet.getSheetByName("Account Management");
  const dataRange = inputSheet.getRange(1, 2, inputSheet.getLastRow(), inputSheet.getLastColumn() - 1).getValues();
  const subjects = dataRange[0];
  const update = dataRange[1];
  const accounts = dataRange.slice(2);

  // Get the update status
  const updateColumnNumbers = update.map((ele, index) => ele === true ? index : -1).filter(index => index !== -1);

  if (updateColumnNumbers.length > 0) { // Check if there are selected subjects
    const outputSheet = SpreadsheetApp.openById("1u4i5rMCWM0mSgDWdf6EmGVi40kb9i2Qx51SE1cupIjU").getSheetByName("Backend - Account Numbers"); //1eJ-7nKeICwYXEBh4oEzLwC6SpkYzCw5ntYxarbD2jxA
    const outputDataRange = outputSheet.getRange(1, 2, outputSheet.getLastRow(), outputSheet.getLastColumn() - 1).getValues();
    const outputData = outputDataRange.slice(1), outputHeader = outputDataRange[0];

    updateColumnNumbers.forEach(columnIndex => {
      const subject = subjects[columnIndex];
      const outputSubjectIndex = outputHeader.indexOf(subject); // Get the index from inputIndices
      const outputColumnData = outputSheet.getRange(2, outputSubjectIndex + 1, outputSheet.getLastRow(), 1).getValues().filter(r => r !== '');
      if (outputSubjectIndex !== -1) { // Check if the subject exists in inputIndices
        const columnData = accounts.map(row => row[outputSubjectIndex]); // Extract data from the specific column
        // Filter out empty values from columnData
        const nonEmptyColumnData = columnData.filter(value => value !== "");

        if (nonEmptyColumnData.length > 0) { // Check if there are non-empty values to set
          // Clear the column before setting new values
          const outputColumn = outputSheet.getRange(2, columnIndex + 2, outputColumnData.length, 1);
          outputColumn.clearContent();

          // Set the non-empty values in the output column
          const valuesToSet = nonEmptyColumnData.map(value => [value]);
          const outputColumnRange = outputSheet.getRange(2, columnIndex + 2, nonEmptyColumnData.length, 1);
          outputColumnRange.setValues(valuesToSet);
        }
      }
    });
  }
}


function clearCellsBelow(sheet, startRow, startColumn) {

  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var numRows = lastRow - startRow + 1;

  if (numRows > 0) {
    sheet.getRange(startRow, startColumn, numRows, lastColumn).clearContent();
  }
}






