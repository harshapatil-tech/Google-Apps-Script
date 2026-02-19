//From Backend-Topic sheet view data in Department managment sheet
function viewDepartmentData() {
  const inputSheet = SpreadsheetApp.openById("1u4i5rMCWM0mSgDWdf6EmGVi40kb9i2Qx51SE1cupIjU").getSheetByName("Backend - Topic");
  const inputDataRange = inputSheet.getRange(2, 1, inputSheet.getLastRow(), inputSheet.getLastColumn()).getValues();
  const inputHeaders = inputDataRange[0], inputData = inputDataRange.slice(1);

  const outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Department Management");
  const outputDataRange = outputSheet.getRange(4, 1, outputSheet.getLastRow(), outputSheet.getLastColumn()).getValues();
  const outputHeaders = outputDataRange[0], outputData = outputDataRange.slice(1);

  const dropdownValue = outputSheet.getRange(1, 3).getValue();

  const inputIndices = {
    srNoIdx: inputHeaders.indexOf("Sr No."),
    departmentIdx: inputHeaders.indexOf("Department"),
    subjectIdx: inputHeaders.indexOf("Subject"),
    topicIdx: inputHeaders.indexOf("Topic"),
    subTopicIdx: inputHeaders.indexOf("SubTopic"),
  };

  const outputIndices = {
    srNoIdx: outputHeaders.indexOf("#"),
    subjectIdx: outputHeaders.indexOf("Subject"),
    topicIdx: outputHeaders.indexOf("Topic"),
    subTopicIdx: outputHeaders.indexOf("SubTopic"),
    addUpdateBoxIdx: outputHeaders.indexOf("Add / Update?"),
    deleteBoxIdx: outputHeaders.indexOf("Delete?"),
  };

  const filteredData = dropdownValue === 'All'
    ? inputData
    : inputData.filter(r => r[inputIndices.departmentIdx].trim().toLowerCase() === dropdownValue.trim().toLowerCase());

  let startRow = 5;
  clearRowsBelow(outputSheet, 4);

  const numRows = filteredData.length;
  const numColumns = outputIndices.deleteBoxIdx + 1; // Number of columns to set checkboxes + 1 for data columns

  const batchValues = [];

  for (let i = 0; i < numRows; i++) {
    const r = filteredData[i];
    const newRow = [
      r[inputIndices.srNoIdx],
      r[inputIndices.subjectIdx],
      r[inputIndices.topicIdx],
      r[inputIndices.subTopicIdx],
    ];

    // Insert checkboxes for "Add / Update?" and "Delete?" columns
    newRow.push(false, false);

    batchValues.push(newRow);

    if (batchValues.length === 100 || i === numRows - 1) {

      const checkboxRange = outputSheet.getRange(startRow, outputIndices.addUpdateBoxIdx + 1, batchValues.length, 2);
      checkboxRange.insertCheckboxes();
      // Set values in batches of 100 rows or at the end of the loop
      outputSheet.getRange(startRow, 1, batchValues.length, numColumns).setValues(batchValues);
      applyCustomFormatting(outputSheet.getRange(startRow, 1, batchValues.length, numColumns));
      startRow += batchValues.length;
      batchValues.length = 0;
    }
  }
}


//Update data in Department managment sheet
function updateDepartmentData() {
  const outputSheet = SpreadsheetApp.openById("1u4i5rMCWM0mSgDWdf6EmGVi40kb9i2Qx51SE1cupIjU").getSheetByName("Backend - Topic");
  const outputDataRange = outputSheet.getRange(2, 1, outputSheet.getLastRow(), outputSheet.getLastColumn()).getValues();
  const outputHeaders = outputDataRange[0], outputData = outputDataRange.slice(1);

  const inputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Department Management");
  const inputDataRange = inputSheet.getRange(4, 1, inputSheet.getLastRow(), inputSheet.getLastColumn()).getValues().filter(r => r.some(Boolean));
  const inputHeaders = inputDataRange[0], inputData = inputDataRange.slice(1);

  const outputIndices = {
    srNoIdx: outputHeaders.indexOf("Sr No."),
    departmentIdx: outputHeaders.indexOf("Department"),
    subjectIdx: outputHeaders.indexOf("Subject"),
    topicIdx: outputHeaders.indexOf("Topic"),
    subTopicIdx: outputHeaders.indexOf("SubTopic"),
  }

  const inputIndices = {
    srNoIdx: inputHeaders.indexOf("#"),
    subjectIdx: inputHeaders.indexOf("Subject"),
    topicIdx: inputHeaders.indexOf("Topic"),
    subTopicIdx: inputHeaders.indexOf("SubTopic"),
    addUpdateBoxIdx: inputHeaders.indexOf("Add / Update?"),
    deleteBoxIdx: inputHeaders.indexOf("Delete?"),
  }

  // const x = inputData.filter(r => r[inputIndices.srNoIdx] === '')

  const filteredDataUpdate = inputData.filter(r => r[inputIndices.addUpdateBoxIdx] === true);
  const filteredDataDelete = inputData.filter(r => r[inputIndices.deleteBoxIdx] === true && r[inputIndices.srNoIdx] !== '');

  const onlySrNosDB = outputData.map(r => r[outputIndices.srNoIdx]);
  const onlySrNosDBInput = inputData.map(r => r[inputIndices.srNoIdx]);
  const dropdownValue = inputSheet.getRange(1, 3).getValue();

  filteredDataDelete.forEach(r => {
    // Now get the serialNumber
    let outputRowIndex = onlySrNosDB.indexOf(r[inputIndices.srNoIdx]) + 3;
    outputSheet.deleteRow(outputRowIndex);
  })

  filteredDataUpdate.forEach((r, index) => {
    let rowIndex = onlySrNosDB.indexOf(r[inputIndices.srNoIdx]) + 3;
    const inputRowIndex = onlySrNosDBInput.indexOf(r[inputIndices.srNoIdx]) + 5;
    if (r[inputIndices.srNoIdx] === '') {
      const lastRow = outputSheet.getLastRow();
      rowIndex = outputSheet.getRange(outputSheet.getLastRow(), outputIndices.srNoIdx + 1).getValue() + 1;
      outputSheet.getRange(lastRow + 1, outputIndices.srNoIdx + 1).setValue(rowIndex);
      outputSheet.getRange(lastRow + 1, outputIndices.departmentIdx + 1).setValue(dropdownValue);
      outputSheet.getRange(lastRow + 1, outputIndices.subjectIdx + 1).setValue(r[inputIndices.subjectIdx]);
      outputSheet.getRange(lastRow + 1, outputIndices.topicIdx + 1).setValue(r[inputIndices.topicIdx]);
      outputSheet.getRange(lastRow + 1, outputIndices.subTopicIdx + 1).setValue(r[inputIndices.subTopicIdx]);
      applyCustomFormatting(outputSheet.getRange(lastRow + 1, 1, 1, outputSheet.getLastColumn()));

    } else {
      outputSheet.getRange(rowIndex, outputIndices.subjectIdx + 1).setValue(r[inputIndices.subjectIdx]);
      outputSheet.getRange(rowIndex, outputIndices.topicIdx + 1).setValue(r[inputIndices.topicIdx]);
      outputSheet.getRange(rowIndex, outputIndices.subTopicIdx + 1).setValue(r[inputIndices.subTopicIdx]);
      inputSheet.getRange(inputRowIndex, inputIndices.addUpdateBoxIdx + 1).setValue(false)
      applyCustomFormatting(outputSheet.getRange(rowIndex, 1, 1, outputSheet.getLastColumn()));
    }
  })
  viewDepartmentData();
}


function clearRowsBelow(sheet, startRow) {

  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var numRows = lastRow - startRow + 1;

  if (numRows > 0) {
    sheet.getRange(startRow + 1, 1, numRows, lastColumn).clearContent().removeCheckboxes();
  }
}










