//From Backend-Other data view data in Other Data sheet
function viewOtherData() {
  const inputSheet = SpreadsheetApp.openById("1u4i5rMCWM0mSgDWdf6EmGVi40kb9i2Qx51SE1cupIjU").getSheetByName("Backend - Other Data");

  //1eJ-7nKeICwYXEBh4oEzLwC6SpkYzCw5ntYxarbD2jxA
  const inputDataRange = inputSheet.getRange(1, 1, inputSheet.getLastRow(), inputSheet.getLastColumn()).getValues();
  const inputHeaders = inputDataRange[0], inputData = inputDataRange.slice(1);

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const outputSheet = spreadsheet.getSheetByName("Other Data");
  const outputDataRange = outputSheet.getRange(3, 2, outputSheet.getLastRow() - 2, outputSheet.getLastColumn() - 1).getValues();
  const updateRow = outputDataRange[0], outputHeaders = outputDataRange[1], outputData = outputDataRange.slice(2);

  const inputIndices = {
    client: inputHeaders.indexOf("Client"),
    mode: inputHeaders.indexOf("Mode"),
    audio: inputHeaders.indexOf("Audio"),
    ratings: inputHeaders.indexOf("Rating"),
    negRevReason: inputHeaders.indexOf("Reasons for negative ratings"),
    clientComplaints: inputHeaders.indexOf("Client Complaints"),
    //mapping : inputHeaders.indexOf("Mapping"),
    discussion: inputHeaders.indexOf("Discussion"),
    identyIdx: inputHeaders.indexOf('SubjectKnowledge_Identify'),
    breakProcessIdx: inputHeaders.indexOf('SubjectKnowledge_Break The Process'),
    explanationIdx: inputHeaders.indexOf('SubjectKnowledge_Explanation'),
    encourageIdx: inputHeaders.indexOf('Tutoring_Encourage'),
    tutoringFlowIdx: inputHeaders.indexOf('Tutoring_Session Flow'),
    socraticIdx: inputHeaders.indexOf('Tutoring_Socratic'),
    greetingIdx: inputHeaders.indexOf('Admin_Greeting/ closing'),
    policiesIdx: inputHeaders.indexOf('Admin_Client policies'),
    englishFlowIdx: inputHeaders.indexOf('Communication_English'),
    effectiveFlowIdx: inputHeaders.indexOf('Communication_Effectiveness'),
    //lowRatedSession: inputHeaders.indexOf("Score of low rated sessions"),
    netTutor: inputHeaders.indexOf("NetTutor Client Ratings (Out of Five)"),
    inputTraining: inputHeaders.indexOf("Input to Training team"),

  }

  const outputIndices = {
    client: outputHeaders.indexOf("Client"),
    mode: outputHeaders.indexOf("Mode"),
    audio: outputHeaders.indexOf("Audio"),
    ratings: outputHeaders.indexOf("Rating"),
    negRevReason: outputHeaders.indexOf("Reasons for negative ratings"),
    clientComplaints: outputHeaders.indexOf("Client Complaints"),
    // mapping : outputHeaders.indexOf("Mapping"),
    discussion: outputHeaders.indexOf("Discussion"),

    identyIdx: outputHeaders.indexOf('SubjectKnowledge_Identify'),
    breakProcessIdx: outputHeaders.indexOf('SubjectKnowledge_Break The Process'),
    explanationIdx: outputHeaders.indexOf('SubjectKnowledge_Explanation'),
    encourageIdx: outputHeaders.indexOf('Tutoring_Encourage'),
    tutoringFlowIdx: outputHeaders.indexOf('Tutoring_Session Flow'),
    socraticIdx: outputHeaders.indexOf('Tutoring_Socratic'),
    greetingIdx: outputHeaders.indexOf('Admin_Greeting/ closing'),
    policiesIdx: outputHeaders.indexOf('Admin_Client policies'),
    englishFlowIdx: outputHeaders.indexOf('Communication_English'),
    effectiveFlowIdx: outputHeaders.indexOf('Communication_Effectiveness'),

    //lowRatedSession: outputHeaders.indexOf("Score of low rated sessions"),
    netTutor: outputHeaders.indexOf("NetTutor Client Ratings (Out of Five)"),
    inputTraining: outputHeaders.indexOf("Input to Training team"),

  }


  let startRowIndex = 5;
  inputData.forEach(r => {
    outputSheet.getRange(startRowIndex, outputIndices.client + 2).setValue(r[inputIndices.client]);
    outputSheet.getRange(startRowIndex, outputIndices.mode + 2).setValue(r[inputIndices.mode]);
    outputSheet.getRange(startRowIndex, outputIndices.audio + 2).setValue(r[inputIndices.audio]);
    outputSheet.getRange(startRowIndex, outputIndices.ratings + 2).setValue(r[inputIndices.ratings]);
    outputSheet.getRange(startRowIndex, outputIndices.negRevReason + 2).setValue(r[inputIndices.negRevReason]);
    outputSheet.getRange(startRowIndex, outputIndices.clientComplaints + 2).setValue(r[inputIndices.clientComplaints]);
    // outputSheet.getRange(startRowIndex, outputIndices.mapping + 2).setValue(r[inputIndices.mapping]);
    outputSheet.getRange(startRowIndex, outputIndices.discussion + 2).setValue(r[inputIndices.discussion]);
    outputSheet.getRange(startRowIndex, outputIndices.identyIdx + 2).setValue(r[inputIndices.identyIdx]);
    outputSheet.getRange(startRowIndex, outputIndices.breakProcessIdx + 2).setValue(r[inputIndices.breakProcessIdx]);
    outputSheet.getRange(startRowIndex, outputIndices.explanationIdx + 2).setValue(r[inputIndices.explanationIdx]);
    outputSheet.getRange(startRowIndex, outputIndices.encourageIdx + 2).setValue(r[inputIndices.encourageIdx]);
    outputSheet.getRange(startRowIndex, outputIndices.tutoringFlowIdx + 2).setValue(r[inputIndices.tutoringFlowIdx]);
    outputSheet.getRange(startRowIndex, outputIndices.socraticIdx + 2).setValue(r[inputIndices.socraticIdx]);
    outputSheet.getRange(startRowIndex, outputIndices.greetingIdx + 2).setValue(r[inputIndices.greetingIdx]);
    outputSheet.getRange(startRowIndex, outputIndices.policiesIdx + 2).setValue(r[inputIndices.policiesIdx]);
    outputSheet.getRange(startRowIndex, outputIndices.englishFlowIdx + 2).setValue(r[inputIndices.englishFlowIdx]);
    outputSheet.getRange(startRowIndex, outputIndices.effectiveFlowIdx + 2).setValue(r[inputIndices.effectiveFlowIdx]);
   // outputSheet.getRange(startRowIndex, outputIndices.lowRatedSession + 2).setValue(r[inputIndices.lowRatedSession]);
    outputSheet.getRange(startRowIndex, outputIndices.netTutor + 2).setValue(r[inputIndices.netTutor]);
    outputSheet.getRange(startRowIndex, outputIndices.inputTraining + 2)
      .setValue(r[inputIndices.inputTraining]);
    startRowIndex++;


  })

}

//Update data in Other Data sheet
function updateOtherData() {
  const outputSheet = SpreadsheetApp.openById("1u4i5rMCWM0mSgDWdf6EmGVi40kb9i2Qx51SE1cupIjU").getSheetByName("Backend - Other Data");
  const outputDataRange = outputSheet.getRange(1, 1, outputSheet.getLastRow(), outputSheet.getLastColumn()).getValues();
  const outputHeaders = outputDataRange[0], outputData = outputDataRange.slice(1);

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = spreadsheet.getSheetByName("Other Data");
  const inputDataRange = inputSheet.getRange(3, 2, inputSheet.getLastRow() - 2, inputSheet.getLastColumn() - 1).getValues();

  const updateRow = inputDataRange[0], inputHeaders = inputDataRange[1], inputData = inputDataRange.slice(2);

  const updateColumnNumbers = updateRow.map((ele, index) => ele === true ? index : -1).filter(index => index !== -1);

  updateColumnNumbers.forEach(columnIndex => {
    const dropdownColumns = inputHeaders[columnIndex];
    const inputDropdownIndex = inputHeaders.indexOf(dropdownColumns);
    const outputDropdownIndex = outputHeaders.indexOf(dropdownColumns);
    if (outputDropdownIndex !== -1) {
      const columnData = inputData.map(row => row[inputDropdownIndex]);
      if (columnData.length > 0) {
        // Clear the column before setting new values
        const outputColumn = outputSheet.getRange(2, outputDropdownIndex + 1, columnData.length, 1);
        outputColumn.clearContent();

        //Set the non-empty values in the output column
        const valuesToSet = columnData.map(value => [value]);
        applyCustomFormatting(outputColumn).setValues(valuesToSet)
      }
    }
  })
}















