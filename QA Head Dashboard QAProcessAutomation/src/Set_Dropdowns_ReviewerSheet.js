// function testSetDropdownsQAReviewAdd() {
//   const ss = SpreadsheetApp.openById("1I6K0B3Dc-Wj4VWo1lTd49hO_a_DPAlgdSHykMrNoM8E");
  
//   const mainSheet = ss.getSheetByName("QA_Review_Add"); // sheet where dropdowns are applied
//   const viewSheet = ss.getSheetByName("QA_Review_Update"); // sheet containing headers
//   const backendSheet = ss.getSheetByName("Backend"); // sheet with all dropdown values
  
//   if (!mainSheet) { throw new Error("Sheet 'QA Review' not found!"); }
//   if (!viewSheet) { throw new Error("Sheet 'QA View' not found!"); }
//   if (!backendSheet) { throw new Error("Sheet 'Backend_Data' not found!"); }

//   // Make sure your sheet names match or create sheets with these names for testing
//   setDropdownsQAReviewAdd(mainSheet, viewSheet, backendSheet, 2);
  
//   SpreadsheetApp.flush();
//   Logger.log("Dropdowns, validations, and checkboxes applied successfully!");
// }

//Set DropDown Values In QA_Review_Add sheet
function setDropdownsQAReviewAdd(sheet, viewSheet, backendSheet, numCols=2){
  const dataRange = viewSheet.getRange(5, 1, viewSheet.getLastRow() - 4, viewSheet.getLastColumn());
  const viewHeaders = dataRange.getValues()[0]; 
  console.log(viewSheet.getName());
  console.log("Headers are:-",viewHeaders);

  const viewIndices = {
    srNoIdx: viewHeaders.indexOf("#") + 1,
    commentsIdx: viewHeaders.indexOf("Comments") + 1,
    studentCommentsIdx: viewHeaders.indexOf("Student's Comments") + 1,
   // totalMinsIdx: viewHeaders.indexOf("Total Hours (In Decimals)") + 1,
    totalMinsIdx: viewHeaders.indexOf("Review Time (Min)") + 1,
    discussionIdx:viewHeaders.indexOf("Discussion") + 1,
    //lowrateScoreIdx : viewHeaders.indexOf("Score of low rated sessions") +1 ,
    netTutorIdx: viewHeaders.indexOf("NetTutor Client Ratings (Out of Five)") + 1,
    inputToTrainingIdx: viewHeaders.indexOf("Input to Training team") + 1
  }
  
  const wholeSheet = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  console.log(sheet.getName());
  const headerRow = wholeSheet[0], dataRows = wholeSheet.slice(1);
  console.log("headers Add:-",headerRow);
  const headerRowIdx = 4
  const backendColumnNames = backendSheet.getRange(headerRowIdx, 1, 1, backendSheet.getLastColumn()).getValues().flat();
  console.log("Backend Headers are",backendColumnNames);
  const colIdxMap = getColumnIndices(headerRow);
  const backendMap = getColumnIndices(backendColumnNames);

  const srNoIdx = headerRow.indexOf('#') + 1;
  const smeNameIdx = headerRow.indexOf('SME Name') + 1;
  const clientIdx = headerRow.indexOf('Client') + 1;
  const subjectIdx = headerRow.indexOf('Subject') + 1;
  const topicIdx = headerRow.indexOf('Topic');
  //const topicIdx = headerRow.indexOf('Topic')+1;

  const subTopicIdx = headerRow.indexOf('Sub-Topic');
  //const subTopicIdx = headerRow.indexOf('Sub-Topic')+1;

  //const reviewDateIdx = headerRow.indexOf('Review Date') + 1;
  const sessionDateIdx = headerRow.indexOf('Session Date') + 1;
  //Added new columns
  //const reviewTATIdx = headerRow.indexOf('Review TAT') + 1;
  //const discussionTATIdx = headerRow.indexOf('Discussion TAT') + 1;
  const accountNumIdx = headerRow.indexOf('Account number') + 1;
 // const boardIdx = headerRow.indexOf('Board#');
  const boardIdx = headerRow.indexOf('Board#')+1;

  const modeIdx = headerRow.indexOf('Mode') + 1;
  const audioIdx = headerRow.indexOf('Audio') + 1;
  const ratingsIdx = headerRow.indexOf('Rating\n(Negative/Positive/Low)') +1;
  const negReviewReasonIdx = headerRow.indexOf('Reason for negative rating') + 1;
  const clientComplaintsIdx = headerRow.indexOf('Client Complaint') + 1;
  const identyIdx = headerRow.indexOf('SubjectKnowledge_Identify') + 1;
  const breakProcessIdx = headerRow.indexOf('SubjectKnowledge_Break The Process') + 1;
  const explanationIdx = headerRow.indexOf('SubjectKnowledge_Explanation') + 1;
  const encourageIdx = headerRow.indexOf('Tutoring_Encourage') + 1;
  const tutoringFlowIdx = headerRow.indexOf('Tutoring_Session Flow') + 1;
  const socraticIdx = headerRow.indexOf('Tutoring_Socratic') + 1;
  const greetingIdx = headerRow.indexOf('Admin_Greeting/ closing') + 1;
  const policiesIdx = headerRow.indexOf('Admin_Client policies') + 1;
  const englishFlowIdx = headerRow.indexOf('Communication_English') + 1;
  const effectiveFlowIdx = headerRow.indexOf('Communication_Effectiveness') + 1;
  const subjectKnowledgePercentIdx = headerRow.indexOf('Subject Knowledge') + 1;
  const tutoringPercentIdx = headerRow.indexOf('Tutoring') + 1;
  const adminPercentIdx = headerRow.indexOf('Admin') + 1;
  const communicationPercentIdx = headerRow.indexOf('Communication') + 1;
  const averageIdx = headerRow.indexOf('Average') + 1;
  const timeInSessionIdx = headerRow.indexOf('Session Time (in minutes)') + 1;
  const commentsIdx = headerRow.indexOf("Comments") + 1;
  const discussionIdx = headerRow.indexOf("Discussion") + 1;
 
  //new columns
  //const discussionDateIdx = headerRow.indexOf("Discussion Date") + 1;
  //const discussionDurationIdx = headerRow.indexOf("Discussion Duration (Min)") + 1;

  const studentCommentsIdx = headerRow.indexOf("Student's Comments") + 1;
  //const lowrateScoreIdx = headerRow.indexOf("Score of low rated sessions") + 1;
  const netTutorIdx = headerRow.indexOf("NetTutor Client Ratings (Out of Five)") + 1;
  //const totalMinsIdx = headerRow.indexOf("Total Hours (In Decimals)") + 1;
  //new column
  const inputToTrainingIdx = headerRow.indexOf("Input to Training team") + 1;
  const totalMinsIdx = headerRow.indexOf("Review Time (Min)") + 1;
  const addIdx = headerRow.indexOf('Add?') + 1;

  const backendsmeNames = getDataFromColName(backendSheet, backendMap, headerRowIdx, 'SME Name');
  const backendclientNames = getDataFromColName(backendSheet, backendMap, headerRowIdx, 'Client');
  const backendSubjects = getDataFromColName(backendSheet, backendMap, headerRowIdx, 'Subject');
  const backendTopics = getDataFromColName(backendSheet, backendMap, headerRowIdx, 'Topic');
  const backendSubTopics = getDataFromColName(backendSheet, backendMap, headerRowIdx, 'Sub-Topic');
  const backendaccountNumbers = getDataFromColName(backendSheet, backendMap, headerRowIdx, 'Account number');
  const backendmodes = getDataFromColName(backendSheet, backendMap, headerRowIdx, 'Mode');
  const backendaudio = getDataFromColName(backendSheet, backendMap, headerRowIdx, 'Audio');
  const backendratings = getDataFromColName(backendSheet, backendMap, headerRowIdx, 'Rating\n(Negative/Positive/Low)');
  const backendnegReviews = getDataFromColName(backendSheet, backendMap, headerRowIdx, 'Reason for negative rating');
  const backendclientComplaints = getDataFromColName(backendSheet, backendMap, headerRowIdx, 'Client Complaint');
  const backendIdentity = getDataFromColName(backendSheet, backendMap, headerRowIdx, "SubjectKnowledge_Identify");
  const breakProcess = getDataFromColName(backendSheet, backendMap, headerRowIdx, "SubjectKnowledge_Break The Process");
  const explanation = getDataFromColName(backendSheet, backendMap, headerRowIdx, "SubjectKnowledge_Explanation");
  const encourage = getDataFromColName(backendSheet, backendMap, headerRowIdx, "Tutoring_Encourage");
  const sessionFlow = getDataFromColName(backendSheet, backendMap, headerRowIdx, "Tutoring_Session Flow");
  const socratic = getDataFromColName(backendSheet, backendMap, headerRowIdx, "Tutoring_Socratic");
  const greeting = getDataFromColName(backendSheet, backendMap, headerRowIdx, "Admin_Greeting/ closing");
  const policies = getDataFromColName(backendSheet, backendMap, headerRowIdx, "Admin_Client policies");
  const commEnglish = getDataFromColName(backendSheet, backendMap, headerRowIdx, "Communication_English");
  const commEffectiveness = getDataFromColName(backendSheet, backendMap, headerRowIdx, "Communication_Effectiveness");
  const backendDiscussion = getDataFromColName(backendSheet, backendMap, headerRowIdx, 'Discussion');
  //console.log("Discussion:-",backendDiscussion)
  //const backendLowRatedScores = getDataFromColName(backendSheet, backendMap, headerRowIdx, 'Score of low rated sessions');
  const backendNetTutor = getDataFromColName(backendSheet, backendMap, headerRowIdx, 'NetTutor Client Ratings (Out of Five)');
  //new column
  const backendInputToTraining = ["-", ...getDataFromColName(backendSheet, backendMap, headerRowIdx, 'Input to Training team') .getValues().flat().filter(String)];

  setDropdownValues(sheet, backendsmeNames, smeNameIdx, numCols);
  setDropdownValues(sheet, backendclientNames, clientIdx, numCols);
  setDropdownValues(sheet, backendSubjects, subjectIdx, numCols);
  //setDropdownValues(sheet, backendTopics, topicIdx, numCols);
 // setDropdownValues(sheet, backendSubTopics, subTopicIdx, numCols);
  setDateValidationRule(sheet, sessionDateIdx, numCols);
  setDropdownValues(sheet, backendaccountNumbers, accountNumIdx, numCols);
  setDropdownValues(sheet, backendmodes, modeIdx, numCols);
  setDropdownValues(sheet, backendaudio, audioIdx, numCols);
  setDropdownValues(sheet, backendratings, ratingsIdx, numCols);
  setDropdownValues(sheet, backendnegReviews, negReviewReasonIdx, numCols);
  setDropdownValues(sheet, backendclientComplaints, clientComplaintsIdx, numCols);
  setDropdownValues(sheet, backendIdentity, identyIdx, numCols);
  setDropdownValues(sheet, breakProcess, breakProcessIdx, numCols);
  setDropdownValues(sheet, explanation, explanationIdx, numCols);
  setDropdownValues(sheet, encourage, encourageIdx, numCols);
  setDropdownValues(sheet, sessionFlow, tutoringFlowIdx, numCols);
  setDropdownValues(sheet, socratic, socraticIdx, numCols);
  setDropdownValues(sheet, greeting, greetingIdx, numCols);
  setDropdownValues(sheet, policies, policiesIdx, numCols);
  setDropdownValues(sheet, commEnglish, englishFlowIdx, numCols);
  setDropdownValues(sheet, commEffectiveness, effectiveFlowIdx, numCols);
  

  let customValues = {
    "Excellent": 1,
    "Acceptable": 2/3,
    "Needs Improvement": 1/3,
    "Unacceptable": 0
  };

  setScores(sheet, subjectKnowledgePercentIdx, identyIdx, breakProcessIdx, explanationIdx, customValues);
  setScores(sheet, tutoringPercentIdx, encourageIdx, tutoringFlowIdx, socraticIdx, customValues);
  
  customValues = {
    "Excellent": 1,
    "Acceptable": 1/2,
    "Needs Improvement": 1/4,
    "Unacceptable": 0
  };

  setScores(sheet, adminPercentIdx, greetingIdx, policiesIdx, null, customValues);
  setScores(sheet, communicationPercentIdx, englishFlowIdx, effectiveFlowIdx, null, customValues);
  setAverage(sheet, subjectKnowledgePercentIdx, tutoringPercentIdx, adminPercentIdx, communicationPercentIdx, averageIdx,numCols);

  //new columns 
  //setReviewAndDiscussionTAT(sheet, sessionDateIdx, reviewDateIdx, discussionDateIdx, reviewTATIdx, discussionTATIdx, numCols);


  setDropdownValues(sheet, backendDiscussion, discussionIdx, numCols);
  //setDropdownValuesList(sheet, [0,1,2,3,4,5], lowrateScoreIdx, numCols);
  setDropdownValuesList(sheet, [0,1,2,3,4,5], netTutorIdx, numCols).setNumberFormat('0');
  //new column 
  setDropdownValuesList(sheet, backendInputToTraining, inputToTrainingIdx, numCols);
  setNumberValidation(sheet, totalMinsIdx, numCols, 'greaterThanOrEqualTo', { min: 0});
  setNumberValidation(sheet, timeInSessionIdx, numCols, 'greaterThanOrEqualTo', { min: 0});
  
  //new columns
  // if (reviewDateIdx > 0) {
  //   setDateValidationRule(sheet, reviewDateIdx, numCols);
  // }

  // if (discussionDateIdx > 0) {
  //   setDateValidationRule(sheet, discussionDateIdx, numCols);
  // }

  // if (discussionDurationIdx > 0) {
  //   setNumberValidation(sheet, discussionDurationIdx, numCols, 'greaterThanOrEqualTo', { min: 0 });
  // }

  setCheckBoxes(sheet, addIdx, numCols);

}


function setDropdownValues(sheet, array, idx, numCols){
  var column = sheet.getRange(numCols, idx, sheet.getLastRow()-1, 1);
  column.setNumberFormat("@")
  var rule = SpreadsheetApp.newDataValidation()
                           .requireValueInRange(array,true) 
                           .setAllowInvalid(false).build();
  return column.clearDataValidations().setDataValidation(rule);
}


function setDropdownValuesList(sheet, array, idx, numCols){
  var column = sheet.getRange(numCols, idx, sheet.getLastRow()-1, 1);
  column.setNumberFormat("@")
  var rule = SpreadsheetApp.newDataValidation()
                           .requireValueInList(array, true)
                           .setAllowInvalid(false).build();
  return column.clearDataValidations().setDataValidation(rule);
}


function setCheckBoxes(sheet, idx, numCols){
  const column = sheet.getRange(numCols, idx, sheet.getLastRow()-1, 1)
  column.insertCheckboxes().setValue(false);
}


function setDateValidationRule(sheet, columnNumber, startRow) {
  var column = sheet.getRange(startRow, columnNumber, sheet.getLastRow());
  // Create a data validation rule for dates
  var rule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .build();

  // Apply the data validation rule to the specified column
  column.setDataValidation(rule);
  column.setNumberFormat("dd-mmm-yy"); 
}


function setNumberValidation(sheet, columnNumber, startRow, validationType, options) {
  var column = sheet.getRange(startRow, columnNumber, sheet.getLastRow());

  // Create a data validation rule based on the validationType
  var rule = SpreadsheetApp.newDataValidation();

  switch (validationType) {
    case 'between':
      if (options && options.min !== undefined && options.max !== undefined) {
        rule.requireNumberBetween(options.min, options.max);
      }
      break;
      
    case 'greaterThanOrEqualTo':
      if (options && options.min !== undefined) {
        rule.requireNumberGreaterThanOrEqualTo(options.min);
      }
      break;

    case 'lessThanOrEqualTo':
      if(options && options.max !== undefined)
        rule.requireNumberLessThanOrEqualTo(options.max);
      break;

    default:
      throw new Error('Invalid validationType');
  }

  // Set common data validation properties
  rule.setAllowInvalid(false);

  // Build and apply the data validation rule to the specified column
  column.setDataValidation(rule.build());
}

//for new columns 
// function setReviewAndDiscussionTAT(sheet, sessionDateIdx, reviewDateIdx, discussionDateIdx, reviewTATIdx, discussionTATIdx, startRow = 2) {
//   const numRows = sheet.getLastRow() - startRow + 1;

//   const reviewTATFormulas = [];
//   const discussionTATFormulas = [];

//   for (let i = 0; i < numRows; i++) {
//     const row = startRow + i;

//     const sessionDateCell = sheet.getRange(row, sessionDateIdx).getA1Notation();
//     const reviewDateCell = sheet.getRange(row, reviewDateIdx).getA1Notation();
//     const discussionDateCell = sheet.getRange(row, discussionDateIdx).getA1Notation();

//     const reviewTATFormula = `=IF(AND(ISDATE(${sessionDateCell}), ISDATE(${reviewDateCell})), (${sessionDateCell}-${reviewDateCell}) , "")`;
//     const discussionTATFormula = `=IF(AND(ISDATE(${reviewDateCell}), ISDATE(${discussionDateCell})), (${reviewDateCell}-${discussionDateCell}), "")`;

//     reviewTATFormulas.push([reviewTATFormula]);
//     discussionTATFormulas.push([discussionTATFormula]);
//   }

//   sheet.getRange(startRow, reviewTATIdx, numRows, 1).setFormulas(reviewTATFormulas);
//   sheet.getRange(startRow, discussionTATIdx, numRows, 1).setFormulas(discussionTATFormulas);
// }


function setScores(sheet, actualColumnIdx, col1Idx, col2Idx, col3Idx = null, customValues = {}) {
  const numRows = sheet.getLastRow();
  const formulas = [];

  for (let i = 2; i <= numRows; i++) { // Start from row 2 (assuming headers in row 1)
      let range1 = sheet.getRange(i, col1Idx)
      let val1 = range1.getA1Notation();
      let range2 = sheet.getRange(i, col2Idx)
      let val2 = range2.getA1Notation();
      let range3 = col3Idx ? sheet.getRange(i, col3Idx) : null;
      let val3 = range3 ? range3.getA1Notation() : "";

      // Initialize the "actual" column to 0 initially
      let formula = `=IF(COUNTA(${val1}, ${val2}, ${val3}) > 0, (` + 
                    `IF(${val1}="Excellent", ${customValues["Excellent"] || 0}, IF(${val1}="Acceptable", ${customValues["Acceptable"] || 0}, IF(${val1}="Needs Improvement", ${customValues["Needs Improvement"] || 0}, IF(${val1}="Unacceptable", ${customValues["Unacceptable"] || 0}, "")))) + ` +
                    `IF(${val2}="Excellent", ${customValues["Excellent"] || 0}, IF(${val2}="Acceptable", ${customValues["Acceptable"] || 0}, IF(${val2}="Needs Improvement", ${customValues["Needs Improvement"] || 0}, IF(${val2}="Unacceptable", ${customValues["Unacceptable"] || 0}, "")))) + `;

      if (col3Idx !== null) {
        formula += `IF(${val3}="Excellent", ${customValues["Excellent"] || 0}, IF(${val3}="Acceptable", ${customValues["Acceptable"] || 0}, IF(${val3}="Needs Improvement", ${customValues["Needs Improvement"] || 0}, IF(${val3}="Unacceptable", ${customValues["Unacceptable"] || 0}, ""))))`;
      } else {
        formula = formula.slice(0, -2); // Remove the trailing " + "
      }

      formula += `)/COUNTA(${val1}, ${val2}, ${val3}), "")`;

      formulas.push([formula]);
  }
  // Set the formulas for the entire "actual" column
  const actualColumnRange = sheet.getRange(2, actualColumnIdx, numRows - 1, 1);
  actualColumnRange.setFormulas(formulas);
}


function setAverage(sheet, val1Idx, val2Idx, val3Idx, val4Idx, columnIdx, numCols) {
  const numRows = sheet.getLastRow();
  const formulas = [];

  const val1Range = sheet.getRange(numCols, val1Idx, numRows - numCols + 1, 1);
  const val2Range = sheet.getRange(numCols, val2Idx, numRows - numCols + 1, 1);
  const val3Range = sheet.getRange(numCols, val3Idx, numRows - numCols + 1, 1);
  const val4Range = sheet.getRange(numCols, val4Idx, numRows - numCols + 1, 1);

  for (let i = 1; i <= numRows - numCols + 1; i++) {
    const val1A1 = val1Range.getCell(i, 1).getA1Notation();
    const val2A1 = val2Range.getCell(i, 1).getA1Notation();
    const val3A1 = val3Range.getCell(i, 1).getA1Notation();
    const val4A1 = val4Range.getCell(i, 1).getA1Notation();

    // Check for division by zero
    // const formula = `=IF(COUNT(${val1A1}:${val4A1})<4, "", IF(SUM(${val1A1}:${val4A1})=0, 0, AVERAGE(${val1A1}:${val4A1})))`;

   const formula = `=IF(COUNT(${val1A1},${val2A1},${val3A1},${val4A1})=0, "", AVERAGEIF({${val1A1},${val2A1},${val3A1},${val4A1}}, "<>"))`;

    formulas.push([formula]);
  }


  // Set the formulas for the entire column at once
  const columnRange = sheet.getRange(numCols, columnIdx, numRows - numCols + 1, 1);
  columnRange.setNumberFormat("0.00%");
  columnRange.setFormulas(formulas);
}


function setMaxValue(sheet, array, idx, numCols){
  const maxVal = Math.max.apply(this, array.filter(e => typeof e === 'number'));
  sheet.getRange(numCols, idx, sheet.getLastRow()-1, 1).setValue(maxVal)
}


function getColumnIndices(columnNames) {
  const columnIndexMap = {};
  for (const columnName of columnNames) {
    columnIndexMap[columnName] = columnNames.indexOf(columnName);
  }
  return columnIndexMap;
}


function getDataFromColName(sheet, map, headerRowIdx, colName){
  return sheet.getRange(headerRowIdx+1, map[colName]+1, sheet.getLastRow()-headerRowIdx, 1);//.getValues().flat().filter(r => r !== "");
}


function getDataFromThreeColumns(sheet, map, headerRowIdx, startColName){
  return sheet.getRange(headerRowIdx+1, map[startColName]+1, sheet.getLastRow(), 3).getValues().filter(r=>Boolean(r));
}





// --------------------------old code---------------------------
// function setDropdownsQAReviewAdd(sheet, viewSheet, backendSheet, numCols=2){


//   const dataRange = viewSheet.getRange(5, 1, viewSheet.getLastRow() - 4, viewSheet.getLastColumn());
//   const viewHeaders = dataRange.getValues()[0]; 

//   const viewIndices = {
//     srNoIdx: viewHeaders.indexOf("#") + 1,
//     commentsIdx: viewHeaders.indexOf("Comments") + 1,
//     studentCommentsIdx: viewHeaders.indexOf("Student's Comments") + 1,
//    // totalMinsIdx: viewHeaders.indexOf("Total Hours (In Decimals)") + 1,
//     totalMinsIdx: viewHeaders.indexOf("Review Time (Min)") + 1,
//     discussionIdx:viewHeaders.indexOf("Discussion") + 1,
//     lowrateScoreIdx : viewHeaders.indexOf("Score of low rated sessions") +1 ,
//     netTutorIdx: viewHeaders.indexOf("NetTutor Client Ratings (Out of Five)") + 1,
//   }

//   const wholeSheet = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();

//   const headerRow = wholeSheet[0], dataRows = wholeSheet.slice(1);
//   const headerRowIdx = 4
//   const backendColumnNames = backendSheet.getRange(headerRowIdx, 1, 1, backendSheet.getLastColumn()).getValues().flat();

//   const colIdxMap = getColumnIndices(headerRow);
//   const backendMap = getColumnIndices(backendColumnNames);

//   const srNoIdx = headerRow.indexOf('#') + 1;
//   const smeNameIdx = headerRow.indexOf('SME Name') + 1;
//   const clientIdx = headerRow.indexOf('Client') + 1;
//   const subjectIdx = headerRow.indexOf('Subject') + 1;
//   const topicIdx = headerRow.indexOf('Topic');
//   const subTopicIdx = headerRow.indexOf('Sub-Topic');
//   const sessionDateIdx = headerRow.indexOf('Session Date') + 1;
//   //new column
//   const reviewDateIdx = headerRow.indexOf("Review Date") + 1;
//   const reviewTATIdx = headerRow.indexOf("Review TAT") + 1;

//   const accountNumIdx = headerRow.indexOf('Account number') + 1;
//   const boardIdx = headerRow.indexOf('Board#');
//   const modeIdx = headerRow.indexOf('Mode') + 1;
//   const audioIdx = headerRow.indexOf('Audio') + 1;
//   const ratingsIdx = headerRow.indexOf('Rating\n(Negative/Positive/Low)') +1;
//   const negReviewReasonIdx = headerRow.indexOf('Reason for negative rating') + 1;
//   const clientComplaintsIdx = headerRow.indexOf('Client Complaint') + 1;
//   const identyIdx = headerRow.indexOf('SubjectKnowledge_Identify') + 1;
//   const breakProcessIdx = headerRow.indexOf('SubjectKnowledge_Break The Process') + 1;
//   const explanationIdx = headerRow.indexOf('SubjectKnowledge_Explanation') + 1;
//   const encourageIdx = headerRow.indexOf('Tutoring_Encourage') + 1;
//   const tutoringFlowIdx = headerRow.indexOf('Tutoring_Session Flow') + 1;
//   const socraticIdx = headerRow.indexOf('Tutoring_Socratic') + 1;
//   const greetingIdx = headerRow.indexOf('Admin_Greeting/ closing') + 1;
//   const policiesIdx = headerRow.indexOf('Admin_Client policies') + 1;
//   const englishFlowIdx = headerRow.indexOf('Communication_English') + 1;
//   const effectiveFlowIdx = headerRow.indexOf('Communication_Effectiveness') + 1;
//   const subjectKnowledgePercentIdx = headerRow.indexOf('Subject Knowledge') + 1;
//   const tutoringPercentIdx = headerRow.indexOf('Tutoring') + 1;
//   const adminPercentIdx = headerRow.indexOf('Admin') + 1;
//   const communicationPercentIdx = headerRow.indexOf('Communication') + 1;
//   const averageIdx = headerRow.indexOf('Average') + 1;
//   const timeInSessionIdx = headerRow.indexOf('Session Time (in minutes)') + 1;
//   const commentsIdx = headerRow.indexOf("Comments") + 1;
//   const discussionIdx = headerRow.indexOf("Discussion") + 1;
//   //new columns
//   const discussionDateIdx = headerRow.indexOf("Discussion Date") + 1;
//   const discussionDurationIdx = headerRow.indexOf("Discussion Duration (Min)") + 1;

//   const studentCommentsIdx = headerRow.indexOf("Student's Comments") + 1;
//   const lowrateScoreIdx = headerRow.indexOf("Score of low rated sessions") + 1;
//   const netTutorIdx = headerRow.indexOf("NetTutor Client Ratings (Out of Five)") + 1;
//   //const totalMinsIdx = headerRow.indexOf("Total Hours (In Decimals)") + 1;
//   const totalMinsIdx = headerRow.indexOf("Review Time (Min)") + 1;
//   const addIdx = headerRow.indexOf('Add?') + 1;

//   const backendsmeNames = getDataFromColName(backendSheet, backendMap, headerRowIdx, 'SME Name');
//   const backendclientNames = getDataFromColName(backendSheet, backendMap, headerRowIdx, 'Client');
//   const backendSubjects = getDataFromColName(backendSheet, backendMap, headerRowIdx, 'Subject');
//   const backendaccountNumbers = getDataFromColName(backendSheet, backendMap, headerRowIdx, 'Account number');
//   const backendmodes = getDataFromColName(backendSheet, backendMap, headerRowIdx, 'Mode');
//   const backendaudio = getDataFromColName(backendSheet, backendMap, headerRowIdx, 'Audio');
//   const backendratings = getDataFromColName(backendSheet, backendMap, headerRowIdx, 'Rating\n(Negative/Positive/Low)');
//   const backendnegReviews = getDataFromColName(backendSheet, backendMap, headerRowIdx, 'Reason for negative rating');
//   const backendclientComplaints = getDataFromColName(backendSheet, backendMap, headerRowIdx, 'Client Complaint');
//   const backendIdentity = getDataFromColName(backendSheet, backendMap, headerRowIdx, "SubjectKnowledge_Identify");
//   const breakProcess = getDataFromColName(backendSheet, backendMap, headerRowIdx, "SubjectKnowledge_Break The Process");
//   const explanation = getDataFromColName(backendSheet, backendMap, headerRowIdx, "SubjectKnowledge_Explanation");
//   const encourage = getDataFromColName(backendSheet, backendMap, headerRowIdx, "Tutoring_Encourage");
//   const sessionFlow = getDataFromColName(backendSheet, backendMap, headerRowIdx, "Tutoring_Session Flow");
//   const socratic = getDataFromColName(backendSheet, backendMap, headerRowIdx, "Tutoring_Socratic");
//   const greeting = getDataFromColName(backendSheet, backendMap, headerRowIdx, "Admin_Greeting/ closing");
//   const policies = getDataFromColName(backendSheet, backendMap, headerRowIdx, "Admin_Client policies");
//   const commEnglish = getDataFromColName(backendSheet, backendMap, headerRowIdx, "Communication_English");
//   const commEffectiveness = getDataFromColName(backendSheet, backendMap, headerRowIdx, "Communication_Effectiveness");
//   const backendDiscussion = getDataFromColName(backendSheet, backendMap, headerRowIdx, 'Discussion');
//   const backendLowRatedScores = getDataFromColName(backendSheet, backendMap, headerRowIdx, 'Score of low rated sessions');
//   const backendNetTutor = getDataFromColName(backendSheet, backendMap, headerRowIdx, 'NetTutor Client Ratings (Out of Five)');
  
//   setDropdownValues(sheet, backendsmeNames, smeNameIdx, numCols);
//   setDropdownValues(sheet, backendclientNames, clientIdx, numCols);
//   setDropdownValues(sheet, backendSubjects, subjectIdx, numCols);
//   setDateValidationRule(sheet, sessionDateIdx, numCols);
//   setDropdownValues(sheet, backendaccountNumbers, accountNumIdx, numCols);
//   setDropdownValues(sheet, backendmodes, modeIdx, numCols);
//   setDropdownValues(sheet, backendaudio, audioIdx, numCols);
//   setDropdownValues(sheet, backendratings, ratingsIdx, numCols);
//   setDropdownValues(sheet, backendnegReviews, negReviewReasonIdx, numCols);
//   setDropdownValues(sheet, backendclientComplaints, clientComplaintsIdx, numCols);
//   setDropdownValues(sheet, backendIdentity, identyIdx, numCols);
//   setDropdownValues(sheet, breakProcess, breakProcessIdx, numCols);
//   setDropdownValues(sheet, explanation, explanationIdx, numCols);
//   setDropdownValues(sheet, encourage, encourageIdx, numCols);
//   setDropdownValues(sheet, sessionFlow, tutoringFlowIdx, numCols);
//   setDropdownValues(sheet, socratic, socraticIdx, numCols);
//   setDropdownValues(sheet, greeting, greetingIdx, numCols);
//   setDropdownValues(sheet, policies, policiesIdx, numCols);
//   setDropdownValues(sheet, commEnglish, englishFlowIdx, numCols);
//   setDropdownValues(sheet, commEffectiveness, effectiveFlowIdx, numCols);

//   let customValues = {
//     "Excellent": 1,
//     "Acceptable": 2/3,
//     "Needs Improvement": 1/3,
//     "Unacceptable": 0
//   };

//   setScores(sheet, subjectKnowledgePercentIdx, identyIdx, breakProcessIdx, explanationIdx, customValues);
//   setScores(sheet, tutoringPercentIdx, encourageIdx, tutoringFlowIdx, socraticIdx, customValues);
  
//   customValues = {
//     "Excellent": 1,
//     "Acceptable": 1/2,
//     "Needs Improvement": 1/4,
//     "Unacceptable": 0
//   };

//   setScores(sheet, adminPercentIdx, greetingIdx, policiesIdx, null, customValues);
//   setScores(sheet, communicationPercentIdx, englishFlowIdx, effectiveFlowIdx, null, customValues);
//   setAverage(sheet, subjectKnowledgePercentIdx, tutoringPercentIdx, adminPercentIdx, communicationPercentIdx, averageIdx,numCols);

//   setDropdownValues(sheet, backendDiscussion, discussionIdx, numCols);
//   setDropdownValuesList(sheet, [0,1,2,3,4,5], lowrateScoreIdx, numCols);
//   setDropdownValuesList(sheet, [0,1,2,3,4,5], netTutorIdx, numCols).setNumberFormat('0');
//   setNumberValidation(sheet, totalMinsIdx, numCols, 'greaterThanOrEqualTo', { min: 0});
//   setNumberValidation(sheet, timeInSessionIdx, numCols, 'greaterThanOrEqualTo', { min: 0});
  
//   //new columns
//   if (discussionDateIdx > 0) {
//     setDateValidationRule(sheet, discussionDateIdx, numCols);
//   }

//   if (discussionDurationIdx > 0) {
//     setNumberValidation(sheet, discussionDurationIdx, numCols, 'greaterThanOrEqualTo', { min: 0 });
//   }

//   setCheckBoxes(sheet, addIdx, numCols);

//   //new column logic 
// // Set Review TAT formula: Session Date - Review Date
// if (reviewDateIdx > 0 && reviewTATIdx > 0) {
//   const numRows = sheet.getLastRow();
//   const formulas = [];

//   for (let i = numCols; i <= numRows; i++) {
//     const sessionCell = sheet.getRange(i, sessionDateIdx).getA1Notation();
//     const reviewCell = sheet.getRange(i, reviewDateIdx).getA1Notation();
//     formulas.push([`=IF(AND(ISDATE(${sessionCell}), ISDATE(${reviewCell})), ${sessionCell}-${reviewCell}, "")`]);
//   }

//   const formulaRange = sheet.getRange(numCols, reviewTATIdx, numRows - numCols + 1, 1);
//   formulaRange.setFormulas(formulas);
//   formulaRange.setNumberFormat("0.0");
// }

// // Set Discussion TAT formula: Review Date - Discussion Date
// if (reviewDateIdx > 0 && discussionDateIdx > 0 && discussionTATIdx > 0) {
//   const numRows = sheet.getLastRow();
//   const formulas = [];

//   for (let i = numCols; i <= numRows; i++) {
//     const reviewCell = sheet.getRange(i, reviewDateIdx).getA1Notation();
//     const discussionCell = sheet.getRange(i, discussionDateIdx).getA1Notation();
//     formulas.push([`=IF(AND(ISDATE(${reviewCell}), ISDATE(${discussionCell})), ${reviewCell}-${discussionCell}, "")`]);
//   }

//   const formulaRange = sheet.getRange(numCols, discussionTATIdx, numRows - numCols + 1, 1);
//   formulaRange.setFormulas(formulas);
//   formulaRange.setNumberFormat("0.0");
// }

// }


// function setDropdownValues(sheet, array, idx, numCols){
//   var column = sheet.getRange(numCols, idx, sheet.getLastRow()-1, 1);
//   column.setNumberFormat("@")
//   var rule = SpreadsheetApp.newDataValidation()
//                            .requireValueInRange(array, true)
//                            .setAllowInvalid(false).build();
//   return column.clearDataValidations().setDataValidation(rule);
// }



// function setDropdownValuesList(sheet, array, idx, numCols){
//   var column = sheet.getRange(numCols, idx, sheet.getLastRow()-1, 1);
//   column.setNumberFormat("@")
//   var rule = SpreadsheetApp.newDataValidation()
//                            .requireValueInList(array, true)
//                            .setAllowInvalid(false).build();
//   return column.clearDataValidations().setDataValidation(rule);
// }

// function setCheckBoxes(sheet, idx, numCols){
//   const column = sheet.getRange(numCols, idx, sheet.getLastRow()-1, 1)
//   column.insertCheckboxes().setValue(false);
// }

// function setDateValidationRule(sheet, columnNumber, startRow) {
//   var column = sheet.getRange(startRow, columnNumber, sheet.getLastRow());
//   // Create a data validation rule for dates
//   var rule = SpreadsheetApp.newDataValidation()
//     .requireDate()
//     .setAllowInvalid(false)
//     .build();

//   // Apply the data validation rule to the specified column
//   column.setDataValidation(rule);
//   column.setNumberFormat("dd-mmm-yy"); 
// }


// function setNumberValidation(sheet, columnNumber, startRow, validationType, options) {
//   var column = sheet.getRange(startRow, columnNumber, sheet.getLastRow());

//   // Create a data validation rule based on the validationType
//   var rule = SpreadsheetApp.newDataValidation();

//   switch (validationType) {
//     case 'between':
//       if (options && options.min !== undefined && options.max !== undefined) {
//         rule.requireNumberBetween(options.min, options.max);
//       }
//       break;
      
//     case 'greaterThanOrEqualTo':
//       if (options && options.min !== undefined) {
//         rule.requireNumberGreaterThanOrEqualTo(options.min);
//       }
//       break;

//     case 'lessThanOrEqualTo':
//       if(options && options.max !== undefined)
//         rule.requireNumberLessThanOrEqualTo(options.max);
//       break;

//     default:
//       throw new Error('Invalid validationType');
//   }

//   // Set common data validation properties
//   rule.setAllowInvalid(false);

//   // Build and apply the data validation rule to the specified column
//   column.setDataValidation(rule.build());
// }

// function setScores(sheet, actualColumnIdx, col1Idx, col2Idx, col3Idx = null, customValues = {}) {
//   const numRows = sheet.getLastRow();
//   const formulas = [];

//   for (let i = 2; i <= numRows; i++) { // Start from row 2 (assuming headers in row 1)
//       let range1 = sheet.getRange(i, col1Idx)
//       let val1 = range1.getA1Notation();
//       let range2 = sheet.getRange(i, col2Idx)
//       let val2 = range2.getA1Notation();
//       let range3 = col3Idx ? sheet.getRange(i, col3Idx) : null;
//       let val3 = range3 ? range3.getA1Notation() : "";

//       // Initialize the "actual" column to 0 initially
//       let formula = `=IF(COUNTA(${val1}, ${val2}, ${val3}) > 0, (` + 
//                     `IF(${val1}="Excellent", ${customValues["Excellent"] || 0}, IF(${val1}="Acceptable", ${customValues["Acceptable"] || 0}, IF(${val1}="Needs Improvement", ${customValues["Needs Improvement"] || 0}, IF(${val1}="Unacceptable", ${customValues["Unacceptable"] || 0}, "")))) + ` +
//                     `IF(${val2}="Excellent", ${customValues["Excellent"] || 0}, IF(${val2}="Acceptable", ${customValues["Acceptable"] || 0}, IF(${val2}="Needs Improvement", ${customValues["Needs Improvement"] || 0}, IF(${val2}="Unacceptable", ${customValues["Unacceptable"] || 0}, "")))) + `;

//       if (col3Idx !== null) {
//         formula += `IF(${val3}="Excellent", ${customValues["Excellent"] || 0}, IF(${val3}="Acceptable", ${customValues["Acceptable"] || 0}, IF(${val3}="Needs Improvement", ${customValues["Needs Improvement"] || 0}, IF(${val3}="Unacceptable", ${customValues["Unacceptable"] || 0}, ""))))`;
//       } else {
//         formula = formula.slice(0, -2); // Remove the trailing " + "
//       }

//       formula += `)/COUNTA(${val1}, ${val2}, ${val3}), "")`;

//       formulas.push([formula]);
//   }


//   // Set the formulas for the entire "actual" column
//   const actualColumnRange = sheet.getRange(2, actualColumnIdx, numRows - 1, 1);
//   actualColumnRange.setFormulas(formulas);
// }

// function setAverage(sheet, val1Idx, val2Idx, val3Idx, val4Idx, columnIdx, numCols) {
//   const numRows = sheet.getLastRow();
//   const formulas = [];

//   const val1Range = sheet.getRange(numCols, val1Idx, numRows - numCols + 1, 1);
//   const val2Range = sheet.getRange(numCols, val2Idx, numRows - numCols + 1, 1);
//   const val3Range = sheet.getRange(numCols, val3Idx, numRows - numCols + 1, 1);
//   const val4Range = sheet.getRange(numCols, val4Idx, numRows - numCols + 1, 1);

//   for (let i = 1; i <= numRows - numCols + 1; i++) {
//     const val1A1 = val1Range.getCell(i, 1).getA1Notation();
//     const val2A1 = val2Range.getCell(i, 1).getA1Notation();
//     const val3A1 = val3Range.getCell(i, 1).getA1Notation();
//     const val4A1 = val4Range.getCell(i, 1).getA1Notation();

//     // Check for division by zero
//     const formula = `=IF(COUNT(${val1A1}:${val4A1})<4, "", IF(SUM(${val1A1}:${val4A1})=0, 0, AVERAGE(${val1A1}:${val4A1})))`;
//     formulas.push([formula]);
//   }

//   // Set the formulas for the entire column at once
//   const columnRange = sheet.getRange(numCols, columnIdx, numRows - numCols + 1, 1);
//   columnRange.setNumberFormat("0.00%");
//   columnRange.setFormulas(formulas);
// }


// function setMaxValue(sheet, array, idx, numCols){
//   const maxVal = Math.max.apply(this, array.filter(e => typeof e === 'number'));
//   sheet.getRange(numCols, idx, sheet.getLastRow()-1, 1).setValue(maxVal)
// }



// function getColumnIndices(columnNames) {
//   const columnIndexMap = {};
//   for (const columnName of columnNames) {
//     columnIndexMap[columnName] = columnNames.indexOf(columnName);
//   }
//   return columnIndexMap;
// }

// function getDataFromColName(sheet, map, headerRowIdx, colName){
//   return sheet.getRange(headerRowIdx+1, map[colName]+1, sheet.getLastRow()-headerRowIdx, 1);//.getValues().flat().filter(r => r !== "");
// }

// function getDataFromThreeColumns(sheet, map, headerRowIdx, startColName){
//   return sheet.getRange(headerRowIdx+1, map[startColName]+1, sheet.getLastRow(), 3).getValues().filter(r=>Boolean(r));

// }




