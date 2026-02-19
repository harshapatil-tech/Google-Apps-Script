// function applyCustomFormatting(range, options) {

//   options = options || {};
  
//   var fontSize = options.fontSize || 10;
//   var fontColor = options.fontColor || 'black';
//   var bgColor = options.bgColor || 'white';
//   var fontWeight = options.fontWeight || 'normal'

//   range.setHorizontalAlignment("center")
//       .setVerticalAlignment("middle")
//       .setWrap(true)
//       .setFontFamily("Roboto")
//       .setFontSize(fontSize)
//       .setFontColor(fontColor)
//       .setFontWeight(fontWeight)
//       .setBorder(true, true, true, true, true, true)
//       .setBackground(bgColor);
//   return range;
// };


// function clearRowsBelow(sheet, startRow) {  
//   var lastRow = sheet.getLastRow();
//   var lastColumn = sheet.getLastColumn();
//   var numRows = lastRow - startRow + 1;
  
//   if (numRows > 0) {
//     sheet.getRange(startRow + 1, 1, numRows, lastColumn).clearContent();
//   }
// }





//________________*** OLD FUNCTION ***____________
// function tempFunction(){ 
//   createBackendReviwerSheetByDepartment("https://docs.google.com/spreadsheets/d/17t_fKuXAkXtoVnB4FmCRDIHyJzdiAzIYGbQR0eEdeQM/edit", "Statistics");
//    ////"https://docs.google.com/spreadsheets/d/1PO0lZfZ_f81reDKTFTifwtmw2UslQ0KFvJ3s2s07K88/edit",
//    //https://docs.google.com/spreadsheets/d/1PgZenNQ2H4hRgcb-nzKeTowGOUCf7hbkZFI6ObJnFUs/edit
// }

// function createBackendReviwerSheetByDepartment(url, department){
//   const spreadsheet = SpreadsheetApp.openByUrl(url);
//   const reviwerName = spreadsheet.getName().split("_").slice(2).join(' ');
//   const backendSheet = spreadsheet.getSheetByName('Backend');
//   const qaReviewUpdateSheet = spreadsheet.getSheetByName('QA_Review_Update');
//   const qaReviewAddSheet = spreadsheet.getSheetByName("QA_Review_Add");
  
//   const [topicData, smeData, accountData, otherData] = getMasterDBData(reviwerName, department)
//   const headerRowIdx = 4 
//   clearRowsBelow(backendSheet, headerRowIdx)
//   const headerRow = backendSheet.getRange(headerRowIdx, 1, 1, backendSheet.getLastColumn()).getValues().flat();
//   const smeNameIdx = headerRow.indexOf('SME Name') + 1;
//   const clientIdx = headerRow.indexOf('Client') + 1;
//   const subjectIdx = headerRow.indexOf('Subject') + 1;
//   const topicIdx = headerRow.indexOf('Topic') + 1;
//   const subTopicIdx = headerRow.indexOf('Sub-Topic') + 1;
//   const accNumIdx = headerRow.indexOf('Account number') + 1;
//   const modeIdx = headerRow.indexOf('Mode') + 1;
//   const audioIdx = headerRow.indexOf('Audio') + 1;
//   const ratingIdx = headerRow.indexOf('Rating\n(Negative/Positive/Low)') + 1;
//   const negRatingReasonIdx = headerRow.indexOf('Reason for negative rating') + 1;
//   const clientComplaintIdx = headerRow.indexOf('Client Complaint') + 1;
  
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

//   const discussionIdx = headerRow.indexOf('Discussion') + 1;
//   const lowRateScoreIdx = headerRow.indexOf("Score of low rated sessions") + 1;
//   const netTutorIdx = headerRow.indexOf("NetTutor Client Ratings (Out of Five)") + 1;
//   const candidateListIdx = headerRow.indexOf('Candidate list') + 1;

//   let [clients, mode, audio, ratings, negReviews, clientComplaints, mappings, discussions, identity, breakProcess, explanation, encourage, tutoringFlow, socratic, greetings, policies, englishFlow, effectiveFlow, lowRateScores, netTutor] = otherData;

//   if(smeData.length>0)
//     backendSheet.getRange(headerRowIdx+1, smeNameIdx, smeData.length, 1).setValues(smeData.map(r=>[r]));

//   if(clients.length>0)
//     backendSheet.getRange(headerRowIdx+1, clientIdx, clients.length, 1).setValues(clients.map(r=>[r]));

//   if (topicData.length>0)
//   backendSheet.getRange(headerRowIdx+1, subjectIdx, topicData.length, 3).setValues(topicData);

//   if(accountData.length>0)
//     backendSheet.getRange(headerRowIdx+1, accNumIdx, accountData.length, 1).setValues(accountData.map(r=>[r]))
 
//   if(mode.length>0)
//     backendSheet.getRange(headerRowIdx+1, modeIdx, mode.length, 1).setValues(mode.map(r=>[r]));
 
//   if(audio.length>0)
//     backendSheet.getRange(headerRowIdx+1, audioIdx, audio.length, 1).setValues(audio.map(r=>[r]));
 
//   if(ratings.length>0)
//     backendSheet.getRange(headerRowIdx+1, ratingIdx, ratings.length, 1).setValues(ratings.map(r=>[r]));
 
//   if(negReviews.length>0)
//     backendSheet.getRange(headerRowIdx+1, negRatingReasonIdx, negReviews.length, 1).setValues(negReviews.map(r=>[r]));

//   if(clientComplaints.length>0)
//     backendSheet.getRange(headerRowIdx+1, clientComplaintIdx, clientComplaints.length, 1).setValues(clientComplaints.map(r=>[r]));

//   if(discussions.length>0)
//     backendSheet.getRange(headerRowIdx+1, discussionIdx, discussions.length, 1).setValues(discussions.map(r=>[r]));

//   if(identity.length>0)
//     backendSheet.getRange(headerRowIdx+1, identyIdx, identity.length, 1).setValues(identity.map(r=>[r]));

//   if(breakProcess.length>0)
//     backendSheet.getRange(headerRowIdx+1, breakProcessIdx, breakProcess.length, 1).setValues(breakProcess.map(r=>[r]));

//   if(explanation.length>0)
//     applyCustomFormatting(backendSheet.getRange(headerRowIdx+1, explanationIdx, explanation.length, 1)).setValues(explanation.map(r=>[r]));

//   if(encourage.length>0)
//     backendSheet.getRange(headerRowIdx+1, encourageIdx, encourage.length, 1).setValues(encourage.map(r=>[r]));

//   if(tutoringFlow.length>0)
//     backendSheet.getRange(headerRowIdx+1, tutoringFlowIdx, tutoringFlow.length, 1).setValues(tutoringFlow.map(r=>[r]));

//   if(socratic.length>0)
//     backendSheet.getRange(headerRowIdx+1, socraticIdx, socratic.length, 1).setValues(socratic.map(r=>[r]));

//   if(greetings.length>0)
//     backendSheet.getRange(headerRowIdx+1, greetingIdx, greetings.length, 1).setValues(greetings.map(r=>[r]));

//   if(policies.length>0)
//     backendSheet.getRange(headerRowIdx+1, policiesIdx, policies.length, 1).setValues(policies.map(r=>[r]));

//   if(englishFlow.length>0)
//     backendSheet.getRange(headerRowIdx+1, englishFlowIdx, englishFlow.length, 1).setValues(englishFlow.map(r=>[r]));

//   if(effectiveFlow.length>0)
//     backendSheet.getRange(headerRowIdx+1, effectiveFlowIdx, effectiveFlow.length, 1).setValues(effectiveFlow.map(r=>[r]));

//   if(lowRateScores.length>0)
//     backendSheet.getRange(headerRowIdx+1, lowRateScoreIdx, lowRateScores.length, 1).setValues(lowRateScores.map(r=>[r]));

//   if(netTutor.length>0)
//     backendSheet.getRange(headerRowIdx+1, netTutorIdx, netTutor.length, 1).setValues(netTutor.map(r=>[r]));

  
//   setDropdownsQAReviewAdd(qaReviewAddSheet, qaReviewUpdateSheet, backendSheet, numCols=2)
//   qaReviewUpdateSetDropdowns(qaReviewUpdateSheet, backendSheet);
//   if(spreadsheet.getOwner().getEmail() != 'Automation@upthink.com')
//     DriveApp.getFileById(spreadsheet.getId()).setOwner("automation@upthink.com");

// }








