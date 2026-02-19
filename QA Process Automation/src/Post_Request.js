//_______________________**********Edited doPost*******------

function doPost(e) {

  const sheetName = "QA DB";
  const outputSheet = MASTER_SHEET.getSheetByName(sheetName);
  const outputSheetHeader = outputSheet.getRange(1, 1, 1, outputSheet.getLastColumn()).getValues().flat();

  const outputSheetIndices = {
    srNoIdx: outputSheetHeader.indexOf('#'),
    departmentIdx: outputSheetHeader.indexOf('Department'),
    reviewerUUIDIdx:outputSheetHeader.indexOf('Reviewer UUID'),
    qaReviwerIdx: outputSheetHeader.indexOf('QA Reviewer'),
    smeUUIDIdx: outputSheetHeader.indexOf('SME UUID'),
    smeNameIdx: outputSheetHeader.indexOf('SME Name'),
    clientIdx: outputSheetHeader.indexOf('Client'),
    subjectIdx: outputSheetHeader.indexOf('Subject'),
    topicIdx: outputSheetHeader.indexOf('Topic'),
    subTopicIdx: outputSheetHeader.indexOf('Sub-Topic'),
    dateIdx: outputSheetHeader.indexOf('Review Date'),
    sessionDateIdx: outputSheetHeader.indexOf('Session Date'),
    //added new columns
    reviewTATIdx: outputSheetHeader.indexOf('Review TAT'),
    discussionTATIdx: outputSheetHeader.indexOf('Discussion TAT'),

    accountNumIdx: outputSheetHeader.indexOf('Account number'),
    boardIdx: outputSheetHeader.indexOf('Board#'),
    modeIdx: outputSheetHeader.indexOf('Mode'),
    audioIdx: outputSheetHeader.indexOf('Audio'),
    ratingsIdx: outputSheetHeader.indexOf('Rating\n(Negative/Positive/Low)'),
    negReviewReasonIdx: outputSheetHeader.indexOf('Reason for negative rating'),
    clientComplaintsIdx: outputSheetHeader.indexOf('Client Complaint'),

    identyIdx: outputSheetHeader.indexOf('SubjectKnowledge_Identify'),
    breakProcessIdx: outputSheetHeader.indexOf('SubjectKnowledge_Break The Process'),
    explanationIdx: outputSheetHeader.indexOf('SubjectKnowledge_Explanation'),
    encourageIdx: outputSheetHeader.indexOf('Tutoring_Encourage'),
    tutoringFlowIdx: outputSheetHeader.indexOf('Tutoring_Session Flow'),
    socraticIdx: outputSheetHeader.indexOf('Tutoring_Socratic'),
    greetingIdx: outputSheetHeader.indexOf('Admin_Greeting/ closing'),
    policiesIdx: outputSheetHeader.indexOf('Admin_Client policies'),
    englishFlowIdx: outputSheetHeader.indexOf('Communication_English'),
    effectiveFlowIdx: outputSheetHeader.indexOf('Communication_Effectiveness'),

    subjectKnowledgePercentIdx: outputSheetHeader.indexOf('Subject Knowledge'),
    tutoringPercentIdx: outputSheetHeader.indexOf('Tutoring'),
    adminPercentIdx: outputSheetHeader.indexOf('Admin'),
    communicationPercentIdx: outputSheetHeader.indexOf('Communication'),
    averageIdx: outputSheetHeader.indexOf('Average'),
    //delete the mapping column
    // mappingIdx: outputSheetHeader.indexOf('Mapping'),
    sessionTimeIdx: outputSheetHeader.indexOf('Session Time (in minutes)'),
    commentsIdx: outputSheetHeader.indexOf('Comments'),
    discussionIdx: outputSheetHeader.indexOf('Discussion'),
    //new column
    discussionDateIdx: outputSheetHeader.indexOf('Discussion Date'),
    discussionDurationIdx: outputSheetHeader.indexOf('Discussion Duration (Min)'),

    studentsCommentsIdx: outputSheetHeader.indexOf("Student's Comments"),
    //scoreLowRatedSessionIdx: outputSheetHeader.indexOf('Score of low rated sessions'),
    clientRatingNetTutorIdx: outputSheetHeader.indexOf('NetTutor Client Ratings (Out of 5)'),
    //replace column name with Review Time Min
    //totalMinsIdx: outputSheetHeader.indexOf('Total Hours (In Decimals)'),
    inputToTrainingTeamIdx: outputSheetHeader.indexOf('Input to Training team'),
    totalMinsIdx: outputSheetHeader.indexOf('Review Time (Min)'),
    sheetIdIdx: outputSheetHeader.indexOf('Sheet ID'),
  };
 
 

  // Check if postData exists
  if (!e.postData) {
    console.error("No postData received");
    return ContentService.createTextOutput(JSON.stringify({ 'status': 'error', 'message': 'No postData received' }));
  } else {
    const contents = JSON.parse(e.postData.contents); // Parse the JSON string into an object
    const source = contents.source;
    //add unique ids in QA Reviewer column
    const qaReviewerId = contents.qaReviewerId;
    const userEmail = contents.userEmail;
    const receivedData = contents.data;
    const sheetName = contents.sheetName;
    const department = contents.department;

    const headerRow = receivedData[0];
    const data = receivedData.slice(1);

    const srNoIdx = headerRow.indexOf('#');

    const smeUUIDIdx = headerRow.indexOf('SME UUID');
    const smeNameIdx = headerRow.indexOf('SME Name');
    const clientIdx = headerRow.indexOf('Client');
    const subjectIdx = headerRow.indexOf('Subject');
    const topicIdx = headerRow.indexOf('Topic');
    const subTopicIdx = headerRow.indexOf('Sub-Topic');
    const reviewDateIdx = headerRow.indexOf('Review Date');
    const sessionDateIdx = headerRow.indexOf('Session Date');
    //new column 
    const reviewTATIdx = headerRow.indexOf('Review TAT');
    const discussionTATIdx = headerRow.indexOf('Discussion TAT');

    const accountNumIdx = headerRow.indexOf('Account number');
    const boardIdx = headerRow.indexOf('Board#');
    const modeIdx = headerRow.indexOf('Mode');
    const audioIdx = headerRow.indexOf('Audio');
    const ratingsIdx = headerRow.indexOf('Rating\n(Negative/Positive/Low)');
    const negReviewReasonIdx = headerRow.indexOf('Reason for negative rating');

    const clientComplaintsIdx = headerRow.indexOf('Client Complaint');
    const identyIdx = headerRow.indexOf('SubjectKnowledge_Identify');
    const breakProcessIdx = headerRow.indexOf('SubjectKnowledge_Break The Process');
    const explanationIdx = headerRow.indexOf('SubjectKnowledge_Explanation');
    const encourageIdx = headerRow.indexOf('Tutoring_Encourage');
    const tutoringFlowIdx = headerRow.indexOf('Tutoring_Session Flow');
    const socraticIdx = headerRow.indexOf('Tutoring_Socratic');
    const greetingIdx = headerRow.indexOf('Admin_Greeting/ closing');
    const policiesIdx = headerRow.indexOf('Admin_Client policies');
    const englishFlowIdx = headerRow.indexOf('Communication_English');
    const effectiveFlowIdx = headerRow.indexOf('Communication_Effectiveness');

    const subjectKnowledgePercentIdx = headerRow.indexOf('Subject Knowledge');
    const tutoringPercentIdx = headerRow.indexOf('Tutoring');
    const adminPercentIdx = headerRow.indexOf('Admin');
    const communicationPercentIdx = headerRow.indexOf('Communication');
    const averageIdx = headerRow.indexOf('Average');
    const sessionTimeIdx = headerRow.indexOf('Session Time (in minutes)');
    const commentsIdx = headerRow.indexOf('Comments');
    const discussionIdx = headerRow.indexOf('Discussion');
    //new column
    const discussionDateIdx = headerRow.indexOf('Discussion Date');
    const discussionDurationIdx = headerRow.indexOf('Discussion Duration (Min)');

    const studentsCommentsIdx = headerRow.indexOf("Student's Comments");
    //const scoreLowRatedSessionIdx = headerRow.indexOf('Score of low rated sessions');
    const clientRatingNetTutorIdx = headerRow.indexOf('NetTutor Client Ratings (Out of Five)');
    //const totalMinsIdx = headerRow.indexOf('Total Hours (In Decimals)');
    const inputToTrainingTeamIdx = headerRow.indexOf('Input to Training team');
    const totalMinsIdx = headerRow.indexOf('Review Time (Min)');

    if (data.length > 0 && source === 'retrieveQAData') {

      const lock = LockService.getScriptLock();
      try {
        // Wait for up to 30 seconds for other processes to finish.
        lock.waitLock(30000);
        const startRow = outputSheet.getLastRow();
        let lastSrNo;

        let rowIdx;
        if (startRow === 1) {
          rowIdx = 2;
          lastSrNo = 1
        } else {
          rowIdx = startRow + 1;
          lastSrNo = outputSheet.getRange(startRow, outputSheetIndices.srNoIdx + 1).getValue() + 1;
        }

        data.forEach((r, index) => {
          console.log(r);
          outputSheet.getRange(rowIdx, outputSheetIndices.srNoIdx + 1).setValue(lastSrNo);
          outputSheet.getRange(rowIdx, outputSheetIndices.departmentIdx + 1).setValue(department);
          //outputSheet.getRange(rowIdx,outputSheetIndices.reviewerUUIDIdx + 1).setValue(qaReviewerId);
          //set unique id
          outputSheet.getRange(rowIdx, outputSheetIndices.reviewerUUIDIdx + 1).setValue(qaReviewerId);

          outputSheet.getRange(rowIdx, outputSheetIndices.qaReviwerIdx + 1).setValue(userEmail); 
          outputSheet.getRange(rowIdx, outputSheetIndices.smeUUIDIdx +1).setValue(r[smeUUIDIdx]);
          outputSheet.getRange(rowIdx, outputSheetIndices.smeNameIdx + 1).setValue(r[smeNameIdx]);
          outputSheet.getRange(rowIdx, outputSheetIndices.clientIdx + 1).setValue(r[clientIdx]);
          outputSheet.getRange(rowIdx, outputSheetIndices.subjectIdx + 1).setValue(r[subjectIdx]);
          outputSheet.getRange(rowIdx, outputSheetIndices.topicIdx + 1).setValue(r[topicIdx]);
          outputSheet.getRange(rowIdx, outputSheetIndices.subTopicIdx + 1).setValue(r[subTopicIdx]);
          // outputSheet.getRange(rowIdx, outputSheetIndices.dateIdx + 1)
          //   .setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yy'));
          
         
          // let reviewDate = "";
          // if (r[reviewDateIdx] !== "") {
          //   reviewDate = new Date(r[reviewDateIdx]);
          //   reviewDate = Utilities.formatDate(reviewDate, Session.getScriptTimeZone(), 'dd-MMM-yy');
          // }
           const reviewDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yy');
           outputSheet.getRange(rowIdx, outputSheetIndices.dateIdx + 1).setValue(reviewDate);

          let sessionDate;
          if (r[sessionDateIdx] !== "") {
            sessionDate = new Date(r[sessionDateIdx]);
            sessionDate = Utilities.formatDate(sessionDate, Session.getScriptTimeZone(), 'dd-MMM-yy');
          } else {
            sessionDate = "";
          }

          outputSheet.getRange(rowIdx, outputSheetIndices.sessionDateIdx + 1).setValue(sessionDate);

           if (outputSheetIndices.reviewTATIdx > -1) {
            const sessionCell = outputSheet.getRange(rowIdx, outputSheetIndices.sessionDateIdx + 1).getA1Notation();
            const reviewCell = outputSheet.getRange(rowIdx, outputSheetIndices.dateIdx + 1).getA1Notation();
            const tatCell = outputSheet.getRange(rowIdx, outputSheetIndices.reviewTATIdx + 1);
            tatCell.setFormula(`=IF(AND(ISDATE(${sessionCell}),ISDATE(${reviewCell})),${sessionCell}-${reviewCell},"")`);
          }
          
          if (outputSheetIndices.discussionTATIdx > -1) {
            const reviewCell = outputSheet.getRange(rowIdx, outputSheetIndices.dateIdx + 1).getA1Notation();
            const discussionCell = outputSheet.getRange(rowIdx, outputSheetIndices.discussionDateIdx + 1).getA1Notation();
            const tatCell = outputSheet.getRange(rowIdx, outputSheetIndices.discussionTATIdx + 1);
            tatCell.setFormula(`=IF(AND(ISDATE(${reviewCell}),ISDATE(${discussionCell})),${reviewCell}-${discussionCell},"")`);
          }
          outputSheet.getRange(rowIdx, outputSheetIndices.accountNumIdx + 1).setValue(r[accountNumIdx]);
          outputSheet.getRange(rowIdx, outputSheetIndices.boardIdx + 1).setValue(r[boardIdx]);
          outputSheet.getRange(rowIdx, outputSheetIndices.modeIdx + 1).setValue(r[modeIdx]);
          outputSheet.getRange(rowIdx, outputSheetIndices.audioIdx + 1).setValue(r[audioIdx]);
          outputSheet.getRange(rowIdx, outputSheetIndices.ratingsIdx + 1).setValue(r[ratingsIdx]);
          outputSheet.getRange(rowIdx, outputSheetIndices.negReviewReasonIdx + 1).setValue(r[negReviewReasonIdx]);
          outputSheet.getRange(rowIdx, outputSheetIndices.clientComplaintsIdx + 1).setValue(r[clientComplaintsIdx]);
          outputSheet.getRange(rowIdx, outputSheetIndices.identyIdx + 1).setValue(r[identyIdx]);
          outputSheet.getRange(rowIdx, outputSheetIndices.breakProcessIdx + 1).setValue(r[breakProcessIdx]);
          outputSheet.getRange(rowIdx, outputSheetIndices.explanationIdx + 1).setValue(r[explanationIdx]);
          outputSheet.getRange(rowIdx, outputSheetIndices.encourageIdx + 1).setValue(r[encourageIdx]);
          outputSheet.getRange(rowIdx, outputSheetIndices.tutoringFlowIdx + 1).setValue(r[tutoringFlowIdx]);
          outputSheet.getRange(rowIdx, outputSheetIndices.socraticIdx + 1).setValue(r[socraticIdx]);
          outputSheet.getRange(rowIdx, outputSheetIndices.greetingIdx + 1).setValue(r[greetingIdx]);
          outputSheet.getRange(rowIdx, outputSheetIndices.policiesIdx + 1).setValue(r[policiesIdx]);
          outputSheet.getRange(rowIdx, outputSheetIndices.englishFlowIdx + 1).setValue(r[englishFlowIdx]);
          outputSheet.getRange(rowIdx, outputSheetIndices.effectiveFlowIdx + 1).setValue(r[effectiveFlowIdx]);

          if (r[subjectKnowledgePercentIdx] !== "")
            outputSheet.getRange(rowIdx, outputSheetIndices.subjectKnowledgePercentIdx + 1).setValue(r[subjectKnowledgePercentIdx] * 100 + '%');

          if (r[tutoringPercentIdx] !== "")
            outputSheet.getRange(rowIdx, outputSheetIndices.tutoringPercentIdx + 1).setValue(r[tutoringPercentIdx] * 100 + '%');

          if (r[adminPercentIdx] !== "")
            outputSheet.getRange(rowIdx, outputSheetIndices.adminPercentIdx + 1).setValue(r[adminPercentIdx] * 100 + '%');

          if (r[communicationPercentIdx] !== "")
            outputSheet.getRange(rowIdx, outputSheetIndices.communicationPercentIdx + 1).setValue(r[communicationPercentIdx] * 100 + '%');

          if (r[averageIdx] !== "")
            outputSheet.getRange(rowIdx, outputSheetIndices.averageIdx + 1).setValue(r[averageIdx] * 100 + '%');

          // outputSheet.getRange(rowIdx, outputSheetIndices.mappingIdx + 1).setValue(r[29]);
          outputSheet.getRange(rowIdx, outputSheetIndices.sessionTimeIdx + 1).setValue(r[sessionTimeIdx]);
          outputSheet.getRange(rowIdx, outputSheetIndices.commentsIdx + 1).setValue(r[commentsIdx]);
          outputSheet.getRange(rowIdx, outputSheetIndices.discussionIdx + 1).setValue(r[discussionIdx]);

          //set date format
          let discussionDate = "";
          if (r[discussionDateIdx] !== "") {
            discussionDate = new Date();
            discussionDate = Utilities.formatDate(discussionDate, Session.getScriptTimeZone(), 'dd-MMM-yy');
          }
          const discussionDateCell = outputSheet.getRange(rowIdx, outputSheetIndices.discussionDateIdx + 1);
          discussionDateCell.setValue(discussionDate);
          discussionDateCell.setNumberFormat("dd-mmm-yy");

          outputSheet.getRange(rowIdx, outputSheetIndices.discussionDurationIdx + 1).setValue(r[discussionDurationIdx]);

          outputSheet.getRange(rowIdx, outputSheetIndices.studentsCommentsIdx + 1).setValue(r[studentsCommentsIdx]);
          //outputSheet.getRange(rowIdx, outputSheetIndices.scoreLowRatedSessionIdx + 1).setValue(r[scoreLowRatedSessionIdx]);
          outputSheet.getRange(rowIdx, outputSheetIndices.clientRatingNetTutorIdx + 1).setValue(r[clientRatingNetTutorIdx]);
          //new column
          outputSheet.getRange(rowIdx, outputSheetIndices.inputToTrainingTeamIdx + 1)
            .setValue(r[inputToTrainingTeamIdx] || "");

          outputSheet.getRange(rowIdx, outputSheetIndices.totalMinsIdx + 1).setValue(r[totalMinsIdx]);
          outputSheet.getRange(rowIdx, outputSheetIndices.sheetIdIdx + 1).setValue(sheetName);
          lastSrNo++;
          rowIdx++;
        });   // END of forEach loop
      } // Try block
      catch (e) {
        // Log any errors and/or return a failure message to the user.
        console.error(e.toString());
        return Browser.msgBox(ContentService.createTextOutput(JSON.stringify({ 'status': 'error', 'message': 'Unable to access the sheet to write data.' })));
      } finally {
        // Ensure the lock is always released, even if there's an error.
        lock.releaseLock();
      }
    }
    // END of received data length check
    else if (data.length > 0 && source === 'updateQAReviewData') {

      const serialNumberList = outputSheet.getRange(2, outputSheetIndices.srNoIdx + 1, outputSheet.getLastRow() - 1).getValues().flat();

      const lock = LockService.getScriptLock();
      try {
        // Wait for up to 30 seconds for other processes to finish.
        lock.waitLock(20000);
        data.forEach(r => {
          const rowIndex = serialNumberList.indexOf(r[srNoIdx]) + 2;
          outputSheet.getRange(rowIndex, outputSheetIndices.ratingsIdx + 1).setValue(r[ratingsIdx]);
          outputSheet.getRange(rowIndex, outputSheetIndices.negReviewReasonIdx + 1).setValue(r[negReviewReasonIdx]);
          outputSheet.getRange(rowIndex, outputSheetIndices.clientComplaintsIdx + 1).setValue(r[clientComplaintsIdx]);

          outputSheet.getRange(rowIndex, outputSheetIndices.discussionIdx + 1).setValue(r[discussionIdx]);

          // let discussionDate = "";
          // if (r[discussionDateIdx] !== "") {
          //   discussionDate = new Date();
          //   discussionDate = Utilities.formatDate(discussionDate, Session.getScriptTimeZone(), 'dd-MMM-yy');
          // }
          // const discussionDateCell = outputSheet.getRange(rowIndex, outputSheetIndices.discussionDateIdx + 1);
          // discussionDateCell.setValue(discussionDate);
          // discussionDateCell.setNumberFormat("dd-mmm-yy");
         
          // --- Discussion Date Logic ---
          let discussionDate = r[discussionDateIdx]; // take the value user entered

          if (discussionDate) {
            // if it's a Date object, format it
            if (discussionDate instanceof Date) {
              discussionDate = Utilities.formatDate(discussionDate, Session.getScriptTimeZone(), 'dd-MMM-yy');
            } else {
              // if it's a string, try to parse it directly
              const parsedDate = new Date(discussionDate);
              if (!isNaN(parsedDate)) {
                discussionDate = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), 'dd-MMM-yy');
              }
            }
          }

          const discussionDateCell = outputSheet.getRange(rowIndex, outputSheetIndices.discussionDateIdx + 1);
          discussionDateCell.setValue(discussionDate || "");
          discussionDateCell.setNumberFormat("dd-mmm-yy");






          outputSheet.getRange(rowIndex, outputSheetIndices.discussionDurationIdx + 1).setValue(r[discussionDurationIdx]);

          // outputSheet.getRange(rowIndex, outputSheetIndices.scoreLowRatedSessionIdx + 1).setValue(r[scoreLowRatedSessionIdx]);

          
          outputSheet.getRange(rowIndex, outputSheetIndices.clientRatingNetTutorIdx + 1).setValue(r[clientRatingNetTutorIdx]);
          //new column
          outputSheet.getRange(rowIndex, outputSheetIndices.inputToTrainingTeamIdx + 1)
            .setValue(r[inputToTrainingTeamIdx]);

          outputSheet.getRange(rowIndex, outputSheetIndices.totalMinsIdx + 1).setValue(r[totalMinsIdx]);

          
        });
      } catch (e) {

        console.error(e.toString());
        return Browser.msgBox(ContentService.createTextOutput(JSON.stringify({ 'status': 'error', 'message': 'Unable to access the sheet to write data.' })));

      } finally {
        // Ensure the lock is always released, even if there's an error.
        lock.releaseLock();
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ 'status': 'success', 'message': 'Data processed successfully' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}






function convertToScriptTimeZone(date) {
  var timeZone = Session.getScriptTimeZone();
  var formattedDate = Utilities.formatDate(date, timeZone, "dd-MMM-yy");
  return new Date(formattedDate);
}







//my edited code
// function doPost(e) {

//   const sheetName = "QA DB";
//   const outputSheet = MASTER_SHEET.getSheetByName(sheetName);
//   const outputSheetHeader = outputSheet.getRange(1, 1, 1, outputSheet.getLastColumn()).getValues().flat();

//   const outputSheetIndices = {
//     srNoIdx: outputSheetHeader.indexOf('#'),
//     departmentIdx: outputSheetHeader.indexOf('Department'),
//     qaReviwerIdx: outputSheetHeader.indexOf('QA Reviewer'),
//     smeNameIdx: outputSheetHeader.indexOf('SME Name'),
//     clientIdx: outputSheetHeader.indexOf('Client'),
//     subjectIdx: outputSheetHeader.indexOf('Subject'),
//     topicIdx: outputSheetHeader.indexOf('Topic'),
//     subTopicIdx: outputSheetHeader.indexOf('Sub-Topic'),
//     dateIdx: outputSheetHeader.indexOf('Review Date'),
//     sessionDateIdx: outputSheetHeader.indexOf('Session Date'),
//     //added new columns
//     reviewTATIdx: outputSheetHeader.indexOf('Review TAT'),
//     discussionTATIdx: outputSheetHeader.indexOf('Discussion TAT'),

//     accountNumIdx: outputSheetHeader.indexOf('Account number'),
//     boardIdx: outputSheetHeader.indexOf('Board#'),
//     modeIdx: outputSheetHeader.indexOf('Mode'),
//     audioIdx: outputSheetHeader.indexOf('Audio'),
//     ratingsIdx: outputSheetHeader.indexOf('Rating\n(Negative/Positive/Low)'),
//     negReviewReasonIdx: outputSheetHeader.indexOf('Reason for negative rating'),
//     clientComplaintsIdx: outputSheetHeader.indexOf('Client Complaint'),

//     identyIdx: outputSheetHeader.indexOf('SubjectKnowledge_Identify'),
//     breakProcessIdx: outputSheetHeader.indexOf('SubjectKnowledge_Break The Process'),
//     explanationIdx: outputSheetHeader.indexOf('SubjectKnowledge_Explanation'),
//     encourageIdx: outputSheetHeader.indexOf('Tutoring_Encourage'),
//     tutoringFlowIdx: outputSheetHeader.indexOf('Tutoring_Session Flow'),
//     socraticIdx: outputSheetHeader.indexOf('Tutoring_Socratic'),
//     greetingIdx: outputSheetHeader.indexOf('Admin_Greeting/ closing'),
//     policiesIdx: outputSheetHeader.indexOf('Admin_Client policies'),
//     englishFlowIdx: outputSheetHeader.indexOf('Communication_English'),
//     effectiveFlowIdx: outputSheetHeader.indexOf('Communication_Effectiveness'),

//     subjectKnowledgePercentIdx: outputSheetHeader.indexOf('Subject Knowledge'),
//     tutoringPercentIdx: outputSheetHeader.indexOf('Tutoring'),
//     adminPercentIdx: outputSheetHeader.indexOf('Admin'),
//     communicationPercentIdx: outputSheetHeader.indexOf('Communication'),
//     averageIdx: outputSheetHeader.indexOf('Average'),
//     //delete the mapping column
//     // mappingIdx: outputSheetHeader.indexOf('Mapping'),
//     sessionTimeIdx: outputSheetHeader.indexOf('Session Time (in minutes)'),
//     commentsIdx: outputSheetHeader.indexOf('Comments'),
//     discussionIdx: outputSheetHeader.indexOf('Discussion'),
//     //new column
//     discussionDateIdx: outputSheetHeader.indexOf('Discussion Date'),
//     discussionDurationIdx: outputSheetHeader.indexOf('Discussion Duration (Min)'),

//     studentsCommentsIdx: outputSheetHeader.indexOf("Student's Comments"),
//     //scoreLowRatedSessionIdx: outputSheetHeader.indexOf('Score of low rated sessions'),
//     clientRatingNetTutorIdx: outputSheetHeader.indexOf('NetTutor Client Ratings (Out of 5)'),
//     //replace column name with Review Time Min
//     //totalMinsIdx: outputSheetHeader.indexOf('Total Hours (In Decimals)'),
//     inputToTrainingTeamIdx: outputSheetHeader.indexOf('Input to Training team'),
//     totalMinsIdx: outputSheetHeader.indexOf('Review Time (Min)'),
//     sheetIdIdx: outputSheetHeader.indexOf('Sheet ID'),
//   };


//   // Check if postData exists
//   if (!e.postData) {
//     console.error("No postData received");
//     return ContentService.createTextOutput(JSON.stringify({ 'status': 'error', 'message': 'No postData received' }));
//   } else {
//     const contents = JSON.parse(e.postData.contents); // Parse the JSON string into an object
//     const source = contents.source;
//     //add unique ids in QA Reviewer column
//     const qaReviewerId = contents.qaReviewerId;
//     //const userEmail = contents.userEmail;
//     const receivedData = contents.data;
//     const sheetName = contents.sheetName;
//     const department = contents.department;

//     const headerRow = receivedData[0];
//     const data = receivedData.slice(1);

//     const srNoIdx = headerRow.indexOf("#")
//     const smeNameIdx = headerRow.indexOf('SME Name');
//     const clientIdx = headerRow.indexOf('Client');
//     const subjectIdx = headerRow.indexOf('Subject');
//     const topicIdx = headerRow.indexOf('Topic');
//     const subTopicIdx = headerRow.indexOf('Sub-Topic');
//     const reviewDateIdx = headerRow.indexOf('Review Date');
//     const sessionDateIdx = headerRow.indexOf('Session Date');
//     //new column 
//     const reviewTATIdx = headerRow.indexOf('Review TAT');
//     const discussionTATIdx = headerRow.indexOf('Discussion TAT');

//     const accountNumIdx = headerRow.indexOf('Account number');
//     const boardIdx = headerRow.indexOf('Board#');
//     const modeIdx = headerRow.indexOf('Mode');
//     const audioIdx = headerRow.indexOf('Audio');
//     const ratingsIdx = headerRow.indexOf('Rating\n(Negative/Positive/Low)');
//     const negReviewReasonIdx = headerRow.indexOf('Reason for negative rating');

//     const clientComplaintsIdx = headerRow.indexOf('Client Complaint');
//     const identyIdx = headerRow.indexOf('SubjectKnowledge_Identify');
//     const breakProcessIdx = headerRow.indexOf('SubjectKnowledge_Break The Process');
//     const explanationIdx = headerRow.indexOf('SubjectKnowledge_Explanation');
//     const encourageIdx = headerRow.indexOf('Tutoring_Encourage');
//     const tutoringFlowIdx = headerRow.indexOf('Tutoring_Session Flow');
//     const socraticIdx = headerRow.indexOf('Tutoring_Socratic');
//     const greetingIdx = headerRow.indexOf('Admin_Greeting/ closing');
//     const policiesIdx = headerRow.indexOf('Admin_Client policies');
//     const englishFlowIdx = headerRow.indexOf('Communication_English');
//     const effectiveFlowIdx = headerRow.indexOf('Communication_Effectiveness');

//     const subjectKnowledgePercentIdx = headerRow.indexOf('Subject Knowledge');
//     const tutoringPercentIdx = headerRow.indexOf('Tutoring');
//     const adminPercentIdx = headerRow.indexOf('Admin');
//     const communicationPercentIdx = headerRow.indexOf('Communication');
//     const averageIdx = headerRow.indexOf('Average');
//     const sessionTimeIdx = headerRow.indexOf('Session Time (in minutes)');
//     const commentsIdx = headerRow.indexOf('Comments');
//     const discussionIdx = headerRow.indexOf('Discussion');
//     //new column
//     const discussionDateIdx = headerRow.indexOf('Discussion Date');
//     const discussionDurationIdx = headerRow.indexOf('Discussion Duration (Min)');

//     const studentsCommentsIdx = headerRow.indexOf("Student's Comments");
//     //const scoreLowRatedSessionIdx = headerRow.indexOf('Score of low rated sessions');
//     const clientRatingNetTutorIdx = headerRow.indexOf('NetTutor Client Ratings (Out of Five)');
//     //const totalMinsIdx = headerRow.indexOf('Total Hours (In Decimals)');
//     const inputToTrainingTeamIdx = headerRow.indexOf('Input to Training team');
//     const totalMinsIdx = headerRow.indexOf('Review Time (Min)');

//     if (data.length > 0 && source === 'retrieveQAData') {

//       const lock = LockService.getScriptLock();
//       try {
//         // Wait for up to 30 seconds for other processes to finish.
//         lock.waitLock(30000);
//         const startRow = outputSheet.getLastRow();
//         let lastSrNo;

//         let rowIdx;
//         if (startRow === 1) {
//           rowIdx = 2;
//           lastSrNo = 1
//         } else {
//           rowIdx = startRow + 1;
//           lastSrNo = outputSheet.getRange(startRow, outputSheetIndices.srNoIdx + 1).getValue() + 1;
//         }

//         data.forEach((r, index) => {
//           outputSheet.getRange(rowIdx, outputSheetIndices.srNoIdx + 1).setValue(lastSrNo);
//           outputSheet.getRange(rowIdx, outputSheetIndices.departmentIdx + 1).setValue(department);
//           //set unique id
//           outputSheet.getRange(rowIdx, outputSheetIndices.qaReviwerIdx + 1).setValue(qaReviewerId);

//           //outputSheet.getRange(rowIdx, outputSheetIndices.qaReviwerIdx + 1).setValue(userEmail);
//           outputSheet.getRange(rowIdx, outputSheetIndices.smeNameIdx + 1).setValue(r[smeNameIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.clientIdx + 1).setValue(r[clientIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.subjectIdx + 1).setValue(r[subjectIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.topicIdx + 1).setValue(r[topicIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.subTopicIdx + 1).setValue(r[subTopicIdx]);
//           // outputSheet.getRange(rowIdx, outputSheetIndices.dateIdx + 1)
//           //   .setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yy'));
          
         
//           let reviewDate = "";
//           if (r[reviewDateIdx] !== "") {
//             reviewDate = new Date(r[reviewDateIdx]);
//             reviewDate = Utilities.formatDate(reviewDate, Session.getScriptTimeZone(), 'dd-MMM-yy');
//           }
//           outputSheet.getRange(rowIdx, outputSheetIndices.dateIdx + 1).setValue(reviewDate);

//           let sessionDate;
//           if (r[sessionDateIdx] !== "") {
//             sessionDate = new Date(r[sessionDateIdx]);
//             sessionDate = Utilities.formatDate(sessionDate, Session.getScriptTimeZone(), 'dd-MMM-yy');
//           } else {
//             sessionDate = "";
//           }

//           outputSheet.getRange(rowIdx, outputSheetIndices.sessionDateIdx + 1).setValue(sessionDate);
//           //new column 
//           outputSheet.getRange(rowIdx, outputSheetIndices.reviewTATIdx + 1).setValue(r[reviewTATIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.discussionTATIdx + 1).setValue(r[discussionTATIdx]);

//           outputSheet.getRange(rowIdx, outputSheetIndices.accountNumIdx + 1).setValue(r[accountNumIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.boardIdx + 1).setValue(r[boardIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.modeIdx + 1).setValue(r[modeIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.audioIdx + 1).setValue(r[audioIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.ratingsIdx + 1).setValue(r[ratingsIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.negReviewReasonIdx + 1).setValue(r[negReviewReasonIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.clientComplaintsIdx + 1).setValue(r[clientComplaintsIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.identyIdx + 1).setValue(r[identyIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.breakProcessIdx + 1).setValue(r[breakProcessIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.explanationIdx + 1).setValue(r[explanationIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.encourageIdx + 1).setValue(r[encourageIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.tutoringFlowIdx + 1).setValue(r[tutoringFlowIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.socraticIdx + 1).setValue(r[socraticIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.greetingIdx + 1).setValue(r[greetingIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.policiesIdx + 1).setValue(r[policiesIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.englishFlowIdx + 1).setValue(r[englishFlowIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.effectiveFlowIdx + 1).setValue(r[effectiveFlowIdx]);

//           if (r[subjectKnowledgePercentIdx] !== "")
//             outputSheet.getRange(rowIdx, outputSheetIndices.subjectKnowledgePercentIdx + 1).setValue(r[subjectKnowledgePercentIdx] * 100 + '%');

//           if (r[tutoringPercentIdx] !== "")
//             outputSheet.getRange(rowIdx, outputSheetIndices.tutoringPercentIdx + 1).setValue(r[tutoringPercentIdx] * 100 + '%');

//           if (r[adminPercentIdx] !== "")
//             outputSheet.getRange(rowIdx, outputSheetIndices.adminPercentIdx + 1).setValue(r[adminPercentIdx] * 100 + '%');

//           if (r[communicationPercentIdx] !== "")
//             outputSheet.getRange(rowIdx, outputSheetIndices.communicationPercentIdx + 1).setValue(r[communicationPercentIdx] * 100 + '%');

//           if (r[averageIdx] !== "")
//             outputSheet.getRange(rowIdx, outputSheetIndices.averageIdx + 1).setValue(r[averageIdx] * 100 + '%');

//           // outputSheet.getRange(rowIdx, outputSheetIndices.mappingIdx + 1).setValue(r[29]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.sessionTimeIdx + 1).setValue(r[sessionTimeIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.commentsIdx + 1).setValue(r[commentsIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.discussionIdx + 1).setValue(r[discussionIdx]);

//           //set date format
//           let discussionDate = "";
//           if (r[discussionDateIdx] !== "") {
//             discussionDate = new Date(r[discussionDateIdx]);
//             discussionDate = Utilities.formatDate(discussionDate, Session.getScriptTimeZone(), 'dd-MMM-yy');
//           }
//           const discussionDateCell = outputSheet.getRange(rowIdx, outputSheetIndices.discussionDateIdx + 1);
//           discussionDateCell.setValue(discussionDate);
//           discussionDateCell.setNumberFormat("dd-mmm-yy");

//           outputSheet.getRange(rowIdx, outputSheetIndices.discussionDurationIdx + 1).setValue(r[discussionDurationIdx]);

//           outputSheet.getRange(rowIdx, outputSheetIndices.studentsCommentsIdx + 1).setValue(r[studentsCommentsIdx]);
//           //outputSheet.getRange(rowIdx, outputSheetIndices.scoreLowRatedSessionIdx + 1).setValue(r[scoreLowRatedSessionIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.clientRatingNetTutorIdx + 1).setValue(r[clientRatingNetTutorIdx]);
//           //new column
//           outputSheet.getRange(rowIdx, outputSheetIndices.inputToTrainingTeamIdx + 1)
//             .setValue(r[inputToTrainingTeamIdx] || "");

//           outputSheet.getRange(rowIdx, outputSheetIndices.totalMinsIdx + 1).setValue(r[totalMinsIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.sheetIdIdx + 1).setValue(sheetName);
//           lastSrNo++;
//           rowIdx++;
//         });   // END of forEach loop
//       } // Try block
//       catch (e) {
//         // Log any errors and/or return a failure message to the user.
//         console.error(e.toString());
//         return Browser.msgBox(ContentService.createTextOutput(JSON.stringify({ 'status': 'error', 'message': 'Unable to access the sheet to write data.' })));
//       } finally {
//         // Ensure the lock is always released, even if there's an error.
//         lock.releaseLock();
//       }
//     }
//     // END of received data length check
//     else if (data.length > 0 && source === 'updateQAReviewData') {

//       const serialNumberList = outputSheet.getRange(2, outputSheetIndices.srNoIdx + 1, outputSheet.getLastRow() - 1).getValues().flat();

//       const lock = LockService.getScriptLock();
//       try {
//         // Wait for up to 30 seconds for other processes to finish.
//         lock.waitLock(20000);
//         data.forEach(r => {
//           const rowIndex = serialNumberList.indexOf(r[srNoIdx]) + 2;
//           outputSheet.getRange(rowIndex, outputSheetIndices.ratingsIdx + 1).setValue(r[ratingsIdx]);
//           outputSheet.getRange(rowIndex, outputSheetIndices.negReviewReasonIdx + 1).setValue(r[negReviewReasonIdx]);
//           outputSheet.getRange(rowIndex, outputSheetIndices.clientComplaintsIdx + 1).setValue(r[clientComplaintsIdx]);
//           //add new columns
//           outputSheet.getRange(rowIndex, outputSheetIndices.reviewTATIdx + 1).setValue(r[reviewTATIdx]);
//           outputSheet.getRange(rowIndex, outputSheetIndices.discussionTATIdx + 1).setValue(r[discussionTATIdx]);

//           outputSheet.getRange(rowIndex, outputSheetIndices.discussionIdx + 1).setValue(r[discussionIdx]);

//           let discussionDate = "";
//           if (r[discussionDateIdx] !== "") {
//             discussionDate = new Date(r[discussionDateIdx]);
//             discussionDate = Utilities.formatDate(discussionDate, Session.getScriptTimeZone(), 'dd-MMM-yy');
//           }
//           const discussionDateCell = outputSheet.getRange(rowIndex, outputSheetIndices.discussionDateIdx + 1);
//           discussionDateCell.setValue(discussionDate);
//           discussionDateCell.setNumberFormat("dd-mmm-yy");

//           outputSheet.getRange(rowIndex, outputSheetIndices.discussionDurationIdx + 1).setValue(r[discussionDurationIdx]);

//           //outputSheet.getRange(rowIndex, outputSheetIndices.scoreLowRatedSessionIdx + 1).setValue(r[scoreLowRatedSessionIdx]);
//           outputSheet.getRange(rowIndex, outputSheetIndices.clientRatingNetTutorIdx + 1).setValue(r[clientRatingNetTutorIdx]);
//           //new column
//           outputSheet.getRange(rowIndex, outputSheetIndices.inputToTrainingTeamIdx + 1)
//             .setValue(r[inputToTrainingTeamIdx]);

//           outputSheet.getRange(rowIndex, outputSheetIndices.totalMinsIdx + 1).setValue(r[totalMinsIdx]);
//         });
//       } catch (e) {

//         console.error(e.toString());
//         return Browser.msgBox(ContentService.createTextOutput(JSON.stringify({ 'status': 'error', 'message': 'Unable to access the sheet to write data.' })));

//       } finally {
//         // Ensure the lock is always released, even if there's an error.
//         lock.releaseLock();
//       }
//     }
//     return ContentService.createTextOutput(JSON.stringify({ 'status': 'success', 'message': 'Data processed successfully' }))
//       .setMimeType(ContentService.MimeType.JSON);
//   }
// }

// function convertToScriptTimeZone(date) {
//   var timeZone = Session.getScriptTimeZone();
//   var formattedDate = Utilities.formatDate(date, timeZone, "dd-MMM-yy");
//   return new Date(formattedDate);
// }

//old code
// function doPost(e) {

//   const sheetName = "QA DB";
//   const outputSheet = MASTER_SHEET.getSheetByName(sheetName);
//   const outputSheetHeader = outputSheet.getRange(1, 1, 1, outputSheet.getLastColumn()).getValues().flat();

//   const outputSheetIndices = {
//     srNoIdx: outputSheetHeader.indexOf('#'),
//     departmentIdx: outputSheetHeader.indexOf('Department'),
//     qaReviwerIdx: outputSheetHeader.indexOf('QA Reviewer'),
//     smeNameIdx: outputSheetHeader.indexOf('SME Name'),
//     clientIdx: outputSheetHeader.indexOf('Client'),
//     subjectIdx: outputSheetHeader.indexOf('Subject'),
//     topicIdx: outputSheetHeader.indexOf('Topic'),
//     subTopicIdx: outputSheetHeader.indexOf('Sub-Topic'),
//     dateIdx: outputSheetHeader.indexOf('Review Date'),
//     sessionDateIdx: outputSheetHeader.indexOf('Session Date'),
//     accountNumIdx: outputSheetHeader.indexOf('Account number'),
//     boardIdx: outputSheetHeader.indexOf('Board#'),
//     modeIdx: outputSheetHeader.indexOf('Mode'),
//     audioIdx: outputSheetHeader.indexOf('Audio'),
//     ratingsIdx: outputSheetHeader.indexOf('Rating\n(Negative/Positive/Low)'),
//     negReviewReasonIdx: outputSheetHeader.indexOf('Reason for negative rating'),
//     clientComplaintsIdx: outputSheetHeader.indexOf('Client Complaint'),

//     identyIdx: outputSheetHeader.indexOf('SubjectKnowledge_Identify'),
//     breakProcessIdx: outputSheetHeader.indexOf('SubjectKnowledge_Break The Process'),
//     explanationIdx: outputSheetHeader.indexOf('SubjectKnowledge_Explanation'),
//     encourageIdx: outputSheetHeader.indexOf('Tutoring_Encourage'),
//     tutoringFlowIdx: outputSheetHeader.indexOf('Tutoring_Session Flow'),
//     socraticIdx: outputSheetHeader.indexOf('Tutoring_Socratic'),
//     greetingIdx: outputSheetHeader.indexOf('Admin_Greeting/ closing'),
//     policiesIdx: outputSheetHeader.indexOf('Admin_Client policies'),
//     englishFlowIdx: outputSheetHeader.indexOf('Communication_English'),
//     effectiveFlowIdx: outputSheetHeader.indexOf('Communication_Effectiveness'),

//     subjectKnowledgePercentIdx: outputSheetHeader.indexOf('Subject Knowledge'),
//     tutoringPercentIdx: outputSheetHeader.indexOf('Tutoring'),
//     adminPercentIdx: outputSheetHeader.indexOf('Admin'),
//     communicationPercentIdx: outputSheetHeader.indexOf('Communication'),
//     averageIdx: outputSheetHeader.indexOf('Average'),
//    // mappingIdx: outputSheetHeader.indexOf('Mapping'),
//     sessionTimeIdx: outputSheetHeader.indexOf('Session Time (in minutes)'),
//     commentsIdx: outputSheetHeader.indexOf('Comments'),
//     discussionIdx: outputSheetHeader.indexOf('Discussion'),
//     //new column
//     discussionDateIdx: outputSheetHeader.indexOf('Discussion Date'),
//     discussionDurationIdx: outputSheetHeader.indexOf('Discussion Duration (Min)'),

//     studentsCommentsIdx: outputSheetHeader.indexOf("Student's Comments"),
//     scoreLowRatedSessionIdx: outputSheetHeader.indexOf('Score of low rated sessions'),
//     clientRatingNetTutorIdx: outputSheetHeader.indexOf('NetTutor Client Ratings (Out of 5)'),
//     //totalMinsIdx: outputSheetHeader.indexOf('Total Hours (In Decimals)'),
//     totalMinsIdx: outputSheetHeader.indexOf('Review Time (Min)'),
//     sheetIdIdx: outputSheetHeader.indexOf('Sheet ID'),
//   };


//   // Check if postData exists
//   if (!e.postData) {
//     console.error("No postData received");
//     return ContentService.createTextOutput(JSON.stringify({ 'status': 'error', 'message': 'No postData received' }));
//   } else {

//     const contents = JSON.parse(e.postData.contents); // Parse the JSON string into an object
//     const source = contents.source;
//     const qaReviewerId = contents.qaReviewerId;
//     //const userEmail = contents.userEmail;
//     const receivedData = contents.data;
//     const sheetName = contents.sheetName;
//     const department = contents.department;

//     const headerRow = receivedData[0];
//     const data = receivedData.slice(1);

//     const srNoIdx = headerRow.indexOf("#")
//     const smeNameIdx = headerRow.indexOf('SME Name');
//     const clientIdx = headerRow.indexOf('Client');
//     const subjectIdx = headerRow.indexOf('Subject');
//     const topicIdx = headerRow.indexOf('Topic');
//     const subTopicIdx = headerRow.indexOf('Sub-Topic');
//     const sessionDateIdx = headerRow.indexOf('Session Date');
//     const accountNumIdx = headerRow.indexOf('Account number');
//     const boardIdx = headerRow.indexOf('Board#');
//     const modeIdx = headerRow.indexOf('Mode');
//     const audioIdx = headerRow.indexOf('Audio');
//     const ratingsIdx = headerRow.indexOf('Rating\n(Negative/Positive/Low)');
//     const negReviewReasonIdx = headerRow.indexOf('Reason for negative rating');

//     const clientComplaintsIdx = headerRow.indexOf('Client Complaint');
//     const identyIdx = headerRow.indexOf('SubjectKnowledge_Identify');
//     const breakProcessIdx = headerRow.indexOf('SubjectKnowledge_Break The Process');
//     const explanationIdx = headerRow.indexOf('SubjectKnowledge_Explanation');
//     const encourageIdx = headerRow.indexOf('Tutoring_Encourage');
//     const tutoringFlowIdx = headerRow.indexOf('Tutoring_Session Flow');
//     const socraticIdx = headerRow.indexOf('Tutoring_Socratic');
//     const greetingIdx = headerRow.indexOf('Admin_Greeting/ closing');
//     const policiesIdx = headerRow.indexOf('Admin_Client policies');
//     const englishFlowIdx = headerRow.indexOf('Communication_English');
//     const effectiveFlowIdx = headerRow.indexOf('Communication_Effectiveness');

//     const subjectKnowledgePercentIdx = headerRow.indexOf('Subject Knowledge');
//     const tutoringPercentIdx = headerRow.indexOf('Tutoring');
//     const adminPercentIdx = headerRow.indexOf('Admin');
//     const communicationPercentIdx = headerRow.indexOf('Communication');
//     const averageIdx = headerRow.indexOf('Average');
//     const sessionTimeIdx = headerRow.indexOf('Session Time (in minutes)');
//     const commentsIdx = headerRow.indexOf('Comments');
//     const discussionIdx = headerRow.indexOf('Discussion');
//     //new column
//     const discussionDateIdx = headerRow.indexOf('Discussion Date');
//     const discussionDurationIdx = headerRow.indexOf('Discussion Duration (Min)');

//     const studentsCommentsIdx = headerRow.indexOf("Student's Comments");
//     const scoreLowRatedSessionIdx = headerRow.indexOf('Score of low rated sessions');
//     const clientRatingNetTutorIdx = headerRow.indexOf('NetTutor Client Ratings (Out of Five)');
//     //const totalMinsIdx = headerRow.indexOf('Total Hours (In Decimals)');
//     const totalMinsIdx = headerRow.indexOf('Review Time (Min)');

//     if (data.length > 0 && source === 'retrieveQAData') {

//       const lock = LockService.getScriptLock();
//       try {
//         // Wait for up to 30 seconds for other processes to finish.
//         lock.waitLock(30000);
//         const startRow = outputSheet.getLastRow();
//         let lastSrNo;

//         let rowIdx;
//         if (startRow === 1) {
//           rowIdx = 2;
//           lastSrNo = 1
//         } else {
//           rowIdx = startRow + 1;
//           lastSrNo = outputSheet.getRange(startRow, outputSheetIndices.srNoIdx + 1).getValue() + 1;
//         }

//         data.forEach((r, index) => {
//           outputSheet.getRange(rowIdx, outputSheetIndices.srNoIdx + 1).setValue(lastSrNo);
//           outputSheet.getRange(rowIdx, outputSheetIndices.departmentIdx + 1).setValue(department);

//           outputSheet.getRange(rowIdx, outputSheetIndices.qaReviwerIdx + 1).setValue(qaReviewerId);

//           //outputSheet.getRange(rowIdx, outputSheetIndices.qaReviwerIdx + 1).setValue(userEmail);
//           outputSheet.getRange(rowIdx, outputSheetIndices.smeNameIdx + 1).setValue(r[smeNameIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.clientIdx + 1).setValue(r[clientIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.subjectIdx + 1).setValue(r[subjectIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.topicIdx + 1).setValue(r[topicIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.subTopicIdx + 1).setValue(r[subTopicIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.dateIdx + 1)
//             .setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yy'));
//           let sessionDate;
//           if (r[sessionDateIdx] !== "") {
//             sessionDate = new Date(r[sessionDateIdx]);
//             sessionDate = Utilities.formatDate(sessionDate, Session.getScriptTimeZone(), 'dd-MMM-yy');
//           } else {
//             sessionDate = "";
//           }

//           outputSheet.getRange(rowIdx, outputSheetIndices.sessionDateIdx + 1).setValue(sessionDate);
//           outputSheet.getRange(rowIdx, outputSheetIndices.accountNumIdx + 1).setValue(r[accountNumIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.boardIdx + 1).setValue(r[boardIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.modeIdx + 1).setValue(r[modeIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.audioIdx + 1).setValue(r[audioIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.ratingsIdx + 1).setValue(r[ratingsIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.negReviewReasonIdx + 1).setValue(r[negReviewReasonIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.clientComplaintsIdx + 1).setValue(r[clientComplaintsIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.identyIdx + 1).setValue(r[identyIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.breakProcessIdx + 1).setValue(r[breakProcessIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.explanationIdx + 1).setValue(r[explanationIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.encourageIdx + 1).setValue(r[encourageIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.tutoringFlowIdx + 1).setValue(r[tutoringFlowIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.socraticIdx + 1).setValue(r[socraticIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.greetingIdx + 1).setValue(r[greetingIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.policiesIdx + 1).setValue(r[policiesIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.englishFlowIdx + 1).setValue(r[englishFlowIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.effectiveFlowIdx + 1).setValue(r[effectiveFlowIdx]);

//           if (r[subjectKnowledgePercentIdx] !== "")
//             outputSheet.getRange(rowIdx, outputSheetIndices.subjectKnowledgePercentIdx + 1).setValue(r[subjectKnowledgePercentIdx] * 100 + '%');

//           if (r[tutoringPercentIdx] !== "")
//             outputSheet.getRange(rowIdx, outputSheetIndices.tutoringPercentIdx + 1).setValue(r[tutoringPercentIdx] * 100 + '%');

//           if (r[adminPercentIdx] !== "")
//             outputSheet.getRange(rowIdx, outputSheetIndices.adminPercentIdx + 1).setValue(r[adminPercentIdx] * 100 + '%');

//           if (r[communicationPercentIdx] !== "")
//             outputSheet.getRange(rowIdx, outputSheetIndices.communicationPercentIdx + 1).setValue(r[communicationPercentIdx] * 100 + '%');

//           if (r[averageIdx] !== "")
//             outputSheet.getRange(rowIdx, outputSheetIndices.averageIdx + 1).setValue(r[averageIdx] * 100 + '%');

//           // outputSheet.getRange(rowIdx, outputSheetIndices.mappingIdx + 1).setValue(r[29]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.sessionTimeIdx + 1).setValue(r[sessionTimeIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.commentsIdx + 1).setValue(r[commentsIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.discussionIdx + 1).setValue(r[discussionIdx]);
//           if (r[discussionDateIdx] !== "") {
//             const discussionDate = new Date(r[discussionDateIdx]);
//             outputSheet.getRange(rowIdx, outputSheetIndices.discussionDateIdx + 1)
//               .setValue(discussionDate)
//               .setNumberFormat('dd-mmm-yy');
//           }
//           outputSheet.getRange(rowIdx, outputSheetIndices.discussionDurationIdx + 1).setValue(r[discussionDurationIdx]);

//           outputSheet.getRange(rowIdx, outputSheetIndices.studentsCommentsIdx + 1).setValue(r[studentsCommentsIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.scoreLowRatedSessionIdx + 1).setValue(r[scoreLowRatedSessionIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.clientRatingNetTutorIdx + 1).setValue(r[clientRatingNetTutorIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.totalMinsIdx + 1).setValue(r[totalMinsIdx]);
//           outputSheet.getRange(rowIdx, outputSheetIndices.sheetIdIdx + 1).setValue(sheetName);
//           lastSrNo++;
//           rowIdx++;
//         });   // END of forEach loop
//       } // Try block
//       catch (e) {
//         // Log any errors and/or return a failure message to the user.
//         console.error(e.toString());
//         return Browser.msgBox(ContentService.createTextOutput(JSON.stringify({ 'status': 'error', 'message': 'Unable to access the sheet to write data.' })));
//       } finally {
//         // Ensure the lock is always released, even if there's an error.
//         lock.releaseLock();
//       }


//     } // END of received data length check
//     else if (data.length > 0 && source === 'updateQAReviewData') {

//       const serialNumberList = outputSheet.getRange(2, outputSheetIndices.srNoIdx + 1, outputSheet.getLastRow() - 1).getValues().flat();

//       const lock = LockService.getScriptLock();
//       try {
//         // Wait for up to 30 seconds for other processes to finish.
//         lock.waitLock(20000);
//         data.forEach(r => {
//           const rowIndex = serialNumberList.indexOf(r[srNoIdx]) + 2;
//           outputSheet.getRange(rowIndex, outputSheetIndices.ratingsIdx + 1).setValue(r[ratingsIdx]);
//           outputSheet.getRange(rowIndex, outputSheetIndices.negReviewReasonIdx + 1).setValue(r[negReviewReasonIdx]);
//           outputSheet.getRange(rowIndex, outputSheetIndices.clientComplaintsIdx + 1).setValue(r[clientComplaintsIdx]);
//           outputSheet.getRange(rowIndex, outputSheetIndices.discussionIdx + 1).setValue(r[discussionIdx]);
//           let discussionDate = "";
//           if (r[discussionDateIdx] !== "") {
//             discussionDate = new Date(r[discussionDateIdx]);
//             discussionDate = Utilities.formatDate(discussionDate, Session.getScriptTimeZone(), 'dd-MMM-yy');
//           }
//           const discussionDateCell = outputSheet.getRange(rowIndex, outputSheetIndices.discussionDateIdx + 1);
//           discussionDateCell.setValue(discussionDate);
//           discussionDateCell.setNumberFormat("dd-mmm-yy");

//           outputSheet.getRange(rowIndex, outputSheetIndices.discussionDurationIdx + 1).setValue(r[discussionDurationIdx]);

//           outputSheet.getRange(rowIndex, outputSheetIndices.scoreLowRatedSessionIdx + 1).setValue(r[scoreLowRatedSessionIdx]);
//           outputSheet.getRange(rowIndex, outputSheetIndices.clientRatingNetTutorIdx + 1).setValue(r[clientRatingNetTutorIdx]);
//           outputSheet.getRange(rowIndex, outputSheetIndices.totalMinsIdx + 1).setValue(r[totalMinsIdx]);
//         });


//       } catch (e) {

//         console.error(e.toString());
//         return Browser.msgBox(ContentService.createTextOutput(JSON.stringify({ 'status': 'error', 'message': 'Unable to access the sheet to write data.' })));

//       } finally {
//         // Ensure the lock is always released, even if there's an error.
//         lock.releaseLock();
//       }
//     }
//     return ContentService.createTextOutput(JSON.stringify({ 'status': 'success', 'message': 'Data processed successfully' }))
//       .setMimeType(ContentService.MimeType.JSON);

//   }
// }

// function convertToScriptTimeZone(date) {
//   var timeZone = Session.getScriptTimeZone();
//   var formattedDate = Utilities.formatDate(date, timeZone, "dd-MMM-yy");
//   return new Date(formattedDate);
// }








