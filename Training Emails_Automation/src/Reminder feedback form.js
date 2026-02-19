function emailAfterTwentyTwoDaysOfTrainingCompletion() {

  const RUNNING_DATE = new Date();
  RUNNING_DATE.setHours(0, 0, 0, 0);

  const CURRENT_SPREADSHEET = SpreadsheetApp.openById("1C2OaNTqEZd07JDsUF7jhqZNrrZpcVrFaG6ygSAsBIww");
  const TRAINING_TRACKER_SHEET = CURRENT_SPREADSHEET.getSheetByName("Training Tracker");
  let [headers_TrainingTracker, data_TrainingTracker] = getDataIndicesFromSheetWithStartingRow(TRAINING_TRACKER_SHEET, 3);

  data_TrainingTracker.forEach((row, index) => {
      const employeeName = row[headers_TrainingTracker["Employee Name"]];
      const employeeEmailId = row[headers_TrainingTracker["Email ID"]];
      const trainingEndDate = row[headers_TrainingTracker["Training End Date"]];
      const feedbackReceivedBool = row[headers_TrainingTracker["Feedback Received"]]; 
      const feedbackEmailSentBool = row[headers_TrainingTracker["Feedback Email Sent?"]];
      
      const trainingEndDateObj = new Date(trainingEndDate);

      if (feedbackReceivedBool === false && feedbackEmailSentBool === true && !isNaN(trainingEndDateObj.getTime())) {
        
        // 22 days after training end date
        var afterTwentyTwoDaysFromTrainingEndDate = new Date(trainingEndDateObj.getTime() + 22 * 24 * 60 * 60 * 1000);

        // Check if today is 22 days after the training end date
        const sendBool = CentralLibrary.getDaysDifference(afterTwentyTwoDaysFromTrainingEndDate, RUNNING_DATE) === 0;

        if (sendBool === true) {
          if (employeeEmailId === undefined || employeeEmailId === "") {
            GmailApp.sendEmail("automation@upthink.com", "Training Plan Asmita", "Email id missing");
          } else {
            const formattedTrainingEndDate = Utilities.formatDate(trainingEndDateObj, "IST", "dd-MMM-YYYY");
            const formattedSendDate = Utilities.formatDate(afterTwentyTwoDaysFromTrainingEndDate, "IST", "dd-MMM-YYYY");
            reminderFeedback(employeeName, employeeEmailId, formattedTrainingEndDate, formattedSendDate);
          }
        }
      }
    });
}

function reminderFeedback(employeeName, employeeEmailId, trainingEndDate, submissionDate) {
  const template = HtmlService.createTemplateFromFile("Reminder Email.html");
  
  const trainingFeedbackFormLink = "https://docs.google.com/forms/d/1Oy9Sa_WeH9y-PxYQQznWlyFfzvgTkGhKULECVYYDebQ/edit";

  template.employeeName = employeeName;
  template.trainingEndDate = trainingEndDate;
  template.submissionDate = submissionDate;
  template.trainingFeedbackFormLink = trainingFeedbackFormLink;
  const messageBody = template.evaluate().getContent();
  const subjectLine = "Reminder: Training Feedback Form";

  const options = {
    htmlBody: messageBody,
    from: "asmita.sane@upthink.com",
  };

  GmailApp.sendEmail(employeeEmailId, subjectLine, "", options);
}

function getDataIndicesFromSheetWithStartingRow(sheet, startRow) {
  let dataRange = sheet.getDataRange().getValues();
  dataRange = dataRange.slice(startRow);
  const headers = dataRange[0];
  const data = dataRange.slice(1);
  return [this.createIndexMap(headers), data];
}

function createIndexMap(headers) {
  return headers.reduce((map, val, index) => {
    if (val !== "") {
      map[val.trim()] = index;
    }
    return map;
  }, {});
}
