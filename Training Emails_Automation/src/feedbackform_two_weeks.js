function emailAfterFourteenDaysOfTrainingCompletion() {

  const RUNNING_DATE = new Date();
  RUNNING_DATE.setHours(0, 0, 0, 0);

  const CURRENT_SPREADSHEET = SpreadsheetApp.openById("1C2OaNTqEZd07JDsUF7jhqZNrrZpcVrFaG6ygSAsBIww");
  const TRAINING_TRACKER_SHEET = CURRENT_SPREADSHEET.getSheetByName("Training Tracker");
  let [headers_TrainingTracker, data_TrainingTracker] = getDataIndicesFromSheetWithStartingRow(TRAINING_TRACKER_SHEET, 3);

  data_TrainingTracker.forEach((row, index) => {
      const employeeName = row[headers_TrainingTracker["Employee Name"]];
      const employeeEmailId = row[headers_TrainingTracker["Email ID"]];
      const moodleCredentials = row[headers_TrainingTracker["Moodle Login ID"]];
      const password = row[headers_TrainingTracker["Password"]];
      const department = row[headers_TrainingTracker["Department"]];
      const trainingStartDate = row[headers_TrainingTracker["Training Start Date"]];
      const trainingStartTime = row[headers_TrainingTracker["Training Start Time"]];
      const trainingEndDate = row[headers_TrainingTracker["Training End Date"]];
      const trainingEmailCheckBoxBool = row[headers_TrainingTracker["Training Email"]]; 
      const feedbackEmailSentBool = row[headers_TrainingTracker["Feedback Email Sent?"]];
      const trainingEmailSent = row[headers_TrainingTracker["Email Sent?"]]; 
      const moodleCredentialsEmailSent = row[headers_TrainingTracker["Credentials Sent?"]];

      const trainingEndDateObj = new Date(trainingEndDate)

      if (feedbackEmailSentBool === "" && trainingEndDate !== "" && trainingEndDate !== undefined && !isNaN(trainingEndDateObj.getTime())){
        
        var pastTwoWeeksFromTrainingEndDate = new Date(trainingEndDateObj.getTime() + 14 * 24 * 60 * 60 * 1000);
        const formSubmitDate = new Date(pastTwoWeeksFromTrainingEndDate.getTime() + 5 * 24 * 60 * 60 * 1000);
        
        const sendBool = CentralLibrary.getDaysDifference(pastTwoWeeksFromTrainingEndDate, RUNNING_DATE) === 0;
        if (sendBool === true){
          if (employeeEmailId === undefined || employeeEmailId === "") {
            GmailApp.sendEmail("automation@upthink.com", "Training Plan Asmita", "Email id missing")
          } else{
            const formattedTrainingEndDate = Utilities.formatDate(trainingEndDateObj, "IST", "dd-MMM-YYYY")
            const formattedFormSubmitDate = Utilities.formatDate(formSubmitDate, "IST", "dd-MMM-YYYY")
            trainingFeedbackEmail(employeeName, employeeEmailId, formattedTrainingEndDate, formattedFormSubmitDate);
            TRAINING_TRACKER_SHEET.getRange(index+5, headers_TrainingTracker["Feedback Email Sent?"] + 1).setValue("Y");
          }
        }

      }


        
    });


}





function trainingFeedbackEmail(employeeName, employeeEmailId, trainingEndDate, submissionDate) {
    const template = HtmlService.createTemplateFromFile("Training Feedback.html");
    
    const trainingFeedbackFormLink = "https://docs.google.com/forms/d/1Oy9Sa_WeH9y-PxYQQznWlyFfzvgTkGhKULECVYYDebQ/edit";

    template.employeeName = employeeName;
    template.trainingEndDate = trainingEndDate;
    template.submissionDate = submissionDate;
    template.trainingFeedbackFormLink = trainingFeedbackFormLink;
    const messageBody = template.evaluate().getContent();
    const subjectLine = "Training Feedback Form.";

    const options = {
      htmlBody: messageBody,
      from: "asmita.sane@upthink.com",
    };
    
    GmailApp.sendEmail(employeeEmailId, subjectLine, "", options)
}



function getDataIndicesFromSheetWithStartingRow(sheet, startRow) {
    let dataRange = sheet.getDataRange().getValues();
    dataRange = dataRange.slice(startRow);
    const headers = dataRange[0];
    const data = dataRange.slice(1);
    return [this.createIndexMap(headers), data];
}

function getDataIndicesFromSheet(sheet) {
    const dataRange = sheet.getDataRange().getValues();
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