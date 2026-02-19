function submit() {
  const emailsClass = new Emails();
  emailsClass.sendEmails();
}



class Emails {

  constructor() {
    this.CURRENT_SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
    this.TRAINING_TRACKER_SHEET = Training.getSheetById(this.CURRENT_SPREADSHEET, 0)
    this.BACKEND_SHEET = Training.getSheetById(this.CURRENT_SPREADSHEET, 1240090851);
  }

  static getSheetById (spreadsheet, sheetId) {
    return spreadsheet.getSheets().find(sheet => sheet.getSheetId() === sheetId);
  }

  sendEmails() {
    let departmentMapper = makeOwnKeyValuePairs(this.BACKEND_SHEET, "Department", "Trainer's Email IDs", "Team Lead Email ID");

    let [headers_TrainingTracker, data_TrainingTracker] = this.getDataIndicesFromSheetWithStartingRow(this.TRAINING_TRACKER_SHEET, 3);

    const emailIdMissingList = []
    let emailIdMissingBool = false;

    const trainingEmailList = [];
    let trainingEmailRequiredBool = false;

    const trainingEmailAlreadySentList = [];
    let trainingEmailAlreadySentBool = false;


    const moodleCredSentList = [];
    let moodleCredSentBool = false;

    const moodleCredList = [];
    let moodleCredBool = false;

    data_TrainingTracker.forEach((row, index) => {
      const employeeName = row[headers_TrainingTracker["Employee Name"]];
      const employeeEmailId = row[headers_TrainingTracker["Email ID"]];
      const moodleCredentials = row[headers_TrainingTracker["Moodle Login ID"]];
      const password = row[headers_TrainingTracker["Password"]];
      const department = row[headers_TrainingTracker["Department"]];
      const trainingStartDate = row[headers_TrainingTracker["Training Start Date"]];
      const trainingStartTime = row[headers_TrainingTracker["Training Start Time"]];
      const trainingEndDate = row[headers_TrainingTracker["Training End Date"]];
      const subjectTrainersArray = departmentMapper[department].trainerSEmailIds.split(",")
      const subjectLeadArray = departmentMapper[department].teamLeadEmailId.split(",")
      const trainingEmailCheckBoxBool = row[headers_TrainingTracker["Training Email"]]; 
      const moodleCredentialsCheckBoxBool = row[headers_TrainingTracker["Moodle Credentials Email"]];
      const trainingEmailSent = row[headers_TrainingTracker["Email Sent?"]]; 
      const moodleCredentialsEmailSent = row[headers_TrainingTracker["Credentials Sent?"]];
      const location = row[headers_TrainingTracker["Location"]];
      const contact = row[headers_TrainingTracker["Personal Contact Number"]];

      if (employeeEmailId === undefined || employeeEmailId === "") {
        emailIdMissingList.push(employeeEmailId);
        emailIdMissingBool = true;
      } else{
        // Training Email

        if(trainingEmailCheckBoxBool === true) { // &&
          // console.log(trainingEmailSent)
          if(trainingStartDate === undefined || trainingStartDate === "" || trainingStartTime === "" || trainingStartTime === undefined){
            trainingEmailList.push(employeeEmailId);
            trainingEmailRequiredBool = true;
          } else if(trainingEmailSent === "Y") {
            // console.log("traing email sent", trainingEmailSent)
            trainingEmailAlreadySentList.push(employeeEmailId);
            trainingEmailAlreadySentBool = true;
          } else if((trainingEmailSent === undefined || trainingEmailSent === "")){
            this.trainingEmail(employeeName, employeeEmailId, 
                        trainingStartDate, trainingStartTime, 
                        subjectTrainersArray, subjectLeadArray, location, contact
                        );
            this.TRAINING_TRACKER_SHEET.getRange(index+5, headers_TrainingTracker["Email Sent?"] + 1).setValue("Y")
            this.TRAINING_TRACKER_SHEET.getRange(index+5, headers_TrainingTracker["Training Email"] + 1).setValue(false);
          }
          
        }
        // Moodle Credentials
        if(moodleCredentialsCheckBoxBool === true && (moodleCredentialsEmailSent === undefined || moodleCredentialsEmailSent === "")) {
          this.moodleCredentialsEmail(employeeName, employeeEmailId, moodleCredentials, password, subjectTrainersArray);
          this.TRAINING_TRACKER_SHEET.getRange(index+5, headers_TrainingTracker["Credentials Sent?"] + 1).setValue("Y");
          this.TRAINING_TRACKER_SHEET.getRange(index+5, headers_TrainingTracker["Moodle Credentials Email"] + 1).setValue(false);
        }else if (moodleCredentialsCheckBoxBool === true && (moodleCredentialsEmailSent === "Y" || moodleCredentialsEmailSent === "Yes")) {
            moodleCredSentList.push(employeeEmailId);
            moodleCredSentBool = true;
        } else if(moodleCredentials === undefined || moodleCredentials === "" || password === undefined || password === "") {
          moodleCredList.push(employeeEmailId);
          moodleCredBool = true;
        }
      }
        
    });

    if(emailIdMissingBool){
      this.alertMessage("Incomplete data \nThe Email ID is missing for", emailIdMissingList);
    }

    if(trainingEmailRequiredBool){
      this.alertMessage("Required fields are incomplete for", trainingEmailList);
    }

    if(trainingEmailAlreadySentBool){
      this.alertMessage("This email has already been sent to", trainingEmailAlreadySentList);
    }    

    if(moodleCredSentBool){
      this.alertMessage("This email has already been sent to", moodleCredSentList);
    }

    if (moodleCredBool){
      this.alertMessage("Required fields are incomplete for", moodleCredList);
    }


  }

  alertMessage(message, emailIdArray) {
    const emailIds = emailIdArray.join(", ");
    message = message + " " + emailIds
    SpreadsheetApp.getUi().alert(message);
  }

  trainingEmail(employeeName, employeeEmailId, trainingStartDate, trainingStartTime, subjectTrainersArray, subjectLeadArray,location,contact) {
    const template = HtmlService.createTemplateFromFile("Training Email.html");
    
    trainingStartDate = Utilities.formatDate(trainingStartDate, "IST", "dd-MMM-YYYY");
    trainingStartTime = Utilities.formatDate(trainingStartTime, "IST", "hh:mm a")

    template.employeeName = employeeName;
    template.trainingStartDate = trainingStartDate;
    template.trainingStartTime = trainingStartTime;
    template.subjectTrainersArray = subjectTrainersArray;
    template.subjectLeadArray = subjectLeadArray
    template.location = location;
    template.personalContactNumber = contact;

    const messageBody = template.evaluate().getContent();
    const subjectLine = "Training Commencement Information.";

    const combinedCCList = subjectTrainersArray.concat(subjectLeadArray);
    const ccEmails = combinedCCList.join(',');
    const options = {
      htmlBody: messageBody,
      // from: "automation@upthink.com",
      from: "asmita.sane@upthink.com",
      cc: ccEmails,
    };
    
    GmailApp.sendEmail(employeeEmailId, subjectLine, "", options);  //
  }


  moodleCredentialsEmail(employeeName, employeeEmailId, loginCredentials, password, subjectTrainersArray) {

    const template = HtmlService.createTemplateFromFile("Moodle Credentials.html");
    template.employeeName = employeeName;
    template.loginCredentials = loginCredentials;
    template.password = password;
    const messageBody = template.evaluate().getContent();
    const subjectLine = "Moodle Account Credentials.";
    // Add asmita.sane@upthink.com to the cc list
    subjectTrainersArray.push("asmita.sane@upthink.com");
    const ccEmails = subjectTrainersArray.join(',');
    const options = {
      htmlBody: messageBody,
      cc: ccEmails,
      from: "asmita.sane@upthink.com",
    };
    
    GmailApp.sendEmail(employeeEmailId, subjectLine, "", options)
  }


  getDataIndicesFromSheetWithStartingRow(sheet, startRow) {
    let dataRange = sheet.getDataRange().getValues();
    dataRange = dataRange.slice(startRow);
    const headers = dataRange[0];
    const data = dataRange.slice(1);
    return [this.createIndexMap(headers), data];
  }

  getDataIndicesFromSheet(sheet) {
    const dataRange = sheet.getDataRange().getValues();
    const headers = dataRange[0];
    const data = dataRange.slice(1);
    return [this.createIndexMap(headers), data];
  }

  createIndexMap(headers) {
    return headers.reduce((map, val, index) => {
      if (val !== "") {
        map[val.trim()] = index;
      }
      return map;
    }, {});
  }
}
