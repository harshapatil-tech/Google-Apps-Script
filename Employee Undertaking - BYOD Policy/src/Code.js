const SPREADSHEET = SpreadsheetApp.openById("1D0NOkjL5kOuzG3tzg4H4PdcTqIuaiZRFEn1DhoR3xJg");
const FORM_RESPONSES_SHEET_ID = 1232633093;
const BYOD_UNDERTAKING_STATUS_ID = 1822982685;
const START_DATE = new Date("2023-12-01")
START_DATE.setHours(0, 0, 0, 0);
const TODAY = new Date();
TODAY.setHours(0, 0, 0, 0);

function onOpen() {
  var ui = SpreadsheetApp.getUi();

  // Or DocumentApp or FormApp.
  ui.createMenu('Custom Menu')
      .addItem('RUN', 'myFunction')
      .addToUi();
}


function myFunction() {
  
  const byodFormResponseSheet = getSheetById(SPREADSHEET, FORM_RESPONSES_SHEET_ID);
  const undertakingStatusSheet = getSheetById(SPREADSHEET, BYOD_UNDERTAKING_STATUS_ID);

  const [byodIndices, byodData] = get_Data_Indices_From_Sheet(byodFormResponseSheet);
  const [undertakingIndices, undertakingData] = get_Data_Indices_From_Sheet(undertakingStatusSheet);

  const byodEmailAdresses = byodData.map(row => row[byodIndices["Email Address"]].trim().toLowerCase());

  undertakingData.forEach( function(row, index) {

    const emailAddress = row[undertakingIndices["Official email ID"]].trim().toLowerCase();
    const reportingManagerEmail = row[undertakingIndices["Reporting Manager Email ID"]]
    const hodEmail = row[undertakingIndices["Department Head Email ID"]]
    const underTakingCheckBoxRange = undertakingStatusSheet.getRange(index + 2, undertakingIndices["Undertaking"] + 1)

    if (emailAddress !== '' && reportingManagerEmail !== '' && hodEmail !== '') {
    
      if(byodEmailAdresses.includes(emailAddress)) {
        // If there in the form responses update the status column
        underTakingCheckBoxRange.check();
        undertakingStatusSheet.getRange(index + 2, undertakingIndices["Status"] + 1).setValue("Agreed");
      }

      else {
        
        // If the byod email not checked then send first email and then check the checkbox
        const name = row[undertakingIndices["Employee Name"]];
        const doj = row[undertakingIndices["DOJ"]];

        const ccEmails = [reportingManagerEmail, hodEmail]
        const byodEmailRange = undertakingStatusSheet.getRange(index + 2, undertakingIndices["BYOD Email"] + 1);
        const firstReminderEmailRange = undertakingStatusSheet.getRange(index + 2, undertakingIndices["Reminder Email 1"] + 1);
        const secondReminderEmailRange = undertakingStatusSheet.getRange(index + 2, undertakingIndices["Reminder Email 2"] + 1);

        const oneDayAfter = daysInAdvance(doj, 1);
        const twoDaysAfter = daysInAdvance(doj, 2);
        const threeDaysAfter = daysInAdvance(doj, 3);
        const fourDaysAfter = daysInAdvance(doj, 4);
        const fiveDaysAfter = daysInAdvance(doj, 5);
        const sixDaysAfter = daysInAdvance(doj, 6);
        const sevenDaysAfter = daysInAdvance(doj, 7);

        if(TODAY.getTime() >= oneDayAfter.getTime() && !underTakingCheckBoxRange.isChecked() && !byodEmailRange.isChecked()) {
          
          send_Email(emailAddress, 
                    name, 
                    Utilities.formatDate(threeDaysAfter, Session.getScriptTimeZone(), "EEEE, dd-MMM-yyyy"), 
                    "BYOD_Email_Template.html", 
                    ccEmails,
                    "BYOD Policy and Employee Productivity Software."
                    );
          byodEmailRange.check();

        }else if (TODAY.getTime() === fourDaysAfter.getTime() && !underTakingCheckBoxRange.isChecked() && byodEmailRange.isChecked() 
                    && !firstReminderEmailRange.isChecked()) {

          send_Email(emailAddress, 
                    name, 
                    Utilities.formatDate(fiveDaysAfter, Session.getScriptTimeZone(), "EEEE, dd-MMM-yyyy"),
                    "First_Reminder_Template.html",
                    ccEmails,
                    "Reminder: BYOD Policy Document & Undertaking Form.");
                    
          firstReminderEmailRange.check();

        }else if(TODAY.getTime() === sixDaysAfter.getTime() && !underTakingCheckBoxRange.isChecked() && byodEmailRange.isChecked() && 
                  firstReminderEmailRange.isChecked() && !secondReminderEmailRange.isChecked()) {

          send_Email(emailAddress, 
                    name, 
                    Utilities.formatDate(sevenDaysAfter, Session.getScriptTimeZone(), "EEEE, dd-MMM-yyyy"), 
                    "Second_Reminder_Template.html",
                    ccEmails,
                    "Reminder: BYOD Policy Document & Undertaking Form.");

          secondReminderEmailRange.check();
        }
        undertakingStatusSheet.getRange(index + 2, undertakingIndices["Status"] + 1).setValue("Pending");
      }
    }
  })
  

}



function byodNotComplete() {

  const undertakingStatusSheet = getSheetById(SPREADSHEET, BYOD_UNDERTAKING_STATUS_ID);

  const [undertakingIndices, undertakingData] = get_Data_Indices_From_Sheet(undertakingStatusSheet);

  const notSubmitted = {}

  undertakingData.forEach((r, index) => {

    const byodEmail = r[undertakingIndices["BYOD Email"]];
    const firstReminderEmail = r[undertakingIndices["Reminder Email 1"]];
    const secondReminderEmail = r[undertakingIndices["Reminder Email 2"]];
    const undertakingEmail = r[undertakingIndices["Undertaking"]];
    const employeeName = r[undertakingIndices["Employee Name"]];
    const hodEmail = r[undertakingIndices["Department Head Email ID"]];
    const reportingManagerEmail = r[undertakingIndices["Reporting Manager Email ID"]];
    const hodName = r[undertakingIndices["Department Head Name"]];

    if (byodEmail && firstReminderEmail && secondReminderEmail && !undertakingEmail) {

      if(!notSubmitted.hasOwnProperty(hodEmail)) {
        notSubmitted[hodEmail] = {
          "hodName" : hodName,
          "reportingManagerEmails" : [],
          "Employee Names" : []
        }
      }
      notSubmitted[hodEmail]["reportingManagerEmails"].push(reportingManagerEmail);
      notSubmitted[hodEmail]["Employee Names"].push(employeeName);
    }

  })

  // Loop over not submitted object
  for (const [key, value] of Object.entries(notSubmitted)) {
    console.log(key, value)
    send_Email_HOD(key,  value.hodName, "BYOD_Not_Completed_Template.html", value)
  }

  
}


function send_Email_HOD(hodEmail, hodName, htmlFile, object) {
  const template = HtmlService.createTemplateFromFile(htmlFile);
  template.name = hodName;
  template.employeeNames = object["Employee Names"];

  const messageBody = template.evaluate().getContent();
  const recipientEmail = hodEmail.toString();

  const options = {
    htmlBody: messageBody,
    from: "hrd@upthink.com",
    cc: object["reportingManagerEmails"].join(", ")
  };

  // GmailApp.sendEmail(recipientEmail, "BYOD Policy Undertaking Status.", "This is a fallback plain text body", options);
}




function send_Email(sme_Email, name, date, htmlFile, ccEmail, subjectLine, pdfFile=true) {
  
  const template = HtmlService.createTemplateFromFile(htmlFile);
  template.name = name;
  template.date = date;
  const messageBody = template.evaluate().getContent();
  const recipientEmail = sme_Email.toString();

  const options = {
    htmlBody: messageBody,
    from: "hrd@upthink.com",
    cc : ccEmail.join(", ")
  };

  // Attach PDF file  10C1H0gRd1_uWx4NDsFHFmeFKlPKcpHs4
  if (pdfFile) {
    const pdfBlob = DriveApp.getFileById("10C1H0gRd1_uWx4NDsFHFmeFKlPKcpHs4").getBlob();
    
    // Specify mimeType and fileName for the attachment
    options.attachments = [{
      fileName: "Bring Your Own Device Policy.pdf",
      content: pdfBlob.getBytes(),
      mimeType: "application/pdf"
    }];
  }

  // GmailApp.sendEmail(recipientEmail, subjectLine, "This is a fallback plain text body", options);
}





function getSheetById(spreadsheet, id) {
  return spreadsheet.getSheets().filter(
    function(s) {return s.getSheetId() === id;}
  )[0];
}


function get_Data_Indices_From_Sheet(sheet) {

  const dataRange = sheet.getDataRange().getValues();
  const headers = dataRange[0], data = dataRange.slice(1);
  return [createMapIndex(headers), data]
}


function createMapIndex(headers) {

  return headers.reduce((mapObj, currVal, currIndex) => {
    mapObj[currVal.trim()] = currIndex;
    return mapObj;
  }, {})
}

function daysInAdvance(today, numOfDays) {
  today = new Date(today)
  today.setHours(0, 0, 0, 0)
  // Calculate the date two days from now
  var laterDate = new Date(today);

  laterDate.setDate(today.getDate() + numOfDays);
  laterDate.setHours(0, 0, 0, 0);
  return laterDate;
}

