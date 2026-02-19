function createTrigger() {
  // Delete any existing triggers for 'doThings' to prevent duplicates
  CentralLibrary.deleteTriggers('doThings');

  // Create a new time-based trigger for 'doThings'
  ScriptApp.newTrigger('doThings')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.TUESDAY)
    .atHour(23)
    .create();
}

function doThings() {

  const scriptProperties = PropertiesService.getScriptProperties();
  const lastExecution = scriptProperties.getProperty('lastExecutionDate');
  const today = new Date();

  // If 'lastExecutionDate' is not set, initialize it
  if (!lastExecution) {
    scriptProperties.setProperty('lastExecutionDate', today.toDateString());
    // Proceed with executing the function for the first time
  } else {
    const lastDate = new Date(lastExecution);
    const diffInDays = Math.floor((today - lastDate) / (1000 * 60 * 60 * 24));

    // Check if 14 days (2 weeks) have passed
    if (diffInDays < 11) {
      // Less than two weeks have passed; do not execute
      console.log("Funcion already executed once within 14 days")
      return;
    } else {
      // Update the last execution date
      scriptProperties.setProperty('lastExecutionDate', today.toDateString());
    }
  }

  const {startDate, endDate} = calculatePreviousTwoWeeksDateRange();
  // = getDateDaysAgo(16);
  // const pastDateEnd = getDateDaysAgo(2);
  const options = { day: '2-digit', month: 'short', year: '2-digit' };
  const formattedStartDate = startDate.toLocaleDateString('en-GB', options).replace(/ /g, ' ');
  const formattedEndDate = endDate.toLocaleDateString('en-GB', options).replace(/ /g, ' ');
  console.log(formattedStartDate, formattedEndDate, startDate, endDate)
  const spreadsheetName = `QMS_English Essay_Biweekly_MIS_Output [${formattedStartDate} - ${formattedEndDate}]`;
  const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet().getId();
  const newSheetId = CentralLibrary.copySpreadsheetToFolder("1Uz9ntmNgiXO1Wc828Mgm_P4eoKueSViV", currentSpreadsheet, spreadsheetName);
  sendFormattedEmail(newSheetId);
  
}




function sendFormattedEmail(spreadsheetId) {
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName('MIS'); // Adjust to your sheet name
  const link = spreadsheet.getUrl();
  // Extracting numbers for Section 1: BF Timesheet vs QMS
  const bfQmsData = [
    sheet.getRange('C7:F8').getValue(),  // # QMS
    sheet.getRange('H7:K8').getValue(),  // # BF
    sheet.getRange('M7:P8').getValue()   // Difference
  ];

  // Extracting numbers for Section 2: QMS vs Manual Entry
  const qmsManualData = [
    sheet.getRange('C15:F16').getValue(), // # QMS
    sheet.getRange('H15:K16').getValue(), // # Manual
    sheet.getRange('M15:P16').getValue()  // Difference
  ];

  // Prepare email content
  const emailContent = HtmlService.createTemplateFromFile('EmailTemplate');
  emailContent.bfQmsData = bfQmsData;
  emailContent.qmsManualData = qmsManualData;
  emailContent.link = link;

  const subject = 'MIS Biweekly Report';
  const recipient = 'deboo.roy@upthink.com';  // Adjust recipient email
  const ccList = ["apurva.yadav@upthink.com", "tejas.jagtap@upthink.com"].join(',');
  // Generate the HTML content and send the email
  const message = emailContent.evaluate().getContent();
  GmailApp.sendEmail(recipient, subject, '', {
    htmlBody: message,
    cc: ccList
  });
  CentralLibrary.shareSpreadsheet(spreadsheet, recipient, 'reader');
  CentralLibrary.shareSpreadsheet(spreadsheet, "apurva.yadav@upthink.com", 'reader');
  CentralLibrary.shareSpreadsheet(spreadsheet, "tejas.jagtap@upthink.com", 'reader');
}



function calculatePreviousTwoWeeksDateRange() {
  var today = new Date();
  var daysSinceSunday = today.getDay(); // 0 (Sunday) to 6 (Saturday)
  var lastSunday = new Date(today.getFullYear(), today.getMonth(), today.getDate() - daysSinceSunday);
  var startDate = new Date(lastSunday.getFullYear(), lastSunday.getMonth(), lastSunday.getDate() - 13); // 13 days before last Sunday
  return {
    startDate: startDate,
    endDate: lastSunday
  };
}




function getDateDaysAgo(numDays) {
  var today = new Date();
  var pastDate = new Date();
  pastDate.setDate(today.getDate() - numDays); // Subtracts 15 days from today

  return pastDate;
}
