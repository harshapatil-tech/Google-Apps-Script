function updateDashboard(spreadsheetId, startDate, endDate, invoiceNum) {
  const ss = SpreadsheetApp.openById("1PJbclFDc1-i6gZ4QAH73FyEJjVNdkzPN3hd2A_ZfF1o");
  const yogeshDashboard = ss.getSheetByName("Yogesh Dashboard");
  const dashoardSheet = ss.getSheetByName("Dashboard");
  const logSheet = ss.getSheetByName("Logs")

  const date = new Date();
  
  // Copy and protect previous sheet
  copyAndProtectRange(dashoardSheet);

  dashoardSheet.getRange(3, 2).setValue(invoiceNum);
  dashoardSheet.getRange(4, 2).setValue(Utilities.formatDate(startDate, "IST", "dd-MMM-yyyy")).setNumberFormat("dd-MMM-yyyy");
  dashoardSheet.getRange(5, 2).setValue(Utilities.formatDate(endDate, "IST", "dd-MMM-yyyy")).setNumberFormat("dd-MMM-yyyy");

  dashoardSheet.getRange("C8").setValue(Utilities.formatDate(date, "IST", "dd-MMM-yyyy")).setNumberFormat("dd-MMM-yyyy");
  dashoardSheet.getRange("C9").setValue(Utilities.formatDate(date, "IST", "dd-MMM-yyyy")).setNumberFormat("dd-MMM-yyyy");
  dashoardSheet.getRange("E8").setValue(Utilities.formatDate(date, "IST", "dd-MMM-yyyy")).setNumberFormat("dd-MMM-yyyy");
  dashoardSheet.getRange("E9").setValue(Utilities.formatDate(date, "IST", "dd-MMM-yyyy")).setNumberFormat("dd-MMM-yyyy");
  dashoardSheet.getRange("F8").setValue("Completed");
  dashoardSheet.getRange("F9").setValue("Review Pending");

  // const spreadsheet = DriveApp.getFileById(spreadsheetId)
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId)
  const item = spreadsheet.getName();
  const url = spreadsheet.getUrl();
  addEditorSilentlyV3(spreadsheetId, "sreenjay.sen@upthink.com")
  addEditorSilentlyV3(spreadsheetId, "yogesh.kadwade@upthink.com")
  const folder = DriveApp.getFolderById("1q1ZT3I3K5JKeoL-Oz3sj2IDtKGH-Rda4");
  folder.addEditor("yogesh.kadwade@upthink.com")
  // const url = spreadsheet.getUrl();
  try {
    const emailBody = createEmailBody("Automation to Admin", {item: item, spreadsheetLink: url})
    Utils.EmailUtil.sendEmail("automation@upthink.com",
                    "Automation",
                    "yogesh.kadwade@upthink.com", 
                    `New BrainFuse Invoice: ${item}`, 
                    emailBody, 
                    [], 
                    [], 
                    "automation@upthink.com",
                    )
    dashoardSheet.getRange("G9").setValue(url);
    // Update Yogesh's dashboard
    yogeshDashboardUpdate(yogeshDashboard, date, invoiceNum, url);
    // Update Log sheet
    logSheetUpdate(logSheet, invoiceNum, date);
    
  } catch (err) {
    console.log("Error occured while sending email or unable to find the dashboard sheet", err);
  }


}



function createEmailBody(templateName, variables) {
  const emailTemplate = HtmlService.createTemplateFromFile(templateName);
  // Dynamically assign variables from the object
  for (const [key, value] of Object.entries(variables)) {
    emailTemplate[key] = value;
  }
  return emailTemplate.evaluate().getContent();
}




function addEditorSilentlyV3(fileId, email) {
  const permission = {
    role:         'writer',
    type:         'user',
    emailAddress: email  // in v3 it’s emailAddress
  };
  // sendNotificationEmail: false will suppress the email
  Drive.Permissions.create(permission, fileId, { sendNotificationEmail: false,  supportsAllDrives: true  });
}




function logSheetUpdate(sheet, invoiceNum, date) {
  const [ headerMap, _ ] = Utils.DataNHeaders.getDataIndicesFromSheet(sheet);
  const lastRow = sheet.getLastRow();
  incrementSrNo(sheet, headerMap, "Sr. No.", lastRow);
  sheet.getRange(lastRow+1, headerMap["Invoice file name"]+1).setValue(invoiceNum);
  sheet.getRange(lastRow+1, headerMap["Expected Run Date"]+1).setValue(date).setNumberFormat("dd-MMM-yyyy")
  sheet.getRange(lastRow+1, headerMap["Data Creation Date"]+1).setValue(date).setNumberFormat("dd-MMM-yyyy");
}



function yogeshDashboardUpdate(sheet, startDate, invoiceNum, url) {
  const lastRow = sheet.getLastRow();
  const [ headers, data] = Utils.DataNHeaders.getDataIndicesFromSheet(sheet);

  const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();

  // Increment the sr no
  incrementSrNo(sheet, headers, "Sr. No.", lastRow);

  sheet.getRange(lastRow+1, headers["Start Date"]+1).setValue(Utilities.formatDate(startDate, "IST", "dd-MMM-yyyy")).setNumberFormat("dd-MMM-yyyy");
  sheet.getRange(lastRow+1, headers["Link"]+1).setValue(url);
  sheet.getRange(lastRow+1, headers["Item"]+1).setValue(invoiceNum);
  sheet.getRange(lastRow+1, headers["Status"]+1).setValue("Review Pending");
  sheet.getRange(lastRow+1, headers["Update?"]+1).setDataValidation(checkboxRule);
  
  const actionValidation = {range : sheet.getRange(lastRow+1, headers["Action"]+1), list: ["Pass", "Fail"]};
  const statusValidation = {range : sheet.getRange(lastRow+1, headers["Status"]+1), list: ["Review Pending", "Pending with Automation", "Review Complete"]};
  dataValidations(actionValidation.range, actionValidation.list);
  dataValidations(statusValidation.range, statusValidation.list);
}


function dataValidations(range, list) {
  const validation = SpreadsheetApp.newDataValidation().setAllowInvalid(false).requireValueInList(list).build()
  range.setDataValidation(validation)
}




function incrementSrNo(sheet, headerMap, headerVal, lastRow) {

  const columnIdx = headerMap[headerVal];
  const prevRange = sheet.getRange(lastRow, columnIdx+1)

  const range = sheet.getRange(lastRow+1, columnIdx+1)
  if (lastRow === 1) {
    range.setValue(1);
  } else {
    range.setFormula(`${prevRange.getA1Notation()}+1`)
  }
}













