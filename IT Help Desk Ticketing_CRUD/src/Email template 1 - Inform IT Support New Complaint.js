// function mainTemplate1(){
//   const spreadsheet = SpreadsheetApp.openById("1iByitSy5R35cu13rupuppctpzV0X8dTzqDSJJP2ilAk");
//   const sheet = spreadsheet.getSheetByName("Master DB");

//   const[dbHeaders, dbData] = CentralLibrary.get_Data_Indices_From_Sheet(sheet);
//   const row = dbData[0];
//   // console.log(row);
  
//   emailTemplate1(sheet, dbHeaders, row, 0);
// }

// function emailTemplate1(sheet, headers, row, i) {
//   if (!headers || !headers["Employee Name"]) {
//     Logger.log("Error: 'Employee Name' header not found. Headers: " + JSON.stringify(headers));
//     return;
//   }

//   const employeeNameIndex = headers["Employee Name"];
//   const ticketNoIndex = headers["Ticket No."];
  
//   if (!row[employeeNameIndex]) {
//     Logger.log(`Row ${i + 1}: 'Employee Name' column is empty or row data is invalid.`);
//     return;
//   }

//   const employeeName = row[employeeNameIndex];
//   Logger.log(`Processing Employee: ${employeeName}`);

//   const recipientEmailIndex = headers["Email Address"];
//   const contactNoIndex = headers["Contact Number"];
//   const issueTypeIndex = headers["Issue Type"];
//   const issueDescriptionIndex = headers["Issue Description"];
//   const requestedDateIndex = headers["Date"];
//   const requestedTimeIndex = headers["Time"];
//   const initialEmailDateIndex = headers["Initial Email Date"];
//   const issueStatusIndex = headers["Issue Status"];

//   const today = new Date();
//   const ticketNumber = row[ticketNoIndex];
//   const initialEmailDate = row[initialEmailDateIndex];
//   const recipientEmail = row[recipientEmailIndex];

//   if (ticketNumber !== "" && !initialEmailDate) {
//     const emailVariables = {
//       employeeName: row[employeeNameIndex],
//       ticketNumber: ticketNumber,
//       contactNumber: row[contactNoIndex],
//       issueType: row[issueTypeIndex],
//       issueDescription: row[issueDescriptionIndex],
//       requestedDate: Utilities.formatDate(row[requestedDateIndex], Session.getScriptTimeZone(), "dd-MMM-yyyy"),
//       requestedTime: Utilities.formatDate(row[requestedTimeIndex], Session.getScriptTimeZone(), "hh:mm a"),
//     };

//     sheet.getRange(i + 2, issueStatusIndex + 1).setValue("1.1 Open");
//     sheet.getRange(i + 2, initialEmailDateIndex + 1)
//       .setValue(Utilities.formatDate(today, Session.getScriptTimeZone(), "dd-MMM-yyyy"))
//       .setNumberFormat("dd-MMM-yyyy");

//     const emailSubject = `New complaint received Ticket No. ${ticketNumber} has been raised.`;
//     const emailBody = createEmailBody("email template 1.html", emailVariables);
//     sendEmail(recipientEmail, emailSubject, emailBody);
//   }
// }

