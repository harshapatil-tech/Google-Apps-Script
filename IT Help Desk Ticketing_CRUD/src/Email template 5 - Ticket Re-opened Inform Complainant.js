// function mainTemplate5(){
//   const spreadsheet = SpreadsheetApp.openById("1iByitSy5R35cu13rupuppctpzV0X8dTzqDSJJP2ilAk");
//   const sheet = spreadsheet.getSheetByName("Master DB");

//   const[dbHeaders, dbData] = CentralLibrary.get_Data_Indices_From_Sheet(sheet);
//   const row = dbData[0];
//   // console.log(row);
  
//   emailTemplate5(sheet, dbHeaders, row, 0);
// }

function emailTemplate5(sheet, headers, row, i){
  const ticketNoIndex = headers["Ticket No."];
  const recipientEmailIndex = headers["Email Address"];
  const laptopNumberIndex = headers["Laptop Number"];
  const itSupportDiagnosisIndex = headers["IT Support diagnosis"];
  const issueDescriptionIndex = headers["Issue Description"];
  const issueStatusIndex = headers["Issue Status"];
  const finalClosureDateIndex = headers["Final Closure Date"];
  const closureEmailDateIndex = headers["Closure Email Date"];
  const reopenTicketDateIndex = headers["Re-open Ticket Date"];
  const departmentIndex = headers["Department"];

  const today = new Date();
  const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd-MMM-yyyy");

  const issueStatus = row[issueStatusIndex];
  const closureEmailDate = row[closureEmailDateIndex];
  const finalClosureDate = row[finalClosureDateIndex];
  const ticketNumber = row[ticketNoIndex];
  const recipientEmail = row[recipientEmailIndex];
  const reopenTicketDate = row[reopenTicketDateIndex];
  const department = row[departmentIndex];

  if (issueStatus === "5.1 Re-open ticket" && closureEmailDate !="" && finalClosureDate === ""){
    const emailVariables = {
         ticketNumber : row[ticketNoIndex],
         emailAddress : row[recipientEmailIndex],
         laptopNumber : row[laptopNumberIndex],
         itSupportDiagnosis : row[itSupportDiagnosisIndex],
         issueDescription : row[issueDescriptionIndex]
      };

    // Update the Issue Status column
    sheet.getRange(i + 2, issueStatusIndex + 1).setValue("5.2 Ticket reopened [A]");
    // Set the Reopen Ticket Date 
    sheet.getRange(i + 2, reopenTicketDateIndex + 1).setValue(formattedDate).setNumberFormat("dd-MMM-yyyy");
    
    const ccListEmail = [ccEmailLists[department], SUPPORT_CC];
    const emailSubject = `Your Ticket No. ${ticketNumber} has been re-opened.`;
    const emailBody = createEmailBody("Email template 5.html", emailVariables);
    sendEmail(recipientEmail, emailSubject, emailBody, ccList = ccListEmail);
  }

  

}