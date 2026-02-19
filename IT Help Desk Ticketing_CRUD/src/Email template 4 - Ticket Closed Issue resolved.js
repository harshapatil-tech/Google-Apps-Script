function emailTemplate4(sheet, headers, row, i){
  // Get column indices
  const ticketNoIndex = headers["Ticket No."];
  console.log(ticketNoIndex);
  const recipientEmailIndex = headers["Email Address"];
  const issueStatusIndex = headers["Issue Status"];
  const initialEmailDateIndex = headers["Initial Email Date"];
  const closureEmailDateIndex = headers["Closure Email Date"];
  const departmentIndex = headers["Department"];
  const raiseTicketDateIndex = headers["Raise Ticket Date"];
  
  const today = new Date();
  const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd-MMM-yyyy");

  const issueStatus = row[issueStatusIndex];
  const initialEmailDate = row[initialEmailDateIndex];
  const closureEmailDate = row[closureEmailDateIndex];
  const ticketNumber = row[ticketNoIndex];
  const recipientEmail = row[recipientEmailIndex];
  const department = row[departmentIndex];
  const raiseTicketDate = row[raiseTicketDateIndex];

  if (issueStatus === "4.1 Close Ticket" && initialEmailDate !== "" && closureEmailDate === ""){  // && raiseTicketDate ===""
    const emailVariables = { ticketNumber };
    console.log("Here")
    sheet.getRange(i + 2, issueStatusIndex + 1).setValue("4.2 Ticket closed [A]");
    sheet.getRange(i + 2, closureEmailDateIndex + 1).setValue(formattedDate).setNumberFormat("dd-MMM-yyyy");
    
    const ccListEmail = [ccEmailLists[department], SUPPORT_CC];
    const emailSubject = `Your Ticket No. ${ticketNumber} has been resolved.`;
    const emailBody = createEmailBody("Email template 4.html", emailVariables);
    sendEmail(recipientEmail, emailSubject, emailBody, ccList = ccListEmail);
  }

}
