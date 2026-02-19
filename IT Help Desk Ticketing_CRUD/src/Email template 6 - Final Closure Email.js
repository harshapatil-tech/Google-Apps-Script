function emailTemplate6(sheet, headers, row, i){
  const ticketNoIndex = headers["Ticket No."];
  const recipientEmailIndex = headers["Email Address"];
  const issueStatusIndex = headers["Issue Status"];
  const closureEmailDateIndex = headers["Closure Email Date"];
  const finalClosureDateIndex = headers["Final Closure Date"];
  const departmentIndex = headers["Department"];

  const today = new Date();
  const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd-MMM-yyyy");

  const issueStatus = row[issueStatusIndex];
  console.log(issueStatus);
  const closureEmailDate = row[closureEmailDateIndex];
  const finalClosureDate = row[finalClosureDateIndex];
  const ticketNumber = row[ticketNoIndex];
  const recipientEmail = row[recipientEmailIndex];
  const department = row[departmentIndex];

  if (issueStatus === "6.1 Final closure" && closureEmailDate !="" && finalClosureDate ===""){
    const emailVariables = { ticketNumber };

    // Update "Issue Status"
    sheet.getRange(i + 2, issueStatusIndex + 1).setValue("6.2 Final closure [A]");

    // Set "Final Closure Date" to today's date
    sheet.getRange(i + 2, finalClosureDateIndex + 1).setValue(formattedDate).setNumberFormat("dd-MMM-yyyy");

    // Send the email
    const ccListEmail = [ccEmailLists[department], SUPPORT_CC];
    const emailSubject = `Your Ticket No. ${ticketNumber} has been resolved.`;
    const emailBody = createEmailBody("Email template 6.html", emailVariables);
    sendEmail(recipientEmail, emailSubject, emailBody, ccList = ccListEmail);
  }
}

