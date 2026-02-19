function emailTemplate2(sheet, headers, row, i){
  const ticketNoIndex = headers["Ticket No."];
  const recipientEmailIndex = headers["Email Address"];
  const laptopNumberIndex = headers["Laptop Number"];
  const employeeNameIndex = headers["Employee Name"];
  const contactNumberIndex = headers["Contact Number"];
  const issueTypeIndex = headers["Issue Type"];
  const hardwareIssueIndex = headers["Hardware Issue"];
  const softwareIssueIndex = headers["Software Issue"];
  const issueDescriptionIndex =  headers["Issue Description"];
  const issueStatusIndex = headers["Issue Status"];
  const initialEmailDateIndex = headers["Initial Email Date"];
  const siliconRentalEmailDateIndex = headers["Silicon Rental Email Date"];
  const closureEmailDateIndex = headers["Closure Email Date"];
  const raiseTicketDateIndex = headers["Raise Ticket Date"];
  const departmentIndex = headers["Department"];

  const today = new Date();
  const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd-MMM-yyyy");

  const issueStatus = row[issueStatusIndex];
  const initialEmailDate = row[initialEmailDateIndex];
  const siliconRentalEmailDate = row[siliconRentalEmailDateIndex];
  const closureEmailDate = row[closureEmailDateIndex];
  const raiseTicketDate = row[raiseTicketDateIndex];
  const ticketNumber = row[ticketNoIndex];
  const recipientEmail = row[recipientEmailIndex];
  const department = row[departmentIndex];

  if(issueStatus === "3.1 Inform Silicon Rental" && initialEmailDate !=="" && siliconRentalEmailDate ==="" && closureEmailDate ===""){
     // Prepare email variables
      const emailVariables = {
        ticketNumber : row[ticketNoIndex],
        emailAddress : row[recipientEmailIndex],
        employeeName : row[employeeNameIndex],
        laptopNumber: row[laptopNumberIndex],
        contactNumber : row[contactNumberIndex],
        issueType : row[issueTypeIndex],
        // hardwareIssue : row[hardwareIssueIndex],
        // softwareIssue : row[softwareIssueIndex],
        hsIssue: row[hardwareIssueIndex] ? row[hardwareIssueIndex] : row[softwareIssueIndex],
        issueDescription : row[issueDescriptionIndex]
      };

      // Set "Silicon Rental Email Date" to today's date
      sheet.getRange(i + 2, siliconRentalEmailDateIndex + 1).setValue(formattedDate).setNumberFormat("dd-MMM-yyyy");

      // Update Issue Status
      const updateStatus = raiseTicketDate ==="" ? "3.2 Silicon Rental Informed [A]": "3.3 Ticket Raised [Silicon] [A]"
      sheet.getRange(i + 2, issueStatusIndex + 1).setValue(updateStatus);
      // const ccListEmail = [recipientEmail, "tejas.jagtap@upthink.com"];
      // Email content
      const ccListEmail = [ccEmailLists[department], SUPPORT_CC, "tejas.jagtap@upthink.com"];
      const emailSubject = `Ticket No. ${ticketNumber} Silicon Rental Solutions`;
      const emailBody = createEmailBody("Email template 2.html", emailVariables);

      // Send email
      sendEmail(siliconRentalTeam["Silicon Rental"].join(", "), emailSubject, emailBody, ccList=ccListEmail);
  }

}