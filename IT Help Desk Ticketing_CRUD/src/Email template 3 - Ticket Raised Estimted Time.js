function sendAutomaticEmails(){
  const spreadsheet = SpreadsheetApp.openById("1iByitSy5R35cu13rupuppctpzV0X8dTzqDSJJP2ilAk");
  const sheet = spreadsheet.getSheetByName("Master DB");

  const[dbHeaders, dbData] = CentralLibrary.get_Data_Indices_From_Sheet(sheet);
  const row = dbData[0];

  const today = new Date();
  const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd-MMM-yyyy");

  dbData.forEach((row, index)=>{
    emailTemplate3(sheet, dbHeaders, row, index);
    updateFinalClosureDates(sheet, dbHeaders, row, index, formattedDate, today);
  })
  // console.log(row);
  
  
}

function emailTemplate3(sheet, headers, row, i){

  // Column indices
  const ticketNoIndex = headers["Ticket No."];
  const recipientEmailIndex = headers["Email Address"];
  const employeeNameIndex = headers["Employee Name"];
  const itSupportDiagnosisIndex = headers["IT Support diagnosis"];
  const estimatedTimeIndex = headers["Estimated Time"];
  const issueStatusIndex = headers["Issue Status"];
  const initialEmailDateIndex = headers["Initial Email Date"];
  const closureEmailDateIndex = headers["Closure Email Date"];
  const raiseTicketDateIndex = headers["Raise Ticket Date"];
  const departmentIndex = headers["Department"];

  const today = new Date();
  const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd-MMM-yyyy");

  const itSupportDiagnosis = row[itSupportDiagnosisIndex];
  const estimatedTime = row[estimatedTimeIndex];
  const initialEmailDate = row[initialEmailDateIndex];
  const closureEmailDate = row[closureEmailDateIndex];
  const raiseTicketDate = row[raiseTicketDateIndex];
  const issueStatus = row[issueStatusIndex];
  const ticketNumber = row[ticketNoIndex];
  const recipientEmail = row[recipientEmailIndex];
  const employeeName = row[employeeNameIndex];
  const department = row[departmentIndex];

  // Check conditions
    if (itSupportDiagnosis !="" && estimatedTime !="" && initialEmailDate !="" && closureEmailDate ==="" && raiseTicketDate ===""){
      
      // Prepare email variables
      const emailVariables = {
        ticketNumber,
        emailAddress: recipientEmail,
        employeeName,
        itSupportDiagnosis,
        estimatedTime,
      };

      // Set "Raise Ticket Date" to today's date
      sheet.getRange(i + 2, raiseTicketDateIndex + 1).setValue(formattedDate).setNumberFormat("dd-MMM-yyyy");

      // Update Issue Status
      const updatedStatus = issueStatus === "3.2 Silicon Rental Informed [A]"
        ? "3.3 Ticket Raised [Silicon] [A]"
        : "2.1 Ticket Raised [A]";
      sheet.getRange(i + 2, issueStatusIndex + 1).setValue(updatedStatus);

      // Email content
      const ccListEmail = [ccEmailLists[department], SUPPORT_CC];
      const emailSubject = `Ticket No. ${ticketNumber} Estimated time is ${estimatedTime}`;
      const emailBody = createEmailBody("Email template 3.html", emailVariables);

      // Send email
      sendEmail(recipientEmail, emailSubject, emailBody, ccList=ccListEmail);

    }
}