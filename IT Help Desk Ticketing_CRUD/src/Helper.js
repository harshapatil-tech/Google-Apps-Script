function createEmailBody(templateName, variables) {
  // Create an HTML template from the specified file
  const emailTemplate = HtmlService.createTemplateFromFile(templateName);
  // Dynamically assign variables from the object
  for (const [key, value] of Object.entries(variables)) {
    emailTemplate[key] = value;
  }
  // Evaluate the template and return the generated HTML content
  return emailTemplate.evaluate().getContent();
}


function sendEmail(recipientEmail, emailSubject, emailBody, ccList = [], bccList = [], replyTo = SENDER_EMAIL_ID) {
  try {
    // Prepare email options
    const options = {
      htmlBody: emailBody,               // HTML content for the email
      cc: Array.isArray(ccList) ? ccList.join(",") : ccList,  // Convert CC list to a comma-separated string
      bcc: Array.isArray(bccList) ? bccList.join(",") : bccList,  // Convert BCC list to a comma-separated string
      replyTo: replyTo,                  // Reply-To address
      from: SENDER_EMAIL_ID,    // Sender email address
      name: "UpThink Support" // Display name for the sender
    };

    // Send the email using GmailApp
    GmailApp.sendEmail(
      recipientEmail,           // Recipient email address
      emailSubject,             // Subject of the email
      "This is a fallback plain text body.", // Fallback plain-text body
      options                   // Options object with additional fields
    );
    // Log success message
    console.log(`Email sent successfully to ${recipientEmail}`);
  } catch (error) {
    // Log the error in case of failure
    console.error(`Failed to send email to ${recipientEmail}: ${error.message}`);
  }
}







// class Helper {

//   static generateTicketNo () {

//   }
// }

