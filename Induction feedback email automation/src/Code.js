function onOpen(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("MENU")
  .addItem("Send Email", "sendEmails")
  .addToUi();
}


function sendEmails() {

  const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const inputSheet = sourceSpreadsheet.getSheetByName("Input Sheet");
  const formResponses1 = sourceSpreadsheet.getSheetByName("Form Responses 1");
  const formResponses2 = sourceSpreadsheet.getSheetByName("Form Responses 2");
  const formResponses3 = sourceSpreadsheet.getSheetByName("Form Responses 3");

  const formResponses1DataRange = formResponses1.getDataRange().getValues();
  const formResponses1HeaderIndices = createIndexMap(formResponses1DataRange[0]);
  const formResponses1AllEmails = formResponses1DataRange.slice(1)
                                                        .map(r => r[formResponses1HeaderIndices["Email Address"]].trim().toLowerCase())
                                                        .filter(r=>r!=="");

  const formResponses2DataRange = formResponses2.getDataRange().getValues();
  const formResponses2HeaderIndices = createIndexMap(formResponses2DataRange[0]);
  const formResponses2AllEmails = formResponses2DataRange.slice(1)
                                                        .map(r => r[formResponses2HeaderIndices["Email Address"]].trim().toLowerCase())
                                                        .filter(r=>r!=="");

  const formResponses3DataRange = formResponses3.getDataRange().getValues();
  const formResponses3HeaderIndices = createIndexMap(formResponses3DataRange[0]);
  const formResponses3AllEmails = formResponses3DataRange.slice(1)
                                                        .map(r => r[formResponses3HeaderIndices["Email Address"]].trim().toLowerCase())
                                                        .filter(r=>r!=="");


  const dataRange = inputSheet.getDataRange().getValues();
  const header = dataRange[0], data = dataRange.slice(1);
  const headerIndices = createIndexMap(header);


  const today = new Date();  // Date Today


  data.forEach( (row, index) => {

    const email = row[headerIndices["Official Email ID"]].trim().toLowerCase();

    if (formResponses1AllEmails.includes(email)) {

      inputSheet.getRange(index+2, headerIndices["Form 1 Emailed"]+1).setValue(true);
      inputSheet.getRange(index+2, headerIndices["Feedback 1 Received"]+1).setValue(true);

    }else if (formResponses2AllEmails.includes(email)) {

      inputSheet.getRange(index+2, headerIndices["Form 2 Emailed"]+1).setValue(true);
      inputSheet.getRange(index+2, headerIndices["Feedback 2 Received"]+1).setValue(true);

    }else if (formResponses3AllEmails.includes(email)) {

      inputSheet.getRange(index+2, headerIndices["Form 3 Emailed"]+1).setValue(true);
      inputSheet.getRange(index+2, headerIndices["Feedback 3 Received"]+1).setValue(true);
    }

  });


  data.forEach( (row, index) => {

    const name = row[headerIndices["Employee Name"]];
    const email = row[headerIndices["Official Email ID"]];
    //console.log(CentralLibrary.getDaysDifference(today, row[headerIndices["DOJ"]]))
    const dayDifference = CentralLibrary.getDaysDifference(today, row[headerIndices["DOJ"]]);

    const form1Response = row[headerIndices["Feedback 1 Received"]];
    const form2Response = row[headerIndices["Feedback 2 Received"]];
    const form3Response = row[headerIndices["Feedback 3 Received"]];

    const form1Emailed = row[headerIndices["Form 1 Emailed"]];
    const form2Emailed = row[headerIndices["Form 2 Emailed"]];
    const form3Emailed = row[headerIndices["Form 3 Emailed"]];

    const exited = row[headerIndices["Exited Employee"]];

    if (exited !== true) {

      if (dayDifference >= 30 && dayDifference <= 60 && form1Emailed === false && form1Response === false) {

          emailTemplate(email, name, "30days.html")
          inputSheet.getRange(index+2, headerIndices["Form 1 Emailed"]+1).setValue(true);
      
      }
      else if (dayDifference >= 60 && dayDifference <= 90 && form2Emailed === false && form2Response === false) {

          emailTemplate(email, name, "60days.html")
          inputSheet.getRange(index+2, headerIndices["Form 2 Emailed"]+1).setValue(true);

      }
      else if (dayDifference >= 90 && form3Emailed === false && form3Response === false) {

          emailTemplate(email, name, "90days.html")
          inputSheet.getRange(index+2, headerIndices["Form 3 Emailed"]+1).setValue(true);

      }
    } // End of checking exited or not.
  });


}


function emailTemplate(emailId, name, htmlFile) {

  const template = HtmlService.createTemplateFromFile(htmlFile);
  template.name = name;
  const messageBody = template.evaluate().getContent();
  const recipientEmail = emailId.toString();
  //console.log(recipientEmail);
  const options = {
    htmlBody: messageBody,
    from:"surekha.more@upthink.com",//"shilpa.kulkarni@upthink.com",
  };

  GmailApp.sendEmail(recipientEmail, "Request feedback for the orientation program.", "This is a fallback plain text body", options);
}




function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent()
}



function createIndexMap(headers) {

  return headers.reduce( (map, value, index) => {
    map[value] = index;
    return map;
  }, {});
}


// function getDaysDifference(targetDate, today) {
//   // Normalize the dates to midnight
//   targetDate.setHours(0, 0, 0, 0);
//   today.setHours(0, 0, 0, 0);

//   var differenceInMilliseconds = targetDate.getTime() - today.getTime();
//   var differenceInDays = Math.floor(differenceInMilliseconds / (1000 * 60 * 60 * 24));
//   return differenceInDays;
// }