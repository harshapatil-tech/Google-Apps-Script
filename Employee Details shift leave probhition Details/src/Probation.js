// GLOBAL VARIABLES
const INPUT_SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();

function sendReminderEmails() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const sheet = INPUT_SPREADSHEET.getSheetByName("Employee Info");
  const inputDataRange = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  let inputHeaders = inputDataRange[0], inputData = inputDataRange.slice(1);

  inputData = inputData.filter(row => row[inputHeaders.indexOf("Status")] === "Active");

  inputData = inputData.filter(row => !row[inputHeaders.indexOf("Designation")].toLowerCase().includes("(apprentice)"));

  // const exemptedEmployees = new Set(["Ananya Sharma", "Harsha Patil", "Prathin Jaggili", "Pusparanjita Mahanta", 
  //                                   "Pravin Shinde", "Ratnesh", "Kevina Rebecah", "Likhitha Gopalam", "T Pal Raj", "Isha", "Nikita Badgujar", 
  //                                   "Shikha Dubey", "Ishita Bhowmick", "Nikita Deshmukh", "Ashi Arjeria", "Archana Das"]);
  // inputData = inputData.filter(row => !exemptedEmployees.has(row[inputHeaders.indexOf("Employee Name")]));

  const inputIndices = {
    empNameIdx: inputHeaders.indexOf("Employee Name"),
    reportingManagerIdx: inputHeaders.indexOf("Reporting Manager"),
    pyramidCategoryIdx: inputHeaders.indexOf("Grade"),
    designationIdx: inputHeaders.indexOf("Designation"),
    dojIdx: inputHeaders.indexOf("DOJ"),
    emailIdx: inputHeaders.indexOf("Official Email ID"),
    functionIdx : inputHeaders.indexOf("Function")
  };

  const reportingManagersMap = {}, emailMap = {};
  inputData.forEach(r => {
    const employee = r[inputIndices.empNameIdx], manager = r[inputIndices.reportingManagerIdx], email = r[inputIndices.emailIdx];
    emailMap[employee] = email;
    if (!reportingManagersMap.hasOwnProperty(employee) && manager !== 'Apurva Yadav' && manager !== 'Amogh Chaphalkar' && manager !== 'Pranav Deshpande') {
      reportingManagersMap[employee] = manager;
    }
  });

  // Object.entries(reportingManagersMap).map(entry => console.log(entry[0], entry[1]));
  // inputData.forEach(each => Logger.log(each));


  inputData.forEach(r => {
      // Convert the joining date to IST timezone
      // today.setHours(0, 0, 0, 0);
      // const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM-yy");
      const joiningDate = new Date(Utilities.formatDate(r[inputIndices.dojIdx], Session.getScriptTimeZone(), "dd-MMM-yy"));
      joiningDate.setHours(0, 0, 0, 0)
      // Calculate 84 days after joining date
      // let eightyFourDaysAfter = new Date(joiningDate);
      // eightyFourDaysAfter.setDate(joiningDate.getDate() + 84);
      // eightyFourDaysAfter = Utilities.formatDate(eightyFourDaysAfter, Session.getScriptTimeZone(), "dd-MMM-yy");

      // // Calculate 89 days after joining date
      // let eightyNineDaysAfter = new Date(joiningDate);
      // eightyNineDaysAfter.setDate(joiningDate.getDate() + 89);
      // eightyNineDaysAfter = Utilities.formatDate(eightyNineDaysAfter, Session.getScriptTimeZone(), "dd-MMM-yy");

      // // Calculate 90 days after joining date
      let ninetyDaysAfter = new Date(joiningDate);
      ninetyDaysAfter.setDate(joiningDate.getDate() + 90);
      ninetyDaysAfter = Utilities.formatDate(ninetyDaysAfter, Session.getScriptTimeZone(), "EEEE, dd-MMM-yy");

      const smeName = r[inputIndices.empNameIdx];
      const reportingManager = r[inputIndices.reportingManagerIdx];
      const designation = r[inputIndices.designationIdx];
      let subjectLine;
      let body;
      
      if (CentralLibrary.getDaysDifference(today, joiningDate) == 89
        && CentralLibrary.getDaysDifference(today, joiningDate) 
        && r[inputIndices.pyramidCategoryIdx] === 'U1' 
        && r[inputIndices.functionIdx] === 'Technical') 
      {
        console.log(smeName)
        subjectLine = `Request confirmation to end the probation period for ${smeName}`;
        body = HtmlService.createHtmlOutputFromFile('89days.html').getContent();    
        body = body.replace("{smeName}", smeName)
                  .replace("{reportingManager}", reportingManager)
                  .replace("{designation}", designation)
                  .replace("{ninetyDays}", ninetyDaysAfter);
        
        sendEmailOnJoiningAnniversary(r, inputIndices, reportingManagersMap, emailMap, subjectLine, body);

      } else if (CentralLibrary.getDaysDifference(today, joiningDate) == 84 
        && r[inputIndices.pyramidCategoryIdx] === 'U1' 
        && r[inputIndices.functionIdx] === 'Technical') 
      {
        console.log(CentralLibrary.getDaysDifference(today, joiningDate))
        subjectLine = `Request confirmation to end the probation period for ${smeName}`;
        body = HtmlService.createHtmlOutputFromFile('84days.html').getContent();
        body = body.replace("{smeName}", smeName)
                  .replace("{reportingManager}", reportingManager)
                  .replace("{designation}", designation)
                  .replace("{ninetyDays}", ninetyDaysAfter);

        sendEmailOnJoiningAnniversary(r, inputIndices, reportingManagersMap, emailMap, subjectLine, body);

      }
    });
}


function sendEmailOnJoiningAnniversary(employeeData, inputIndices, reportingManagersMap, emailMap, subject, body) {
  const candidateName = employeeData[inputIndices.empNameIdx];
  const managerList = getReportingManagers(candidateName, reportingManagersMap);
  const managerEmails = managerList.map(manager => emailMap[manager]);
  const receiverEmail = managerEmails[0];
  const ccEmails = [...managerEmails.slice(1), "tanaya.adulkar@upthink.com", "deepti.tonpe@upthink.com"];
  
  const options = {
    to: receiverEmail, // Receiver's email address
    cc: ccEmails.join(','), // Comma-separated list of CC email addresses
    subject: subject, // Email subject
    htmlBody: body, // Email body
    from: "deepti.tonpe@upthink.com",
  };

  // Uncomment the following line to send the email
  GmailApp.sendEmail(receiverEmail, subject, body, options);
}

// Define the getReportingManagers function here
function getReportingManagers(employee, reportingManagerMap) {
  const reportingManagers = [];
  if (employee in reportingManagerMap) {
    const directManager = reportingManagerMap[employee];
    if (directManager !== "") {
      reportingManagers.push(directManager);
      // Recursively find reporting managers of the direct manager
      const indirectManagers = getReportingManagers(directManager, reportingManagerMap);

      // Filter out empty strings from indirectManagers before pushing them
      const nonEmptyIndirectManagers = indirectManagers.filter(manager => manager !== "");
      reportingManagers.push(...nonEmptyIndirectManagers);
    }
  }
  return reportingManagers;
}




function applyCustomFormatting(range, options) {

  options = options || {};
  
  var fontSize = options.fontSize || 10;
  var fontColor = options.fontColor || 'black';
  var bgColor = options.bgColor || 'white';
  var fontWeight = options.fontWeight || 'normal'

  range.setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true)
      .setFontFamily("Roboto")
      .setFontSize(fontSize)
      .setFontColor(fontColor)
      .setFontWeight(fontWeight)
      .setBorder(true, true, true, true, true, true)
      .setBackground(bgColor);
  return range;
};