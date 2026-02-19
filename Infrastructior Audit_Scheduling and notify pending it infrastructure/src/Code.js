function sendEmailWithSheetData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Audit_Requests");
  
  // Use get_Data_Indices_From_Sheet to get headers and data
  const [headers, data] = get_Data_Indices_From_Sheet(sheet);
  
  // Define required field names
  const requiredFields = [
    "Audit Date", "Device", "Processor", "RAM (GB)", "Operating System", 
    "Device Power Backup (In Min)", "Antivirus Installed", "Power Backup", "Power Backup Time (In Min)",
    "Mode of Internet", "Internet Connectivity Speed (Download & Upload in Mbps)", 
    "Alternate Arrangement of Network", "Alternate Arrangement of System", 
    "Recommendation", "Audit Status"
  ];
  
  // Iterate through each row of data
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    
    const triggerValue = row[headers["WFH Infrastructure Audit Request Trigger"]];
    const auditCreatedValue = row[headers["WFH Infrastructure Audit Request Created?"]];
    
    // Process only rows where the trigger is 'True' and audit not yet created
    if (triggerValue === true && auditCreatedValue !== 'Y') {
      // Check if all required fields are populated
      const missingFields = requiredFields.some(field => !row[headers[field]]);

      if (missingFields) {
        Logger.log('Skipping row ' + (i + 1) + ' due to missing required fields.');
        continue;
      }
      
      // Extract data from the row
      const emailData = {
        employeeName: row[headers["Candidate Name"]],
        auditDate : Utilities.formatDate(row[headers["Audit Date"]], Session.getScriptTimeZone(), "dd-MMM-yyyy"),
        meetLink: row[headers["Google Meet Link"]],
        hardwareSpecificationLaptopModelNoMake: row[headers["Device"]],
        processor: row[headers["Processor"]],
        ram: row[headers["RAM (GB)"]],
        softwareSpecification: row[headers["Operating System"]],
        powerBackupforLaptop: row[headers["Device Power Backup (In Min)"]],
        antivirusInstalledornot: row[headers["Antivirus Installed"]],
        powerBackupInverterorUPSStatus: row[headers["Power Backup"]],
        powerBackupTime: row[headers["Power Backup Time (In Min)"]],
        modeofInternet: row[headers["Mode of Internet"]],
        internetConnectivitySpeed: row[headers["Internet Connectivity Speed (Download & Upload in Mbps)"]],
        alternetArrangementincaseofnetworkFailuare: row[headers["Alternate Arrangement of Network"]],
        alternetArrangementincaseofsystemFailuare: row[headers["Alternate Arrangement of System"]],
        recommendation: row[headers["Recommendation"]],
        wFHAuditStatus: row[headers["Audit Status"]],
        reAudit: row[headers["WFH Infrastructure Audit Request Trigger"]],
        remark: row[headers["Remark"]],
        additionalComment : row[headers["Additional Comment (If Any)"]]
      };

      // Generate HTML body using template
      const template = HtmlService.createTemplateFromFile('auditRequest.html');
      Object.assign(template, emailData);
      const body = template.evaluate().getContent();

      let toList = [row[headers["Email Address"]]];
      let ccList = [];
      if(row[headers["Email Address"]] == "careers@upthink.com") {
        toList = toList.concat(["deepti.tonpe@upthink.com", "surekha.more@upthink.com", "tanaya.adulkar@upthink.com"]);
        ccList = ccList.concat(["sachin.tribhuwan@upthink.com", "tejas.jagtap@upthink.com"]);
      } else {
        ccList = ccList.concat(["tejas.jagtap@upthink.com"])
      }

      const subjectLine = `Work from Home Infrastructure Audit: ${row[headers["Candidate Name"]]} | ${row[headers["Department of the Auditee"]]}`;
      // 
      // sendEmail(toList, ccList, "Work from Home Infrastructure Audit", body);
      sendEmail(toList, ccList, subjectLine, body);

      // Mark the row as processed by updating the "WFH Infrastructure Audit Request Created?" column
      sheet.getRange(i + 2, headers["WFH Infrastructure Audit Request Created?"] + 1).setValue('Y');
      sheet.getRange(i + 2, headers["WFH Infrastructure Audit Request Trigger"] + 1).setValue(false);
      Logger.log('Audit request created for row ' + (i + 2));
    }
  }
}

function sendEmail(toList, ccList, subject, body) {
  try {
    const to = Array.isArray(toList) ? toList.join(',') : '';
    const cc = Array.isArray(ccList) ? ccList.join(',') : '';

    MailApp.sendEmail({
      to: to,
      cc: cc,
      subject: subject,
      htmlBody: body
    });

    Logger.log(`Email sent successfully to: ${to}`);
  } catch (error) {
    Logger.log(`Failed to send email: ${error.message}`);
  }
}

function createIndexMap(headers) {
  return headers.reduce((map, val, index) => {
    if (val !== "")
      map[val.trim()] = index;
    return map;
  }, {});
}

function get_Data_Indices_From_Sheet(sheet, startRow = 0) {
  const dataRange = sheet.getDataRange().getValues().slice(startRow);
  const headers = dataRange[0], data = dataRange.slice(1);
  return [createIndexMap(headers), data];
}

function onOpen() {
  // Add a custom menu
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('MENU')
      .addItem('Send Email', 'sendEmailWithSheetData')
      .addToUi();
}











 