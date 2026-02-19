function onUpdateButtonClick() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recruiterSheet = ss.getSheetByName("Recruiter");
  const auditSheet = ss.getSheetByName("Audit_Requests");
  const recruiterHeaders = recruiterSheet.getRange(4, 1, 1, recruiterSheet.getLastColumn()).getValues()[0];
  const recruiterData = recruiterSheet.getRange(5, 1, recruiterSheet.getLastRow() - 1, recruiterSheet.getLastColumn()).getValues();
  const auditHeaders = auditSheet.getRange(1, 1, 1, auditSheet.getLastColumn()).getValues()[0];
  const auditData = auditSheet.getRange(2, 1, auditSheet.getLastRow() - 1, auditSheet.getLastColumn()).getValues();

  for (let i = 0; i < recruiterData.length; i++) {
    const entry = new RecruitersEntry(recruiterData[i], recruiterHeaders);
    const email = entry.data["Candidate's Email ID"];
    const auditIndex = auditData.findLastIndex(r => r[auditHeaders.indexOf("Candidate's Email ID")] === email); //find index in Audit_Request where the email match


    if (auditIndex !== -1 && entry.shouldUpdate()) {

      //if candidate found and update checkbox is true
      let updatedRow = entry.updateAuditRow(auditHeaders, auditData[auditIndex]);
      console.log(updatedRow)
      if (entry.shouldSendEmail()) {
        updatedRow = entry.sendIfEmailNeeded(auditHeaders, updatedRow);
        //Logger.log(entry.shouldSendEmail());
      }

      const dojIdx = auditHeaders.indexOf("DOJ");
      const followupIdx = auditHeaders.indexOf("Followup Status");
      const daysBeforeRemIdx = auditHeaders.indexOf("Days before reminder");
      const reminderDateIdx = auditHeaders.indexOf("Reminder Date");
      const f1Idx = auditHeaders.indexOf("F1");

      console.log(updatedRow[f1Idx])

      auditSheet.getRange(auditIndex + 2, dojIdx + 1).setValue(updatedRow[dojIdx]).setNumberFormat("dd-MMM-yyyy");
      auditSheet.getRange(auditIndex + 2, followupIdx + 1).setValue(updatedRow[followupIdx]);
      auditSheet.getRange(auditIndex + 2, daysBeforeRemIdx + 1).setValue(updatedRow[daysBeforeRemIdx]);
      auditSheet.getRange(auditIndex + 2, reminderDateIdx + 1).setValue(updatedRow[reminderDateIdx]).setNumberFormat("dd-MMM-yyyy");
      auditSheet.getRange(auditIndex + 2, f1Idx + 1).setValue(updatedRow[f1Idx]);
      //auditSheet.getRange(auditIndex + 2, 1, 1, updatedRow.length).setValues([updatedRow]);//don't
    }
  }
  SpreadsheetApp.getUi().alert("Data Suceessfully updated");
}


class RecruitersEntry {
  constructor(row, headers) {
    this.headers = headers;
    this.row = row;
    this.data = this.parseData();
  }

  parseData() {
    const data = {};
    this.headers.forEach((header, i) => {
      data[header] = this.row[i];
    });
    return data;
  }


  isTrue(field) {
    //return String(this.data[field]).toLowerCase() === 'true';
    const value = this.data[field];
    return value === true || String(value).toLowerCase() === 'true';
  }


  calculateReminderDate() {
    const dojRaw = this.data["DOJ"];
    const doj = dojRaw instanceof Date ? dojRaw : Utilities.parseDate(dojRaw, Session.getScriptTimeZone(), "dd-MMM-yyyy");
    const daysBefore = parseInt(this.data["Days before reminder"], 10);

    if (!isNaN(doj.getTime()) && !isNaN(daysBefore)) {
      const reminderDate = new Date(doj);
      reminderDate.setDate(reminderDate.getDate() + daysBefore);
      return Utilities.formatDate(reminderDate, Session.getScriptTimeZone(), "dd-MMM-yyyy");
    }
    return "";
  }


  shouldUpdate() {
    return this.isTrue("Update");
  }


  shouldSendEmail() {
    return this.isTrue("Email Candidate");
  }


  updateAuditRow(auditHeaders, auditRow) {
    const updated = auditRow.slice();

    updated[auditHeaders.indexOf("DOJ")] = this.data["DOJ"];
    updated[auditHeaders.indexOf("Days before reminder")] = this.data["Days before reminder"];
    updated[auditHeaders.indexOf("Reminder Date")] = this.calculateReminderDate();
    updated[auditHeaders.indexOf("Followup Status")] = this.data["Followup Status"];
    updated[auditHeaders.indexOf("Comments")] = this.data["Comments"];
    return updated;
  }


  sendIfEmailNeeded(auditHeaders, auditRow) {
    const status = (this.data["Followup Status"] || "");
    const email = this.data["Candidate's Email ID"];
    const getIndex = (label) => auditHeaders.indexOf(label);


    if (status === "F1 Done" && !auditRow[getIndex("F1")]) {
      this.sendEmail(email, "Template2");
      auditRow[getIndex("F1")] = "Y";
    }
    else if (status === "F2 Done" && auditRow[getIndex("F1")] && !auditRow[getIndex("F2")]) {
      this.sendEmail(email, "Template2");
      auditRow[getIndex("F2")] = "Y";
    }
    else if (status === "Not Connected" && !auditRow[getIndex("NC")]) {
      this.sendEmail(email, "Template3");
      auditRow[getIndex("NC")] = "Y";
    }
    else if (status === "postponed" && !auditRow[getIndex("P")]) {
      this.sendEmail(email, "Template4");
      auditRow[getIndex("P")] = "Y";
    }
    return auditRow;
  }
  // sendEmail(email, template) {
  //   if (!email) return;
  //   email = "harsha.patil@upthink.com";
  //   GmailApp.sendEmail(email, "Audit Followup", `Please refer to this:${template}`);
  //   //Logger.log(`Sent ${template} email to ${email}`);
  // }

  sendEmail(email, templateName) {
    if (!email) return;
    const template = HtmlService.createTemplateFromFile(templateName);
    template.Candidate_Name = this.data["Candidate Name"];
    const htmlBody = template.evaluate().getContent();
    email = "surekha.more@upthink.com";
    GmailApp.sendEmail(email, "Audit Followup", "Please view this email in HTML format.", {
      htmlBody: htmlBody
    });
  }


}