const EMAIL_SENDER = "surekha.more@upthink.com"

function sendReminderEmail() {
  const emailar = new RemainderEmail("Audit_Requests", EMAIL_SENDER);
  emailar.processReminders();
}


class RemainderEmail {
  constructor(sheetName, email) {
    this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    this.email = email;
    this.data = this.sheet.getDataRange().getValues();
    this.today = new Date();
    this.today.setHours(0, 0, 0, 0);

    const headers = this.data[0];
    this.REMINDER_DATE_COL = headers.indexOf("Reminder Date") + 1;
    this.R1_COL = headers.indexOf("R1") + 1;
    this.R2_COL = headers.indexOf("R2") + 1;
    this.STATUS_COL = headers.indexOf("Followup Status") + 1;
    this.EMAIL_COL = headers.indexOf("Candidate's Email ID") + 1;

  }


  processReminders() {
    for (let i = 1; i < this.data.length; i++) {
      const row = this.data[i];
      const reminderDate = row[this.REMINDER_DATE_COL - 1];
      const r1 = row[this.R1_COL - 1];
      const r2 = row[this.R2_COL - 1];
      const status = row[this.STATUS_COL - 1];
      const email = row[this.EMAIL_COL - 1];

      if (!reminderDate || status === "Closed") continue;

      const daysSinceReminder = Math.floor((this.today - new Date(reminderDate).setHours(0, 0, 0, 0)) / (1000 * 60 * 60 * 24));
      // const originalEmail = row[this.EMAIL_COL - 1];

      // if (!r1 && daysSinceReminder >= 0) {
      //   this.sendEmail(originalEmail, 1, i + 1);
      //   this.sheet.getRange(i + 1, this.R1_COL).setValue("Y");
      // } else if (r1 && !r2 && daysSinceReminder >= 3) {
      //   this.sendEmail(originalEmail, 2, i + 1);
      //   this.sheet.getRange(i + 1, this.R2_COL).setValue("Y");
      // } else if (r1 && r2 && daysSinceReminder >= 6 && (daysSinceReminder - 6) % 3 === 0) {
      //   this.sendEmail(originalEmail, 3, i + 1);
      // }


      if (!r1 && daysSinceReminder >= 0) {
        this.sendEmail(email, 1, i + 1);
        this.sheet.getRange(i + 1, this.R1_COL).setValue("Y");
      }
      else if (r1 && !r2 && daysSinceReminder >= 3) {
        this.sendEmail(email, 2, i + 1);
        this.sheet.getRange(i + 1, this.R2_COL).setValue("Y");
      }
      else if (r1 && r2 && daysSinceReminder >= 6 && (daysSinceReminder - 6) % 3 === 0) {
        this.sendEmail(email, 3, i + 1);
      }
    }
  }


  sendEmail(originalEmail, round, row) {
    const subject = `Remainder Email (Round ${round})`;
    const body = `This is Remainder Round ${round}.\nPlease take the required followup actions. \n\n(Originally meant for:${originalEmail})\n(Row:${row})`;

    if (originalEmail) {
      GmailApp.sendEmail(this.email, subject, body);
      Logger.log(`Test Email (Round ${round}) sent to ${this.email} for candidate ${originalEmail}`);
    }
  }
  
}




