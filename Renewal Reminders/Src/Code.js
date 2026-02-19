function renewalReminders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = sheet.getSheetByName("Renewal_Reminders");
  const templateText = sheet.getSheetByName("Template_Text").getRange(1, 1).getValue();
  const logSheet = sheet.getSheetByName("Logs");
  const reminderSheet = sheet.getSheetByName("ReminderDates")

  // const lastIndex = getLastRowWithEntry(inputSheet);
  const inputSheetDataRange = inputSheet.getRange(1, 1, inputSheet.getLastRow(), inputSheet.getLastColumn()).getValues();
  const inputSheetHeaders = inputSheetDataRange[0]; data = inputSheetDataRange.slice(1);
  const inputSheetIndices = createIndexMap(inputSheetHeaders);

  // Setting today's and tomorrow's date
  const today = new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);
  

  const reminderHeader = reminderSheet.getRange(1, 1, 1, 2).getValues().flat();
  const remDatesIdx = reminderHeader.indexOf("Dates") + 1;
  

  data.forEach((row, index) => {

    if (
        row[inputSheetIndices['Particulars']] != "" 
        && row[inputSheetIndices['Due Date']] != "" 
        && row[inputSheetIndices["Mail (TO)"]] !== ""
    ) {
      
      const numDaysRem = getDaysDifference(row[inputSheetIndices['Due Date']], today);
      const reminders = [
        { idx: inputSheetIndices["Reminder 1"], days: reminderSheet.getRange(2, remDatesIdx).getValue() },
        { idx: inputSheetIndices["Reminder 2"], days: reminderSheet.getRange(3, remDatesIdx).getValue() },
        { idx: inputSheetIndices["Reminder 3"], days: reminderSheet.getRange(4, remDatesIdx).getValue() },
      ];

      if (getDaysDifference(today, row[inputSheetIndices['Due Date']]) === 1) {

        // const newDate = new Date();
        const renewalFrequency = row[inputSheetIndices["Frequency in Months"]]
        const newRenewalDate = addMonthsToDate(row[inputSheetIndices['Due Date']], renewalFrequency)
        inputSheet.getRange(index+2, inputSheetIndices["Due Date"]+1).setValue(newRenewalDate);
        // inputSheet.getRange(index+2, inputSheetIndices["Reminder 1"]+1).setValue(false);
        // inputSheet.getRange(index+2, inputSheetIndices["Reminder 2"]+1).setValue(false);
        // inputSheet.getRange(index+2, inputSheetIndices["Reminder 3"]+1).setValue(false);
      }

      else {

        reminders.forEach(reminder => {
          
          if (row[reminder.idx] && numDaysRem === reminder.days) {
            const currentParticulars = row[inputSheetIndices["Particulars"]];
            const currentExpiryDate = Utilities.formatDate(row[inputSheetIndices['Due Date']], Session.getScriptTimeZone(), "EEEE dd MMMM yyyy");
            
            const messageBody = templateText
              .replace("{owner}", row[inputSheetIndices["Owner"]])
              .replace("{particulars}", currentParticulars)
              .replace("{expiry_date}", currentExpiryDate)
              .replace("{num_of_days}", numDaysRem);

            const sender = "automation+renewal_reminders@upthink.com";
            const options = {
              from: sender,
              cc: row[inputSheetIndices["Mail (CC)"]],
            };

            const subjectLine = `Renewal reminder for ${currentParticulars}, Expiry Date: ${currentExpiryDate}`;
            GmailApp.sendEmail(row[inputSheetIndices["Mail (TO)"]], subjectLine, messageBody, options);

            const [rfc822id, timeStamp] = getMessageUrl(row[inputSheetIndices["Mail (TO)"]]);
            logSheetFillData(logSheet, currentParticulars, timeStamp, reminder.idx, rfc822id);
          }
        });
      }
    }
  });
}


