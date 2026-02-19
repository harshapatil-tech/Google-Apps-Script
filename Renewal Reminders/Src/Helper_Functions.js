function logSheetFillData(sheet, particulars,timeStamp, remNo, rfcId){

  const data = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  const headers = data[0];
  
  const srNoIdx = headers.indexOf("Sr. No.") + 1;
  const particularsIdx = headers.indexOf("Particulars") + 1;
  const timeStampIdx = headers.indexOf("Email Sent Timestamp") + 1;
  const remNoIdx = headers.indexOf("Reminder No.") + 1;
  const rfcIdIdx = headers.indexOf("RFC ID") + 1;

  const lastRow = sheet.getLastRow();
  const lastSrNo = sheet.getRange(lastRow, srNoIdx).getValue();

  if (data.length == 1)
    sheet.getRange(lastRow+1, srNoIdx).setWrap(true).setValue(1)
  else
    sheet.getRange(lastRow+1, srNoIdx).setWrap(true).setValue(lastSrNo+1)
  sheet.getRange(lastRow+1, particularsIdx).setWrap(true).setValue(particulars)
  sheet.getRange(lastRow+1, timeStampIdx).setWrap(true).setValue(timeStamp)
  sheet.getRange(lastRow + 1, remNoIdx).setWrap(true).setValue(getReminderText(remNo));
  sheet.getRange(lastRow+1, rfcIdIdx).setWrap(true).setValue(rfcId)
}


function getReminderText(remNo) {
  if (remNo === 7) {
    return "Reminder 1";
  } else if (remNo === 8) {
    return "Reminder 2";
  } else if (remNo === 9) {
    return "Reminder 3";
  }
}


function getDaysDifference(targetDate, today) {
  // Normalize the dates to midnight
  targetDate.setHours(0, 0, 0, 0);
  today.setHours(0, 0, 0, 0);

  var differenceInMilliseconds = targetDate.getTime() - today.getTime();
  var differenceInDays = Math.floor(differenceInMilliseconds / (1000 * 60 * 60 * 24));
  return differenceInDays;
}


const getMessageUrl = (currentMailTo) => {
  const threads = GmailApp.search(`to:${currentMailTo}`, 0, 1);
  const messageId = threads[0].getId()
  const timeStamp = threads[0].getLastMessageDate()
  const message = GmailApp.getMessageById(messageId);
  const rfc822Id = message.getHeader('Message-Id');
  const searchQuery = `rfc822msgid:${rfc822Id.replace(/[<>]/g, '')}`;
  return [searchQuery, timeStamp];
};



function getLastRowWithEntry(sheet) {

  var lastRow = sheet.getLastRow();
  var headers = sheet.getRange(1, 1, 1, 6).getValues()[0]; // Assuming headers are in the first row
  
  var srNoIndex = headers.indexOf("Sr. No.") + 1;

  let lastIndex = 0;
  var range = sheet.getRange(2, srNoIndex, lastRow - 1, 1); // Assuming data starts from row 2
  var values = range.getValues().flat();

  for (var i = values.length - 1; i >= 0; i--) {
    if (values[i] !== "") {
      lastIndex = i + 2; // Adding 2 to account for the header row offset and 0-based index
      break;
    }
  }
  return lastIndex;
}


function createIndexMap(headers) {
  return headers.reduce((map, val, index) => {
    map[val] = index;
    return map;
  }, {});
}



function addMonthsToDate(inputDate, numMonths) {
  const newDate = new Date(inputDate);
  newDate.setHours(0, 0, 0, 0)
  // Calculate the new month and year
  const currentMonth = newDate.getMonth();
  const newMonth = (currentMonth + numMonths) % 12; // Get the new month
  const newYear = newDate.getFullYear() + Math.floor((currentMonth + numMonths) / 12); // Get the new year
  
  // Set the new month and year
  newDate.setMonth(newMonth);
  newDate.setFullYear(newYear);
  
  return newDate;
}


