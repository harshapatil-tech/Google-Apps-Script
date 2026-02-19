function copyAndProtectRange(sheet) {
  const source = sheet.getRange('A3:G11');
  const target = sheet.getRange('A23:G31');
  
  // 0) Remove any existing range protections on A23:G31
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  protections.forEach(function(protection) {
    if (protection.getRange().getA1Notation() === target.getA1Notation()) {
      protection.remove(); 
    }
  });
  
  // 1) Copy everything from A3:G11 into A23:G31
  source.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  
  // 2) Protect that new block
  const protection = target.protect().setDescription('Locked copy of A3:G11');
  
  // 3) Strip out all other editors...
  protection.getEditors().forEach(function(user) {
    protection.removeEditor(user);
  });
  
  // 4) …then add yourself back as sole editor
  const me = Session.getEffectiveUser();
  protection.addEditor(me);
  
  // 5) And disallow anyone in your domain from editing
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}


function getNext14Days(startDate) {
  const next14Days = [];
  const oneDayInMs = 24 * 60 * 60 * 1000; // Number of milliseconds in a day
  
  for (let i = 0; i < 14; i++) {
    // Calculate the new date by adding 'i' days in milliseconds
    const currentDate = new Date(startDate.getTime() + i * oneDayInMs);
    const formattedDate = formatDate(currentDate)
    next14Days.push(formattedDate);
  }
  return next14Days;
}


function formatDate(date) {
  const day = date.getDate().toString().padStart(2, '0');
  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const month = monthNames[date.getMonth()];
  const year = date.getFullYear();
  return day + '-' + month + '-' + year;
}



function get_Data_Indices_From_Sheet(sheet) {

  const dataRange = sheet.getDataRange().getValues();
  const headers = dataRange[0], data = dataRange.slice(1);
  return [createMapIndex(headers), data]
}


function createMapIndex(headers) {

  return headers.reduce((mapObj, currVal, currIndex) => {
    mapObj[currVal.trim()] = currIndex;
    return mapObj;
  }, {})
}

//date function sunday
function calculatePreviousTwoWeeksDateRange() {
  var today = new Date();
  var daysSinceSunday = today.getDay(); // 0 (Sunday) to 6 (Saturday)
  var lastSunday = new Date(today.getFullYear(), today.getMonth(), today.getDate() - daysSinceSunday);
  var startDate = new Date(lastSunday.getFullYear(), lastSunday.getMonth(), lastSunday.getDate() -13); // 13 days before last Sunday
  return {
    startDate: startDate,
    endDate: lastSunday
  };
}


function getTotalArrayLsit(values){
  const firstArray = values[0]; // Get the first array
  const lastArray = values[values.length - 1]; // Get the last array
  return firstArray.map((element, index) => [element, lastArray[index]]);
}


function getSeason(month) {
  switch (month) {
    case 'January':
    case 'February':
    case 'March':
    case 'April':
    case 'May':
      return 'Spring';
    case 'June':
    case 'July':
    case 'August':
      return 'Summer';
    case 'September':
    case 'October':
    case 'November':
    case 'December':
      return 'Fall';
    default:
      return 'Invalid month';
  }
}


function getFinancialYear(month, year) {
  // Create a mapping of month names to their corresponding indexes
  const monthMap = {
    January: 0,
    February: 1,
    March: 2,
    April: 3,
    May: 4,
    June: 5,
    July: 6,
    August: 7,
    September: 8,
    October: 9,
    November: 10,
    December: 11
  };

  // Get the month index from the month name
  const monthIndex = monthMap[month];

  // Determine the financial year based on the month index
  let financialYear;
  year = parseFloat(year)
  if (monthIndex >= 3) {
    // For months July to December, use the same year and the next year
    financialYear = year + "-" + (year + 1).toString().substring(2);
  } else {
    // For months January to June, use the previous year and the current year
    financialYear = (year - 1) + "-" + year.toString().substring(2);
  }

  return financialYear;
}


function getColumnLetter(columnNumber) {
  var dividend = columnNumber;
  var columnName = '';
  var modulo;

  while (dividend > 0) {
    modulo = (dividend - 1) % 26;
    columnName = String.fromCharCode(65 + modulo) + columnName;
    dividend = Math.floor((dividend - modulo) / 26);
  }

  return columnName;
}


function multiplyTwoColumns(col1, col2, row) {
  return `=${getColumnLetter(col1)}${row}*${getColumnLetter(col2)}${row}`;
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




function extractNumberFromString(inputString) {
    if (!inputString) return ""; // return empty string if undefined or null
    const matches = inputString.match(/\d+/);
    return matches ? matches[0] : "";
}



function dateValidation(range) {
    const validation = SpreadsheetApp.newDataValidation()
                                      .requireDate()
                                      .setAllowInvalid(true)
                                      .build();

    return range.setDataValidation(validation).setNumberFormat('dd-MMM-YYYY');
}




function copySheetToFolder(sourceFileId, targetFolderId, startDate) {
  
  // Open the source file
  var sourceFile = DriveApp.getFileById(sourceFileId);
  
  // Get the target folder
  var targetFolder = DriveApp.getFolderById(targetFolderId);
  
  // Make a copy of the file in the target folder
  const fourteenDaysBefore = new Date(startDate);
  fourteenDaysBefore.setDate(startDate.getDate())
  const endDate = new Date(fourteenDaysBefore);
  endDate.setDate(endDate.getDate() + 13);
  const name = `Brainfuse_Subjectwise_Hours_${formatDate(fourteenDaysBefore)}_${formatDate(endDate)}`
  const newSpreadsheet = sourceFile.makeCopy(name, targetFolder);
  var spreadsheet = SpreadsheetApp.openById(newSpreadsheet.getId())
  var sheets = spreadsheet.getSheets()
  sheets.forEach(sheet => {
        if (sheet.getName().trim() !== 'Calculus' && sheet.getName().trim() !== "Statistics" &&
          sheet.getName().trim() !== 'English' && sheet.getName().trim() !== "Chemistry" &&
          sheet.getName().trim() !== 'Physics' && sheet.getName().trim() !== "Biology" && sheet.getName().trim() !== "Writing_Lab" &&
          sheet.getName().trim() !== 'Finance' && sheet.getName().trim() !== "Economics" &&
          sheet.getName().trim() !== 'Computer Science' && sheet.getName().trim() !== "Intro Accounting" && sheet.getName().trim() !== "Summary"){
        
          spreadsheet.deleteSheet(sheet)
        }
  });
  Logger.log("File copied to folder with ID: " + targetFolderId);
}



function triggerGitHubWorkflow() {
  var owner = 'your_username';
  var repo = 'your_repository';
  var workflowId = 'your_workflow_id';
  var accessToken = 'your_personal_access_token';
  var url = 'https://api.github.com/repos/' + owner + '/' + repo + '/actions/workflows/' + workflowId + '/dispatches';
  var payload = {
    ref: 'main'
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'token ' + accessToken
    },
    payload: JSON.stringify(payload)
  };
  var response = UrlFetchApp.fetch(url, options);
  Logger.log(response.getContentText());
}



