function getData(fileId) {
  const ss = SpreadsheetApp.openById(fileId);
  const inputSheetName = ss.getName();
  console.log("Spreadsheet Name", inputSheetName);
  const [client, subject, yearMonth] = inputSheetName.split("_").slice(1, -1);
  const [month, year] = yearMonth.split("'");
  const inputSheet = ss.getSheetByName('Summary');

  const masterHeader = inputSheet.getRange(3, 1, 1, inputSheet.getLastColumn()).getValues().flat();
  const masterData = inputSheet.getRange(1, 1, inputSheet.getLastRow(), inputSheet.getLastColumn()).getValues();
  const smeName = masterHeader.indexOf("Name of Tutor");
  const totHrs = masterHeader.indexOf("Total");

  // Headers for different types of shifts
  const headers = inputSheet.getRange(1, 1, inputSheet.getLastRow(), 1).getValues().flat();

  var onlineShiftDay = "Summary_Day Shift_Online+Extended";
  var onlineShiftNight = "Summary_Night Shift_Online+Extended";
  var dayNight = "Summary_Day+Night Shift_Online+Extended";
  var dayShiftPreSchedule = "Summary_Day Shift_Pre-Scheduled";
  var nightShiftPreSchedule = "Summary_Night Shift_Pre-Scheduled";

  const dayShiftSMEDetails = smeDetails(headers, masterData, onlineShiftDay, onlineShiftNight, smeName, totHrs);
  const nightShiftSMEDetails = smeDetails(headers, masterData, onlineShiftNight, dayNight, smeName, totHrs);
  const preDaySMEDetails = smeDetails(headers, masterData, dayShiftPreSchedule, nightShiftPreSchedule, smeName, totHrs);
  const preNightSMEDetails = smeDetails(headers, masterData, nightShiftPreSchedule, "Total Hours (Day+Night Shift)", smeName, totHrs);

  function rowCreation(smeDetails, mode) {
    const arr = [];
    for (const i of smeDetails) {
      const name = i[0];
      const totHrs = i[1];
      arr.push([name, mode, totHrs, client, subject, convertToLongMonth(month), 
                convertToFourDigitYear(year), mode.split(" ")[0].split("_")[1]]);
    }
    return arr;
  }

  const outputData1 = rowCreation(dayShiftSMEDetails, onlineShiftDay);
  const outputData2 = rowCreation(nightShiftSMEDetails, onlineShiftNight);
  const outputData3 = rowCreation(preDaySMEDetails, dayShiftPreSchedule);
  const outputData4 = rowCreation(preNightSMEDetails, nightShiftPreSchedule);

  // Set output arrays into output datasheet
  const outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Summary")
  const startRow = outputSheet.getLastRow() + 1;

  outputData = [...outputData1, ...outputData2, ...outputData3, ...outputData4]
  // outputData = [...outputData1, ...outputData2]

  // Set outputData1
  const numRows1 = outputData.length;
  const numColumns1 = outputData[0].length;
  outputSheet.getRange(startRow, 1, numRows1, numColumns1)
             .setHorizontalAlignment("center")
             .setVerticalAlignment("middle")
             .setBorder(true, true, true, true, true, true)
             .setWrap(true)
             .setFontFamily("Roboto")
             .setValues(outputData)

  return true;
  
}





function smeDetails(headers, masterData, startIndex, endIndex, smeName, totHrs){

  // const startIdx = headers.findIndex(header => typeof header === 'string' && header.includes(startIndex));
  // const endIdx = headers.findIndex(header => typeof header === 'string' && header.includes(endIndex));
  
  const particularShiftObject = headers
      .map((ele, index) => ({ header: ele, index }))
      .filter(({ header, index }) =>
        // index > startIdx && index < endIdx
        index > headers.indexOf(startIndex) && index < headers.indexOf(endIndex)
      );

  // console.log("Shift Object", particularShiftObject)
  const indices = getIndicesBetweenSrNoAndTotal(particularShiftObject);
  let smeDetails = masterData.filter((ele, idx) => indices.includes(idx+1))
                              .filter(ele => ele[smeName] && !ele[smeName].match(/^T\d/i) && ele[smeName] !== 'x' && ele[smeName] !== 'X')
                            //  .filter(ele => !ele[smeName].match(/^T\d/i) && ele[smeName] != 'x' && ele[smeName] != 'X')
                             .map(ele => [ele[smeName], ele[totHrs]]);

  return smeDetails
}


/**
 * Start with and check for Total - MAYBE a BETTER IDEA
 */

function getIndicesBetweenSrNoAndTotal(objArray) {
  // Find the index of 'Sr. No' and 'Total'
  // console.log("Objet array", objArray)
  const srNoIndex = objArray.find(item => item.header === 'Sr. No.').index + 1;
  const totalIndex = objArray.find(item => item.header === 'Total').index;
  // console.log(srNoIndex, totalIndex)
  // Filter indices between 'Sr. No' and 'Total'
  const indices = objArray
    .filter(({ index }) => index > srNoIndex && index < totalIndex)
    .map(({ index }) => index);
  return indices
}

function convertToFourDigitYear(twoDigitYear) {
  var currentYear = new Date().getFullYear();
  var prefix = currentYear.toString().slice(0, 2);
  var fourDigitYear = prefix + twoDigitYear;
  return parseInt(fourDigitYear);
}


function convertToLongMonth(shortMonth) {
  var date = new Date();
  var monthNumber = new Date(Date.parse(shortMonth + " 1, " + date.getFullYear())).getMonth();
  var longMonth = date.toLocaleString("en-US", { month: "long" });
  
  // Iterate over all the months until the short month is found
  while (date.getMonth() !== monthNumber) {
    date.setMonth(date.getMonth() + 1);
    longMonth = date.toLocaleString("en-US", { month: "long" });
  }
  return longMonth;
}

