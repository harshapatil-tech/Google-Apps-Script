function myFun(){
  const processor = new SpreadsheetProcessor("1HAlOhOdIcZ0xlBWO15wBmMH_bAqnQpvk_psX5WwGDyg");
  processor.processShifts();
  return true;
}

// function getData(fileId) {
//   const processor = new SpreadsheetProcessor(fileId);
//   processor.processShifts();
//   return true;
// }

class SpreadsheetProcessor {
  constructor(fileId) {
    this.ss = SpreadsheetApp.openById(fileId);
    this.inputSheetName = this.ss.getName();
    [this.client, this.subject, this.yearMonth] = this.inputSheetName.split("_").slice(1, -1);
    [this.month, this.year] = this.yearMonth.split("'");
    this.inputSheet = this.ss.getSheetByName('Summary');
    // console.log(this.masterHeader);

    this.inputSheetHeader = this.inputSheet.getRange(3, 1, 1, this.inputSheet.getLastColumn()).getValues().flat();
    // console.log(this.inputSheetHeader);
    this.inputSheetData = this.inputSheet.getRange(1, 1, this.inputSheet.getLastRow(), this.inputSheet.getLastColumn()).getValues();
    this.smeNameIndex = this.inputSheetHeader.indexOf("Name of Tutor");
    this.totalHoursIndex = this.inputSheetHeader.indexOf("Total");
  }
  
  //this method to convert two digit year to four digit year
  static convertToFourDigitYear(twoDigitYear) {
    const currentYear = new Date().getFullYear();       //get the current year
    const prefix = currentYear.toString().slice(0, 2);    //extracts the first two digit of the current year
    return parseInt(prefix + twoDigitYear);     //then return the four digit year
  }

  //this method convert short month name to long month name
  static convertToLongMonth(shortMonth) {
    const date = new Date(Date.parse(shortMonth + " 1"));
    return date.toLocaleString("en-US", { month: "long" });     //return the full month name in english
  }

  smeDetails(headers, inputSheetData, startIndex, endIndex) {
    const particularShiftObject = headers
      .map((ele, index) => ({ header: ele, index }))
      .filter(({ header, index }) =>
        index > headers.indexOf(startIndex) && index < headers.indexOf(endIndex)
      );
    // console.log(particularShiftObject)
    // console.log(inputSheetData);
    const indices = this.getIndicesBetweenSrNoAndTotal(particularShiftObject);
    return inputSheetData
      .filter((_, idx) => indices.includes(idx + 1))
      .filter(row => row[this.smeNameIndex] && !row[this.smeNameIndex].match(/^T\d/i) && row[this.smeNameIndex].toLowerCase() !== 'x')
      .map(row => [row[this.smeNameIndex], row[this.totalHoursIndex]]);
  }

  getIndicesBetweenSrNoAndTotal(objArray) {
    const srNoIndex = objArray.find(item => item.header === 'Sr. No.').index + 1;   //find the index of the sr no column
    const totalIndex = objArray.find(item => item.header === 'Total').index;    //find the index of the total column
    return objArray
      .filter(({ index }) => index > srNoIndex && index < totalIndex)
      .map(({ index }) => index);
  }

  //this method to create rows for output based on sme details and mode
  rowCreation(smeDetails, mode) {
    return smeDetails.map(([name, totHrs]) => [
      name,
      mode,
      totHrs,
      this.client,
      this.subject,
      SpreadsheetProcessor.convertToLongMonth(this.month),
      SpreadsheetProcessor.convertToFourDigitYear(this.year),
      mode.split(" ")[0].split("_")[1]
    ]);
  }

  processShifts() {
    const headers = this.inputSheet.getRange(1, 1, this.inputSheet.getLastRow(), 1).getValues().flat();
    // let [headers, data] = CentralLibrary.get_Data_Indices_From_Sheet(this.inputSheet, 2)

    //various shift modes
    const onlineShiftDay = "Summary_Day Shift_Online+Extended";
    const onlineShiftNight = "Summary_Night Shift_Online+Extended";
    const dayNight = "Summary_Day+Night Shift_Online+Extended";
    const dayShiftPreSchedule = "Summary_Day Shift_Pre-Scheduled";
    const nightShiftPreSchedule = "Summary_Night Shift_Pre-Scheduled";
    // console.log(headers);
    // headers = Object.keys(headers);

    //extract sme details for different shifts
    const dayShiftSMEDetails = this.smeDetails(headers, this.inputSheetData, onlineShiftDay, onlineShiftNight, this.smeNameIndex, this.totalHoursIndex);
    const nightShiftSMEDetails = this.smeDetails(headers, this.inputSheetData, onlineShiftNight, dayNight, this.smeNameIndex, this.totalHoursIndex);
    const preDaySMEDetails = this.smeDetails(headers, this.inputSheetData, dayShiftPreSchedule, nightShiftPreSchedule, this.smeNameIndex, this.totalHoursIndex);
    const preNightSMEDetails = this.smeDetails(headers, this.inputSheetData, nightShiftPreSchedule, "Total Hours (Day+Night Shift)", this.smeNameIndex, this.totalHoursIndex);

    //output data for combining rows for all shifts
    const outputData = [
      ...this.rowCreation(dayShiftSMEDetails, onlineShiftDay),
      ...this.rowCreation(nightShiftSMEDetails, onlineShiftNight),
      ...this.rowCreation(preDaySMEDetails, dayShiftPreSchedule),
      ...this.rowCreation(preNightSMEDetails, nightShiftPreSchedule)
    ];

    this.writeToOutputSheet(outputData);
  }

  writeToOutputSheet(outputData) {
    const outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Summary");    //get the summary sheet for the output
    const startRow = outputSheet.getLastRow() + 1;

    outputSheet.getRange(startRow, 1, outputData.length, outputData[0].length)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true)
      .setWrap(true)
      .setFontFamily("Roboto")
      .setValues(outputData);
  }
}







// function getData(fileId){

// }

// class DataProcessor{
//   constructor(fileId){
//     this.ss = SpreadsheetApp.openById(fileId);
//     this.inputSheetName = this.ss.getName();
//     this.inputSheet = this.ss.getSheetByName('Summary');
//     const [client, subject, yearMonth] = this.inputSheetName.split("_").slice(1, -1);
//     const [month, year] = yearMonth.split("'");
//     this.client = client;
//     this.subject = subject;
//     this.month = month;
//     this.year = year;

//     [this.masterHeader, this.masterData] = CentralLibrary.get_Data_Indices_From_Sheet(this.inputSheet);
//     this.smeName = masterHeader["Name of Tutor"];
//     this.totHrs = masterHeader["Total"];

    


//   }
// }