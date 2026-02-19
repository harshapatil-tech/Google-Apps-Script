function main() {
  const SPREADSHEET = SpreadsheetApp.openById("1XH5eGI4thuz-yAsHDeVbI90KFX2AWee3w556yUcGe6I");
  const backendSheet = SPREADSHEET.getSheetById(204064486);
  const webscrapingBooleanRange = backendSheet.getRange("B1");
  const webscrapingBooleanValue = webscrapingBooleanRange.getValue();

  const {startDate, endDate} = calculatePreviousTwoWeeksDateRange();
  console.log(`${startDate.toDateString()} to ${endDate.toDateString()}`);

  if (webscrapingBooleanValue === 1) {
    
    // Fetches English data to "English Timesheet Data" spreadsheet for Deboo
    englishData(SPREADSHEET);
    copyData();
    const data = new getData(SPREADSHEET, startDate);
    const SUBJETS = [
      "Calculus", "Statistics", "Physics", "Intro Accounting",
      "Chemistry", 
      "Biology", 
      "Computer Science",
      "Finance","Economics", "English"
                    ]

    console.log("Executed");
    SUBJETS.forEach(department => {
      if (department === "English")
        data.createSheetFormat(department, 0);
      else
        data.createSheetFormat(department, 1);
      if (department === "English") {
        data.englishWritingLab();
        data.createSheetFormat(department, 0);
      }
    });

    const [invoiceNum, fileId] = createInvoice();
    updateDashboard(fileId, startDate, endDate, invoiceNum)
    copySheetToFolder(SPREADSHEET.getId(), "1IhydA4xQXfP0CQ4zMFBwWcbKiiQJbInH", startDate);
    webscrapingBooleanRange.setValue([[0]]) 
  }
  settingBrainFuse_Occupancy_non_Occupanacy();
}


class getData {

  constructor(spreadsheet, startDate) {
    this.SPREADSHEET = spreadsheet
    this.scriptTimeZone = Session.getScriptTimeZone();
    this.startDate = startDate
  }


  createSheetFormat(subject, dualFlag = 1) {

    const subjectSheet = this.SPREADSHEET.getSheetByName(subject);

    if(subject === "Calculus")
      subject = "Mathematics"
    
    subjectSheet.clear();
    subjectSheet.getDataRange().clearDataValidations()
    
    const [singleSubjectUniqueAccountNumbers, singleSubjectData] = this.getSubjectData(subject, "single");
    this.tabularData(subjectSheet, singleSubjectUniqueAccountNumbers, singleSubjectData, 1);
    
    if (dualFlag ==1 ) {
      const [dualSubjectUniqueAccountNumbers, dualSubjectData] = this.getSubjectData(subject, "dual");
      this.tabularData(subjectSheet, dualSubjectUniqueAccountNumbers, dualSubjectData, subjectSheet.getLastRow()+7);
    }
    
  }


  englishWritingLab() {

    const writingLabSheet = this.SPREADSHEET.getSheetByName("Writing_Lab");
    writingLabSheet.clear();
    writingLabSheet.deleteRows(2, writingLabSheet.getMaxRows()-1);

    const allData = this.SPREADSHEET.getSheetByName("Summary").getDataRange().getValues();
    const subjectData = allData.filter(r => r[0].trim().toLowerCase() === "english" && r[3].trim().toLowerCase() === "task assignment");
    
    //make unique list of tutor account no.
    const uniqueAccountNumbers = [... new Set(subjectData.map( r => r[1]))].sort((a,b) => a-b);
    const subjectObject = {};

    subjectData.forEach(r => {
      const date = Utilities.formatDate(r[4], this.scriptTimeZone, "dd-MMM-yyyy"); // Assuming the date is at index 4
      const status = r[3]; // Assuming the status is at index 3
      const tutorId = r[1]; // Assuming the tutor ID is at index 1

      if (!subjectObject[date]) {
        subjectObject[date] = {};
      }

      if (!subjectObject[date][tutorId]) {
        subjectObject[date][tutorId] = {
          values : 0,  //total hours/essays
          count : 0    //number of records
        };
      }

      subjectObject[date][tutorId].values += r[6];
      subjectObject[date][tutorId].count ++;
      
    });
    // Logger.log(subjectObject["01-Dec-2023"])

    let startRow = 1;

    uniqueAccountNumbers.forEach((accountNumber, index) => {

      applyCustomFormatting(writingLabSheet.getRange(startRow+1, index+2), {"bgColor":"#6fb43f", "fontWeight":"bold"}).setValues([[accountNumber]]);
      applyCustomFormatting(writingLabSheet.getRange(startRow+2, index+2), {"bgColor":"#6fb43f", "fontWeight":"bold"}).setValues([["# of Essays"]]);
      
    })
    applyCustomFormatting(writingLabSheet.getRange(startRow+1, 1, 2), {"bgColor":"#6fb43f", "fontWeight":"bold"}).merge().setValue("Date");

    const next14Days = getNext14Days(this.startDate);
    
    writingLabSheet.getRange(startRow+3, 1, 14, 1).setValues(next14Days.map(r => [r]));
    let lastRow = writingLabSheet.getLastRow();

    let columnTotalRange = writingLabSheet.getRange(lastRow+3, 1)
    columnTotalRange.setValue("Total");
    dateValidation(writingLabSheet.getRange(startRow+3, 1, 14, 1))
    applyCustomFormatting(writingLabSheet.getRange(startRow+3, 1, 14+3, 1), {"bgColor":"#6fb43f", "fontWeight":"bold"})
    applyCustomFormatting(writingLabSheet.getRange(startRow+3, 2, 14+3, uniqueAccountNumbers.length))

    this.setDataWritingLab(writingLabSheet, subjectObject, startRow, uniqueAccountNumbers.length);

    lastRow = writingLabSheet.getLastRow();
    writingLabSheet.insertRowAfter(lastRow);
    writingLabSheet.getRange(lastRow+1, 1, 1, writingLabSheet.getLastColumn()).clearFormat();

    startRow = lastRow + 7;

    uniqueAccountNumbers.forEach((accountNumber, index) => {

      applyCustomFormatting(writingLabSheet.getRange(startRow+1, index+2), {"bgColor":"#6fb43f", "fontWeight":"bold"}).setValues([[accountNumber]]);
      applyCustomFormatting(writingLabSheet.getRange(startRow+2, index+2), {"bgColor":"#6fb43f", "fontWeight":"bold"}).setValues([["# of Hours"]]);
      
    })
    applyCustomFormatting(writingLabSheet.getRange(startRow+1, 1, 2), {"bgColor":"#6fb43f", "fontWeight":"bold"}).merge().setValue("Date");

    writingLabSheet.getRange(startRow+3, 1, 14, 1).setValues(next14Days.map(r => [r]));
    lastRow = writingLabSheet.getLastRow();

    columnTotalRange = writingLabSheet.getRange(lastRow+3, 1)
    columnTotalRange.setValue("Total")
    dateValidation(writingLabSheet.getRange(startRow+3, 1, 14, 1))
    applyCustomFormatting(writingLabSheet.getRange(startRow+3, 1, 14+3, 1), {"bgColor":"#6fb43f", "fontWeight":"bold"})
    
    applyCustomFormatting(writingLabSheet.getRange(startRow+3, 2, 14+3, uniqueAccountNumbers.length))

    this.setDataWritingLab(writingLabSheet, subjectObject, startRow, uniqueAccountNumbers.length, false);
    

  }

  setDataWritingLab(sheet, data, headerRow, totalAccounts, essayHours=true) {
    
    const dateValues = sheet.getRange(headerRow + 3, 1, 14, 1).getDisplayValues().flat();
    const accounts = sheet.getRange(headerRow + 1, 2, 1, totalAccounts).getDisplayValues().flat();
    
    dateValues.forEach((date, index) => {
      
      if (data.hasOwnProperty(date)) {
        const outerValues = data[date];
        for (const accountId of accounts) {
          const accountID_columnNum = accounts.indexOf(accountId) + 2;
          if (outerValues.hasOwnProperty(accountId)) {
            const innerValues = outerValues[accountId];
            if (essayHours === false)
              sheet.getRange(index + 3 + headerRow, accountID_columnNum).setValue(innerValues.values);
            else
              sheet.getRange(index + 3 + headerRow, accountID_columnNum).setValue(innerValues.count);
          }
          else {
            sheet.getRange(index + 3 + headerRow, accountID_columnNum).setValue(0);
          }
        }
      }
    });
    const lastRow = sheet.getLastRow();
    for(let index=2; index<=totalAccounts+1; index++) {
      const start = sheet.getRange(headerRow+3, index).getA1Notation();
      const end = sheet.getRange(headerRow+3+14-1, index).getA1Notation();
      sheet.getRange(lastRow, index).setFormula(`=SUM(${start}:${end})`);
    }
    applyCustomFormatting(sheet.getRange(lastRow, 2, 1, sheet.getLastColumn()-1), {"bgColor":"#6fb43f", "fontWeight":"bold" });
  }

  setData(sheet, data, headerRow, totalColumn) {
    Logger.log(sheet.getRange(headerRow + 3, 1, 14, 1).getA1Notation())
    const dateValues = sheet.getRange(headerRow + 3, 1, 14, 1).getDisplayValues().flat();
    const accounts = sheet.getRange(headerRow + 1, 1, 1, totalColumn+2).getDisplayValues().flat();  // POSSSIBLE FOR CHANGE - COLUMN RANGE
    
    dateValues.forEach((date, index) => {

      if (data.hasOwnProperty(date)) {
        const outerValues = data[date];

        let totalOccupancy = 0;
        let totalNoOccupancy = 0;
        let totalOffline = 0;
        
        for (const accountId of accounts) {
          
          if (outerValues.hasOwnProperty(accountId)) {
            
            const innerValues = outerValues[accountId];
            const accountID_columnNum = accounts.indexOf(accountId) + 1;
            // Logger.log(innerValues)
            for (const value of innerValues) {
              if (Object.keys(value)[0].trim().toLowerCase() === "std - occupancy") {
                const occupancyValue = Object.values(value)[0];
                const range = sheet.getRange(index + 3 + headerRow, accountID_columnNum)
                const val = range.getValue();
                range.setValue(val + occupancyValue);
                totalOccupancy += occupancyValue;
              } else if (Object.keys(value)[0].trim().toLowerCase() === "std - no-occupancy") {
                const noOccupancyValue = Object.values(value)[0];
                const range = sheet.getRange(index + 3 + headerRow, accountID_columnNum+1)
                const val = range.getValue();
                range.setValue(val + noOccupancyValue);
                totalNoOccupancy += noOccupancyValue;
              } else if (Object.keys(value)[0].trim().toLowerCase() === "offline") {
                const offlineValue = Object.values(value)[0];
                const range = sheet.getRange(index + 3 + headerRow, accountID_columnNum+2)
                const val = range.getValue();
                range.setValue(val + offlineValue);
                totalOffline += offlineValue;
              }

            }
          }
          
        }

        // Set the total values in the last row for each date
        sheet.getRange(index + 3 + headerRow, totalColumn).setValue(totalOccupancy);
        sheet.getRange(index + 3 + headerRow, totalColumn + 1).setValue(totalNoOccupancy);
        sheet.getRange(index + 3 + headerRow, totalColumn + 2).setValue(totalOffline);
      }
    });

    const lastRow = sheet.getLastRow()
    for(let index=2; index<=accounts.length; index++) {

      const start = sheet.getRange(headerRow+3, index).getA1Notation();
      const end = sheet.getRange(headerRow+3+14-1, index).getA1Notation();
      sheet.getRange(lastRow, index).setFormula(`=SUM(${start}:${end})`);
    }
    applyCustomFormatting(sheet.getRange(lastRow, 2, 1, sheet.getLastColumn()-1), {"bgColor":"#6fb43f", "fontWeight":"bold" });

  }

  //Make structure of sheet
  tabularData(subjectSheet, uniqueAccountNumbers, subjectData, startRow=1, buffer=0) {
  
    if (Object.keys(subjectData).length > 0) {
      uniqueAccountNumbers.forEach((accountNumber, index) => {
        // Find the columns for the current account number
        const columns = [3 * index + 2, 3 * index + 3, 3 * index + 4];

        // Set values for the current account number
        subjectSheet.getRange(startRow+1, columns[0], 1, 3).setValues([[accountNumber, "", ""]]).merge();

        // Set labels for the next row of unmerged columns
        subjectSheet.getRange(startRow+2, columns[0]).setValue("Std - Occupancy");
        subjectSheet.getRange(startRow+2, columns[1]).setValue("Std - No-Occupancy");
        subjectSheet.getRange(startRow+2, columns[2]).setValue("Offline");
      });

      const totalColumn = 3 * uniqueAccountNumbers.length + 2;
      const rowTotalRange = subjectSheet.getRange(startRow+1, totalColumn, 1, 3)
      const rowTotalColumnNum = rowTotalRange.getColumn();
      rowTotalRange.setValues([["Total", "", ""]]).merge();
      subjectSheet.getRange(startRow + 2, totalColumn).setValue("Std - Occupancy");
      subjectSheet.getRange(startRow + 2, totalColumn + 1).setValue("Std - No-Occupancy");
      subjectSheet.getRange(startRow + 2, totalColumn + 2).setValue("Offline");
      
      subjectSheet.getRange(startRow+1, 1, 2).merge().setValue("Date");
      applyCustomFormatting(subjectSheet.getRange(startRow+1, 1, 2, totalColumn+2), {"bgColor":"#6fb43f", "fontWeight":"bold"})
      
      const next14Days = getNext14Days(this.startDate);
      const start = next14Days[0].split("-").slice(0,2).join("_");
      const end = next14Days[next14Days.length-1].split("-").slice(0,2).join("_");
      const title = `Summary_${start}_${end}`
    
      
      subjectSheet.getRange(startRow+3, 1, 14, 1).setValues(next14Days.map(r => [r]));
      let lastRow = subjectSheet.getLastRow();
      
      const columnTotalRange = subjectSheet.getRange(lastRow+3, 1)
      columnTotalRange.setValue("Total")
      
      dateValidation(subjectSheet.getRange(startRow+3, 1, 14, 1))
      applyCustomFormatting(subjectSheet.getRange(startRow+3, 1, 14+3, 1), {"bgColor":"#6fb43f", "fontWeight":"bold"})
      applyCustomFormatting(subjectSheet.getRange(startRow+3, 2, 14+3, rowTotalColumnNum+1));
      //call setData to fill actual values
      this.setData(subjectSheet, subjectData, startRow, rowTotalColumnNum);

      lastRow = subjectSheet.getLastRow();
      subjectSheet.insertRowAfter(lastRow);
      subjectSheet.getRange(lastRow+1, 1, 1, subjectSheet.getLastColumn()).clearFormat();
    }
  }


  getSubjectData(subject, tutoringType="single") {
    
    const allData = this.SPREADSHEET.getSheetByName("Summary").getDataRange().getValues();
    const subjectData = allData.filter(r => r[0].trim().toLowerCase() === subject.trim().toLowerCase());

    const activitySingle = subjectData.filter(r => r[2].trim().toLowerCase() === tutoringType);
    // Logger.log(activitySingle);
    
    //Make list of unique Tutor ids
    const uniqueAccountNumbers = [... new Set(activitySingle.map( r => r[1]))].sort((a,b) => a-b);
    const subjectObject = {};

    activitySingle.forEach(r => {
      const date = Utilities.formatDate(r[4], this.scriptTimeZone, "dd-MMM-yyyy"); // Assuming the date is at index 4
      const status = r[3]; // Assuming the status is at index 3
      const tutorId = r[1]; // Assuming the tutor ID is at index 1

      if (!subjectObject[date]) {
        subjectObject[date] = {};
      }

      if (!subjectObject[date][tutorId]) {
        subjectObject[date][tutorId] = [];
      }

      if (subject.trim().toLowerCase() !== "english") {
        if (status.trim().toLowerCase() === "ia-tutored")
          subjectObject[date][tutorId].push({"Std - Occupancy": r[6]});
        else if (status.trim().toLowerCase() === "ia-waited")
          subjectObject[date][tutorId].push({"Std - No-Occupancy": r[6]});
        else if (status.trim().toLowerCase() === "task assignment")
          subjectObject[date][tutorId].push({"Offline": r[6]});
      } else {
        if (status.trim().toLowerCase() === "ia-tutored")
          subjectObject[date][tutorId].push({"Std - Occupancy": r[6]});
        else if (status.trim().toLowerCase() === "ia-waited")
          subjectObject[date][tutorId].push({"Std - No-Occupancy": r[6]});
      }
    });
    return [uniqueAccountNumbers, subjectObject];
  }
}









function createTrigger() {
  // Delete any existing triggers for 'doThings' to prevent duplicates
  CentralLibrary.deleteTriggers('main');

  // Create a new time-based trigger for 'doThings'
  ScriptApp.newTrigger('main')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.TUESDAY)
    .atHour(20) // 20 represents 8 PM in 24-hour format
    .create();
}


function englishData(inputSpreadsheet){
  const scraperSummarySheet = inputSpreadsheet.getSheetByName("Summary");
  const scraperHeader = scraperSummarySheet.getRange(1, 1, 1, scraperSummarySheet.getLastColumn()).getValues().flat();
  const scraperData = scraperSummarySheet
                        .getRange(2, 1, scraperSummarySheet.getLastRow(), scraperSummarySheet.getLastColumn())
                        .getValues();

  const englishTimeSheet = SpreadsheetApp
                                .openById("1biOn_YoHLtHWDNjixRqkApeqXJEAnjKWy_S_b618adE")
                                .getSheetByName("Data");
  
  const dataRange = englishTimeSheet.getRange(2, 1, englishTimeSheet.getLastRow(), englishTimeSheet.getLastColumn()-1);
  dataRange.clearContent();
  const filteredEnglishData = scraperData
                                .filter(row => row[scraperHeader.indexOf("Department")].trim().toLowerCase() == "english");
  
  englishTimeSheet.getRange(2, 1, filteredEnglishData.length, filteredEnglishData[0].length).setValues(filteredEnglishData);
}