function run () {
  const x = new DepartWise_Sheets_Creator();
  x.run();
}

class DepartWise_Sheets_Creator {

  constructor () {
    this.today = new Date();
    // Flag to create new spreadsheets
    this.createNewSpreadsheets = 0;
    this.monthSheetUpdate = 0;
    this.prevMonth = (this.today.getMonth() === 0) ? 11 : this.today.getMonth() - 1;
    this.prevMonthShort = new Date(this.today.getFullYear(),this.prevMonth, 1).toLocaleString("en-US", { month: "short" });
    this.year = (this.today.getMonth() === 0) ? this.today.getFullYear() - 1 : this.today.getFullYear();
    if(this.today.getDate() === 13 && this.today.getMonth() === 4) {
      this.createNewSpreadsheets = 1;
    }
    if (this.today.getDate() === 13)
      this.monthSheetUpdate = 1;
    this.monthlySheetUpdater = new Monthly_Sheet_Updater();
  }

  _getSubjectName(log) {
    let subject = log.split("_")[2]
    // Replace camelCase with space separated words
    subject = subject.replace(/([a-z])([A-Z])/g, '$1 $2');
    return subject
  }

  run () {
    if (this.createNewSpreadsheets) {
      for (const fileId of this._getPreviousYearFolder()) {
        this._forEachSpreadsheet(fileId);
      }
    }
    if (this.monthSheetUpdate === 1)
      for(const fileId of this._getPreviousYearFolder()) {
        const ss = SpreadsheetApp.openById(fileId);
        console.log(ss.getName())
        const deptName = this._getSubjectName(ss.getName());
        if (DEPARTMENTS.includes(deptName)) {
          const empForDept = this.monthlySheetUpdater.getDataByDept(deptName);
          // Get the names from previous month Schedules
          const scheduleFiles = this.findFilesByDeptName(deptName);
          if (scheduleFiles.length > 0) {
            const scheduleFile = scheduleFiles[0];
            // console.log(scheduleFile)
            const scheduleTutorNames = this.getTutorNames(scheduleFile);
            const filteredEmps = empForDept.filter(row=> {
              return scheduleTutorNames.includes(row[0])
            })            
            const sheet = ss.getSheetByName(`${this.prevMonthShort}_${this.year.toString().slice(-2)}`);
            this.monthlySheetUpdater.setName(sheet, filteredEmps);
          }
        }
      }
    


  }

  getTutorNames(file) {
    const sheet = SpreadsheetApp.open(file).getSheetByName("Tutor_Names");
    const [header, data] = CentralLibrary.get_Data_Indices_From_Sheet(sheet);
    const allTutorNames = this._getColumnValues(header, data, "Name of Tutors")
    return allTutorNames.slice(0, allTutorNames.findIndex(r=> r==="X"));
  }

  _getColumnValues(headers, rows, headerName) {
    const idx = headers[headerName];
    if (idx === undefined) throw new Error(`Header "${headerName}" not found.`);
    return rows.map(r => r[idx]).filter(v => v);
  }


  findFilesByDeptName(deptName) {
    const rootFolder = DriveUtil.findFolderById(SCHEDULES_ROOT_ID);
    const deptFolder = DriveUtil.findFolder(rootFolder, deptName);
    const year = this.year;
    const yearFolder = DriveUtil.findFolder(deptFolder, `${deptName}_${year}`);
    const prevMonth = this.prevMonth;
    const prevMonthShort = this.prevMonthShort;
    const monthFolderNomenclature = `${(prevMonth+1).toString().padStart(2, '0')}${prevMonthShort}-${year.toString().slice(-2)}`
    const monthFolder = DriveUtil.findFolder(yearFolder, monthFolderNomenclature);
    const scheduleFiles = DriveUtil.searchFiles(monthFolder, `BF_${deptName}`);
    return scheduleFiles;

      // this.monthlySheetUpdater.fillUpCurrMonthNames()
  }

  _forEachSpreadsheet(id) {
    const ss = SpreadsheetApp.openById(id);
    const sheets = ss.getSheets();
    this._createAnnualSheets(ss);
    this._deletePreviousYearSheets(ss);
  }


    /**
   * Create 12 monthly sheets (May→Apr) in the given spreadsheet,
   * fill rows 1-2 with Month and days, row 3 with detailed headers,
   * and rows 4-200 with the required formulas.
   */
  _createAnnualSheets(ss) {
    const startYear = this.today.getFullYear();
    const monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const startIndex = 3;  // May is index 4 (0-based)

    // detailed headers for row 3
    const headers = [
      'Name of Tutor',
      'Hrs per week',
      'Hrs per day',
      'Expected Days',
      'Expected hours',
      'Final Tutoring Hours',
      'Difference in hrs (Expected - Final Tutoring)',
      'Difference in Days (Expected - Final Tutoring)',
      'Difference in Hrs (TMS - Expected)',
      'Leaves (From Keka)',
      "Weekly off's",
      'Holidays (Days)',
      'Holidays Working (Hrs)',
      'Actual Tutoring Hours (From Timetable)',
      'TMS Hrs',
      'Extra Work (Days)',
      'Final Extra Work (Days)'
    ];

    const startRow = 4;
    const lastRow = 200;
    const numRows = lastRow - startRow + 1;

    for (let i = 0; i < 12; i++) {
      const mIndex = (startIndex + i) % 12;
      const yearOffset = Math.floor((startIndex + i) / 12);
      const year = startYear + yearOffset;
      const yy = year.toString().slice(-2);
      const sheetName = `${monthNames[mIndex]}_${yy}`;

      // get or create sheet
      let sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        sheet.clear();
      } else {
        sheet = ss.insertSheet(sheetName);
      }

      // rows 1 & 2
      const monthLabel = `1-${monthNames[mIndex]}`;
      const daysInMonth = new Date(year, mIndex + 1, 0).getDate();
      sheet.getRange('A1')
        .setBorder(true, true, true, true, true, true)          // all borders on
        .setHorizontalAlignment("center")                       // center-align
        .setFontFamily("Roboto")                                // Roboto font
        .setFontSize(11)                                        // size 11
        .setValue('Month');
      sheet.getRange('B1')
        .setBorder(true, true, true, true, true, true)         
        .setHorizontalAlignment("center")                       
        .setFontFamily("Roboto")                               
        .setFontSize(11)                                
      .setValue(monthLabel);
      sheet.getRange('A2')
          .setBorder(true, true, true, true, true, true)         
          .setHorizontalAlignment("center")                       
          .setFontFamily("Roboto")                               
          .setFontSize(11)  
          .setValue('No of days in Month');
      sheet.getRange('B2')
          .setBorder(true, true, true, true, true, true)         
          .setHorizontalAlignment("center")                       
          .setFontFamily("Roboto")                               
          .setFontSize(11)  
          .setValue(daysInMonth);

      // row 3 headers
      sheet.getRange(3, 1, 1, headers.length)
            .setBorder(true, true, true, true, true, true)          // all borders on
            .setHorizontalAlignment("center")                       // center-align
            .setFontFamily("Roboto")                                // Roboto font
            .setFontSize(11)                                        // size 11
            .setValues([headers]);

      // rows 4–200 formulas
      sheet.getRange(startRow, 3, numRows, 1).setFormula('=B4/6');
      sheet.getRange(startRow, 4, numRows, 1).setFormula('=$B$2-$J4-$K4-$L4');
      sheet.getRange(startRow, 5, numRows, 1).setFormula('=$C4*$D4');
      sheet.getRange(startRow, 6, numRows, 1).setFormula('=$N4-$M4');
      sheet.getRange(startRow, 7, numRows, 1).setFormula('=F4-E4');
      sheet.getRange(startRow, 8, numRows, 1).setFormula('=IFERROR((G4/C4),"0.00")');
      sheet.getRange(startRow, 9, numRows, 1).setFormula('=(O4-M4)-E4');
      sheet.getRange(startRow, 16, numRows, 1).setFormula('=IFERROR((I4/C4),"0.00")');

      sheet.hideColumns(3, 7)
    }
  }


  /**
   * Remove sheets named Apr-YY … Mar-YY for the previous financial year.
  */
  _deletePreviousYearSheets(ss) {
    const monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const today = new Date(); // Assuming 'today' is properly initialized
    const currYear = today.getFullYear();
    const prevStart = (today.getMonth() >= 3) ? currYear - 1 : currYear - 2; // April (index 3) check

    const namesToDelete = [];
    for (let i = 11; i >= 2; i--) { // Loop from Dec to Mar
      const mIndex = i % 12;
      const year = mIndex >= 3 ? prevStart : prevStart + 1;
      const yy = year.toString().slice(-2);
      namesToDelete.push(`${monthNames[mIndex]}_${yy}`);
    }

    ss.getSheets().forEach(sheet => {
      if (namesToDelete.includes(sheet.getName())) {
        ss.deleteSheet(sheet);
      }
    });
  }



  _getPreviousYearFolder () {
    const prevFinYr = getPrevFinYear(this.today);
    const currFinYear = getCurrFinYear(this.today); 
    // const rootFolder = DriveApp.getFolderById(ROOT_FOLDER_ID);
    const rootFolder = DriveUtil.findFolderById(ROOT_FOLDER_ID)
    const prevYearFolder = DriveUtil.findFolder(rootFolder, prevFinYr);
    const currYearFolder = DriveUtil.getOrCreateFolder(rootFolder, currFinYear);
    const copiedFiles = DriveUtil.copyFilesFromFolderToFolder(prevYearFolder, currYearFolder, prevFinYr, "2025-26");
    return copiedFiles;
  }


}






