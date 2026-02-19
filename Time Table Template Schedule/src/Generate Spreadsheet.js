// function genNewSpreadsheet () {
//   const activeSpreadsheet = new ActiveSpreadsheet();
//   const deptValue = activeSpreadsheet.fetchValue("B4");
//   const client    = activeSpreadsheet.fetchValue("C4");
//   const ssType    = activeSpreadsheet.fetchValue("D4");
//   const empMap    = Helper.mapSubjectToEmployees();

//   // 1) decide which departments to run
//   const departments = deptValue === "all"
//     ? SUBJECTS
//     : [ deptValue ];

//   const created = [];
//   const notCreated = []

//   // 2) for each dept, run the right routine
//   departments.forEach(rawDept => {
//     const dept = rawDept.trim().toLowerCase();
//     if (ssType === "timetable") {
//       const manager = new SpreadsheetManager(dept, empMap, client, ssType)
//       const createFlag = manager.runTimetable();
//       if (createFlag === 1)
//         created.push(Helper.capitalizeFirstLetter(dept));
//       else if (createFlag === 0)
//         notCreated.push(Helper.capitalizeFirstLetter(dept));
//     }
//     else if (ssType === "qc") {
//       const qcManager = new CreateQC(dept, client, ssType)
//       qcManager.copyIfNeeded();
//     }
//     else {
//       throw new Error(`Unknown sheet type: ${ssType}`);
//     }
//   });

//   const createdLog    = created.length > 0 ? `Timetable(s) created for the client: ${client.toUpperCase()}, department: ${created.join(", ")}` : "";
//   const notCreatedLog = notCreated.length > 0 ? `Timetable(s) for the client: ${client.toUpperCase()}, department ${notCreated.join(", ")} already exists` : "";
//   const finalLog = createdLog + "\n" + notCreatedLog;
//   SpreadsheetApp.getUi().alert(finalLog);
// }




class CreateQC {


  constructor (dept, client, ssType) {
    this.dept = dept, this.client = client, this.ssType = ssType;
    this.rootFolder = DriveApp.getFolderById(GOOGLE_DRIVE_FOLDER_ID);
    this.nextYear = DateUtil.nextYear();
  }

  copyIfNeeded () {
    const nextMonthTimetable = this._getTimetableNextMonth();
    console.log(nextMonthTimetable)
    if (nextMonthTimetable === undefined) {
      SpreadsheetApp.getUi().alert("No Time table found for the next month");
      return;
    }

    const deptFolder  = DriveUtil.getChildFolder(this.rootFolder, this.dept);
    const finalQC  = DriveUtil.getChildFolder(deptFolder, "final qc")
    const yearFolder = DriveUtil.getChildFolder(finalQC, this.nextYear);
    const monthFolder = DriveUtil.getOrCreateFolder(yearFolder, DateUtil.nextMonthKey());
    const timeTableName = nextMonthTimetables.getName();
    const copiedFile = DriveUtil.getOrCreateFile(monthFolder, nextMonthTimetable[0], timeTableName);
    if (copiedFile.action === "exists") {
      SpreadsheetApp.getUi().alert("QC for the client and month already exists");
    }
  
    
  }

  _getTimetableNextMonth () {
    const deptFolder  = DriveUtil.getChildFolder(this.rootFolder, this.dept);
    const yearFolder  = DriveUtil.getChildFolder(deptFolder, this.nextYear);
    const nextMonthTimetable = DriveUtil.getChildFolder(yearFolder, DateUtil.nextMonthKey().toLowerCase());
    console.log(nextMonthTimetable.getName())
    const fileName = `Time table_${this.client.toUpperCase()}_${Helper.capitalizeFirstLetter(this.dept)}_${DateUtil.nextMonthName("short")}'25`;
    const file = DriveUtil.getFiles(nextMonthTimetable, fileName)
    return file;
  }
}









class SummaryBuilder {


  static process(sheet, date) {
    // 1) Determine next‐month year & month
    const nextMonthNum = DateUtil.getNextMonthNumber(date);
    const nextYearFull = DateUtil.getNextYear(date);
    // const nextMonthNum = this.currentMonthNum === 12 ? 1 : this.currentMonthNum + 1;
    // const nextYearFull = this.currentMonthNum === 12
    //   ? this.currentYear + 1
    //   : this.currentYear;
    console.log(nextMonthNum, nextYearFull)

    // 2) Compute days in next month
    const daysInMonth = new Date(nextYearFull, nextMonthNum, 0).getDate();

    // 3) Build the dd-MMM-yy strings for row 3
    const tz = Session.getScriptTimeZone();
    const datesRow = [];
    for (let d = 1; d <= daysInMonth; d++) {
      const dt = new Date(nextYearFull, nextMonthNum - 1, d);
      datesRow.push( [Utilities.formatDate(dt, tz, "dd-MMM-yy")] );
    }
    
    // Find rows where column 1 contains "Sr. No."
    const lastRow = sheet.getLastRow();
    const range = sheet.getRange(1, 1, lastRow, 1);
    const values = range.getValues();

    for (let row = 0; row < values.length; row++) {
        if (values[row][0] === "Sr. No.") {
            // Update dates for columns 3 to 34 (adjust if needed based on your sheet structure)
            sheet.getRange(row + 1, 3, 1, datesRow.length).setValues([ datesRow ]);
        }
    }
  }

}



/**
 * Applies formatting across all sheets.
 */
class Formatting {
  static format(ss, font) {
    // ——— DELETE ALL NON-ESSENTIAL SHEETS ———
    const keep = new Set([
      'Tutor_Names',
      'Online_Wk_1',
      'Online_Wk_2',
      'Online_Wk_3',
      'Online_Wk_4',
      'Online_Wk_5',
      'Online_Wk_6',
      'Extended',
      'Summary'
    ]);
    ss.getSheets().forEach(sh => {
      if (!keep.has(sh.getName())) {
        ss.deleteSheet(sh);
      }
    });
    ss.getSheets().forEach(sh => sh.getDataRange().setFontFamily(font));
  }

}


/**
 * Compares and writes additions/deletions for Tutor_Names sheet.
 */
class NameDiff {
  static process(sheet, empsFromBackend) {
    const lastRow = sheet.getLastRow();
    const existingAll = sheet
      .getRange(2, 2, lastRow - 1, 1)
      .getValues()
      .flat()
      .map(n => n.toString().trim())
      .filter(n => n);
    
    const idx = existingAll.indexOf("X");
    const existing = existingAll.slice(0, idx)
    const add = empsFromBackend.filter(n => !existing.includes(n));
    const del = existing.filter(n => !empsFromBackend.includes(n));

    // clear old diff cols
    sheet.getRange(2, 8, lastRow - 1, 2).clearContent();
    if (add.length) sheet.getRange(2, 8, add.length, 1).setValues(add.map(n => [n]));
    if (del.length) sheet.getRange(2, 9, del.length, 1).setValues(del.map(n => [n]));

    return existingAll;
  }
}






/**
 * Fetches the department, client, spreadsheet type from current spreadsheet
 */
class ActiveSpreadsheet {

  constructor() {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.sheet = this.ss.getActiveSheet();
    /**
     * Gets the value for a specific range 
    */
    this.fetchValue = (range) => this.sheet.getRange(range).getValue().toString().trim().toLowerCase();
  }
}


