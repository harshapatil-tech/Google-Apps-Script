// class SpreadsheetManager {

//   constructor(dept, employeeMap, client, ssType) {
//     this.dept = dept, this.client = client, this.ssType = ssType;
//     this.employees = Helper.getEmployeesByDept(employeeMap, dept);
//     this.rootFolder = DriveApp.getFolderById(GOOGLE_DRIVE_FOLDER_ID);
//     this.year = DateUtil.currentYear();
//     this.monthKey = DateUtil.currentMonthKey();
//     this.nextMonthKey = DateUtil.nextMonthKey();
//     console.log(this.year, this.monthKey, this.nextMonthKey);
//   }

//   runTimetable () {
//     const deptFolder  = DriveUtil.getChildFolder(this.rootFolder, this.dept);
//     const yearFolder  = DriveUtil.getChildFolder(deptFolder, this.year);
//     const monthFolder = DriveUtil.getChildFolder(yearFolder, this.monthKey.toLowerCase())
//     const nextMonthFolder = DriveUtil.getOrCreateFolder(yearFolder, this.nextMonthKey);
//     const currMonthFileName = `Time table_${this.client.toUpperCase()}_${Helper.capitalizeFirstLetter(this.dept)}_${DateUtil.currentMonthName("short")}'25`;
//     const currentMonthFile = DriveUtil.getFiles(monthFolder, currMonthFileName);
//     if (currentMonthFile === undefined) {
//       SpreadsheetApp.getUi().alert(`Please check the name of the time table of the current month. New file not created for ${Helper.capitalizeFirstLetter(this.dept)}`);
//       return;
//     }
//     const newFileName = `Time table_${this.client.toUpperCase()}_${Helper.capitalizeFirstLetter(this.dept)}_${DateUtil.nextMonthName("short")}'25`;
//     const copy = DriveUtil.getOrCreateFile(nextMonthFolder, currentMonthFile, newFileName);
//     if (copy.action === "exists") {
//     // SpreadsheetApp.getUi().alert("Timetable for the client, department and month already exists");
//       return 0;
//     } 
//     Spreadsheet_Processor.process(copy.file.getId(), this.employees);
//     return 1;
//   }
// }
