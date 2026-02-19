function genNewSpreadsheet () {
  const activeSpreadsheet = new ActiveSpreadsheet();
  const deptValue = activeSpreadsheet.fetchValue("B4");
  const client    = activeSpreadsheet.fetchValue("C4");
  const ssType    = activeSpreadsheet.fetchValue("D4");
  const empMap    = Helper.mapSubjectToEmployees();

  // 1) decide which departments to run
  const departments = deptValue === "all" 
    ? SUBJECTS
    : [ deptValue ];

  const created = [];
  const notCreated = []

  // 2) for each dept, run the right routine
  departments.forEach(rawDept => {
    const dept = rawDept.trim().toLowerCase();
    if (ssType === "timetable") {
      const manager = new SpreadsheetManager(dept, empMap, client, ssType)
      manager.runTimetable();
    }
  });

}



class SpreadsheetManager {

  constructor(dept, employeeMap, client, ssType) {
    this.dept = dept, this.client = client, this.ssType = ssType;
    this.employees = Helper.getEmployeesByDept(employeeMap, dept);
    this.rootFolder = DriveApp.getFolderById(GOOGLE_DRIVE_FOLDER_ID);
    this.date = new Date();
    this.year = DateUtil.getCurrentYear(this.date);
    this.monthKey = DateUtil.getMonthKey(this.date);
    this.nextMonthKey = DateUtil.getNextMonthKey(this.date);
    console.log(this.year, this.monthKey, this.nextMonthKey);
  }

  runTimetable () {
    const deptFolder  = DriveUtil.getChildFolder(this.rootFolder, this.dept);
    const yearFolder  = DriveUtil.getChildFolder(deptFolder, this.year);
    const monthFolder = DriveUtil.getChildFolder(yearFolder, this.monthKey.toLowerCase());
    const nextMonthFolder = DriveUtil.getOrCreateFolder(yearFolder, this.nextMonthKey);
    const shortMonth = new Intl.DateTimeFormat("en-US", { month: "short" }).format(this.date);
    const yy = String(this.date.getFullYear()).slice(-2);
    const currMonthFileName =   `Time table_${this.client.toUpperCase()}_${Helper.capitalizeFirstLetter(this.dept)}_${shortMonth}'${yy}`;
    // const currMonthFileName = `Time table_${this.client.toUpperCase()}_${Helper.capitalizeFirstLetter(this.dept)}_${DateUtil.getMonthKey(this.date, "{MONTH}'{YY}")}`;
    const currentMonthFile = DriveUtil.getFiles(monthFolder, currMonthFileName);
    console.log(monthFolder.getName(), currMonthFileName)

    if (currentMonthFile === undefined) {
      SpreadsheetApp.getUi().alert(`Please check the name of the time table of the current month. New file not created for ${Helper.capitalizeFirstLetter(this.dept)}`);
      return;
    }
    const newFileName = `Time table_${this.client.toUpperCase()}_${Helper.capitalizeFirstLetter(this.dept)}_${DateUtil.getNextMonthKey(this.date, "{MONTH}'{YY}")}`;

    // We need to copy the template file to the next month folder
    const template = DriveApp.getFileById("1k27-jiqtgxxi7kdbofMSehE3HU3LD67_9vrQCNPWKyg");
    const copy = DriveUtil.getOrCreateFile(nextMonthFolder, template, newFileName);
    // if (copy.action === "exists") {
    //   SpreadsheetApp.getUi().alert("Timetable for the client, department and month already exists");
    //   return 0;
    // } else if(copy.action === "copied") {
    //   // If a new file has been created, we have to update the file
    //   // const previousData = new CopyData(currentMonthFile.getId());
    //   // const data = previousData.run();
    //   // const pastData = new PasteData(copy.file.getId(), data);
    //   // pastData.run();
    // }
    const previousData = new CopyData(currentMonthFile.getId());
    const data = previousData.run();
    //console.log(JSON.stringify(data, null, 2));

    const pastData = new PasteData(copy.file.getId(), data);
    pastData.run();

    Spreadsheet_Processor.process(copy.file.getId(), this.employees, this.date);
    return 1;


  }
}





class Spreadsheet_Processor {

  static process(spreadsheetId, employeesFromBackend, date) {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    // 1. Uniform font
    Formatting.format(ss, 'Roboto');
    // 2. Tutor names diff
    const tutorSheet = ss.getSheetByName('Tutor_Names');
    const tutorNames = NameDiff.process(tutorSheet, employeesFromBackend);
    // 3. Online week date fill
    OnlineWkProcessor.process(ss, tutorSheet, date);
    new OnlineWkSummaryUpdater(ss);
    // 4. (optional) Summary sheet updates
    SummaryBuilder.process(ss.getSheetByName('Summary'), date);
    SummaryBuilder.process(ss.getSheetByName('Extended'), date);
    const extendedUpdater = new ExtendedUpdater(ss)
      extendedUpdater.addOrDeleteColumn();
      const summaryUpdater = new SummaryUpdater(ss)
      summaryUpdater.addOrDeleteColumn();
  }
}









