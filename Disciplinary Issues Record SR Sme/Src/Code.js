function run() {
  const empDetails = EmployeeDetailsForDisciplinary.getEmployeeDetails();
  const DEPARTMENTS = ["Mathematics", "Statistics", "Biology", "Computer Science", "Business", "English", "Chemistry", "Physics"];
  // const DEPARTMENTS = ["Statistics"]
  for (const dept of DEPARTMENTS) {
    new Disciplinary(dept, empDetails[dept]);
  }
}

class Disciplinary {
  
  constructor (dept, deptEmployees) {
    this.backendSheetName = "Employee List";
    // this.
    const rootFolderId = "1nOsaYwgBWvN3tPiRGaBG30ClNhroe6_b";
    const rootFolder = DriveApp.getFolderById(rootFolderId);
    const subjectFolder = this._getChildFolder(rootFolder, dept);
    const files = this._getFilesInFolder(subjectFolder);
    for (const file of files) {
      this.updateSpreadSheet(file, deptEmployees);
    }
  }

  updateSpreadSheet(file, deptEmployees, updateNames=false) {
    const ss = SpreadsheetApp.open(file);
    const backendSheet = this._getOrCreateFile(ss);
    const [ empIdentifierRange, data ] = this._updateBackendSheet(backendSheet, deptEmployees)

    if (ss.getName().trim().toLocaleLowerCase().includes("sr"))
      this._updateMasterSheet(ss.getSheetByName("Issues_Sr. SME"), empIdentifierRange, data);
    else
      this._updateMasterSheet(ss.getSheetByName("Issues_SME"), empIdentifierRange, data);

    if (updateNames) {
      if (ss.getName().trim().toLocaleLowerCase().includes("sr"))
        this.runSometime(ss.getSheetByName("Issues_Sr. SME"), data);
      else
        this.runSometime(ss.getSheetByName("Issues_SME"), data);
    }
  }

  runSometime(sheet, backendData) {
    if (sheet === undefined || sheet === null)
      return;
    
    const backendMap = Object.fromEntries(backendData.map(r => [r[0], r[1]]));

    const [ headers, data ] = CentralLibrary.get_Data_Indices_From_Sheet(sheet);

    data.forEach((row, idx) => {
      const uuid = row[headers["UUID"]];
      const smeName = row[headers["SME Name"]];
      if (uuid !== "") {
        const backendName = backendMap[uuid];
        if (backendName.trim().toLowerCase() !== smeName.trim().toLowerCase()) {
          sheet.getRange(idx+2, headers["SME Name"]+1).setValue(backendName);
        }
      }
    })
  }


  _getOrCreateFile(spreadsheet) {
    const sheets = spreadsheet.getSheets();
    const found = sheets.find(sheet => sheet.getName() === this.backendSheetName);
    if (found === undefined)
      spreadsheet.insertSheet(this.backendSheetName);
    return spreadsheet.getSheetByName(this.backendSheetName);
  }

  _updateMasterSheet(sheet, empIdentifierRange, backendData) {
    if (sheet === undefined || sheet === null)
      return;

    const backendMap = Object.fromEntries(backendData.map(r => [r[1], r[0]]));
    
    const [ headers, data ] = CentralLibrary.get_Data_Indices_From_Sheet(sheet);
    const smeColumnIdx = headers["SME Name"]
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(empIdentifierRange)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, smeColumnIdx+1, sheet.getLastRow()-1, 1).setDataValidation(rule);

    // data.forEach((row, idx) => {
    //   const smeName = row[headers["SME Name"]];
    //   if (!smeName)
    //     return;
    //   if (backendMap[smeName]) {
    //     sheet.getRange(idx+2, headers["UUID"]+1).setValue(backendMap[smeName]);
    //   }
    // })

    // this.protectColumnB(sheet)
  }

  _updateBackendSheet(sheet, deptEmployees) {
    if (sheet === undefined || sheet === null)
      return;
    sheet.getRange(1, 1, 1, 3).setValues([["UUID", "Employee Identifier", "Employee Name"]])
    const [ headers, data ] = CentralLibrary.get_Data_Indices_From_Sheet(sheet);

    //. Get all the previous UUIDs present in the backend sheet
    const uuids = data.map(r => r[headers["UUID"]]);
    const presentEmpIdentifiers = data.map(r => r[headers["Employee Identifier"]]);
  
    const values = [];
    for (const [ uuid, [empIdentifier, empName] ] of Object.entries(deptEmployees)) {
      if (!uuids.includes(uuid)) {
        values.push([uuid, empIdentifier, empName]);
      }
      else {
        if (!presentEmpIdentifiers.includes(empIdentifier)) {
          sheet.getRange(uuids.indexOf(uuid)+2, headers["Employee Identifier"]+1).setValue(empIdentifier);
        }
      }
      
    }
    if (values.length > 0)
      sheet.getRange(sheet.getLastRow()+1, 1, values.length, values[0].length).setValues(values);
    // this._hideAndProtectSheet(sheet)
    return [ sheet.getRange(2, 2, sheet.getMaxRows()-1, 1),  data];
  }


  _hideAndProtectSheet(sheet, editors = ["automation@upthink.com"]) {
    sheet.hideSheet(); // Hide the sheet in UI (not secure)
    const protection = sheet.protect();
    protection.setDescription("Protected by script");
    protection.setWarningOnly(false); // Actually restrict, not just warn
    protection.addEditors(editors); // Set who can edit

    // Remove owner and all other editors except the script owner
    // Optionally, also remove domain editors:
    // protection.removeEditors(protection.getEditors().filter(
    //   editor => !editors.includes(editor.getEmail())
    // ));
  }

  protectColumnB(sheet, editors = ["automation@upthink.com"]) {
    // Get range for column B (entire column)
    const range = sheet.getRange("B:B");

    // Remove previous protections on this range
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (let i = 0; i < protections.length; i++) {
      if (protections[i].getRange().getA1Notation() === "B:B") {
        protections[i].remove();
      }
    }

    sheet.hideColumn(range);
    // Set new protection
    const protection = range.protect().setDescription('Protected Column B');
    protection.setWarningOnly(false); // Enforce protection (not just warning)
    protection.addEditors(editors);   // Allow only these editors
    // Optionally remove yourself from editors (not recommended unless you want to lose access)
    // protection.removeEditors(protection.getEditors().filter(
    //   editor => !editors.includes(editor.getEmail())
    // ));
  }

  

  _getChildFolder(rootFolder, childFolderNomenclature) {
    const childFolders = rootFolder.getFolders();
    while(childFolders.hasNext()) {
      const folder = childFolders.next();
      const folderName = folder.getName();
      if ( folderName.includes(childFolderNomenclature) ) {
        return folder;
      }
    }
  }

  _getFilesInFolder(rootFolder) {
    const spreadsheets = [];
    const files = rootFolder.getFiles();
    while(files.hasNext()) {
      const file = files.next();
      if (file.getName().trim().toLowerCase().includes("disciplinary"))
        spreadsheets.push(file);
    }
    return spreadsheets;
  }

}


class EmployeeDetailsForDisciplinary {

  static getEmployeeDetails() {
    const ss = SpreadsheetApp.openById("11b5A_ZDi9DskTiKJ0j4fBDkRbY-0I2h7rVRxr6KVK1o");
    const empInfoSheet = ss.getSheetByName("Employee Info");
    const [headers,data] = CentralLibrary.get_Data_Indices_From_Sheet(empInfoSheet);
    
    const deptWiseObject = {};

    data.forEach( row => {
      const deptVal = row[headers["Department"]].trim();
      const uuidVal = row[headers["Unique ID"]].trim();
      const array = [ row[headers["Employee Identifier"]], row[headers["Employee Name"]] ];
      if (!deptWiseObject[deptVal]) {
        deptWiseObject[deptVal] = {};
      }
      deptWiseObject[deptVal][uuidVal] = array;
    })

    return deptWiseObject;
  }
}