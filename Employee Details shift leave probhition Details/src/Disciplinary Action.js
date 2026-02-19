function runDisciplinaryAction() {
  const DEPARTMENTS = ["Mathematics", "Statistics", "Biology", "Computer Science", "Business", "English", "Chemistry", "Physics"];
  const empDetails = EmployeeDetailsForDisciplinary.getEmployeeDetails();
  const disciplinaryAction = new DisciplinaryAction();
  disciplinaryAction.updateSpreadsheet(SpreadsheetApp.openById("1mtl_umhrm4mH7TLK2FX3d0DcXkmeK6eCTJ70trASBMU"))
}

class DisciplinaryAction {
  
  constructor () {
    const disciplinaryActionSS = SpreadsheetApp.openById("1mtl_umhrm4mH7TLK2FX3d0DcXkmeK6eCTJ70trASBMU");
    this.backendSheetName = "Backend Employee Data"
  }

  updateSpreadsheet(ss) {
    const backendSheet = this._getOrCreateBackendSheet(ss);
    const [ empIdentifierRange, data ] = this._updateBackendSheet(backendSheet, deptEmployees)

    this._updateMasterSheet(ss.getSheetByName("Issues_SME"), empIdentifierRange, data);

    this.runSometime(ss.getSheetByName("Issues_SME"), data);

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
    this._hideAndProtectSheet(sheet)
    return [ sheet.getRange(2, 2, sheet.getMaxRows()-1, 1),  data];
  }

  _getOrCreateBackendSheet(ss) {
    const sheets = ss.getSheets();
    const found = sheets.find(sheet => sheet.getName() === this.backendSheetName);
    if (found === undefined) {
      ss.insertSheet(this.backendSheetName);
    }
    return ss.getSheetByName(this.backendSheetName);
  }
}


// class EmployeeDetailsForDisciplinary {

//   static getEmployeeDetails() {
//     const ss = SpreadsheetApp.openById("11b5A_ZDi9DskTiKJ0j4fBDkRbY-0I2h7rVRxr6KVK1o");
//     const empInfoSheet = ss.getSheetByName("Employee Info");
//     const [headers,data] = CentralLibrary.get_Data_Indices_From_Sheet(empInfoSheet);
    
//     const deptWiseObject = {};

//     data.forEach( row => {
//       const deptVal = row[headers["Department"]].trim();
//       const uuidVal = row[headers["Unique ID"]].trim();
//       const array = [ row[headers["Employee Identifier"]], row[headers["Employee Name"]] ];
//       if (!deptWiseObject[deptVal]) {
//         deptWiseObject[deptVal] = {};
//       }
//       deptWiseObject[deptVal][uuidVal] = array;
//     })

//     return deptWiseObject;
//   }
// }