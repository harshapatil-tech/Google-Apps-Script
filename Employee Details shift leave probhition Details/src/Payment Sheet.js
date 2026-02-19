function runPaymentDetails() {
  const employees = new EmployeeSheet();
  // console.log(employees.employeeDetails)
  const paymentSheet = new PaymentSheet();
  paymentSheet.run(employees.employeeDetails);
}

class PaymentSheet {

  constructor () {
    this.today = new Date();
    this.next = new Date(
      this.today.getFullYear(),
      this.today.getMonth() + 1,
      1
    );
    this.ss = SpreadsheetApp.openById("1AmjNZq6ryYkJWZ4rElJYZIVsNg97su_1p9IP4adMk3s");
    this.startRow = 2;
  }


  run(employeeMap) {
    this.createSheetIfNotExist();
    const sheet = this.ss.getSheetByName(this._findMonthKey(this.next));
    const [headers, data] = CentralLibrary.get_Data_Indices_From_Sheet(sheet, this.startRow-1);
    const presentUUID = data.map(row => row[headers["UUID"]]);
    
    // for (const[uuid, arr] of Object.entries(employeeMap)) {
    //   const name = arr[0];
    //   if (presentNames.includes(name)) {
    //     const found = presentNames.findIndex(ele => ele === name) 
    //     sheet.getRange(found+3, headers["UUID"]+1).setValue(uuid);
        
    //   }
    // }

    // 3) Delete rows for employees separated before next month (bottom-up)
    for (let i = data.length - 1; i >= 0; i--) {
      const rowUUID = data[i][headers["UUID"]];
      const emp = employeeMap[rowUUID];
      if (emp) {
        const [_name, _dept, status, dol, _gender] = emp;
        if (status === "Separated" && new Date(dol) < this.next) {
          sheet.deleteRow(i+3)
          // sheet.deleteRow(i + 2); // +2 because data starts at row 2
        }
      }
    }

    // 3) collect any employees that aren’t yet on the sheet
    const toAdd = [];
    for (const [uuid, [name, dept, status, dol, gender]] of Object.entries(employeeMap)) {
      const alreadyThere = presentUUID.includes(uuid);
      const leftBeforeNext = (status === "Separated" && new Date(dol) < this.next);
      if (!alreadyThere && !leftBeforeNext) {
        // build the new row however your sheet expects it—
        // here we’re just doing [Name, Department, Status, DOL]
        toAdd.push([ uuid, name, dept, gender ]);
      }
    }
    
    // 4) append each missing employee
    toAdd.forEach(row => {
      const lastRow = sheet.getLastRow() + 1
      // sheet.getRange(sheet.getLastRow(), 1).copyTo(
      //       sheet.getRange(lastRow, 1),
      //       SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
      //       false
      //     );
      sheet.getRange(lastRow, headers["UUID"]+1).setValue(row[0]);
      sheet.getRange(lastRow, headers["Name"]+1).setValue(row[1]);
      sheet.getRange(lastRow, headers["Subject"]+1).setValue(row[2]);
      sheet.getRange(lastRow, headers["Gender"]+1).setValue(row[3]);
    })

    this._updateSerialNumber(sheet)

  }

  _updateSerialNumber(sheet) {
    const lastRow = sheet.getLastRow();
    const firstDataRow = this.startRow + 1;
    sheet.getRange(firstDataRow, 1).setValue(1);
    // 2) only proceed if there’s at least one data row
    if (lastRow >= firstDataRow) {
      const numRows = lastRow - firstDataRow;         // number of data rows
      
      // 3) write the “previous row + 1” formula into column A of every data row
      //    R1C1 “=R[-1]C+1” means “this cell = value one row above, same column, plus 1”
      sheet
        .getRange(firstDataRow+1, 1, numRows, 1)
        .setFormulaR1C1("=R[-1]C+1");
    }

    
  }

  /**
   * Returns next month’s short name (e.g. "Jun") and two-digit year (e.g. "25")
   */
  _findMonthKey(date) {
    // short month name, e.g. "Jun"
    const shortMonth = date.toLocaleString('en-US', { month: 'short' });
    // two-digit year, e.g. "25"
    const shortYear  = date.toLocaleString('en-US', { year: '2-digit' });
    return `${shortMonth}${shortYear}`;
  }


  /**
   * Create sheet for the next month
  */
  createSheetIfNotExist() {
    const nextMonthKey = this._findMonthKey(this.next);
    const currMonthKey = this._findMonthKey(this.today);
    const sheets = this.ss.getSheets();
    const found = sheets.find(sheet => sheet.getName() === nextMonthKey);
    if (found === undefined) {
      const copy = this.ss.getSheetByName(currMonthKey).copyTo(this.ss);
      copy.setName(nextMonthKey);
    }
  }
}




class EmployeeSheet {

  constructor() {
    const empSheet = SpreadsheetApp.openById("11b5A_ZDi9DskTiKJ0j4fBDkRbY-0I2h7rVRxr6KVK1o").getSheetByName("Employee Info"); 
    const [headers, data] = CentralLibrary.get_Data_Indices_From_Sheet(empSheet);
    this.employeeDetails = this._getEmployeeDetails(headers, data);
  }

  getEmployeesByDept(department) {
    return this.employeeDetails[department];
  }


  _getEmployeeDetails(headers, data) {
    
    // data = data.filter(row => row[headers["Status"]] === "Separated");
    return Object.fromEntries(
      data.map(row=> {
        return [row[headers["Unique ID"]], [
          row[headers["Employee Identifier"]], 
          row[headers["Department"]], 
          row[headers["Status"]], 
          row[headers["DOL"]],
          row[headers["Gender"]]
          ]
        
      ]})
    )


  }
}









