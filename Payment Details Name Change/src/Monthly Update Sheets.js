function runMonthly() {
  const upd = new Monthly_Sheet_Updater();
  const maths = upd.getDataByDept("Statistics");
  console.log(maths)
}



class Monthly_Sheet_Updater {
  constructor () {
    const employeeDetailsSheet = SpreadsheetApp
      .openById("11b5A_ZDi9DskTiKJ0j4fBDkRbY-0I2h7rVRxr6KVK1o")
      .getSheetByName("Employee Info");
    const [headers, data] = CentralLibrary.get_Data_Indices_From_Sheet(employeeDetailsSheet);
    this.employeeData = this._getEmployeeDetails(headers, data);
  }

  _formatData(headers, rawData) {
    // 1) Build a name→identifier map
    const nameToIdentifier = rawData.reduce((map, row) => {
      const nameCol = headers["Employee Name"];
      const idCol   = headers["Employee Identifier"];
      map[row[nameCol]] = row[idCol];
      return map;
    }, {});

    // 2) Replace each row’s “Reporting Manager” with the manager’s identifier
    const mgrCol = headers["Reporting Manager"];
    rawData.forEach(row => {
      const mgrName = row[mgrCol];
      // if we have an identifier for that name, swap it in
      if (mgrName && nameToIdentifier[mgrName]) {
        row[mgrCol] = nameToIdentifier[mgrName];
      }
      // otherwise leave it as-is (or blank if you prefer)
    });

  }


  _getEmployeeDetails(headers, data) {
    this._formatData(headers, data)
    const reqHeaders = ["Employee Identifier", "Reporting Manager", "Designation", "Department", "Days", "Hours"];
    return data.map(row => {
      return reqHeaders.reduce((acc, header) => {
        if (header in headers) {
          let value = row[headers[header]];
          if (header === "Days" && (value === undefined || value === '')) {
            value = 6;
          }
          if (header === "Hours" && (value === undefined || value === '')) {
            value = 42.5;
          }
          if (header === "Department" && (value === "Economics" || value === "Finance" || value === "Accounts")) {
            value = "Business";
          }
          acc.push(value);
        }
        return acc;
      }, []);
    });
  }

  /**
   * Sorts an array of employee-rows so that each manager
   * appears immediately before any rows whose row[1] === manager.
   * @param {Array<Array>} rows — each row is [id, manager, dept, days, hours]
   * @returns {Array<Array>} a new array in manager→reports order
   */
  _sortByHierarchy(data) {
    // 1) build lookup maps
    const idToRow    = {};
    const childrenMap = {};

    data.forEach(r => {
      const id = r[0], manager = r[1];
      idToRow[id] = r;
      // ensure a list even if manager === undefined or ''
      (childrenMap[manager] = childrenMap[manager] || []).push(id);
    });

    // 2) find “roots”: those whose manager isn’t in our set
    const roots = data
      .filter(r => !idToRow[r[1]])
      .map(r => r[0])
      // optional: alphabetical top-level order
      .sort((a, b) => a.localeCompare(b));
    
    const ordered = [];
    function visit(id) {
      ordered.push(idToRow[id]);
      const kids = childrenMap[id] || [];
      // optional: alphabetical order among siblings
      kids.sort((a, b) => a.localeCompare(b))
          .forEach(childId => visit(childId));
    }
    roots.forEach(rootId => visit(rootId));
    return ordered;
  }


  _joinData(data) {
    data = data.filter(r => !r[2].trim().toLowerCase().includes("domain specialist"))
    const seniorSMEs = data.filter(r=> r[2].trim().toLowerCase().includes("senior"))
    const seniorIds = new Set(seniorSMEs.map(r => r[0]));
    const remSMEs = data.filter(r=> !seniorIds.has(r[0]));
    return [...seniorSMEs, ...this._sortByHierarchy(remSMEs)];
  }


  /**
   * Get the rows for one department, already sorted
   * manager→direct reports.
   */
  getDataByDept(department) {
    const deptRows = this.employeeData.filter(r => r[3] === department);
    return this._joinData(deptRows);
  }

  /**
   * Clear the existing 2-column block and write
   * each row as [ identifier, hours ].
   */
  setName(sheet, rows) {
    const startRow = 4;            // where your block begins
    const numRows  = rows.length;

    // 1) clear cols A–C
    sheet.getRange(startRow, 1, sheet.getLastRow(), 2).clearContent();

    // 2) write [ ID, Hours ]
    //    NOTE: I’m assuming rows are [id, mgr, desig, dept, days, hours]
    const toWrite = rows.map(r => [ r[0], r[5] ]);
    sheet.getRange(startRow, 1, numRows, 2).setValues(toWrite);
    sheet.getRange(startRow, 1, numRows, sheet.getLastColumn())
        .setBorder(true, true, true, true, true, true)          // all borders on
        .setHorizontalAlignment("center")                       // center-align
        .setFontFamily("Roboto")                                // Roboto font
        .setFontSize(11);                                       // size 11
    // // 3) build a formula for each row:  = (this row’s B) / (that row’s Days)
    // //    RC[-1]  = column B in same row
    // //    r[4]    = the “Days” value for this row
    // const formulas = rows.map(r => [`=RC[-1]/${r[4]}`]);

    // // 4) write all of column C at once
    // sheet
    //   .getRange(startRow, 3, numRows, 1)
    //   .setFormulasR1C1(formulas);
  }

}
