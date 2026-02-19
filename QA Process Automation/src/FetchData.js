//fetch the employee details to smeDB sheet for unique ids
function fetchEmployeeData() {

  const empSheetObj = CentralLibrary.DataAndHeaders(EMPLOYEE_DETAIL_SHEET_ID);
  const smeSheetObj = CentralLibrary.DataAndHeaders(MASTER_DB_SPREADSHEET_ID);

  empSheetObj.getSheetById(EMPLOYEE_TAB_ID);
  smeSheetObj.getSheetById(SME_DB_TAB_ID);

  const [empHeaders, empData] = empSheetObj.getDataIndicesFromSheet();
  const [smeHeaders, smeData] = smeSheetObj.getDataIndicesFromSheet();

  const newRows = empData.map(row => {
    return {
      [smeHeaders["Unique ID"]]: row[empHeaders["Unique ID"]],
      [smeHeaders["Email ID"]]: row[empHeaders["Official Email ID"]],
      [smeHeaders["Department"]]: row[empHeaders["Department"]],
      //[smeHeaders["SME Name"]]: row[empHeaders["Employee Name"]],
      [smeHeaders["SME Name"]]: row[empHeaders["Employee Identifier"]],
      [smeHeaders["Grade"]]: row[empHeaders["Grade"]],
      [smeHeaders["Designation"]]: row[empHeaders["Designation"]],
      [smeHeaders["Reporting Manager"]]: row[empHeaders["Reporting Manager"]],
      [smeHeaders["Active?"]]: "",
      [smeHeaders["Added Date"]]: "",
      [smeHeaders["QA Reviewer"]]: "",
      [smeHeaders["QA Reviewer 2"]]: "",
      [smeHeaders["Removed Date"]]: "",
      [smeHeaders["Sheet creation date"]]: "",
      [smeHeaders["SME Sheet Link"]]: ""
    };
  });
  //console.log(newRows);

  const smeHeaderKeys = Object.keys(smeHeaders).sort((a, b) => smeHeaders[a] - smeHeaders[b]);
  const finalRows = newRows.map(obj => smeHeaderKeys.map(header => obj[smeHeaders[header]] ?? ""));

  if (finalRows.length > 0) {
    const smeSheet = smeSheetObj.sheet;
    smeSheet.getRange(2, 1, finalRows.length, finalRows[0].length).setValues(finalRows);
    smeSheet.getRange(2, smeHeaders["Active?"] + 1, finalRows.length, 1).insertCheckboxes().setValue(false);
  } else {
    Logger.log("No data to write.");
  }
}
