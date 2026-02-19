//Set DropDwon Values in QA_Review_Update sheet 
function qaReviewUpdateSetDropdowns(qaReviewUpdateSheet, backendSheet) {

  //const smeList = [... new Set(backendSheet.getRange(5, 1, backendSheet.getLastRow(), 1).getValues().flat().filter(Boolean)), 'All'];
  const headerRow = backendSheet.getRange(4, 1, 1, backendSheet.getLastColumn()).getValues()[0];
  const smeNameColIndex = headerRow.indexOf("SME Name") + 1; // 1-based index
  const smeList = [...new Set(backendSheet.getRange(5, smeNameColIndex, backendSheet.getLastRow() - 4, 1).getValues().flat().filter(Boolean)), "All"];
  console.log("smeList", smeList)
  var cell = qaReviewUpdateSheet.getRange(1, 4);
  //cell.clearContent();
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(smeList).setAllowInvalid(false).build();
  cell.clearDataValidations().setDataValidation(rule)

  const timeStamp = new Date();
  let startDate = new Date(timeStamp.getTime() - 7 * 24 * 60 * 60 * 1000);
  const endDate = Utilities.formatDate(timeStamp, Session.getScriptTimeZone(), 'dd-MMM-yy');
  startDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'dd-MMM-yy');
  const startDateCell = qaReviewUpdateSheet.getRange(1, 7);
  const endDateCell = qaReviewUpdateSheet.getRange(2, 7);
  startDateCell.clearContent();
  startDateCell.setValue(startDate);
  endDateCell.clearContent();
  endDateCell.setValue(endDate);
}