function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Actions')
      .addItem('Archive Data', 'copyData')
      .addItem("Create Departmentwise Sheets", "main")
      .addToUi();
}

function copyData() {
  const SPREADSHEET = SpreadsheetApp.openById("1azmcGWS2os6jdsQXN1bOS6euctPi9DGuGuKbnFrn4UQ");
  const archiveSheet = SPREADSHEET.getSheetByName("Brainfuse Timesheet Archive");
  const scraperSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Summary");

  const [scraperIndices, scraperData] = get_Data_Indices_From_Sheet(scraperSheet);
  const [archiveIndices, archiveData] = get_Data_Indices_From_Sheet(archiveSheet);

  let lastRowIndex = archiveSheet.getLastRow();

  const scriptProperties = PropertiesService.getScriptProperties();
  const lastExecutionDate = scriptProperties.getProperty('lastExecutionDate');

  if (lastExecutionDate) {
    const daysSinceLastExecution = Math.floor((new Date() - new Date(lastExecutionDate)) / (1000 * 60 * 60 * 24));

    if (daysSinceLastExecution <= 1) {
      SpreadsheetApp.getUi().alert("Function already executed in the last 14 days.");
      return;
    } 
  }

  const newRows = scraperData.map(scraperRow => [
    scraperRow[scraperIndices["Department"]],
    scraperRow[scraperIndices["Account No."]],
    scraperRow[scraperIndices["Type"]],
    scraperRow[scraperIndices["Activity Type"]],
    scraperRow[scraperIndices["Start Date"]],
    scraperRow[scraperIndices["Start Time"]],
    scraperRow[scraperIndices["Hours"]],
    scraperRow[scraperIndices["Comments"]]
  ]);

  archiveSheet.getRange(lastRowIndex + 1, archiveIndices["Department"] + 1, newRows.length, newRows[0].length).setValues(newRows);
  
  //save current date time
  scriptProperties.setProperty('lastExecutionDate', new Date().toISOString());
}



