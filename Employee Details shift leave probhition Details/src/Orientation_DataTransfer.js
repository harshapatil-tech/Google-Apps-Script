// Something to do with exited employees


function copyData() {

  const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const employeeDetailsSheet = sourceSpreadsheet.getSheetByName("Employee Info");

  const destinationSpreadsheet = SpreadsheetApp.openById("1qMcyOzLzuYk2nlUhY9if2JX9XvV1_qlgfkmzihe5eXk");
  const destinationSheet = destinationSpreadsheet.getSheetByName("Input Sheet");

  const dataRange = employeeDetailsSheet.getDataRange().getValues();
  const headerIndices = createIndexMap(dataRange[0]);
  const data = dataRange.slice(1);

  const sourceEmails = data.filter(r => r[headerIndices["Status"]] === "Active").map(r => r[headerIndices["Official Email ID"]] ).filter(r=>r!=='');

  const startDate = new Date("2023-07-01");
  startDate.setHours(0, 0, 0, 0);


  const destinationDataRange = destinationSheet.getDataRange().getValues();
  const destinationHeaders = destinationDataRange[0], destinationData = destinationDataRange.slice(1);
  const destinationHeaderIndices = createIndexMap(destinationHeaders);
  console.log("Destination Header Indices:", destinationHeaderIndices);


  const dateValidation = SpreadsheetApp.newDataValidation()
                                      .requireDate()
                                      .setAllowInvalid(false)
                                      .build();


  const destinationAllEmails = destinationData.map(r => r[destinationHeaderIndices["Official Email ID"]]).filter(r=>r!=='');
  


  destinationAllEmails.forEach( (email, idx) => {
    if(!sourceEmails.includes(email)) {
      destinationSheet.getRange(idx+2, destinationHeaderIndices["Exited Employee"]+1).setValue(true)
    }
    else{
      destinationSheet.getRange(idx+2, destinationHeaderIndices["Exited Employee"]+1).setValue(false)
    }
  })


  const dataToWrite = [];

  data.forEach(row => {

    const companyDomain = row[headerIndices["Official Email ID"]].split("@")[1];
    
    if (row[headerIndices["DOJ"]] >= startDate && companyDomain === "upthink.com") {
    
      const innerArray = [];
      innerArray.push(row[headerIndices["Employee Name"]]);
      innerArray.push(row[headerIndices["Grade"]]);
      innerArray.push(row[headerIndices["Designation"]]);
      innerArray.push(row[headerIndices["Department"]]);
      innerArray.push(row[headerIndices["DOJ"]]);
      innerArray.push(row[headerIndices["Official Email ID"]]);

      dataToWrite.push(innerArray);
    
    }

  });


  let lastRow = destinationSheet.getLastRow();

  dataToWrite.forEach(r => {
  
    if(!destinationAllEmails.includes(r[5])) {

      destinationSheet.getRange(lastRow+1, destinationHeaderIndices["Employee Name"]+1).setValue(r[0]);
      destinationSheet.getRange(lastRow+1, destinationHeaderIndices["Grade"]+1).setValue(r[1]);
      destinationSheet.getRange(lastRow+1, destinationHeaderIndices["Designation"]+1).setValue(r[2]);
      destinationSheet.getRange(lastRow+1, destinationHeaderIndices["Department"]+1).setValue(r[3]);
      destinationSheet.getRange(lastRow+1, destinationHeaderIndices["DOJ"]+1).setValue(r[4]);
      destinationSheet.getRange(lastRow+1, destinationHeaderIndices["Official Email ID"]+1).setValue(r[5]);
      // destinationSheet.getRange(lastRow+1, destinationHeaderIndices["DOJ"]+1).setDataValidation(dateValidation);
      destinationSheet.getRange(lastRow+1, destinationHeaderIndices["Form 1 Emailed"]+1).insertCheckboxes();
      destinationSheet.getRange(lastRow+1, destinationHeaderIndices["Form 2 Emailed"]+1).insertCheckboxes();
      destinationSheet.getRange(lastRow+1, destinationHeaderIndices["Form 3 Emailed"]+1).insertCheckboxes();
      destinationSheet.getRange(lastRow+1, destinationHeaderIndices["Feedback 1 Received"]+1).insertCheckboxes();
      destinationSheet.getRange(lastRow+1, destinationHeaderIndices["Feedback 2 Received"]+1).insertCheckboxes();
      destinationSheet.getRange(lastRow+1, destinationHeaderIndices["Feedback 3 Received"]+1).insertCheckboxes();
      destinationSheet.getRange(lastRow+1, destinationHeaderIndices["Exited Employee"]+1).insertCheckboxes()
      destinationSheet.getRange(lastRow+1, destinationHeaderIndices["Meet 1"]+1).insertCheckboxes()
      destinationSheet.getRange(lastRow+1, destinationHeaderIndices["Meet 2"]+1).insertCheckboxes()
      destinationSheet.getRange(lastRow+1, destinationHeaderIndices["Meet 3"]+1).insertCheckboxes()
      lastRow += 1;
    }
  });

  applyCustomFormatting(destinationSheet.getRange(2, 1, destinationSheet.getLastRow()-1, destinationSheet.getLastColumn()));

  

}




function createIndexMap(headers) {
  return headers.reduce((map, val, index) => {
    map[val] = index;
    return map;
  }, {});
}
