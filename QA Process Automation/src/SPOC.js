// function onEdit(e) {
//   // Set the name of your master database sheet
//   var masterSheetName = "1eJ-7nKeICwYXEBh4oEzLwC6SpkYzCw5ntYxarbD2jxA";
  
//   // Set the name of the destination sheet (e.g., "MathsSheet")
//   var destinationSheetName = "1jxjxLsycvImu3a5QC2V88kXF7Txgne-JpHj-yq41jkk";
  
//   // Set the column number for the department (assuming it's column B)
//   var departmentColumn = 2;
  
//   // Get the edited range
//   var editedRange = e.range;
//   Logger.log(editedRange)
//   // Get the active sheet
//   var activeSheet = editedRange.getSheet();
  
//   // Check if the edit occurred in the master sheet and in the specified column
//   if (activeSheet.getName() == "QA DB" && editedRange.getColumn() == departmentColumn) {
//     // Get the data from the edited row
//     var rowData = activeSheet.getRange(editedRange.getRow(), 1, 1, activeSheet.getLastColumn()).getValues()[0];
    
//     // Check if the department is "Maths"
//     if (rowData[departmentColumn - 1].toLowerCase() === "mathematics") {
//       // Get the destination sheet
//       var destinationSheet = SpreadsheetApp.openById("1jxjxLsycvImu3a5QC2V88kXF7Txgne-JpHj-yq41jkk").getSheetByName("Sheet1");
//       Logger.log(destinationSheet)
//       // Find the row in the destination sheet that corresponds to the edited row in the master sheet
//       var destinationRow = findDestinationRow(destinationSheet, rowData);
      
//       // If the row exists in the destination sheet, update it; otherwise, append a new row
//       if (destinationRow !== -1) {
//         destinationSheet.getRange(destinationRow, 1, 1, destinationSheet.getLastColumn()).setValues([rowData]);
//       } else {
//         destinationSheet.appendRow(rowData);
//       }
//     }
//   }
// }

// function findDestinationRow(destinationSheet, rowData) {
//   // Get the values from the destination sheet
//   var destinationData = destinationSheet.getDataRange().getValues();
  
//   // Iterate through the destination sheet to find the row that matches the data
//   for (var i = 0; i < destinationData.length; i++) {
//     if (isEqual(destinationData[i], rowData)) {
//       return i + 1; // Return the row number (add 1 because arrays are 0-indexed, but rows are 1-indexed)
//     }
//   }
  
//   return -1; // Return -1 if the row is not found
// }

// function isEqual(arr1, arr2) {
//   // Check if two arrays are equal
//   return JSON.stringify(arr1) === JSON.stringify(arr2);
// }
