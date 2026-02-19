function myFunction() {
  const masterDbSheet = SpreadsheetApp.openById("11b5A_ZDi9DskTiKJ0j4fBDkRbY-0I2h7rVRxr6KVK1o").getSheetByName("Employee Info");
  const [masterHeaders, masterData] = CentralLibrary.get_Data_Indices_From_Sheet(masterDbSheet);

  const trainingPlanSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Training Tracker");
  const [trainingHeaders, trainingData] = CentralLibrary.get_Data_Indices_From_Sheet(trainingPlanSheet, 3);
  
  masterData.map(row => row[masterHeaders["Unique ID"]], row[masterHeaders[""]])
}
