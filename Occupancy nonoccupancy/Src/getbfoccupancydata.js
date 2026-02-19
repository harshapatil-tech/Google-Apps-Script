// Dont Delete for now
function getBFOccupancyData() {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = spreadSheet.getSheetByName("Summary")
  const [headerRowMap, data] = get_Data_Indices_From_Sheet(summarySheet);
  
  const leastDate = new Date(Math.min.apply(null, data.map(r => new Date(r[headerRowMap["Start Date"]]))));
  console.log("least date",leastDate);
  const maxDate = new Date(Math.max.apply(null, data.map(r => new Date(r[headerRowMap["Start Date"]]))));
  console.log("maxdate:-",maxDate);
  
  const dataObject = {};
  
  data.forEach(singleRow => {
    const rowDate = new Date(singleRow[headerRowMap["Start Date"]]); 
    //console.log("Row Date:", rowDate,  "| Account:", singleRow[headerRowMap["Account No."]], "| Dept:", singleRow[headerRowMap["Department"]]);
      //const accountNum = singleRow[headerRowMap["Account No."]];
     let accountNum = singleRow[headerRowMap["Account No."]];
     accountNum = accountNum.toString().replace(/\D/g, ''); 


      let singleDual = singleRow[headerRowMap["Type"]];
     
      let department = singleRow[headerRowMap["Department"]];
     
      if (department === "Mathematics") {
        department = "Calculus"
      }
      if (department === "Intro Accounting") {
        department = "Accounting"
      }
      let occupancyType = singleRow[headerRowMap["Activity Type"]];
      const hours = singleRow[headerRowMap["Hours"]];

      if (occupancyType === 'IA-Waited' || occupancyType === "IA-Tutored") {

        if (singleDual === "Dual")
          singleDual = "Multiple"

        if (occupancyType === "IA-Tutored")
          occupancyType = "Occupancy"

        if (occupancyType === "IA-Waited")
          occupancyType = "Non-Occupancy"

        if (!dataObject.hasOwnProperty(department)) {

          dataObject[department] = {}
        }
        if (!dataObject[department].hasOwnProperty(singleDual)) {
          dataObject[department][singleDual] = {}
        }

        if (!dataObject[department][singleDual].hasOwnProperty(accountNum)) {
          dataObject[department][singleDual][accountNum] = {}
        }

        if (!dataObject[department][singleDual][accountNum].hasOwnProperty(occupancyType)) {
          dataObject[department][singleDual][accountNum][occupancyType] = 0
        }
        
        dataObject[department][singleDual][accountNum][occupancyType] += hours;
      } 
    });

     //console.log("From:", leastDate, "To:", maxDate);
     //console.log("Final Data Object:", JSON.stringify(dataObject, null, 2));

    return [dataObject, leastDate, maxDate];

}
