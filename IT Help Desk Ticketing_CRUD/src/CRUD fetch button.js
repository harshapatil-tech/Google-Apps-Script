/**
 * fetching CRUD data.
 */
function getCrudData() {
  const ticketing = new CrudDataManager();
  ticketing.getCRUDDataUpdated();
}

// Defines the CRUD sheet and sets up its headers
class CrudSheet {
  constructor(crudSheet){
    this.crudSheet = crudSheet;   // Reference to the CRUD sheet
    this.headers = CRUD_HEADERS;  // Predefined headers for the CRUD sheet
  }

  // Writes all headers to the CRUD sheet at the specified row (default = 8)
  writeHeaders(startRow){
    CentralLibrary.applyCustomFormatting(this.crudSheet.getRange(startRow, 1, 1, this.headers.length)).setValues([this.headers]);
  }
}


/**
 * Manages data fetching, filtering, and updating of the CRUD sheet.
 */
class CrudDataManager {
  constructor() {
    // Opens the main spreadsheet and master database using their IDs
    this.spreadsheet = SpreadsheetApp.openById("1fL71mTIzO8gw1quS-lgCiLnYyYHRGpESoiksUQTpbHQ");
    this.masterDBSpreadsheet = SpreadsheetApp.openById("1iByitSy5R35cu13rupuppctpzV0X8dTzqDSJJP2ilAk");

    // Get references to different sheets in the spreadsheet
    this.crudSheet = this.spreadsheet.getSheetByName('CRUD');
    this.dropdownSheet = this.spreadsheet.getSheetByName('Backend');
    this.masterDBSheet = this.masterDBSpreadsheet.getSheetByName('Master DB');

    // Initializes the CrudSheet object and writes headers
    const crudSheet = new CrudSheet(this.crudSheet);
    crudSheet.writeHeaders(8);

    // Fetches headers and data from different sheets
    [this.crudHeaders, this.crudData] = CentralLibrary.get_Data_Indices_From_Sheet(this.crudSheet, 8 - 1);
    [this.dropdownHeaders, this.dropdownData] = CentralLibrary.get_Data_Indices_From_Sheet(this.dropdownSheet);
    [this.dbHeaders, this.dbData] = CentralLibrary.get_Data_Indices_From_Sheet(this.masterDBSheet);

    // Precompute the dropdown object (list of possible values for each relevant column)
    this.dropdownObject = this.dropdowns(this.dropdownHeaders, this.dropdownData);
    // console.log(this.dropdownObject);

    // Mapping issue statuses
    this.statusMap = {
      "Open": ["1.1 Open"],
      "Silicon Rental": ["3.1 Inform Silicon Rental", "3.2 Silicon Rental Informed [A]"],
      "Ticket raised": ["2.1 Ticket Raised [A]"],
      "Closed & Final closure": ["4.1 Close Ticket", "4.2 Ticket closed [A]", "6.1 Final closure", "6.2 Final closure [A]"],
      "Re-opened": ["5.1 Re-open ticket", "5.2 Ticket reopened [A]"],
      "All_without Closed": ["1.1 Open", "3.1 Inform Silicon Rental", "3.2 Silicon Rental Informed [A]", "2.1 Ticket Raised [A]", "5.1 Re-open ticket", "5.2 Ticket reopened [A]"],

    };
  }

/**
 * Clears existing data from row 9 onward, fetches new data, updates issue counts, and sets dropdown validations.
 */
  getCRUDDataUpdated(){
     // Clear from row 9 onward
    this.crudSheet
      .getRange(9, 1, this.crudSheet.getLastRow(), this.crudSheet.getLastColumn())
      .clear()
      .clearDataValidations()
      .clearContent()
      .clearFormat();

    //fetch and update the crud sheet
    const updatedData = this.fetchData();
    if(updatedData.length === 0){
      console.log("no data fetched to crud sheet");
      return;
    }else{
      console.log(`fetched ${updatedData.length} rows in crud sheet`);
    }

    this.updateIssueStatusCounts();


    // Set the dropdown validations
    this.setDropdown(this.crudSheet, this.crudHeaders, this.dropdownObject);
  }
  
  // Creates dropdown list from dropdown sheet data
  dropdowns(header, data){
    const dropdownMapper = {};

    dropdownMapper['Issue Status'] = data.map(row => row[header['Issue Status (For CRUD)']]).filter(Boolean);
    dropdownMapper['Estimated Time'] = data.map(row => row[header['Estimated Time']]).filter(Boolean);
    return dropdownMapper;
  } 

  //Applies the dropdown validation rules to each relevant column in the CRUD sheet.
  setDropdown(crudSheet, header, dropdownObject){
    const dropdownColumns = [
      'Issue Status',
      'Estimated Time',
    ];
  
    dropdownColumns.forEach(columnName =>{
      // console.log(columnName, header)
      const colIndex = header[columnName.trim()] ;
      // console.log(columnName)
       if (colIndex !== undefined) {
        const dropdownList = dropdownObject[columnName];
         if (!dropdownList || dropdownList.length === 0) {
        console.log(`No dropdown data available for column: ${columnName}`);
        return;
      }
      const range = crudSheet.getRange(9, colIndex + 1, crudSheet.getLastRow()- 9 + 1, 1);
      // console.log(`Applying dropdown to column: ${columnName}`);
      range.setDataValidation(this.dropdownValidationRule(dropdownObject[columnName]));
    }else {
      console.log(`Column not found in header for: ${columnName}`);
    }
    
  })
  }

  //Creates and returns a data validation rule for a dropdown list.
  dropdownValidationRule(dropdownList) {
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(dropdownList, true)
    .setAllowInvalid(false)
    .build();
  }


  // Fetches data from the master database sheet and filters it based on selected criteria
  fetchData(){
    const siliconRentalCell = this.crudSheet.getRange(2, 5);
    const siliconRentalChecked = siliconRentalCell.getValue();
    const selectedStatus = this.crudSheet.getRange(4, 5).getValue().trim();
    // console.log(selectedStatus);
   
    // const crudData = this.crudSheet.getRange(9, 1, 1, this.crudSheet.getLastColumn()).getValues();
    // const masterDBData = this.masterDBSheet.getRange(2, 1, 1, this.masterDBSheet.getLastColumn()).getValues();
     

    const timestampIndex = this.dbHeaders["Timestamp"];
    const siliconRentalEmailDateIndex = this.dbHeaders["Silicon Rental Email Date"];
    const finalClosureDateIndex = this.dbHeaders["Final Closure Date"];
    const closureEmailDateIndex = this.dbHeaders["Closure Email Date"];
    const issueStatusIndex = this.dbHeaders["Issue Status"]; 

    let startDate = this.crudSheet.getRange(2, 2).getValue();
    startDate = new Date(startDate);
    startDate.setHours(0, 0, 0, 0);
    startDate = startDate.getTime();

    let endDate = this.crudSheet.getRange(4, 2).getValue();
    endDate = new Date(endDate);
    endDate.setHours(0, 0, 0, 0);
    endDate = endDate.getTime();
    
    let filteredData = [];

    if (siliconRentalChecked === true) {
        filteredData = this.dbData.filter(row => {
            let timestamp = new Date(row[timestampIndex]);
            timestamp.setHours(0, 0, 0, 0);
            timestamp = timestamp.getTime();

            return (
                row[siliconRentalEmailDateIndex] !== "" &&
                (row[finalClosureDateIndex] === "" || row[closureEmailDateIndex] === "") &&
                (timestamp >= startDate && timestamp <= endDate)
            );
        });
    } else {
        if (selectedStatus === "Closed & Final closure") {

            filteredData = this.dbData.filter(row => {
                let timestamp = new Date(row[timestampIndex]);
                timestamp.setHours(0, 0, 0, 0);
                timestamp = timestamp.getTime();

                return (
                    (timestamp >= startDate && timestamp <= endDate)
                );
            });
        } else {
            filteredData = this.dbData.filter(row => {
                let timestamp = new Date(row[timestampIndex]);
                timestamp.setHours(0, 0, 0, 0);
                timestamp = timestamp.getTime();
                
                return (
                    (row[finalClosureDateIndex] === "" || row[closureEmailDateIndex] === "") &&
                    (timestamp >= startDate && timestamp <= endDate)
                );
            });
        }
    }

    // if(siliconRentalChecked === true){
    //   filteredData = this.dbData.filter(row =>{
    //     let timestamp = new Date(row[timestampIndex])
    //     timestamp.setHours(0, 0, 0, 0);
    //     timestamp = timestamp.getTime();

    //     return(
    //       row[siliconRentalEmailDateIndex] !=="" &&
    //       (row[finalClosureDateIndex] === "" || row[closureEmailDateIndex] ==="") &&
    //     (timestamp >= startDate && timestamp <= endDate)
    //     );
       
    //   });
    // }

    // else{
    //   filteredData = this.dbData.filter(row =>{
    //     let timestamp = new Date(row[timestampIndex]);
    //     timestamp.setHours(0, 0, 0, 0);
    //     timestamp = timestamp.getTime();
    //     return(
    //     (row[finalClosureDateIndex] === "" || row[closureEmailDateIndex] === "") &&
    //     (timestamp >=startDate && timestamp <=endDate)
    //     )
    //   })
    // }

    // console.log(filteredData);
    
    if(selectedStatus.trim() !== "All"){
     filteredData = this.fetchFilteredData(filteredData,this.statusMap[selectedStatus]) 
    }
    //  console.log("Filtered Data", filteredData);
   
    const formatedData = this.crudMapping(this.dbHeaders, filteredData);
    // const currentIssueStatus = this.fetchFilteredData(siliconRentalChecked, statusMap[selectedStatus]);
    // console.log("current issue status", currentIssueStatus);
    this.updateCrudSheet(formatedData);
    return formatedData;


  }

  // Filters data based on issue status
  fetchFilteredData(data, statusList) {
    const issueStatusIndex = this.dbHeaders["Issue Status"];
    
    if (issueStatusIndex === -1) {
      console.log("Issue Status column not found");
      return [];
    }

    return data.filter((row, index) => {
      return statusList.includes(row[issueStatusIndex]);
    });
  }


  //CRUD ticket counts
   updateIssueStatusCounts() {
    // const data = this.masterDBSheet.getDataRange().getValues();
    // const headers = data[0];

    // index of the issue status column
    const issueStatusIndex = this.dbHeaders["Issue Status"];
    // exit if column is not found
    if (issueStatusIndex === -1) return;

    // Count occurrences of each Issue Status
    let issueStatusCounts = {};
    for (let i = 0; i < this.dbData.length; i++) {
      let status = this.dbData[i][issueStatusIndex];    // get the issue status
      // console.log("Status", status);
      
      // if issue status is not found the default count is 0
      if (status) {
        issueStatusCounts[status] = (issueStatusCounts[status] || 0) + 1;
      }
    }

    // Map counts to the CRUD sheet
    // const crudData = this.crudSheet.getDataRange().getValues();
    const labelCount = this.crudSheet.getRange(1, 8, 4, 1).getValues().flat();
    // console.log("LabelCount", labelCount);
    const labelMapping = {
      '# Open Tickets' : "1.1 Open",
      '# Tickets in "3.1 Inform Silicon Rental"': "3.1 Inform Silicon Rental",
      '# Tickets in "4.1 Close ticket"': "4.1 Close Ticket",
      '# Tickets in "5.1 Re-open ticket"': "5.1 Re-open ticket"
    };
    
    // the crud sheet data to update ticket counts 
    for (let i = 0; i < labelCount.length; i++) {
      const label = labelCount[i]; // Column "H" contains ticket labels
      // console.log("Label", label)
      if (labelMapping[label]) {
        const count = issueStatusCounts[labelMapping[label]] || 0;
        // console.log("Count", count);
        this.crudSheet.getRange(i + 1, 9).setValue(count); // Column I
      }
    }
  }



  //map db data into the expected CRUD layout/structure
  crudMapping(header, data){
    const modifiedArray = [];

    data.forEach(row => {
      const ticketNo = row[header["Ticket No."]];
      const timestamp = row[header["Timestamp"]];
      const emailAddress = row[header["Email Address"]];
      const employeeName = row[header["Employee Name"]];
      const department = row[header["Department"]];
      const contactNumber = row[header["Contact Number"]];
      const location = row[header["Location"]];
      const deviceType = row[header["Device Type"]];
      const detailedAddress = row[header["Detailed Address"]];
      const laptopBrand = row[header["Laptop Brand"]];
      const laptopNumber = row[header["Laptop Number"]];
      const laptopNumberPicture = row[header["Laptop Number Picture"]];
      const issueType = row[header["Issue Type"]];
      const hardwareIssue = row[header["Hardware Issue"]];
      const softwareIssue = row[header["Software Issue"]];
      const issueDescription = row[header["Issue Description"]];
      // const date = row[header["Date"]];
      const date = Utilities.formatDate(row[header["Date"]], Session.getScriptTimeZone(), "dd-MMM-yyyy");
      const time = row[header["Time"]];
      // const time = Utilities.formatDate(row[header["Time"]], Session.getScriptTimeZone(), "HH:mm:ss");
      const itSupportDiagnosis = row[header["IT Support diagnosis"]];
      const estimatedTime = row[header["Estimated Time"]];
      const initialEmailDate = row[header["Initial Email Date"]];
      const raiseTicketDate = row[header["Raise Ticket Date"]];
      const siliconRentalEmailDate = row[header["Silicon Rental Email Date"]];
      const closureEmailDate = row[header["Closure Email Date"]];
      const reopenTicketDate = row[header["Re-open Ticket Date"]];
      const finalClosureDate = row[header["Final Closure Date"]];
      const issueStatus = row[header["Issue Status"]];
      const remarks = row[header["Remarks"]];
      const lastUpdateDate = row[header["Last Update Date"]];

      modifiedArray.push([ticketNo, timestamp, employeeName, contactNumber, deviceType, hardwareIssue, softwareIssue, itSupportDiagnosis, estimatedTime, issueStatus, "", emailAddress, location, date, time, issueType, department, detailedAddress, laptopBrand, laptopNumber, laptopNumberPicture, issueDescription, remarks]);

    });
    return modifiedArray;
  }


  // Updates the CRUD sheet with new data
  updateCrudSheet(data){
    
    if (data.length === 0) {
      SpreadsheetApp.getUi().alert("No data to display");
      return;
    }
    const startRow = 9;
    const range = this.crudSheet.getRange(startRow, 1, data.length, data[0].length);
    range.setValues(data);

    this.crudSheet.getRange(startRow, this.crudHeaders['Timestamp'] + 1, data.length, 1).setNumberFormat("dd-MMM-yyyy");
    this.crudSheet.getRange(startRow, this.crudHeaders['Date'] + 1, data.length, 1).setNumberFormat("dd-MMM-yyyy");
    this.crudSheet.getRange(startRow, this.crudHeaders['Time'] + 1, data.length, 1).setNumberFormat("hh:mm AM/PM");
    // Insert checkboxes in the 'Update?' column
    const updateColumnIndex = this.crudHeaders["Update?"];
    if (updateColumnIndex !== undefined) {
      const checkboxRange = this.crudSheet.getRange(startRow, updateColumnIndex + 1, data.length, 1);
      checkboxRange.insertCheckboxes();
      // console.log(`Checkboxes inserted in 'Update?' column at column index: ${updateColumnIndex + 1}`);
    } else {
      console.log("'Update?' column not found in headers.");
    }

    // Apply borders to the data range
    range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
    // console.log("Borders applied to the fetched data range.");


  }

}







