let NUM_DAYS_TRAINING = 31

function fetch() {

  // Call the class and its methods
  const training = new Training();
  const employeeDetails = training.getEmployeeDetails();
  training.setData_TrainingTrackerSheet(employeeDetails)
}


class Training {

  constructor () {
    this.EMPLOYEE_DETAILS_SPREADSHEET = SpreadsheetApp.openById("11b5A_ZDi9DskTiKJ0j4fBDkRbY-0I2h7rVRxr6KVK1o");
    this.TRAINING_FEEDBACK_SPREADSHEET = SpreadsheetApp.openById("1o4yXPRS4ZMVabk5UGQgD7GBi4p9odUpAnZ08zs9bEWU");
    this.CURRENT_SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
    this.TRAINING_TRACKER_SHEET = Training.getSheetById(this.CURRENT_SPREADSHEET, 0)
    this.BACKEND_SHEET = Training.getSheetById(this.CURRENT_SPREADSHEET, 1240090851);
    this.START_DATE = new Date(2024, 7, 1);
    this.TODAY = new Date();
  }
  
  static getSheetById (spreadsheet, sheetId) {
    return spreadsheet.getSheets().find(sheet => sheet.getSheetId() === sheetId);
  }


  setData_TrainingTrackerSheet(employeeDetails) {
    const feedbackEmailAddresses = this.checkForFeedBack();

    let [headers_TrainingTracker, data_TrainingTracker] = this.getDataIndicesFromSheetWithStartingRow(this.TRAINING_TRACKER_SHEET, 3);
    let [headers_Backend, data_Backend] = this.getDataIndicesFromSheet(this.BACKEND_SHEET);
   const backendInfo = {};


    backendInfo[headers_TrainingTracker["Reasons"]] = data_Backend.map(row => {
      return row[headers_Backend["Reasons"]];
    }).filter(Boolean);

    backendInfo[headers_TrainingTracker["Status"]] = data_Backend.map(row => {
      return row[headers_Backend["Status"]];
    }).filter(Boolean);
    
    const trainingTrackerUniqueIds = [... new Set(data_TrainingTracker.map(row => row[headers_TrainingTracker["Unique ID"]]))];


    employeeDetails.forEach((employee) => {
      
      if (!trainingTrackerUniqueIds.includes(employee.UniqueID)) {
        // Logger.log(`Missing Unique ID in training tracker: ${employee.UniqueID}`);
        const newRow = [];
        const {login, password} = loginAndPassword(employee.EmployeeName);
        newRow[headers_TrainingTracker["Unique ID"]] = employee.UniqueID;
        newRow[headers_TrainingTracker["Mon_YY"]] = employee.MonYY;
        newRow[headers_TrainingTracker["Department"]] = employee.Department;
        newRow[headers_TrainingTracker["Employee Name"]] = employee.EmployeeName;
        newRow[headers_TrainingTracker["Email ID"]] = employee.EmailID;
        newRow[headers_TrainingTracker["Personal Contact Number"]] = employee.PhoneNumber
        newRow[headers_TrainingTracker["Location"]] = employee.Location;
        newRow[headers_TrainingTracker["Gender"]] = employee.Gender;
        newRow[headers_TrainingTracker["Moodle Login ID"]] = login;
        newRow[headers_TrainingTracker["Password"]] = password;
        const statusValues = backendInfo[headers_TrainingTracker["Status"]];
        const reasonsValues = backendInfo[headers_TrainingTracker["Reasons"]];

        const lastRowIndex = this.TRAINING_TRACKER_SHEET.appendRow(newRow).getLastRow();
        // Create dropdown lists for 'Status' and 'Reasons'
        const statusCell = this.TRAINING_TRACKER_SHEET.getRange(lastRowIndex, headers_TrainingTracker["Status"] + 1); 
        const reasonsCell = this.TRAINING_TRACKER_SHEET.getRange(lastRowIndex, headers_TrainingTracker["Reasons"] + 1); 
        // Set data validation for 'Status'
        const statusRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(statusValues, true) // 'true' for showing dropdown arrow
          .setAllowInvalid(false)
          .build();
        statusCell.setDataValidation(statusRule).setValue("Initiation");

        // Set data validation for 'Reasons'
        const reasonsRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(reasonsValues, true) // 'true' for showing dropdown arrow
          .setAllowInvalid(false)
          .build();
        reasonsCell.setDataValidation(reasonsRule);

        this.TRAINING_TRACKER_SHEET.getRange(lastRowIndex, headers_TrainingTracker["Training Email"] + 1).insertCheckboxes();
        this.TRAINING_TRACKER_SHEET.getRange(lastRowIndex, headers_TrainingTracker["Moodle Credentials Email"] + 1).insertCheckboxes();


        const dateValidationRule = SpreadsheetApp.newDataValidation()
            .requireDate()
            .build();

        const timeValidationRule = SpreadsheetApp.newDataValidation()
          .requireFormulaSatisfied('=AND(ISNUMBER(A1), TEXT(A1, "HH:MM") = A1)')
          .build();

        // Add the current date
        const startDateCell = this.TRAINING_TRACKER_SHEET.getRange(lastRowIndex, headers_TrainingTracker["Training Start Date"] + 1);
        startDateCell.setDataValidation(dateValidationRule).setNumberFormat("dd-MMM-yyyy");

        const endDateCell = this.TRAINING_TRACKER_SHEET.getRange(lastRowIndex, headers_TrainingTracker["Training End Date"] + 1);
        endDateCell.setDataValidation(dateValidationRule).setNumberFormat("dd-MMM-yyyy");

      } else {
        const index = trainingTrackerUniqueIds.indexOf(employee.UniqueID);
        const currentRow = index + 5;
        // Number of days from the date of joining (if only end date is not mentioned),
        const trainingStartDate = this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Training Start Date"] + 1).getValue();
        const trainingEndDate   = this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Training End Date"] + 1).getValue();
        const leaves   = this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Leaves"] + 1).getValue();
        const emailSentBoolean = this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Email Sent?"] + 1).getValue();

        if (this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Department"] + 1).getValue() === "Statistic")
          NUM_DAYS_TRAINING = 61;

        let numOfDays;
        if (trainingStartDate !== "" && trainingStartDate !== undefined)
        {
          if (trainingEndDate === "" && leaves === "") {
              
              numOfDays = CentralLibrary.getDaysDifference(this.TODAY, new Date(trainingStartDate));
              // console.log(employee, numOfDays, trainingStartDate, currentRow, index)
              // console.log(this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Training Duration in Days"] + 1).getValue())
              this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Training Duration in Days"] + 1).setValue(numOfDays);
              if (numOfDays > NUM_DAYS_TRAINING) {
                this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Training Duration in Days"] + 1).setBackground("#e06666");
              } else{
                this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Training Duration in Days"] + 1).setBackground("#ffffff");
              }

          } else if ((trainingEndDate !== "" && trainingEndDate !== undefined) && (leaves === "" || leaves === undefined)) {
              numOfDays = CentralLibrary.getDaysDifference(new Date(trainingEndDate), new Date(trainingStartDate));
              this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Training Duration in Days"] + 1).setValue(numOfDays);
              if (numOfDays > NUM_DAYS_TRAINING) {
                // this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Training Duration in Days"] + 1).setValue("Extended");
                this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Training Duration in Days"] + 1).setBackground("#e06666");
              } else{
                this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Training Duration in Days"] + 1).setBackground("#ffffff");
                // this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Training Duration in Days"] + 1).setValue();
              }

          } else if ((trainingEndDate !== "" && trainingEndDate !== undefined) && (leaves !== "" || leaves !== undefined)) {
              numOfDays = CentralLibrary.getDaysDifference(new Date(trainingEndDate), new Date(trainingStartDate))
              numOfDays -= leaves;
              this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Training Duration in Days"] + 1).setValue(numOfDays);
              if (numOfDays > NUM_DAYS_TRAINING) {
                this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Training Duration in Days"] + 1).setBackground("#e06666");
              } else{
                this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Training Duration in Days"] + 1).setBackground("#ffffff");
              }

          }else if ((trainingEndDate === "") && (leaves !== "" || leaves !== undefined)) {
              numOfDays = CentralLibrary.getDaysDifference(this.TODAY, new Date(trainingStartDate))
              numOfDays -= leaves;
              this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Training Duration in Days"] + 1).setValue(numOfDays);
              if (numOfDays > NUM_DAYS_TRAINING) {
                this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Training Duration in Days"] + 1).setBackground("#e06666");
              } else{
                this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Training Duration in Days"] + 1).setBackground("#ffffff");
              }
          }
        }

        // check for feedback received and change the status of the candidates whose feedback has been received
        if (feedbackEmailAddresses.includes(employee.EmailID)) {
          this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Feedback Received"] + 1).setValue("Y");
        } else {
          this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Feedback Received"] + 1).setValue("");
        }

        const traineeStatus = this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Status"] + 1).getValue();
        // Change the current status each candidate is at the moment.
        if (traineeStatus !== "Extended" && traineeStatus !== "Discontinued" && traineeStatus !== "Left"){
          if(emailSentBoolean.toLowerCase() === "y") {
            // console.log(trainingEndDate);
            this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Status"] + 1).setValue("In Progress");
            if(trainingEndDate !== "" && numOfDays > NUM_DAYS_TRAINING)
              this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Status"] + 1).setValue("Extended");
            else if(trainingEndDate === "" && numOfDays <= NUM_DAYS_TRAINING)
              this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Status"] + 1).setValue("In Progress");
            else if(trainingEndDate === "" && numOfDays > NUM_DAYS_TRAINING)
              this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Status"] + 1).setValue("Extended");
            else if(trainingEndDate !== "" && numOfDays <= NUM_DAYS_TRAINING)
              this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Status"] + 1).setValue("Completed");

          } else if (emailSentBoolean === "") {

            if(trainingEndDate !== "" && numOfDays <= NUM_DAYS_TRAINING)
              this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Status"] + 1).setValue("Completed");
            else if(trainingEndDate !== "" && numOfDays > NUM_DAYS_TRAINING)
              this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Status"] + 1).setValue("Extended");
            else if(trainingEndDate === "" && numOfDays <= NUM_DAYS_TRAINING)
              this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Status"] + 1).setValue("Completed");
            else if(trainingEndDate === "" && numOfDays > NUM_DAYS_TRAINING)
              this.TRAINING_TRACKER_SHEET.getRange(currentRow, headers_TrainingTracker["Status"] + 1).setValue("Extended");
          }
        }

      } // first else block ends
    });
    
  }

  checkForFeedBack() {
    const sheet = this.TRAINING_FEEDBACK_SPREADSHEET.getSheetByName("Form Responses 1");
    const [headers, data] = this.getDataIndicesFromSheet(sheet);
    const emailAdresses = data.map(oneRow => oneRow[headers["Email Address"]]);
    return emailAdresses;
  }


  getEmployeeDetails() {
    const sheet = this.EMPLOYEE_DETAILS_SPREADSHEET.getSheetByName("Employee Info");
    const [headers, data] = this.getDataIndicesFromSheet(sheet);



    const activeEmployees = data.filter(row => row[headers["Function"]] !== "Enabling" && CentralLibrary.getDaysDifference(row[headers["DOJ"]], this.START_DATE) >= 0);
    // activeEmployees.forEach(row=> Logger.log(row[headers["DOJ"]].getYear()));
    const trainingData = activeEmployees.map(row => {
      return {
        UniqueID: row[headers["Unique ID"]],
        MonYY: `${CentralLibrary.monthNumToMonthName(new Date(row[headers["DOJ"]]).getMonth()+1)}_${new Date(row[headers["DOJ"]]).getFullYear()}`,
        Department: row[headers["Department"]],
        EmployeeName: row[headers["Employee Name"]],
        EmailID: row[headers["Official Email ID"]],
        Gender: row[headers["Gender"]],
        PhoneNumber : row[headers["Phone Number"]],
        Location : row[headers["Location"]]
      };
    });
    return trainingData;
  }

  getDataIndicesFromSheetWithStartingRow(sheet, startRow) {
    let dataRange = sheet.getDataRange().getValues();
    dataRange = dataRange.slice(startRow);
    const headers = dataRange[0];
    const data = dataRange.slice(1);
    return [this.createIndexMap(headers), data];
  }

  getDataIndicesFromSheet(sheet) {
    const dataRange = sheet.getDataRange().getValues();
    const headers = dataRange[0];
    const data = dataRange.slice(1);
    return [this.createIndexMap(headers), data];
  }

  createIndexMap(headers) {
    return headers.reduce((map, val, index) => {
      if (val !== "") {
        map[val.trim()] = index;
      }
      return map;
    }, {});
  }

}













