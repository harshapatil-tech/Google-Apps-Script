function main() {

  const misData = new MIS_DATA();
   const dataObject = misData.getData();
   misData.setData(dataObject);

}


/**
 * Represents data management for MIS (Management Information System) in Google Apps Script.
 * This class provides methods to set, get, and combine MIS data from various sources.
 * @constructor*/

class MIS_DATA {

  /**
   * Constructor for initializing MIS_DATA with default values.
   * Sets up various properties such as sheet IDs, dates, and folder IDs.
  */

  constructor() {

    this.CLIENT_HOURS_SHEET = "1wtjTbww9IM4FcMB6-N-Tv23oDDkFcPBWmePjhzCcglY";
    this.BRAINFUSE_RATINGS_SHEET = "1rVWwjDGkQWeYwQ58EmXyeD9dr5WFV_fhnpc5FJfJEPY";
    this.OCCUPANCY_SHEET = "1WTAwoyzkLtzKOEPqejOVNs7_3xzRMBEHKTumlNwMd-M";
    this.BILLABILITY_SHEET = "1N1SqJg9zG90QlQViNNH3JXZ_Y0ebXMAM1P_X3-QSQV8";//"1rIkKafzr_xE4wK6ZNsqP13N7sOVuwyY8gMKJoYr4a9E"; 
    this.QA_RATINGS_SHEET = "1eJ-7nKeICwYXEBh4oEzLwC6SpkYzCw5ntYxarbD2jxA";
    this.DATE = new Date();                     // new Date(2024, 9, 1);
    this.YEAR = this.DATE.getFullYear();
    this.PREVIOUS_YEAR = this.YEAR;
    this.MONTH = this.DATE.getMonth();
    this.PREVIOUS_MONTH = this.MONTH - 1;
    console.log("Previous month is: ",this.PREVIOUS_MONTH);
    if (this.MONTH === 1) {
      this.PREVIOUS_YEAR -= 1;
      this.MONTH = CentralLibrary.monthNumToMonthName.call(this, 1);
      this.PREVIOUS_MONTH = CentralLibrary.monthNumToMonthName.call(this, 12);
      console.log(this.MONTH, this.PREVIOUS_MONTH);
    }else{
      this.MONTH = CentralLibrary.monthNumToMonthName.call(this, this.MONTH);
      this.PREVIOUS_MONTH = CentralLibrary.monthNumToMonthName.call(this, this.PREVIOUS_MONTH);
      
    }
    this.FIN_YEAR = CentralLibrary.getFinancialYear(this.MONTH, this.YEAR);
    this.PARENT_FOLDER_ID = "16G65r1bwp5bRKB6N3PW2nITm-7QH5F24";
    this.SUBJECT_LIST = [
                          "Mathematics", "Statistics", "Physics", "Chemistry", "Biology", "Intro Accounting", 
                          "Finance", "Economics","Computer Science", "English", "Total"
                        ];

    // this.SUBJECT_LIST = [
    //                       "Total"
    //                     ];


    console.log(this.MONTH, this.FIN_YEAR, this.PREVIOUS_MONTH, this.YEAR)

    // this.labGrading = new LabGrading(this.MONTH, this.YEAR);
  }

  setDataIndividualSheet(fileId, key, values, dataObject, lastColumn=14){
    const spreadsheet = SpreadsheetApp.openById(fileId);
    if(spreadsheet.getName().split("|")[2].trim().toLowerCase() === "total") {
      this.setDepartmentwiseDataForTotalSheet(spreadsheet.getSheetByName("Departmentwise Data"), dataObject, lastColumn)
    }
      
    const allSheets = spreadsheet.getSheets();
    let sheet = ""
    if(allSheets.some(sheet => sheet.getName() === "Sheet1"))
      sheet = spreadsheet.getSheetByName("Sheet1").setName("Backend DB");
    else
      sheet = spreadsheet.getSheetByName("Backend DB");

    if (lastColumn === 0 || lastColumn === 1) {
      // sheet.clear();
    CentralLibrary.applyCustomFormatting.call(this, sheet.getRange(2, 1), {"fontWeight": "bold" }).setValue("Client Hours")
    CentralLibrary.applyCustomFormatting.call(this, sheet.getRange(3, 1), {"fontWeight": "bold" }).setValue("Scheduled Hours")   
    CentralLibrary.applyCustomFormatting.call(this, sheet.getRange(4, 1), {"fontWeight": "bold" }).setValue("Actual Hours")
    CentralLibrary.applyCustomFormatting.call(this, sheet.getRange(7, 1), {"fontWeight": "bold" }).setValue("Occupancy");    
    CentralLibrary.applyCustomFormatting.call(this, sheet.getRange(8, 1), {"fontWeight": "bold" }).setValue("Non-Occupancy");
    CentralLibrary.applyCustomFormatting.call(this, sheet.getRange(11, 1), {"fontWeight": "bold" }).setValue("Positive Ratings");    
    CentralLibrary.applyCustomFormatting.call(this, sheet.getRange(12, 1), {"fontWeight": "bold" }).setValue("Negative Ratings");    
    CentralLibrary.applyCustomFormatting.call(this, sheet.getRange(13, 1), {"fontWeight": "bold" }).setValue("Total Students");    
    CentralLibrary.applyCustomFormatting.call(this, sheet.getRange(16, 1), {"fontWeight": "bold" }).setValue("Billable Hours");   
    CentralLibrary.applyCustomFormatting.call(this, sheet.getRange(17, 1), {"fontWeight": "bold" }).setValue("Non-Billabale Hours");    
    CentralLibrary.applyCustomFormatting.call(this, sheet.getRange(20, 1), {"fontWeight": "bold" }).setValue("Leakage Percentage");    
    CentralLibrary.applyCustomFormatting.call(this, sheet.getRange(21, 1), {"fontWeight": "bold" }).setValue("Occupancy Percentage");    
    CentralLibrary.applyCustomFormatting.call(this, sheet.getRange(22, 1), {"fontWeight": "bold" }).setValue("Negative Ratings Percentage");    
    CentralLibrary.applyCustomFormatting.call(this, sheet.getRange(23, 1), {"fontWeight": "bold" }).setValue("Billability Percentage");
    CentralLibrary.applyCustomFormatting.call(this, sheet.getRange(24, 1), {"fontWeight": "bold" }).setValue("QA Ratings Percentage");
    sheet.setColumnWidth(1, 200);
    lastColumn = 2
    }else {
      sheet.deleteColumn(2);
    }
    
    CentralLibrary.applyCustomFormatting.call(this, sheet.getRange(1, lastColumn), {"bgColor": "#6fb43f", "fontWeight": "bold" })
                  .setNumberFormat("MMM-YYYY").setValue(key);
    sheet.getRange(2, lastColumn).setValue(values["Client Received Hours"])
    sheet.getRange(3, lastColumn).setValue(values["Scheduled Hours"])
    sheet.getRange(4, lastColumn).setValue(values["Actual Delivered Hours"]);
    sheet.getRange(7, lastColumn).setValue(values["Occupancy"]);
    sheet.getRange(8, lastColumn).setValue(values["Non-Occupancy"]);
    sheet.getRange(11, lastColumn).setValue(values["Positive Ratings"]);
    sheet.getRange(12, lastColumn).setValue(values["Negative Ratings"]);
    sheet.getRange(13, lastColumn).setValue(values["Total Students"]);
    sheet.getRange(16, lastColumn).setValue(values["Billable Hours"]);
    sheet.getRange(17, lastColumn).setValue(values["Non-Billabale Hours"]);
    // sheet.getRange(16, lastColumn).setValue("");
    // sheet.getRange(17, lastColumn).setValue("");
    sheet.getRange(20, lastColumn).setNumberFormat("0.00%").setValue(values["Leakage Percentage"] + "%");
    sheet.getRange(21, lastColumn).setNumberFormat("0.0%").setValue(values["Occupancy Percentage"] + "%");
    sheet.getRange(22, lastColumn).setNumberFormat("0.0%").setValue(values["Negative Ratings Percentage"] +"%");
    // sheet.getRange(23, lastColumn).setNumberFormat("0.0%").setValue("");
    sheet.getRange(23, lastColumn).setNumberFormat("0.0%").setValue(values["Billability Percentage"] + "%");
    sheet.getRange(24, lastColumn).setNumberFormat("0.0%").setValue(values["QA Ratings Percentage"] + "%");
    
  }

  /**
   * Sets data for MIS based on the provided data object.
   * @param {Object} dataObject - The data object containing MIS data.
  */

  setData(dataObject) {
    for (const [sheetKey, subjectObject] of Object.entries(dataObject)){
      for (const[subject, values] of Object.entries(subjectObject)) {
        if (!this.SUBJECT_LIST.includes(subject.trim()))
          continue;
        // if (subject.trim() != "Total")
        //   continue;
        const folderId = CentralLibrary.createFolderIfNotExists.call(this, this.PARENT_FOLDER_ID, subject);
        Logger.log(subject)
        // if(subject === "Intro Accounting") {
          // Copy previous month sheet - Needs to change
          let fileId = ""
          var folder = DriveApp.getFolderById(folderId);
          var files = folder.getFiles()
          if(!files.hasNext()){
            fileId = CentralLibrary.createSheetInFolder.call(this, folderId, `MIS-${subject}-${sheetKey}`);
          }else {
            while (files.hasNext()) {
              var file = files.next();
              const fileName = file.getName();
              if (fileName == `MIS | Monthly Report | ${subject} | ${this.PREVIOUS_MONTH} ${this.PREVIOUS_YEAR}`){
                console.log("File name", fileName)
                const newFile = file.makeCopy();
                newFile.setName(`MIS | Monthly Report | ${subject} | ${this.MONTH} ${this.YEAR}`);
                fileId = newFile.getId();
                this.setDataIndividualSheet(fileId, sheetKey, values, dataObject);
                // New logic for lab grading
                if (subject == "Chemistry") {
                  //const {labs, mapper} = this.labGrading.chemistry();
                  //this.labGrading.setChemistry(fileId, labs, mapper)
                }
                
                if (subject == "Biology") {
                  //const {labs, mapper} = this.labGrading.biology();
                  //this.labGrading.setBiology(fileId, labs, mapper)
                }

                if (subject == "Physics") {
                  //const {labs, mapper} = this.labGrading.physics();
                  //this.labGrading.setPhysics(fileId, labs, mapper)
                }

                if (subject == "English") {
                  //const {labs, mapper} = this.labGrading.english();
                  //this.labGrading.setEnglish(fileId, labs, mapper)
                }
                break;
              }
            }
          }
      }
    }
  }


  transformData(data) {
    const selectedKeys = [
        "Actual Delivered Hours",
        "Leakage Percentage",
        "Occupancy Percentage",
        "Negative Ratings Percentage",
        "Billability Percentage",
        "QA Ratings Percentage",
    ];
    let transformed = {};

    Object.keys(data).forEach(month => {
        if (!transformed.hasOwnProperty(month)) {
            transformed[month] = {};
        }

        selectedKeys.forEach(metric => {
            transformed[month][metric] = {};
            Object.keys(data[month]).forEach(subject => {
                if (data[month][subject].hasOwnProperty(metric)) {
                    transformed[month][metric][subject] = data[month][subject][metric];
                }
            });
        });
    });

    return transformed;
  }


  getQARatingsData() {

    const subjectObject = {}
    const sheet = SpreadsheetApp.openById(this.QA_RATINGS_SHEET).getSheetByName("QA DB");
    let [headers, data] = CentralLibrary.get_Data_Indices_From_Sheet(sheet);

    data = data.filter(row => row[headers["Session Date"]] !== '' && CentralLibrary.monthNumToMonthName(row[headers["Session Date"]].getMonth()+1) === this.MONTH)
    data = data.filter(row => row[headers["Total Hours (In Decimals)"]] === '' || row[headers["SME Name"]] !== '');
    
    for (let department of this.SUBJECT_LIST) {
      let departmentData, subject;
      if (department === "Finance" || department === "Economics") {
        
        subject = department;
        department = "Business";
        departmentData = data.filter(row => row[headers["Department"]] === "Business" && row[headers["Subject"]] === subject);
        
      } else if (department === "Intro Accounting") {
        
        subject = "Intro Accounting";
        department = "Business";
        departmentData = data.filter(row => row[headers["Department"]] === "Business" && 
                                                  row[headers["Subject"]] !== "Finance" && 
                                                  row[headers["Subject"]] !== "Economics"
                                          );
      } else {

        subject = department;
        departmentData = data.filter(row => row[headers["Department"]] === department);
      
      }
      subjectObject[subject] = this.getAverageForDepartment(headers, departmentData, "Subject Knowledge", "Tutoring", "Admin", "Communication");
    }

    this.getTotalAverage(subjectObject);
    return subjectObject;

  }

  getTotalAverage(subjectObject) {


    for (const [ subject, subjectValues ] of Object.entries(subjectObject)) {

      if (subject !== "Total") {

        for (const [category, categoryObject] of Object.entries(subjectValues.categoryAverages)) {
          const {percentages, hours} = categoryObject;
          subjectObject["Total"].categoryAverages[category].percentages += percentages * hours
          subjectObject["Total"].categoryAverages[category].hours += hours;
        }

      }
      
    }
    
    let weightedPercentage = 0, totHours = 0
    for (const [key, object] of Object.entries(subjectObject["Total"].categoryAverages)) {

      object.percentages = object.percentages / object.hours;
      weightedPercentage += object.percentages * object.hours;
      totHours += object.hours

    }
    subjectObject["Total"]["QA Ratings Percentage"] = weightedPercentage / totHours;
    
  } 


  getAverageForDepartment(headers, data, ...categories) {
    // Initialize an object to hold the averages for each category
    const categoryAverages = categories.reduce((acc, category) => ({
      ...acc,
      [category]: {
        percentages: 0,
        hours: 0,
      },
    }), {});
    

    // Helper function to calculate average
    const avgCalculation = (arr) => arr.length > 0 ? 100 * (arr.reduce((acc, val) => acc + val, 0) / arr.length) : 0;

    // Iterate over each category to calculate averages
    categories.forEach(category => {
      const categoryValues = data.map(row => parseFloat(row[headers[category]])).filter(num => !isNaN(num));
      const average = avgCalculation(categoryValues);
      categoryAverages[category].percentages = average;
      categoryAverages[category].hours = categoryValues.length; // Assuming you want to count valid entries as 'hours'
    });

    
    // Optionally, you can calculate the overall average across all categories if needed
    const total = categories.reduce((acc, category) => acc + categoryAverages[category].percentages * categoryAverages[category].hours, 0);
    const totalCount = categories.reduce((acc, category) => acc + categoryAverages[category].hours, 0);
    const overallAverage = totalCount > 0 ? total / totalCount : 0;
    // You can return the calculated averages if needed
    
    return { categoryAverages, "QA Ratings Percentage" : overallAverage };
  }


  setDepartmentwiseDataForTotalSheet(sheet, dataObject, lastColumn) {
    
    // const sheet = SpreadsheetApp.openById(fileId).getSheetByName("Departmentwise Data");
    const transforemedData = this.transformData(dataObject);

    let index = 2;
    sheet.deleteColumn(2)
    for (const [monthKey, outerObject] of Object.entries(transforemedData)) {
    
      for (const [key, valueObject] of Object.entries(outerObject)) {
        
        if (key === "Actual Delivered Hours" || key === "Leakage Percentage" || key === "Occupancy Percentage" || 
            key === "Negative Ratings Percentage" || key === "Billability Percentage" || key === "QA Ratings Percentage")
        { 
          CentralLibrary.applyCustomFormatting(sheet.getRange(index, 1), { "bgColor" : "#6d9eeb", "fontWeight": "bold" }).setValue(key)
          CentralLibrary.applyCustomFormatting(sheet.getRange(index, lastColumn), { "bgColor": "#6fb43f", "fontWeight": "bold" })
                        .setNumberFormat("MMM-YYYY").setValue(monthKey)
          for (const [subject, value] of Object.entries(valueObject)) {
            
            if (this.SUBJECT_LIST.includes(subject) && subject !== "Total") {
              
              index += 1;
              CentralLibrary.applyCustomFormatting(sheet.getRange(index, 1)).setValue(subject);
              if (key === "Actual Delivered Hours")
                CentralLibrary.applyCustomFormatting(sheet.getRange(index, lastColumn)).setValue(value);
              else if (key === "Leakage Percentage")
                CentralLibrary.applyCustomFormatting(sheet.getRange(index, lastColumn)).setNumberFormat("0.00%").setValue(value + "%")
              else
                CentralLibrary.applyCustomFormatting(sheet.getRange(index, lastColumn)).setNumberFormat("0.0%").setValue(value + "%")
            }
          }
        }
        index += 2;
      }
    }
    
  }


  getData() {
    return this.combineKeys(this.getRatingsData(), this.getOccupancyData(), this.getClientHours(), this.getBillabilityData(), this.getQARatingsData());
  }

  
  getRatingsData() {
    
    const subjectObject = {}
    const inputSheet = SpreadsheetApp.openById(this.BRAINFUSE_RATINGS_SHEET).getSheetByName("Positive_N_Negative_Ratings");
    const [headers, data] = CentralLibrary.get_Data_Indices_From_Sheet.call(this, inputSheet);
    const filteredData = data.filter(row => row[headers["Year"]] === this.YEAR && row[headers["Month"]] === this.MONTH);

    const subjects = [...new Set(data.map(row => row[headers["Department"]]))].filter(Boolean);
    subjects.forEach(subject => {
      subject = subject.trimRight()
      subjectObject[subject] = { "Total Students": 0, "Positive Ratings": 0 , "Negative Ratings": 0, "Negative Ratings Percentage": 0 };
      const subjectData = filteredData.filter(row => row[headers["Department"]] === subject)
      subjectObject[subject]["Total Students"] = subjectData.map(row => row[headers["Total Students"]]).reduce((acc, curr) => acc + curr, 0);
      subjectObject[subject]["Positive Ratings"] = subjectData.map(row => row[headers["Positive Ratings"]]).reduce((acc, curr) => acc + curr, 0);
      subjectObject[subject]["Negative Ratings"] = subjectData.map(row => row[headers["Negative Ratings"]]).reduce((acc, curr) => acc + curr, 0);
      let negativeRatingsPercentage = 0;

      if (subjectObject[subject]["Total Students"] !== 0)
        negativeRatingsPercentage = 
          ((subjectObject[subject]["Negative Ratings"] / subjectObject[subject]["Total Students"])*100).toFixed(1);

      subjectObject[subject]["Negative Ratings Percentage"] = negativeRatingsPercentage;
    });
    return subjectObject;

  }

  getOccupancyData() {
    const subjectObject = {};
    const inputSheet = SpreadsheetApp.openById(this.OCCUPANCY_SHEET).getSheetByName("Master_Data");
    const [headers, data] = CentralLibrary.get_Data_Indices_From_Sheet.call(this, inputSheet);
    const filteredData = data.filter(row => row[headers["Financial Year"]] === this.FIN_YEAR && row[headers["Month"]] === this.MONTH);
    const subjects = [...new Set(data.map(row => row[headers["Subject"]]))].filter(Boolean);
    subjects.forEach(subject => {
      subject = subject.trimRight()
      const subjectData = filteredData.filter(row => row[headers["Subject"]] === subject);
      if (subject === "Calculus" || subject === "Mathematics")
        subject = "Mathematics"
      if (subject === "Intro Accounting" || subject === "Accounting")
        subject = "Intro Accounting"
      subjectObject[subject] = { "Occupancy": 0, "Non-Occupancy": 0, "Occupancy Percentage": 0 };
        subjectData.forEach(row => {
          if (row[headers["Occupancy"]] === "Occupancy") {
            subjectObject[subject]["Occupancy"] += row[headers["Hours"]];
          } else if (row[headers["Occupancy"]] === "Non-Occupancy") {
            subjectObject[subject]["Non-Occupancy"] += row[headers["Hours"]];
          }
          let occupancyPercentage = 0

          if (subjectObject[subject]["Non-Occupancy"] + subjectObject[subject]["Occupancy"] !== 0)
            occupancyPercentage = 
            ((subjectObject[subject]["Occupancy"] / (subjectObject[subject]["Non-Occupancy"] + subjectObject[subject]["Occupancy"])) * 100)
                .toFixed(1)

          subjectObject[subject]["Occupancy Percentage"] = occupancyPercentage;
        });
    });
    
    return subjectObject
  }


  getClientHours() {
    const subjectObject = {}
    const inputSheet = SpreadsheetApp.openById(this.CLIENT_HOURS_SHEET).getSheetByName("Master_Data");
    const [headers, data] = CentralLibrary.get_Data_Indices_From_Sheet.call(this, inputSheet);
    const filteredData = data.filter(row => row[headers["Client"]] === "Brainfuse" && row[headers["Year"]] === this.YEAR && row[headers["Month"]] === this.MONTH)
    const subjects = [...new Set(data.map(row => row[headers["Subject"]]))].filter(Boolean);
    subjects.forEach(subject => {
      subject = subject.trimRight()
      subjectObject[subject] = this.filterBySubject(headers, filteredData, subject)
    })
  
    return subjectObject;
  }

  filterBySubject(headers, data, subject) {
    const object = {"Client Received Hours":0, "Scheduled Hours": 0, "Actual Delivered Hours": 0, "Leakage Percentage":0 };
    data = data.filter(row => row[headers["Subject"]] === subject);
    const clientReceivedHours = data.map(row => row[headers["Client Received Hours"]]).reduce((acc, curr) => acc + curr, 0);
    const clientScheduledHours = data.map(row => row[headers["Scheduled Hours"]]).reduce((acc, curr) => acc + curr, 0);
    const actualDeliveredHours = data.map(row => row[headers["Actual Delivered Hours"]]).reduce((acc, curr) => acc + curr, 0);
    let leakagePercentage = 0
    if (actualDeliveredHours != 0)
      leakagePercentage = (((clientScheduledHours - actualDeliveredHours) / actualDeliveredHours) * 100).toFixed(2);
    object["Client Received Hours"] = clientReceivedHours;
    object["Scheduled Hours"] = clientScheduledHours;
    object["Actual Delivered Hours"] = actualDeliveredHours;
    object["Leakage Percentage"] = leakagePercentage;
    
    return object;
  }
 
  getBillabilityData() {
    const subjectObject = {}
    const inputSheet = SpreadsheetApp.openById(this.BILLABILITY_SHEET).getSheetByName("Custom Time Logs_Hours_1");
    const [headers, data] = CentralLibrary.get_Data_Indices_From_Sheet.call(this, inputSheet);
    const subjects = [...new Set(data.map(row => row[headers["Department"]]))].filter(Boolean);
    subjects.forEach(subject => {
      subject = subject.trimRight();
      const subjectData = data.filter(row => row[headers["Department"]] === subject )
      if (subject === "Accounts")
        subject = "Intro Accounting"
      subjectObject[subject] = { "Billable Hours": 0, "Non-Billabale Hours": 0, "Billability Percentage": 0 };
      
      subjectData.forEach(row => {
        if (row[headers["Billability"]] === "Billable") {
          subjectObject[subject]["Billable Hours"] += parseFloat(this.convertTimeToDecimal(row[headers["Hour(s)"]]).toFixed(2));
          // subjectObject[subject]["Billable Hours"] += row[headers["Hour(s)"]];
        } else if (row[headers["Billability"]] === "Non - Billable") {
          subjectObject[subject]["Non-Billabale Hours"] += parseFloat(this.convertTimeToDecimal(row[headers["Hour(s)"]]).toFixed(2));
          // subjectObject[subject]["Non-Billabale Hours"] += row[headers["Hour(s)"]];
        }
        let billabilityPercentage = 0;
        if(subjectObject[subject]["Billable Hours"] + subjectObject[subject]["Non-Billabale Hours"] !== 0) {
          billabilityPercentage = ((subjectObject[subject]["Billable Hours"] / (subjectObject[subject]["Billable Hours"] + subjectObject[subject]["Non-Billabale Hours"])) * 100).toFixed(1);
        }
        
        subjectObject[subject]["Billability Percentage"] = billabilityPercentage;
      })
    });

    return subjectObject;
  }
  
  combineKeys(...datasets) {
    const combinedData = {};
    const monthKey = `${this.convertMonthFormat(this.MONTH)}-${this.YEAR}`;

    // Find all unique keys across all datasets
    const allKeys = datasets.reduce((keys, dataset) => {
        Object.keys(dataset).forEach(subject => {
            keys.push(...Object.keys(dataset[subject]));
        });
        return keys;
    }, []);

    // Remove duplicates and sort keys alphabetically
    const uniqueKeys = [...new Set(allKeys)].sort();

    // Iterate over each dataset
    datasets.forEach(dataset => {
        // Iterate over each subject in the dataset
        for (const subject in dataset) {
            // If the subject already exists in combinedData, merge the keys, otherwise, add it to combinedData
            combinedData[subject] = { ...(combinedData[subject] || {}), ...dataset[subject] };

            // For each unique key, if the subject doesn't have it, set its value to 0
            uniqueKeys.forEach(key => {
                if (!(key in combinedData[subject])) {
                    combinedData[subject][key] = 0;
                }
            });
        }
    });

    const totalData = {
      "QA Ratings Percentage" : combinedData["Total"]["QA Ratings Percentage"].toFixed(2),
      };
    uniqueKeys.forEach(key => {
      
      if (key !== "Billability Percentage" && key !== "Negative Ratings Percentage" 
          && key !== "Negative Ratings Percentage" && key !== "Occupancy Percentage" 
          && key !== "QA Ratings Percentage" && key !== "categoryAverages")
        totalData[key] = Object.values(combinedData).reduce((sum, subjectData) => sum + (parseFloat(subjectData[key]) || 0), 0);
    });
    combinedData["Total"] = totalData;
    
    combinedData["Total"]["Leakage Percentage"] = 
      (((combinedData["Total"]["Client Received Hours"] - combinedData["Total"]["Actual Delivered Hours"]) / combinedData["Total"]["Actual Delivered Hours"]) * 100)
            .toFixed(2);
    combinedData["Total"]["Occupancy Percentage"] = 
      ((combinedData["Total"]["Occupancy"] / (combinedData["Total"]["Non-Occupancy"] + combinedData["Total"]["Occupancy"])) * 100).toFixed(1)
    combinedData["Total"]["Negative Ratings Percentage"] = ((combinedData["Total"]["Negative Ratings"] / combinedData["Total"]["Total Students"])*100).toFixed(1);
    combinedData["Total"]["Billability Percentage"] = 
      ((combinedData["Total"]["Billable Hours"] / (combinedData["Total"]["Billable Hours"] + combinedData["Total"]["Non-Billabale Hours"])) * 100).toFixed(1)
    const result = {};

    result[monthKey] = combinedData;
    return result;
  }


  convertTimeToDecimal(time) {
    // Split the time string into hours and minutes
    var parts = time.split(':');
    var hours = parseInt(parts[0]);
    var minutes = parseInt(parts[1]);

    // Calculate the decimal representation of time
    var decimalTime = hours + minutes / 60;

    return decimalTime;
  }

  convertMonthFormat(longMonthName) {
    const monthMap = {
        "January": "Jan",
        "February": "Feb",
        "March": "Mar",
        "April": "Apr",
        "May": "May",
        "June": "Jun",
        "July": "Jul",
        "August": "Aug",
        "September": "Sep",
        "October": "Oct",
        "November": "Nov",
        "December": "Dec"
    };

    return monthMap[longMonthName] || longMonthName;
  }

}











