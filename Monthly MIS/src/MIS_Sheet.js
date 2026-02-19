function report_main() {

  const report = new MIS_Report();
  report.sheetLoop()
}

class MIS_Report {

  constructor() {

    this.PARENT_FOLDER = "16G65r1bwp5bRKB6N3PW2nITm-7QH5F24";
    this.SUBJECT_LIST = [
                          // "Mathematics", 
                          // "Statistics", 
                          "Physics", 
                          "Chemistry", 
                          "Biology", 
                          // "Intro Accounting", 
                          // "Finance", 
                          // "Economics",
                          // "Computer Science", 
                          "English", 
                          "Total"
                        ];

    this.DATE = new Date();
    this.YEAR = this.DATE.getFullYear();
    this.MONTH = this.DATE.getMonth();
    if (this.MONTH === 0) {
      this.YEAR -= 1;
      this.MONTH = CentralLibrary.monthNumToMonthName.call(this, this.MONTH + 12)
    }else{
      this.MONTH = CentralLibrary.monthNumToMonthName.call(this, this.MONTH);
    }
    this.MONTH_YEAR_KEY = `${this.MONTH} ${this.YEAR}`;
    console.log(this.MONTH_YEAR_KEY)
  }

  sheetLoop () {

    const folderObject = this.listSubfolders(this.PARENT_FOLDER)
    for (const [folderName, fileObject] of Object.entries(folderObject)) {

      if (this.SUBJECT_LIST.includes(folderName)) {

        for (const [fileName, fileId] of Object.entries(fileObject)) {
          const fileNameSplit = fileName.split("|");
          if (fileNameSplit[3] !== undefined && fileNameSplit[3].trim() === this.MONTH_YEAR_KEY) {
              console.log(fileId)
              this.modifyEachSheet(fileId)
            
          }
        }        
      }
    }
  }

  modifyEachSheet(fileId) {

    const spreadsheet = SpreadsheetApp.openById(fileId);
    const allSheetNames = spreadsheet.getSheets().map(sheet => sheet.getName());
    let reportSheet;
    
    if (!allSheetNames.includes("Report"))
      reportSheet = spreadsheet.insertSheet("Report");
    else
      reportSheet = spreadsheet.getSheetByName("Report");
    
    const dataSheet = spreadsheet.getSheetByName("Backend DB");
    const data = dataSheet.getDataRange().getValues();
    reportSheet.setHiddenGridlines(true)


    const tableData = this.tableChart(spreadsheet, data);
    this.setTableChart(tableData, reportSheet)

    const existingCharts = reportSheet.getCharts();
    for (const chart of existingCharts) {
      reportSheet.removeChart(chart);
    }
    
    this.createChart1(spreadsheet);
    this.createChart2(spreadsheet);
    this.createChart3(spreadsheet);
  }


  setTableChart(data, sheet) {
    CentralLibrary.applyCustomFormatting(sheet.getRange(1, 1, 4, 1)).setValues([[""], ["Current"], ["Y-o-Y"], ["M-o-M"] ])
    // CentralLibrary.applyCustomFormatting(sheet.getRange(1, 2, data.length, data[0].length)).setValues(data);

    const tutoringHours = data.map(arr => arr[0])
    tutoringHours.forEach((val, index) => {

      if (index === 0 || index === 1)
        CentralLibrary.applyCustomFormatting(sheet.getRange(index+1, 2)).setValue(val);
      else
        CentralLibrary.applyCustomFormatting(sheet.getRange(index+1, 2)).setValue(`${val}(${this.calculateMoMChange(tutoringHours[1], val).toFixed(1)}%)`);
    })

    const percentValues = data.map(arr => arr.slice(1));
    // console.log(percentValues)
    percentValues.forEach((arr, index) => {
      
      if (index === 0)
        CentralLibrary.applyCustomFormatting(sheet.getRange(index+1, 3, 1, data[0].length-1)).setValues([arr])
      else {
        arr.forEach((val, idx) => {
          if (val === "" || val === 0)
            CentralLibrary.applyCustomFormatting(sheet.getRange(index+1, 3+idx)).setValue("-");
          else
            CentralLibrary.applyCustomFormatting(sheet.getRange(index+1, 3+idx).setNumberFormat("0.00%")).setValue(val + "%");
            
        })
      }
        
    });
  }


  //Old code
  tableChart(spreadsheet, data) {

    const dept = spreadsheet.getName().split("|")[2].trim();
    console.log(dept);

    // const headerMap = data.map(r => r[0]).reduce((object,  curr, index) => {
    //   object[curr] = index;
    //   return object; 
    // }, {})

    const headerMap = data.reduce((object, row, index) => {
      const key = row[0] ? row[0].toString().trim() : "";
      if (key) object[key] = index;
      return object;
    }, {});
    Logger.log(headerMap);

    const tutoringHours = data[headerMap["Actual Hours"]].slice(1);
    const positiveRatings = data[headerMap["Positive Ratings"]].slice(1);
    const negativeRatings = data[headerMap["Negative Ratings"]].slice(1);
    const negativeRatingStudentsPercent = this.findPercent(negativeRatings, positiveRatings);

    const leakagePercentages = data[headerMap["Leakage Percentage"]].slice(1);
    const occupancyPercentages = data[headerMap["Occupancy Percentage"]].slice(1);
    const negativeRatingPercentages = data[headerMap["Negative Ratings Percentage"]].slice(1);
    const billabilityPercentages = data[headerMap["Billability Percentage"]].slice(1);
    const qaRatingsPercentages = data[headerMap["QA Ratings Percentage"]].slice(1);


    const tableData = [];
    Logger.log(JSON.stringify(data));


    tableData.push([
      "BF-Tutoring Hours Done",
      "Leakage %",
      "Billability %",
      "Occupancy %",
      "Negative Ratings % (#Students)",
      "Negative Ratings % (#Ratings)",
      "QA Ratings %"
    ]);
    // Current month
    tableData.push([
      tutoringHours[tutoringHours.length - 1],
      leakagePercentages[leakagePercentages.length - 1] * 100,
      billabilityPercentages[billabilityPercentages.length - 1] * 100,
      occupancyPercentages[occupancyPercentages.length - 1] * 100,
      negativeRatingPercentages[negativeRatingPercentages.length - 1] * 100,
      negativeRatingStudentsPercent[negativeRatingStudentsPercent.length - 1],
      qaRatingsPercentages[qaRatingsPercentages.length - 1] * 100,
    ]);

    tableData.push([
      tutoringHours[0],
      leakagePercentages[0] * 100,
      billabilityPercentages[0] * 100,
      occupancyPercentages[0] * 100,
      negativeRatingPercentages[0] * 100,
      negativeRatingStudentsPercent[0],
      qaRatingsPercentages[0] * 100,
    ]);
    tableData.push([
      tutoringHours[tutoringHours.length - 2],
      leakagePercentages[leakagePercentages.length - 2] * 100,
      billabilityPercentages[billabilityPercentages.length - 2] * 100,
      occupancyPercentages[occupancyPercentages.length - 2] * 100,
      negativeRatingPercentages[negativeRatingPercentages.length - 2] * 100,
      negativeRatingStudentsPercent[negativeRatingStudentsPercent.length - 2],
      qaRatingsPercentages[qaRatingsPercentages.length - 2] * 100,
    ])

    // New code for lab grading
    if (dept == "English" || dept == "Chemistry" || dept == "Physics" || dept == "Biology") {
      tableData[0].unshift("Labs Done");
      const labDoneArray = data[headerMap["Labs Done"]].slice(1);
      tableData[1].unshift(labDoneArray[labDoneArray.length - 1]);
      tableData[2].unshift(labDoneArray[0]);
      tableData[3].unshift(labDoneArray[labDoneArray.length - 2]);
    }

    return tableData;

  }

 
  calculateMoMChange(current, previous) {
    if (previous !== 0 && previous !== '')
      return ((current - previous) / previous) * 100;
    else
      return current * 100;
  }


  createChart1(spreadsheet) {
    const sheet = spreadsheet.getSheetByName("Report");
    const backendSheet = spreadsheet.getSheetByName("Backend DB");

    // Create the chart
    const chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.COMBO)

      // Add labels and data series
      .addRange(backendSheet.getRange('A1:N1')) // Labels
      .addRange(backendSheet.getRange('A4:N4')) // Data series 1
      .addRange(backendSheet.getRange('A20:N20')) // Data series 2
      .setNumHeaders(1)
      // Set chart position and style
      .setPosition(8, 1, 0, 0)
      .setTransposeRowsAndColumns(true)
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_ROWS)

      .setOption('series', {
        0: {
          type: 'bars',
          targetAxisIndex: 0,
          color: '#C9DAF8' // Blue for bars
        },
        1: {
          type: 'line',
          dataLabel: 'value',
          targetAxisIndex: 1,
          color: '#0000FF',
          pointSize: 6,
        }
      })
    chartBuilder
      .setOption('vAxes', {
        1: { title: 'Leakage Percentage', gridlines: { count: 8 }, minValue: 0 },
        0: { title: 'Actual Hours' },

      })

    // Set chart title and style
    chartBuilder
      .setOption('title', 'Actual Hours vs Leakage Percentage')
      .setOption('titleTextStyle', { color: 'black', fontSize: 14, bold: true, alignment: 'center' })

      // Correct way to set legend text and style
      .setOption('legend',
        {
          position: 'bottom',
          textStyle: { fontSize: 12 },
          alignment: 'center',
        });

    // Insert the chart into the sheet
    sheet.insertChart(chartBuilder.build());
  }

  createChart2(spreadsheet) {
    const sheet = spreadsheet.getSheetByName("Report");
    const backendSheet = spreadsheet.getSheetByName("Backend DB");

    // Create the chart
    const chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      // Add labels and data series
      .addRange(backendSheet.getRange('A1:N1')) // Labels
      .addRange(backendSheet.getRange('A21:N21')) // Data series 1
      .addRange(backendSheet.getRange('A23:N23')) // Data series 2
      .setNumHeaders(1)
      // Set chart position and style
      .setPosition(8, 8, 0, 0)
      .setTransposeRowsAndColumns(true)
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_ROWS)
      // .setOption('treatLabelsAsText', true)
      // Configure series (without specifying legend names here)
      .setOption('series', {
        0: {
          type: 'line',
          targetAxisIndex: 0,
          dataLabel: 'value',
          pointSize: 6,
          color: '#F28C28',

        },
        1: {
          type: 'line',
          targetAxisIndex: 1,
          dataLabel: 'value',
          pointSize: 6,
          dataLabelPlacement: "below",
          color: '#A0522D'
        }
      })
      .setOption('hAxis', {
        gridlines: { count: 10, color: "#fff2cc" },
        textStyle: { fontSize: 12, fontName: "Roboto" },
        slantedText: true,
        slantedTextAngle: 45, // Sets the slant angle for the horizontal axis labels to 60 degrees

      })

      .setOption('vAxis', {
        title: "Percentages",
        gridlines: { color: "#fff2cc" }, // This will hide all gridlines
      })
      // Set chart title and style
      .setOption('title', 'Occupancy vs Billability')
      .setOption('titleTextStyle', { color: 'black', fontSize: 14, bold: true, italic: false, alignment: 'center', fontName: "Roboto" })

      // Correct way to set legend text and style
      .setOption('legend',
        {
          position: 'bottom',
          textStyle: { fontSize: 12 }
        });

    // Insert the chart into the sheet
    sheet.insertChart(chartBuilder.build());
  }

  createChart3(spreadsheet) {
    const sheet = spreadsheet.getSheetByName("Report");
    const backendSheet = spreadsheet.getSheetByName("Backend DB");

    // Create the chart
    const chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)

      // Add labels and data series
      .addRange(backendSheet.getRange('A1:N1')) // Labels
      .addRange(backendSheet.getRange('A22:N22')) // Data series 1
      .addRange(backendSheet.getRange('A24:N24')) // Data series 2
      .setNumHeaders(1)
      // Set chart position and style
      .setPosition(28, 5, 0, 0)
      .setTransposeRowsAndColumns(true)
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_ROWS)
      .setOption('series', {
        0: {
          targetAxisIndex: 0,
          dataLabel: 'value',
          pointSize: 5,
          dataLabelPlacement: "below",
          color: '#F28C28',
        },
        1: {
          targetAxisIndex: 1,
          dataLabel: 'value',
          pointSize: 5,
          dataLabelPlacement: "top",
          color: '#0000FF',
        }
      })
      .setOption('hAxis', {
        gridlines: { count: 10, color: "#fff2cc" },
        textStyle: { fontSize: 12, fontName: "Roboto" },
        slantedText: true,
        slantedTextAngle: 45 // Sets the slant angle for the horizontal axis labels to 60 degrees

      })
      .setOption('vAxis', {
        title: "Percentages",
        gridlines: { color: "#fff2cc" },
      })
      // Set chart title and style
      .setOption('title', 'Negative Ratings vs QA-Ratings')
      .setOption('titleTextStyle', { color: 'black', fontSize: 14, bold: true, italic: false, alignment: 'center', fontName: "Roboto" })

      // Correct way to set legend text and style
      .setOption('legend',
        {
          position: 'bottom',
          alighment: 'center',
          textStyle: { fontSize: 12 }
        });



    // Insert the chart into the sheet
    sheet.insertChart(chartBuilder.build());
  }



  findPercent(arr1, arr2) {
    if (arr1.length !== arr2.length) {
      throw new Error("Arrays must have the same length");
    }

    return arr1.map((value, index) => ((100 * value) / (value + arr2[index])).toFixed(2));
  }


  listSubfolders(parentFolderId) {
    var parentFolder = DriveApp.getFolderById(parentFolderId);
    var subfolders = parentFolder.getFolders();
    const folderObject = {}
    while (subfolders.hasNext()) {
      var subfolder = subfolders.next();
      folderObject[subfolder.getName()] = this.listFiles(subfolder.getId());
    }
    return folderObject;
  }

  listFiles(folderId) {
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();
    const fileObject = {}
    while (files.hasNext()) {
      var file = files.next();
      // Logger.log("File Name: " + file.getName());
      fileObject[file.getName()] = file.getId();
      // Process the file as needed
    }
    return fileObject;
  }


}






  // createChart(spreadsheet, data) {
  //   const sheet = spreadsheet.getSheetByName("Report");
  //   const backendSheet = spreadsheet.getSheetByName("Backend DB");
  //   // Create the chart
  //   const chartBuilder = sheet.newChart()
  //   chartBuilder
  //     .setChartType(Charts.ChartType.COMBO)
  //     .addRange(backendSheet.getRange('A1:M1')) // Labels
  //     .addRange(backendSheet.getRange('A4:M4')) // Data series
  //     .setPosition(5, 5, 0, 0) // Position where the chart will be placed in the sheet
  //     .setTransposeRowsAndColumns(true)
  //     .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_ROWS)

  
  //   // Insert the chart into the sheet
  //   sheet.insertChart(chartBuilder.build());
  // }
