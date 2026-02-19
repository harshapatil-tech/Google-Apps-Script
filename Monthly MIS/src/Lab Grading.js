function labGrading() {
  const x = new LabGrading('Aug', 2025);
  //const { labs, mapper } = x.biology();

  // const { labs, mapper } = x.chemistry();
  const { labs, mapper } = x.physics();
  
  console.log(mapper)
  // x.setBiology("1Q2N-SLCHo6U28H9X2eEMLHwpNNEOl4040paLOA5HcwE", labs, mapper)
  // x.setChemistry("1btYYohna-ijR9bK7acNh0MDKP1h1rHv3NOW5-PGlcgQ", labs, mapper)
  x.setPhysics("1JtDE_aXgfov6obc4dxLnIuCiyUOD0ieJCc9MHeSCA3E", labs, mapper)

}

class LabGrading {

  constructor(month, year) {
    this.month = month;
    this.year = year;

    console.log(this.year, this.month)
  }

  //OLD CODE
  chemistry() {
    const chemistry_2025 = SpreadsheetApp.openById("1ykoy8z-FNDhZ8Ht1M8GkZEBVlXxvbBM8pb1GZvMaEQ0");
    const sheet = chemistry_2025.getSheetByName("Sheet1");
    const [ headers, data ] = CentralLibrary.get_Data_Indices_From_Sheet(sheet);

    const mapper = {};
    const labNums = new Set();

    data.forEach(row => {
      const date   = row[headers["Year"]];
      const month  = row[headers["Month"]];
      // Normalize to string and trim whitespace
      const rawLab = row[headers["Lab No."]];
      const labNum = rawLab != null
        ? rawLab.toString()
        : "";

      // Skip if labNum is empty after trimming
      if (labNum === "") return;

      mapper[date] = mapper[date] || {};
      mapper[date][month] = mapper[date][month] || {};
      mapper[date][month][labNum] = (mapper[date][month][labNum] || 0) + 1;
      labNums.add(labNum)
    });

    const labs = [...labNums].sort();
    return {labs, mapper};
  }

  setChemistry (fileId, labs, mapper) {
    const ss = SpreadsheetApp.openById(fileId);
    const sheet = ss.getSheetByName("Backend DB");

    sheet.getRange(26, 1, labs.length, 1).setValues(labs.map(lab => [lab]));

    const months = sheet.getRange(1, 2, 1, sheet.getLastColumn()).getValues().flat().filter(Boolean);

    const monthMap = months.reduce((map, key, index) => {
      if (key != "") {
        map[key] = index+2;
      }
      return map;
    }, {})


    for (const [month, col] of Object.entries(monthMap)) {
      const d = new Date(month);
      const Y   = d.getFullYear().toString();
      const M   = d.toLocaleString('default', { month: 'short' });
      // const M   = d.toLocaleString('default', { month: 'long' });

      if (mapper[Y] && mapper[Y][M]) {
        labs.forEach((lab, idx) => {
          const row   = 26 + idx;
          const count = mapper[Y][M][lab];
          sheet.getRange(row, col).setValue(count);
        });
      }
    }
  }


  // ************************************/

  // chemistry() {
  //   const chemistry_2025 = SpreadsheetApp.openById("1ykoy8z-FNDhZ8Ht1M8GkZEBVlXxvbBM8pb1GZvMaEQ0");
  //   const sheet = chemistry_2025.getSheetByName("Sheet1");
  //   const [headers, data] = CentralLibrary.get_Data_Indices_From_Sheet(sheet);

  //   const mapper = {};
  //   const labNums = new Set();

  //   data.forEach(row => {
  //     const date = row[headers["Year"]];
  //     const month = row[headers["Month"]];
  //     // Normalize to string and trim whitespace
  //     const rawLab = row[headers["Lab No."]];
  //     const labNum = rawLab != null
  //       ? rawLab.toString()
  //       : "";

  //     // Skip if labNum is empty after trimming
  //     if (labNum === "") return;

  //     mapper[date] = mapper[date] || {};
  //     mapper[date][month] = mapper[date][month] || {};
  //     mapper[date][month][labNum] = (mapper[date][month][labNum] || 0) + 1;
  //     labNums.add(labNum)
  //   });

  //   const labs = [...labNums].sort();
  //   return { labs, mapper };
  // }

  // setChemistry(fileId, labs, mapper) {
  //   const ss = SpreadsheetApp.openById(fileId);
  //   const sheet = ss.getSheetByName("Backend DB");

  //   // Clear old data below row 26
  //   sheet.getRange(26, 1, sheet.getMaxRows() - 25, sheet.getLastColumn()).clearContent();

  //   // Column A → labs list
  //   sheet.getRange(26, 1, labs.length, 1).setValues(labs.map(lab => [lab]));

  //   // Months headers in row 1 (B1 → onwards)
  //   const months = sheet.getRange(1, 2, 1, sheet.getLastColumn())
  //     .getValues().flat().filter(Boolean);

  //   const monthMap = months.reduce((map, key, index) => {
  //     if (key !== "") {
  //       map[key] = index + 2; // col index in sheet
  //     }
  //     return map;
  //   }, {});

  //   // Fill data
  //   for (const [month, col] of Object.entries(monthMap)) {
  //     const d = new Date(month);
  //     const Y = d.getFullYear().toString();
  //     const M = d.toLocaleString('default', { month: 'short' });

  //     labs.forEach((lab, idx) => {
  //       const row = 26 + idx;
  //       const count = (mapper[Y] && mapper[Y][M] && mapper[Y][M][lab]) || 0; // default 0
  //       sheet.getRange(row, col).setValue(count);
  //     });
  //   }

  //   //  Add "Labs Done" totals row
  //   const lastLabRow = 26 + labs.length - 1;
  //   const labsDoneRow = lastLabRow + 1;

  //   sheet.getRange(labsDoneRow, 1).setValue("Labs Done");

  //   for (const col of Object.values(monthMap)) {
  //     const sumFormula = `=SUM(${sheet.getRange(26, col).getA1Notation()}:${sheet.getRange(lastLabRow, col).getA1Notation()})`;
  //     sheet.getRange(labsDoneRow, col).setFormula(sumFormula);
  //   }
  // }


  // physics() {
  //   const chemistry_2025 = SpreadsheetApp.openById("1BScRgi6G7-gNHWQVJ7jOkigfDMjCdDyvjg1TXbTPfRA");
  //   const sheet = chemistry_2025.getSheetByName("Sheet1");
  //   const [headers, data] = CentralLibrary.get_Data_Indices_From_Sheet(sheet);

  //   const mapper = {};
  //   const labNums = new Set();

  //   data.forEach(row => {
  //     // const d   = new Date(row[headers["Date"]])
  //     // const date = d.getFullYear();

  //     const date = row[headers["Year"]]
  //     const month = row[headers["Month"]]
  //     // const month = d.toLocaleString('default', { month: 'short' });
  //     // Normalize to string and trim whitespace
  //     const rawLab = row[headers["Lab No."]];
  //     const labNum = rawLab != null
  //       ? rawLab.toString()
  //       : "";

  //     // Skip if labNum is empty after trimming
  //     if (labNum === "") return;

  //     mapper[date] = mapper[date] || {};
  //     mapper[date][month] = mapper[date][month] || {};
  //     mapper[date][month][labNum] = (mapper[date][month][labNum] || 0) + 1;
  //     labNums.add(labNum)
  //   });

  //   const labs = [...labNums].sort();
  //   return { labs, mapper };
  // }


  // setPhysics(fileId, labs, mapper) {
  //   const ss = SpreadsheetApp.openById(fileId);
  //   const sheet = ss.getSheetByName("Backend DB");

  //   // sheet.getRange(26, 1, labs.length, 1).setValues(labs.map(lab => [lab]));

  //   sheet.getRange(26, 1, sheet.getMaxRows() - 25, sheet.getLastColumn()).clearContent();

  //   sheet.getRange(26, 1, labs.length, 1).setValues(labs.map(lab => [lab]));

  //   const months = sheet.getRange(1, 2, 1, sheet.getLastColumn()).getValues().flat().filter(Boolean);

  //   const monthMap = months.reduce((map, key, index) => {
  //     if (key != "") {
  //       map[key] = index + 2;
  //     }
  //     return map;
  //   }, {})


  //   for (const [month, col] of Object.entries(monthMap)) {
  //     const d = new Date(month);
  //     const Y = d.getFullYear().toString();
  //     const M = d.toLocaleString('default', { month: 'short' });
  //     // const M   = d.toLocaleString('default', { month: 'long' });

  //     labs.forEach((lab, idx) => {
  //       const row = 26 + idx;
  //       const count = (mapper[Y] && mapper[Y][M] && mapper[Y][M][lab]) || 0; // default 0
  //       sheet.getRange(row, col).setValue(count);
  //     });
  //   }
  //   //  Add "Labs Done" totals row
  //   const lastLabRow = 26 + labs.length - 1;
  //   const labsDoneRow = lastLabRow + 1;

  //   sheet.getRange(labsDoneRow, 1).setValue("Labs Done");

  //   for (const col of Object.values(monthMap)) {
  //     const sumFormula = `=SUM(${sheet.getRange(26, col).getA1Notation()}:${sheet.getRange(lastLabRow, col).getA1Notation()})`;
  //     sheet.getRange(labsDoneRow, col).setFormula(sumFormula);
  //   }
  // }



  //old code
  physics() {
    const chemistry_2025 = SpreadsheetApp.openById("1BScRgi6G7-gNHWQVJ7jOkigfDMjCdDyvjg1TXbTPfRA");
    const sheet = chemistry_2025.getSheetByName("Sheet1");
    const [ headers, data ] = CentralLibrary.get_Data_Indices_From_Sheet(sheet);

    const mapper = {};
    const labNums = new Set();

    data.forEach(row => {
      // const d   = new Date(row[headers["Date"]])
      // const date = d.getFullYear();

      const date = row[headers["Year"]]
      const month  = row[headers["Month"]]
      // const month = d.toLocaleString('default', { month: 'short' });
      // Normalize to string and trim whitespace
      const rawLab = row[headers["Lab No."]];
      const labNum = rawLab != null
        ? rawLab.toString()
        : "";

      // Skip if labNum is empty after trimming
      if (labNum === "") return;

      mapper[date] = mapper[date] || {};
      mapper[date][month] = mapper[date][month] || {};
      mapper[date][month][labNum] = (mapper[date][month][labNum] || 0) + 1;
      labNums.add(labNum)
    });

    const labs = [...labNums].sort();
    return {labs, mapper};
  }


  setPhysics (fileId, labs, mapper) {
    const ss = SpreadsheetApp.openById(fileId);
    const sheet = ss.getSheetByName("Backend DB");

    // sheet.getRange(26, 1, labs.length, 1).setValues(labs.map(lab => [lab]));

    const months = sheet.getRange(1, 2, 1, sheet.getLastColumn()).getValues().flat().filter(Boolean);

    const monthMap = months.reduce((map, key, index) => {
      if (key != "") {
        map[key] = index+2;
      }
      return map;
    }, {})


    for (const [month, col] of Object.entries(monthMap)) {
      const d = new Date(month);
      const Y   = d.getFullYear().toString();
      const M   = d.toLocaleString('default', { month: 'short' });
      // const M   = d.toLocaleString('default', { month: 'long' });

      if (mapper[Y] && mapper[Y][M]) {
        labs.forEach((lab, idx) => {
          const row   = 26 + idx;
          const count = mapper[Y][M][lab];
          sheet.getRange(row, col).setValue(count);
        });
      }
    }
  }



  //old code
  biology() {
    const chemistry_2025 = SpreadsheetApp.openById("12mX5RWKvbe4LqRv6MZiTo6xBKFWdUf-aG6qrKZkEF5A");
    const sheet = chemistry_2025.getSheetByName("Sheet1");
    const [ headers, data ] = CentralLibrary.get_Data_Indices_From_Sheet(sheet);

    const mapper = {};
    console.log("mapper",mapper);
    const labNums = new Set();

    data.forEach(row => {
      // const d   = new Date(row[headers["Date"]])
      // const date = d.getFullYear();

      const date = row[headers["Year"]]
      const month  = row[headers["Month"]]
      // const month = d.toLocaleString('default', { month: 'short' });
      // Normalize to string and trim whitespace
      // const rawLab = row[headers["Lab No."]];
      let courseName = row[headers["Course name"]]
      // courseName = courseName.slice(0, 6);
      const labNum = courseName != null
        ? (courseName.charAt(0).toUpperCase() + courseName.slice(1).toLowerCase()).slice(0,6).toString()
        : "";

      // Skip if labNum is empty after trimming
      if (labNum === "") return;

      mapper[date] = mapper[date] || {};
      mapper[date][month] = mapper[date][month] || {};
      mapper[date][month][labNum] = (mapper[date][month][labNum] || 0) + 1;
      labNums.add(labNum)
    });

    const labs = [...labNums].sort();
    console.log("Labs are:-",labs);
    return {labs, mapper};
  }

  setBiology (fileId, labs, mapper) {
    const ss = SpreadsheetApp.openById(fileId);
    const sheet = ss.getSheetByName("Backend DB");

    // sheet.getRange(26, 1, labs.length, 1).setValues(labs.map(lab => [lab]));

    const months = sheet.getRange(1, 2, 1, sheet.getLastColumn()).getValues().flat().filter(Boolean);

    const monthMap = months.reduce((map, key, index) => {
      if (key != "") {
        map[key] = index+2;
      }
      return map;
    }, {})


    for (const [month, col] of Object.entries(monthMap)) {
      const d = new Date(month);
      const Y   = d.getFullYear().toString();
      const M   = d.toLocaleString('default', { month: 'short' });
      // const M   = d.toLocaleString('default', { month: 'long' });

      if (mapper[Y] && mapper[Y][M]) {
        labs.forEach((lab, idx) => {
          const row   = 26 + idx;
          const count = mapper[Y][M][lab];
          sheet.getRange(row, col).setValue(count);
        });
      }
    }
  }


  // //***************************************************/


  // biology() {
  //   const chemistry_2025 = SpreadsheetApp.openById("12mX5RWKvbe4LqRv6MZiTo6xBKFWdUf-aG6qrKZkEF5A");
  //   const sheet = chemistry_2025.getSheetByName("Sheet1");

  //   const [headers, data] = CentralLibrary.get_Data_Indices_From_Sheet(sheet);

  //   const mapper = {};
  //   console.log("mapper", mapper);
  //   const labNums = new Set();

  //   data.forEach(row => {
  //     // const d   = new Date(row[headers["Date"]])
  //     // const date = d.getFullYear();

  //     const date = row[headers["Year"]]
  //     // console.log("Date is:-",date);
  //     const month = row[headers["Month"]]
  //     // console.log("Month is:-",month);
  //     // const month = d.toLocaleString('default', { month: 'short' });
  //     // Normalize to string and trim whitespace
  //     // const rawLab = row[headers["Lab No."]];
  //     let courseName = row[headers["Course name"]]
  //     // courseName = courseName.slice(0, 6);
  //     const labNum = courseName != null
  //       ? (courseName.charAt(0).toUpperCase() + courseName.slice(1).toLowerCase()).slice(0, 6).toString()
  //       : "";

  //     // Skip if labNum is empty after trimming
  //     if (labNum === "") return;

  //     mapper[date] = mapper[date] || {};
  //     mapper[date][month] = mapper[date][month] || {};
  //     mapper[date][month][labNum] = (mapper[date][month][labNum] || 0) + 1;
  //     labNums.add(labNum)
  //   });

  //   const labs = [...labNums].sort();
  //   console.log("Labs are:-", labs);
  //   return { labs, mapper };
  // }

  // setBiology(fileId, labs, mapper) {
  //   const ss = SpreadsheetApp.openById(fileId);
  //   const sheet = ss.getSheetByName("Backend DB");

  //   sheet.getRange(26, 1, sheet.getMaxRows() - 25, sheet.getLastColumn()).clearContent();

  //   // Labs list in column A starting row 26
  //   sheet.getRange(26, 1, labs.length, 1).setValues(labs.map(lab => [lab]));

  //   // Months header (Row 1, from col B onwards)
  //   const months = sheet.getRange(1, 2, 1, sheet.getLastColumn())
  //     .getValues().flat().filter(Boolean);

  //   const monthMap = months.reduce((map, key, index) => {
  //     if (key != "") map[key] = index + 2; // col index (B=2, C=3, etc.)
  //     return map;
  //   }, {});

  //   for (const [month, col] of Object.entries(monthMap)) {
  //     const d = new Date(month);
  //     const Y = d.getFullYear().toString();
  //     const M = d.toLocaleString('default', { month: 'short' });

  //     if (mapper[Y] && mapper[Y][M]) {
  //       // Fill lab counts
  //       labs.forEach((lab, idx) => {
  //         const row = 26 + idx;
  //         const count = mapper[Y][M][lab] || 0;
  //         sheet.getRange(row, col).setValue(count);
  //       });

  //       // ---- Labs Done Row ----
  //       const lastLabRow = 26 + labs.length - 1;
  //       const labsDoneRow = lastLabRow + 1;

  //       // Label in column A
  //       sheet.getRange(labsDoneRow, 1).setValue("Labs Done");

  //       // Formula in this month column
  //       const sumFormula = `=SUM(${sheet.getRange(26, col).getA1Notation()}:${sheet.getRange(lastLabRow, col).getA1Notation()})`;
  //       sheet.getRange(labsDoneRow, col).setFormula(sumFormula);

  //     }
  //   }
  // }





  english() {
    const englishSS = SpreadsheetApp.openById("1neTzUxzlkLFO-POU2LHkUR7ba5EK56cvo493uRjTAOY");
    const sheet = englishSS.getSheetByName("Essay_details");
    const [headers, data] = CentralLibrary.get_Data_Indices_From_Sheet(sheet);

    const mapper = {};
    const labNums = new Set();

    data.forEach(row => {
      const d = new Date(row[headers["Created At (IST)"]])
      const date = d.getFullYear();

      // const date = row[headers["Year"]]
      // const month  = row[headers["Month"]]
      const month = d.toLocaleString('default', { month: 'short' });
      // Normalize to string and trim whitespace
      // const rawLab = row[headers["Lab No."]];
      let courseName = row[headers["Client"]]
      // courseName = courseName.slice(0, 6);
      const labNum = courseName != null
        ? courseName.toString()
        : "";

      // Skip if labNum is empty after trimming
      if (labNum === "") return;

      mapper[date] = mapper[date] || {};
      mapper[date][month] = mapper[date][month] || {};
      mapper[date][month][labNum] = (mapper[date][month][labNum] || 0) + 1;
      labNums.add(labNum)
    });

    const labs = [...labNums].sort();
    return { labs, mapper };
  }

  setEnglish(fileId, labs, mapper) {
    const ss = SpreadsheetApp.openById(fileId);
    const sheet = ss.getSheetByName("Backend DB");

    // sheet.getRange(26, 1, labs.length, 1).setValues(labs.map(lab => [lab]));

    const months = sheet.getRange(1, 2, 1, sheet.getLastColumn()).getValues().flat().filter(Boolean);

    const monthMap = months.reduce((map, key, index) => {
      if (key != "") {
        map[key] = index + 2;
      }
      return map;
    }, {})


    for (const [month, col] of Object.entries(monthMap)) {
      const d = new Date(month);
      const Y = d.getFullYear().toString();
      const M = d.toLocaleString('default', { month: 'short' });
      // const M   = d.toLocaleString('default', { month: 'long' });

      if (mapper[Y] && mapper[Y][M]) {
        labs.forEach((lab, idx) => {
          const row = 26 + idx;
          const count = mapper[Y][M][lab];
          sheet.getRange(row, col).setValue(count);
        });
      }
    }
  }

}


