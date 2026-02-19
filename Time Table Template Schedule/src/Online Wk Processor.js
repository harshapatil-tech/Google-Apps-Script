
//old code
class OnlineWkProcessor {

  static process(ss, tutorSheet, date) {

    const scheduleSheets = ss.getSheets().filter(sh => sh.getName().includes('Online_Wk'));

    //new code
    // const scheduleSheets = ss.getSheets()
    //   .filter(sh => sh.getName().includes('Online_Wk'))
    //   .sort((a, b) => {
    //     const wa = parseInt(a.getName().match(/\d+/)[0], 10);
    //     const wb = parseInt(b.getName().match(/\d+/)[0], 10);
    //     return wa - wb;
    //   });

    //console.log("Sheet order after sort:",scheduleSheets.map(s => s.getName()));

    if (!scheduleSheets.length) return;

    // 1) compute next month/year
    const { nextMonthNum, nextYearFull } = this._getNextMonthAndYear(date);

    // 2) find first week Monday
    const firstWeekMonday = this._computeFirstWeekMonday(nextYearFull, nextMonthNum);

    // 3) build tutor-dropdown rule
    const validationRule = this._buildValidationRule(tutorSheet);

    // 4) process each week sheet
    const offsetMap = { Mon: 0, Tue: 1, Wed: 2, Thu: 3, Fri: 4, Sat: 5, Sun: 6 };
    console.log("offsetmap",offsetMap);
    scheduleSheets.forEach((sheet, weekIndex) => {
     // sheet.getDataRange().clearDataValidations();
      // this._resetHeaderMerges(sheet);
      // this._freezeColumns(sheet, 4);
      // this._applyHeaderMergesAndCopy(sheet);

     // this._setTitle(sheet, nextYearFull);
      this._setTitle(sheet, nextYearFull, date);
     
      
      const weekMonday = new Date(firstWeekMonday);
      weekMonday.setDate(firstWeekMonday.getDate() + 7 * weekIndex);
      // console.log("Sheet:", sheet.getName());
      // console.log("Week index:", weekIndex);
      // console.log("Week Monday:", weekMonday);

      this._stampDates(sheet, weekMonday, nextMonthNum, offsetMap);
      this._fillDropdownBlocks(sheet, validationRule);
      this._applyFormulas(sheet);
    });
  }

  // Helpers ------------------------------------------------
  static _resetHeaderMerges(sheet) {
    [1, 2, 3, 111, 112].forEach(r => sheet.getRange(`B${r}:L${r}`).breakApart());
  }

  static _applyHeaderMergesAndCopy(sheet) {
    const configs = [
      { mergeRow: 1, valueRow: 1 },
      { mergeRow: 2, valueRow: 2 },
      { mergeRow: 3, valueRow: 3 },
      { mergeRow: 111, valueRow: 111 },
      { mergeRow: 112, valueRow: 112 }
    ];
    configs.forEach(({ mergeRow, valueRow }) => {
      const range = sheet.getRange(`B${valueRow}`);
      const val = range.getValue();
      sheet.getRange(`B${valueRow}:D${valueRow}`).clear();
      sheet.getRange(`E${mergeRow}:L${mergeRow}`).merge().setValue(val);
    });
  }

  static _getNextMonthAndYear(date) {
    const nextMonthNum = DateUtil.getNextMonthNumber(date);
    const nextYearFull = DateUtil.getNextYear(date);
    console.log(nextMonthNum, nextYearFull);
    console.log("Input date:", date);
    console.log("Computed month:", nextMonthNum);
    console.log("Computed year:", nextYearFull);
    return { nextMonthNum, nextYearFull }
  }
  
  // static _getNextMonthAndYear(date) {
  //   const base = DateUtil.baseDate();
  //   const currMonth = base.getMonth() + 1;
  //   const currYear  = base.getFullYear();
  //   const nextMonthNum = currMonth === 12 ? 1 : currMonth + 1;
  //   const nextYearFull = currMonth === 12 ? currYear + 1 : currYear;
  //   return { nextMonthNum, nextYearFull };
  // }

  // static _setTitle(sheet, yearFull) {
  //   const nextName = DateUtil.getNextMonthName("long", this.date);
  //   const suffix = String(yearFull).slice(-2);
  //   sheet.getRange("E1:L1").setValue(`Time table for ${nextName}'${suffix}`);
  // }
 
  //new code
  static _setTitle(sheet, yearFull, baseDate) {

  const nextName = DateUtil.getNextMonthName("long", baseDate);
  
  sheet.setFrozenColumns(0);
  sheet.getRange("B1:L1").breakApart();

  sheet.getRange("B1:L1")
    .merge()
    .setValue(`Time table for ${nextName}'${yearFull}`)
    .setHorizontalAlignment("center")
    .setFontWeight("bold");
}

    
  static _computeFirstWeekMonday(year, month) {
    const firstOfMonth = new Date(year, month - 1, 1);
    const weekday = firstOfMonth.getDay();
    const daysBack = (weekday - 1 + 7) % 7;
    const firstMonday = new Date(firstOfMonth);
    firstMonday.setDate(1 - daysBack);
    return firstMonday;
  }

//old code
  static _buildValidationRule(sheet) {
    const tutorRange = sheet.getRange("B2:B");
    return SpreadsheetApp.newDataValidation()
      .requireValueInRange(tutorRange, true)
      .build();
  }


  //new code
  // static _buildValidationRule(sheet) {

  //   const lastRow = sheet.getLastRow();
  //   const tutorRange = sheet.getRange(2, 2, lastRow - 1, 1); // B2:B(lastRow)

  //   return SpreadsheetApp.newDataValidation()
  //     .requireValueInRange(tutorRange, true)
  //     .build();
  // }


  static _applyFormulas(sheet) {
    const formula = '=COUNTIFS(RC[-7]:RC[-1],"<>"&"X",RC[-7]:RC[-1],"<>")*RC[-9]';
    sheet.getRange(7, 12, 100, 1).setFormulaR1C1(formula);
    sheet.getRange(116, 12, 100, 1).setFormulaR1C1(formula);
  }

 
  static _freezeColumns(sheet, numCols) {
    sheet.setFrozenColumns(numCols);
  }

  //old code
  static _stampDates(sheet, weekMonday, targetMonth, offsetMap) {
 
    const lastCol = 11;
    const startCol = 5
    const headers = sheet.getRange(5, startCol, 1, lastCol - startCol + 1).getValues()[0];
    const weekStart = weekMonday.getMonth() + 1;

    headers.forEach((day, idx) => {
      if (!(day in offsetMap)) return;
      const d = new Date(weekMonday);
      d.setDate(weekMonday.getDate() + offsetMap[day]);
      if (d.getMonth() + 1 === targetMonth || weekStart === targetMonth) {
        const fmt = Utilities.formatDate(d, Session.getScriptTimeZone(), "dd-MMM-yy");
        sheet.getRange(4, idx + startCol).setValue(fmt);
      }
    });
  }


   //old code
  static _fillDropdownBlocks(sheet, validationRule) {
    const startCol = 5, numCols = 7;
    //new
    //const startCol = 4, numCols = 7;
    const blocks = [{ row: 7, rows: 100 }, { row: 116, rows: 100 }];
    const dayHdr = sheet.getRange(4, startCol, 1, numCols).getValues()[0];
    
    blocks.forEach(b => {

      const vals = sheet.getRange(b.row, startCol, b.rows, numCols).getValues();
      const grid = vals.map((row, i) =>
        row.map((cell, j) => cell ? "X" : "")       // && dayHdr[j]
      );
      //new code
      // const grid = vals.map(row =>
      //   row.map(cell => cell || "")
      // );

      const range = sheet.getRange(b.row, startCol, b.rows, numCols);
      if (sheet.getName() === "Online_Wk_1")
        range.setValues(grid)
      range
        .setDataValidation(validationRule)
        .setBorder(true, true, true, true, true, true)
        .setBackground("white")
        .setFontWeight("normal");
    });
  }


  //new code
  // static _fillDropdownBlocks(sheet, validationRule) {

  //   const startCol = 5;
  //   const numCols = 7;

  //   const blocks = [
  //     { row: 7, rows: 100 },
  //     { row: 116, rows: 100 }
  //   ];

  //   blocks.forEach(b => {

  //     const range = sheet.getRange(b.row, startCol, b.rows, numCols);

  //     range.clearContent();

  //     range
  //       .setDataValidation(validationRule)
  //       .setBorder(true, true, true, true, true, true)
  //       .setBackground("white")
  //       .setFontWeight("normal");

  //   });
  // }

}








