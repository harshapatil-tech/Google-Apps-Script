class SummaryUpdater {

  constructor (ss) {
    this.ss       = ss;
    this.sheet    = ss.getSheetByName('Summary');
    this.tutorSheetName = "Tutor_Names";
    this.headers  = [
      'Summary_Day Shift_Online+Extended',
      'Summary_Night Shift_Online+Extended',
      'Summary_Day+Night Shift_Online+Extended'
    ];
    this.curentTutorNames = this._getCleanTutorNames();
  }

  addOrDeleteColumn() {
    let firstTwoColumns; firstTwoColumns = this.sheet.getRange(1, 1, this.sheet.getLastRow(), 2).getValues();
    

    let firstBlockStart, firstBlockEnd;
    
    firstTwoColumns = this.sheet.getRange(1, 1, this.sheet.getLastRow(), 2).getValues()
    firstBlockStart = firstTwoColumns.findIndex(ele => ele[0] === this.headers[0]) + 2;
    firstBlockEnd = firstTwoColumns.findIndex(ele => ele[0] === this.headers[1]) - 3;
    this._toDelete(this.curentTutorNames, firstTwoColumns.slice(firstBlockStart, firstBlockEnd).map(ele => ele[1]), firstBlockStart);

    firstTwoColumns = this.sheet.getRange(1, 1, this.sheet.getLastRow(), 2).getValues()
    firstBlockStart = firstTwoColumns.findIndex(ele => ele[0] === this.headers[0]) + 2;
    firstBlockEnd = firstTwoColumns.findIndex(ele => ele[0] === this.headers[1]) - 3;
    this._toAdd(this.curentTutorNames, firstTwoColumns.slice(firstBlockStart, firstBlockEnd).map(ele => ele[1]), firstBlockStart, "firstBlock");

    


    let secondBlockStart, secondBlockEnd;
    
    firstTwoColumns = this.sheet.getRange(1, 1, this.sheet.getLastRow(), 2).getValues()
    secondBlockStart = firstTwoColumns.findIndex(ele => ele[0] === this.headers[1]) + 2;
    secondBlockEnd = firstTwoColumns.findIndex(ele => ele[0] === this.headers[2]) - 3;
    this._toDelete(this.curentTutorNames, firstTwoColumns.slice(secondBlockStart, secondBlockEnd).map(ele => ele[1]), secondBlockStart);

    firstTwoColumns = this.sheet.getRange(1, 1, this.sheet.getLastRow(), 2).getValues()
    secondBlockStart = firstTwoColumns.findIndex(ele => ele[0] === this.headers[1]) + 2;
    secondBlockEnd = firstTwoColumns.findIndex(ele => ele[0] === this.headers[2]) - 3;
    this._toAdd(this.curentTutorNames, firstTwoColumns.slice(secondBlockStart, secondBlockEnd).map(ele => ele[1]), secondBlockStart, "secondBlock");



    let thirdBlockStart, thirdBlockEnd;
    
    firstTwoColumns = this.sheet.getRange(1, 1, this.sheet.getLastRow(), 2).getValues()
    thirdBlockStart = firstTwoColumns.findIndex(ele => ele[0] === this.headers[2]) + 2;
    thirdBlockEnd = this.sheet.getLastRow()-1;
    this._toDelete(this.curentTutorNames, firstTwoColumns.slice(thirdBlockStart, thirdBlockEnd).map(ele => ele[1]), thirdBlockStart);

    firstTwoColumns = this.sheet.getRange(1, 1, this.sheet.getLastRow(), 2).getValues()
    thirdBlockStart = firstTwoColumns.findIndex(ele => ele[0] === this.headers[2]) + 2;
    thirdBlockEnd = this.sheet.getLastRow()-1;
    this._toAdd(this.curentTutorNames, firstTwoColumns.slice(thirdBlockStart, thirdBlockEnd).map(ele => ele[1]), thirdBlockStart);

    
    
     // Update Sr nos
    firstTwoColumns = this.sheet.getRange(1, 1, this.sheet.getLastRow(), 2).getValues()
    firstBlockStart = firstTwoColumns.findIndex(ele => ele[0] === this.headers[0]) + 2+1;
    firstBlockEnd = firstTwoColumns.findIndex(ele => ele[0] === this.headers[1]) - 3;
    secondBlockStart = firstTwoColumns.findIndex(ele => ele[0] === this.headers[1]) + 2+1;
    secondBlockEnd = firstTwoColumns.findIndex(ele => ele[0] === this.headers[2]) - 3;
    this._updateSrNosInBlock(firstTwoColumns.findIndex(ele => ele[0] === this.headers[2]) + 3, this.sheet.getLastRow()-1);
    this._updateSrNosInBlock(secondBlockStart, secondBlockEnd);
    this._updateSrNosInBlock(firstBlockStart, firstBlockEnd);

    this._fillInTotalFormulas(firstTwoColumns.findIndex(ele => ele[0] === this.headers[2]) + 3, this.sheet.getLastRow()-1, firstBlockStart, secondBlockStart)
  }

  _toAdd(list1, list2, blockStart, blockIdx="thirdBlock") {
    console.log("List1", list1);
    console.log("List2", list2);
    const toAdd    = [];
    // 2. Find missing in list2 → need to be added at the index they appear in list1
    list1.forEach((name, idx) => {
      if (!list2.includes(name)) {
        toAdd.push({ name, index: idx + blockStart+1 });
      }
    });
    
    toAdd.forEach(row => {
      const previousSrNo = this.sheet.getRange(row.index-1, 1).getValue();
      this.sheet.insertRowAfter(row.index-1)
      this.sheet.getRange(row.index, 1).setValue(previousSrNo+1)
      this.sheet.getRange(row.index, 2).setValue(row.name);

      if (blockIdx == "thirdBlock") {
        // 2) now copy all of the formulas from the row above → the new row
        const startCol = 3;               // C
        const endCol   = 33;              // AG
        const numCols  = endCol - startCol + 1;
        const aboveRow = row.index - 1;
        const newRow   = row.index;

        this.sheet
          .getRange(aboveRow, startCol, 1, numCols)
          .copyTo(
            this.sheet.getRange(newRow, startCol, 1, numCols),
            SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
            false
          );
      }
    });
  }

  _toDelete(list1, list2, blockStart) {
    const toRemove = [];

    const occurrences = {};
    list2.forEach((name, idx) => {
      if (!name) return;
      if (!occurrences[name]) occurrences[name] = [];
      occurrences[name].push(idx);
    });
    Object.keys(occurrences).forEach(name => {
      const indices = occurrences[name]; 
      if (indices.length > 1) {
        // keep the first occurrence; delete every index after that
        indices.slice(1).forEach(relIdx => {
          // absolute row = (relIdx + blockStart + 1), matching your existing formula
          toRemove.push({ name, index: relIdx + blockStart + 1 });
        });
      }
    });


    // 1. Find extras in list2 → need to be removed at their current indices
    list2.forEach((name, idx) => {
      if (!list1.includes(name)) {
        toRemove.push({ name, index: idx + blockStart+1 });
      }
    });
    // // 3. Sort removals descending so deleting earlier entries doesn't shift later indices
    toRemove.sort((a, b) => b.index - a.index);
    // Delete now
    toRemove.forEach(row => {
      this.sheet.deleteRow(row.index)
    });

  }


  _fillInTotalFormulas(startRow, endRow, firstBlockStart, secondBlockStart) {
    const rowCount = endRow - startRow + 1;
    if (rowCount <= 0) return;

    // Total number of columns we need (from column C = 3) to the sheet's last column:
    const lastCol = this.sheet.getLastColumn();
    const numCols  = lastCol - 3 + 1;  // e.g. if lastCol=33 (AG), then numCols=31 (C..AG)

    // Build a 2D array of formulas: one row per data row, one entry per column C..lastCol
    const formulas = [];
    for (let i = 0; i < rowCount; i++) {
      const formulasRow = [];
      // Compute the two “source” rows in block 1 and block 2 for this offset
      const sourceRow1 = firstBlockStart  + i;
      const sourceRow2 = secondBlockStart + i;

      for (let col = 3; col <= lastCol; col++) {
        const colLetter = Helper.columnToLetter(col);
        // e.g. "=C4+C51", "=D5+D52", etc.
        formulasRow.push(
          `=${colLetter}${sourceRow1}+${colLetter}${sourceRow2}`
        );
      }
      formulas.push(formulasRow);
    }

    // Write all formulas in one shot, starting at (startRow, col=3), spanning rowCount×numCols
    this.sheet
      .getRange(startRow, 3, rowCount, numCols)
      .setFormulas(formulas);
  }





  _updateSrNosInBlock(startRow, endRow, firstBlockStart, secondBlockStart) {
    const rowCount = endRow - startRow+1;

    if (rowCount <= 0) return;
    // build [[1],[2],…]
    const seq = Array.from({length: rowCount}, (_, i) => [i + 1]);
  
    this.sheet
        .getRange(startRow, 1, rowCount, 1)
        .setValues(seq);

    const totalCol = this.sheet.getLastColumn();

    // Precompute the column letters once:
    const colStart = Helper.columnToLetter(3);   // yields "C"
    const colEnd   = Helper.columnToLetter(totalCol-1);  // yields "AG"

    // Build a 2D array of formula‐strings, one row per data‐row
    const formulas = [];
    for (let r = startRow; r <= endRow; r++) {
      // SUM from C to AG on the same row:
      formulas.push([ `=SUM(${colStart}${r}:${colEnd}${r})` ]);
    }
    // Now write them all at once into column “totalCol”, rows startRow..endRow:
    this.sheet
        .getRange(startRow, totalCol, rowCount, 1)
        .setFormulas(formulas);
    // this.sheet.getRange(endRow+1, 3, 1, 32).setFormula(`=SUM(C${startRow}:C${endRow})`);
  }


  _getCleanTutorNames() {
    const sheet = this.ss.getSheetByName(this.tutorSheetName);
    const [headers, rows] = CentralLibrary.get_Data_Indices_From_Sheet(sheet);
    const rawNames = rows
      .map(r => r[headers["Name of Tutors"]])
      .filter(Boolean)
      // .filter(v => v && v !== "X" && v !== "Y");
    return rawNames;
  }
}



