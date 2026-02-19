class ExtendedUpdater {

  constructor (ss) {
    this.ss       = ss;
    this.sheet    = ss.getSheetByName('Extended');
    this.tutorSheetName = "Tutor_Names";
    this.headers  = [
      'Extended_Day Shift',
      'Extended_Night Shift',
      'Total Extended Hours (Day+Night Shift)'
    ];
    this.curentTutorNames = this._getCleanTutorNames();
  }

  addOrDeleteColumn() {
    let firstTwoColumns = this.sheet.getRange(1, 1, this.sheet.getLastRow(), 2).getValues();

    let secondBlockStart, secondBlockEnd;
    secondBlockStart = firstTwoColumns.findIndex(ele => ele[0] === this.headers[1]) + 2;
    secondBlockEnd = firstTwoColumns.findIndex(ele => ele[0] === this.headers[2]) - 3;
    console.log(secondBlockStart, secondBlockEnd)
    this._toDelete(this.curentTutorNames, firstTwoColumns.slice(secondBlockStart, secondBlockEnd).map(ele => ele[1]), secondBlockStart);
    
    firstTwoColumns = this.sheet.getRange(1, 1, this.sheet.getLastRow(), 2).getValues();
    secondBlockStart = firstTwoColumns.findIndex(ele => ele[0] === this.headers[1]) + 2;
    secondBlockEnd = firstTwoColumns.findIndex(ele => ele[0] === this.headers[2]) - 3;
    this._toAdd(this.curentTutorNames, firstTwoColumns.slice(secondBlockStart, secondBlockEnd).map(ele => ele[1]), secondBlockStart);

    let firstBlockStart, firstBlockEnd;

    firstTwoColumns = this.sheet.getRange(1, 1, this.sheet.getLastRow(), 2).getValues();
    firstBlockStart = firstTwoColumns.findIndex(ele => ele[0] === this.headers[0]) + 2;
    firstBlockEnd = firstTwoColumns.findIndex(ele => ele[0] === this.headers[1]) - 3;
    this._toDelete(this.curentTutorNames, firstTwoColumns.slice(firstBlockStart, firstBlockEnd).map(ele => ele[1]), firstBlockStart);

    firstTwoColumns = this.sheet.getRange(1, 1, this.sheet.getLastRow(), 2).getValues();
    firstBlockStart = firstTwoColumns.findIndex(ele => ele[0] === this.headers[0]) + 2;
    firstBlockEnd = firstTwoColumns.findIndex(ele => ele[0] === this.headers[1]) - 3;
    this._toAdd(this.curentTutorNames, firstTwoColumns.slice(firstBlockStart, firstBlockEnd).map(ele => ele[1]), firstBlockStart);


    firstTwoColumns = this.sheet.getRange(1, 1, this.sheet.getLastRow(), 2).getValues()
    this._updateSrNosInBlock(firstTwoColumns.findIndex(ele => ele[0] === this.headers[1]) + 2+1, firstTwoColumns.findIndex(ele => ele[0] === this.headers[2]) - 3);
    this._updateSrNosInBlock(firstTwoColumns.findIndex(ele => ele[0] === this.headers[0]) + 2+1, firstTwoColumns.findIndex(ele => ele[0] === this.headers[1]) - 3);
  }

  _toAdd(list1, list2, blockStart) {
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
    });
  }

  _toDelete(list1, list2, blockStart) {
    const toRemove = [];
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

  _updateSrNosInBlock(startRow, endRow) {
    const rowCount = endRow - startRow+1;
    if (rowCount <= 0) return;
    // build [[1],[2],…]
    const seq = Array.from({length: rowCount}, (_, i) => [i + 1]);
  
    this.sheet
        .getRange(startRow, 1, rowCount, 1)
        .setValues(seq);
    this.sheet.getRange(endRow+1, 3, 1, 32).setFormula(`=SUM(C${startRow}:C${endRow})`);
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




// function runCode() {
//   const ss = SpreadsheetApp.openById("1NmJo2DBNius1CBDKhjURxoO_PrDrsz_QNOIUMzSdyBM");
//   const y = new ExtendedUpdater(ss);
//   y.addOrDeleteColumn()
// }

