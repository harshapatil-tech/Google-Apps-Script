/**
 * Updates the “Tutor_Names” sheet by inserting any new tutor names
 * (and their “(Instant)” variants) into the proper X/Y-delimited blocks.
 */
class SpreadsheetUpdater {
  /**
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
   * @param {{ add: string[], delete: string[] }} changes
   */
  constructor(spreadsheet, changes) {
    this.ss        = spreadsheet;
    this.sheet     = this.ss.getSheetByName("Tutor_Names");
    this.additions = changes.add || [];
    this.deletions = changes.delete || [];

  }

  /**
   * Run the update in either “add” or “delete” mode.
   * @param {"add"|"delete"|"both"} action
   */
  updateTutorNamesSheet(action = "both") {
    // 1) fetch headers + all rows
    const [headers, rows] = CentralLibrary.get_Data_Indices_From_Sheet(this.sheet);
    const rawNames  = this._getColumnValues(headers, rows, "Name of Tutors");
    const sheetAdds = this._getColumnValues(headers, rows, "Suggestion_Addition");

    // 2) locate X/Y markers
    const xPos = rawNames.indexOf("X");
    const yPos = rawNames.indexOf("Y");
    if (xPos < 0 || yPos < 0 || xPos > yPos) {
      throw new Error("X/Y markers not found or out of order.");
    }

    // 3) split blocks
    let normalBlock  = rawNames.slice(0, xPos);
    let instantBlock = rawNames.slice(xPos + 1, yPos);

    // 4) handle deletion
    if ((action === "delete" || action === "both") && this.deletions.length) {
      normalBlock  = normalBlock.filter(n => !this.deletions.includes(n));
      instantBlock = instantBlock.filter(n => {
        const base = n.replace(/ \(I\)$/, "");
        return !this.deletions.includes(base);
      });
    }

    // 5) handle additions
    let toAdd = [];
    if ((action === "add" || action === "both") && this.additions.length) {
      toAdd = this.additions
        .filter(n => sheetAdds.includes(n) && !normalBlock.includes(n));
      normalBlock  = normalBlock.concat(toAdd);
      instantBlock = instantBlock.concat(toAdd.map(n => `${n} (I)`));
    }

    // if nothing to do, exit
    if (!toAdd.length && !this.deletions.length) return;

    // 6) sort blocks
    const cmp = (a, b) => a.localeCompare(b);
    normalBlock.sort(cmp);
    instantBlock.sort(cmp);

    // 7) reassemble with markers
    const updatedNames = [
      ...normalBlock,
      "X",
      ...instantBlock,
      "Y"
    ];

    // 8) write back
    const output = updatedNames.map((nm, i) => [i + 1, nm]);
    const lastRow = this.sheet.getLastRow();
    if (lastRow >= 2) {
      this.sheet.getRange(2, 1, lastRow - 1, 2).clearContent();
    }
    this.sheet.getRange(2, 1, output.length, 2).setValues(output);
  }

  _getColumnValues(headers, rows, headerName) {
    const idx = headers[headerName];
    if (idx === undefined) throw new Error(`Header "${headerName}" not found.`);
    return rows.map(r => r[idx]).filter(v => v);
  }
}






