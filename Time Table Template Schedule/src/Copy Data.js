//Last cross month week copy,tutor names copy
function test_Copy() {
  const ss = SpreadsheetApp.openById("1CjY1E9Dx611R8_T3anEorWqLyT5RfzXRe-nuI1r49es");
  const copyData = new CopyData(ss);
  let x = copyData.run();
}


function setupErrorLogging() {
  ErrorLogger.init({
    projectId: '',
    logName:   'apps_script_errors',
    environment: 'production',
    // optional extras:
    batchSize:      5,
    flushIntervalSec: 60,
    sheet: { use: false },             // disable spreadsheet fallback
    alert: { use: true, to: 'automation@upthink.com' }
  });
}



class CopyData {

  constructor (id) {
    this.spreadsheet = SpreadsheetApp.openById(id);
  }

  run() {
    return {
      tutorNames: this.getTutorNames(),
      lastWeek: this.getSourceSheets(this.spreadsheet)
    }
  }

  getTutorNames () {
    try {
      const tutorNamesSheet = this.spreadsheet.getSheetByName("Tutor_Names");
      const [headers, _] = CentralLibrary.get_Data_Indices_From_Sheet(tutorNamesSheet);
      const tutorNames = tutorNamesSheet
                              .getRange(2, headers["Name of Tutors"]+1, tutorNamesSheet.getLastRow()-1, 1)
                              .getValues()
                              .map(ele => ele[0])
                              // .flatten()
                              .filter(Boolean);

      return tutorNames;
    } catch (e){
      ErrorLogger.error(e, {
      functionName: 'CopyData.getTutorNames',
        customData: { sheetName: 'Tutor_Names' }
      });
      throw e;
    }

  }

  /* @param {Spreadsheet} ss  The source Spreadsheet
   * @return {{ lastOnlineWeek: Sheet,
   *            tutorNames:    Sheet,
   *            extended:      Sheet,
   *            summary:       Sheet }}
   */
  getSourceSheets(ss) {
    try {
      let lastOnlineWeek = null;

      // Look from Online_Wk_6 down to Online_Wk_1 for the week that spans two months
      for (let i = 6; i >= 1; i--) {
        const name = `Online_Wk_${i}`;
        const sheet = ss.getSheetByName(name);
        if (!sheet) continue;

        // Grab row 4, columns 5–11 (E4:K4)
        const dates = sheet.getRange(4, 5, 1, 7).getValues()[0];
        // Filter to actual Date objects, pull out months, dedupe
        const months = dates
          .filter(cell => cell instanceof Date)
          .map(d => d.getMonth());
        const uniqueMonths = months.filter((m, idx, arr) => arr.indexOf(m) === idx);

        // If we see more than one month in that week, it’s the “last” one
        if (uniqueMonths.length > 1) {
          lastOnlineWeek = sheet;
          break;
        }
      }

      const block1 = lastOnlineWeek.getRange(7, 2, 100, 11-2+1).getValues();
      const block2 = lastOnlineWeek.getRange(116, 5, 100, 11-2+1).getValues();

      return {
        block1,
        block2
      }
    } catch (e) {
      // ErrorLogger.error(e, {
      //   functionName: 'CopyData.getSourceSheets',
      //   customData: { spreadsheetId: ss.getId() }
      //});
      // Re-throw so upstream can handle or our wrapper can catch it
      throw e;
    }
    
  }


}



class PasteData {

  constructor(id, data) {
    this.ss = SpreadsheetApp.openById(id);
    this.data = data;
  }

  run() {
    this.setTutorNames();
    this.setOnlineWk1();
  }

  setTutorNames() {
    const sheet = this.ss.getSheetByName("Tutor_Names");
    const [header, data] = CentralLibrary.get_Data_Indices_From_Sheet(sheet);
    const tutorNames = this.data.tutorNames.map(ele => [ele]);
    const numTutors = tutorNames.length;
    sheet.getRange(2, header["Name of Tutors"]+1, numTutors, 1).setValues(tutorNames);
    
    // If there are no tutors, bail out
    if (numTutors === 0) return;

    const formulas = [];
    for (let i = 0; i < numTutors; i++) {
      const row = 2 + i;
      if (row === 2) {
        // A2 = 1
        formulas.push(["=1"]);
      } else {
        // e.g. A3 = A2+1, A4 = A3+1, …
        const prevRow = row - 1;
        formulas.push([`=A${prevRow}+1`]);
      }
    }

    // Write those formulas into column A, rows 2..(2+numTutors−1)
    sheet
      .getRange(2, 1, numTutors, 1)
      .setFormulas(formulas);
  }

  setOnlineWk1(){
    const onlineWk1 = this.ss.getSheetByName("Online_Wk_1");
    onlineWk1.getRange(7, 2, 100, 11-2+1).setValues(this.data.lastWeek.block1);
    onlineWk1.getRange(116, 2, 100, 11-2+1).setValues(this.data.lastWeek.block2);
  }


}































