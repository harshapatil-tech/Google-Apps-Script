/**
 * Constants – replace with your own IDs if different.
 */
const TEMPLATE_ID     = '13GQ_U1QOsB1JfP3Kjw-Na1qjFSASyEdKXPIWkaALKDQ';
const FOLDER_ID       = '1cjIsSoXNcAaYqcloA9G7n4zUJa36ZjPr';
const INPUT_SHEET_ID  = '1XH5eGI4thuz-yAsHDeVbI90KFX2AWee3w556yUcGe6I';


/**
 * Entry point: Create a new invoice copy, then copy dates and account data.
 */ 
function createInvoice() {
  // 1) Determine next invoice name
  const nextName = getNextInvoiceName();
  
  // 2) Copy the template into the target folder
  const newFile = copyTemplate(nextName);
  const newSpreadsheet = SpreadsheetApp.openById(newFile.getId());
  
  // 3) Open the input spreadsheet
  const inputSpreadsheet = SpreadsheetApp.openById(INPUT_SHEET_ID);
  
  // 4) For each tab mapping, copy dates and account blocks; collect failures
  const failures = [];
  const tabMappings = getTabMappings();
  
  for (let i = 0; i < tabMappings.length; i++) {
    const { inputName, outputName } = tabMappings[i];
    const inputSheet  = inputSpreadsheet.getSheetByName(inputName);
    const outputSheet = newSpreadsheet.getSheetByName(outputName);
  
    if (!inputSheet || !outputSheet) {
      // If either sheet is missing, record as failure
      failures.push(`${inputName} (entire tab missing in input or output)`);
      continue;
    }
    console.log(inputName)
    if (inputSheet.getLastRow() === 0){
      continue;
    }
    // 3a) Copy the 14 dates into both tables (Table 1 and Table 2) of the output sheet
    copyDatesForTab(inputSheet, outputSheet);
    
    // 3b) Copy account‐by‐account data for Table 1 and, if present, Table 2
    copyAccountsForTab(inputSheet, outputSheet, failures);

    addExtraAccountsForTab(inputSheet, outputSheet);

  }
  
  // --- At the end of createInvoice(), after all tabs have been copied ---
  // (Assumes that every subject tab has now already had its dates written into A4:A17)
  // Pick any one of those tabs (e.g. “Calculus”) to read the 14‐day range from column A:
  const refSheet = newSpreadsheet.getSheetByName('Calculus');
  if (refSheet) {
    // Read the first and last of the 14 dates:
    const startDate = refSheet.getRange(4, 1).getValue();  // A4
    const endDate   = refSheet.getRange(17, 1).getValue(); // A17

    // Format as “MMM_dd” in the spreadsheet’s time zone:
    const startS = Utilities.formatDate(startDate,'GMT',"MMM_dd");
    const endS   = Utilities.formatDate(endDate,'GMT', 'MMM_dd');

    // Now write into the merged cell C6 (preserving its formatting):
    const summarySheet = newSpreadsheet.getSheetByName('Summary');
    if (summarySheet) {
      summarySheet.getRange('C6').setValue(`Summary_${startS}_${endS}`);
    }
  }

  // 5) After all tabs are done, show a single dialog listing any failures
  if (failures.length > 0) {
    showFailureDialog(failures);
  } else {
    SpreadsheetApp.getUi().alert('All data copied successfully.');
  }

  return [ nextName, newFile.getId() ];
}


/**
 * 1) Scan the target folder for existing files named "####_BF_Invoice" and return the next name.
 *    Format: zero-padded 4 digits + '_BF_Invoice' (e.g. '0124_BF_Invoice').
 */
function getNextInvoiceName() {
  
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files  = folder.getFiles();
  console.log("Execute")
  let maxNum = 0;
  while (files.hasNext()) {
    const file     = files.next();
    const fileName = file.getName();
    const match    = fileName.match(/^(\d{4})_BF_Invoice$/);
    if (match) {
      const num = parseInt(match[1], 10);
      if (num > maxNum) maxNum = num;
    }
  }
  
  const nextNum = maxNum + 1;
  // Zero-pad to 4 digits:
  const nextStr = ('000' + nextNum).slice(-4);
  return `${nextStr}_BF_Invoice`;
}


/**
 * 2) Copy the template file into the folder and rename it to newName.
 *    Returns the newly created Drive file.
 */
function copyTemplate(newName) {
  const templateFile = DriveApp.getFileById(TEMPLATE_ID);
  const targetFolder = DriveApp.getFolderById(FOLDER_ID);
  const copyFile     = templateFile.makeCopy(newName, targetFolder);
  DriveApp.getFileById(copyFile.getId()).moveTo(targetFolder);
  return copyFile;
}


/**
 * Returns an array of objects defining input→output tab names.
 */
function getTabMappings() {
  return [
    { inputName: 'Calculus',         outputName: 'Calculus' },
    { inputName: 'Statistics',       outputName: 'Statistics' },
    { inputName: 'Physics',          outputName: 'Physics' },
    { inputName: 'Chemistry',        outputName: 'Chemistry' },
    { inputName: 'Biology',          outputName: 'Biology' },
    { inputName: 'Intro Accounting', outputName: 'Accounting' },
    { inputName: 'Economics',        outputName: 'Economics' },
    { inputName: 'Finance',          outputName: 'Finance' },
    { inputName: 'Engineering',      outputName: 'Engineering' },
    { inputName: 'Computer Science', outputName: 'Computer Science' },
    { inputName: 'Adobe&IT',         outputName: 'Adobe&IT' },
    { inputName: 'SEO',              outputName: 'SEO' },
    { inputName: 'Writing_Lab',      outputName: 'Writing Lab (Asynchronous)' },
    { inputName: 'English',          outputName: 'Writing (Live Tutoring)' }
  ];
}


/**
 * Normalize a JavaScript Date to exactly local‐midnight
 * in the given time zone (e.g. 'Asia/Kolkata'), so it
 * won’t “flip back” a day when displayed.
 */
function normalizeDateToLocal(dateObj, timeZone) {
  const ymd = Utilities.formatDate(dateObj, timeZone, 'yyyy-MM-dd');
  return new Date(ymd + 'T00:00:00');
}


/**
 * 3a) Copy the 14 dates from inputSheet into every table in outputSheet,
 *     but first normalize each Date to local‐midnight. (Asia/Kolkata)
 *
 *     - Table 1: input rows 4–17, col A → output rows 4–17, col A
 *     - Table 2 (if present): input rows 30–43, col A → output rows 36–49, col A
 */
function copyDatesForTab(inputSheet, outputSheet) {
  const timeZone = 'Asia/Kolkata'; // ensure this matches your spreadsheet’s time zone

  // ==== TABLE 1 (read from rows 4–17, col A) ====
  const rawDates1 = inputSheet.getRange(4, 1, 14, 1).getValues();
  const normDates1 = rawDates1.map(row => {
    const d = row[0];
    if (d instanceof Date && !isNaN(d)) {
      // Normalize to midnight IST…
      const localMidnight = normalizeDateToLocal(d, 'Asia/Kolkata');
      // …then add exactly 1 day:
      localMidnight.setDate(localMidnight.getDate() + 1);
      return [ localMidnight ];
    } else {
      return [ d ];
    }
  });

  // Paste normalized dates into output rows 4–17, col A
  outputSheet.getRange(4, 1, 14, 1).setValues(normDates1);

  // --- after you’ve already called setValues(normDates1) on rows 4–17 ---

  // ==== TABLE 2 (if it exists) ====
  const headerRow2Vals = inputSheet.getRange(28, 1, 1, inputSheet.getLastColumn()).getValues()[0];
  const hasTable2 = headerRow2Vals.some(cell => cell !== '' && cell !== null);
  if (hasTable2) {
    const rawDates2 = inputSheet.getRange(30, 1, 14, 1).getValues();
    const normDates2 = rawDates2.map(row => {
      const d = row[0];
      if (d instanceof Date && !isNaN(d)) {
        const localMidnight = normalizeDateToLocal(d, 'Asia/Kolkata');
        localMidnight.setDate(localMidnight.getDate() + 1);
        return [ localMidnight ];
      } else {
        return [ d ];
      }
    });

    // Paste normalized dates into output rows 36–49, col A
    outputSheet.getRange(36, 1, 14, 1).setValues(normDates2);
  }
}


/**
 * 3b) Copy account‐by‐account data from inputSheet to outputSheet.
 *     If an account in input has no exact match (case‐insensitive) in output header, record failure.
 *
 *   We handle two tables:
 *    • Table 1 header row at input row 2 / output row 2; data from row 4 / row 4
 *    • Table 2 header row at input row 28 / output row 34; data from row 30 / row 36
 *
 *   For “Writing_Lab” (output “Writing Lab (Asynchronous)”), each account uses 1 column.
 *   For all other tabs, each account uses 3 columns.
 */
function copyAccountsForTab(inputSheet, outputSheet, failures) {
  // ===== TABLE 1 =====
  const inHeaderRow1  = 2;
  const outHeaderRow1 = 2;
  const inDataStart1  = 4;
  const outDataStart1 = 4;
  copySingleTable(
    inputSheet, outputSheet,
    inHeaderRow1, outHeaderRow1,
    inDataStart1, outDataStart1,
    1, failures
  );
  
  // ===== TABLE 2 =====
  const inHeaderRow2  = 28;
  const outHeaderRow2 = 34;
  const inDataStart2  = 30;
  const outDataStart2 = 36;
  
  // Only attempt if input row 28 has any non‐blank cells
  const headerRow2Vals = inputSheet.getRange(inHeaderRow2, 1, 1, inputSheet.getLastColumn()).getValues()[0];
  const hasTable2 = headerRow2Vals.some(cell => cell !== '' && cell !== null);
  if (hasTable2) {
    copySingleTable(
      inputSheet, outputSheet,
      inHeaderRow2, outHeaderRow2,
      inDataStart2, outDataStart2,
      2, failures
    );
  }
}


/**
 * Helper: copy one table (tableNo = 1 or 2) from input→output.
 *   - inHeaderRow: row in inputSheet with account header cells
 *   - outHeaderRow: row in outputSheet with account header cells
 *   - inDataStart: first row of actual data (14×blockWidth) in input
 *   - outDataStart: first row of actual data (14×blockWidth) in output
 *   - tableNo: 1 or 2 (for logging failures)
 *   - failures: array to push failure messages into
 *
 *   We skip column A (dates), so we begin scanning at column 2.
 *   We detect “Writing_Lab” → set blockWidth = 1; otherwise blockWidth = 3.
 *   Matching is case‐insensitive: both input and output headers are lowercased.
 */
function copySingleTable(inputSheet, outputSheet,
                         inHeaderRow, outHeaderRow,
                         inDataStart, outDataStart,
                         tableNo, failures) {
  // Determine blockWidth: 1 for Writing_Lab, otherwise 3
  const isWritingLab = inputSheet.getName() === 'Writing_Lab';
  const blockWidth   = isWritingLab ? 1 : 3;
  
  // Number of columns in input header (we scan columns 2..inLastCol)
  const inLastCol  = inputSheet.getLastColumn();
  // Number of columns in output header (we read the full width)
  const outLastCol = outputSheet.getLastColumn();
  
  // Read input header row (row inHeaderRow, columns 1..inLastCol)
  const inHeaderVals  = inputSheet.getRange(inHeaderRow, 1, 1, inLastCol).getValues()[0];
  // Read output header row (row outHeaderRow, columns 1..outLastCol)
  const outHeaderVals = outputSheet.getRange(outHeaderRow, 1, 1, outLastCol).getValues()[0];
  
  // Build an array of lowercase output headers for case‐insensitive matching:
  const outHeaderLower = outHeaderVals.map(cell => String(cell || '').trim().toLowerCase());
  
  // Now iterate across each column in the input header row, starting at column 2
  for (let c = 2; c <= inLastCol; c += blockWidth) {
    const rawAcct = inHeaderVals[c - 1];
    if (rawAcct !== '' && rawAcct !== null) {
      // Normalize the input account to lowercase for matching:
      const acctLower = String(rawAcct).trim().toLowerCase();
      
      // Search for that lowercase acct in the lowercase output header row
      const outIndex = outHeaderLower.indexOf(acctLower);
      if (outIndex !== -1) {
        // Found it! Zero‐based outIndex, so actual column = outIndex + 1
        const outBlockStart = outIndex + 1;
        
        // Copy the 14×blockWidth block from input → output
        const inRange  = inputSheet.getRange(inDataStart,  c, 14, blockWidth);
        const values   = inRange.getValues();
        const outRange = outputSheet.getRange(outDataStart, outBlockStart, 14, blockWidth);
        outRange.setValues(values);
      } else {
        // Not found → record a failure
        failures.push(`${rawAcct} | ${inputSheet.getName()} | Table ${tableNo}`);
      }
    }
    // Next iteration: advance by blockWidth (each account’s data spans blockWidth columns)
  }
}


/**
 * 5) If any account blocks failed, show a dialog with all failures listed.
 */
function showFailureDialog(failures) {
  const ui = SpreadsheetApp.getUi();
  let message = 'Unable to paste data for the following accounts:\n';
  failures.forEach((line, idx) => {
    message += `${idx + 1}) ${line}\n`;
  });
  ui.alert(message);
}





/**
 * 3c) After copying all matching accounts, add any extra ones
 *     before the “Total” column in the output sheet.
 */
function addExtraAccountsForTab(inputSheet, outputSheet) {
  // define both Table-1 and Table-2 row mappings
  const tables = [
    { inHeaderRow: 2,  outHeaderRow: 2,  inDataStart: 4,  outDataStart: 4  },
    { inHeaderRow: 28, outHeaderRow: 34, inDataStart: 30, outDataStart: 36 }
  ];
  const inLastCol = inputSheet.getLastColumn();

  tables.forEach(({inHeaderRow, outHeaderRow, inDataStart, outDataStart}) => {
    // read entire headers
    const inHeaderVals  = inputSheet.getRange(inHeaderRow, 1, 1, inLastCol).getValues()[0];
    const outLastCol    = outputSheet.getLastColumn();
    let outHeaderVals   = outputSheet.getRange(outHeaderRow, 1, 1, outLastCol).getValues()[0];
    let outHeaderLower  = outHeaderVals.map(c => String(c||'').trim().toLowerCase());

    // find where “Total” lives (zero-based)
    let totalIdx = outHeaderLower.indexOf('total');
    if (totalIdx < 0) return;  // no Total? skip

    // determine how many columns each account block uses
    const isWritingLab = inputSheet.getName() === 'Writing_Lab';
    const blockWidth   = isWritingLab ? 1 : 3;

    // scan each input‐account block
    for (let c = 2; c <= inLastCol; c += blockWidth) {
      const rawAcct   = inHeaderVals[c - 1];
      const acctLower = String(rawAcct||'').trim().toLowerCase();
      // skip blanks and those already copied
      if (!rawAcct || outHeaderLower.includes(acctLower)) continue;

      // *** we have an extra account → insert its block ***

      // re-read the output header row in case it grew
      outHeaderVals  = outputSheet.getRange(outHeaderRow, 1, 1, outputSheet.getLastColumn()).getValues()[0];
      outHeaderLower = outHeaderVals.map(c => String(c||'').trim().toLowerCase());
      totalIdx       = outHeaderLower.indexOf('total');
      const insertPos = totalIdx + 1;       // 1-based column

      // 1) insert empty columns before “Total”
      outputSheet.insertColumnsBefore(insertPos, blockWidth);

      // 2) copy the header labels from input → output
      const headerBlock = inputSheet
        .getRange(inHeaderRow, c, 1, blockWidth)
        .getValues();
      outputSheet
        .getRange(outHeaderRow, insertPos, 1, blockWidth)
        .setValues(headerBlock);

      // 3) copy the 14×blockWidth data block
      const dataBlock = inputSheet
        .getRange(inDataStart, c, 14, blockWidth)
        .getValues();
      outputSheet
        .getRange(outDataStart, insertPos, 14, blockWidth)
        .setValues(dataBlock);
    }
  });
}









