// function updateSheetBfromA() {
//   // ==== CONFIG ====
//   const SPREADSHEET_A_ID = "1Sjka9f3Z_w8jX6jzk8CCM1g1scZ42rRjcGR5Nh3ZSbI";   
//   const SHEET_A_NAME = "Headcount";                         
//   const SPREADSHEET_B_ID = "1yAVztZBtGYPugT62jef9jbQS5FyjnRdkWQy-dKeIwSg";   
//   const SHEET_B_NAME = "Input Sheet";                       

//   const MAX_HEADER_SCAN_ROWS = 12;

//   const KEY_NAMES_A = ['uuid', 'employee identifier', 'unique id', 'uniqueid'];
//   const KEY_NAMES_B = ['unique id', 'uniqueid', 'uuid', 'employee identifier'];

//   const normalize = v => {
//     if (v === null || v === undefined) return "";
//     return String(v).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
//   };

//   const findHeaderRowIndex = (rawValues, candidateNames) => {
//     const rowsToCheck = Math.min(MAX_HEADER_SCAN_ROWS, rawValues.length);
//     for (let i = 0; i < rowsToCheck; i++) {
//       const row = rawValues[i];
//       for (let j = 0; j < row.length; j++) {
//         if (candidateNames.indexOf(normalize(row[j])) !== -1) {
//           return i; 
//         }
//       }
//     }
//     return -1;
//   };

//   const findColIndexByCandidates = (headersRow, candidateNames) => {
//     for (let c = 0; c < headersRow.length; c++) {
//       if (candidateNames.indexOf(normalize(headersRow[c])) !== -1) return c;
//     }
//     return -1;
//   };

//   try {
//     console.log("Opening spreadsheets...");
//     const ssA = SpreadsheetApp.openById(SPREADSHEET_A_ID);
//     const ssB = SpreadsheetApp.openById(SPREADSHEET_B_ID);
//     const sheetA = ssA.getSheetByName(SHEET_A_NAME);
//     const sheetB = ssB.getSheetByName(SHEET_B_NAME);

//     if (!sheetA) {
//       console.log("ERROR: Sheet A not found. Check SPREADSHEET_A_ID and SHEET_A_NAME.");
//       return;
//     }
//     if (!sheetB) {
//       console.log("ERROR: Sheet B not found. Check SPREADSHEET_B_ID and SHEET_B_NAME.");
//       return;
//     }

//     const rawA = sheetA.getDataRange().getValues();
//     const rawB = sheetB.getDataRange().getValues();

//     if (rawA.length === 0) {
//       console.log("ERROR: Sheet A is empty.");
//       return;
//     }
//     if (rawB.length === 0) {
//       console.log("ERROR: Sheet B is empty.");
//       return;
//     }

//     const headerRowIndexA = findHeaderRowIndex(rawA, KEY_NAMES_A);
//     const headerRowIndexB = findHeaderRowIndex(rawB, KEY_NAMES_B);

//     if (headerRowIndexA === -1 || headerRowIndexB === -1) {
//       console.log("ERROR: Could not detect header row(s).");
//       console.log(" headerRowIndexA =", headerRowIndexA, " headerRowIndexB =", headerRowIndexB);
//       console.log(" Searched keys for A:", KEY_NAMES_A, "  for B:", KEY_NAMES_B);
//       console.log(" --- first few rows of A for debugging ---");
//       console.log(rawA.slice(0, Math.min(MAX_HEADER_SCAN_ROWS, rawA.length)));
//       console.log(" --- first few rows of B for debugging ---");
//       console.log(rawB.slice(0, Math.min(MAX_HEADER_SCAN_ROWS, rawB.length)));
//       return;
//     }

//     const headerRowA_1based = headerRowIndexA + 1;
//     const headerRowB_1based = headerRowIndexB + 1;

//     console.log("Detected header rows -> A row:", headerRowA_1based, " B row:", headerRowB_1based);

//     const lastRowA = sheetA.getLastRow();
//     const lastColA = sheetA.getLastColumn();
//     const blockA = sheetA.getRange(headerRowA_1based, 1, lastRowA - headerRowA_1based + 1, lastColA).getValues();
//     const headersA = blockA[0].map(h => (h === null || h === undefined) ? "" : String(h));
//     const bodyA = blockA.slice(1);

//     const lastRowB = sheetB.getLastRow();
//     const lastColB = sheetB.getLastColumn();
//     const blockB = sheetB.getRange(headerRowB_1based, 1, lastRowB - headerRowB_1based + 1, lastColB).getValues();
//     const headersB = blockB[0].map(h => (h === null || h === undefined) ? "" : String(h));
//     const bodyB = blockB.slice(1);

//     const uuidColA = findColIndexByCandidates(headersA, KEY_NAMES_A);
//     const uuidColB = findColIndexByCandidates(headersB, KEY_NAMES_B);

//     if (uuidColA === -1 || uuidColB === -1) {
//       console.log("ERROR: Could not locate UUID/Unique ID column inside detected header rows.");
//       console.log(" uuidColA =", uuidColA, " uuidColB =", uuidColB);
//       console.log(" headersA:", headersA);
//       console.log(" headersB:", headersB);
//       return;
//     }

//     console.log("UUID column indexes -> A:", uuidColA, " B:", uuidColB);


//     const mapA = {};
//     for (let i = 0; i < bodyA.length; i++) {
//       const row = bodyA[i];
//       const rawUuid = row[uuidColA];
//       const key = normalize(rawUuid);
//       if (key) {
//         mapA[key] = row;
//       }
//     }
//     console.log("Mapped rows from A (by UUID):", Object.keys(mapA).length);


//     let rowsMatched = 0;
//     let cellsUpdated = 0;
//     for (let r = 0; r < bodyB.length; r++) {
//       const rowB = bodyB[r];
//       const rawUuidB = rowB[uuidColB];
//       const keyB = normalize(rawUuidB);
//       if (!keyB) continue;

//       const rowA = mapA[keyB];
//       if (!rowA) continue; 
      

//       rowsMatched++;

//       for (let c = 0; c < headersB.length; c++) {
//         const hdrBnorm = normalize(headersB[c]);
//         if (!hdrBnorm) continue;

//         let colAindex = -1;
//         for (let ca = 0; ca < headersA.length; ca++) {
//           if (normalize(headersA[ca]) === hdrBnorm) {
//             colAindex = ca;
//             break;
//           }
//         }
//         if (colAindex !== -1) {
//           const valA = rowA[colAindex];
//           if (rowB[c] !== valA) {
//             bodyB[r][c] = valA;
//             cellsUpdated++;
//           }
//         }
//       }
//     }

//     console.log("Rows matched:", rowsMatched, "Cells updated:", cellsUpdated);

//     if (cellsUpdated > 0) {
//       const outBlock = [headersB].concat(bodyB);
//       sheetB.getRange(headerRowB_1based, 1, outBlock.length, outBlock[0].length).setValues(outBlock);
//       console.log("Wrote updates back to Sheet B starting at row", headerRowB_1based);
//     } else {
//       console.log("No changes detected — nothing written to Sheet B.");
//     }

//     console.log(" updateSheetBfromA completed.");
//   } catch (err) {
//     console.log("ERROR: Exception in updateSheetBfromA ->", err && err.message ? err.message : err);
//     if (err && err.stack) console.log(err.stack);
//   }
// }
function updateTargetFromMaster() {
  const masterId = "1Sjka9f3Z_w8jX6jzk8CCM1g1scZ42rRjcGR5Nh3ZSbI";  
  const targetId = "1yAVztZBtGYPugT62jef9jbQS5FyjnRdkWQy-dKeIwSg"; 

  const masterSheetName = "Headcount"; 
  const targetSheetName = "Input Sheet"; 

  const masterSS = SpreadsheetApp.openById(masterId);
  const targetSS = SpreadsheetApp.openById(targetId);
  const masterSheet = masterSS.getSheetByName(masterSheetName);
  const targetSheet = targetSS.getSheetByName(targetSheetName);


  const masterData = masterSheet.getDataRange().getValues();
  const targetData = targetSheet.getDataRange().getValues();

  const masterHeader = masterData[0];
  const targetHeader = targetData[0];
  const masterRows = masterData.slice(1);
  const targetRows = targetData.slice(1);


  const masterUUIDIndex = 51; 
  
  const targetUUIDIndex = 0;  
  
  const existingUUIDs = new Set(targetRows.map(row => row[targetUUIDIndex]));

  const newRows = masterRows.filter(row => !existingUUIDs.has(row[masterUUIDIndex]));

  if (newRows.length === 0) {
    Logger.log(" No new rows to update. Target is already up to date.");
    return;
  }

  targetSheet.getRange(targetSheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);

  Logger.log(` ${newRows.length} new rows added to Target sheet.`);
}
