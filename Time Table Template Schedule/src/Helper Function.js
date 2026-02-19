// class Helper {


//   /**
//  * Convert a 1-based column index into its A1 letter (e.g. 1→"A", 3→"C", 27→"AA", 33→"AG").
//  */
//   static columnToLetter(colIndex) {
//     let letter = "";
//     while (colIndex > 0) {
//       const rem = (colIndex - 1) % 26;
//       letter = String.fromCharCode(65 + rem) + letter;  // 65 = "A"
//       colIndex = Math.floor((colIndex - 1) / 26);
//     }
//     return letter;
//   }





//   static getEmployeesByDept(obj, dept) {
//     return (obj[dept] || []).map(n => n.trim());
//   }


//   /**
//    * Capitalizes the first letter of a string.
//    *
//    * @param {string} str  The input string.
//    * @return {string}     The string with its first character upper-cased.
//    */
//   static capitalizeFirstLetter(str) {
//     if (typeof str !== 'string' || str.length === 0) return str;
//     return str.charAt(0).toUpperCase() + str.slice(1);
//   }

//   static mapSubjectToEmployees() {
//     const ss = SpreadsheetApp.openById("11b5A_ZDi9DskTiKJ0j4fBDkRbY-0I2h7rVRxr6KVK1o");
//     const employeeInfoSheet = ss.getSheetByName("Employee Info");
//     const [headers, data] = CentralLibrary.get_Data_Indices_From_Sheet(employeeInfoSheet);
//     const hashMap = {};
//     for (const row of data) {
//       let dept = row[headers["Department"]].trim();
//       if (dept === "Economics" || dept === "Finance" || dept === "Accounts")
//         dept = "Business";
//       const status = row[headers["Status"]].trim();
//       if (status == "Active") {
//         if (SUBJECTS.includes(dept)) {
//           dept = dept.toLowerCase();
//           if (dept in hashMap) {
//             hashMap[dept].push(row[headers["Employee Identifier"]])
//           } else {
//             hashMap[dept] = [];
//           }
//         }
//       }
//     }
//     return hashMap;
//   }

//   static logXlsmCommentsAndPopulateColumns(rootFolder, xlsmFileId, fileName) {
//     // const myDriveRoot = DriveApp.getRootFolder().getId();
//     const token       = ScriptApp.getOAuthToken();
//     // 1) Copy & convert via Drive API v3
//     const url = `https://www.googleapis.com/drive/v3/files/${xlsmFileId}/copy?supportsAllDrives=true`;
//     const requestBody = {
//       name:     fileName,
//       mimeType: 'application/vnd.google-apps.spreadsheet',
//       parents:  [ rootFolder ]
//     };
//     const response = UrlFetchApp.fetch(url, {
//       method:             'post',
//       contentType:        'application/json',
//       headers:            { Authorization: 'Bearer ' + token },
//       payload:            JSON.stringify(requestBody),
//       muteHttpExceptions: true
//     });
//     const copy = JSON.parse(response.getContentText());
//     if (!copy.id) {
//       throw new Error('Copy failed: ' + response.getContentText());
//     }
//     return copy.id;
//   }

//   static modifySS(id, department) {
//     const hashMap     = Helper.mapSubjectToEmployees();
//     const employees = hashMap[department]
//     const instantEmps = hashMap[department].map(emp => emp+" (I)");
//     const empArray = [...employees, "X", ...instantEmps, "Y"];
//     const empData = empArray.map(emp=> [emp]);
//     console.log(empData)

//     const ss = SpreadsheetApp.openById(id);
//     const sheets = ss.getSheets();

//     const tutorNameSheet = ss.getSheetByName("Tutor_Names");
//     tutorNameSheet.getRange(2, 2, tutorNameSheet.getLastRow(), 1).clear();
//     tutorNameSheet.getRange(2, 2, empData.length, 1).setValues(empData);

//     sheets.forEach(sheet => {
//       if (sheet.getName().includes("Online_Wk_")){
//         if (!sheet) {
//           Logger.log(`Sheet "${name}" not found—skipping.`);
//           return;
//         }
//         // 4) Insert a column after col 3
//         // sheet.deleteColumn(5);
//         const targetCol = 4;
//         // 5) Read all notes
//         const notes = sheet.getDataRange().getNotes();
//         // 6) Pull first note-per-row
//         // const out = notes.map(row =>
//         //   [ row.find(cellNote => cellNote)?.split(':\n')[1] || '' ]
//         // );
//         // 6) Pull only the “CtutorXXX” line from each row’s notes
//         const out = notes.map(row => {
//           // pick the first non-empty note in the row
//           const cellNote = row.find(n => n);
//           if (!cellNote) return [''];

//           // look for a line that starts with “Ctutor” (case-insensitive)
//           const match = cellNote.match(/^CwTutor\d+/mi);
//           return [ match ? match[0] : '' ];
//         });

//         // 7) Write back
//         sheet.getRange(1, targetCol, out.length, 1).setValues(out);
//         Logger.log(`Done – notes for "${sheet.getName()}" copied into column ${targetCol}`);
//       }
//     });
//   }

// }

