

// function replacePlaceholders() {
//   const docId = "1mxE-zlNgS3nDB2MqLqgB4bNawf2K1-y_"; 
//   const folderId = RELIEVING_LETTER_FOLDER_ID;

//   DriveApp.getFilesByName("")
  
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const sheet = ss.getSheetByName("Employee Data");
//   const data = sheet.getDataRange().getValues();
//   const headers = data[0];

//   const updateColIndex = 20; 
//   const sentColIndex = 21;  

//   for (let i = 1; i < data.length; i++) {
//     const row = data[i];

//     if (row[updateColIndex] === true) {
//       try {
//         const employeeName = row[2]; 
//         const empId = row[1];        
//         const currentAddress = row[18]; 
//         const resignationDate = row[12]; 
//         const lastWorkingDate = row[10]; 
//         const doj = row[9];          
//         const designation = row[6];  
//         const hrName = row[4];       

//         const formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM dd, yyyy");

//         const fileName = `${employeeName}_Relieving Letter`;
//         const copiedFile = DriveApp.getFileById(docId).makeCopy(fileName, DriveApp.getFolderById(folderId));
//         const copyId = copiedFile.getId();

//         const doc = DocumentApp.openById(copyId);
//         const body = doc.getBody();

//         // Replace placeholders
//         const replacements = {
//           "<<Date>>": formattedDate,
//           "<<Employee ID>>": empId || "",
//           "<<Employee Name>>": employeeName || "",
//           "<<Employee Address>>": currentAddress || "",
//           "<<Resignation date>>": formatDateSafe(resignationDate),
//           "<<Last Working Date>>": formatDateSafe(lastWorkingDate),
//           "<<Date of Joining>>": formatDateSafe(doj),
//           "<<Designation>>": designation || "",
//           "<<HR Name>>": hrName || ""
//         };

//         for (const [placeholder, value] of Object.entries(replacements)) {
//           if (!value) {
//             console.warn(`Warning: Missing value for placeholder ${placeholder} in row ${i + 1}`);
//           }
//           body.replaceText(placeholder, value);
//         }

//         doc.saveAndClose();

//         sheet.getRange(i + 1, sentColIndex + 1).setValue("SENT");

//       } catch (error) {
//         console.error(`Error on row ${i + 1}: ${error.message}`);
//       }
//     }
//   }

//   SpreadsheetApp.getUi().alert('Relieving letters generated and saved successfully.');
// }

// function formatDateSafe(dateValue) {
//   if (!(dateValue instanceof Date)) {
//     try {
//       dateValue = new Date(dateValue);
//       if (isNaN(dateValue)) return "";
//     } catch {
//       return "";
//     }
//   }
//   return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), "MMMM dd, yyyy");
// }
