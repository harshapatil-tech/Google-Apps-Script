const RELIEVING_LETTER_FOLDER_ID = "15ovsrPMjx7o_nkUh0I6S4r5FSHqngNWK";


function createRelievingLetter() {

  const year = new Date().getFullYear();
  console.log(year)

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hrInputSheet = ss.getSheetByName("Input Sheet");
  //const hrInputSheet = ss.getSheetByName("Copy of Input Sheet");
  
  let successCount = 0;   
  let errorCount = 0; 

  const [hrInputHeaders, hrInputData] = CentralLibrary.get_Data_Indices_From_Sheet(hrInputSheet, 2);
  
  const hrNameDesignationMap = getHrNameDesignationMap(ss);

  const rootFolder = DriveApp.getFolderById(RELIEVING_LETTER_FOLDER_ID);

  const folderOpt = rootFolder.getFoldersByName(`${year}`);
  if ( folderOpt.hasNext() === false ) {
    SpreadsheetApp.getUi().alert("Please a create a folder for current year / The current year folder could not be found");
    return;
  }   
  const folder = folderOpt.next();

  //const templateFile = DriveApp.getFilesByName("Relieving Letter.docx").next();
  const templateFile = DriveApp.getFilesByName("Relieving Letter_Template").next();

  
  const requiredFields = [
    "New Emp Id", "Employee Name", "Current Address", "Date of Resignation",
    "Date of Leaving", "DOJ", "Designation", "HR Name"
  ];

  //let errorCount = 0;

  hrInputData.forEach((row, i) => {

    const rowIndex = i + 4;
    const createLetter = row[hrInputHeaders["Create Relieving Letter?"]];
    const letterCreated = row[hrInputHeaders["Letter Created"]];
    
    if (createLetter === true && !letterCreated) {

      Logger.log(`Processing row ${rowIndex} for Employee ID: ${row[hrInputHeaders["New Emp Id"]]}`);
      const missingFields = requiredFields.filter(field => !row[hrInputHeaders[field]] || row[hrInputHeaders[field]] === "");

      if (missingFields.length > 0) {
        Logger.log(`Missing fields for row ${rowIndex}: ${missingFields.join(", ")}`);
        errorCount++;
        return;
      }
      try {

        // const tempDocFile = templateFile.makeCopy(`Relieving_Letter_${row[hrInputHeaders["New Emp Id"]]}`, folder);

        // Logger.log(`Temporary Google Doc created: ${tempDocFile.getName()}`);
        // console.log(tempDocFile.getId())
        // Utilities.sleep(10000); // pause 2 seconds

        const tempDocFile = createDocFile(templateFile, row[hrInputHeaders["New Emp Id"]], folder)

        console.log("Temporary file", tempDocFile.getName());

        //const doc = DocumentApp.openById(tempDocFile.id);
        const doc = DocumentApp.openById(tempDocFile.getId());

        const body = doc.getBody();
        console.log(doc.getName())

        const format = d => {
          if (!d) return "";
          const date = new Date(d);
          date.setHours(0, 0, 0, 0);
          return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd-MM-yyyy");
        };

        const hrName = row[hrInputHeaders["HR Name"]] || "";
        const hrDesignation = hrNameDesignationMap[hrName] || "";
        console.log(hrDesignation);

        body.replaceText("<<EmpCode>>", row[hrInputHeaders["New Emp Id"]]);
        body.replaceText("<<FinancialYear>>", getFinancialYear())
        body.replaceText("<<Date>>", format(new Date()));
        body.replaceText("<<Employee ID>>", row[hrInputHeaders["New Emp Id"]]);
        body.replaceText("<<Employee Name>>", row[hrInputHeaders["Employee Name"]]);
        body.replaceText("<<Employee Address>>", row[hrInputHeaders["Current Address"]]);
        body.replaceText("<<Resignation date>>", format(row[hrInputHeaders["Date of Resignation"]]));
        body.replaceText("<<Last Working Date>>", format(row[hrInputHeaders["Date of Leaving"]]));
        body.replaceText("<<Date of Joining>>", format(row[hrInputHeaders["DOJ"]]));
        body.replaceText("<<HR Name>>", hrName);
        body.replaceText("<<Designation>>", row[hrInputHeaders["Designation"]]);
        body.replaceText("<<HR Designation>>", hrDesignation);
        doc.saveAndClose();
        Logger.log(`Placeholders replaced and document saved.`);

       // exportDocToDocxAndDelete(doc.getId(), folder);

        hrInputSheet.getRange(rowIndex, hrInputHeaders["Letter Created"] + 1).setValue("Y");
        hrInputSheet.getRange(rowIndex, hrInputHeaders["Create Relieving Letter?"] + 1).setValue(false);

        Logger.log(`Sheet updated for row ${rowIndex}: Letter Created = Y, Checkbox = FALSE`);

        successCount++;


      }
      catch (err) {
        Logger.log(` Error creating letter for row ${rowIndex}: ${err}`);
        errorCount++;
      }
    } 
    // else {
    //   Logger.log(`Skipping row ${rowIndex} - Checkbox is unchecked or letter already created.`);
    // }

  });
  Logger.log(`Summary: ${successCount} success(es), ${errorCount} error(s).`);
  SpreadsheetApp.getUi().alert(`${successCount} letter(s) successfully created & ${errorCount} had error(s).`);
}


function getFinancialYear() {
    const today = new Date();
    const financialYearStr = ""
    const year = today.getFullYear();
    const month = today.getMonth() +  1;
    if (month <= 3) {
        financialYear = `${year-1}-${year.toString().slice(-2)}`
    } else {
        financialYear = `${year}-${(year+1).toString().slice(-2)}`
    }
    return financialYear;
}



/**
 * Exports a native Google Doc (by ID) into a .docx in the given folder, 
 * then trashes (deletes) the original Google Doc.
 *
 * This version explicitly passes { alt: "media" } to Drive.Files.export
 * so that the API returns the binary content of the .docx.
 *
 * @param {string}                 googleDocId       The ID of the Google Document to export.
 * @param {GoogleAppsScript.Drive.Folder} destinationFolder  A DriveApp Folder object where the .docx will be created.
 * @return {GoogleAppsScript.Drive.File}      The newly created `.docx` File object.
 */
// function exportDocToDocxAndDelete(googleDocId, destinationFolder) {
//   // Validate that a file ID is provided
//   if (googleDocId == null) throw new Error("No file ID."); // Throw error if null

//   var file = DriveApp.getFileById(googleDocId); // Get the file from Google Drive by ID
//   var mime = file.getMimeType(); // Get the MIME type of the file
//   var format = ""; // Initialize variable for format type
//   var ext = ""; // Initialize variable for file extension

//   // Determine the format and extension based on the MIME type
//   switch (mime) {
//     case "application/vnd.google-apps.document":
//       format = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"; // .docx format
//       ext = ".docx"; // Extension for Word documents
//       break;
//     case "application/vnd.google-apps.spreadsheet":
//       format = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"; // .xlsx format
//       ext = ".xlsx"; // Extension for Excel spreadsheets
//       break;
//     case "application/vnd.google-apps.presentation":
//       format = "application/vnd.openxmlformats-officedocument.presentationml.presentation"; // .pptx format
//       ext = ".pptx"; // Extension for PowerPoint presentations
//       break;
//     default:
//       return null; // Return null if unsupported MIME type
//   }

//   // Construct the URL for exporting the file in the specified format
//   var url = "https://www.googleapis.com/drive/v3/files/" + googleDocId + "/export?mimeType=" + format;
//   const baseName = file.getName();
//   const finalName = baseName.toLowerCase().endsWith('.docx')
//     ? baseName
//     : baseName + '.docx';

//   // Fetch the file as a Blob using the constructed URL and OAuth token for authorization
//   var blob = UrlFetchApp.fetch(url, {
//     method: "get", // HTTP method for fetching
//     headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() }, // Include authorization token
//     muteHttpExceptions: true // Prevent HTTP exceptions from being thrown
//   }).getBlob().setName(finalName); // Get the file as a Blob

//   // return blob; // Return the Blob containing the converted file

//   // 3) Create the .docx file in the destination folder
//   const finalFile = destinationFolder.createFile(blob);
//   Logger.log(`Exported .docx created: ${finalFile.getName()} (ID=${finalFile.getId()})`);

//   // 4) Trash the original Google Doc so only the .docx remains
//   DriveApp.getFileById(googleDocId).setTrashed(true);
//   Logger.log(`Trashed original Google Doc (ID=${googleDocId}).`);

//   return finalFile;
// }

//Direct copy the Template file no google docs conversion
function createDocFile(templateFile, empName, folder) {
  // Assume templateFile is still your .docx in DriveApp
  //const originalId = templateFile.getId();
  const newName = `Relieving_Letter_${empName}`;

  // Build the resource for the converted copy:
  // const resource = {
  //   name: newName,                           // new file’s name
  //   mimeType: MimeType.GOOGLE_DOCS,           // convert into a native Google Doc
  //   parents: [{ id: folder.getId() }]         // place it in your “2025” folder
  // };

  // Drive.Files.copy will convert the .docx → Google Doc
  // const copiedFile = Drive.Files.copy(resource, originalId);
  // Logger.log(`Converted copy (Google Doc) created: ${newName} (ID=${copiedFile.id})`);

  // DriveApp.getFileById(copiedFile.id).moveTo(folder);
  // Logger.log(`Moved converted Google Doc (ID=${copiedFile.name}) into folder "${folder.getName()}".`);

  const copiedFile = templateFile.makeCopy(newName, folder);

  Logger.log(`Google Doc copy created: ${newName} (ID=${copiedFile.getId()})`);

  return copiedFile;

}


function getHrNameDesignationMap(ss) {
  const dropdownSheet = ss.getSheetByName("Drop Downs");
  const [headers, data] = CentralLibrary.get_Data_Indices_From_Sheet(dropdownSheet, 1);

  const map = {};
  data.forEach(row => {
    const name = row[headers["HR Name"]];
    const designation = row[headers["HR Designation"]];
    if (name && designation) {
      map[name] = designation;
    }
  });
  console.log(map);
  return map;
}


//my tempory folder for testing
////const RELIEVING_LETTER_FOLDER_ID = "1bzluq-t6W_EV2hkZkmTag0IVsSGKeFcF";

// function checkTemplateType() {
//   const file = DriveApp.getFilesByName("Relieving Letter_Template").next();
//   Logger.log(file.getMimeType());
// }


