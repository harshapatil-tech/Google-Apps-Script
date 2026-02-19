 //Code for the addReviewer and Remove Reviewer from Reviewer managment sheet
class ReviewerManager {
  constructor(sheetId, managementTabId, indexTabId, headerRow) {
    const spreadSheet = CentralLibrary.DataAndHeaders(sheetId);

    this.spreadSheet = spreadSheet;
    //console.log(this.spreadSheet);

    // Management Sheet (Add & Remove reviewers)
    const management = spreadSheet.getSheetById(managementTabId);
    //console.log(management);

    this.managementSheet = management.sheet;
    //console.log(this.managementSheet);

    const [headers, data] = management.getDataIndicesFromSheet(headerRow);
    //console.log(headers, data);

    this.reviewerHeaders = headers;
    //console.log(this.reviewerHeaders);

    this.reviewerData = data;
    //console.log(this.reviewerData);


    // Index Sheet
    const index = spreadSheet.getSheetById(indexTabId);
    //console.log(index);

    this.indexSheet = index.sheet;
    //console.log(this.indexSheet);

    const [indexHeaders, indexData] = index.getDataIndicesFromSheet(1);
    // console.log(indexHeaders, indexData);

    this.indexHeaders = indexHeaders;
    //console.log(this.indexHeaders);

    this.indexData = indexData;
    //console.log(this.indexData);
    this.status = "";
    //console.log("The status is:-", this.status);
    this.emailLinkArray = [];
    // console.log("The email link array :-", this.emailLinkArray);
    this.removedIndexes = [];
    //console.log("The Remove Index is:-", +this.removedIndexes);
  }

  //Reviwer is add when email id, department & Add? is fill in Reviewer Managment sheet
  reviewerAdd() {
    this.reviewerData.forEach((row, idx) => {
      if (row[this.reviewerHeaders["Add?"]] === true) {
        const emailId = row[this.reviewerHeaders["Email ID"]];
        console.log("The Email id is:-", emailId);
        const department = row[this.reviewerHeaders["Department"]];
        console.log("The Department is:-", department);
        const rowIndex = idx + 4;
        //console.log("The Rowindex is:-", rowIndex);

        if (!emailId || !department) return;

        const [sheetLink, reviewerStatus] = getReviewerStatus(emailId, department);
        // console.log("Returned values from getReviewerStatus:", sheetLink, reviewerStatus);

        if (sheetLink && reviewerStatus) {
          // Already exists but was reactivated
          this.status += reviewerStatus + "\n";
          this.emailLinkArray.push([emailId, department, sheetLink]);
        } else if (reviewerStatus && !sheetLink) {
          // Already active — not creating new sheet
          this.status += reviewerStatus + "\n";
        } else {
          // Brand new reviewer, create sheet
          const [reviewerSheetLink, reviewerSheetId, sheetID, reviewerName] = setReviewerSheet_CreateNewSheet(emailId, department);

          setValuesAtStart(reviewerSheetId, department, sheetID, emailId);
          //createBackendReviwerSheetByDepartment(reviewerSheetLink, department);
          createBackendReviewerSheet(reviewerSheetLink, department);

          this.emailLinkArray.push([emailId, department, reviewerSheetLink]);

          this.status += `A new sheet for reviewer ${reviewerName} has been created\n`; //reviwerName
        }

        applyCustomFormatting(this.managementSheet.getRange(rowIndex, 2, 1, 3)).clearContent();
      }
    });
    this.updateIndexSheet();
    SpreadsheetApp.getUi().alert(this.status);
  }

  //After Adding Reviewer Update data in index sheet
  updateIndexSheet() {
    let lastRow = this.indexSheet.getLastRow();
    //console.log("The lastRow is:-", lastRow);

    this.emailLinkArray.forEach(r => {
      const srNo = (lastRow === 2) ? 1 : this.indexSheet.getRange(lastRow, this.indexHeaders["#"] + 1).getValue() + 1;
      //console.log("SrNo is :-", srNo);

      const nextRow = lastRow + 1;
      //console.log("The nextRow is:-", nextRow);

      this.indexSheet.getRange(nextRow, this.indexHeaders["#"] + 1).setValue(srNo);
      this.indexSheet.getRange(nextRow, this.indexHeaders["QA Reviewer Email"] + 1).setValue(r[0]);
      this.indexSheet.getRange(nextRow, this.indexHeaders["Department"] + 1).setValue(r[1]);
      this.indexSheet.getRange(nextRow, this.indexHeaders["Sheet Link"] + 1).setValue(r[2]);
      lastRow++;
    });

  }

  //Reviewer is removed when email id, department & Remove? id fill in Reviwer Managment sheet
  reviewerRemove() {
    this.reviewerData.forEach((row, idx) => {
      if (row[this.reviewerHeaders["Remove?"]] === true) {
        const emailId = row[this.reviewerHeaders["Email ID"]];
        console.log("The email id is:-", emailId);
        const department = row[this.reviewerHeaders["Department"]];
        console.log("The department is:-", department);
        const rowIndex = idx + 10;
        //console.log("The Row index is:-", rowIndex);

        if (!emailId || !department) return;

        const [archivedFile, message] = archiveReviewer(emailId, department);
        // console.log("Returned value from archiveReviewer", archivedFile, message);

        this.status += message + "\n";
        console.log("The status is:-", this.status);

        applyCustomFormatting(this.managementSheet.getRange(rowIndex, 2, 1, 3)).clearContent();

        if (archivedFile !== undefined) {
          const foundIndex = this.indexData.findIndex(r =>
            r[this.indexHeaders["QA Reviewer Email"]] === emailId &&
            r[this.indexHeaders["Sheet Link"]] === archivedFile);
          //console.log("The found index is:-", foundIndex);

          if (foundIndex !== -1) {
            this.removedIndexes.push(foundIndex);
          }
        }
      }
    });

    this.removedIndexes.sort((a, b) => b - a).forEach(index => this.indexData.splice(index, 1));
    this.indexData = this.indexData.filter(row => row.some(cell => cell !== ''));

    for (let i = 0; i < this.indexData.length; i++) {
      this.indexData[i][this.indexHeaders["#"]] = i + 1;
    }
    this.clearAndWriteIndexSheet();
    SpreadsheetApp.getUi().alert(this.status);
  }

  //clear the old data of index sheet and update new
  clearAndWriteIndexSheet() {
    applyCustomFormatting(this.indexSheet.getRange(3, 1, this.indexSheet.getLastRow(), 4)).clearContent();
    if (this.indexData.length > 0) {
      this.indexSheet.getRange(3, 1, this.indexData.length, 4).setValues(this.indexData);
    }
  }
}


function addReviewer() {
  const add = new ReviewerManager(
    QA_HEAD_DASHBORD_ID,
    REVIEWER_MANAGEMENT_TAB_ID,
    REVIEWER_INDEX_TAB_ID,
    2
  );
  add.reviewerAdd();
}


function removeReviewer() {
  const remove = new ReviewerManager(
    QA_HEAD_DASHBORD_ID,
    REVIEWER_MANAGEMENT_TAB_ID,
    REVIEWER_INDEX_TAB_ID,
    8
  );
  remove.reviewerRemove();
}


//_______________________________________________________________________________________________

// function getReviewerStatus(emailId, department) {
//   const spreadsheet = SpreadsheetApp.openById(MASTER_DB_SPREADSHEET_ID);
//   const sheet = spreadsheet.getSheetById(REVIEWER_DB_TAB_ID);
//  // const sheet = spreadsheet.getSheetByName("Reviewer DB");
//   const dataRange = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
//   const header = dataRange[0], data = dataRange.slice(1);

//   const emailIdx = header.indexOf("Email ID");
//   const departmentIdx = header.indexOf("Department");
//   const reviewerIdx = header.indexOf("Reviewer Name")
//   const reviewerSheetLink = header.indexOf("Reviewer Sheet Link");
//   const activeStatusIdx = header.indexOf("Active?");
//   const removedDateIdx = header.indexOf("Removed Date");

//   const rowIndex = data.findIndex(r => r[emailIdx] === emailId && r[departmentIdx] === department);
//   const activeStatus = sheet.getRange(rowIndex + 2, activeStatusIdx + 1).getValue()

//   if (rowIndex !== -1) {
//     const sheetLink = sheet.getRange(rowIndex + 2, reviewerSheetLink + 1).getValue();
//     const reviewer = sheet.getRange(rowIndex + 2, reviewerIdx + 1).getValue();
//     if (sheetLink === '') {
//       return [undefined, undefined];
//     } else {
//       // Move the file from ex reviewer to the department folder
//       if (!activeStatus) {
//         const unarchiveSheet = SpreadsheetApp.openByUrl(sheetLink)
//         const sheetId = unarchiveSheet.getId();
//         DriveApp.getFileById(sheetId).addEditor(emailId);
//         const fileinFolder = DriveApp.getFileById(sheetId).getParents();
//         const folderId = fileinFolder.next().getId();
//         const parentFolder = DriveApp.getFolderById(folderId).getParents();
//         const parentFolderId = parentFolder.next().getId();
//         const folder = DriveApp.getFolderById(parentFolderId);
//         DriveApp.getFileById(sheetId).moveTo(folder);
//         sheet.getRange(rowIndex + 2, activeStatusIdx + 1).setValue(true);
//         return [sheetLink, `Reviewer ${reviewer} has been unarchieved`]
//       } else {
//         return [undefined, `Reviewer ${reviewer} already exist`]
//       }
//     }
//   }
//   else
//     return [undefined, undefined];
// }



// function setReviewerSheet_CreateNewSheet(emailId, department) {
//   const spreadsheet = SpreadsheetApp.openById("1u4i5rMCWM0mSgDWdf6EmGVi40kb9i2Qx51SE1cupIjU")
//   const sheet = spreadsheet.getSheetByName("Reviewer DB");
//   const dataRange = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
//   const header = dataRange[0], data = dataRange.slice(1);

//   const reviewerDBIndices = {
//     srNoIdx: header.indexOf("#"),
//     emailIdx: header.indexOf('Email ID'),
//     departmentIdx: header.indexOf('Department'),
//     reviewerNameIdx: header.indexOf('Reviewer Name'),
//     reviewerSheetLink: header.indexOf("Reviewer Sheet Link"),
//     addedDateIdx: header.indexOf('Added Date'),
//     activeIdx: header.indexOf('Active?'),
//     sheetIdx: header.indexOf('Sheet ID'),
//     uniqueIdIdx: header.indexOf('Unique ID')

//   }
//   console.log("index of :-", reviewerDBIndices);

//   const rowIndex = data.findIndex(r => r[reviewerDBIndices.emailIdx] === emailId &&
//     r[reviewerDBIndices.departmentIdx].trim().toLowerCase() === department.trim().toLowerCase());

//   const sortedSheetIDs = data.map(r => r[reviewerDBIndices.sheetIdx])
//     .filter(Boolean)
//     .sort((str1, str2) => {
//       const num1 = parseInt(str1.slice(1)); // Extract the numeric part from the string
//       const num2 = parseInt(str2.slice(1));
//       return num1 - num2; // Compare the numeric values for sorting
//     });

//   let nextSheetID;

//   if (sortedSheetIDs.length > 0) {
//     nextSheetID = parseInt(sortedSheetIDs[sortedSheetIDs.length - 1].slice(1)) + 1;
//     nextSheetID = 'R' + nextSheetID;
//   } else
//     nextSheetID = 'R1'


//   let sheetName;
//   let reviewerSheetId;
//   let reviewerSheetLink;
//   let reviewerName;
//   let uniqueId;

//   if (rowIndex !== -1) {   // If the reviwer is present but does not have a sheet;
//     reviewerName = data[rowIndex][reviewerDBIndices.reviewerNameIdx];
//     uniqueId = data[rowIndex][reviewerDBIndices.uniqueIdIdx];

//     sheet.getRange(rowIndex + 2, reviewerDBIndices.addedDateIdx + 1)
//       .setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yy'));
//     sheet.getRange(rowIndex + 2, reviewerDBIndices.activeIdx + 1).setValue(true);

//     sheetName = `QA_Reviewer_${reviewerName.split(' ').join("_")}`;
//     // reviewerSheetId = createCopyOfSpreadsheet("1WDmG0nebB6NgCLUOs41wwutbcbpB9tFlFF4f1t7Xld4", sheetName, department, "Reviewer Sheets");
//     reviewerSheetId = createCopyOfSpreadsheet("17t_fKuXAkXtoVnB4FmCRDIHyJzdiAzIYGbQR0eEdeQM", sheetName, department, "Reviewer Sheets");
//     reviewerSheetLink = `https://docs.google.com/spreadsheets/d/${reviewerSheetId}/edit`;

//     sheet.getRange(rowIndex + 2, reviewerDBIndices.reviewerSheetLink + 1).setValue(reviewerSheetLink);
//     sheet.getRange(rowIndex + 2, reviewerDBIndices.sheetIdx + 1).setValue(nextSheetID);

//   } else {  // If reviwer not present
//     const smeDBDatabase = spreadsheet.getSheetByName("SME DB");
//     const smeDataRange = smeDBDatabase.getRange(1, 1, smeDBDatabase.getLastRow(), smeDBDatabase.getLastColumn()).getValues();
//     const smeHeader = smeDataRange[0], smeData = smeDataRange.slice(1);
//     const smeDBIndices = {
//       emailIdx: smeHeader.indexOf("Email ID"),
//       nameIdx: smeHeader.indexOf("SME Name"),
//       uniqueIdIdx: smeHeader.indexOf("Unique ID")
//     }

//     const reviewerRow = smeData.find(r => r[smeDBIndices.emailIdx] === emailId);
//     if (!reviewerRow) throw new Error("Reviewer not found in SME DB");

//     reviewerName = reviewerRow[smeDBIndices.nameIdx];
//     uniqueId = reviewerRow[smeDBIndices.uniqueIdIdx];

//     const lastRow = sheet.getLastRow() + 1;
//     reviewerName = smeData.filter(r => r[smeDBIndices.emailIdx] === emailId).map(r => r[smeDBIndices.nameIdx])[0];
//     const lastSrNo = sheet.getRange(lastRow - 1, reviewerDBIndices.srNoIdx + 1).getValue();
//     sheet.getRange(lastRow, reviewerDBIndices.srNoIdx + 1).setValue(lastSrNo + 1);
//     sheet.getRange(lastRow, reviewerDBIndices.emailIdx + 1).setValue(emailId);
//     sheet.getRange(lastRow, reviewerDBIndices.reviewerNameIdx + 1).setValue(reviewerName);
//     sheet.getRange(lastRow, reviewerDBIndices.departmentIdx + 1).setValue(department);
//     sheet.getRange(lastRow, reviewerDBIndices.uniqueIdIdx + 1).setValue(uniqueId);
//     sheet.getRange(lastRow, reviewerDBIndices.addedDateIdx + 1)
//       .setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yy'));
//     sheet.getRange(lastRow, reviewerDBIndices.activeIdx + 1).insertCheckboxes().setValue(true);

//     sheetName = `QA_Reviewer_${reviewerName.split(' ').join("_")}`;
//     // reviewerSheetId = createCopyOfSpreadsheet("1WDmG0nebB6NgCLUOs41wwutbcbpB9tFlFF4f1t7Xld4", sheetName, department, "Reviewer Sheets");
//      reviewerSheetId = createCopyOfSpreadsheet("17t_fKuXAkXtoVnB4FmCRDIHyJzdiAzIYGbQR0eEdeQM", sheetName, department, "Reviewer Sheets");
//     reviewerSheetLink = `https://docs.google.com/spreadsheets/d/${reviewerSheetId}/edit`;

//     sheet.getRange(lastRow, reviewerDBIndices.reviewerSheetLink + 1).setValue(reviewerSheetLink);
//     sheet.getRange(lastRow, reviewerDBIndices.sheetIdx + 1).setValue(nextSheetID);
//     applyCustomFormatting(sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()))
//   }
//   return [reviewerSheetLink, reviewerSheetId, nextSheetID, reviewerName];
// }


// function setValuesAtStart(sheetIdx, department, uniqueID, emailId) {
//   const reviewerSheet = SpreadsheetApp.openById(sheetIdx);
//   const reviewerSheetUrl = reviewerSheet.getUrl();
//   const rubric = reviewerSheet.getSheetByName('Rubric')
//   const backendSheet = reviewerSheet.getSheetByName("Backend")
//   const masterDB = SpreadsheetApp.openById("1u4i5rMCWM0mSgDWdf6EmGVi40kb9i2Qx51SE1cupIjU")
//   // masterDB.addEditor(emailId);   // CHANGE IT 1eJ-7nKeICwYXEBh4oEzLwC6SpkYzCw5ntYxarbD2jxA

//   DriveApp.getFileById(sheetIdx).addEditor(emailId);  // CHANGE IT

//   defineProtectionFunctionalities(rubric);
//   defineProtectionFunctionalities(backendSheet, hide = true)
//   const departmentRow = backendSheet.getRange(1, 1, 1, backendSheet.getLastColumn()).getValues().flat();
//   const departmentIdx = departmentRow.indexOf('Department') + 1;
//   backendSheet.getRange(1, departmentIdx).offset(0, 1).setValue(department);

//   const sheetIdRow = backendSheet.getRange(2, 1, 1, backendSheet.getLastColumn()).getValues().flat();
//   const sheetIdIdx = sheetIdRow.indexOf('Sheet ID') + 1;
//   backendSheet.getRange(2, sheetIdIdx).offset(0, 1).setValue(uniqueID);  // According to the number of sheets in the folder
// }


// function defineProtectionFunctionalities(sheet, hide = false) {
//   let protection;
//   if (hide)
//     protection = sheet.hideSheet().protect();
//   else
//     protection = sheet.protect();
//   // Optionally, you can customize the protection settings
//   protection.setDescription('Protected Sheet : Only the owner can edit this sheet');
//   protection.setWarningOnly(false); // Show a warning when trying to edit
//   protection.removeEditors(protection.getEditors());
// }


// function harsha(){
//   createCopyOfSpreadsheet("17t_fKuXAkXtoVnB4FmCRDIHyJzdiAzIYGbQR0eEdeQM",
//     "fsdfsdfsd",
//     "Mathematics",
//     "Reviewer Sheets"
//   )
// }


// function createCopyOfSpreadsheet(sourceSpreadsheetId, newSpreadsheetName, department, innerFolderName) {
//   // Open the source spreadsheet
//   const sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
//   let folderId = createFolderIfNotExists("14oM-13qyzcSDEzFo1CdtXRVSIH61b3Yx", department);
//   folderId = createFolderIfNotExists(folderId, innerFolderName);
//   const fileId = createSheetInFolder(folderId, sourceSpreadsheet, fileName = newSpreadsheetName);
//   return fileId
// }


// function createSheetInFolder(folderId, sourceSpreadsheet, sheetName = "New Sheet") {
//   var folder = DriveApp.getFolderById(folderId);

//   // Check if a file with the same name already exists in the folder
//   var existingFiles = folder.getFilesByName(sheetName);
//   if (existingFiles.hasNext()) {
//     // If a file with the same name exists, return its URL without creating a new one
//     return existingFiles.next().getId();
//   }

//   // Create a new Google Sheets file
//   var newSpreadsheet = sourceSpreadsheet.copy(sheetName);

//   // Move the newly created Google Sheets file to the specified folder
//   DriveApp.getFileById(newSpreadsheet.getId()).moveTo(folder);

//   // Return the URL of the created Google Sheets file
//   return newSpreadsheet.getId();
// }



// function archiveReviewer(emailId, department) {
//   const spreadsheet = SpreadsheetApp.openById("1u4i5rMCWM0mSgDWdf6EmGVi40kb9i2Qx51SE1cupIjU"); //1eJ-7nKeICwYXEBh4oEzLwC6SpkYzCw5ntYxarbD2jxA
//   const sheet = spreadsheet.getSheetByName("Reviewer DB");
//   const dataRange = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
//   const header = dataRange[0], data = dataRange.slice(1);

//   const emailIdx = header.indexOf("Email ID");
//   const reviewerIdx = header.indexOf("Reviewer Name")
//   const departmentIdx = header.indexOf("Department");
//   const reviewerSheetLink = header.indexOf("Reviewer Sheet Link");
//   const activeStatus = header.indexOf("Active?");
//   const removedDateIdx = header.indexOf("Removed Date");

//   const rowIndex = data.findIndex(r => r[emailIdx] === emailId && r[departmentIdx] === department);

//   if (rowIndex === -1) { // If there is no reviewer no question of deleting
//     return [undefined, "Reviewer does not exist!"];
//   } else {
//     // Check if the sheet is already archieved
//     if (sheet.getRange(rowIndex + 2, activeStatus + 1).getValue() === false) {
//       const reviewerName = sheet.getRange(rowIndex + 2, reviewerIdx + 1).getValue();
//       const sheetLink = sheet.getRange(rowIndex + 2, reviewerSheetLink + 1).getValue();
//       return [undefined, `Reviewer ${reviewerName} is already archieved!`];
//     } else {
//       const reviewerName = sheet.getRange(rowIndex + 2, reviewerIdx + 1).getValue()
//       const sheetLink = sheet.getRange(rowIndex + 2, reviewerSheetLink + 1).getValue();
//       sheet.getRange(rowIndex + 2, activeStatus + 1).setValue(false);
//       const archiveSheet = SpreadsheetApp.openByUrl(sheetLink);
//       const editors = archiveSheet.getEditors();
//       editors.forEach(editor => {
//         const email = editor.getEmail();
//         archiveSheet.removeEditor(email);
//       })
//       const sheetId = archiveSheet.getId();
//       today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yy')
//       sheet.getRange(rowIndex + 2, removedDateIdx + 1).setValue(today);
//       // Move the file to ex reviewer
//       const folders = DriveApp.getFileById(sheetId).getParents();
//       const parentFolderId = folders.next().getId();
//       const archieveFolderId = createFolderIfNotExists(parentFolderId, "Ex Reviewer Sheets");
//       const folder = DriveApp.getFolderById(archieveFolderId);
//       DriveApp.getFileById(sheetId).moveTo(folder);

//       return [sheetLink, `Reviewer ${reviewerName} has been archieved!`];
//     }
//   }
// }



// /**
//  * This function creates a folder inside the parent folder given the id of the parent folder and name of the folder
//  * that is to be created. It returns the id the new folder created.
//  * If folder with the same name already exists inside the parent folder then it just returns the id of the folder
//  * whose name was passed.
//  * @params : parentFolderId, newFolderName
//  */

// function createFolderIfNotExists(parentFolderId, newFolderName) {
//   var parentFolder = getFolderIfExists(parentFolderId);

//   if (parentFolder) {
//     var folders = parentFolder.getFoldersByName(newFolderName);

//     if (!folders.hasNext()) {
//       var newFolder = parentFolder.createFolder(newFolderName);
//       Logger.log("Folder created: " + newFolderName);
//       return newFolder.getId(); // Return the ID of the newly created folder
//     } else {
//       var existingFolder = folders.next();
//       Logger.log("Folder already exists: " + newFolderName);
//       return existingFolder.getId(); // Return the ID of the existing folder
//     }
//   } else {
//     Logger.log("Parent folder not found: " + parentFolderId);
//     return null; // Indicate failure or handle the error accordingly
//   }
// }


// /**
//  * Creates a google sheet inside the folder whose id is passed as a parameter with name passed as sheetName parameter.
//  * If the sheetName is not passed "New Sheet" will be the name of the new sheet created.
//  * It returns the id of the newly created sheet.
//  * If a sheet with the same name is already present in the folder then it returns the id of the sheet
//  * @params : folderId, sheetName
//  */
// // function createSheetInFolder(folderId, sheetName = "New Sheet") {
// //   var folder = DriveApp.getFolderById(folderId);

// //   // Check if a file with the same name already exists in the folder
// //   var existingFiles = folder.getFilesByName(sheetName);
// //   if (existingFiles.hasNext()) {
// //     // If a file with the same name exists, return its URL without creating a new one
// //     return existingFiles.next().getId();
// //   }

// //   // Create a new Google Sheets file
// //   var newSpreadsheet = SpreadsheetApp.create(sheetName);

// //   // Move the newly created Google Sheets file to the specified folder
// //   DriveApp.getFileById(newSpreadsheet.getId()).moveTo(folder);

// //   // Return the URL of the created Google Sheets file
// //   return newSpreadsheet.getId();
// // }



// /**
//  * Gets the folder if the folder id is given.
//  */
// function getFolderIfExists(folderId) {
//   try {
//     return DriveApp.getFolderById(folderId);
//   } catch (e) {
//     return null;
//   }
// }














//-----------------------------------
//OLD addreviwer Code
// function addReviewer() {
//   const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
//   const sheet = spreadsheet.getActiveSheet();
//   const indexSheet = spreadsheet.getSheetByName("Index");
//   const dataRange = sheet.getRange(3, 1, 4, 4).getValues();
//   const header = dataRange[0];
//   let data = dataRange.slice(1);

//   const indexDataRange = indexSheet.getRange(2, 1, indexSheet.getLastRow(), 4).getValues();
//   const indexHeader = indexDataRange[0];
//   let indexData = indexDataRange.slice(1);

//   const indexSrNo = indexHeader.indexOf("#");
//   const indexReviewerEmail = indexHeader.indexOf("QA Reviewer Email");
//   const indexDepartmentIdx = indexHeader.indexOf("Department");
//   const indexSheetLink = indexHeader.indexOf("Sheet Link");

//   const emailIdx = header.indexOf("Email ID");
//   const departmentIdx = header.indexOf("Department");
//   const addIdx = header.indexOf("Add?");

//   let status = ""
//   const emailLinkArray = []
//   data.forEach((r, index) => {
//     if((r[addIdx] === true) && (r[emailIdx] != '' && r[departmentIdx] != '')){
//       const emailId = r[emailIdx];
//       const department = r[departmentIdx];
//       const rowIndex = index + 4;
//       // If reviwer exists then show they exists if not then go on to create a sheet.
//       const [sheetLink, reviewerStatus] = getReviewerStatus(emailId, department);

//       if (reviewerStatus !== undefined && sheetLink !== undefined){
//         status += reviewerStatus + "\n";
//         emailLinkArray.push([emailId, department, sheetLink]);
//       }else if(reviewerStatus !== undefined && sheetLink === undefined) {
//         status += reviewerStatus + "\n"
//       }else {
//         const [reviewerSheetLink, reviewerSheetId, sheetID, reviewerName] = setReviewerSheet_CreateNewSheet(emailId, department);
//         setValuesAtStart(reviewerSheetId, department, sheetID, emailId);
//         createBackendReviwerSheetByDepartment(reviewerSheetLink, department);
//         emailLinkArray.push([emailId, department, reviewerSheetLink])
//         status += `A new sheet for reviewer ${reviewerName} has been created` + "\n"
//       }

//       applyCustomFormatting(sheet.getRange(rowIndex, 2, 1, 3)).clearContent();
//     }
//   });

//   let lastRow = indexSheet.getLastRow();
//   emailLinkArray.forEach(r => {
//     if(lastRow === 2){
//       indexSheet.getRange(lastRow + 1, indexSrNo+1).setValue(1);
//       indexSheet.getRange(lastRow + 1, indexReviewerEmail+1).setValue(r[0]);
//       indexSheet.getRange(lastRow + 1, indexDepartmentIdx+1).setValue(r[1]);
//       indexSheet.getRange(lastRow + 1, indexSheetLink+1).setValue(r[2]);
//       lastRow++;
//     }else{
//       // Get the last row Index
//       const lastSrNo = indexSheet.getRange(lastRow, indexSrNo+1).getValue();
//       indexSheet.getRange(lastRow + 1, indexSrNo+1).setValue(lastSrNo + 1);
//       indexSheet.getRange(lastRow + 1, indexReviewerEmail+1).setValue(r[0]);
//       indexSheet.getRange(lastRow + 1, indexDepartmentIdx+1).setValue(r[1]);
//       indexSheet.getRange(lastRow + 1, indexSheetLink+1).setValue(r[2]);
//       lastRow++;
//     }
//   })
//   SpreadsheetApp.getUi().alert(status);
// }






//_________________________________
// old removeReviwer code
// function removeReviewer(){
//   const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
//   const sheet = spreadsheet.getActiveSheet();
//   const indexSheet = spreadsheet.getSheetByName("Index");
//   const dataRange = sheet.getRange(9, 1, 4, 4).getValues();
//   const header = dataRange[0];
//   let data = dataRange.slice(1);

//   const indexDataRange = indexSheet.getRange(2, 1, indexSheet.getLastRow(), 4).getValues();
//   const indexHeader = indexDataRange[0];
//   let indexData = indexDataRange.slice(1);

//   const emailIdx = header.indexOf("Email ID");
//   const departmentIdx = header.indexOf("Department");
//   const removeIdx = header.indexOf("Remove?");

//   const indexSrNo = indexHeader.indexOf("#");
//   const indexReviewerEmail = indexHeader.indexOf("QA Reviewer Email");
//   const indexDepartmentIdx = indexHeader.indexOf("Department");
//   const indexSheetLink = indexHeader.indexOf("Sheet Link");

//   const removedIndexes = [];
//   let status = "";

//   data.forEach((r, index) => {
//     if(r[removeIdx] === true && r[emailIdx] !== '' && r[departmentIdx] !== ''){
//       const emailId = r[emailIdx];
//       const department = r[departmentIdx];
//       // const reviewerStatus = getReviewerStatus(emailId, department, archieveReviewer=true);
//       const [archivedFile, message] = archiveReviewer(emailId, department);
//       status += message + "\n";
//       // Clear the corresponding row
//       applyCustomFormatting(sheet.getRange(index + 10, 2, 1, 3)).clearContent();
//       if (archivedFile !== undefined) {
//         const foundIndex = indexData.findIndex(r => r[indexReviewerEmail] === emailId && r[indexSheetLink] === archivedFile);
//         if (foundIndex !== -1) {
//           removedIndexes.push(foundIndex);
//         }
//       }
//     }
//   });

//   // Delete rows in reverse order to avoid shifting issues
//   removedIndexes.sort((a, b) => b - a).forEach(index => {
//     indexData.splice(index, 1);
//   });

//   // Remove empty rows from indexData
//   indexData = indexData.filter(row => row.some(cell => cell !== ''));
//   // Update the serial numbers in index sheet
//   for (let i = 0; i < indexData.length; i++) {
//     indexData[i][indexSrNo] = i + 1;
//   }

//   // Clear the existing data in the index sheet, starting from row 3
//   applyCustomFormatting(indexSheet.getRange(3, 1, indexSheet.getLastRow(), 4)).clearContent();

//   // Write the updated data back to the index sheet
//   if (indexData.length > 0) {
//     indexSheet.getRange(3, 1, indexData.length, 4).setValues(indexData);
//   }

//   SpreadsheetApp.getUi().alert(status);
// }

