//ADD and REMOVE sme from SME Managment sheet
class SMESheetManager {
  constructor(localSheetId, smeTabId, indexTabId, backendSheetId, backendTabId,smeHeaderRow) {
    
    this.today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM-yy");
    //console.log(this.today);

    const localWrapper = CentralLibrary.DataAndHeaders(localSheetId);

    //SME Managment sheet
    const smeWrapper = localWrapper.getSheetById(smeTabId);
    //console.log(smeWrapper);
    this.smeSheet = smeWrapper.sheet;
    //console.log(this.smeSheet);
    const [smeHeaders, smeData] = smeWrapper.getDataIndicesFromSheet(smeHeaderRow);
    //console.log(smeHeaders, smeData);
    this.smeHeaders = smeHeaders;
    //console.log(this.smeHeaders);
    this.smeData = smeData;
    //console.log(this.smeData);

    //sme DB Sheet
    const backendWrapper = CentralLibrary.DataAndHeaders(backendSheetId).getSheetById(backendTabId);
    //console.log(backendWrapper);
    this.backendSheet = backendWrapper.sheet;
    //console.log(this.backendSheet);
    const [backendHeaders, backendData] = backendWrapper.getDataIndicesFromSheet(0);
    //console.log(backendHeaders,backendData);
    this.backendHeaders = backendHeaders;
    //console.log(this.backendHeaders);
    this.backendData = backendData;
    //console.log(this.backendData);

    //Index-sme sheet
    const indexWrapper = localWrapper.getSheetById(indexTabId);
    //console.log(indexWrapper);
    this.indexSheet = indexWrapper.sheet;
    //console.log(this.indexSheet);
    const [indexHeaders, indexData] = indexWrapper.getDataIndicesFromSheet(1);
    //console.log(indexHeaders, indexData);
    this.indexHeaders = indexHeaders;
    //console.log(this.indexHeaders);
    this.indexData = indexData;
    //console.log(this.indexData);
  }


//Fill email id, department & Add? in SME Managment sheet for adding sme
  smeAdd() {
    let status = "";
    //console.log("The status is:-",status);
    const emailLinkArray = [];
    //console.log("emailLink array is:-",emailLinkArray);

    this.smeData.forEach((row, inputSheetIndex) => {
      if (row[this.smeHeaders["Add?"]] === true) {
        const emailId = row[this.smeHeaders["Email ID"]];
        //console.log("The email id is:-",emailId);//Logger.log("Email id" + emailId); 
        const department = row[this.smeHeaders["Department"]];
        //console.log("Department is:-",department);


        this.backendData.forEach((r, index) => {
          const rowIndex = index + 2;
          //console.log("Row index is:-",rowIndex);
          if (
            r[this.backendHeaders["Email ID"]] === emailId &&
            (r[this.backendHeaders["Department"]] === department || r[this.backendHeaders["Department"]] === "Others")
          ) {
            const isActive = r[this.backendHeaders["Active?"]];
            //console.log("is Active is:-",isActive);
            const hasSheet = r[this.backendHeaders["SME Sheet Link"]];
            //console.log("The sheet link is",hasSheet);

            if (isActive === true) {
              status += 'User already Active\n';
            }
            else {
              this.backendSheet.getRange(rowIndex, this.backendHeaders["Active?"] + 1).setValue(true);
              this.backendSheet.getRange(rowIndex, this.backendHeaders["Added Date"] + 1).setValue(this.today);

              if (!hasSheet) {
                const name = r[this.backendHeaders["SME Name"]].toString().split(" ");
                //console.log("Sme Name is:-",name);
                const sheetName = name.length > 1 ? `QA_SME_${name[0]}_${name[name.length - 2]}_${name[name.length - 1]}` : `QA_SME_${name[0]}`;
                //console.log("QA sme sheet",sheetName);
                const newSheetId = createCopyOfSpreadsheet(TEMPLATE_SHEET_ID, sheetName, department, "SME Sheets");
                //console.log("The sheet id is:-",newSheetId);
                const sheetLink = `https://docs.google.com/spreadsheets/d/${newSheetId}/edit`;
                //console.log("Sheet link:-",sheetLink);
                this.backendSheet.getRange(rowIndex, this.backendHeaders["Sheet creation date"] + 1).setValue(this.today);
                this.backendSheet.getRange(rowIndex, this.backendHeaders["SME Sheet Link"] + 1).setValue(sheetLink);
                smeBackend(sheetLink, emailId);
                //console.log("Sheetlink and emailid",sheetLink, emailId);
                DriveApp.getFileById(newSheetId).addEditor(emailId);
                emailLinkArray.push([emailId, department, sheetLink]);
                status += `A new sheet for SME ${name.join(" ")} has been created\n`; 
               
              } else {
                const sheetLink = r[this.backendHeaders["SME Sheet Link"]];
                //console.log(sheetLink);
                const sheetId = SpreadsheetApp.openByUrl(sheetLink).getId();
                //console.log(sheetId);
                const file = DriveApp.getFileById(sheetId);
                //console.log("The file is:-",file);
                const folder = file.getParents().next().getParents().next();
                //console.log("The folder is:-",folder);
                file.moveTo(folder);
                file.addEditor(emailId);
                emailLinkArray.push([emailId, department, sheetLink]);
                status += `SME moved from archived folder to the main folder\n`;
              }
            }
          }

        });
        applyCustomFormatting(this.smeSheet.getRange(inputSheetIndex + 5, 2, 1, 3)).clearContent();
      }
    });

    let lastRow = this.indexSheet.getLastRow();
    //console.log("Last row is:-",lastRow);

    emailLinkArray.forEach(r => {
      const nextRow = lastRow + 1;
      //console.log("next row is:-",nextRow);
      let srNo = 1;
      //console.log("sr no is:-",srNo);

      if (lastRow > 2) {
        srNo = this.indexSheet.getRange(lastRow, this.indexHeaders["#"] + 1).getValue() + 1;
        //console.log(srNo);
      }

      this.indexSheet.getRange(nextRow, this.indexHeaders["#"] + 1).setValue(srNo);
      this.indexSheet.getRange(nextRow, this.indexHeaders["SME Email"] + 1).setValue(r[0]);
      this.indexSheet.getRange(nextRow, this.indexHeaders["Department"] + 1).setValue(r[1]);
      this.indexSheet.getRange(nextRow, this.indexHeaders["Sheet Link"] + 1).setValue(r[2]);
      lastRow++;
    });

   SpreadsheetApp.getUi().alert(status);
    //     try {
    //     SpreadsheetApp.getUi().alert(status);
    //   } catch (e) {
    //     Logger.log(status);
    //  }
  }


  //Fill email id, department & Remove? in SME Managment sheet for remove sme
  smeRemove() {
    let status = "";
    console.log("Status",status);

    const removedIndexes = [];
    console.log("Remove Indexes",removedIndexes);

    //const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM-yy");
    const today = this.today;
    console.log("The Date is :-" +today);

    this.smeData.forEach((row, inputSheetIndex) => {
    
      if (row[this.smeHeaders["Remove?"]] === true) {
        const emailId = row[this.smeHeaders["Email ID"]];
        console.log("The email id is :-" +emailId);
        const department = row[this.smeHeaders["Department"]];
        console.log("Department is:-" +department);

        this.backendData.forEach((r, index) => {
          if (emailId === r[this.backendHeaders["Email ID"]] &&
            department === r[this.backendHeaders["Department"]]) {
            if (r[this.backendHeaders["Active?"]] === true) {
              const rowIndex = index + 2;
              //console.log("The row index is:-" +rowIndex);

              //Deactivate SME 
              this.backendSheet.getRange(rowIndex, this.backendHeaders["Active?"] + 1).setValue(false);
              this.backendSheet.getRange(rowIndex, this.backendHeaders["Removed Date"] + 1).setValue(today);

              const sheetLink = r[this.backendHeaders["SME Sheet Link"]];
              console.log("The sheet link is" +sheetLink);

              const sheetId = SpreadsheetApp.openByUrl(sheetLink).getId();
              console.log("The Sheet Id is :-" +sheetId);

              const file = DriveApp.getFileById(sheetId);

              
              // Archive folder logic
              const parentFolderId = file.getParents().next().getId();
              
              const archiveFolderId = createFolderIfNotExists(parentFolderId, "Exited SME Sheets");

              const archiveFolder = DriveApp.getFolderById(archiveFolderId);
              
              file.moveTo(archiveFolder);
              status += `SME File moved to Archived\n`;


              // Remove from index sheet
              const foundIndex = this.indexData.findIndex(r => r[this.indexHeaders["SME Email"]] === emailId && r[this.indexHeaders["Sheet Link"]] === sheetLink);
              console.log("The found index is:-" +foundIndex);

              if(foundIndex!=-1) removedIndexes.push(foundIndex);
            }
            else{
              status +=`SME Already Archived\n`;
            }
          }
       });

       //clear input in the managment sheet
       applyCustomFormatting(this.smeSheet.getRange(inputSheetIndex + 18, 2, 1, 3)).clearContent();
      }
 });
    //  remove archived entries from index sheet
     removedIndexes.sort((a,b)=> b-a).forEach(index =>{
        this.indexData.splice(index,1);
     });

    //  clean and renumber data
   this.indexData = this.indexData.filter(row => row.some(cell => cell !== ""));
   //console.log("Index data is:-"+this.indexData); 

   for (let i = 0; i < this.indexData.length; i++) {
    this.indexData[i][this.indexHeaders["#"]] = i + 1;
  }
   
    // Clear and update the index sheet
  this.indexSheet.getRange(3, 1, this.indexSheet.getLastRow() - 2, 4).clearContent();
  if (this.indexData.length > 0) {
    this.indexSheet.getRange(3, 1, this.indexData.length, 4).setValues(this.indexData);
  }
     
    SpreadsheetApp.getUi().alert(status); 
    // try {
    //     SpreadsheetApp.getUi().alert(status);
    //   } catch (e) {
    //     Logger.log(status);
    //  }
  }
}


function addSME() {
  const sme1 = new SMESheetManager(
    QA_HEAD_DASHBORD_ID,
    SME_MANAGEMENT_TAB_ID,
    INDEX_SME_TAB_ID,
    MASTER_DB_SPREADSHEET_ID,
    SME_DB_TAB_ID,
    3
  );
  sme1.smeAdd();
}


function removeSME() {
  const sme2 = new SMESheetManager(
    QA_HEAD_DASHBORD_ID,
    SME_MANAGEMENT_TAB_ID,
    INDEX_SME_TAB_ID,
    MASTER_DB_SPREADSHEET_ID,
    SME_DB_TAB_ID,
    16
  );
  sme2.smeRemove();
}


function setValuesQAReviewSheet(backendSheet, qaReviewSheet){
  const smeNamesList = backendSheet.getRange(2, 1, backendSheet.getLastRow()).getValues().flat().filter(Boolean);
  const smeDropdownCell = qaReviewSheet.getRange(1, 4);
  smeDropdownCell.clearContent();
  smeDropdownCell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(smeNamesList).setAllowInvalid(false).build());

  const timeStamp = new Date();
  let  startDate = new Date(timeStamp.getTime() - 7 * 24 * 60 * 60 * 1000);
  const endDate = Utilities.formatDate(timeStamp, Session.getScriptTimeZone(), 'dd-MMM-yy');
  startDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'dd-MMM-yy');
  const startDateCell = qaReviewSheet.getRange(1, 7);
  const endDateCell = qaReviewSheet.getRange(2, 7);
  startDateCell.clearContent();
  startDateCell.setValue(startDate)
  endDateCell.clearContent();
  endDateCell.setValue(endDate);
}

function testFunction(){
  smeBackend("https://docs.google.com/spreadsheets/d/15kjqzaEIm5OmiSwLibUjW-_LfQlO45qkn4bkp2OEdDc/edit", "mrinmoyee.maity@upthink.com")
}


const masterDBSME = SpreadsheetApp.openById("1u4i5rMCWM0mSgDWdf6EmGVi40kb9i2Qx51SE1cupIjU").getSheetByName("SME DB");


function smeBackend(url, smeEmail){
  const spreadsheet = SpreadsheetApp.openByUrl(url)
  const sheetId = spreadsheet.getId();
  console.log("Sheet Id: ", sheetId);
  const backendSheet = spreadsheet.getSheetByName('Backend');
  const qaReviewSheet = spreadsheet.getSheetByName("QA Review");
  
  
  const masterDataRange = masterDBSME.getRange(1, 1, masterDBSME.getLastRow(), masterDBSME.getLastColumn()).getValues();
  const masterHeader = masterDataRange[0], masterData = masterDataRange.slice(1);

  const inputIndices = {
    srNo : masterHeader.indexOf("Sr. No."),
    emailId : masterHeader.indexOf("Email ID"),
    department : masterHeader.indexOf("Department"),
    sme : masterHeader.indexOf("SME Name"),
    reviewer : masterHeader.indexOf("QA Reviewer"),
    grade : masterHeader.indexOf("Grade"),
    designation : masterHeader.indexOf("Designation"),
    reportingManager : masterHeader.indexOf("Reporting Manager"),
    addedDateIdx : masterHeader.indexOf("Added Date"),
    activeIdx : masterHeader.indexOf("Active?"),
    removedDateIdx : masterHeader.indexOf("Removed Date"),
    smeSheetLink : masterHeader.indexOf("SME Sheet Link"),
  }

  let smeList = []
  //Check which departments does not have others
  let uniqueDepartments = [];

  // Create a map to keep track of departments with 'Others' entries
  let departmentsWithOthers = {};

  masterData.forEach(r => {
    if (r[inputIndices.grade] === 'Others') {
      departmentsWithOthers[r[inputIndices.department]] = true;
    }
  });

  masterData.forEach(r => {
    if (!departmentsWithOthers[r[inputIndices.department]]) {
      if (!uniqueDepartments.includes(r[inputIndices.department])) {
        uniqueDepartments.push(r[inputIndices.department]);
      }
    }
  });

  
  let grade, department, technicalDeisgnation;
  masterData.forEach(r=>{
    if(r[inputIndices.emailId] === smeEmail){
      grade = r[inputIndices.grade];
      technicalDeisgnation = r[inputIndices.designation].split("-")[0].trim().toLowerCase();
      department = r[inputIndices.department];
    }
  })

  smeList.push(masterData.filter(r=>r[inputIndices.emailId] === smeEmail).map(r=>r[inputIndices.sme])[0]);


  // ONLY SMEs WHOSE ADDED DATE IS PRESENT OR WHOSE REMOVED DATE IS GREATER THAN ADDED DATE
  // CHANGE THIS - NEW ORG STRUCTURE

  masterData.forEach(r =>{
    if(r=> (r[inputIndices.removedDateIdx]==='' && r[inputIndices.activeIdx]===false) || 
            (r[inputIndices.activeIdx]===true && r[inputIndices.removedDateIdx] !=='')){
      if(department === r[inputIndices.department] && department !== 'Others'){
        if(grade === 'U1' && technicalDeisgnation === "associate"){
          if(r[inputIndices.grade] === 'U1' && r[inputIndices.designation].split("-")[0].trim().toLowerCase() === "junior associate"){
            smeList.push(r[inputIndices.sme]);     
          }
        }else if(grade === 'U2' && technicalDeisgnation === "subject matter expert"){
          if((r[inputIndices.grade] === 'U1' && r[inputIndices.designation].split("-")[0].trim().toLowerCase() === "associate") || 
              (r[inputIndices.grade] === 'U1' && r[inputIndices.designation].split("-")[0].trim().toLowerCase() === "junior associate")){
            smeList.push(r[inputIndices.sme])
          }
        }else if (grade === 'U2' && technicalDeisgnation === "senior subject matter expert"){
          if((r[inputIndices.grade] === 'U2' && r[inputIndices.designation].split("-")[0].trim().toLowerCase() === "subject matter expert") || 
              (r[inputIndices.grade] === 'U1' && r[inputIndices.designation].split("-")[0].trim().toLowerCase() === "associate") || 
              (r[inputIndices.grade] === 'U1' && r[inputIndices.designation].split("-")[0].trim().toLowerCase() === "junior associate")){
            smeList.push(r[inputIndices.sme]);
          }
        }else if(grade === 'U3'){
          if(r[inputIndices.grade]==='U1' || r[inputIndices.grade]==='U2' || r[inputIndices.grade]==='U3'){
            smeList.push(r[inputIndices.sme])
          }
        }
      }else if (department === 'Others' && grade === 'U3') {
        if (r[inputIndices.grade] === 'U1' || r[inputIndices.grade] === 'U2' || r[inputIndices.grade] === 'U3') {
          const uniqueSMEs = new Set(); // Use a Set to store unique SME names
          uniqueDepartments.forEach(uniqueDept => {
            masterData.forEach(innerR => {
              if (innerR[inputIndices.department] === uniqueDept && innerR[inputIndices.grade] !== 'Others') {
                uniqueSMEs.add(innerR[inputIndices.sme]); // Add SME name to the Set
              }
            });
          });

          // Convert the Set back to an array
          smeList = Array.from(uniqueSMEs);
        }
      }
    }
  });


  clearRowsBelow(backendSheet, 1)
  let rowIndex = 2;
  smeList.forEach(name => {
    backendSheet.getRange(rowIndex, 1).setValue(name);
    rowIndex ++; 
  });

  // Protect backend sheet
  defineProtectionFunctionalities(backendSheet, hide=true);
  setValuesQAReviewSheet(backendSheet, qaReviewSheet);
  if (spreadsheet.getOwner().getEmail() !== AUTOMATION_EMAIL)
    DriveApp.getFileById(sheetId).setOwner(AUTOMATION_EMAIL)
  spreadsheet.addEditor(EDITOR_EMAIL)
}


/*_____________________________________OLD CPODE addsme()_________________________________________ */
// function addSME() {
//   const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
//   const sheet = spreadsheet.getSheetByName("SME Management");
//   const dataRange = sheet.getRange(4, 1, 11, 4).getValues();
//   const headers = dataRange[0], data = dataRange.slice(1);

//   const indexSheet = spreadsheet.getSheetByName("Index - SME");
//   const indexSheetRange = indexSheet.getRange(2, 1, indexSheet.getLastRow(), indexSheet.getLastColumn()).getValues();
//   const indexHeader = indexSheetRange[0], indexData = indexSheetRange.slice(1);

//   const smeBackendSheet = SpreadsheetApp.openById("1eJ-7nKeICwYXEBh4oEzLwC6SpkYzCw5ntYxarbD2jxA").getSheetByName("SME DB");
//   const backendDataRange = smeBackendSheet.getRange(1, 1, smeBackendSheet.getLastRow(), smeBackendSheet.getLastColumn()).getValues();
//   const backendHeaders = backendDataRange[0], backendData = backendDataRange.slice(1);


//   const sheetIndices = {
//     emailIdx : headers.indexOf("Email ID"),
//     departmentIdx : headers.indexOf("Department"),
//     addIdx : headers.indexOf("Add?")
//   }

//   const indexIndices = {
//     indexSrNo : indexHeader.indexOf("#"),
//     smeEmail : indexHeader.indexOf("SME Email"),
//     indexDepartmentIdx : indexHeader.indexOf("Department"),
//     indexSheetLink : indexHeader.indexOf("Sheet Link"),
//   }

//   const backendSheetIndices = {
//     emailIdx : backendHeaders.indexOf("Email ID"),
//     departmentIdx : backendHeaders.indexOf("Department"),
//     smeNameIdx : backendHeaders.indexOf("SME Name"),
//     category : backendHeaders.indexOf("Pyramid Category"),
//     addedDateIdx : backendHeaders.indexOf("Added Date"),
//     activeIdx : backendHeaders.indexOf("Active?"),
//     sheetCreationDateIdx : backendHeaders.indexOf("Sheet creation date"),
//     sheetLinkIdx : backendHeaders.indexOf("SME Sheet Link"),
//     tlIdx : backendHeaders.indexOf("TL"),
//     smeCIdx : backendHeaders.indexOf("SME C"),
//     smeBIdx : backendHeaders.indexOf("SME B"),
//   }


//   let status = "";
//   const emailLinkArray = [];
//   // Get the data from the addSME range.
//   data.forEach((row, inputSheetIndex) =>{
//     if(row[sheetIndices.addIdx] === true){
//       const emailId = row[sheetIndices.emailIdx];
//       const department = row[sheetIndices.departmentIdx];
//       backendData.forEach((r, index) => {
//         if(r[backendSheetIndices.emailIdx] === emailId && (r[backendSheetIndices.departmentIdx] === department || r[backendSheetIndices.departmentIdx] ==='Others')){  // CHECK THIS
//           // Check active or not
//           const rowIndex = index + 2;
//           const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM-yy");
//           if(r[backendSheetIndices.activeIdx] === true){
//             status += "User already active" + "\n";
//           }else{  // If user not active
//             smeBackendSheet.getRange(rowIndex, backendSheetIndices.activeIdx + 1).setValue(true);
//             smeBackendSheet.getRange(rowIndex, backendSheetIndices.addedDateIdx + 1).setValue(today);
//             if(r[backendSheetIndices.sheetLinkIdx] === ''){
//               //Create a sheet and write the sheet link and addsheet creation date.
//               const name = r[backendSheetIndices.smeNameIdx].split(" ");
//               const sheetName = name.length > 1 ? `QA_SME_${name[0]}_${name[name.length-1]}` : `QA_SME_${name[0]}`;
//               const sheetId = createCopyOfSpreadsheet("1qY7TlB_PwO0HptiOoMjhcQ-fLsa7zAwAUeZopApmGO4", sheetName, department, "SME Sheets");
//               const sheetLink = reviewerSheetLink = `https://docs.google.com/spreadsheets/d/${sheetId}/edit`;
//               smeBackend(sheetLink, emailId)
//               smeBackendSheet.getRange(rowIndex, backendSheetIndices.sheetCreationDateIdx + 1).setValue(today);
//               smeBackendSheet.getRange(rowIndex, backendSheetIndices.sheetLinkIdx + 1).setValue(sheetLink);
//               emailLinkArray.push([emailId, department, sheetLink]);
//               status += `A new sheet for SME ${name.join(" ")} has been created` + "\n";
//               DriveApp.getFileById(sheetId).addEditor(emailId); // Needs to change
//             }else{
//               smeBackendSheet.getRange(rowIndex, backendSheetIndices.addedDateIdx + 1).setValue(today);
//               smeBackendSheet.getRange(rowIndex, backendSheetIndices.activeIdx + 1).setValue(true);
//               const sheetLink = smeBackendSheet.getRange(rowIndex, backendSheetIndices.sheetLinkIdx+1).getValue();
//               const sheetId = SpreadsheetApp.openByUrl(sheetLink).getId()
//               // Move the file to ex reviewer
//               DriveApp.getFileById(sheetId).addEditor(emailId);
//               const fileinFolder = DriveApp.getFileById(sheetId).getParents();
//               const folderId = fileinFolder.next().getId();
//               const parentFolder = DriveApp.getFolderById(folderId).getParents();
//               const parentFolderId = parentFolder.next().getId();
//               const folder = DriveApp.getFolderById(parentFolderId);
//               DriveApp.getFileById(sheetId).moveTo(folder);
//               DriveApp.getFileById(sheetId).moveTo(folder);
//               status += "SME moved from archieved folder to the main folder" + "\n";
//               emailLinkArray.push([emailId, department, sheetLink]);
//               DriveApp.getFileById(sheetId).addEditor(emailId); // Needs to change
//             }
//           }
//         }
//       })
      
//       applyCustomFormatting(sheet.getRange(inputSheetIndex+5, 2, 1, 3)).clearContent();
//     }
//   })

//   let lastRow = indexSheet.getLastRow();
  
//   emailLinkArray.forEach(r => {
//     if(lastRow === 2){
//       indexSheet.getRange(lastRow + 1, indexIndices.indexSrNo+1).setValue(1);
//       indexSheet.getRange(lastRow + 1, indexIndices.smeEmail+1).setValue(r[0]);
//       indexSheet.getRange(lastRow + 1, indexIndices.indexDepartmentIdx+1).setValue(r[1]);
//       indexSheet.getRange(lastRow + 1, indexIndices.indexSheetLink+1).setValue(r[2]);
//       lastRow++;
//     }else{
//       // Get the last row Index
//       const lastSrNo = indexSheet.getRange(lastRow, indexIndices.indexSrNo+1).getValue();
//       indexSheet.getRange(lastRow + 1, indexIndices.indexSrNo+1).setValue(lastSrNo + 1);
//       indexSheet.getRange(lastRow + 1, indexIndices.smeEmail+1).setValue(r[0]);
//       indexSheet.getRange(lastRow + 1, indexIndices.indexDepartmentIdx+1).setValue(r[1]);
//       indexSheet.getRange(lastRow + 1, indexIndices.indexSheetLink+1).setValue(r[2]);
//       lastRow++;
//     }
//   })

//   SpreadsheetApp.getUi().alert(status);
// }

/*________________________________________________________________________________________________ */





/*_______________________________________________OLD CODE Removesme()___________________________*/
// function removeSME(){
//   const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
//   const inputSheet = spreadsheet.getSheetByName("SME Management");
//   const indexSheet = spreadsheet.getSheetByName("Index - SME")
//   const dataRange = inputSheet.getRange(17, 1, 6, 4).getValues();
//   const headers = dataRange[0], data = dataRange.slice(1);

//   const smeBackendSheet = SpreadsheetApp.openById("1eJ-7nKeICwYXEBh4oEzLwC6SpkYzCw5ntYxarbD2jxA").getSheetByName("SME DB");
//   const backendDataRange = smeBackendSheet.getRange(1, 1, smeBackendSheet.getLastRow(), smeBackendSheet.getLastColumn()).getValues();
//   const backendHeaders = backendDataRange[0], backendData = backendDataRange.slice(1);

//   const indexDataRange = indexSheet.getRange(2, 1, indexSheet.getLastRow(), 4).getValues();
//   const indexHeader = indexDataRange[0];
//   let indexData = indexDataRange.slice(1);

//   const sheetIndices = {
//     emailIdx : headers.indexOf("Email ID"),
//     departmentIdx : headers.indexOf("Department"),
//     removeIdx : headers.indexOf("Remove?")
//   }

//   const indexSrNo = indexHeader.indexOf("#");
//   const indexSmeEmail = indexHeader.indexOf("SME Email");
//   const indexDepartmentIdx = indexHeader.indexOf("Department");
//   const indexSheetLink = indexHeader.indexOf("Sheet Link");

//   const backendSheetIndices = {
//     emailIdx : backendHeaders.indexOf("Email ID"),
//     departmentIdx : backendHeaders.indexOf("Department"),
//     smeNameIdx : backendHeaders.indexOf("SME Name"),
//     category : backendHeaders.indexOf("Pyramid Category"),
//     removedDateIdx : backendHeaders.indexOf("Removed Date"),
//     activeIdx : backendHeaders.indexOf("Active?"),
//     sheetCreationDateIdx : backendHeaders.indexOf("Sheet creation date"),
//     sheetLinkIdx : backendHeaders.indexOf("SME Sheet Link"),
//     tlIdx : backendHeaders.indexOf("TL"),
//     smeCIdx : backendHeaders.indexOf("SME C"),
//     smeBIdx : backendHeaders.indexOf("SME B"),
//   }

//   const removedIndexes = [];
//   let status = "";

//   const today = Utilities.formatDate(new Date, Session.getScriptTimeZone(), 'dd-MMM-yy');
//   data.forEach((row, inputSheetIndex) => {
//     if(row[sheetIndices.removeIdx] === true){
//       const emailId = row[sheetIndices.emailIdx];
//       const department = row[sheetIndices.departmentIdx];
//       backendData.forEach((r, index) => {
//         if(emailId === r[backendSheetIndices.emailIdx] && department === r[backendSheetIndices.departmentIdx]){
//           if(r[backendSheetIndices.activeIdx] === true){
//             const rowIndex = index + 2;
//             smeBackendSheet.getRange(rowIndex, backendSheetIndices.activeIdx + 1).setValue(false);
//             smeBackendSheet.getRange(rowIndex, backendSheetIndices.removedDateIdx + 1).setValue(today);
//             const sheetLink = smeBackendSheet.getRange(rowIndex, backendSheetIndices.sheetLinkIdx+1).getValue();
//             const sheetId = SpreadsheetApp.openByUrl(sheetLink).getId()
//             // Move the file to ex reviewer
//             // const folder = DriveApp.getFolderById("1kkSbDnMiOPxHzaviNMmciyXqMhzvZFUD");
//             const folders = DriveApp.getFileById(sheetId).getParents();
//             const parentFolderId = folders.next().getId();
//             const archieveFolderId = createFolderIfNotExists(parentFolderId, "Exited SME Sheets");
//             const folder = DriveApp.getFolderById(archieveFolderId);
//             DriveApp.getFileById(sheetId).moveTo(folder);
//             DriveApp.getFileById(sheetId).moveTo(folder);
//             status += "SME File moved to Archieved" + "\n"
//             const foundIndex = indexData.findIndex(r => r[indexSmeEmail] === emailId && r[indexSheetLink] === sheetLink);
//             if (foundIndex !== -1)
//               removedIndexes.push(foundIndex);
//             // removedIndexes.push([emailId, department, sheetLink])
//           }else{
//             status += "SME Already Archieved" + "\n"
            
//           }  
//         }
//       })
//       applyCustomFormatting(inputSheet.getRange(inputSheetIndex+18, 2, 1, 3)).clearContent();
//     }
//   })

//   // Delete rows in reverse order to avoid shifting issues
//   removedIndexes.sort((a, b) => b - a).forEach(index => {
//     indexData.splice(index, 1);
//   });

//   // Remove empty rows from indexData
//   indexData = indexData.filter(row => row.some(cell => cell !== ''));

//   // Remove empty rows from indexData
//   indexData = indexData.filter(row => row.some(cell => cell !== ''));

//   // Update the serial numbers in index sheet
//   for (let i = 0; i < indexData.length; i++) {
//     indexData[i][indexSrNo] = i + 1;
//   }

//   // Clear the existing data in the index sheet, starting from row 3
//   indexSheet.getRange(3, 1, indexSheet.getLastRow() - 1, 4).clearContent();

//   // Write the updated data back to the index sheet
//   if (indexData.length > 0) {
//     indexSheet.getRange(3, 1, indexData.length, 4).setValues(indexData);
//   }
//   SpreadsheetApp.getUi().alert(status);
// }












