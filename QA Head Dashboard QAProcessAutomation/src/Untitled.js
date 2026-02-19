// function onOpen() {
//   var ui = SpreadsheetApp.getUi();
//   ui.createMenu('Custom Menu')
//       .addItem('Run SME Functions 1', 'addSMEFromList1')
//       .addItem('Run SME Functions 2', 'addSMEFromList2')
//       .addItem('Run SME Functions 3', 'addSMEFromList3')
//       // .addItem('Run Function 2', 'oneTimeReviewerSheet2')
//       .addToUi();
// };

function oneTimeReviewerSheet1(){
  const sheet = SpreadsheetApp.openById("1u4i5rMCWM0mSgDWdf6EmGVi40kb9i2Qx51SE1cupIjU").getSheetByName("Reviewer DB"); //1eJ-7nKeICwYXEBh4oEzLwC6SpkYzCw5ntYxarbD2jxA
  const inputDataRange = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  const inputHeaders = inputDataRange[0], inputData = inputDataRange.slice(1);

  const inputIndices = {
    emailIdx : inputHeaders.indexOf("Email ID"),
    departmentIdx : inputHeaders.indexOf("Department"),
    activeIdx : inputHeaders.indexOf("Active?"),
    srNoIdx : inputHeaders.indexOf("#"),
  }

  console.log("Input indices",inputIndices);
  const indexSheet = SpreadsheetApp.openById("1S8BI536OII-foUX3CjvsJMZui_OVCNjVigB9SNxaE5o").getSheetByName("Index"); //1FrKuzuyN6uo1UP41MAvD9zLvPUMNTOzBYz4iwHl7VhY
  const indexDataRange = indexSheet.getRange(2, 1, indexSheet.getLastRow(), 4).getValues();
  const indexHeader = indexDataRange[0];
  let indexData = indexDataRange.slice(1);

  const indexSrNo = indexHeader.indexOf("#");
  const indexReviewerEmail = indexHeader.indexOf("QA Reviewer Email");
  const indexDepartmentIdx = indexHeader.indexOf("Department");
  const indexSheetLink = indexHeader.indexOf("Sheet Link");


  const emailLinkArray = [];
  inputData.forEach(r => {
    if(r[inputIndices.activeIdx] !== true && r[inputIndices.srNoIdx] <= 20){
      const department = r[inputIndices.departmentIdx], emailId = r[inputIndices.emailIdx];
      const [reviewerSheetLink, reviewerSheetId, sheetID, reviewerName] = setReviewerSheet_CreateNewSheet(emailId, department);
      setValuesAtStart(reviewerSheetId, department, sheetID, emailId);
      //createBackendReviwerSheetByDepartment(reviewerSheetLink, department); 
      createBackendReviewerSheet(reviewerSheetLink,department)
      emailLinkArray.push([emailId, department, reviewerSheetLink]);
    }
  })

  let lastRow = indexSheet.getLastRow();
  emailLinkArray.forEach(r => {
    if(lastRow === 2){
      indexSheet.getRange(lastRow + 1, indexSrNo+1).setValue(1);
      indexSheet.getRange(lastRow + 1, indexReviewerEmail+1).setValue(r[0]);
      indexSheet.getRange(lastRow + 1, indexDepartmentIdx+1).setValue(r[1]);
      indexSheet.getRange(lastRow + 1, indexSheetLink+1).setValue(r[2]);
      lastRow++;
    }else{
      // Get the last row Index
      const lastSrNo = indexSheet.getRange(lastRow, indexSrNo+1).getValue();
      indexSheet.getRange(lastRow + 1, indexSrNo+1).setValue(lastSrNo + 1);
      indexSheet.getRange(lastRow + 1, indexReviewerEmail+1).setValue(r[0]);
      indexSheet.getRange(lastRow + 1, indexDepartmentIdx+1).setValue(r[1]);
      indexSheet.getRange(lastRow + 1, indexSheetLink+1).setValue(r[2]);
      lastRow++;
    }
  })
}


function oneTimeReviewerSheet2(){
  const sheet = SpreadsheetApp.openById("1u4i5rMCWM0mSgDWdf6EmGVi40kb9i2Qx51SE1cupIjU").getSheetByName("Reviewer DB"); 
  //"1eJ-7nKeICwYXEBh4oEzLwC6SpkYzCw5ntYxarbD2jxA"
  const inputDataRange = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  const inputHeaders = inputDataRange[0], inputData = inputDataRange.slice(1);

  const inputIndices = {
    emailIdx : inputHeaders.indexOf("Email ID"),
    departmentIdx : inputHeaders.indexOf("Department"),
    activeIdx : inputHeaders.indexOf("Active?"),
    srNoIdx : inputHeaders.indexOf("#"),
  }
  

  const indexSheet = SpreadsheetApp.openById("1S8BI536OII-foUX3CjvsJMZui_OVCNjVigB9SNxaE5o").getSheetByName("Index");
  //"1FrKuzuyN6uo1UP41MAvD9zLvPUMNTOzBYz4iwHl7VhY"
  const indexDataRange = indexSheet.getRange(2, 1, indexSheet.getLastRow(), 4).getValues();
  const indexHeader = indexDataRange[0];
  let indexData = indexDataRange.slice(1);

  const indexSrNo = indexHeader.indexOf("#");
  const indexReviewerEmail = indexHeader.indexOf("QA Reviewer Email");
  const indexDepartmentIdx = indexHeader.indexOf("Department");
  const indexSheetLink = indexHeader.indexOf("Sheet Link");


  const emailLinkArray = [];
  inputData.forEach(r => {
    if(r[inputIndices.activeIdx] !== true && r[inputIndices.srNoIdx] > 20){
      const department = r[inputIndices.departmentIdx], emailId = r[inputIndices.emailIdx];
      const [reviewerSheetLink, reviewerSheetId, sheetID, reviewerName] = setReviewerSheet_CreateNewSheet(emailId, department);
      setValuesAtStart(reviewerSheetId, department, sheetID, emailId);
     // createBackendReviwerSheetByDepartment(reviewerSheetLink, department);
      createBackendReviewerSheet(reviewerSheetLink,department)

      emailLinkArray.push([emailId, department, reviewerSheetLink]);
    }
  })

  let lastRow = indexSheet.getLastRow();
  emailLinkArray.forEach(r => {
    if(lastRow === 2){
      indexSheet.getRange(lastRow + 1, indexSrNo+1).setValue(1);
      indexSheet.getRange(lastRow + 1, indexReviewerEmail+1).setValue(r[0]);
      indexSheet.getRange(lastRow + 1, indexDepartmentIdx+1).setValue(r[1]);
      indexSheet.getRange(lastRow + 1, indexSheetLink+1).setValue(r[2]);
      lastRow++;
    }else{
      // Get the last row Index
      const lastSrNo = indexSheet.getRange(lastRow, indexSrNo+1).getValue();
      indexSheet.getRange(lastRow + 1, indexSrNo+1).setValue(lastSrNo + 1);
      indexSheet.getRange(lastRow + 1, indexReviewerEmail+1).setValue(r[0]);
      indexSheet.getRange(lastRow + 1, indexDepartmentIdx+1).setValue(r[1]);
      indexSheet.getRange(lastRow + 1, indexSheetLink+1).setValue(r[2]);
      lastRow++;
    }
  })
}



function addSMEFromList1() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // const sheet = spreadsheet.getSheetByName("SME DB");
  // const dataRange = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  // const headers = dataRange[0], data = dataRange.slice(1);

  const indexSheet = spreadsheet.getSheetByName("Index - SME");
  const indexSheetRange = indexSheet.getRange(2, 1, indexSheet.getLastRow(), indexSheet.getLastColumn()).getValues();
  const indexHeader = indexSheetRange[0], indexData = indexSheetRange.slice(1);

  const smeBackendSheet = SpreadsheetApp.openById("1u4i5rMCWM0mSgDWdf6EmGVi40kb9i2Qx51SE1cupIjU").getSheetByName("SME DB");
  //"1eJ-7nKeICwYXEBh4oEzLwC6SpkYzCw5ntYxarbD2jxA"
  const backendDataRange = smeBackendSheet.getRange(1, 1, smeBackendSheet.getLastRow(), smeBackendSheet.getLastColumn()).getValues();
  const backendHeaders = backendDataRange[0], backendData = backendDataRange.slice(1);


  // const sheetIndices = {
  //   emailIdx : headers.indexOf("Email ID"),
  //   departmentIdx : headers.indexOf("Department"),
  //   addIdx : headers.indexOf("Add?")
  // }

  const indexIndices = {
    indexSrNo : indexHeader.indexOf("#"),
    smeEmail : indexHeader.indexOf("SME Email"),
    indexDepartmentIdx : indexHeader.indexOf("Department"),
    indexSheetLink : indexHeader.indexOf("Sheet Link"),
  }

  const backendSheetIndices = {
    emailIdx : backendHeaders.indexOf("Email ID"),
    departmentIdx : backendHeaders.indexOf("Department"),
    smeNameIdx : backendHeaders.indexOf("SME Name"),
    category : backendHeaders.indexOf("Pyramid Category"),
    addedDateIdx : backendHeaders.indexOf("Added Date"),
    activeIdx : backendHeaders.indexOf("Active?"),
    sheetCreationDateIdx : backendHeaders.indexOf("Sheet creation date"),
    sheetLinkIdx : backendHeaders.indexOf("SME Sheet Link"),
    tlIdx : backendHeaders.indexOf("TL"),
    smeCIdx : backendHeaders.indexOf("SME C"),
    smeBIdx : backendHeaders.indexOf("SME B"),
  }


  let status = "";
  const emailLinkArray = [];
  // Get the data from the addSME range.
  backendData.forEach((row, inputSheetIndex) =>{
    if(inputSheetIndex <= 80){
      const emailId = row[backendSheetIndices.emailIdx];
      const department = row[backendSheetIndices.departmentIdx];
      backendData.forEach((r, index) => {
        if(r[backendSheetIndices.emailIdx] === emailId && (r[backendSheetIndices.departmentIdx] === department || r[backendSheetIndices.departmentIdx] ==='Others')){  // CHECK THIS
          // Check active or not
          const rowIndex = index + 2;
          const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM-yy");
          if(r[backendSheetIndices.activeIdx] === true){
            status += "User already active" + "\n";
          }else{  // If user not active
            smeBackendSheet.getRange(rowIndex, backendSheetIndices.activeIdx + 1).setValue(true);
            smeBackendSheet.getRange(rowIndex, backendSheetIndices.addedDateIdx + 1).setValue(today);
            if(r[backendSheetIndices.sheetLinkIdx] === ''){
              //Create a sheet and write the sheet link and addsheet creation date.
              const name = r[backendSheetIndices.smeNameIdx].split(" ");
              const sheetName = name.length > 1 ? `QA_SME_${name[0]}_${name[name.length-1]}` : `QA_SME_${name[0]}`;
              const sheetId = createCopyOfSpreadsheet("1qY7TlB_PwO0HptiOoMjhcQ-fLsa7zAwAUeZopApmGO4", sheetName, department, "SME Sheets");
              const sheetLink = reviewerSheetLink = `https://docs.google.com/spreadsheets/d/${sheetId}/edit`;
              smeBackend(sheetLink, emailId)
              smeBackendSheet.getRange(rowIndex, backendSheetIndices.sheetCreationDateIdx + 1).setValue(today);
              smeBackendSheet.getRange(rowIndex, backendSheetIndices.sheetLinkIdx + 1).setValue(sheetLink);
              emailLinkArray.push([emailId, department, sheetLink]);
              status += `A new sheet for SME ${name.join(" ")} has been created` + "\n";
              DriveApp.getFileById(sheetId).addEditor(emailId); // emailId - Needs to change
            }else{
              smeBackendSheet.getRange(rowIndex, backendSheetIndices.addedDateIdx + 1).setValue(today);
              smeBackendSheet.getRange(rowIndex, backendSheetIndices.activeIdx + 1).setValue(true);
              const sheetLink = smeBackendSheet.getRange(rowIndex, backendSheetIndices.sheetLinkIdx+1).getValue();
              const sheetId = SpreadsheetApp.openByUrl(sheetLink).getId()
              // Move the file to ex reviewer
              DriveApp.getFileById(sheetId).addEditor(emailId);
              const fileinFolder = DriveApp.getFileById(sheetId).getParents();
              const folderId = fileinFolder.next().getId();
              const parentFolder = DriveApp.getFolderById(folderId).getParents();
              const parentFolderId = parentFolder.next().getId();
              const folder = DriveApp.getFolderById(parentFolderId);
              DriveApp.getFileById(sheetId).moveTo(folder);
              DriveApp.getFileById(sheetId).moveTo(folder);
              status += "SME moved from archieved folder to the main folder" + "\n";
              emailLinkArray.push([emailId, department, sheetLink]);
              DriveApp.getFileById(sheetId).addEditor(emailId); // Needs to change
            }
          }
        }
      })
    }
      // applyCustomFormatting(sheet.getRange(inputSheetIndex+5, 2, 1, 3)).clearContent();
  })

  let lastRow = indexSheet.getLastRow();
  
  emailLinkArray.forEach(r => {
    if(lastRow === 2){
      indexSheet.getRange(lastRow + 1, indexIndices.indexSrNo+1).setValue(1);
      indexSheet.getRange(lastRow + 1, indexIndices.smeEmail+1).setValue(r[0]);
      indexSheet.getRange(lastRow + 1, indexIndices.indexDepartmentIdx+1).setValue(r[1]);
      indexSheet.getRange(lastRow + 1, indexIndices.indexSheetLink+1).setValue(r[2]);
      lastRow++;
    }else{
      // Get the last row Index
      const lastSrNo = indexSheet.getRange(lastRow, indexIndices.indexSrNo+1).getValue();
      indexSheet.getRange(lastRow + 1, indexIndices.indexSrNo+1).setValue(lastSrNo + 1);
      indexSheet.getRange(lastRow + 1, indexIndices.smeEmail+1).setValue(r[0]);
      indexSheet.getRange(lastRow + 1, indexIndices.indexDepartmentIdx+1).setValue(r[1]);
      indexSheet.getRange(lastRow + 1, indexIndices.indexSheetLink+1).setValue(r[2]);
      lastRow++;
    }
  })

  SpreadsheetApp.getUi().alert(status);
}

function addSMEFromList2() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // const sheet = spreadsheet.getSheetByName("SME DB");
  // const dataRange = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  // const headers = dataRange[0], data = dataRange.slice(1);

  const indexSheet = spreadsheet.getSheetByName("Index - SME");
  const indexSheetRange = indexSheet.getRange(2, 1, indexSheet.getLastRow(), indexSheet.getLastColumn()).getValues();
  const indexHeader = indexSheetRange[0], indexData = indexSheetRange.slice(1);

  const smeBackendSheet = SpreadsheetApp.openById("1u4i5rMCWM0mSgDWdf6EmGVi40kb9i2Qx51SE1cupIjU").getSheetByName("SME DB");
  //"1eJ-7nKeICwYXEBh4oEzLwC6SpkYzCw5ntYxarbD2jxA"
  const backendDataRange = smeBackendSheet.getRange(1, 1, smeBackendSheet.getLastRow(), smeBackendSheet.getLastColumn()).getValues();
  const backendHeaders = backendDataRange[0], backendData = backendDataRange.slice(1);


  // const sheetIndices = {
  //   emailIdx : headers.indexOf("Email ID"),
  //   departmentIdx : headers.indexOf("Department"),
  //   addIdx : headers.indexOf("Add?")
  // }

  const indexIndices = {
    indexSrNo : indexHeader.indexOf("#"),
    smeEmail : indexHeader.indexOf("SME Email"),
    indexDepartmentIdx : indexHeader.indexOf("Department"),
    indexSheetLink : indexHeader.indexOf("Sheet Link"),
  }

  const backendSheetIndices = {
    emailIdx : backendHeaders.indexOf("Email ID"),
    departmentIdx : backendHeaders.indexOf("Department"),
    smeNameIdx : backendHeaders.indexOf("SME Name"),
    category : backendHeaders.indexOf("Pyramid Category"),
    addedDateIdx : backendHeaders.indexOf("Added Date"),
    activeIdx : backendHeaders.indexOf("Active?"),
    sheetCreationDateIdx : backendHeaders.indexOf("Sheet creation date"),
    sheetLinkIdx : backendHeaders.indexOf("SME Sheet Link"),
    tlIdx : backendHeaders.indexOf("TL"),
    smeCIdx : backendHeaders.indexOf("SME C"),
    smeBIdx : backendHeaders.indexOf("SME B"),
  }


  let status = "";
  const emailLinkArray = [];
  // Get the data from the addSME range.
  backendData.forEach((row, inputSheetIndex) =>{
    if(inputSheetIndex > 80 && inputSheetIndex <= 160){
      const emailId = row[backendSheetIndices.emailIdx];
      const department = row[backendSheetIndices.departmentIdx];
      backendData.forEach((r, index) => {
        if(r[backendSheetIndices.emailIdx] === emailId && (r[backendSheetIndices.departmentIdx] === department || r[backendSheetIndices.departmentIdx] ==='Others')){  // CHECK THIS
          // Check active or not
          const rowIndex = index + 2;
          const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM-yy");
          if(r[backendSheetIndices.activeIdx] === true){
            status += "User already active" + "\n";
          }else{  // If user not active
            smeBackendSheet.getRange(rowIndex, backendSheetIndices.activeIdx + 1).setValue(true);
            smeBackendSheet.getRange(rowIndex, backendSheetIndices.addedDateIdx + 1).setValue(today);
            if(r[backendSheetIndices.sheetLinkIdx] === ''){
              //Create a sheet and write the sheet link and addsheet creation date.
              const name = r[backendSheetIndices.smeNameIdx].split(" ");
              const sheetName = name.length > 1 ? `QA_SME_${name[0]}_${name[name.length-1]}` : `QA_SME_${name[0]}`;
              const sheetId = createCopyOfSpreadsheet("1qY7TlB_PwO0HptiOoMjhcQ-fLsa7zAwAUeZopApmGO4", sheetName, department, "SME Sheets");
              const sheetLink = reviewerSheetLink = `https://docs.google.com/spreadsheets/d/${sheetId}/edit`;
              smeBackend(sheetLink, emailId)
              smeBackendSheet.getRange(rowIndex, backendSheetIndices.sheetCreationDateIdx + 1).setValue(today);
              smeBackendSheet.getRange(rowIndex, backendSheetIndices.sheetLinkIdx + 1).setValue(sheetLink);
              emailLinkArray.push([emailId, department, sheetLink]);
              status += `A new sheet for SME ${name.join(" ")} has been created` + "\n";
              DriveApp.getFileById(sheetId).addEditor(emailId); // emailId - Needs to change
            }else{
              smeBackendSheet.getRange(rowIndex, backendSheetIndices.addedDateIdx + 1).setValue(today);
              smeBackendSheet.getRange(rowIndex, backendSheetIndices.activeIdx + 1).setValue(true);
              const sheetLink = smeBackendSheet.getRange(rowIndex, backendSheetIndices.sheetLinkIdx+1).getValue();
              const sheetId = SpreadsheetApp.openByUrl(sheetLink).getId()
              // Move the file to ex reviewer
              DriveApp.getFileById(sheetId).addEditor(emailId);
              const fileinFolder = DriveApp.getFileById(sheetId).getParents();
              const folderId = fileinFolder.next().getId();
              const parentFolder = DriveApp.getFolderById(folderId).getParents();
              const parentFolderId = parentFolder.next().getId();
              const folder = DriveApp.getFolderById(parentFolderId);
              DriveApp.getFileById(sheetId).moveTo(folder);
              DriveApp.getFileById(sheetId).moveTo(folder);
              status += "SME moved from archieved folder to the main folder" + "\n";
              emailLinkArray.push([emailId, department, sheetLink]);
              DriveApp.getFileById(sheetId).addEditor(emailId); // Needs to change
            }
          }
        }
      })
    }
      // applyCustomFormatting(sheet.getRange(inputSheetIndex+5, 2, 1, 3)).clearContent();
  })

  let lastRow = indexSheet.getLastRow();
  
  emailLinkArray.forEach(r => {
    if(lastRow === 2){
      indexSheet.getRange(lastRow + 1, indexIndices.indexSrNo+1).setValue(1);
      indexSheet.getRange(lastRow + 1, indexIndices.smeEmail+1).setValue(r[0]);
      indexSheet.getRange(lastRow + 1, indexIndices.indexDepartmentIdx+1).setValue(r[1]);
      indexSheet.getRange(lastRow + 1, indexIndices.indexSheetLink+1).setValue(r[2]);
      lastRow++;
    }else{
      // Get the last row Index
      const lastSrNo = indexSheet.getRange(lastRow, indexIndices.indexSrNo+1).getValue();
      indexSheet.getRange(lastRow + 1, indexIndices.indexSrNo+1).setValue(lastSrNo + 1);
      indexSheet.getRange(lastRow + 1, indexIndices.smeEmail+1).setValue(r[0]);
      indexSheet.getRange(lastRow + 1, indexIndices.indexDepartmentIdx+1).setValue(r[1]);
      indexSheet.getRange(lastRow + 1, indexIndices.indexSheetLink+1).setValue(r[2]);
      lastRow++;
    }
  })

  SpreadsheetApp.getUi().alert(status);
}



function addSMEFromList3() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // const sheet = spreadsheet.getSheetByName("SME DB");
  // const dataRange = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  // const headers = dataRange[0], data = dataRange.slice(1);

  const indexSheet = spreadsheet.getSheetByName("Index - SME");
  const indexSheetRange = indexSheet.getRange(2, 1, indexSheet.getLastRow(), indexSheet.getLastColumn()).getValues();
  const indexHeader = indexSheetRange[0], indexData = indexSheetRange.slice(1);

  const smeBackendSheet = SpreadsheetApp.openById("1u4i5rMCWM0mSgDWdf6EmGVi40kb9i2Qx51SE1cupIjU").getSheetByName("SME DB");
  //"1eJ-7nKeICwYXEBh4oEzLwC6SpkYzCw5ntYxarbD2jxA"
  const backendDataRange = smeBackendSheet.getRange(1, 1, smeBackendSheet.getLastRow(), smeBackendSheet.getLastColumn()).getValues();
  const backendHeaders = backendDataRange[0], backendData = backendDataRange.slice(1);


  // const sheetIndices = {
  //   emailIdx : headers.indexOf("Email ID"),
  //   departmentIdx : headers.indexOf("Department"),
  //   addIdx : headers.indexOf("Add?")
  // }

  const indexIndices = {
    indexSrNo : indexHeader.indexOf("#"),
    smeEmail : indexHeader.indexOf("SME Email"),
    indexDepartmentIdx : indexHeader.indexOf("Department"),
    indexSheetLink : indexHeader.indexOf("Sheet Link"),
  }

  const backendSheetIndices = {
    emailIdx : backendHeaders.indexOf("Email ID"),
    departmentIdx : backendHeaders.indexOf("Department"),
    smeNameIdx : backendHeaders.indexOf("SME Name"),
    category : backendHeaders.indexOf("Pyramid Category"),
    addedDateIdx : backendHeaders.indexOf("Added Date"),
    activeIdx : backendHeaders.indexOf("Active?"),
    sheetCreationDateIdx : backendHeaders.indexOf("Sheet creation date"),
    sheetLinkIdx : backendHeaders.indexOf("SME Sheet Link"),
    tlIdx : backendHeaders.indexOf("TL"),
    smeCIdx : backendHeaders.indexOf("SME C"),
    smeBIdx : backendHeaders.indexOf("SME B"),
  }


  let status = "";
  const emailLinkArray = [];
  // Get the data from the addSME range.
  backendData.forEach((row, inputSheetIndex) =>{
    if(inputSheetIndex > 160){
      const emailId = row[backendSheetIndices.emailIdx];
      const department = row[backendSheetIndices.departmentIdx];
      backendData.forEach((r, index) => {
        if(r[backendSheetIndices.emailIdx] === emailId && (r[backendSheetIndices.departmentIdx] === department || r[backendSheetIndices.departmentIdx] ==='Others')){  // CHECK THIS
          // Check active or not
          const rowIndex = index + 2;
          const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM-yy");
          if(r[backendSheetIndices.activeIdx] === true){
            status += "User already active" + "\n";
          }else{  // If user not active
            smeBackendSheet.getRange(rowIndex, backendSheetIndices.activeIdx + 1).setValue(true);
            smeBackendSheet.getRange(rowIndex, backendSheetIndices.addedDateIdx + 1).setValue(today);
            if(r[backendSheetIndices.sheetLinkIdx] === ''){
              //Create a sheet and write the sheet link and addsheet creation date.
              const name = r[backendSheetIndices.smeNameIdx].split(" ");
              const sheetName = name.length > 1 ? `QA_SME_${name[0]}_${name[name.length-1]}` : `QA_SME_${name[0]}`;
              const sheetId = createCopyOfSpreadsheet("1qY7TlB_PwO0HptiOoMjhcQ-fLsa7zAwAUeZopApmGO4", sheetName, department, "SME Sheets");
              const sheetLink = reviewerSheetLink = `https://docs.google.com/spreadsheets/d/${sheetId}/edit`;
              smeBackend(sheetLink, emailId)
              smeBackendSheet.getRange(rowIndex, backendSheetIndices.sheetCreationDateIdx + 1).setValue(today);
              smeBackendSheet.getRange(rowIndex, backendSheetIndices.sheetLinkIdx + 1).setValue(sheetLink);
              emailLinkArray.push([emailId, department, sheetLink]);
              status += `A new sheet for SME ${name.join(" ")} has been created` + "\n";
              DriveApp.getFileById(sheetId).addEditor(emailId); // emailId - Needs to change
            }else{
              smeBackendSheet.getRange(rowIndex, backendSheetIndices.addedDateIdx + 1).setValue(today);
              smeBackendSheet.getRange(rowIndex, backendSheetIndices.activeIdx + 1).setValue(true);
              const sheetLink = smeBackendSheet.getRange(rowIndex, backendSheetIndices.sheetLinkIdx+1).getValue();
              const sheetId = SpreadsheetApp.openByUrl(sheetLink).getId()
              // Move the file to ex reviewer
              DriveApp.getFileById(sheetId).addEditor(emailId);
              const fileinFolder = DriveApp.getFileById(sheetId).getParents();
              const folderId = fileinFolder.next().getId();
              const parentFolder = DriveApp.getFolderById(folderId).getParents();
              const parentFolderId = parentFolder.next().getId();
              const folder = DriveApp.getFolderById(parentFolderId);
              DriveApp.getFileById(sheetId).moveTo(folder);
              DriveApp.getFileById(sheetId).moveTo(folder);
              status += "SME moved from archieved folder to the main folder" + "\n";
              emailLinkArray.push([emailId, department, sheetLink]);
              DriveApp.getFileById(sheetId).addEditor(emailId); // Needs to change
            }
          }
        }
      })
    }
      // applyCustomFormatting(sheet.getRange(inputSheetIndex+5, 2, 1, 3)).clearContent();
  })

  let lastRow = indexSheet.getLastRow();
  
  emailLinkArray.forEach(r => {
    if(lastRow === 2){
      indexSheet.getRange(lastRow + 1, indexIndices.indexSrNo+1).setValue(1);
      indexSheet.getRange(lastRow + 1, indexIndices.smeEmail+1).setValue(r[0]);
      indexSheet.getRange(lastRow + 1, indexIndices.indexDepartmentIdx+1).setValue(r[1]);
      indexSheet.getRange(lastRow + 1, indexIndices.indexSheetLink+1).setValue(r[2]);
      lastRow++;
    }else{
      // Get the last row Index
      const lastSrNo = indexSheet.getRange(lastRow, indexIndices.indexSrNo+1).getValue();
      indexSheet.getRange(lastRow + 1, indexIndices.indexSrNo+1).setValue(lastSrNo + 1);
      indexSheet.getRange(lastRow + 1, indexIndices.smeEmail+1).setValue(r[0]);
      indexSheet.getRange(lastRow + 1, indexIndices.indexDepartmentIdx+1).setValue(r[1]);
      indexSheet.getRange(lastRow + 1, indexIndices.indexSheetLink+1).setValue(r[2]);
      lastRow++;
    }
  })

  SpreadsheetApp.getUi().alert(status);
}

















// function smeBackend(url, smeEmail){
//   const spreadsheet = SpreadsheetApp.openByUrl(url)
//   const sheetId = spreadsheet.getId();
//   console.log("Sheet Id: ", sheetId);
//   const backendSheet = spreadsheet.getSheetByName('Backend');
//   const qaReviewSheet = spreadsheet.getSheetByName("QA Review");
  
  
//   const masterDataRange = masterDBSME.getRange(1, 1, masterDBSME.getLastRow(), masterDBSME.getLastColumn()).getValues();
//   const masterHeader = masterDataRange[0], masterData = masterDataRange.slice(1);

//   const inputIndices = {
//     srNo : masterHeader.indexOf("Sr. No."),
//     emailId : masterHeader.indexOf("Email ID"),
//     department : masterHeader.indexOf("Department"),
//     sme : masterHeader.indexOf("SME Name"),
//     reviewer : masterHeader.indexOf("QA Reviewer"),
//     grade : masterHeader.indexOf("Grade"),
//     designation : masterHeader.indexOf("Designation"),
//     reportingManager : masterHeader.indexOf("Reporting Manager"),
//     addedDateIdx : masterHeader.indexOf("Added Date"),
//     activeIdx : masterHeader.indexOf("Active?"),
//     removedDateIdx : masterHeader.indexOf("Removed Date"),
//     smeSheetLink : masterHeader.indexOf("SME Sheet Link"),
//   }

//   let smeList = []
//   //Check which departments does not have others
//   let uniqueDepartments = [];

//   // Create a map to keep track of departments with 'Others' entries
//   let departmentsWithOthers = {};

//   masterData.forEach(r => {
//     if (r[inputIndices.pyramidCategory] === 'Others') {
//       departmentsWithOthers[r[inputIndices.department]] = true;
//     }
//   });

//   masterData.forEach(r => {
//     if (!departmentsWithOthers[r[inputIndices.department]]) {
//       if (!uniqueDepartments.includes(r[inputIndices.department])) {
//         uniqueDepartments.push(r[inputIndices.department]);
//       }
//     }
//   });

  
//   let pyramidCategory, department;
//   masterData.forEach(r=>{
//     if(r[inputIndices.emailId] === smeEmail){
//       pyramidCategory = r[inputIndices.pyramidCategory];
//       department = r[inputIndices.department];
//     }
//   })

//   smeList.push(masterData.filter(r=>r[inputIndices.emailId] === smeEmail).map(r=>r[inputIndices.sme])[0]);


//   // ONLY SMEs WHOSE ADDED DATE IS PRESENT OR WHOSE REMOVED DATE IS GREATER THAN ADDED DATE
//   // CHANGE THIS - NEW ORG STRUCTURE

//   masterData.forEach(r =>{
//     if(r=> (r[inputIndices.removedDateIdx]==='' && r[inputIndices.activeIdx]===false) || 
//             (r[inputIndices.activeIdx]===true && r[inputIndices.removedDateIdx] !=='')){
//       if(department === r[inputIndices.department] && department !== 'Others'){
//         if(pyramidCategory === 'SME B'){
//           if(r[inputIndices.pyramidCategory] === 'SME A'){
//             smeList.push(r[inputIndices.sme])
            
//           }
//         }else if(pyramidCategory === 'SME C'){
//           if(r[inputIndices.pyramidCategory] === 'SME A' || r[inputIndices.pyramidCategory] === 'SME B'){
//             smeList.push(r[inputIndices.sme])
//           }
//         }else if(pyramidCategory === 'Others'){
//           if(r[inputIndices.pyramidCategory]==='SME A' || r[inputIndices.pyramidCategory]==='SME B' || r[inputIndices.pyramidCategory]==='SME C'){
//             smeList.push(r[inputIndices.sme])
//           }
//         }
//       }else if (department === 'Others' && pyramidCategory === 'Others') {
//         if (r[inputIndices.pyramidCategory] === 'SME A' || r[inputIndices.pyramidCategory] === 'SME B' || r[inputIndices.pyramidCategory] === 'SME C') {
//           const uniqueSMEs = new Set(); // Use a Set to store unique SME names
//           uniqueDepartments.forEach(uniqueDept => {
//             masterData.forEach(innerR => {
//               if (innerR[inputIndices.department] === uniqueDept && innerR[inputIndices.pyramidCategory] !== 'Others') {
//                 uniqueSMEs.add(innerR[inputIndices.sme]); // Add SME name to the Set
//               }
//             });
//           });

//           // Convert the Set back to an array
//           smeList = Array.from(uniqueSMEs);
//         }
//       }
//     }
//   });


//   clearRowsBelow(backendSheet, 1)
//   let rowIndex = 2;
//   smeList.forEach(name => {
//     backendSheet.getRange(rowIndex, 1).setValue(name);
//     rowIndex ++; 
//   });

//   // Protect backend sheet
//   defineProtectionFunctionalities(backendSheet, hide=true);
//   setVlauesQAReviewSheet(backendSheet, qaReviewSheet);
//   DriveApp.getFileById(sheetId).setOwner("automation@upthink.com")
//   spreadsheet.addEditor("supriya.pawar@upthink.com")
//   // SpreadsheetApp.openById(sheetId).  //("automation@upthink.com")
// }





















