const LOG_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs")

function main() {
  const allFiles = listFoldersInFolder();
  for(const [key, value] of Object.entries(allFiles)){
    for (const [subject, id] of Object.entries(value.subjects)){
      getAllFilesInFolder(id);
    }
  }
}

function listAllSubfolders(parentFolder, object) {
  const subFolders = parentFolder.getFolders(); 
  while (subFolders.hasNext()) { 
    var folder = subFolders.next();
    if (folder.getName().match(/^\d{4}$/)) {  
      var yearObject = {
        id: folder.getId(),
        subjects: {}
      };
      yearObject.subjects = listAllSubfolders(folder, {});
      object[folder.getName()] = yearObject;  
    } else {
      object[folder.getName()] = folder.getId();
    }
  }
  return object;
}

function listFoldersInFolder() {
  var folderId = "18_aFs7a9u_NxOZ-MUwHA4PadJYM2tVEk"; // ID of your parent folder
  var parentFolder = DriveApp.getFolderById(folderId);
  const object = {};
  return listAllSubfolders(parentFolder, object);
}





function getAllFilesInFolder(folderId) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  
  const logSheetHeaders = LOG_SHEET.getRange(1, 1, 1, LOG_SHEET.getLastColumn()).getValues().flat();
  console.log(logSheetHeaders)
  const srNoIdx = logSheetHeaders.indexOf('Sr. No.');
  const sheetNameIdx = logSheetHeaders.indexOf('Sheet Name');
  const sheetIdIdx = logSheetHeaders.indexOf('Sheet Id');

  let lastRow;
  try{
    lastRow = LOG_SHEET.getLastRow()
  }catch(e){
    lastRow = 1
  }
  let srNoColumn = LOG_SHEET.getRange(2, sheetIdIdx + 1, lastRow, 1).getValues();
  let maxSrNo = Math.max(...srNoColumn.flat());

  // Check if the log sheet is empty
  if (maxSrNo === 0) {
    maxSrNo = 0;
  }

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    var fileId = file.getId();
    
    var existingRow = getExistingRow(LOG_SHEET, fileName, sheetNameIdx);
    if (existingRow !== -1) {
      console.log(`Skipping file "${fileName}" (ID: ${fileId}) as it is already present in the log sheet.`);
      continue;
    }
  
    var fileExtension = fileName.substring(fileName.lastIndexOf('.') + 1);
    var newFileId;
    
    // Check if the file has the desired extension (e.g., .xls)
    if (fileExtension == 'xls' || fileExtension == 'xlsm' || fileExtension == 'xlsx') {
      var fileBlob = file.getBlob();
      var [filePart1, filePart2] = fileBlob.getName().split(".");
      var resource = {
        title: filePart1,
        mimeType: 'application/vnd.google-apps.spreadsheet'
      };
  
      var newFile = Drive.Files.insert(resource, fileBlob);
      newFileId = newFile.id;
      getData(newFileId);
  
      // Optional: Trash the new Google Sheet instead of the original file
      var newSheetFile = DriveApp.getFileById(newFileId);
      newSheetFile.setTrashed(true);
    } else {
        getData(file.getId())
      // newFileId = fileId;
      // getData(fileId);
    }
  
    // Add the file details to the log sheet
    maxSrNo++;
    var newRow = [maxSrNo, fileName, fileId];
    LOG_SHEET.appendRow(newRow);
  }
}



function getExistingRow(logSheet, fileName, sheetNameIdx) {
  const dataRange = logSheet.getDataRange();
  const values = dataRange.getValues();
  
  for (let i = 1; i < values.length; i++) {
    const row = values[i];    
    const fileNameCell = row[sheetNameIdx];
    
    if (fileNameCell === fileName) {
      console.log(fileNameCell, fileName)
      return i + 1; // Return the row number (1-based index) + 1 to match the sheet row index
    }
  }
  
  return -1; // Return -1 if the file is not found in the log sheet
}

