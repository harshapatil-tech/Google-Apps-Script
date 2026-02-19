// function main() {
//   const folderId = "18_aFs7a9u_NxOZ-MUwHA4PadJYM2tVEk"; // Replace with your folder ID
//   const logSheetName = "Logs";
//   const processor = new DriveFileProcessor(folderId, logSheetName);
//   processor.processAllFiles();
// }

class DriveFileProcessor{
  constructor(folderId, logSheetName){
    this.folderId = folderId;
    this.logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(logSheetName); 

    this.logSheetHeaders = CentralLibrary.get_Data_Indices_From_Sheet(this.logSheet); 
    // console.log(this.logSheetHeaders);
    
  }

  listAllSubfolders(parentFolder, object){
    const subFolders = parentFolder.getFolders(); //iterate for all subfolders
    while(subFolders.hasNext()){  //iterates through each subfolder
      const folder = subFolders.next();
      if(folder.getName().match(/^\d{4}$/)){  //check foldername matches a 4 digit year pattern
        const yearObject = {
          id : folder.getId(),
          subjects : {}
        };
        yearObject.subjects = this.listAllSubfolders(folder, {});
        object[folder.getName()] = yearObject;  //adding year folder and subfolder to the object
      }else{
        object[folder.getName()] = folder.getId();
      }
    }
    return object;
  }

  listFoldersInFolder(){
    const parentFolder = DriveApp.getFolderById(this.folderId);
    const object = {};
    return this.listAllSubfolders(parentFolder, object);
  }


  getAllFilesInFolder(folderId){
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles(); //get all files in folder
    const [logSheetHeaders, logData] = CentralLibrary.get_Data_Indices_From_Sheet(this.logSheet)
    console.log(logSheetHeaders);
    let lastRow = this.logSheet.getLastRow();  //fetch last row in the logsheet
    const srNoColumn = this.logSheet.getRange(2, logSheetHeaders["Sr. No."] + 1, lastRow, 1).getValues();  //get all values from srno column
    // console.log(srNoColumn);
    let maxSrNo = Math.max(...srNoColumn.flat());   //find highest sr no in the sheet

    if(maxSrNo === 0){
      maxSrNo = 0;
    }

    while(files.hasNext()){
      const file = files.next();  //fetch each file in the folder
      const fileName = file.getName();  //get file name
      const fileId = file.getId();  //get file id
      console.log(fileName)
      const existingRow = this.getExistingRow(fileName)  //check if file is already logged
      if (existingRow !== -1) {
      console.log(`Skipping file "${fileName}" (ID: ${fileId}) as it is already present in the log sheet.`);
      continue; 
      }
      const fileExtension = fileName.substring(fileName.lastIndexOf('.') + 1); // extracts file extension.
      let newFileId;

      if (fileExtension == 'xls' || fileExtension == 'xlsm'){
        const fileBlob = file.getBlob();
        const [filePart1] = fileBlob.getName().split('.'); // extracts file name without extension.
        const resource = {
        title: filePart1,
        mimeType: 'application/vnd.google-apps.spreadsheet'
        };
        var newFile = Drive.Files.insert(resource, fileBlob);  //convert excel to google sheet
        newFileId = newFile.id;
        // this.getData(newFileId);  //processes the new file

        // Optional: Trash the new Google Sheet instead of the original file
        const newSheetFile = DriveApp.getFileById(newFileId); 
        newSheetFile.setTrashed(true);  
      }else{
        // this.getData(fileId)
      }

      maxSrNo++;  
      const newRow = [maxSrNo, fileName, fileId]; // prepares a new log row.
      this.logSheet.appendRow(newRow); // append the new row to the log sheet.

    }
  }

  getExistingRow(fileName){
    const dataRange = this.logSheet.getDataRange();  // get all data in the logsheet
    const values = dataRange.getValues();  
    for(let i=1; i<values.length; i++){
      const row = values[i];
      const fileNameCell = row[this.sheetNameIdx]

      if(fileNameCell === fileName){
        return + 1;  //Return the row number (1-based index) + 1 to match the sheet row index
      }
    }
    return -1;  //Return -1 if the file is not found in the log sheet
  }

  processAllFiles() {
    const allFiles = this.listFoldersInFolder();
    for (const [key, value] of Object.entries(allFiles)) {
      for (const [subject, id] of Object.entries(value.subjects)) {
        this.getAllFilesInFolder(id);
      }
    }
  }

  // getData(fileId) {
  // console.log(`Processing data for file ID: ${fileId}`); 
  // }

}