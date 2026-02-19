function replicateCurrentMonthTemplate() {
  const replicate = new ReplicaTemplate();
  replicate.loopOver();
}


/**
 * Class to replicate every months's template to next month
 */
class ReplicaTemplate {

  constructor () {
    const date = new Date();
    this.currentMonthNum = (date.getMonth() + 1).toString().padStart(2, '0');
    this.currentMonthName = date.toLocaleString('en-US', { month: 'short' });
    this.currentYear = date.getFullYear();
    // build a lookup of normalized keys
    this.SUBJECT_KEYS = new Set(
      SUBJECTS.map(s => s.trim().toLowerCase())
    );
    this.SUBJECT_KEYS.forEach(i => console.log(i))
    //Logger.log(this.SUBJECT_KEYS)
  }

  /**
   * Function to loop over the schedule folder
   */
  loopOver () {
    // Get the root folder using root folder id
    const rootFolder = DriveApp.getFolderById(GOOGLE_DRIVE_FOLDER_ID);
    // Loop over all the folders inside the root folder
    const subjectFolders = rootFolder.getFolders();
    while(subjectFolders.hasNext()) {
      const subjectFolder = subjectFolders.next();
      const rawSubjectFolderName = subjectFolder.getName().trim().toLowerCase();
      console.log("raw sub is:-",rawSubjectFolderName);
      // normalize the folder name into the same “key” form
      // here we strip off a trailing “ schedules” (if present)
      // then convert spaces → underscores
      const key = rawSubjectFolderName
        .replace(/\s*schedules?$/, "")
        .trim();
      if (this.SUBJECT_KEYS.has(key)) {
        const yearFolders = subjectFolder.getFolders();
        while (yearFolders.hasNext()) {
          const yearFolder = yearFolders.next();
          const rawYearFolderName = yearFolder.getName().trim().toLowerCase().split(" ")[2];
          if (rawYearFolderName != undefined && rawYearFolderName.startsWith("table_") && rawYearFolderName.endsWith(`_${this.currentYear}`)) {
            const monthFolders = yearFolder.getFoldersByName(`${this.currentMonthNum}${this.currentMonthName}-${this.currentYear.toString().slice(-2)}`);
            while(monthFolders.hasNext()) {
              const monthFolder = monthFolders.next();
              const files = monthFolder.getFiles();
              while(files.hasNext()) {
                const file = files.next();
                const fileName = file.getName();
                if (fileName.startsWith("Time table_BF") && !fileName.endsWith(".xlsx")) {
                  // Get the already available copy
                  Helper.modifySS(file.getId(), key);
                  break;
                }else if (fileName.startsWith("Time table_BF") && fileName.endsWith(".xlsx")) {
                  const rawFileName = file.getName().replace(/\.xlsx$/, "");
                  console.log("Raw file name:", rawFileName)
                  const copiedId = Helper.logXlsmCommentsAndPopulateColumns(monthFolder, file.getId(), rawFileName);
                  Helper.modifySS(copiedId, key);
                  break;
                } 
              } 
            }
          }
        }
      }
    }
  }
}






//new code
// function replicateCurrentMonthTemplate() {
//   const replicate = new ReplicaTemplate();
//   replicate.loopOver();
// }


// /**
//  * Class to replicate every months's template to next month
//  */
// class ReplicaTemplate {

//   constructor () {
//     this.date = new Date();
//     this.currentMonthNum = (this.date.getMonth() + 1).toString().padStart(2, '0');
//     this.currentMonthName = this.date.toLocaleString('en-US', { month: 'short' });
//     this.currentYear = this.date.getFullYear();
//     // build a lookup of normalized keys
//     this.SUBJECT_KEYS = new Set(
//       SUBJECTS.map(s => s.trim().toLowerCase())
//     );
//     this.SUBJECT_KEYS.forEach(i => console.log(i))
//     //Logger.log(this.SUBJECT_KEYS)
//   }

//   /**
//    * Function to loop over the schedule folder
//    */
//   loopOver () {
//     // Get the root folder using root folder id
//     const rootFolder = DriveApp.getFolderById(GOOGLE_DRIVE_FOLDER_ID);
//     // Loop over all the folders inside the root folder
//     const subjectFolders = rootFolder.getFolders();
//     while(subjectFolders.hasNext()) {
//       const subjectFolder = subjectFolders.next();
//       const rawSubjectFolderName = subjectFolder.getName().trim().toLowerCase();
//       console.log("raw sub is:-",rawSubjectFolderName);
//       // normalize the folder name into the same “key” form
//       // here we strip off a trailing “ schedules” (if present)
//       // then convert spaces → underscores
//       const key = rawSubjectFolderName
//         .replace(/\s*schedules?$/, "")
//         .trim();
//       if (this.SUBJECT_KEYS.has(key)) {
//         const yearFolders = subjectFolder.getFolders();
//         while (yearFolders.hasNext()) {
//           const yearFolder = yearFolders.next();
//           const rawYearFolderName = yearFolder.getName().trim().toLowerCase().split(" ")[2];
//           if (rawYearFolderName != undefined && rawYearFolderName.startsWith("table_") && rawYearFolderName.endsWith(`_${this.currentYear}`)) {
//             const monthFolders = yearFolder.getFoldersByName(`${this.currentMonthNum}${this.currentMonthName}-${this.currentYear.toString().slice(-2)}`);
//             while(monthFolders.hasNext()) {
//               const monthFolder = monthFolders.next();
//               const files = monthFolder.getFiles();
//               while(files.hasNext()) {
//                 const file = files.next();
//                 const fileName = file.getName();
//                 if (fileName.startsWith("Time table_BF") && !fileName.endsWith(".xlsx")) {
//                   // Get the already available copy
//                   Helper.modifySS(file.getId(), key);
//                   // Fill dates in Online_Wk sheets with correct dates and days
//                   const ss = SpreadsheetApp.openById(file.getId());
//                   const tutorSheet = ss.getSheetByName("Tutor_Names");
//                   OnlineWkProcessor.process(ss, tutorSheet, this.date);
//                   new OnlineWkSummaryUpdater(ss);
//                   break;
//                 }else if (fileName.startsWith("Time table_BF") && fileName.endsWith(".xlsx")) {
//                   const rawFileName = file.getName().replace(/\.xlsx$/, "");
//                   console.log("Raw file name:", rawFileName)
//                   const copiedId = Helper.logXlsmCommentsAndPopulateColumns(monthFolder, file.getId(), rawFileName);
//                   Helper.modifySS(copiedId, key);
//                   // Fill dates in Online_Wk sheets with correct dates and days
//                   const ss = SpreadsheetApp.openById(copiedId);
//                   const tutorSheet = ss.getSheetByName("Tutor_Names");
                  
//                   // OnlineWkProcessor calculates NEXT month, so for Feb XLSX we need to pass Jan date
//                   const monthNum = parseInt(this.currentMonthNum);
//                   const yearNum = this.currentYear;
//                   // Pass previous month so processor calculates current month
//                   let prevMonthNum = monthNum - 1;
//                   let prevYearNum = yearNum;
//                   if (prevMonthNum < 1) {
//                     prevMonthNum = 12;
//                     prevYearNum = yearNum - 1;
//                   }
//                   const xlsxDate = new Date(prevYearNum, prevMonthNum - 1, 1);
//                   console.log("XLSX conversion date (previous month):", xlsxDate);
                  
//                   OnlineWkProcessor.process(ss, tutorSheet, xlsxDate);
//                   new OnlineWkSummaryUpdater(ss);
//                   break;
//                 } 
//               } 
//             }
//           }
//         }
//       }
//     }
//   }
// }




