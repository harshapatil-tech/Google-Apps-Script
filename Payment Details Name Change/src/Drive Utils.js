class DriveUtil {
  static findFolder(root, key) {
    const folders = root.getFolders();
    while (folders.hasNext()) {
      const f = folders.next();
      const folderName = f.getName();
      if (folderName === key || folderName.includes(key)) return f;
    }
    return null;
  }

  static findFolderById(folderId) {
    try {
      return DriveApp.getFolderById(folderId)
    } catch {
      throw new Error(`Folder with ID "${folderId}" not found or inaccessible.`);
    }
  }

  static getOrCreateFolder (root, nameKey) {
    const found = DriveUtil.findFolder(root, nameKey);
    return found || root.createFolder(nameKey);
  }

  /**
   * Copies all files from sourceFolder into destinationFolder,
   * replacing occurrences of prevYearRange in each filename
   * with currYearRange.
   *
   * @param {Folder} sourceFolder
   * @param {Folder} destinationFolder
   * @param {string} prevYearRange   e.g. "2024-25"
   * @param {string} currYearRange   e.g. "2025-26"
   * @return {File[]}  the newly created files
   */
  static copyFilesFromFolderToFolder(sourceFolder, destFolder,
                                     prevRange, currRange) {
    const fileIds = [];
    const files = sourceFolder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const originalName = file.getName();
      const newName = originalName.replace(prevRange, currRange);

      // see if it already exists in destFolder
      const existing = destFolder.getFilesByName(newName);
      let targetFile;
      if (existing.hasNext()) {
        targetFile = existing.next();
      } else {
        // make a fresh copy
        targetFile = file.makeCopy(newName, destFolder);
      }

      fileIds.push(targetFile.getId());
    }
    return fileIds;
  }

  /**
   * Searches for files inside a folder whose names include the given key.
   * @param {Folder} folder      The Drive Folder object to search within.
   * @param {string} key         Substring to match in file names.
   * @return {File[]}            Array of matching File objects.
   */
    /**
   * Searches for files inside a folder whose names include the given key,
   * excluding files with the specified exception extension.
   * @param {Folder} folder      The Drive Folder object to search within.
   * @param {string} key         Substring to match in file names.
   * @param {string} exceptionExt File extension to exclude (e.g. ".xlsm").
   * @return {File[]}            Array of matching File objects.
   */
  static searchFiles(folder, key, exceptionExt = ".xlsm") {
    const matches = [];
    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const name = file.getName();
      if (name.includes(key) && !name.toLowerCase().endsWith(exceptionExt.toLowerCase())) {
        matches.push(file);
      }
    }
    return matches;
  }


// 	3.1 Inform Silicon Rental
// May 12, 2025, 10:15:59 AM	Error	TypeError: Assignment to constant variable.
//     at emailTemplate2(Email template 2 - Ticket Raised Silicon Rental Solutions:53:19)
//     at updateMasterDB(CRUD update button:204:7)
//     at update(CRUD update button:170:12)
//     at updateButton(CRUD update button:5:12)


  // /**
  //  * Copies all files from sourceFolder into destinationFolder.
  //  *
  //  * @param {Folder} sourceFolder     The Drive folder whose files you want to copy.
  //  * @param {Folder} destinationFolder The Drive folder to receive the copies.
  //  * @return {File[]}                 An array of the newly created File objects.
  //  */
  // static copyFilesFromFolderToFolder(sourceFolder, destinationFolder) {
  //   const copiedFiles = [];
  //   const files = sourceFolder.getFiles();
  //   while (files.hasNext()) {
  //     const file = files.next();
  //     // makeCopy(name, destination) returns the new File
  //     const copy = file.makeCopy(file.getName(), destinationFolder);
  //     copiedFiles.push(copy);
  //   }
  //   return copiedFiles;
  // }
}