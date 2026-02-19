class DriveUtil {

  static getChildFolder(parentFolder, folderKey) {
    const folders = parentFolder.getFolders();
    while(folders.hasNext()) {
      const folder = folders.next();
      const folderName = folder.getName().trim().toLowerCase();
      if (folderName == folderKey || folderName.includes(folderKey)) {
        return folder
      }
    }
    return null;
  }

  static getOrCreateFolder(parentFolder, folderName) {
    const iterator = parentFolder.getFoldersByName(folderName);
    return iterator.hasNext() ? iterator.next() : parentFolder.createFolder(folderName);
  }


  static getOrCreateFile (targetFolder, sourceFile, newFileName) {
    const iterator = targetFolder.getFilesByName(newFileName);
    if (iterator.hasNext()) {
      return { 
        file: iterator.next(), 
        action: 'exists'        // indicate we found an existing file
      };
    } else {
      const copy = sourceFile.makeCopy(newFileName, targetFolder);
      return { 
        file: copy, 
        action: 'copied'        // indicate we just made a new copy
      };
    }
  }


  /**
   * Iterates all files in a Drive folder, calling your callback
   * for every file whose name contains `namePattern` (if given)
   * and whose extension is NOT in `excludeExts`.
   *
   * @param {Folder}   folder           The Drive Folder to scan.
   * @param {string}   stringPattern    Substring to match on file names.
   * @param {function(File)} callback   Invoked for each matching file.
   * @param {string[]} excludeExts      Extensions to skip (including leading dot). Default: [".xlsm"]
  */
  static copyFilesForEach(folder, stringPattern, callback, excludeExts = [".xlsx"]) { //xlsm
    const files      = folder.getFiles();
    const pattern    = stringPattern.trim().toLowerCase();
    const excludeSet = new Set(excludeExts.map(ext => ext.toLowerCase()));

    while (files.hasNext()) {
      const file      = files.next();
      const rawName   = file.getName().trim();
      const nameLC    = rawName.toLowerCase();
      // find extension (including the dot), or empty string if none
      const dotIndex = nameLC.lastIndexOf(".");
      const ext      = dotIndex > -1 ? nameLC.slice(dotIndex) : "";

      // only skip if there *is* an extension AND it's in the exclude list
      if (dotIndex > -1 && excludeSet.has(ext)) {
        continue;
      }
      // if a pattern was provided, skip names that don't include it
      if (!nameLC.includes(pattern)) {
        continue;
      }

      // yay—this file matches!
      callback(file);
    }
  }


  /**
   * Returns all files in `folder` whose names include `stringPattern`
   * (case-insensitive) and whose extensions (if any) are NOT in `excludeExts`.
   *
   * @param {Folder}    folder         The Drive folder to scan.
   * @param {string}    stringPattern  Substring to match (optional).
   * @param {string[]}  excludeExts    Extensions *with* leading “.” to skip.
   *                                   Default is [".xlsm"].
   * @return {File[]}                  Array of matching File objects.
   */
  static getFiles(folder, stringPattern = "", excludeExts = [".xlsx"]) {    //xlsm
    const files      = folder.getFiles();
    const pattern    = stringPattern.trim().toLowerCase();
    const excludeSet = new Set(excludeExts.map(e => e.toLowerCase()));
    const results    = [];

    while (files.hasNext()) {
      const file      = files.next();
      const rawName   = file.getName().trim();
      const nameLC    = rawName.toLowerCase();

      // find extension (including the dot), or empty string if none
      const dotIndex = nameLC.lastIndexOf(".");
      const ext      = dotIndex > -1 ? nameLC.slice(dotIndex) : "";

      // skip only if there *is* an extension AND it's excluded
      if (dotIndex > -1 && excludeSet.has(ext)) {
        continue;
      }

      // if a pattern was provided, skip names that don't include it
      if (pattern && !nameLC.includes(pattern)) {
        continue;
      }

      return file;
    }
  }

}
