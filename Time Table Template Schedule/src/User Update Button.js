class UpdateNames {
  constructor() {
    // your “current” date context
    this.date         = new Date(2025, 3, 28);
    this.currentYear  = this.date.getFullYear();
    this.DATA_FOLDER_ID = GOOGLE_DRIVE_FOLDER_ID;

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Create Spreadsheet");
    // [ ID, subject, name, isAddition, isDeletion, includeRow ]
    this.data = sheet.getRange(12, 1, 5, 6).getValues();
    this.namesBySubject = this._groupNames();
  }

  // turn a folder or subject name into “underscore_key” form without regex
  _normalizeKey(name) {
    let normalized = name.trim().toLowerCase();
    if (normalized.endsWith(' schedules')) {
      normalized = normalized.slice(0, normalized.length - ' schedules'.length);
    } else if (normalized.endsWith(' schedule')) {
      normalized = normalized.slice(0, normalized.length - ' schedule'.length);
    }
    const parts = normalized.split(' ').filter(function(p) { return p; });
    return parts.join('_');
  }

  // build { subject: { add: [], delete: [] } }
  _groupNames() {
    return this.data.reduce(function(acc, row) {
      const subject    = row[1];
      const name       = row[2];
      const isAddition = row[3];
      const isDeletion = row[4];
      const includeRow = row[5];
      if (!includeRow) return acc;

      if (!acc[subject]) {
        acc[subject] = { add: [], delete: [] };
      }
      if (isAddition) acc[subject].add.push(name);
      if (isDeletion) acc[subject].delete.push(name);
      return acc;
    }, {});
  }

  getNamesBySubject() {
    return this.namesBySubject;
  }

  getAddNames(subject) {
    return this.namesBySubject[subject]?.add || [];
  }

  getDeleteNames(subject) {
    return this.namesBySubject[subject]?.delete || [];
  }

  /**
   * @param {string} subjectKey  normalized folder key e.g. "statistics"
   * @returns {string|null} name of the file in the highest‐year, highest‐month folder
   */
  getLatestFileName(subjectKey) {
    // 1️⃣ find the subject folder
    const root    = DriveApp.getFolderById(this.DATA_FOLDER_ID);
    const folders = root.getFolders();
    let subjectFolder = null;
    while (folders.hasNext()) {
      const f   = folders.next();
      const key = this._normalizeKey(f.getName());
      if (key === subjectKey) {
        subjectFolder = f;
        break;
      }
    }
    if (!subjectFolder) return null;

    // 2️⃣ pick the folder whose last 4 chars parse to the largest year
    const yearFolders       = subjectFolder.getFolders();
    let latestYear          = -Infinity;
    let latestYearFolder    = null;
    while (yearFolders.hasNext()) {
      const yf      = yearFolders.next();
      const name    = yf.getName().trim();
      const yrStr   = name.substring(name.length - 4);
      const yearNum = parseInt(yrStr, 10);
      if (!isNaN(yearNum) && yearNum > latestYear) {
        latestYear       = yearNum;
        latestYearFolder = yf;
      }
    }
    if (!latestYearFolder) return null;

    // 3️⃣ within that, pick the folder whose first 2 chars parse to the largest month
    const monthFolders      = latestYearFolder.getFolders();
    let latestMonth         = -Infinity;
    let latestMonthFolder   = null;
    while (monthFolders.hasNext()) {
      const mf       = monthFolders.next();
      const mName    = mf.getName().trim();
      const mmStr    = mName.substring(0, 2);
      const monthNum = parseInt(mmStr, 10);
      if (!isNaN(monthNum) && monthNum > latestMonth) {
        latestMonth       = monthNum;
        latestMonthFolder = mf;
      }
    }
    if (!latestMonthFolder) return null;

    // 4️⃣ grab all file names, sort, and return the last
    const files = latestMonthFolder.getFiles();
    const names = [];
    while (files.hasNext()) {
      const file = files.next()
      const name = file.getName();
      if (name.toLowerCase().endsWith('.xlsx')) continue; //xlsm
      names.push(file.getId());
    }
    if (names.length === 0) return null;
    names.sort();
    return names[names.length - 1];
  }

  /**
   * @param {string} subjectKey  normalized folder key e.g. "statistics"
   * @returns {string|null} name of the alphabetically last file in
   *                      the folder matching this.date’s year/month
   */
  getCurrentFileName(subjectKey) {
    // 1️⃣ find subject folder
    const root    = DriveApp.getFolderById(this.DATA_FOLDER_ID);
    const folders = root.getFolders();
    let subjectFolder = null;
    while (folders.hasNext()) {
      const f   = folders.next();
      const key = this._normalizeKey(f.getName());
      if (key === subjectKey) {
        subjectFolder = f;
        break;
      }
    }
    if (!subjectFolder) return null;

    // 2️⃣ find the year folder containing this.currentYear
    const yearFolders = subjectFolder.getFolders();
    let yearFolder    = null;
    const yearStr     = this.currentYear.toString();
    while (yearFolders.hasNext()) {
      const yf = yearFolders.next();
      if (yf.getName().indexOf(yearStr) !== -1) {
        yearFolder = yf;
        break;
      }
    }

    if (!yearFolder) return null;

    // 3️⃣ build the exact month folder name: MM<MonthName>-YY
    const mm         = (this.date.getMonth() + 1).toString().padStart(2, '0');
    const monthName  = this.date.toLocaleString('default', { month: 'short' });
    const yy         = yearStr.slice(-2);
    const folderName = mm + monthName + '-' + yy;
    console.log(folderName)
    const months = yearFolder.getFoldersByName(folderName);
    if (!months.hasNext()) return null;
    const monthFolder = months.next();

    // 4️⃣ list, sort, and return the last file name
    const files = monthFolder.getFiles();
    const names = [];
    while (files.hasNext()) {
      const file = files.next()
      const name = file.getName();
      if (name.toLowerCase().endsWith('.xlsx')) continue; //xlsm
      names.push(file.getId());
    }
    if (names.length === 0) return null;
    names.sort();
    return names[names.length - 1];
  }
}
