function runAllUpdates() {
  runTimetableAllUpdates();
  // runQCAllUpdates();
}


/**
 * DRIVER: run both the tutor-names update and the OnlineWk summaries
 */
function runQCAllUpdates() {
  const updater        = new UpdateNamesQC();
  const namesBySubject = updater.getNamesBySubject();
  console.log('Name changes:', namesBySubject);

  Object.keys(namesBySubject).forEach(subject => {
    const key    = updater._normalizeKey(subject);
    const latest = updater.getLatestFileName(key);
    console.log(`Latest file for ${subject}:`, latest);
    const changes = namesBySubject[subject];
    let ss;
    if (latest && (changes.delete.length || changes.add.length)) {
      ss = SpreadsheetApp.openById(latest);
      const deleteTutorNames = new SpreadsheetUpdater(ss, changes);
      deleteTutorNames.updateTutorNamesSheet("both")
      // deleteTutorNames.updateTutorNamesSheet("add");
      new OnlineWkSummaryUpdater(ss);
      const extendedUpdater = new ExtendedUpdater(ss)
      extendedUpdater.addOrDeleteColumn();
      const summaryUpdater = new SummaryUpdater(ss)
      summaryUpdater.addOrDeleteColumn();
    }
    const current = updater.getCurrentFileName(key);
    console.log(`Current file for ${subject}:`, current);
    if (current && changes.add.length) {
      ss = SpreadsheetApp.openById(current);
      const addTutorNames = new SpreadsheetUpdater(ss, changes);
      addTutorNames.updateTutorNamesSheet("add");
      new OnlineWkSummaryUpdater(ss);
      const extendedUpdater = new ExtendedUpdater(ss)
      extendedUpdater.addOrDeleteColumn();
      const summaryUpdater = new SummaryUpdater(ss)
      summaryUpdater.addOrDeleteColumn();
    }
  });
}









/**
 * DRIVER: run both the tutor-names update and the OnlineWk summaries
 */
function runTimetableAllUpdates() {
  const updater        = new UpdateNames();
  const namesBySubject = updater.getNamesBySubject();
  console.log('Name changes:', namesBySubject);

  Object.keys(namesBySubject).forEach(subject => {
    const key    = updater._normalizeKey(subject);
    const latest = updater.getLatestFileName(key);
    console.log(`Latest file for ${subject}:`, latest);
    const changes = namesBySubject[subject];
    let ss;
    if (latest && (changes.delete.length || changes.add.length)) {
      ss = SpreadsheetApp.openById(latest);
      const deleteTutorNames = new SpreadsheetUpdater(ss, changes);
      deleteTutorNames.updateTutorNamesSheet("both")
      // deleteTutorNames.updateTutorNamesSheet("add");
      // new OnlineWkSummaryUpdater(ss);
      // const extendedUpdater = new ExtendedUpdater(ss)
      // extendedUpdater.addOrDeleteColumn();
      const summaryUpdater = new SummaryUpdater(ss)
      summaryUpdater.addOrDeleteColumn();
    }
    // const current = updater.getCurrentFileName(key);
    // console.log(`Current file for ${subject}:`, current);
    // if (current && changes.add.length) {
    //   ss = SpreadsheetApp.openById(current);
    //   const addTutorNames = new SpreadsheetUpdater(ss, changes);
    //   addTutorNames.updateTutorNamesSheet("add");
    //   new OnlineWkSummaryUpdater(ss);
    //   const extendedUpdater = new ExtendedUpdater(ss)
    //   extendedUpdater.addOrDeleteColumn();
    //   const summaryUpdater = new SummaryUpdater(ss)
    //   summaryUpdater.addOrDeleteColumn();
    // }
  });
}