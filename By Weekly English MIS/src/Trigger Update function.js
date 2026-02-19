/**
 * Sets up a biweekly trigger to run the updateBFTimesheetData function every Tuesday at 7:00 PM.
 */
// function createBiweeklyTrigger() {
//   // First, delete any existing triggers for updateBFTimesheetData to avoid duplicates
//   var allTriggers = ScriptApp.getProjectTriggers();
//   for (var i = 0; i < allTriggers.length; i++) {
//     if (allTriggers[i].getHandlerFunction() == 'updateBFTimesheetData') {
//       ScriptApp.deleteTrigger(allTriggers[i]);
//     }
//   }

//   // Create a new trigger
//   var startDate = new Date('April 08, 2025 22:00:00'); // Set your start date and time
//   ScriptApp.newTrigger('updateBFTimesheetData')
//     .timeBased()
//     .onWeekDay(ScriptApp.WeekDay.TUESDAY)
//     .atHour(22)
//     .nearMinute(0)
//     .everyWeeks(2)
//     .create();
// }



function createWeeklyTrigger() {
  // Create a trigger that fires every Tuesday at 22:00.
  ScriptApp.newTrigger('runTaskEveryFourteenDays')
           .timeBased()
           .onWeekDay(ScriptApp.WeekDay.TUESDAY)
           .atHour(22)
           .create();
}






function runTaskEveryFourteenDays() {
  var props = PropertiesService.getScriptProperties();
  var lastRun = props.getProperty('LAST_RUN');
  var now = new Date();
  
  if (lastRun) {
    var lastDate = new Date(lastRun);
    var twoWeeksInMs = 12 * 24 * 60 * 60 * 1000; // 12 days in milliseconds
    var timeDiff = now.getTime() - lastDate.getTime();
    
    if (timeDiff < twoWeeksInMs) {
      // Not enough time has passed; exit the function.
      Logger.log("Biweekly condition not met. Exiting.");
      return;
    }
  }
  updateBFTimesheetData();
  // Place your scheduled code here.
  Logger.log("Biweekly task running at " + now);
  
  // Update the last execution time.
  props.setProperty('LAST_RUN', now.toISOString());
}












