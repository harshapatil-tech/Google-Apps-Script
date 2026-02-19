const ROOT_FOLDER_ID = "1FqWrMYMMs8qRpz2x6Imq923-p2zpNRBX";

const DEPARTMENTS = ["Biology", "Business", "Chemistry", "Computer Science", "English", "Mathematics", "Physics", "Statistics"];

const SCHEDULES_ROOT_ID = "1GsRppc-WVzGYJJ9BSLazbCP4OS42R5ud";



/**
 * Returns a string like "2024-25" for the academic/budget year spanning
 * the previous and current calendar years.
 *
 * @param {Date=} optDate  Optional date reference (defaults to today).
 * @return {string}        Formatted "YYYY-YY" string.
 */
function getPrevFinYear(optDate) {
  // 1. Use provided date or default to now
  var now = optDate ? new Date(optDate) : new Date();
  
  // 2. Compute years
  var prevFull = now.getFullYear() - 1;  
  var currShort = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yy');
  
  // 3. Combine into "YYYY-YY"
  return prevFull + '-' + currShort;
}



/**
 * Returns a string like "2025-26" for the academic/budget year spanning
 * the current and next calendar years.
 *
 * @param {Date=} optDate  Optional date reference (defaults to today).
 * @return {string}        Formatted "YYYY-YY" string.
 */
function getCurrFinYear(optDate) {
  // 1. Use provided date or default to now
  var now = optDate ? new Date(optDate) : new Date();
  
  // 2. Compute years
  var currFull = now.getFullYear();              // e.g. 2025
  var nextFull = currFull + 1;                   // e.g. 2026
  // take the last two digits of nextFull:
  var nextShort = nextFull.toString().slice(-2); // "26"
  
  // 3. Combine into "YYYY-YY"
  return `Demo-${currFull}-${nextShort}`;  // To be deleted in production
  return currFull + '-' + nextShort;
}