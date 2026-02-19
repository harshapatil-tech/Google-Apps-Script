/**
 * Calculates the start and end dates for the previous two full weeks (Monday to Sunday).
 * @return {Object} - An object containing startDate and endDate (lastSunday).
 */
function calculatePreviousTwoWeeksDateRange() {
  var today = new Date();
  var daysSinceSunday = today.getDay(); // 0 (Sunday) to 6 (Saturday)
  var lastSunday = new Date(today.getFullYear(), today.getMonth(), today.getDate() - daysSinceSunday);
  var startDate = new Date(lastSunday.getFullYear(), lastSunday.getMonth(), lastSunday.getDate() - 13); // 13 days before last Sunday
  return {
    startDate: startDate,
    endDate: lastSunday
  };
}

/**
 * Parses a cell value into a Date object.
 * @param {any} cellValue - The cell value to parse as a date.
 * @return {Date|null} - The parsed Date object or null if invalid.
 */
function parseDate(cellValue) {
  if (cellValue instanceof Date) {
    return cellValue;
  } else if (typeof cellValue === 'string' && cellValue.trim() !== '') {
    var dateValue = new Date(cellValue);
    if (!isNaN(dateValue)) {
      return dateValue;
    }
  }
  return null;
}

/**
 * Checks if a date is within a specified date range.
 * @param {Date} date - The date to check.
 * @param {Date} startDate - The start date of the range.
 * @param {Date} endDate - The end date of the range.
 * @return {Boolean} - True if date is within the range, false otherwise.
 */
function isDateInRange(date, startDate, endDate) {
  if (!(date instanceof Date)) return false;
  // Normalize the date by setting time to midnight
  var normalizedDate = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  
  return normalizedDate >= startDate && normalizedDate <= endDate;
}

/**
 * Clears all data from a sheet starting from row 2, preserving headers.
 * @param {Sheet} sheet - The sheet to clear data from.
 */
function clearSheetData(sheet) {
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  
  if (lastRow < 2) {
    // No data to clear
    return;
  }
  
  // Define the range to clear: from row 2 to the last row, across all columns
  var rangeToClear = sheet.getRange(2, 1, lastRow - 1, lastColumn);
  
  // Clear the contents of the specified range
  rangeToClear.clearContent();
}

/**
 * Parses a time value (Date object or string) into a Date object representing the time.
 * @param {Date|String} timeValue - The time value.
 * @return {Date|null} - The Date object or null if invalid.
 */
function parseTime(timeValue) {
  if (timeValue instanceof Date && !isNaN(timeValue)) {
    // It's a Date object, extract time
    return new Date(0, 0, 0, timeValue.getHours(), timeValue.getMinutes(), timeValue.getSeconds());
  } else if (typeof timeValue === 'string') {
    // Handle AM/PM format or HH:MM:SS
    var date = new Date('1/1/2000 ' + timeValue);
    if (!isNaN(date)) {
      return new Date(0, 0, 0, date.getHours(), date.getMinutes(), date.getSeconds());
    }
  }
  return null;
}

/**
 * Formats a Date object representing a time into "HH:MM:SS" format.
 * @param {Date} time - The Date object.
 * @return {String} - The formatted time string.
 */
function formatTime(time) {
  if (!(time instanceof Date)) {
    return "";
  }
  var hours = padZero(time.getHours());
  var minutes = padZero(time.getMinutes());
  var seconds = padZero(time.getSeconds());
  return hours + ":" + minutes + ":" + seconds;
}

/**
 * Formats a time difference in milliseconds into HH:MM:SS format.
 * @param {Number} millis - The time difference in milliseconds.
 * @return {String} - The formatted time difference.
 */
function formatTimeDifference(millis) {
  var totalSeconds = Math.floor(Math.abs(millis) / 1000);
  var hours = Math.floor(totalSeconds / 3600);
  var minutes = Math.floor((totalSeconds % 3600) / 60);
  var seconds = totalSeconds % 60;
  var sign = millis < 0 ? "-" : "";
  return sign + padZero(hours) + ":" + padZero(minutes) + ":" + padZero(seconds);
}

/**
 * Pads a number with leading zero if less than 10.
 * @param {Number} num - The number to pad.
 * @return {String} - The padded number as a string.
 */
function padZero(num) {
  return (num < 10 ? "0" : "") + num;
}

/**
 * Formats a date object into YYYY-MM-DD format.
 * @param {Date} date - The date object.
 * @return {String} - The formatted date string.
 */
function formatDate(date) {
  if (!(date instanceof Date)) {
    return "";
  }
  var year = date.getFullYear();
  var month = padZero(date.getMonth() + 1);
  var day = padZero(date.getDate());
  return year + "-" + month + "-" + day;
}

/**
 * Checks if two dates are the same (ignoring time).
 * @param {Date} date1 - The first date.
 * @param {Date} date2 - The second date.
 * @return {Boolean} - True if dates are the same, false otherwise.
 */
function isSameDate(date1, date2) {
  if (!(date1 instanceof Date) || !(date2 instanceof Date)) {
    return false;
  }
  return date1.getFullYear() === date2.getFullYear() &&
         date1.getMonth() === date2.getMonth() &&
         date1.getDate() === date2.getDate();
}


