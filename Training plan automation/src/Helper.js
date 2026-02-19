function sendEmails() {
  CURRENT_SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
  TRAINING_TRACKER_SHEET = Training.getSheetById(this.CURRENT_SPREADSHEET, 0)

  const [headers_TrainingTracker, data_TrainingTracker] = CentralLibrary.get_Data_Indices_From_Sheet(this.TRAINING_TRACKER_SHEET);

  let possibleData = data_TrainingTracker.filter(row => row[headers_TrainingTracker] !== "Sent")
  

}

function protectCell(row, column) {
  // Define the spreadsheet and the range to protect
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getRange(); // Change this to the cell you want to protect

  // Protect the range
  var protection = range.protect().setDescription('Protected cell');

  // Ensure the current user is an editor before removing others. Otherwise, if the user does not have edit access to the range, this script will throw an error.
  var me = Session.getEffectiveUser();
  protection.addEditor(me);
  protection.removeEditors(protection.getEditors());
  
  // Optional: Set the warning only option to allow only owners to modify the protected range
  protection.setWarningOnly(false);

  Logger.log('Cell A1 has been protected and is now uneditable by others.');
}


function makeOwnKeyValuePairs(sheet, keyColumn, ...valueColumns) {
  let [headers, data] = CentralLibrary.get_Data_Indices_From_Sheet(sheet);
  const result = {};

  // Normalize keyColumn and valueColumns
  const normalizedKeyColumn = normalizeColumnName(keyColumn);
  const normalizedValueColumns = valueColumns.map(normalizeColumnName);

  // Create a mapping of normalized column names to header indices
  const headerMap = Object.fromEntries(Object.entries(headers).map(([colName, index]) => [normalizeColumnName(colName), index]));

  // Initialize result structure
  data.forEach(row => {
    const key = row[headerMap[normalizedKeyColumn]] || ''; // Handle empty or missing keys

    // Skip rows with empty key
    if (!key.trim()) {
      return;
    }

    // Initialize the key in result if not already present
    if (!result[key]) {
      result[key] = {};
    }

    // Process each value column
    normalizedValueColumns.forEach(col => {
      const value = row[headerMap[col]] || '';
      if (!result[key][col]) {
        result[key][col] = '';
      }
      result[key][col] += value ? (result[key][col] ? ',' : '') + value : '';
    });
  });

  return result;
}

// Function to normalize column names
function normalizeColumnName(columnName) {
  return columnName
    .replace(/[^a-zA-Z0-9]/g, ' ')   // Replace non-alphanumeric characters with space
    .split(' ')                      // Split by spaces
    .map((word, index) => 
      index === 0 ? word.toLowerCase() : word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()
    )                               // Convert to camelCase
    .join('');
}
