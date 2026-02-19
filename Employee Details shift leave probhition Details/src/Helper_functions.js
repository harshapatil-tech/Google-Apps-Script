/**
 * Compares two arrays of arrays for deep equality.
 *
 * @param {Array<Array>} arr1 - The first array of arrays to compare.
 * @param {Array<Array>} arr2 - The second array of arrays to compare.
 * @return {boolean} - Returns true if both arrays are identical, false otherwise.
 */
function areArraysEqual(arr1, arr2) {
  // Check if both are arrays
  if (!Array.isArray(arr1) || !Array.isArray(arr2)) {
    return false;
  }
  
  // Check if they have the same number of rows
  if (arr1.length !== arr2.length) {
    return false;
  }
  
  // Iterate through each row
  for (let i = 0; i < arr1.length; i++) {
    const row1 = arr1[i];
    const row2 = arr2[i];
    
    // Check if both rows are arrays
    if (!Array.isArray(row1) || !Array.isArray(row2)) {
      return false;
    }
    
    // Check if both rows have the same number of columns
    if (row1.length !== row2.length) {
      return false;
    }
    
    // Compare each cell in the row
    for (let j = 0; j < row1.length; j++) {
      // Handle cases where data types might differ (e.g., number vs. string)
      if (String(row1[j]).trim() !== String(row2[j]).trim()) {
        return false;
      }
    }
  }
  
  // If all checks pass, the arrays are equal
  return true;
}

