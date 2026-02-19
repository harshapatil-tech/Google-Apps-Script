function updateFinalClosureDates(sheet, dbHeaders, row, i, formattedDate, today){
  
  const closureEmailDateIndex = dbHeaders["Closure Email Date"];
  const finalClosureDateIndex = dbHeaders["Final Closure Date"];

  // Extract values from "Closure Email Date" and "Final Closure Date" columns for the current row
  const closureEmailDate = row[closureEmailDateIndex];
  const finalClosureDate = row[finalClosureDateIndex];

  // Proceed only if:
  // - "Closure Email Date" is NOT blank
  // - "Final Closure Date" is blank
  if (closureEmailDate && !finalClosureDate) {
    const emailDate = new Date(closureEmailDate);
    

    // Calculate the difference in days between today and the closure email date:
    // (today - emailDate) → difference in milliseconds
    // (1000 * 60 * 60 * 24) → milliseconds in a day
    // Math.floor → rounds down to the nearest whole number
    const diffDays = Math.floor((today - emailDate) / (1000 * 60 * 60 * 24));
    
    // If the closure email was sent more than 10 days ago
    if (diffDays > 10) {
      // console.log(i+2)
      // Update the "Final Closure Date" column with today's formatted date
      sheet.getRange(i + 2, finalClosureDateIndex + 1).setValue(formattedDate);
      // updatesMade++;
    }
  }
}