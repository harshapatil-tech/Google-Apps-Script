function splitIntoPeriods(startDate, endDate, periodDays = 14) {
  const periods = [];
  let currentStart = new Date(startDate); // clone to avoid mutation
  let remainingDaysInPeriod = periodDays;

  while (currentStart <= endDate) {
    const endOfMonth = new Date(currentStart.getFullYear(), currentStart.getMonth() + 1, 0);
    let currentEnd;

    // Days left in this month
    const daysLeftInMonth = Math.ceil((endOfMonth - currentStart) / (1000 * 60 * 60 * 24)) + 1;

    if (daysLeftInMonth >= remainingDaysInPeriod) {
      // Period fits within current month
      currentEnd = new Date(currentStart);
      currentEnd.setDate(currentStart.getDate() + remainingDaysInPeriod - 1);
      periods.push({ from: new Date(currentStart), to: new Date(currentEnd) });

      // Prepare for next period
      currentStart = new Date(currentEnd);
      currentStart.setDate(currentStart.getDate() + 1);
      remainingDaysInPeriod = periodDays;
    } else {
      // Period spans month boundary
      currentEnd = endOfMonth;
      periods.push({ from: new Date(currentStart), to: new Date(currentEnd) });

      // Remaining days carry to next month
      remainingDaysInPeriod -= daysLeftInMonth;

      currentStart = new Date(currentEnd);
      currentStart.setDate(currentStart.getDate() + 1);
    }
  }

  return periods;
}


function brainfuseAccountWise(spreadSheet) {
  const ss = spreadSheet.getSheetByName("Summary");
  if (!ss) throw new Error("Summary sheet not found");

  const data = ss.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const accounts = {}; // group by subject + account

  // Step 1: Group all rows by subject (department) and account
  rows.forEach(row => {
    let department = row[0]; // Subject
    const account = row[1];
    const type = row[2].trim().toLowerCase();
    const activityType = row[3];
    const startDate = row[4];
    const hours = row[6] || 0;

    if (department.trim().toLowerCase() === "intro accounting") department = "Accounting";
    if (department.trim().toLowerCase() === "mathematics") department = "Calculus";

    if (!accounts[department]) accounts[department] = {};
    if (!accounts[department][account]) accounts[department][account] = [];

    accounts[department][account].push({
      date: new Date(startDate),
      occupancy: activityType.includes("Tutored") ? hours : 0,
      nonOccupancy: activityType.includes("Waited") ? hours : 0,
      accountType: (type === "single" ? "Single" : "Multiple"),
      subject: department
    });
  });

  // Step 2: Create global periods based on all dates in sheet
  const allDates = rows.map(r => new Date(r[4]));
  const minDate = new Date(Math.min(...allDates));
  const maxDate = new Date(Math.max(...allDates));
  const globalPeriods = splitIntoPeriods(minDate, maxDate, 14); // 14-day period logic

  const resultArray = [];

  // Step 3: Iterate subjects and accounts
  for (const subject in accounts) {
    for (const account in accounts[subject]) {
      const accountData = accounts[subject][account];

      globalPeriods.forEach(period => {
        const periodOccupancy = accountData
          .filter(d => d.date >= period.from && d.date <= period.to)
          .reduce((sum, d) => sum + d.occupancy, 0);

        const periodNonOccupancy = accountData
          .filter(d => d.date >= period.from && d.date <= period.to)
          .reduce((sum, d) => sum + d.nonOccupancy, 0);

        // Include even if zero hours (optional: remove if you want only non-zero)
        const monthName = Utilities.formatDate(period.from, Session.getScriptTimeZone(), "MMMM");
        const year = Utilities.formatDate(period.from, Session.getScriptTimeZone(), "yyyy");

        resultArray.push({
          [monthName]: {
            "Occupancy": {
              hours: periodOccupancy,
              subject,
              account,
              firstDay: Utilities.formatDate(period.from, Session.getScriptTimeZone(), "dd-MMM-yyyy"),
              lastDay: Utilities.formatDate(period.to, Session.getScriptTimeZone(), "dd-MMM-yyyy"),
              month: monthName,
              finYear: getFinancialYear(monthName, year),
              accountType: accountData[0].accountType
            },
            "Non-Occupancy": {
              hours: periodNonOccupancy,
              subject,
              account,
              firstDay: Utilities.formatDate(period.from, Session.getScriptTimeZone(), "dd-MMM-yyyy"),
              lastDay: Utilities.formatDate(period.to, Session.getScriptTimeZone(), "dd-MMM-yyyy"),
              month: monthName,
              finYear: getFinancialYear(monthName, year),
              accountType: accountData[0].accountType
            }
          }
        });
      });
    }
  }
  return resultArray;
}






// function splitIntoPeriods(startDate, endDate, periodDays = 14) {
//   const periods = [];
//   let currentStart = new Date(startDate); // clone to avoid mutation
//   let remainingDaysInPeriod = periodDays;

//   while (currentStart <= endDate) {
//     const endOfMonth = new Date(currentStart.getFullYear(), currentStart.getMonth() + 1, 0);
//     let currentEnd;

//     // Days left in this month
//     const daysLeftInMonth = Math.ceil((endOfMonth - currentStart) / (1000 * 60 * 60 * 24)) + 1;

//     if (daysLeftInMonth >= remainingDaysInPeriod) {
//       // Period fits within current month
//       currentEnd = new Date(currentStart);
//       currentEnd.setDate(currentStart.getDate() + remainingDaysInPeriod - 1);
//       periods.push({ from: new Date(currentStart), to: new Date(currentEnd) });

//       // Prepare for next period
//       currentStart = new Date(currentEnd);
//       currentStart.setDate(currentStart.getDate() + 1);
//       remainingDaysInPeriod = periodDays;
//     } else {
//       // Period spans month boundary
//       currentEnd = endOfMonth;
//       periods.push({ from: new Date(currentStart), to: new Date(currentEnd) });

//       // Remaining days carry to next month
//       remainingDaysInPeriod -= daysLeftInMonth;

//       currentStart = new Date(currentEnd);
//       currentStart.setDate(currentStart.getDate() + 1);
//     }
//   }

//   return periods;
// }


// function brainfuseAccountWise(spreadSheet) {
//   const ss = spreadSheet.getSheetByName("Summary");
//   if (!ss) throw new Error("Summary sheet not found");

//   const data = ss.getDataRange().getValues();
//   const headers = data[0];
//   const rows = data.slice(1);

//   const accounts = {}; // group by subject + account

//   // Step 1: Group all rows by subject (department) and account
//   rows.forEach(row => {
//     let department = row[0]; // Subject
//     const account = row[1];
//     const type = row[2].trim().toLowerCase();
//     const activityType = row[3];
//     const startDate = row[4];
//     const hours = row[6] || 0;

//     if (department.trim().toLowerCase() === "intro accounting") department = "Accounting";
//     if (department.trim().toLowerCase() === "mathematics") department = "Calculus";

//     if (!accounts[department]) accounts[department] = {};
//     if (!accounts[department][account]) accounts[department][account] = [];

//     accounts[department][account].push({
//       date: new Date(startDate),
//       occupancy: activityType.includes("Tutored") ? hours : 0,
//       nonOccupancy: activityType.includes("Waited") ? hours : 0,
//       accountType: (type === "single" ? "Single" : "Multiple"),
//       subject: department
//     });
//   });

//   // Step 2: Create global periods based on all dates in sheet
//   const allDates = rows.map(r => new Date(r[4]));
//   const minDate = new Date(Math.min(...allDates));
//   const maxDate = new Date(Math.max(...allDates));
//   const globalPeriods = splitIntoPeriods(minDate, maxDate, 14); // 14-day period logic

//   const resultArray = [];

//   // Step 3: Iterate subjects and accounts
//   for (const subject in accounts) {
//     for (const account in accounts[subject]) {
//       const accountData = accounts[subject][account];

//       globalPeriods.forEach(period => {
//         const periodOccupancy = accountData
//           .filter(d => d.date >= period.from && d.date <= period.to)
//           .reduce((sum, d) => sum + d.occupancy, 0);

//         const periodNonOccupancy = accountData
//           .filter(d => d.date >= period.from && d.date <= period.to)
//           .reduce((sum, d) => sum + d.nonOccupancy, 0);

//         // Include even if zero hours (optional: remove if you want only non-zero)
//         const monthName = Utilities.formatDate(period.from, Session.getScriptTimeZone(), "MMMM");
//         const year = Utilities.formatDate(period.from, Session.getScriptTimeZone(), "yyyy");

//         resultArray.push({
//           [monthName]: {
//             "Occupancy": {
//               hours: periodOccupancy,
//               subject,
//               account,
//               firstDay: Utilities.formatDate(period.from, Session.getScriptTimeZone(), "dd-MMM-yyyy"),
//               lastDay: Utilities.formatDate(period.to, Session.getScriptTimeZone(), "dd-MMM-yyyy"),
//               month: monthName,
//               finYear: getFinancialYear(monthName, year),
//               accountType: accountData[0].accountType
//             },
//             "Non-Occupancy": {
//               hours: periodNonOccupancy,
//               subject,
//               account,
//               firstDay: Utilities.formatDate(period.from, Session.getScriptTimeZone(), "dd-MMM-yyyy"),
//               lastDay: Utilities.formatDate(period.to, Session.getScriptTimeZone(), "dd-MMM-yyyy"),
//               month: monthName,
//               finYear: getFinancialYear(monthName, year),
//               accountType: accountData[0].accountType
//             }
//           }
//         });
//       });
//     }
//   }
//   console.log("resultArray: " + JSON.stringify(resultArray, null, 2));
//   return resultArray;
// }

