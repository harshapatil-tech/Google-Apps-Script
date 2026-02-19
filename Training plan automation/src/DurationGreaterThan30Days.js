function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;

  const TRAINING_TRACKER_SHEET_NAME = "Training Tracker"; // Update this with your sheet name
  const headersRow = 4; // Update this if your headers are in a different row

  // Ensure we're working on the correct sheet
  if (sheet.getName() !== TRAINING_TRACKER_SHEET_NAME) return;

  // Get the header names
  const headers = sheet.getRange(headersRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headers_TrainingTracker = headers.reduce((acc, header, index) => {
    acc[header] = index + 1; // +1 to make it 1-based index
    return acc;
  }, {});

  const row = range.getRow();
  const column = range.getColumn();

  if (row > headersRow) { // Ensure we're not in the header row
    const startTimeCol = headers_TrainingTracker["Training Start Date "];
    const endTimeCol = headers_TrainingTracker["Training End Date"];
    const durationCol = headers_TrainingTracker["Training Duration in Days"];
    const leavesCol = headers_TrainingTracker["Leaves"];
    const departmentCol = headers_TrainingTracker["Department"];

    if (column === startTimeCol || column === endTimeCol || column == leavesCol) {
      const startTimeCell = sheet.getRange(row, startTimeCol);
      const endTimeCell = sheet.getRange(row, endTimeCol);
      const durationCell = sheet.getRange(row, durationCol);
      const leavesCell = sheet.getRange(row, leavesCol);
      const departmentCell = sheet.getRange(row, departmentCol);

      const startTime = startTimeCell.getValue();
      const endTime = endTimeCell.getValue();
      const leaves = leavesCell.getValue();
      department = departmentCell.getValue();
      if (department === "Statistics")
        NUM_DAYS_TRAINING = 60;
      if (startTime && endTime && leaves === "") {
        const durationInMs = new Date(endTime) - new Date(startTime);
        let durationInDays = CentralLibrary.getDaysDifference(endTime, startTime);
        durationInDays = durationInDays;
        console.log("executed non leaves", durationInDays)
        durationCell.setValue(durationInDays);
        if (durationInDays > NUM_DAYS_TRAINING)
          durationCell.setBackground("#e06666")
        else
          durationCell.setBackground("#ffffff")
      } else if (startTime && endTime && leaves !== "") {

        let durationInDays = CentralLibrary.getDaysDifference(endTime, startTime);
        durationInDays -= leaves;
        console.log("executed", durationInDays)               
        durationCell.setValue(durationInDays);
        if (durationInDays > NUM_DAYS_TRAINING)
          durationCell.setBackground("#e06666")
        else
          durationCell.setBackground("#ffffff")
      }else {
        durationCell.setValue(""); // Clear the cell if either time is missing
      }
    }
  }
}
