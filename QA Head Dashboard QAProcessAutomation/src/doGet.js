function pushChangesParallel() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const reviewerIndexSheet = spreadsheet.getSheetByName("Index");

  const dataRangeReviewer = reviewerIndexSheet.getRange(2, 1, reviewerIndexSheet.getLastRow() - 1, reviewerIndexSheet.getLastColumn()).getValues();
  const headerReviewer = dataRangeReviewer[0];
  const dataReviewer = dataRangeReviewer.slice(1);

  const reviewerIndices = {
    srNoIdx: headerReviewer.indexOf("#"),
    emailIdx: headerReviewer.indexOf("QA Reviewer Email"),
    departmentIdx: headerReviewer.indexOf("Department"),
    sheetLinkIdx: headerReviewer.indexOf("Sheet Link"),
  };

  const filteredData = dataReviewer.filter(r =>
    r[reviewerIndices.srNoIdx] !== ''
    && r[reviewerIndices.emailIdx] !== '' && r[reviewerIndices.emailIdx] !== 'sreenjay.sen@upthink.com'
    && r[reviewerIndices.departmentIdx] !== '' && r[reviewerIndices.sheetLinkIdx] !== ''
  );

  // Split data into chunks
  const chunkSize = 10; // Adjust as needed
  const chunks = [];
  for (let i = 0; i < filteredData.length; i += chunkSize) {
    // Extract only department and link for each row in the chunk
    const chunk = filteredData.slice(i, i + chunkSize).map(r => ({
      department: r[reviewerIndices.departmentIdx],
      link: r[reviewerIndices.sheetLinkIdx]
    }));
    chunks.push(chunk);
  }

  // Clean up existing triggers
  const existingTriggers = ScriptApp.getProjectTriggers();
  existingTriggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processChunk') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Schedule triggers
  chunks.forEach((chunk, index) => {
    const triggerTime = new Date(new Date().getTime() + index * 60000); // 1 minute interval between triggers
    const trigger = ScriptApp.newTrigger('processChunk')
      .timeBased()
      .at(triggerTime)
      .create();

    PropertiesService.getScriptProperties().setProperty(`chunk_${index}`, JSON.stringify(chunk));
    PropertiesService.getScriptProperties().setProperty(trigger.getUniqueId(), `chunk_${index}`);
    console.log(`Scheduled trigger for chunk ${index} with ID: ${trigger.getUniqueId()}`);
  });
}

function processChunk() {
  try {
    const chunkIndex = getChunkIndex(); // Retrieve chunk index safely
    console.log("Chunk Index", chunkIndex);
    if (chunkIndex === undefined) {
      throw new Error('Chunk index could not be determined.');
    }
    const chunk = JSON.parse(PropertiesService.getScriptProperties().getProperty(`chunk_${chunkIndex}`));

    if (!chunk) {
      throw new Error(`Chunk data not found for index chunk_${chunkIndex}`);
    }

    chunk.forEach(r => {
      const link = r.link;
      const department = r.department;
      //createBackendReviwerSheetByDepartment(link, department);
      createBackendReviewerSheet(link,department)


    });

    // Clean up properties
    PropertiesService.getScriptProperties().deleteProperty(PropertiesService.getScriptProperties().getProperty(`chunk_${chunkIndex}`));
    PropertiesService.getScriptProperties().deleteProperty(`chunk_${chunkIndex}`);

  } catch (error) {
    console.error('Error processing chunk:', error);
  }
}

function getChunkIndex() {
  const triggerId = ScriptApp.getProjectTriggers().find(trigger => trigger.getHandlerFunction() === 'processChunk').getUniqueId();
  console.log(PropertiesService.getScriptProperties().getProperty(triggerId))
  return PropertiesService.getScriptProperties().getProperty(triggerId).split("_")[1];
}
