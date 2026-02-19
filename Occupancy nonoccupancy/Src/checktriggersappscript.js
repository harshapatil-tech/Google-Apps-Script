function auditScraperSheetById() {
  const SCRAPER_SHEET_ID = "1XH5eGI4thuz-yAsHDeVbI90KFX2AWee3w556yUcGe6I"; //  main Scraper sheet ID
  const ss = SpreadsheetApp.openById(SCRAPER_SHEET_ID);

  // === Get all sheet names in Scraper ===
  const sheetNames = ss.getSheets().map(s => s.getName());

  // === Prepare or reset Audit_Scraper sheet ===
  let auditSheet = ss.getSheetByName("Audit_Scraper");
  if (!auditSheet) auditSheet = ss.insertSheet("Audit_Scraper");
  else auditSheet.clear();

  auditSheet.getRange(1, 1, 1, 5).setValues([
    ["Function Name", "Accessed Sheets", "Trigger Type", "Frequency / Notes", "External Sheets"]
  ]);

  const functionsInfo = [];

  // === 1 Scan all functions in the project ===
  const allFuncs = Object.getOwnPropertyNames(this).filter(f => typeof this[f] === "function");

  allFuncs.forEach(fnName => {
    try {
      const fnText = this[fnName].toString();

      // Sheets accessed within the Scraper file
      const accessedSheets = sheetNames.filter(sName =>
        fnText.includes(`getSheetByName("${sName}")`) || fnText.includes(`getSheetByName('${sName}')`)
      );

      // External sheet IDs found via openById()
      const externalSheets = new Set(); // use Set to avoid duplicates
      const regex = /openById\s*\(\s*["']([a-zA-Z0-9-_]+)["']\s*\)/g;
      let match;
      while ((match = regex.exec(fnText)) !== null) {
        const id = match[1];
        try {
          const fileName = DriveApp.getFileById(id).getName();
          // Clean up "Copy of ..." or duplicate names
          const cleanName = fileName.replace(/^Copy of\s*/i, "").trim();
          externalSheets.add(cleanName);
        } catch (e) {
          externalSheets.add(id); // fallback to ID if file not accessible
        }
      }

      if (accessedSheets.length > 0 || externalSheets.size > 0) {
        functionsInfo.push([
          fnName,
          accessedSheets.join(", ") || "N/A",
          "Manual Call",
          "",
          [...externalSheets].join(", ") || "None"
        ]);
      }
    } catch (e) {
      Logger.log("Error while scanning " + fnName + ": " + e);
    }
  });

  // === 2 Scan all Triggers in the project ===
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(tr => {
    const fn = tr.getHandlerFunction();
    let type = "Manual / Unknown";
    let freq = "";

    try {
      const eventType = tr.getEventType();
      type = eventType ? eventType.toString() : "CLOCK";
    } catch (e) {}

    // Detect trigger frequency
    if (tr.getTriggerSource && tr.getTriggerSource() === ScriptApp.TriggerSource.CLOCK) {
      freq = "Time-based (scheduled)";
    }

    const foundFn = functionsInfo.find(f => f[0] === fn);
    if (foundFn) {
      foundFn[2] = type;
      foundFn[3] = freq;
    } else {
      functionsInfo.push([fn, "Unknown / Check Function", type, freq, ""]);
    }
  });

  // === 3 Write results to Audit_Scraper sheet ===
  if (functionsInfo.length > 0) {
    auditSheet.getRange(2, 1, functionsInfo.length, 5).setValues(functionsInfo);
  } else {
    auditSheet.getRange(2, 1).setValue("No functions or triggers detected for this project.");
  }

  // === 4 Logger + Completion message ===
  functionsInfo.forEach(f => Logger.log(f.join(" | ")));
  auditSheet.getRange(auditSheet.getLastRow() + 2, 1).setValue(" Scraper Sheet Audit Complete!");
  Logger.log(" Audit complete! Check 'Audit_Scraper' tab in the Scraper sheet.");
}
