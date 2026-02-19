/**
 * Creates appointment and NDA letters for employees based on the data in the active sheet.
 * It checks for required fields and updates the sheet accordingly.
 */
function createButton() {
  // Get the currently active spreadsheet and its active sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Retrieve headers and data from the sheet using CentralLibrary function
  const [header, data] = CentralLibrary.get_Data_Indices_From_Sheet(sheet, 2);
  
  // Initialize arrays to keep track of valid rows and errors
  const validRows = [];
  const errors = [];
  
  let createdCount = 0; // Counter for the number of letters successfully created

  // Loop through each row of data
  for (let i = 0; i < data.length; i++) {
    var row = data[i]; // Get the current row of data
    const trigger = row[header["Appointment & NDA Letter Trigger"]]; // Check if the trigger is true
    const created = row[header["Appointment & NDA Letter Created?"]]; // Check if the letter has already been created

    // Proceed only if the trigger is true and the letter has not been created yet
    if (trigger === true && created == "") {
      let allFieldsFilled = true; // Flag to check if all required fields are filled
      // List of required fields that must be filled to create a letter
      const requiredFields = [
        // "Employee Name", "Designation", "Department", "DOJ", 
        // "PAN Card No", "Permanent Address", "Days", "Hours", 
        // "Code", "Probation Tenure", "Notice Period before 1year", 
        // "Notice Period after 1year", "HR Name", "Signing Authority"
      ];

      // Check if all required fields are filled
      for (let j = 0; j < requiredFields.length; j++) {
        // If any required field is empty, set the flag to false and exit the loop
        if (row[header[requiredFields[j]]] == "") {
          allFieldsFilled = false;
          break;
        }
      }
      
      // If all required fields are filled, generate the letter
      if (allFieldsFilled) {
        // Call the function to generate the appointment letter
        generateLetter(row, header); // Generate the letter based on the current row's data
        
        // Log the values for debugging
        console.log(sheet.getRange(i + 4, 1, 1).getValues());

        // Set "Y" in the 'Appointment & NDA Letter Created?' column to indicate success
        sheet.getRange(i + 4, header["Appointment & NDA Letter Created?"] + 1).setValue("Y");
        // Reset the trigger to false after processing
        sheet.getRange(i + 4, header["Appointment & NDA Letter Trigger"] + 1).setValue(false);
        
        createdCount++; // Increment the success counter
      } else {
        // Collect the row number if required fields are not filled
        errors.push(i + 4); // Store the 1-based row index where the error occurred
      }
    }
  }

  // After processing all rows, display a summary message about the results
  if (errors.length > 0) {
    Browser.msgBox(
      createdCount + " letters were successfully created & " +
      errors.length + " letters had errors. Errors occurred in rows: " +
      errors.join(", ")
    );
  } else {
    Browser.msgBox(createdCount + " letters were successfully created with no errors.");
  }
}


/**
 * Generates a letter for an employee based on the provided data row and header mapping.
 * @param {Array} row - An array containing employee data from the spreadsheet.
 * @param {Object} header - An object mapping header names to their respective indices.
 */
function generateLetter(row, header) {
  const templateId = "18Rw2YMcbA8IoNNGRIYjFMO67ZEZSCudcO31Hb1XV1yk"   //"1gz219sPDO9FW4PFF7Ric1MwemIDx42RT"; // ID of the template document
  // https://docs.google.com/document/d/18Rw2YMcbA8IoNNGRIYjFMO67ZEZSCudcO31Hb1XV1yk/edit?tab=t.0
  const outputFolder = DriveApp.getFolderById('1i-3jUiRQtHBwG_JN0aG0j5836qK5K3rm'); // Output folder for generated letters  
  // const templateFile = DriveApp.getFileById(docId); // Get the template file from Google Drive
  
  // Prepare a resource object for copying the template to Google Docs format
  // const resource = {
  //   title: 'Appointment & NDA Letter_ 1.3 - ' + row[header["Employee Name"]],
  //   mimeType: MimeType.GOOGLE_DOCS,
  //   parents: [{ id: outputFolder.getId() }] // Specify the parent folder for the new document
  // };

  const fileName = 'Appointment & NDA Letter_ 1.3 - ' + row[header["Employee Name"]];

  // Use Advanced Drive Service to copy the .docx template to Google Docs format
  const newFile = DriveApp.getFileById(templateId).makeCopy(fileName, outputFolder);

  // const copiedFile = Drive.Files.copy(resource, docId); // Copy the file
  const newDocFileId = newFile.getId(); // Get the ID of the newly created document
  console.log("New google doc file id: ", newDocFileId); // Log the new document ID for debugging
  
  let newDoc; // Declare variable for the new document
  const maxRetries = 3; // Set maximum number of retries for opening the document
  // Attempt to open the new Google Doc, with retries if it fails
  for (let i = 0; i < maxRetries; i++) {
    try {
      newDoc = DocumentApp.openById(newDocFileId); // Open the document by its ID
      console.log('Document opened successfully.'); // Log success message
      break; // Exit loop if successful
    } catch (error) {
      // Log warning if the attempt fails
      console.warn(`Attempt ${i + 1} to open the document failed. Retrying...`);
      Utilities.sleep(1000); // Wait 1 second before retrying
    }
  }

  // If unable to open the document after all retries, throw an error
  if (!newDoc) {
    throw new Error("Failed to open the new document after multiple attempts.");
  }

  // Get the body of the new document to perform text replacements
  const body = newDoc.getBody();
  
  // Perform text replacements in the document based on the employee's data
  const date = new Date(); // Get the current date for potential use in the letter
  body.replaceText('<Name>', row[header["Employee Name"]]); // Replace placeholders with actual data
  body.replaceText('<Employee Name>', row[header["Employee Name"]]);
  body.replaceText('<Employee Designation>', row[header["Designation"]]);
  body.replaceText('<Department>', row[header["Department"]]);
  body.replaceText('<DOJ>', Utilities.formatDate(row[header["DOJ"]], Session.getScriptTimeZone(), 'MMMM dd, YYYY')); // Format the DOJ
  body.replaceText('<HR Name>', row[header["HR Name"]]);
  body.replaceText('<Code>', row[header["Code"]]);
  body.replaceText('<Year>', getFinancialYear(new Date(row[header["DOJ"]]))); // Get financial year based on DOJ
  // body.replaceText('<Address1>', row[header["Permanent Address"]]);
  body.replaceText('<Address>', row[header["Current Address"]]);
  body.replaceText('<probation tenure>', row[header["Probation Tenure"]]);
  body.replaceText('<6>', row[header["Days"]]);
  body.replaceText('<36>', row[header["Hours"]]);
  // body.replaceText('<notice period after 1 year>', row[header["Notice Period after 1year"]]);
  // body.replaceText('<notice period before 1 year>', row[header["Notice Period before 1year"]]);
  body.replaceText('<HR Designation>', getHrDesignation(row[header["HR Name"]])); // Get HR designation
  body.replaceText('<Signing Authority Name>', row[header["Signing Authority"]]);
  body.replaceText('<Signing Authority>', row[header["Signing Authority"]]);
  body.replaceText('<Signing Authority Designation>', getSigningAuthority(row[header["Signing Authority"]])); // Get signing authority designation
  body.replaceText('<Designation>', getSigningAuthority(row[header["Signing Authority"]])); // Replace with authority designation
  body.replaceText('<PAN Card No>', row[header["PAN Card No."]]); // Replace PAN Card No
  body.replaceText('<PAN>', row[header["PAN Card No."]]); // Additional replacement for PAN
  body.replaceText('<Appointment Letter Reference>', row[header["Code"]]); // Replace reference code

  // Save and close the newly modified document
  newDoc.saveAndClose();
  
  // Export the modified Google Doc as a .docx file
  // const docxBlob = convToMicrosoft(newDocFileId); // Convert the document to .docx format
  
  // Define the filename for the .docx file
  // const docxFileName = 'Appointment & NDA Letter_ 1.3 - ' + row[header["Employee Name"]] + '.docx';
  // const docxFileName = 'Appointment & NDA Letter_ 1.3 - ' + row[header["Employee Name"]];
  // Save the .docx file to the specified Drive folder
  // outputFolder.createFile(docxBlob).setName(docxFileName);

  // Optionally, delete the temporary Google Doc to clean up
  // DriveApp.getFileById(newDocFileId).setTrashed(true);
  
}


/**
 * Converts a Google Drive file to a specified Microsoft format.
 * @param {string} fileId - The ID of the file to be converted.
 * @returns {Blob|null} The converted file as a Blob, or null if the format is unsupported.
 * @throws Will throw an error if the file ID is null.
 */
function convToMicrosoft(fileId) {
  // Validate that a file ID is provided
  if (fileId == null) throw new Error("No file ID."); // Throw error if null
  
  var file = DriveApp.getFileById(fileId); // Get the file from Google Drive by ID
  var mime = file.getMimeType(); // Get the MIME type of the file
  var format = ""; // Initialize variable for format type
  var ext = ""; // Initialize variable for file extension

  // Determine the format and extension based on the MIME type
  switch (mime) {
    case "application/vnd.google-apps.document":
      format = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"; // .docx format
      ext = ".docx"; // Extension for Word documents
      break;
    case "application/vnd.google-apps.spreadsheet":
      format = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"; // .xlsx format
      ext = ".xlsx"; // Extension for Excel spreadsheets
      break;
    case "application/vnd.google-apps.presentation":
      format = "application/vnd.openxmlformats-officedocument.presentationml.presentation"; // .pptx format
      ext = ".pptx"; // Extension for PowerPoint presentations
      break;
    default:
      return null; // Return null if unsupported MIME type
  }

  // Construct the URL for exporting the file in the specified format
  var url = "https://www.googleapis.com/drive/v3/files/" + fileId + "/export?mimeType=" + format;

  // Fetch the file as a Blob using the constructed URL and OAuth token for authorization
  var blob = UrlFetchApp.fetch(url, {
    method: "get", // HTTP method for fetching
    headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()}, // Include authorization token
    muteHttpExceptions: true // Prevent HTTP exceptions from being thrown
  }).getBlob(); // Get the file as a Blob
  
  return blob; // Return the Blob containing the converted file
}
