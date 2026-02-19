const DATE_COLUMN_INDEX = 11//9
const SME_NAME_COLUMN_INDEX = 5//3
const SME_ID_COLUMN_INDEX = 4
const MASTER_SHEET = SpreadsheetApp.openById("1u4i5rMCWM0mSgDWdf6EmGVi40kb9i2Qx51SE1cupIjU");
//"1eJ-7nKeICwYXEBh4oEzLwC6SpkYzCw5ntYxarbD2jxA";
const REVIEWER_ID_INDEX = 2;  //2

function doGet(e) {
  var sheetName = "QA DB";
  var sheet = MASTER_SHEET.getSheetByName(sheetName);
  var fullData = sheet.getDataRange().getValues();
  const header = fullData[0];
  let data = fullData.slice(1);

  // Extract and validate the parameters from 'e'.
  var userType = e.parameter.userType;
  //var smeName = e.parameter.smeName;
  var smeId = e.parameter.smeId;
  //var reviewerEmail = e.parameter.reviewerEmail || "";
  var qaReviewerId = e.parameter.qaReviewerId || "";
  var startDate = e.parameter.startDate;
  var endDate = e.parameter.endDate;
  console.log("parameters", e.parameter);



  // Convert startDate and endDate to the script's timezone
  if (startDate && endDate) {
    // startDate = convertToScriptTimeZone(new Date(startDate));
    // endDate = convertToScriptTimeZone(new Date(endDate));
    startDate = new Date(startDate).getTime();
    endDate = new Date(endDate).getTime();

  }

  // Depending on the user type, filter the data accordingly
  if (startDate && endDate) {
    switch (userType) {
      case "sme":
        if (smeId) {
          //data = data.filter(row => row[SME_NAME_COLUMN_INDEX] === smeName);
          data = data.filter(row => row[SME_ID_COLUMN_INDEX] === smeId);

          data = data.filter(row => {
            var rowDate = new Date(row[DATE_COLUMN_INDEX]);
            // rowDate = convertToScriptTimeZone(rowDate);
            rowDate = rowDate.getTime();

            //return row[SME_NAME_COLUMN_INDEX] === smeName && rowDate >= startDate && rowDate <= endDate;
            return row[SME_ID_COLUMN_INDEX] === smeId && rowDate >= startDate && rowDate <= endDate;
          });
        } else {
          data = { "error": "Missing SME name" };
        }
        break;

      case "reviewer":
        if (qaReviewerId) {     //qaReviewerId
         console.log("qareviwer id",qaReviewerId);
         console.log("sme id",smeId);
          data = data.filter(row => row[REVIEWER_ID_INDEX] === qaReviewerId);
          data = data.filter(row => {
            var rowDate = new Date(row[DATE_COLUMN_INDEX]);
            // rowDate = convertToScriptTimeZone(rowDate);
            rowDate = rowDate.getTime();
            Logger.log("Row Date RAW: " + row[DATE_COLUMN_INDEX] + " | Type: " + typeof row[DATE_COLUMN_INDEX]);

            if (smeId === 'All') {
             return row[REVIEWER_ID_INDEX] === qaReviewerId; //&& rowDate >= startDate && rowDate <= endDate;
            }
            else {
              // return row[REVIEWER_ID_INDEX] === qaReviewerId && row[SME_NAME_COLUMN_INDEX] === smeName
              //   && rowDate >= startDate && rowDate <= endDate;

              return row[REVIEWER_ID_INDEX] === qaReviewerId && row[SME_ID_COLUMN_INDEX] === smeId;
                //&& rowDate >= startDate && rowDate <= endDate;
            }
          });
        } else {
          data = { "error": "Missing reviewer id" };
        }
        break;

      default:
        console.log("Error: Invalid user type");
        data = { "error": "Invalid user type" };
    }
  } else {
    console.log("Error: Missing date parameters");
    data = { "error": "Missing date parameters" };
  }

  var outputData = [header].concat(data);

  return ContentService.createTextOutput(JSON.stringify(outputData))
    .setMimeType(ContentService.MimeType.JSON);
}

// function testDoGet() {
//   const mockEvent = {
//     parameter: {
//       userType: "reviewer",
//       smeName: "Harsha Patil (E510)", 
//       qaReviewerId: "ef3418c4-1c12-4013-9d51-00076843f6eb", 
//       startDate: "03-Jul-25",
//       endDate: "04-Jul-25"    
//     }
//   };

//   const response = doGet(mockEvent);
//   Logger.log("content:-"+response.getContent());
// }



//------------------------------------------------------------------------------
//old code do get 
// const DATE_COLUMN_INDEX = 9
// const SME_NAME_COLUMN_INDEX = 3
// const MASTER_SHEET = SpreadsheetApp.openById("1u4i5rMCWM0mSgDWdf6EmGVi40kb9i2Qx51SE1cupIjU");
// //"1eJ-7nKeICwYXEBh4oEzLwC6SpkYzCw5ntYxarbD2jxA";
// const REVIEWER_EMAIL_INDEX = 2

// function doGet(e) {
//   var sheetName = "QA DB";
//   var sheet = MASTER_SHEET.getSheetByName(sheetName);
//   var fullData = sheet.getDataRange().getValues();
//   const header = fullData[0];
//   let data = fullData.slice(1);

//   // Extract and validate the parameters from 'e'.
//   var userType = e.parameter.userType;
//   var smeName = e.parameter.smeName;
//   var reviewerEmail = e.parameter.reviewerEmail || "";
//   var startDate = e.parameter.startDate;
//   var endDate = e.parameter.endDate;
 


//   // Convert startDate and endDate to the script's timezone
//   if (startDate && endDate) {
//     startDate = convertToScriptTimeZone(new Date(startDate));
//     endDate = convertToScriptTimeZone(new Date(endDate));
    
//   }
  
//   // Depending on the user type, filter the data accordingly
//   if (startDate && endDate) {
//     switch(userType) {
//       case "sme":
//         if (smeName) {
//           data = data.filter(row => row[SME_NAME_COLUMN_INDEX] === smeName);
//           data = data.filter(row => {
//             var rowDate = new Date(row[DATE_COLUMN_INDEX]);
//             rowDate = convertToScriptTimeZone(rowDate);
            
//             return row[SME_NAME_COLUMN_INDEX] === smeName && rowDate >= startDate && rowDate <= endDate;
//           });
//         } else {
//           data = {"error": "Missing SME name"};
//         }
//         break;

//       case "reviewer":
//         if (reviewerEmail) {
          
//           data = data.filter(row => row[REVIEWER_EMAIL_INDEX] === reviewerEmail);
//           data = data.filter(row => {
//             var rowDate = new Date(row[DATE_COLUMN_INDEX]);
//             rowDate = convertToScriptTimeZone(rowDate);
            
//             Logger.log("Row Date RAW: " + row[DATE_COLUMN_INDEX] + " | Type: " + typeof row[DATE_COLUMN_INDEX]);

//             if (smeName === 'All'){
//               return row[REVIEWER_EMAIL_INDEX] === reviewerEmail && rowDate >= startDate && rowDate <= endDate;
//             }
//             else {
//               return row[REVIEWER_EMAIL_INDEX] === reviewerEmail && row[SME_NAME_COLUMN_INDEX] === smeName 
//               && rowDate >= startDate && rowDate <= endDate;
//             }
//           });
//           // data = data.filter(row => {

//           //   var rowDate = new Date(row[DATE_COLUMN_INDEX]);
//           //   rowDate = convertToScriptTimeZone(rowDate);
//           //   return rowDate >= startDate && rowDate <= endDate;
             
//           // });
//         } else {
//           data = {"error": "Missing reviewer email"};
//         }
//         break;

//       default:
//         console.log("Error: Invalid user type");
//         data = {"error": "Invalid user type"};
//     }
//   } else {
//     console.log("Error: Missing date parameters");
//     data = {"error": "Missing date parameters"};
//   }

// //   var outputData = [header].concat(data);
  
// //   return ContentService.createTextOutput(JSON.stringify(outputData))
// //                        .setMimeType(ContentService.MimeType.JSON);
//  }


