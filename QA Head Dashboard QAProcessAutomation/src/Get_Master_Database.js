//Get All Data from the Sheets 
class MasterDB {
  constructor(sheetId, topicTabId, smeTabId, accountTabId, otherTabId, reviewerTabId) {
    const spreadSheet = CentralLibrary.DataAndHeaders(sheetId);
    //console.log(spreadSheet);

    //1.Topic sheet
    const topicWrapper = spreadSheet.getSheetById(topicTabId);
    //console.log(topicWrapper);
    this.topicSheet = topicWrapper.sheet;
    //console.log(this.topicSheet);
    const [topicHeaders, topicData] = topicWrapper.getDataIndicesFromSheet(1);
    //console.log(topicHeaders,topicData);
    this.topicHeaders = topicHeaders;
    this.topicData = topicData;

    //2.SME Sheet
    const smeWrapper = spreadSheet.getSheetById(smeTabId);
    //console.log(smeWrapper);
    this.smeSheet = smeWrapper.sheet;
    const [smeHeaders, smeData] = smeWrapper.getDataIndicesFromSheet(0);
    //console.log(smeHeaders,smeData);
    this.smeHeaders = smeHeaders;
    this.smeData = smeData;

     //3.Account Sheet
    const accWrapper = spreadSheet.getSheetById(accountTabId);
    //console.log(accWrapper);
    this.accountSheet = accWrapper.sheet;
    const [accountHeaders, accountData] = accWrapper.getDataIndicesFromSheet(0);
    //console.log(accountHeaders,accountData);
    this.accountHeaders = accountHeaders;
    this.accountData = accountData;

    //4.Other sheet
    const otherWrapper = spreadSheet.getSheetById(otherTabId);
    //console.log(otherWrapper);
    this.otherSheet = otherWrapper.sheet;
    //console.log(this.otherSheet);
    const [otherHeaders, otherData] = otherWrapper.getDataIndicesFromSheet(0);
    //console.log(otherHeaders,otherData);
    this.otherHeaders = otherHeaders;
    //console.log(this.otherHeaders);
    this.otherData = otherData;
    //console.log(this.otherdata);

    //5.Reviwer Sheet
    const reviewerWrapper = spreadSheet.getSheetById(reviewerTabId);
    //console.log(reviewerWrapper);
    this.reviewerSheet = reviewerWrapper.sheet;
    const [reviewerHeaders, reviewerData] = reviewerWrapper.getDataIndicesFromSheet(0);
    //console.log(reviewerHeaders, reviewerData);
    this.reviewerHeaders = reviewerHeaders;
    this.reviewerData = reviewerData;
  }


  //Filtered reviwer Data By Reviewer name and email id
  getReviewerEmailByName(reviewerName) {  
    const nameIdx = this.reviewerHeaders["Reviewer Name"];
    //console.log("Name are:-",nameIdx);
    const emailIdx = this.reviewerHeaders["Email ID"];
    //console.log("Email is:-",emailIdx);
    if (nameIdx === undefined || emailIdx === undefined) return null;

    const matches = this.reviewerData.filter(r =>
      //r[nameIdx]?.toString().trim().toLowerCase() === name.trim().toLowerCase()
      r[nameIdx]?.toString().trim().toLowerCase() === reviewerName.trim().toLowerCase()

    );
    //console.log("matches are",matches);

    return matches.length > 0 ? matches[0][emailIdx] : null;
  }


  //filtered other Data 
  getOtherDataFiltered() {
     const data = this.otherData;
     //console.log(data);
     const val = key => data.map(r => r[this.otherHeaders[key]]).filter(r => r !== "");
     //console.log("Filtered Other Data:", [val("Client")]);

      return [
      val("Client"),
      val("Mode"),
      val("Audio"),
      val("Rating"),
      val("Reasons for negative ratings"),
      val("Client Complaints"),
     // val("Mapping"),
      val("Discussion"),
      val("SubjectKnowledge_Identify"),
      val("SubjectKnowledge_Break The Process"),
      val("SubjectKnowledge_Explanation"),
      val("Tutoring_Encourage"),
      val("Tutoring_Session Flow"),
      val("Tutoring_Socratic"),
      val("Admin_Greeting/ closing"),
      val("Admin_Client policies"),
      val("Communication_English"),
      val("Communication_Effectiveness"),
      //val("Score of low rated sessions"),
      val("NetTutor Client Ratings (Out of Five)"),
      val("Input to Training team")
    ];
  }
 
  //filtered topicdata, smeData, accountData 
  filterFor(reviewerName, department){
    const topicRows = this.topicData;
    //console.log("Topic Rows are:-",topicRows);
    const deptIdx = this.topicHeaders['Department'];
    //console.log("department index:-",deptIdx);
    const subjectIdx = this.topicHeaders['Subject'];
    //console.log("Subject index:-",subjectIdx);
    const topicIdx = this.topicHeaders['Topic'];
    //console.log("Topic index:-",topicIdx);
    const subTopicIdx = this.topicHeaders['SubTopic'];
    //console.log("Sub topic index:-",subTopicIdx);
    
    //1.topicData
    const filteredTopics = topicRows
      .filter(r => r[deptIdx] === department)
      .map(r => [r[subjectIdx], r[topicIdx], r[subTopicIdx]]);
    //console.log("Filter topic Data:-",filteredTopics);
    
    //2.smeData
     
    const smeData = this.smeData.filter(r =>
      (r[this.smeHeaders["QA Reviewer"]] === reviewerName ||
        r[this.smeHeaders["QA Reviewer 2"]] === reviewerName) &&
      r[this.smeHeaders["Department"]] === department &&
      (
        r[this.smeHeaders["Active?"]] === true ||
        (!r[this.smeHeaders["Active?"]] && !r[this.smeHeaders["Removed Date"]])
      )
    ).map(r => [r[this.smeHeaders["Unique ID"]], r[this.smeHeaders["SME Name"]]]);
   console.log("filter smeData:-", smeData);

   const accountRows = this.getAccountNumbersByDepartment(department); //"Statistics"
   //console.log("Filtered account Data:-",accountRows);

   const otherData = this.getOtherDataFiltered();
   //console.log("Filtered Other Data:-",otherData);

   const email = this.getReviewerEmailByName(reviewerName);//"Harsha Patil (E510)"
   console.log("Filtered Reviwer Data:-",email);

   return {
    filteredTopics,
    smeData,
    accountRows,
    otherData,
    reviewerEmail: email
  };
 }


  //3.accountdata
  getAccountNumbersByDepartment(department) {
  const deptHeaders = this.accountHeaders;
  //console.log("Account Headers:", deptHeaders); 

  const deptData = this.accountData;
  //console.log("Account Data (first 3 rows):", deptData.slice(0, 3)); 
  const deptIdx = deptHeaders[department];
  //console.log(`Index for department "${department}":`, deptIdx); 
  if (deptIdx === undefined) {
    //console.warn(`Department '${department}' not found in account headers.`);
    return [];
  }

  const rawValues = deptData.map(r => r[deptIdx]);
  //console.log(`Raw values for "${department}":`, rawValues); // Before filtering

  const filteredValues = rawValues.filter(r => Boolean(r) && r !== "NA");
  //console.log(`Filtered account numbers for "${department}":`, filteredValues); 
  return filteredValues;
    
  } 
}


//class BackendPopulator
class BackendPopulator{
  static populate(sheet, qaAdd, qaUpdate, [topicData, smeData, accountData, otherData]){
  const headerRowIdx = 4;
  console.log(sheet.getName())
  console.log("Header index:-",headerRowIdx);
  clearRowsBelow(sheet, headerRowIdx);
  const header = sheet.getRange(headerRowIdx, 1, 1, sheet.getLastColumn()).getValues();
  console.log("Heades are:-",header);
  

   const headerRow = header[0];
   const idx = name => headerRow.indexOf(name) + 1;

  //const idx = name => header.indexOf(name) + 1;
  // console.log("Column index for 'uniqueID':", idx("Unique ID"));
  
  const fill = (data, name, colSpan = 1) => {
  const columnIndex = idx(name); // 1-based column
  if (columnIndex <= 0) {
    console.warn(`Column "${name}" not found in header row.`);
    return;
  }
  if (data.length > 0) {
    sheet.getRange(headerRowIdx + 1, columnIndex, data.length, colSpan)
         .setValues(colSpan === 1 ? data.map(r => [r]) : data);
  }
};

     const [
      clients, mode, audio, ratings, negReviews, clientComplaints, //mappings,
      discussions, identity, breakProcess, explanation, encourage, tutoringFlow,
      socratic, greetings, policies, englishFlow, effectiveFlow,  netTutor,inputToTraining
    ] = otherData;     //lowRateScores,
   
    //fill(smeData, "Unique ID", 2); // Unique ID + SME Name
    fill(smeData.map(r => [r[0]]), "Unique ID");
    fill(smeData.map(r => [r[1]]), "SME Name");
    fill(clients, "Client");
    fill(topicData, "Subject", 3); // Subject + Topic + SubTopic
    fill(topicData.map(r => [r[1]]), "Topic"); // Topic
    fill(topicData.map(r => [r[2]]), "Sub-Topic");
    fill(accountData, "Account number");
    fill(mode, "Mode");
    fill(audio, "Audio");
    fill(ratings, "Rating\n(Negative/Positive/Low)");
    fill(negReviews, "Reason for negative rating");
    fill(clientComplaints, "Client Complaint");
    fill(discussions, "Discussion");
    fill(identity, "SubjectKnowledge_Identify");
    fill(breakProcess, "SubjectKnowledge_Break The Process");
    if (explanation.length > 0)
      applyCustomFormatting(sheet.getRange(headerRowIdx + 1, idx("SubjectKnowledge_Explanation"), explanation.length, 1)).setValues(explanation.map(r => [r]));
    fill(encourage, "Tutoring_Encourage");
    fill(tutoringFlow, "Tutoring_Session Flow");
    fill(socratic, "Tutoring_Socratic");
    fill(greetings, "Admin_Greeting/ closing");
    fill(policies, "Admin_Client policies");
    fill(englishFlow, "Communication_English");
    fill(effectiveFlow, "Communication_Effectiveness");
    //fill(lowRateScores, "Score of low rated sessions");
    fill(netTutor, "NetTutor Client Ratings (Out of Five)");
    fill(inputToTraining,"Input to Training team");

    setDropdownsQAReviewAdd(qaAdd, qaUpdate, sheet, 2);
    qaReviewUpdateSetDropdowns(qaUpdate, sheet);
  }
}


function setReviewerUniqueId(backendSheet, reviewerName, department, smeData, smeHeaders) {
  
  const reviewerRow = smeData.find(r =>
    r[smeHeaders["SME Name"]] === reviewerName &&
    r[smeHeaders["Department"]] === department
  );

  const reviewerUniqueId = reviewerRow
    ? reviewerRow[smeHeaders["Unique ID"]]
    : "Not Found";

  
  backendSheet.getRange("E1").setValue("Reviewer Unique ID");
  backendSheet.getRange("F1").setValue(reviewerUniqueId);
}


function createBackendReviewerSheet(url, department) {
  const ss = SpreadsheetApp.openByUrl(url);
  console.log(ss.getName())
  const reviewerName = ss.getName().split("_").slice(2).join(" ");
 
  const db = new MasterDB(
    //SME_DB_SPREADSHEET_ID,
    MASTER_DB_SPREADSHEET_ID,
    BACKEND_TOPIC_TAB_ID,
    // SME_TAB_ID,
    SME_DB_TAB_ID,
    BACKEND_ACCOUNT_TAB_ID,
    BACKEND_OTHER_TAB_ID,
    // REVIEWER_TAB_ID,
    REVIEWER_DB_TAB_ID
 );
  
  const {
    filteredTopics: topics,
    smeData: smes,
    accountRows: accounts,
    otherData: other
  } = db.filterFor(reviewerName, department);

  const backend = ss.getSheetByName("Backend");
  const qaAdd = ss.getSheetByName("QA_Review_Add");
  const qaUpdate = ss.getSheetByName("QA_Review_Update");

  setReviewerUniqueId(backend, reviewerName, department, db.smeData, db.smeHeaders);
  BackendPopulator.populate(backend, qaAdd, qaUpdate, [topics, smes, accounts, other]);
  
  if (ss.getOwner().getEmail() !== AUTOMATION_EMAIL) {
    DriveApp.getFileById(ss.getId()).setOwner(AUTOMATION_EMAIL);
  }
}

function tempFunction() {
  createBackendReviewerSheet(
  //"https://docs.google.com/spreadsheets/d/17t_fKuXAkXtoVnB4FmCRDIHyJzdiAzIYGbQR0eEdeQM/edit","Mathematics"

   "https://docs.google.com/spreadsheets/d/1I6K0B3Dc-Wj4VWo1lTd49hO_a_DPAlgdSHykMrNoM8E/edit","Mathematics"
   
  );
 
}





//___________________________OLD CODE____________________________________________________
// function getMasterDBData(reviwerName, department){
//   const spreadsheet = SpreadsheetApp.openById("1u4i5rMCWM0mSgDWdf6EmGVi40kb9i2Qx51SE1cupIjU");
//   const topicSheet = getSheetById(spreadsheet, 606654122);
//   //const reviwerSheet = getSheetById(spreadsheet, 541621095);
//   const smeSheet = getSheetById(spreadsheet, 1777024406);
//   const accountSheet = getSheetById(spreadsheet, 1310786449)
//   const otherSheet = getSheetById(spreadsheet, 154120686)
  
//   const topicData = getTopicData(department, topicSheet);
//   //const reviwerData = getReviwerData(reviwerName, reviwerSheet);
//   const smeData = getSMEData(reviwerName, department, smeSheet);
//   Logger.log("SME Data: " + JSON.stringify(smeData));
//   const accountData = getAccountNums(department, accountSheet);
//   const otherData = getOtherData(otherSheet)
//   return [topicData, smeData, accountData, otherData]
// }


// function getTopicData(department, sheet){
//   const totalData = sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
//   const header = totalData[0], data = totalData.slice(1);
//   const departmentIdx = header.indexOf('Department');
//   const subjectIdx = header.indexOf('Subject');
//   const topicIdx = header.indexOf('Topic');
//   const subTopicIdx = header.indexOf('SubTopic');
  
//   return data.filter(row => row[departmentIdx].trim().toLowerCase() === department.trim().toLowerCase())
//              .map(r => [r[subjectIdx], r[topicIdx], r[subTopicIdx]]);
// }


// function getReviwerData(name, sheet){
//   const totalData = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
//   const header = totalData[0], data = totalData.slice(1);
//   const emailIdx = header.indexOf('Email ID');
//   const reviwerNameIdx = header.indexOf('Reviewer Name');

//   return data.filter(r=> r[reviwerNameIdx].trim().toLowerCase() === name.trim().toLowerCase())[0][emailIdx];
// }


// function getSMEData(reviwerName, department, sheet){
//   const totalData = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
//   const header = totalData[0], data = totalData.slice(1);
  
//   const departmentIdx = header.indexOf('Department');
//   const smeNameIdx = header.indexOf('SME Name');
//   const reviewerIdx = header.indexOf('QA Reviewer');
//   const reviewer2Idx = header.indexOf('QA Reviewer 2');
//   const addedDateIdx = header.indexOf("Added Date");
//   const activeIdx = header.indexOf("Active?");
//   const removedDateIdx = header.indexOf("Removed Date");
 
//   return data.filter(r=> (r[reviewerIdx] === reviwerName || r[reviewer2Idx] === reviwerName) && (r[departmentIdx] === department))
//              .filter(r=> (r[removedDateIdx] === '' && r[activeIdx] === false) || 
//                          (r[activeIdx]===true && r[removedDateIdx] ==='') ||
//                          (r[activeIdx]===true && r[removedDateIdx] !==''))
//              .map(r => r[smeNameIdx]);
  
// }


// function getAccountNums(department, sheet){
//   const totalData = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
//   const header = totalData[0], data = totalData.slice(2);
//   const departmentIdx = header.indexOf(department);

//   return data.map(r => r[departmentIdx]).filter(r => Boolean(r));
// }


// function getOtherData(sheet){
//   const totalData = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
//   const header = totalData[0], data = totalData.slice(1);

//   const clientIdx = header.indexOf('Client');
//   const modeIdx = header.indexOf('Mode');
//   const audioIdx = header.indexOf('Audio');
//   const ratingIdx = header.indexOf('Rating');
//   const negReviewReasonIdx = header.indexOf('Reasons for negative ratings');
//   const clientComplaintsIdx = header.indexOf('Client Complaints');
//   const mappingIdx = header.indexOf('Mapping');
//   const discussionIdx = header.indexOf('Discussion');
//   const identyIdx = header.indexOf('SubjectKnowledge_Identify');
//   const breakProcessIdx = header.indexOf('SubjectKnowledge_Break The Process');
//   const explanationIdx = header.indexOf('SubjectKnowledge_Explanation');
//   const encourageIdx = header.indexOf('Tutoring_Encourage');
//   const tutoringFlowIdx = header.indexOf('Tutoring_Session Flow');
//   const socraticIdx = header.indexOf('Tutoring_Socratic');
//   const greetingIdx = header.indexOf('Admin_Greeting/ closing');
//   const policiesIdx = header.indexOf('Admin_Client policies');
//   const englishFlowIdx = header.indexOf('Communication_English');
//   const effectiveFlowIdx = header.indexOf('Communication_Effectiveness');
//   const lowRateScoreIdx = header.indexOf("Score of low rated sessions");
//   const netTutorIdx = header.indexOf("NetTutor Client Ratings (Out of Five)");

//   const clients = data.map(r => r[clientIdx]).filter(r => Boolean(r));
//   const mode = data.map(r => r[modeIdx]).filter(r => Boolean(r));
//   const audio = data.map(r => r[audioIdx]).filter(r => Boolean(r));
//   const ratings = data.map(r => r[ratingIdx]).filter(r => Boolean(r));
//   const negReviews = data.map(r => r[negReviewReasonIdx]).filter(r => Boolean(r));
//   const clientComplaints = data.map(r => r[clientComplaintsIdx]).filter(r => Boolean(r));
//   const mappings = data.map(r => r[mappingIdx]).filter(r => Boolean(r));
//   const discussions = data.map(r => r[discussionIdx]).filter(r => Boolean(r));

//   const identity = data.map(r => r[identyIdx]).filter(r => r !== "");
//   const breakProcess = data.map(r => r[breakProcessIdx]).filter(r => r !== "");
//   const explanation = data.map(r => r[explanationIdx]).filter(r => r !== "");

//   const encourage = data.map(r => r[encourageIdx]).filter(r => r !== "");
//   const tutoringFlow = data.map(r => r[tutoringFlowIdx]).filter(r => r !== "");
//   const socratic = data.map(r => r[socraticIdx]).filter(r => r !== "");

//   const greetings = data.map(r => r[greetingIdx]).filter(r => r !== "");
//   const policies = data.map(r => r[policiesIdx]).filter(r => r !== "");

//   const englishFlow = data.map(r => r[englishFlowIdx]).filter(r => r !== "");
//   const effectiveFlow = data.map(r => r[effectiveFlowIdx]).filter(r => r !== "");

//   const lowRateScores = data.map(r => r[lowRateScoreIdx]).filter(r => r !== "");
//   const netTutor = data.map(r => r[netTutorIdx]).filter(r => r !== "");

//   return [clients, mode, audio, ratings, negReviews, clientComplaints, mappings, discussions, identity, breakProcess, explanation, 
//           encourage, tutoringFlow, socratic, greetings, policies, englishFlow, effectiveFlow, lowRateScores, netTutor];

// }

// function getSheetById(spreadsheet, id) {
//   return spreadsheet.getSheets().filter(
//     function(s) {return s.getSheetId() === id;}
//   )[0];
// }