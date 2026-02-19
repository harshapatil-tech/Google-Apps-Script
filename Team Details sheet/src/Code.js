const DEPARTMENTS = ["Mathematics", "Statistics", "Physics", "Chemistry", "Biology", "Accounts", "Finance", "Economics", "Computer Science", "English"]


function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Update")
    .addItem("Update Backend", "run")
    .addToUi();
}


function run() {
  const employeeDetails = EmployeeDetails.getEmployeeData();
  const teamDetails = new TeamDetails();
  // teamDetails.updateBackendSheet(employeeDetails);
  teamDetails.updateTeamDataSheet();
}


class TeamDetails {

  constructor () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    this.sheet = ss.getSheetByName("Backend");
    this.teamDataSheet = ss.getSheetByName("Copy of Team_Data");
  }


  updateTeamDataSheet() {
    const [backendHeaders, backendData] = get_Data_Indices_From_Sheet(this.sheet, 1);
    const [teamHeaders, teamData] = get_Data_Indices_From_Sheet(this.teamDataSheet, 2);

    const backendNames = backendData.map(r => r[backendHeaders["Employee Identifier"]]).filter(Boolean);
    const teamNames = teamData.map(r => r[teamHeaders["Name of Tutor"]]).filter(Boolean);

    // Remove "(E10)" etc.
    const clean = name => String(name).replace(/\(.*?\)/g, "").trim();

    let matched = [];
    let onlyBackend = [];
    let onlyTeam = [...teamNames];

    backendNames.forEach(backName => {
      const backClean = clean(backName);

      const found = teamNames.find(teamName => clean(teamName) === backClean);

      if (found) {
        matched.push([backName, found]);
        onlyTeam = onlyTeam.filter(x => x !== found);
      } else {
        onlyBackend.push(backName);
      }
    });

    // SORT for readability
    matched.sort((a, b) => clean(a[0]).localeCompare(clean(b[0])));
    onlyBackend.sort((a, b) => clean(a).localeCompare(clean(b)));
    onlyTeam.sort((a, b) => clean(a).localeCompare(clean(b)));

    // OUTPUT
    console.log("========== MATCHED ==========");
    matched.forEach(pair => console.log(`${pair[0]}  ==  ${pair[1]}`));

    console.log("========== ONLY IN BACKEND ==========");
    onlyBackend.forEach(name => console.log(name));

    console.log("========== ONLY IN TEAM_DATA ==========");
    onlyTeam.forEach(name => console.log(name));

    let replaceMap = {};
    matched.forEach(([backendName, teamName]) => {
      replaceMap[clean(teamName)] = backendName;
    });

    let updatedCount = 0;

    // Replace names in teamData
    for (let i = 0; i < teamData.length; i++) {
      let currentName = teamData[i][teamHeaders["Name of Tutor"]];
      if (!currentName) continue;

      let cleaned = clean(currentName);

      if (replaceMap[cleaned]) {
        // Replace with backend full name
        teamData[i][teamHeaders["Name of Tutor"]] = replaceMap[cleaned];
        updatedCount++;
      }
    }

    // Write back to sheet if updates happened
    if (updatedCount > 0) {
      this.teamDataSheet
        .getRange(3, 1, teamData.length, teamData[0].length)
        .setValues(teamData);

      console.log("Names replaced:", updatedCount);
    } else {
      console.log("No replacements needed.");
    }

  }

  updateBackendSheet(employeeDetails) {
    const [ headers, data ] = CentralLibrary.get_Data_Indices_From_Sheet(this.sheet);
    const uuids = new Set( data.map(row => row[headers["UUID"]]) );
    const recAdd = data.map(row => row[headers["Recommended Addtions"]]).filter(Boolean);
    const recDel = data.map(row => row[headers["Recommended Deletions"]]).filter(Boolean);

    const updatedEmpList = []
    data.forEach(row => {
      const empIdentifier = row[headers["UUID"]];
      if ( !employeeDetails[ empIdentifier ] ) {
        this.sheet.getRange( recDel.length+2, headers["Recommended Deletions"]+1 ).setValue(empIdentifier);
        recDel.push(empIdentifier);
      } else {
        updatedEmpList.push([empIdentifier, employeeDetails[empIdentifier][0]]);
      }
    });
    console.log(updatedEmpList)
    this.sheet.getRange(2, 1, this.sheet.getLastRow(), 2).clearContent();
    this.sheet.getRange(2, 1, Object.keys(updatedEmpList).length, 2).setValues(updatedEmpList);

    for ( const [ key, [ empIdentifier, doj ] ] of Object.entries(employeeDetails) ) {
      if ( !uuids.has(key) ) {
        const lastRow = this.sheet.getLastRow();
        this.sheet.getRange(lastRow+1, 1, 1, 2).setValues([[key, empIdentifier]]);
        this.sheet.getRange( recAdd.length+2, headers["Recommended Addtions"]+1 ).setValue(empIdentifier);
        this.sheet.getRange( recAdd.length+2, headers["Date of Joining"]+1 ).setValue(Utilities.formatDate(doj, "IST", "dd-MMM-yyyy")).setNumberFormat("dd-MMM-yyyy");
        recAdd.push(empIdentifier);
      }
    }
  }
}


function get_Data_Indices_From_Sheet(sheet, headerRow = 1) {
  const all = sheet.getDataRange().getValues();

  // headerRow is 1-based → Convert to 0-based index
  const headerIndex = headerRow - 1;

  const headers = all[headerIndex];
  const data = all.slice(headerIndex + 1);

  return [createIndexMap(headers), data];
}

function createIndexMap(headers) {
  return headers.reduce((map, val, index) => {
    const key = String(val).trim();   // convert to string safely
    if (key !== "")
      map[key] = index;
    return map;
  }, {});
}




class EmployeeDetails {

  static getEmployeeData () {
    const sheet = SpreadsheetApp.openById("11b5A_ZDi9DskTiKJ0j4fBDkRbY-0I2h7rVRxr6KVK1o").getSheetByName("Employee Info");
    const [headers, data] = CentralLibrary.get_Data_Indices_From_Sheet(sheet);
    const filteredData = data.filter( row => DEPARTMENTS.includes(row[headers["Department"]]) && row[headers["Status"]] === "Active" );
    return Object.fromEntries(filteredData.map(row => [ row[headers["Unique ID"]], [ row[headers["Employee Identifier"]], row[headers["DOJ"]] ] ]));
  }
}










