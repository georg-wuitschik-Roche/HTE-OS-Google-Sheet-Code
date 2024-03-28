/*jshint sub:true*/

/**
 * Hotelplanner: Produces a Quantos input file that resets the remaining dose count to 999 for all heads that have less than 200 dosings left. 
 * If there are less than two heads present on the hotel with less than 200 dosings left, the two heads with the lowest residual number of dosings will be selected.  
 */
function createDosingCountResetList() {
  const currentTime = new Date();
  var hotelPlannerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hotelplanner");
  var hotelPlannerRange = hotelPlannerSheet.getRange(2, 10, 32, 7).getValues();
  const quantosFolder = DriveApp.getFolderById(globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["QUANTOSfOLDERiD"]);

  //quantosHeads contains information about solid starting materials that will be dosed using Quantos Chronect
  var quantosHeads = [[
    "",
    "Analysis Method",
    "Dosing Head Tray",
    "Dosing Head Pos.",
    "Substance",
    "Lot ID",
    "Filled Quantity [mg]",
    "Expiration Date",
    "Retest Date",
    "Dose Limit",
    "Tap Before Dosing?",
    "Intensity [%]",
    "Duration [s]"]]; //later filled with dosing head information to write to Quantos heads
  var dosingCount = 0;
  hotelPlannerRange.sort(function (a, b) { // the file that was modified the latest ends up on top.
    return b[3] - a[3];
  });

  for (let row = 0; row < hotelPlannerRange.length; row++) {
    if (hotelPlannerRange[row][1] == "??") {
      hotelPlannerRange.splice(row, 1);
      row--;
    }
  }

  for (let row = 0; row < hotelPlannerRange.length; row++) {
    if (hotelPlannerRange[row][3] > 200 && hotelPlannerRange.length > 2) {
      hotelPlannerRange.splice(row, 1);
      row--;
    }
  }
  console.log(hotelPlannerRange);
  for (let iteration = 1; iteration < 3; iteration++) {//the count first needs to be set to 0 and afterwards to 999
    for (let row = 0; row < hotelPlannerRange.length; row++) {

      quantosHeads.push([
        quantosHeads.length,
        "C:\\Users\\Public\\Documents\\Chronos\\Methods\\HeadWrite in Sequence.cam",
        "Heads",
        hotelPlannerRange[row][0], //Position on the hotel
        hotelPlannerRange[row][1], //Head ID
        hotelPlannerRange[row][6], //Abbreviated Component Name
        hotelPlannerRange[row][5], //Filled Quantity 
        "",
        "",
        dosingCount,   // dosing limit, i.e how many times the head can be dosed, 999 is the maximum
        "True", // Tap before dosing
        40,     // Tapping intensity
        2]);

    }
    dosingCount = 999; //after setting the heads to 0, now run again through all the heads and set them to 999
  }
  console.log(quantosHeads);
  if (quantosHeads.length > 1) {  // only create a file if there's more data present than the header 
    var quantosHeadsXML = createQuantosXml(quantosHeads);  // converts the quantosHeads array into a Quantos digestible XML-string  
    var quantosHeadsFile = quantosFolder.createFile('Head reset ' + currentTime + ".csl", quantosHeadsXML);     // contains all solids for writing dosing heads
  }

}




/**
 * Hotelplanner: This function is used to deselect all checkboxes on the left side of the hotel planner
 */
function uncheckAllCheckboxes() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hotelplanner").getRange(2, 8, 32, 1).uncheck();
}

/**
 * Hotelplanner: creates the file containing the instructions for Quantos which heads to take off the hotel and drop (into a box in front of the hotel at the right)
 */
function createHeadDropOffFile() {

  // connect to sheet and get relevant data from Quantos Hotel table
  // lots of help taken from: https://spreadsheet.dev/working-with-checkboxes-in-google-sheets-using-google-apps-script
  var date = Utilities.formatDate(new Date(), "CET", 'MMM_dd'); // used for file creation (date input)

  var hotelPlannerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hotelplanner");
  var hotelPlannerRange = hotelPlannerSheet.getRange(2, 9, 32, 2).getValues();
  var arr = []; // array to be populated with input data for Quantos
  var quantosXmlArray = [["", "Analysis Method", "Device", "Dosing Head Tray", "Dosing Head Pos."]];
  // console.log(arr);
  // console.log(hotelPlannerRange)

  //loop over the data and populate arr with the heads to be removed
  for (var row = 0; row < hotelPlannerRange.length; row++) {
    if (hotelPlannerRange[row][0] == "x") {
      arr.push(row + 1);
      //console.log(arr);
    }
  }

  // add init-robot line to arr at the beginning

  quantosXmlArray.push(
    [1,
      "C:\\Users\\Public\\Documents\\Chronos\\Methods\\Utilities\\Init Robot.cam",
      "Quantos",
      "",
      ""
    ]
  );

  // iterate through all elements of arr

  for (var item = 0; item < arr.length; item++) {

    quantosXmlArray.push(
      [item + 2,
        "C:\\Users\\Public\\Documents\\Chronos\\Methods\\Utilities\\Head From Rack To DropOff.cam",
        "Quantos",
        "Heads",
      arr[item]
      ]
    );
  }

  // add init-robot line to arr at the end

  quantosXmlArray.push(
    [item + 2,
      "C:\\Users\\Public\\Documents\\Chronos\\Methods\\Utilities\\Init Robot.cam",
      "Quantos",
      "",
      ""
    ]
  );
  // console.log(quantosXmlArray);

  // csl/xml file related:

  var quantosFolder = DriveApp.getFolderById(globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["QUANTOSfOLDERiD"]);
  var quantosXML = createQuantosXml(quantosXmlArray);
  var quantosXmlFile = quantosFolder.createFile('HeadDropOff Selection ' + date + ".xml", quantosXML); // good for debugging but not needed
  var quantosCslFile = quantosFolder.createFile('HeadDropOff Selection ' + date + ".csl", quantosXML);

}

/**
 * Hotelplanner: returns an array with the number of vials present on the array of plateIDs that is handed over. 
 * @param {Array} last10PlateIds Line from the PlateIngredients Sheet
 *  * @param {Array} checkBoxValue dummyParameter to force re-calculation of the formula on change
 * @return {number} Formatted Plate Ingredient as String depending on the type of compound.
 * @customfunction
 */
function getNumberOfVialsOnPlateArray(last10PlateIds, checkBoxValue) {

  // setup the variables
  var elnId = "";
  var plateNumber = 0;

  //var sqlString = "select STRING_AGG(concat(ELN_ID,'_', PLATENUMBER, '$$', COUNT(*)), '£') WITHIN GROUP (group BY ELN_ID, PLATENUMBER) FROM (SELECT DISTINCT ELN_ID, PLATENUMBER, Coordinate FROM wells_prod where  ActualVolume is null and (";


  var sqlString = "select concat(ELN_ID,'_', PLATENUMBER, '$$', COUNT(*))  FROM (SELECT DISTINCT ELN_ID, PLATENUMBER, Coordinate FROM wells_prod where ";
  for (var row = 0; row < last10PlateIds.length; row++) {
    [elnId, plateNumber] = last10PlateIds[row][0].split("_");
    sqlString += "(ELN_ID = '" + elnId + "' and PLATENUMBER = " + plateNumber + ")";
    if (row < last10PlateIds.length - 1) sqlString += " or ";
  }
  sqlString += " ) AS internalQuery group by ELN_ID, PLATENUMBER";

  // connect to the database
  const idOfCurrentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet().getId();
  var connector = new mssql_jdbc_api(   // connect to the database
    globalVariableDict[idOfCurrentSpreadsheet]["DBsERVERiP"],
    globalVariableDict[idOfCurrentSpreadsheet]["DBpORT"],
    globalVariableDict[idOfCurrentSpreadsheet]["DBnAME"],
    globalVariableDict[idOfCurrentSpreadsheet]["DBuSERNAME"],
    globalVariableDict[idOfCurrentSpreadsheet]["DBpASSWORD"]);

  // send the queries

  connector.executeQuery(sqlString);
  const sqlResultsTable = connector.getResultsAsArray();

  for (row = 0; row < sqlResultsTable.length; row++) {
    sqlResultsTable[row] = sqlResultsTable[row][0].split("$$");
  }

  var returnTable = [];
  var plateIdFoundFlag = "no";
  for (row = 0; row < last10PlateIds.length; row++) {
    plateIdFoundFlag = "no";
    for (var sqlRow = 0; sqlRow < sqlResultsTable.length; sqlRow++) {
      if (last10PlateIds[row] == sqlResultsTable[sqlRow][0]) {
        returnTable.push(sqlResultsTable[sqlRow][1]);
        plateIdFoundFlag = "yes";
      }
    }
    if (plateIdFoundFlag == "no") returnTable.push("Plate ID not found");
  }
  return returnTable;
}

/**
 * Hotelplanner: returns an array with the number of undosed solids not dosed so far for the plateIDs handed over. 
 * @param {Array} last10PlateIds array of PlateIds to be checked
 * @param {Array} checkBoxValue dummyParameter to force re-calculation of the formula on change
 * @return {Array} Array of PlateID, number of missing dosings, Component ID and BatchID
 * @customfunction
 */
function getUndosedSolids(last10PlateIds, checkBoxValue) {

  // setup the variables
  var elnId = "";
  var plateNumber = 0;

  // select STRING_AGG(concat(ELN_ID,'_', PLATENUMBER, '$$', COUNT(*)), '£') WITHIN GROUP (group BY ELN_ID, PLATENUMBER) FROM (SELECT DISTINCT ELN_ID, PLATENUMBER, Coordinate FROM wells_prod where  ActualVolume is null and (";

  var sqlString = "SELECT concat(ELN_ID,'_', PLATENUMBER,'$$', COUNT(*) , '$$', Component_ID,'$$', Batch_ID) FROM wells_prod where ActualVolume is null and DosingTimestamp is null and (";
  for (var row = 0; row < last10PlateIds.length; row++) { // generate the part of the sql String that specifies which plates to look for
    [elnId, plateNumber] = last10PlateIds[row][0].split("_");
    sqlString += "(ELN_ID = '" + elnId + "' and PLATENUMBER = " + plateNumber + ")";
    if (row < last10PlateIds.length - 1) sqlString += " or ";
  }
  sqlString += " ) group by ELN_ID, PLATENUMBER, Component_ID, Batch_ID order by ELN_ID, PLATENUMBER";


  // connect to the database
  const idOfCurrentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet().getId();
  var connector = new mssql_jdbc_api(   // connect to the database
    globalVariableDict[idOfCurrentSpreadsheet]["DBsERVERiP"],
    globalVariableDict[idOfCurrentSpreadsheet]["DBpORT"],
    globalVariableDict[idOfCurrentSpreadsheet]["DBnAME"],
    globalVariableDict[idOfCurrentSpreadsheet]["DBuSERNAME"],
    globalVariableDict[idOfCurrentSpreadsheet]["DBpASSWORD"]);

  // send the queries

  connector.executeQuery(sqlString);
  const sqlResultsTable = connector.getResultsAsArray();


  for (row = 0; row < sqlResultsTable.length; row++) {  //unpack the string
    sqlResultsTable[row] = sqlResultsTable[row][0].split("$$");
  }


  return sqlResultsTable;
}

