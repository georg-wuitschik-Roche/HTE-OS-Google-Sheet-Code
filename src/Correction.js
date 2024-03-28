/*jshint sub:true*/
/**
 * Correction Sheet: Moves the selected dosing file in the Correction Sheet to the list of "read-in" dosing files. 
 * 
 */

function moveDosingFile() {
  var col = 23; // refers to Col W in "DropdownTables"
  var correction = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Correction");
  var ddTables = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DropdownTables");
  var fileID = correction.getRange(2, 1).getValue(); // get the value in the Corrections Dropdown menu (the currently active one)
  var columnToCheck = ddTables.getRange("W:W").getValues();
  var lastRow = getLastRowSpecial(columnToCheck);
  ddTables.getRange(lastRow + 1, col).setValue(fileID); //schreib den Wert (Dosing sequence) ans Ende der Tabelle - check!
}

/**
 * Correction Sheet: Moves the selected dosing file in the Correction Sheet back to the list of "fresh" dosing files by removing it from the list of already read-in files in the sheets Dropdown Tables.
 * Triggered by the "Return the Dosing File Button" in the Correction Sheet
 */
function returnDosingFile() {
  var ddTables = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DropdownTables");
  var correction = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Correction");
  var col = 23; // refers to Col W in "DropdownTables" 
  var input = correction.getRange(24, 1).getValue(); // get the value in the 2nd Corrections Dropdown menu ("already written Dosing File")
  var columnToCheck = ddTables.getRange("W:W").getValues(); // get all the dosing files that have been written already

  for (var line = 0; line < columnToCheck.length; line++) {
    if (columnToCheck[line][0] === input) {
      columnToCheck.splice(line, 1);
      // console.log(columnToCheck[line][0]); used to find the infinite loop
      line--;
    }
  }

  ddTables.getRange(1, col, columnToCheck.length, 1).setValues(columnToCheck);
}

/**
 * Correction Sheet: read the Excel-file specified in A2 and write it to columns T to AI, triggered by the button "Load File"
 */
function readDosingFile() {
  var fileData;
  var correctionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Correction");
  var fileID = correctionSheet.getRange(50, 1).getValue();   // contains the file id of the Excel file specified in A2
  var taraFails = []; // contains all the rows for which the taring of the balance failed and thus contain no material although the file records a weight of > 30 g. 
  var taraFailsString = "";
  SpreadsheetApp.getActiveSpreadsheet().toast('Retrieving File', 'Status', 3);
  try {
    fileData = getAndAppendData(fileID);    //read data from the Excel-file specified
  } catch (err) {
    Browser.msgBox('Error ' + err + ' triggered when opening the Excel-file. This can happen if the Excelfile is corrupted because Quantos crashed during execution. In this case it may help to open the file in Excel, repair it, save a copy into the respective dosing results folder, scan the folder and try to read the repaired file. Also, converting the csv-file in the underlying temp-folder into an Excel file and opening that may help and may contain all the dosing information.');
    return;
  }


  var fileContent = fileData[0];

  // Go through the content of the file and correct instances where taring of the balance didn't happen resulting in a fill weight of 30+ g being recorded when in fact nothing ended up in the vial.

  for (var row = 1; row < fileContent.length; row++) { // the first row contains the column headers, thus start from 1

    if (fileContent[row][8] > 30000 && fileContent[row][10] > 100) {  // actual weight needs to be > 30 g and the deviation needs to be > 100%, thus even if the target weight was > 30 g it would not trigger this if, since a deviation that great doesn't happen at this scale.
      taraFails.push(fileContent[row]); // for alerting the user later
      taraFailsString += "Vial: " + fileContent[row][1] + ", Substance: " + fileContent[row][4] + ",\\n";

      fileContent[row][14] = fileContent[row][14] + ", Tara fail, old values: " + fileContent[row][8] + " mg, deviation: " + fileContent[row][10] + "%";
      fileContent[row][8] = 0;
      fileContent[row][10] = -100;

    }

    if (fileContent[row][14].includes("does not hold")) {  //  if a head runs dry during dosing because overdosing happened during the sequence, the the system erroneously counts the dosing as valid when it should be -100%

      fileContent[row][8] = 0;
      fileContent[row][10] = -100;

    }

  }



  var excelFileTitle = fileData[1];
  var plateInformation = String(fileContent[1][2]).split(" - ");   //Array which contains the plate ID as first element, "Initial" or "Correction" as second and, if the second is "Correction" the vialoption as third element. Only for Correction files cases can arise in which the formulas in the Correction sheet can't conclusively determine the plate type

  if (plateInformation.length > 2 && (plateInformation[1] == "Initial" || plateInformation[1] == "Correction")) { //unlikely that a split using two space-characters will result in an array of length 3
    correctionSheet.getRange(7, 1).setValue(plateInformation[0]);
    correctionSheet.getRange(19, 1).setValue(plateInformation[1]);

    switch (plateInformation[2]) {                           // depending on the vial volume chosen, this distinguishes between 1 and 1.2 mL vials. 
      case "96 wells 1 mL gold":
        correctionSheet.getRange(11, 1).setValue("96 gold"); // set value for plate type
        correctionSheet.getRange(13, 1).setValue("1 mL"); // set value for vial volume
        break;
      case "96 wells 1.2 mL gold":
        correctionSheet.getRange(11, 1).setValue("96 gold"); // set value for plate type
        correctionSheet.getRange(13, 1).setValue("1.2 mL"); // set value for vial volume        
        break;
      case "24 wells 1 mL gold":
        correctionSheet.getRange(11, 1).setValue("24 gold"); // set value for plate type
        correctionSheet.getRange(13, 1).setValue("1 mL"); // set value for vial volume     
        break;
      case "24 wells 1.2 mL gold":
        correctionSheet.getRange(11, 1).setValue("24 gold"); // set value for plate type
        correctionSheet.getRange(13, 1).setValue("1.2 mL"); // set value for vial volume
        break;
      case "24 wells 4 mL gold":
        correctionSheet.getRange(11, 1).setValue("24 gold"); // set value for plate type
        correctionSheet.getRange(13, 1).setValue("4 mL"); // set value for vial volume
        break;
      case "24 wells 8 mL gold":
        correctionSheet.getRange(11, 1).setValue("24 gold"); // set value for plate type
        correctionSheet.getRange(13, 1).setValue("8 mL"); // set value for vial volume
        break;
      case "48 wells 2 mL gold":
        correctionSheet.getRange(11, 1).setValue("48 gold"); // set value for plate type
        correctionSheet.getRange(13, 1).setValue("2 mL"); // set value for vial volume

        break;
      default:
        correctionSheet.getRange(11, 1).setFormula("=A53"); // set formula for plate type
        correctionSheet.getRange(13, 1).setFormula("=A58"); // set formula for vial volume
        Browser.msgBox("No assignment to plate type and vial volume could be made based on data found in the sample ID column of " + excelFileTitle + ". Falling back to calculated values. Correct using the dropdowns, if neccessary.");
        break;
    }
    SpreadsheetApp.getActiveSpreadsheet().toast('Writing Data', 'Status');
    correctionSheet.getRange(1, 20, correctionSheet.getLastRow(), 16).clearContent();    // remove whatever was there previously in the right part of the sheet.
    correctionSheet.getRange(1, 20, fileContent.length, fileContent[0].length).setValues(fileContent);    //write data to right part of the sheet
    correctionSheet.getRange(2, 28, fileContent.length, 1).setNumberFormat('#,##0.00');   // sometimes the content of some of the cells would be interpreted as date which results in problems in Spotfire
    correctionSheet.getRange(2, 2, 999, 1).setValue("true");  //reset checkboxes in column B to "checked"
    SpreadsheetApp.flush();

    if (taraFailsString.length > 1) Browser.msgBox("For these dosings, the taring of the balance failed and the dosing result was corrected to reflect the vials not containing any material: \\n\\n" + taraFailsString);
    var response = Browser.msgBox("Do you want to write the actual weights to the Wells-Sheet?", Browser.Buttons.YES_NO);
    if (response == "yes") {
      writeActualWeightsToWellsTable();
    }


  } else {
    for (row = 0; row < fileContent.length; row++) {
      fileContent[row][2] = excelFileTitle;
    }

    // Set Values for Plate ID, Lower Boundary, Plate Type and Vial Volume

    correctionSheet.getRange(11, 1).setFormula("=A53"); // set formula for plate type
    correctionSheet.getRange(13, 1).setFormula("=A58"); // set formula for vial volume
    SpreadsheetApp.getActiveSpreadsheet().toast('Writing Data', 'Status');
    correctionSheet.getRange(1, 20, correctionSheet.getLastRow(), 16).clearContent();    // remove whatever was there previously in the right part of the sheet.
    correctionSheet.getRange(1, 20, fileContent.length, fileContent[0].length).setValues(fileContent);    //write data to right part of the sheet
    correctionSheet.getRange(2, 28, fileContent.length, 1).setNumberFormat('#,##0.00');   // sometimes the content of some of the cells would be interpreted as date which results in problems in Spotfire
    correctionSheet.getRange(2, 2, 999, 1).setValue("true");  //reset checkboxes in column B to "checked"
  }
}

/**
 * Correction Sheet: Accepts the file id of an Excel-file from gDrive, converts the Excel-file to a gSheet and returns its content as an array together with the original filename.
 * @param {String} fileId fileID of the EXcel file to be read in.
 * @return {Array} Array containing the content of the file and the filename of the original Excel file.
 */
function getAndAppendData(fileID) {
  var excelFile = DriveApp.getFileById(fileID);

  var blob = excelFile.getBlob();
  var excelFileTitle = excelFile.getName().replace(/.xlsx?/, "");
  var resource = { title: excelFileTitle };  // Modified
  var sourceSpreadsheet = Drive.Files.insert(resource, blob, { convert: true });  // Modified

  // Also I added below script.
  var sourceSheet = SpreadsheetApp.openById(sourceSpreadsheet.id).getSheets()[0];

  var fileContent = sourceSheet.getDataRange().getValues();
  //fileContent.splice(0, 1)
  Drive.Files.trash(sourceSpreadsheet.id);


  return [fileContent, excelFileTitle];

}

/**
 * Correction Sheet: go through the folder with the Quantos dosing results, collect the filenames and write them to the dropdown tables sheet
 */
function getAndListFilesInFolder() {
  var arr, f, file, folderName, subFolders, id, mainFolder, name, ID, sh, thisSubFolder, url;

  sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DropdownTables");

  arr = [];

  mainFolder = DriveApp.getFolderById(globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["QUANTOSdOSINGrESULTSiD"]);
  subFolders = mainFolder.getFolders();
  folderName = mainFolder.getName();

  SpreadsheetApp.getActiveSpreadsheet().toast('Scanning Folders', 'Status');
  while (subFolders.hasNext()) {
    thisSubFolder = subFolders.next();
    f = thisSubFolder.getFiles();
    folderName = thisSubFolder.getName();

    while (f.hasNext()) {
      file = f.next();
      name = file.getName();
      url = file.getUrl();
      ID = file.getId();

      if (name.substring(name.length - 4, name.length) == "xlsx") {
        arr.push([name, url, ID, folderName]);
      }
    }
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('Writing Data', 'Status');
  sh.getRange(2, 18, arr.length, arr[0].length).setValues(arr);
}

function writeQuantosCorrectionCslFile() {

  // connect to Correction sheet and get the data of the first 18 columns
  var correctionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Correction");
  var sheetData = correctionSheet.getRange(2, 1, correctionSheet.getLastRow(), 18).getValues();


  // assign some values in the first column to variables

  var nameDosingLogFile = sheetData[0][0];
  var plateID = sheetData[5][0];
  if (plateID == "") {
    Browser.msgBox("Please select a Plate ID from the Dropdown list in cell A7 and then try again.");
    return;
  }

  var plateType = sheetData[9][0];
  var vialVolume = sheetData[11][0];
  var dosingType = sheetData[19][0];
  var toleranceMode = "MinusPlus";

  var preDoseTapping = 6;      // length in seconds of the predose tapping
  var preDoseStrength = 40;           // strength in % of maximum tapping strength
  var useFrontDoor = "True";
  var useSideDoors = "False";


  var vialOption = "96 wells 1 mL gold";           // describes the type of vial used on Quantos, gold = analytical sales plates or green unchained labs, black would be another option which we don't use typically. 
  switch (plateType + " " + vialVolume) {                           // depending on the vial volume chosen, this distinguishes between 1 and 1.2 mL vials. 
    case "96 gold 1 mL":
      vialOption = "96 wells 1 mL gold";
      break;
    case "96 gold 1.2 mL":
      vialOption = "96 wells 1.2 mL gold";
      break;
    case "24 gold 1 mL":
      vialOption = "24 wells 1 mL gold";
      break;
    case "24 gold 1.2 mL":
      vialOption = "24 wells 1.2 mL gold";
      break;
    case "24 gold 4 mL":
      vialOption = "24 wells 4 mL gold";
      break;
    case "24 gold 8 mL":
      vialOption = "24 wells 8 mL gold";
      break;
    case "48 gold 2 mL":
      vialOption = "48 wells 2 mL gold";
      break;
    default:
      vialOption = "96 wells 1 mL gold";    //Even if a vial is selected that doesn't exist for 96-well plates, it reverts to 1.0 mL vials and warns the user
      Browser.msgBox("There are no " + data[27][0] + " mL vials for 96-well plates. Standard 1 mL vials are selected instead for the purpose of building the Quantos input file.");
      break;
  }
  var mass = 0;
  //var quantosXmlArray = [["", "Analysis Method", "Dosing Tray Type", "Tolerance Mode", "Substance Name", "Dosing Vial Tray", "Dosing Vial Pos.", "Dosing Vial Pos. [Axx]", "Amount [mg]", "Tolerance [%]", "Sample ID", "Comment"],
  // [1, "C:\\Users\\Public\\Documents\\Chronos\\Methods\\Set Config.cam", vialOption, toleranceMode, "", "", "", "", "", "", "", ""]]; //later filled with dosing information to generate the Quantos csl-file

  var quantosXmlArray = [["",
    "Analysis Method",
    "Dosing Tray Type",
    "PreDose Tap Duration [s]",
    "PreDose Tap Intensity [%]",
    "Tolerance Mode",
    "Device",
    "Use Front Door?",
    "Use Side Doors?",
    "Substance Name",
    "Dosing Vial Tray",
    "Dosing Vial Pos.",
    "Dosing Vial Pos. [Axx]",
    "Amount [mg]",
    "Tolerance [%]",
    "Sample ID",
    "Comment"],
  [1,
    "C:\\Users\\Public\\Documents\\Chronos\\Methods\\Set Config.cam",
    vialOption,
    preDoseTapping,
    preDoseStrength,
    toleranceMode,
    "Quantos",
    useFrontDoor,
    useSideDoors,
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    ""]];

  // go through the array and collect data for the CSL-file
  var correctionTray = sheetData[0][2];

  for (var row = 0; row < sheetData.length; row++) {
    if (sheetData[row][3] == "") { break; }
    // include rows where the box in column 2 is ticked
    if (sheetData[row][1] == true) {
      if (sheetData[row][10] < 0) { sheetData[row][10] = 0; } // in case a negative mass is recorded by the balance
      mass = parseFloat(sheetData[row][9] - sheetData[row][10]).toFixed(3);
      quantosXmlArray.push([quantosXmlArray.length, "C:\\Users\\Public\\Documents\\Chronos\\Methods\\Dosing Method.cam", "", "", "", "", "", "", "", sheetData[row][6], correctionTray, "", sheetData[row][3].split(' - ')[1], mass, parseFloat(7 / (parseFloat(mass) + 0.03) + 3).toFixed(0), plateID + " - Correction - " + vialOption, sheetData[row][8]]); // will be used to generate the csl input file
    }
  }

  // generate csl-file content, use it to write file and place it in the "Quantos Correction Dosings" folder in "Project Data" id 1dbqcbJcuVRaEtwAjejkPw4YBLn7GH6KO

  var folder = DriveApp.getFolderById(globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["QUANTOScORRECTIONdOSINGiD"]);

  var quantosXML = createQuantosXml(quantosXmlArray);
  //  var quantosXmlFile = folder.createFile('Quantos xml ' + plateID + "_Corr_based_on_" + nameDosingLogFile + ".xml", quantosXML)     // only used for debugging purposes
  var quantosCslFile = folder.createFile(plateID + "_Corr_based_on_" + nameDosingLogFile + ".csl", quantosXML);

}

/**
 * Correction Sheet: new version of the function that writes the actual weights of initial dosings and corrections to the wells table 
 */
function writeActualWeightsToWellsTable() {
  // connect to sheets
  var correctionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Correction");
  //var wellsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Wells");
  var dropdownTablesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DropdownTables");

  // read dosing data and wells data
  SpreadsheetApp.getActiveSpreadsheet().toast('Reading Data', 'Status', 20);
  var correctionSheetLastRow = correctionSheet.getLastRow();
  var correctionSheetData = correctionSheet.getRange(2, 20, correctionSheetLastRow, 20).getValues();
  //var wellsSheetData = wellsSheet.getDataRange().getValues()
  var dosingType = correctionSheet.getRange(19, 1).getValue();
  var plateID = correctionSheet.getRange(7, 1).getValue();
  if (plateID == "") {
    Browser.msgBox("Please select a Plate ID from the Dropdown list in cell A7 and then try again.");
    return;
  }

  var elnId = "";
  var plateNumber = 1;
  var notebookNumber = "";
  var experimentNumber = 1;
  var tableToWriteTo = "";
  const idOfCurrentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet().getId();

  [elnId, plateNumber] = plateID.split('_');
  [notebookNumber, experimentNumber] = elnId.split("-");   // needed to know which table to write to

  if (notebookNumber == "ELN046807") {
    tableToWriteTo = "solubility_wells";
  } else { // will be ELNL032036 in most cases for the most cases
    tableToWriteTo = globalVariableDict[idOfCurrentSpreadsheet]["WELLStABLEnAME"];
  }


  //Read list of files where the actual weights have already been added to the Wells sheet and add the current file name to it:  
  var columnToCheck = dropdownTablesSheet.getRange("W:W").getValues();
  var lastRow = getLastRowSpecial(columnToCheck);
  var fileName = correctionSheet.getRange(2, 1).getValue();

  for (var row = 0; row < lastRow + 1; row++) {   //Check if the current fileName is already in the list and abort if it is - prevents writing the same file twice to the Wells-Sheet
    if (columnToCheck[row] == fileName) {
      Browser.msgBox("The actual weights from " + fileName + " have already been written to the Wells-Sheet. The script is aborted.");
      return;
    }
  }

  var connector = new mssql_jdbc_api(   // connect to the database
    globalVariableDict[idOfCurrentSpreadsheet]["DBsERVERiP"],
    globalVariableDict[idOfCurrentSpreadsheet]["DBpORT"],
    globalVariableDict[idOfCurrentSpreadsheet]["DBnAME"],
    globalVariableDict[idOfCurrentSpreadsheet]["DBuSERNAME"],
    globalVariableDict[idOfCurrentSpreadsheet]["DBpASSWORD"]);

  connector.executeQuery("SELECT LEN(STRING_AGG(DosingTimestamp, '')) AS dosings  FROM " + tableToWriteTo + " where ELN_ID = '" + elnId + "' and PLATENUMBER = " + plateNumber);
  //adds up the number of characters in the dosing timestamps for this plate and will be a large number for plates that already contain dosings and 0, if no dosings have been written
  const lengthOfCumulatedDosingString = connector.getResultsAsArray()[0][0];
  const initialOrCorrectionFlag = correctionSheetData[0][2].split(" - ")[1]; //theoretically initial and corrections could be mixed in a sequence, but this shouldn't happen and thus only the first row is being looked at.

  if (lengthOfCumulatedDosingString > 0 && initialOrCorrectionFlag == "Initial") { // only entered if there is already dosing data present in the database and if it's an initial dosing

    var dosingCounts = {};

    for (var row = 0; row < correctionSheetData.length; row++) { // create a dictionary of all dosinghead IDs as keys and the number of dosings present in the file as values
      if (correctionSheetData[row][4] == "") break;
      if (correctionSheetData[row][4] in dosingCounts) {
        dosingCounts[correctionSheetData[row][4]]++;
      } else {
        dosingCounts[correctionSheetData[row][4]] = 1;
      }
    }

    connector.executeQuery("select Component_ID, Batch_ID,  count(Batch_ID) as missing_dosings  FROM " + tableToWriteTo + " where ELN_ID = '" + elnId + "' and PLATENUMBER = " + plateNumber + " and ActualVolume is null and DosingTimestamp is null group by Component_ID, Batch_ID");
    //counts the number of lines of solids for which no dosing data was recorded for each combination of Component ID and Batch ID.
    const tableOfMissingDosings = connector.getResultsAsArray();

    //Go through the dosingCounts dictionary and check for each entry whether the number of dosings present in the dosing file for the compound in question is smaller or equal than the number
    //of entries for this component / Batch ID combination in the database for which no dosing has been recorded so far. 
    var dosingHeadIdFound = "no";
    for (var dosingHeadId in dosingCounts) {
      dosingHeadIdFound = "no";
      for (var row = 0; row < tableOfMissingDosings.length; row++) {
        if (dosingHeadId == tableOfMissingDosings[row][0] + "@" + String(tableOfMissingDosings[row][1].toString().replace(',', ':')).substring(String(tableOfMissingDosings[row][1]).length - 14)) {
          dosingHeadIdFound = "yes";
          if (dosingCounts[dosingHeadId] > tableOfMissingDosings[row][2]) {
            Browser.msgBox("There are " + dosingCounts[dosingHeadId] + " dosings of " + dosingHeadId + " present in the dosing file, but only " + tableOfMissingDosings[row][2] + " dosings are missing in the database. Aborting script now.");
            return;
          }
        }
      }
      if (dosingHeadIdFound == "no") {
        var response = Browser.msgBox("There are " + dosingCounts[dosingHeadId] + " dosings of " + dosingHeadId + " present in the dosing file, but no dosings are missing in the database for this compound. This may happen, if the batch ID dosed is not identical to the one planned in FileGenerator. In this case, click 'Yes', otherwise consider aborting by clicking 'No'. Do you want to continue?",
          Browser.Buttons.YES_NO);
        if (response == Browser.Buttons.NO) {
          return;
        } // otherwise continue
      }
    }
  }


  if ((lengthOfCumulatedDosingString == 0 || !lengthOfCumulatedDosingString) && initialOrCorrectionFlag == "Correction") {
    Browser.msgBox("You're trying to read correction dosings results in, but there are no initial dosings present. Make sure you read in all initial dosings before reading in correction dosings. ABORTING NOW...");
    return;
  }

  //Go through the different dosings and convert them into the dictionary that can be fed to the database
  var truncatedBatchId = "";
  var queryString = "";
  SpreadsheetApp.getActiveSpreadsheet().toast('Writing Data', 'Status', 20);

  const start = new Date();
  for (var row = 0; row < correctionSheetData.length; row++) {
    if (correctionSheetData[row][0] == "") break;
    [componentId, truncatedBatchId] = correctionSheetData[row][4].split('@');
    plateCoordinate = correctionSheetData[row][17];
    mass = parseFloat(correctionSheetData[row][8]);
    if (mass < 0) mass = 0.00001; //negative values are corrected and the 0.00001 signals that the actual value in the dosing file was <0 ( maybe useful for data analysis at some point)
    mass = parseFloat(mass + 0.0001);
    dosingString = correctionSheetData[row][14] + "_" + correctionSheetData[row][15];

    if (initialOrCorrectionFlag == "Initial") { //Overwrite the current mass and dosing string, if it's an initial dosing.
      queryString += "UPDATE " + tableToWriteTo + " set ActualMass = " + mass +
        ", DosingTimestamp = '" + dosingString.replaceAll("'", "") +
        "' where ELN_ID ='" + elnId +
        "' and PLATENUMBER = " + plateNumber +
        " and Component_ID = " + componentId +
        " and Coordinate ='" + plateCoordinate + "'; ";

    } else if (initialOrCorrectionFlag == "Correction") {

      queryString += "UPDATE " + tableToWriteTo + " set ActualMass = ActualMass + " + mass +
        ", DosingTimestamp =  CONCAT(ISNULL(DosingTimestamp,''), '" + "; " + dosingString.replaceAll("'", "") + "')" +
        " where ELN_ID ='" + elnId +
        "' and PLATENUMBER = " + plateNumber +
        " and Component_ID = " + componentId +
        " and Coordinate ='" + plateCoordinate + "'";
    } else {
      console.error("No clear assignment of Dosing status to initial or correction possible for " + correctionSheetData[row]);
    }

    if (row % 50 == 0 && row > 0) { // If too many rows are written at once, the SQL-string becomes too long and the server will come back with an error. Thus, not more than 100 rows are written at once. 
      connector.execute(queryString);  //This is ~10 times faster than writing individual lines, but requires that every single write operation is successful.
      queryString = "";
    }
  }
  if (queryString.length > 0) { // 99% of the time, there will be a residual string that needs to be written.
    connector.execute(queryString);  //This is ~10 times faster than writing individual lines, but requires that every single write operation is successful.
  }
  const end = new Date();


  dropdownTablesSheet.getRange(lastRow + 1, 23).setValue(fileName);
  SpreadsheetApp.getActiveSpreadsheet().toast('Finished, Time elapsed: ' + (end - start) + ' ms for ' + row + ' rows.', 'Status');

}

/************************************************************************
* taken from https://yagisanatode.com/2019/05/11/google-apps-script-get-the-last-row-of-a-data-range-when-other-columns-have-content-like-hidden-formulas-and-check-boxes/
* Gets the last row number based on a selected column range values
*
* @param {array} range : takes a 2d array of a single column's values
*
* @returns {number} : the last row number with a value. 
*
*/

function getLastRowSpecial(range) {
  var rowNum = 0;
  var blank = false;
  for (var row = 0; row < range.length; row++) {

    if (range[row][0] === "" && !blank) {
      rowNum = row;
      blank = true;
    } else if (range[row][0] !== "") {
      blank = false;
    }
  }
  return rowNum;
}
