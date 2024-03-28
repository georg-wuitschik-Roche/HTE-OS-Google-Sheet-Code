/*jshint sub:true*/
/**
 * FileGenerator: Creates the Quantos Inputfile in the Excel-xml format by taking the data found in wellsDictionary and writing it into a temporary gSheet and exporting it as xlsx. 
 * @param {Object} wellsDictionary dictionary containing all the dosing information.
 * @param {String} folderID id of the folder in which the Excelfile is to be created ( globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["CHEATSHEETSfOLDERiD"] ).
 * @param {String} fileName name of the new Excel file and temporary gSheet.
 * @param {Number} rowsOnPlate number of rows on the plate.
 * @param {Number} columnsOnPlate number of columns on the plate. 
 */
function generateQuantosXlsx(wellsDictionary, folderID = globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["CHEATSHEETSfOLDERiD"], fileName = "", rowsOnPlate = 0, columnsOnPlate = 0) {

  //These variables could be moved into a config sheet as to make them more easily accessible for modification.
  var trayType = "Gold96_8x30";  // true for the 96-well plate
  var numberOfWells = 96;
  var columnShift = 0; // While the coordinates used in the gSheet for a plate with 6 columns and more than 4 rows are the ones of a virtual 6x8 plate because the Junior 6-tip cant dose in the middle of a plate, for Quantos we simply use the 96-well template and shift the coordinates 3 columns to the right. 
  if (rowsOnPlate < 5 && columnsOnPlate < 7) { // If there are less than 5 rows and less than 7 columns, then the 24-well insert is used
    trayType = "Gold24_8x30";
    numberOfWells = 24;
  } else if (columnsOnPlate < 7) { // more than 4 rows and less than 6 columns means that a 6 column, 8 row-plate is used that is equivalent to a 96-well plate in which only the middle 6 columns are used (needed because Lea studio cannot handle dosing with the 6-tip in the middle of a plate)
    columnShift = 3;
    columnsOnPlate = 12;           // Junior thinks there are 6 columns on the plate, but in reality it is a 96-well plate which this rectifies,
  }
  var outOfToleranceAction = "Dose remaining substances";
  var algorithm = "Advanced";
  var tappingBeforeDosing = "off";
  var negativeTolerance = -5;
  var positiveTolerance = 5;
  var plateLocation = "Vials1";

  var data = [[trayType], [''], ['Out of tolerance action'], [outOfToleranceAction], [''], ['Algorithm'], ['Tapping before dosing'], [''], ['Negative tolerance [%]'], ['Positive tolerance [%]'], [''], [''], ['Position']];      //array that is later written into the temporary gSheet and which contains all the data needed for Quantos
  var comparator = [[trayType], [''], ['Out of tolerance action'], [outOfToleranceAction], [''], ['Algorithm'], ['Tapping before dosing'], [''], ['Negative tolerance [%]'], ['Positive tolerance [%]'], [''], [''], ['Position']];      //array that is later written into the temporary gSheet and which contains all the data needed for Quantos

  var rowToLetterAssignment = { 1: "A", 2: "B", 3: "C", 4: "D", 5: "E", 6: "F", 7: "G", 8: "H" };
  var coordinateFoundFlag = "no";

  // Generate a new Google Sheet and place it somewhere


  var fileID = createSpreadsheet(folderID, fileName);
  // Create a new sheet, name it and connect to it. 
  var spreadSheet = SpreadsheetApp.openById(fileID).getActiveSheet();    //.getActiveSpreadsheet().getSheetByName("Sheet 1");

  // Write the coordinates into col A 

  for (var plateRow = 1; plateRow < rowsOnPlate + 1; plateRow++) {
    for (var plateColumn = 1; plateColumn < columnsOnPlate + 1; plateColumn++) {
      data.push([rowToLetterAssignment[plateRow] + plateColumn]);
      comparator.push([rowToLetterAssignment[plateRow] + (plateColumn - columnShift)]);
    }
  }

  // Go through the different components and fill column B onward

  // wellsDictionary[key][0] = Header Information
  //            wellsDictionary[data[row][3]+data[row][4]] = [[data[row][3], data[row][4], data[row][16], data[row][18],data[row][19],data[row][20] ],[]];
  //                                                             component name   limit/level       comp name       limit/level    dose as        concentr      unit        solvent

  // wellsDictionary[key][1] = Dosing Information   [row, column, coordinate, volume/mass]
  // wellsDictionary[key][2] = Dosing Boundaries    [firstRow, firstColumn, firstCoordinate,lastRow, lastColumn, lastCoordinate]

  for (var key in wellsDictionary) {

    if (wellsDictionary[key][0][2] != "Solid") { continue; } //Only solids should be considered. 
    //The first 13 rows of each column contain header information
    for (var line = 0; line < 13; line++) {
      switch (line) {
        case 0:
          data[line].push(trayType);
          comparator[line].push("");
          break;
        case 1:
          data[line].push(plateLocation);
          comparator[line].push("");
          break;
        case 2:
          data[line].push('');
          comparator[line].push("");
          break;
        case 3:
          data[line].push('');
          comparator[line].push("");
          break;
        case 4:
          data[line].push('');
          comparator[line].push("");
          break;
        case 5:
          data[line].push(algorithm);
          comparator[line].push("");
          break;
        case 6:
          data[line].push(tappingBeforeDosing);
          comparator[line].push("");
          break;
        case 7:
          data[line].push('');
          comparator[line].push("");
          break;
        case 8:
          data[line].push(negativeTolerance);
          comparator[line].push("");
          break;
        case 9:
          data[line].push(positiveTolerance);
          comparator[line].push("");
          break;
        case 10:
          data[line].push('');
          comparator[line].push("");
          break;
        case 11:
          data[line].push("Substance [mg]");
          comparator[line].push("");
          break;
        case 12:
          data[line].push(wellsDictionary[key][0][7]);
          comparator[line].push("");
          break;
      }
    }
    // Check for each coordinate whether it's present in the wellsDictionary and if so write the corresponding weight to the data array. If it's not present, put an empty string in the slot. 
    for (line = 13; line < 13 + numberOfWells; line++) {
      coordinateFoundFlag = "no";
      for (var well = 0; well < wellsDictionary[key][1].length; well++) {
        if (wellsDictionary[key][1][well][2] == comparator[line][0]) {
          coordinateFoundFlag = "yes";
          data[line].push(wellsDictionary[key][1][well][3]);
          break;
        }
      }
      if (coordinateFoundFlag == "no") { data[line].push(''); }
    }
  }

  spreadSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  SpreadsheetApp.flush();  //makes sure that the data is written to the spreadsheet before building the xlsx
  // Save an xlsx-version of the sheet in the subfolder for this experiment
  makeCopyxlsx(folderID, fileID, fileName);
  // Delete the temporary gSheet    
  try { DriveApp.getFileById(fileID).setTrashed(true); }
  catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Error deleting temp Quantos gSheet in folder with error: ' + e, 'Error', 3);
  }
}

/**
 * FileGenerator: Creates a new Excel file in the given folder with the given filename (analogous to the createSpreadsheet-function ) unless it already exists.
 * copied from https://yagisanatode.com/2018/07/08/google-apps-script-how-to-create-folders-in-directories-with-driveapp/ and amended to create Spreadsheet, not folder
 * Used to create the Excel version of the Quantos input file for debugging purposes. 
 * @param {String} folderID id of the folder in which the Excelfile is to be created.
 * @param {String} fileID id of the gsheet used to create the Excel file.
 * @param {String} fileName name of the new Excel file.
 */
function makeCopyxlsx(folderID = globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["CHEATSHEETSfOLDERiD"], fileID = "", fileName = "") {
  // adapted from https://stackoverflow.com/questions/49963584/export-a-google-sheet-to-google-drive-in-excel-format-with-apps-script

  var destination = DriveApp.getFolderById(folderID);

  var url = "https://docs.google.com/spreadsheets/d/" + fileID + "/export?format=xlsx&access_token=" + ScriptApp.getOAuthToken();


  //The following line is neccessary according to the source, but at least for now it isn't.
  //var res = UrlFetchApp.fetch(url, {headers: {Authorization: "Bearer " + ScriptApp.getOAuthToken()}});
  var blob = UrlFetchApp.fetch(url).getBlob().setName(fileName + ".xlsx"); // Modified
  destination.createFile(blob);
}

/**
 * FileGenerator: Generates the xml-string later to be written to the Quantos input file
 * 
 * @param {Array} quantosXmlArray array containing info on all the individual solid dosings.
 * @return {String} String to be written to the final xml-file.
 */
function createQuantosXml(quantosXmlArray) {
  //  adapted from https://qiita.com/mhgp/items/576a3f5f84ee5544cc33
  // Generates the xml-document for Quantos Chronect

  var grdSettingsArray = [["chkRepeatSchedule", "False"], ["chkPrioritySchedule", "False"], ["chkOverlappedSchedule", "False"]];  // settings information, later used to generate the second worksheet 
  var cellElement; // will later contain the xml cellElement
  var dataElement; // will later hold the xml dataElement
  var defaultNamespace = XmlService.getNamespace('urn:schemas-microsoft-com:office:spreadsheet');   // not clear why there are two name spaces and what they do
  var root = XmlService.createElement('Workbook', defaultNamespace);
  var namespaceSs = XmlService.getNamespace('ss', 'urn:schemas-microsoft-com:office:spreadsheet');

  var grdSampleList = XmlService.createElement('Worksheet', defaultNamespace)
    .setAttribute('Name', 'grdSampleList', namespaceSs);
  root.addContent(grdSampleList);

  var grdSampleListTable = XmlService.createElement('Table', defaultNamespace);
  grdSampleList.addContent(grdSampleListTable);
  var rowElement; // will later contain the xml row element
  for (var row = 0; row < quantosXmlArray.length; row++) {    // cycle through the array containing the individual dosings and write them into individual rows

    rowElement = XmlService.createElement('Row', defaultNamespace);
    grdSampleListTable.addContent(rowElement);

    for (var cell = 0; cell < quantosXmlArray[row].length; cell++) {   // Write the individual cells
      cellElement = XmlService.createElement('Cell', defaultNamespace);
      rowElement.addContent(cellElement);

      dataElement = XmlService.createElement('Data', defaultNamespace)
        .setAttribute('Type', 'String', namespaceSs)
        .setText(quantosXmlArray[row][cell]);
      cellElement.addContent(dataElement);
    }
  }
  var grdSettings = XmlService.createElement('Worksheet', defaultNamespace)
    .setAttribute('Name', 'grdSettings', namespaceSs);
  root.addContent(grdSettings);

  var grdSettingsTable = XmlService.createElement('Table', defaultNamespace);
  grdSettings.addContent(grdSettingsTable);

  for (row = 0; row < grdSettingsArray.length; row++) {
    rowElement = XmlService.createElement('Row', defaultNamespace);
    grdSettingsTable.addContent(rowElement);

    for (let cell = 0; cell < grdSettingsArray[row].length; cell++) {         //   write the settings sheet

      cellElement = XmlService.createElement('Cell', defaultNamespace);
      rowElement.addContent(cellElement);

      dataElement = XmlService.createElement('Data', defaultNamespace)
        .setAttribute('Type', 'String', namespaceSs)
        .setText(grdSettingsArray[row][cell]);
      cellElement.addContent(dataElement);
    }
  }

  var document = XmlService.createDocument(root);
  var content = XmlService.getCompactFormat().setEncoding('UTF-8').setOmitDeclaration(true).format(document);
  var payload = '<?xml version="1.0"?>' + '<?mso-application progid="Excel.Sheet"?>' + content;
  return payload;
}

//****************************

/**
 * FileGenerator: Creates the LEA xml file
 * part of the LEA-xml generation, currently not used since liquid dosing / sampling is performed manually
 * 
 * @return {Array} contains the xml-Text and the locationsOnSourcePlates dictionary.
 */
function createLeaXml(chemicalsAndMixtures, wellsDictionary, projectName, stepName, reactionType, platePurpose, diluentVolume, sampleVolume, ELNID, operator, plateNumber, filledRows, filledColumns, stirrerRpm, reactionTime, temperature, overage, postDesignCreatorFlag, rowShift, columnShift, rowsOnReactionPlate, columnsOnReactionPlate) {

  if (filledColumns % 6 != 0) {   // Because of the 6-tip, only designs with 6 or 12 columns are supported
    //Browser.msgBox("This script currently only supports plate designs containing 6 or 12 columns for Junior input file generation. Your design contains " + filledColumns + " columns. Both the Quantos and LC-MS input files will be written normally.")
    return "";
  }


  var colors = [32896, 8388608, 32768, 8388736, 12648384, 16711935, 32960, 4210752, 12648447, 16776960, 65280, 16711680, 12640511, 16777152, 4210816, 8388672, 12632319, 16512, 49344, 65535, 16761024, 255, 33023, 0]; //color values extracted from sample xml-files
  var colorCounter = 0;
  var colorPeriod = colors.length;  // The colors used for labelling the different ingredients are cycled over the array, since it's unknown wheter deviating from these values extracted from legit xml files will lead to problems
  //  var overage = 2.5                // Multiplier with which the required amounts of solvents/liquids/solutions are multiplied to get to the amounts put into the source plates. This value can now be controlled by the user and is now handed down to the functio
  var ingredientTypes = [];        // will be populated with the different component roles that end up as LSTypes

  var rowToLetterAssignment = { 1: "A", 2: "B", 3: "C", 4: "D", 5: "E", 6: "F", 7: "G", 8: "H" };   //translates row numbers into the corresponding letter

  var locationsOnSourcePlates = {
    "96_0.8": [[8, 12, 0.8], [[]]],
    "48_2.0": [[6, 8, 2.0], [[]]],
    "24_4.0": [[4, 6, 4.0], [[]]],
    "24_8.0": [[4, 6, 8.0], [[]]]
  };   //available plate types: 96 with max 0.8 mL per vial, 48  with max 2 mL, 24 with max 4 mL , 24 with max 8 mL

  var coordinateFirstWell = [1 + rowShift, 1 + columnShift];      // If the plate is not full, the experiment containing vials are centered and thus the first well will not be [1,1] (A1)
  var coordinateLastWell = [rowShift + filledRows, columnShift + filledColumns];
  var wellInTheMiddle = rowToLetterAssignment[coordinateFirstWell[0] + Math.round(filledRows / 2 + 0.01) + 1] + coordinateFirstWell[1];          // trick to avoid tripping up indentation and text highlighting below the sampling Dictionary (occurs if the content of the variable is put directly into the dictionary!)


  /*var coordinateFirstWell = [1 + Math.round((8- filledRows)/2 - 0.01), 1 + Math.round((12- filledColumns)/2 + 0.01)]      // If the plate is not full, the experiment containing vials are centered and thus the first well will not be [1,1] (A1)
  var coordinateLastWell = [ Math.round((8- filledRows)/2 - 0.01)+filledRows,  Math.round((12- filledColumns)/2 + 0.01)+filledColumns]
  var wellInTheMiddle = rowToLetterAssignment[coordinateFirstWell[0]+Math.round(filledRows/2+0.01)+1]+coordinateFirstWell[1]          // trick to avoid tripping up indentation and text highlighting below the sampling Dictionary (occurs if the content of the variable is put directly into the dictionary!)
  
  */
  var diluentInfo = []; // will later be filled with the info on the diluent
  var samplingDictionary = {}; // will below be filled with info in the samplingDictionary
  if (1.2 * diluentVolume * filledRows * filledColumns / 6 < 8000) {    // the diluent also needs to be allocated to a source plate. This code does that by creating an array in analogy to chemicalsAndMixtures and a dictionary in analogy to wellsDictionary
    diluentInfo = [[0, 0, 0, "Solution", 0, 0, 0, 0, 0, 0, diluentVolume * filledRows * filledColumns, "MeCN/water sample diluent", ""]];
    samplingDictionary = { "MeCN/water sample diluent": [[], [[0, 0, 0, diluentVolume]], [coordinateFirstWell[0], coordinateFirstWell[1], rowToLetterAssignment[coordinateFirstWell[0]] + coordinateFirstWell[1], coordinateLastWell[0], coordinateLastWell[1], rowToLetterAssignment[coordinateLastWell[0]] + coordinateLastWell[1]]] };
  } else {
    diluentInfo = [[0, 0, 0, "Solution", 0, 0, 0, 0, 0, 0, diluentVolume * filledRows * filledColumns / 2, "MeCN/water sample diluent 1", ""],
    [0, 0, 0, "Solution", 0, 0, 0, 0, 0, 0, diluentVolume * filledRows * filledColumns / 2, "MeCN/water sample diluent 2", ""]];
    samplingDictionary = {
      "MeCN/water sample diluent 1": [[], [[0, 0, 0, diluentVolume]], [coordinateFirstWell[0], coordinateFirstWell[1], rowToLetterAssignment[coordinateFirstWell[0]] + coordinateFirstWell[1], coordinateFirstWell[0] + Math.round(filledRows / 2 + 0.01), coordinateLastWell[1], rowToLetterAssignment[coordinateFirstWell[0] + Math.round(filledRows / 2 + 0.01)] + coordinateLastWell[1]]],
      "MeCN/water sample diluent 2": [[], [[0, 0, 0, diluentVolume]], [coordinateFirstWell[0] + Math.round(filledRows / 2 + 0.01) + 1, coordinateFirstWell[1], wellInTheMiddle, coordinateLastWell[0], coordinateLastWell[1], rowToLetterAssignment[coordinateLastWell[0]] + coordinateLastWell[1]]]
    };
  }


  for (var element = 0; element < diluentInfo.length; element++) { //put the diluent on source plates with a standard overage of 20%
    locationsOnSourcePlates = amendSourcePlates(1.2, diluentInfo[element], locationsOnSourcePlates);
  }


  var parameterParameters = [ // ['Chronect dispense weight','Number',0,513] Contains the variable bits of the individual parameters, significance largely unknown
    ['StartReactionTimer', 'Number', 0, 0],
    ['StirRate', 'Stir Rate', 3585, 3585],
    ['HeatingTemp', 'Temperature', 1281, 1281],
    ['Delay', 'Time', 770, 770]];

  var nameReactionPlate = "reaction plate";

  // This array contains information on the different decks that are present on the deck and will be amended with additional plates depending on the number and volume of diluent/solvents/solutions present in the experiment
  var platesOnDeck = [[nameReactionPlate, rowsOnReactionPlate, columnsOnReactionPlate, projectName, 1],
  ['sampling plate', 8, 12, projectName, 2]];

  var root = XmlService.createElement('LibraryDesign') //root element, meaning of the attributes unknown
    .setAttribute('ID', '0')
    .setAttribute('ConCheck', '0')
    .setAttribute('LinkID', '')
    .setAttribute('PersistState', '0');

  //*************    Header Information  

  if (postDesignCreatorFlag) {
    root.addContent(XmlService.createElement('Name').setText(stepName));  // some of this info will need to be variable before deployment
  } else { root.addContent(XmlService.createElement('Name').setAttribute('Null', '1')); }
  root.addContent(XmlService.createElement('CreatedBy').setText("Georg Wuitschik"));
  root.addContent(XmlService.createElement('LastModifiedBy').setText(operator));
  if (postDesignCreatorFlag) { root.addContent(XmlService.createElement('Project').setText(projectName)); } else { root.addContent(XmlService.createElement('Project').setAttribute('Null', '1')); }
  root.addContent(XmlService.createElement('Notebook').setText(ELNID));
  root.addContent(XmlService.createElement('Pages').setText(plateNumber));
  root.addContent(XmlService.createElement('Comments').setText(platePurpose));
  root.addContent(XmlService.createElement('Staff').setAttribute('Null', '1'));
  if (postDesignCreatorFlag) {
    root.addContent(XmlService.createElement('Keywords').setText(reactionType));  //not clear why a date is in there in the original version and what its significance is
  } else { root.addContent(XmlService.createElement('Keywords').setAttribute('Null', '1')); }
  root.addContent(XmlService.createElement('Type').setAttribute('Null', '1'));
  root.addContent(XmlService.createElement('CreationDate').setText('2020/02/13 08:53:56'));
  root.addContent(XmlService.createElement('LastModificationDate').setText(Utilities.formatDate(new Date(), 'GMT+1', "yyyy/MM/dd HH:mm:ss")));
  root.addContent(XmlService.createElement('ProductVersion').setText('9.1.0.0'));
  root.addContent(XmlService.createElement('SchemaVersion').setText('3.0'));
  root.addContent(XmlService.createElement('Status').setText('0'));
  root.addContent(XmlService.createElement('Flags').setText('0'));
  root.addContent(XmlService.createElement('SavedBy').setText(operator));
  root.addContent(XmlService.createElement('SaveDate').setText('2020/02/13 08:53:56'));
  root.addContent(XmlService.createElement('LastSavedBy').setText(operator));
  root.addContent(XmlService.createElement('LastSaveDate').setText(Utilities.formatDate(new Date(), 'GMT+1', "yyyy/MM/dd HH:mm:ss")));
  root.addContent(XmlService.createElement('OriginID').setText('0'));
  root.addContent(XmlService.createElement('OriginDBID').setAttribute('Null', '1'));

  // ************** LS Types section   different categories for the ingredients (solvent, base, catalyst...)

  var LSTypes = XmlService.createElement('LSTypes'); //Contains info on the types of ingredients present
  root.addContent(LSTypes);

  for (var key in chemicalsAndMixtures) {                 // Create an array of unique component roles (Ligand, Base, Solvent...) present on the plate and register the component roles as LSType
    if (ingredientTypes.indexOf(chemicalsAndMixtures[key][0]) == -1) {
      ingredientTypes.push(chemicalsAndMixtures[key][0]);
      var LSType = XmlService.createElement('LSType');
      LSTypes.addContent(LSType);
      LSType.addContent(XmlService.createElement('Name').setText(chemicalsAndMixtures[key][0]));
    }
  }

  // **************    LS Chemicals section (individual chemicals, their properties and total amount dispensed)

  var LSChemicals = XmlService.createElement('LSChemicals');    // Defines the different chemicals used and how much in total is present
  root.addContent(LSChemicals);

  // plateIngredientsDictionary[data[row][3]+'_'+data[row][7]+'_'+data[row][18]+'_'+data[row][19]+'_'+data[row][20]] = [data[row][1],data[row][10], data[row][11],data[row][16],data[row][18],data[row][19],data[row][20],data[row][22],    0,         0,           0,       data[row][3]  , , data[row][4],      data[row][24]    ] // new entry to dictionary is created with component name as key and array of component role, MW, density, Dose as,  Concentration, Unit, Solvent, Solution Density and trailing zeros for mass, volume liquid, volume solution as value,
  //                            Comp name        Batch ID          Concentration     Unit               Solvent            0            1             2             3               4            5         6               7            8,         9,          10                11            12                13
  //                                                                                                                      Role              MW            density      dose as         conc          unit      solvent     sol density      mass    liquid vol      sol. vol      Comp Name       Limit/Level       to be evaporated?
  var doseUsing6Tip = true;

  for (let key in chemicalsAndMixtures) {
    if (chemicalsAndMixtures[key][12] === true || chemicalsAndMixtures[key][12] === false) { chemicalsAndMixtures[key][12] = ''; } //This is mostly '', may contain L1, L2... if different levels of reagents are employed, but should not contain true, false which it would for the starting materials. Thus, they're filtered out.
    doseUsing6Tip = true;
    if ((chemicalsAndMixtures[key][11] != chemicalsAndMixtures[key][6]) && //this is triggered, if the entry in question belongs to a solvent that is used for solution preparation which doesn't need to be placed on a plate
      (wellsDictionary[chemicalsAndMixtures[key][11] + chemicalsAndMixtures[key][12]][0][2] != "Solid") || (wellsDictionary[chemicalsAndMixtures[key][11] + chemicalsAndMixtures[key][12]][0][2] != "NonDose Solid")) {
      for (var dosingRange = 2; dosingRange < wellsDictionary[chemicalsAndMixtures[key][11] + chemicalsAndMixtures[key][12]].length; dosingRange++) {
        if ((wellsDictionary[chemicalsAndMixtures[key][11] + chemicalsAndMixtures[key][12]][dosingRange][4] - wellsDictionary[chemicalsAndMixtures[key][11] + chemicalsAndMixtures[key][12]][dosingRange][1] + 1) % 6 != 0) {
          doseUsing6Tip = false;     // if the width of any of the dosing ranges is not divisible by 6, the 6-tip can't be used and the flag is set to false. If all dosing ranges are multiples of 6 columns wide, the true remains
          wellsDictionary[chemicalsAndMixtures[key][11] + chemicalsAndMixtures[key][12]][0][2] = "no6Tip";
        }
      }

    }
    chemicalsAndMixtures[key].push(doseUsing6Tip);    // append the flag to ChemicalsAndMixtures, so when cycling through it later, the program can recognize which liquids/solutions should not be allocated to source plates or dosed using the single tip
  }

  for (let key in chemicalsAndMixtures) {                        // Iterates over all the different compounds present in chemicalsAndMixtures and creates an entry (LSChemical) for each one
    if (chemicalsAndMixtures[key][13] === true) {  //if the compound is preplated as a solution and then evaporated, it's treated as a solid for the purpose of the Junior file (no dosing step is generated for allocation to a source plate) 
      chemicalsAndMixtures[key][3] = "evaporated Solution";
    }


    var LSChemical = XmlService.createElement('LSChemical')
      .setAttribute('ID', '0')
      .setAttribute('ConCheck', '0')
      .setAttribute('PersistState', '0')
      .setAttribute('LinkID', '');
    LSChemicals.addContent(LSChemical);
    LSChemical.addContent(XmlService.createElement('Version').setText(2));
    LSChemical.addContent(XmlService.createElement('Name').setText(chemicalsAndMixtures[key][11] + chemicalsAndMixtures[key][12]));
    LSChemical.addContent(XmlService.createElement('Color').setText(colors[colorCounter % colorPeriod]));                                      // cycle through the different colors defined in the colors array
    colorCounter++;
    switch (chemicalsAndMixtures[key][3]) {        // Depending on the dosing mode, different info is written. 
      case "Solid":
      case "NonDose Solid":
      case "evaporated Solution":
        LSChemical.addContent(XmlService.createElement('UnitDispensed').setText(257)
          .setAttribute("Name", "mg"));
        LSChemical.addContent(XmlService.createElement('AmountDispensed').setText(parseFloat(chemicalsAndMixtures[key][8].toFixed(2))));   // total mass of the chemical in question dosed as solid
        LSChemical.addContent(XmlService.createElement('Density').setText("0."));
        LSChemical.addContent(XmlService.createElement('MolW').setText(chemicalsAndMixtures[key][1]));                                     // molecular weight of the compound
        LSChemical.addContent(XmlService.createElement('SourceArrayName').setAttribute("Null", "1"));
        LSChemical.addContent(XmlService.createElement('SourceArrayID').setText(0));
        LSChemical.addContent(XmlService.createElement('Position').setText(1));
        LSChemical.addContent(XmlService.createElement('WSAmount').setText(parseFloat(chemicalsAndMixtures[key][8].toFixed(2))));         // total mass of the chemical in question dosed as solid (don't know why it repeats)
        LSChemical.addContent(XmlService.createElement('WSUnit').setText(257)
          .setAttribute("Name", "mg"));
        break;
      case "Liquid":
        LSChemical.addContent(XmlService.createElement('UnitDispensed').setText(513).setAttribute("Name", "ul"));
        LSChemical.addContent(XmlService.createElement('AmountDispensed').setText(parseFloat(chemicalsAndMixtures[key][9].toFixed(2))));  // total volume of the chemical in question dosed as neat liquid
        LSChemical.addContent(XmlService.createElement('Density').setText(chemicalsAndMixtures[key][2]));
        LSChemical.addContent(XmlService.createElement('MolW').setText(chemicalsAndMixtures[key][1]));
        LSChemical.addContent(XmlService.createElement('SourceArrayName').setAttribute("Null", "1"));
        LSChemical.addContent(XmlService.createElement('SourceArrayID').setText(0));
        LSChemical.addContent(XmlService.createElement('Position').setText(1));
        LSChemical.addContent(XmlService.createElement('WSAmount').setText(parseFloat(chemicalsAndMixtures[key][9].toFixed(2))));         // total volume of the chemical in question dosed as neat liquid
        LSChemical.addContent(XmlService.createElement('WSUnit').setText(513).setAttribute("Name", "ul"));

        // ** Allocate the liquid to a source plate

        if (chemicalsAndMixtures[key][6] != chemicalsAndMixtures[key][11] && chemicalsAndMixtures[key][14] === true) {         //     Allocate all liquids to a source plate unless the liquid is dissolved in itself (unlikely to happen, but would lead to allocating the compound both as liquid and as solution)
          locationsOnSourcePlates = amendSourcePlates(overage, chemicalsAndMixtures[key], locationsOnSourcePlates);
        } else { Logger.log(chemicalsAndMixtures[key][11] + " was not included"); }

        break;
      case "Solution":
        LSChemical.addContent(XmlService.createElement('UnitDispensed').setText(0).setAttribute("Name", "undefined"));
        LSChemical.addContent(XmlService.createElement('AmountDispensed').setText("0.")); // if the compound is part of a solution (true for both, solvent and solute), then the amount dispensed is set to "0." PROBLEM: If the solvent is also present as a pure compound, then in the real xml, the total amount (and unit) of the pure compound is displayed
        LSChemical.addContent(XmlService.createElement('Density').setText(chemicalsAndMixtures[key][7]));                                 // density of the solution                                                      Right now, two entries for the solvent in question are prepared, one for each case (liquid and solution)
        LSChemical.addContent(XmlService.createElement('MolW').setText(chemicalsAndMixtures[key][1]));
        LSChemical.addContent(XmlService.createElement('SourceArrayName').setAttribute("Null", "1"));
        LSChemical.addContent(XmlService.createElement('SourceArrayID').setText(0));
        LSChemical.addContent(XmlService.createElement('Position').setText(1));
        LSChemical.addContent(XmlService.createElement('WSAmount').setText("0."));
        LSChemical.addContent(XmlService.createElement('WSUnit').setText(0).setAttribute("Name", "undefined"));

        // ** Allocate the solution to a source plate

        if (chemicalsAndMixtures[key][6] != chemicalsAndMixtures[key][11] && chemicalsAndMixtures[key][14] === true) {          //     Allocate all solutions to a source plate, the solvent part is not allocated to avoid duplication 
          locationsOnSourcePlates = amendSourcePlates(overage, chemicalsAndMixtures[key], locationsOnSourcePlates);
        } else { Logger.log(chemicalsAndMixtures[key][11] + " was not included"); }

        break;
    }

    LSChemical.addContent(XmlService.createElement('WSBarcode').setAttribute("Null", "1"));

    var LSEquivalences = XmlService.createElement('LSEquivalences');
    LSChemical.addContent(LSEquivalences);

    var LSEquivalence = XmlService.createElement('LSEquivalence');
    LSEquivalences.addContent(LSEquivalence);

    LSEquivalence.addContent(XmlService.createElement('TypeID').setText(chemicalsAndMixtures[key][0]));      // register the component role as LSEquivalence
    LSEquivalence.addContent(XmlService.createElement('Value').setText("0."));
    LSEquivalence.addContent(XmlService.createElement('Unit').setText(2561).setAttribute("Name", "none"));
    LSChemical.addContent(XmlService.createElement('NamedAttributes'));
  }












  locationsOnSourcePlates = optimizeSourcePlates(locationsOnSourcePlates); // minimizes the number of plates used by moving smaller volume components to a bigger plate, if it results in an overall reduction of plates

  // **************    LS Mixtures section  (Mixtures/solutions of compounds and total amount dispensed)

  var LSMixtures = XmlService.createElement('LSMixtures');  // If there are solutions present, they're entered here
  root.addContent(LSMixtures);
  // Add up to two entries for the MeCN/water mix sample diluent (one entry wouldn't provide the volume needed to fill 96 sample vials from an 24 vial plate with 8 ml each)
  var LSMixture; // will contain the LSMixture xml-element;
  for (let key in samplingDictionary) {

    LSMixture = XmlService.createElement('LSMixture')
      .setAttribute('ID', '0')
      .setAttribute('ConCheck', '0')
      .setAttribute('PersistState', '0')
      .setAttribute('LinkID', '');
    LSMixtures.addContent(LSMixture);
    LSMixture.addContent(XmlService.createElement('Version').setText(2));
    LSMixture.addContent(XmlService.createElement('Name').setText(key));
    LSMixture.addContent(XmlService.createElement('Color').setText(colors[colorCounter % colorPeriod]));  // pick the next color from the colors array
    colorCounter++;
    LSMixture.addContent(XmlService.createElement('DensityCalculation').setText(0).setAttribute("Name", "Known"));
    LSMixture.addContent(XmlService.createElement('UnitDispensed').setText(513).setAttribute("Name", "ul"));
    LSMixture.addContent(XmlService.createElement('AmountDispensed').setText(diluentVolume * filledRows * filledColumns / Object.keys(samplingDictionary).length));                               // total solution volume dispensed
    LSMixture.addContent(XmlService.createElement('SourceArrayName').setAttribute("Null", "1"));
    LSMixture.addContent(XmlService.createElement('SourceArrayID').setText(0));
    LSMixture.addContent(XmlService.createElement('Position').setText(1));
    LSMixture.addContent(XmlService.createElement('WSAmount').setText("0."));
    LSMixture.addContent(XmlService.createElement('WSUnit').setText(0).setAttribute("Name", "undefined"));
    LSMixture.addContent(XmlService.createElement('SourceArrayDBID').setAttribute("Null", "1"));
    LSMixture.addContent(XmlService.createElement('Density').setText("1.").setAttribute("Name", "Known"));
    LSMixture.addContent(XmlService.createElement('DensityCorrection').setText("0."));
  }


  for (let key in chemicalsAndMixtures) {    // iterate through all the components 
    if (chemicalsAndMixtures[key][3] != "Solution" || chemicalsAndMixtures[key][6] == chemicalsAndMixtures[key][11]) { continue; } // skip this chemical, if the compound is not dosed as solution or if the name of the solvent is the same as the name of the component, true for all liquids/solids as well as solvents used in solutions
    LSMixture = XmlService.createElement('LSMixture')
      .setAttribute('ID', '0')
      .setAttribute('ConCheck', '0')
      .setAttribute('PersistState', '0')
      .setAttribute('LinkID', '');
    LSMixtures.addContent(LSMixture);
    LSMixture.addContent(XmlService.createElement('Version').setText(2));
    LSMixture.addContent(XmlService.createElement('Name').setText(chemicalsAndMixtures[key][4] + chemicalsAndMixtures[key][5] + " " + chemicalsAndMixtures[key][11] + chemicalsAndMixtures[key][12] + " in " + chemicalsAndMixtures[key][6]));     //  legit according to Vera's test files 
    LSMixture.addContent(XmlService.createElement('Color').setText(colors[colorCounter % colorPeriod]));  // pick the next color from the colors array
    colorCounter++;
    LSMixture.addContent(XmlService.createElement('DensityCalculation').setText(1).setAttribute("Name", "Ideal mixing"));
    LSMixture.addContent(XmlService.createElement('UnitDispensed').setText(513).setAttribute("Name", "ul"));
    LSMixture.addContent(XmlService.createElement('AmountDispensed').setText(parseFloat(chemicalsAndMixtures[key][10].toFixed(2)))); // total solution volume dispensed
    LSMixture.addContent(XmlService.createElement('SourceArrayName').setAttribute("Null", "1"));
    LSMixture.addContent(XmlService.createElement('SourceArrayID').setText(0));
    LSMixture.addContent(XmlService.createElement('Position').setText(1));
    LSMixture.addContent(XmlService.createElement('WSAmount').setText(parseFloat(chemicalsAndMixtures[key][10].toFixed(2))));        // total solution volume dispensed
    LSMixture.addContent(XmlService.createElement('WSUnit').setText(513).setAttribute("Name", "ul"));
    LSMixture.addContent(XmlService.createElement('SourceArrayDBID').setAttribute("Null", "1"));
    LSMixture.addContent(XmlService.createElement('Density').setText(chemicalsAndMixtures[key][7]).setAttribute("Name", "Ideal mixing"));    // solution density 
    LSMixture.addContent(XmlService.createElement('DensityCorrection').setText("0."));
    var LSComponents = XmlService.createElement('LSComponents');
    LSMixture.addContent(LSComponents);
    var LSComponent = XmlService.createElement('LSComponent');
    LSComponents.addContent(LSComponent);
    LSComponent.addContent(XmlService.createElement('Source').setText(chemicalsAndMixtures[key][11] + chemicalsAndMixtures[key][12]));   // name of solute and Level (L1, L2...)
    LSComponent.addContent(XmlService.createElement('Value').setText(chemicalsAndMixtures[key][4]));   // concentration
    LSComponent.addContent(XmlService.createElement('Unit').setText(2817).setAttribute("Name", "mol/l"));     // needs to be variable depending on solvent unit: PROBLEM: Not clear which solvent units are supported and how they're called

    if (chemicalsAndMixtures[key][2] == '-') { // true, if the solute in question is a solid
      LSComponent.addContent(XmlService.createElement('WSAmount').setText(parseFloat(chemicalsAndMixtures[key][8].toFixed(2))));    // total mass of the solid solute
      LSComponent.addContent(XmlService.createElement('WSUnit').setText(257).setAttribute("Name", "mg"));
    } else {
      LSComponent.addContent(XmlService.createElement('WSAmount').setText(parseFloat(chemicalsAndMixtures[key][9].toFixed(2))));    // total volume of the liquid solute
      LSComponent.addContent(XmlService.createElement('WSUnit').setText(513).setAttribute("Name", "ul"));
    }
    LSComponent.addContent(XmlService.createElement('WSBarcode').setAttribute("Null", "1"));
    LSComponent = XmlService.createElement('LSComponent');
    LSComponents.addContent(LSComponent);
    LSComponent.addContent(XmlService.createElement('Source').setText(chemicalsAndMixtures[key][6]));   // component ID solvent
    LSComponent.addContent(XmlService.createElement('Value').setText("0."));
    LSComponent.addContent(XmlService.createElement('Unit').setText(2819).setAttribute("Name", "remainder"));
    LSComponent.addContent(XmlService.createElement('WSAmount').setText(parseFloat((chemicalsAndMixtures[key][10] - chemicalsAndMixtures[key][8]) / 1000).toFixed(4)));  //approximate volume of the solvent (vol Solution - mass) in mL 
    LSComponent.addContent(XmlService.createElement('WSUnit').setText(514).setAttribute("Name", "ml"));   // There seems to be no strict rule when to use mL vs uL, mL was adopted because it was found in Rick's example file
    LSComponent.addContent(XmlService.createElement('WSBarcode').setAttribute("Null", "1"));
  }
  LSMixture.addContent(XmlService.createElement('NamedAttributes'));

  // **************    LS Parameters section (Parameters defined, not sure how important)

  var LSParameters = XmlService.createElement('LSParameters');
  root.addContent(LSParameters);
  if (postDesignCreatorFlag) { // parameters are only defined, if processing steps should be generated (Checkbox in File Generator sheet)
    for (var constituents = 0; constituents < parameterParameters.length; constituents++) { //Loops through the parameterParameters array and creates the parameters like heating temmp, stir rate, delay, start timer

      var LSParameter = XmlService.createElement('LSParameter') //root element of the parameter
        .setAttribute('ID', '0')
        .setAttribute('ConCheck', '0')
        .setAttribute('PersistState', '0')
        .setAttribute('LinkID', '');

      LSParameter.addContent(XmlService.createElement('Version').setText(2));
      LSParameter.addContent(XmlService.createElement('Name').setText(parameterParameters[constituents][0]));
      LSParameter.addContent(XmlService.createElement('Type').setText(parameterParameters[constituents][1]));
      LSParameter.addContent(XmlService.createElement('Description').setAttribute('Null', '1'));
      LSParameter.addContent(XmlService.createElement('CanVaryAcrossRows').setText(0));
      LSParameter.addContent(XmlService.createElement('CanVaryAcrossCols').setText(0));
      LSParameter.addContent(XmlService.createElement('DefaultUnit').setText(parameterParameters[constituents][2]));
      LSParameter.addContent(XmlService.createElement('SourceUnit').setText(parameterParameters[constituents][3]));
      LSParameter.addContent(XmlService.createElement('Expression').setAttribute('Null', '1'));
      LSParameter.addContent(XmlService.createElement('DecimalPlaces').setText(2));
      LSParameters.addContent(LSParameter);
    }
  }
  //*************  LS Libraries section (the plates used on deck)

  var LSLibraries2 = XmlService.createElement('LSLibraries2');
  root.addContent(LSLibraries2);

  var index = 3;

  for (key in locationsOnSourcePlates) {                   // go through the different types of source plates and add the plates found on there to the platesOnDeck array
    if (locationsOnSourcePlates[key][1].length > 0) {
      for (var sourcePlate = 0; sourcePlate < locationsOnSourcePlates[key][1].length; sourcePlate++) {
        platesOnDeck.push([key + " mL " + (sourcePlate + 1), locationsOnSourcePlates[key][0][0], locationsOnSourcePlates[key][0][1], projectName, index]);
        index++;
      }
    }
  }

  var plateName = "reaction plate";   //initialize variables representing the array elements, initial names are irrelevant
  var rows = rowsOnReactionPlate;
  var columns = columnsOnReactionPlate;
  var project = "Test";
  var color = 255;

  for (let constituents = 0; constituents < platesOnDeck.length; constituents++) {        // go through the platesOnDeck array and write them to the xml

    plateName = platesOnDeck[constituents][0];    // updates the variables with the array elements drawn from the respective row in the array
    rows = platesOnDeck[constituents][1];
    columns = platesOnDeck[constituents][2];
    project = platesOnDeck[constituents][3];
    index = platesOnDeck[constituents][4];
    color = colors[colorCounter % colorPeriod];
    colorCounter++;

    var plate = XmlService.createElement('LSLibrary') //root element of the plate
      .setAttribute('ID', '0')
      .setAttribute('ConCheck', '0')
      .setAttribute('PersistState', '0')
      .setAttribute('LinkID', '');
    LSLibraries2.addContent(plate);  // link the LS library section generated just now to the LSLibraries2 chapter of the file
    plate.addContent(XmlService.createElement('Version').setText(2));
    plate.addContent(XmlService.createElement('Index').setText(index));
    plate.addContent(XmlService.createElement('LibraryID').setText(0));
    plate.addContent(XmlService.createElement('Name').setText(plateName));
    plate.addContent(XmlService.createElement('Description').setAttribute('Null', '1'));
    plate.addContent(XmlService.createElement('Rows').setText(rows));
    plate.addContent(XmlService.createElement('Columns').setText(columns));
    plate.addContent(XmlService.createElement('Shape').setText(0));
    plate.addContent(XmlService.createElement('Flags').setText(0));
    plate.addContent(XmlService.createElement('Color').setText(color));
    plate.addContent(XmlService.createElement('Notebook').setAttribute('Null', '1'));
    plate.addContent(XmlService.createElement('Pages').setAttribute('Null', '1'));
    if (postDesignCreatorFlag) { plate.addContent(XmlService.createElement('Project').setText(projectName)); } else { plate.addContent(XmlService.createElement('Project').setAttribute('Null', '1')); }
    plate.addContent(XmlService.createElement('SynID').setText(0));
    plate.addContent(XmlService.createElement('SynConCheck').setText(0));
    plate.addContent(XmlService.createElement('IsSourceLibrary').setText(0));
    plate.addContent(XmlService.createElement('Staff').setAttribute('Null', '1'));
    plate.addContent(XmlService.createElement('OriginLibDBID').setAttribute('Null', '1'));
    plate.addContent(XmlService.createElement('OriginLibID').setText(0));
    var LSSourceMgr = XmlService.createElement('LSSourceMgr');
    plate.addContent(LSSourceMgr);
    LSSourceMgr.addContent(XmlService.createElement('LSTypes'))
      .addContent(XmlService.createElement('LSChemicals'))
      .addContent(XmlService.createElement('LSMixtures'))
      .addContent(XmlService.createElement('LSParameters'));
    plate.addContent(XmlService.createElement('NamedAttributes'));

    var LSLibraryElements = XmlService.createElement('LSLibraryElements');
    plate.addContent(LSLibraryElements);
    var position = 1; //counts through the positions on the plate

    for (var row = 1; row < rows + 1; row++) {         // generate the plate in question by iterating over all positions in rows and columns
      for (var column = 1; column < columns + 1; column++) {
        var LSLibraryElement = XmlService.createElement('LSLibraryElement')  // one LSLibraryElement = one well
          .setAttribute('ID', '0')
          .setAttribute('ConCheck', '0')
          .setAttribute('PersistState', '0')
          .setAttribute('LinkID', String(position));
        LSLibraryElement.addContent(XmlService.createElement('Position').setText(position));
        LSLibraryElement.addContent(XmlService.createElement('Status').setText(0));
        LSLibraryElement.addContent(XmlService.createElement('Flags').setText(0));
        LSLibraryElement.addContent(XmlService.createElement('Name').setText(plateName + " (" + rowToLetterAssignment[row] + column + ")"));
        LSLibraryElements.addContent(LSLibraryElement);
        position++;
      }
    }
  }

  //********************     LS Layers which can be solid/liquid/solution dosings, stirring, heating, sampling

  var LSLayers = XmlService.createElement('LSLayers');  //In this section, the different processing steps are defined, be it solid dosings, liquid dosings, stirring, sampling or heating/cooling
  root.addContent(LSLayers);

  var layerIndex = 1;  // layers are numbered

  //  locationsOnSourcePlates   plate type : [   [rowCount,colCount, maxVol], [        [               [                vialVolume, ComponentName+Level            ]]]]
  //                                         [0]   [0][0]                      [1]        [1][0]       [1][0][0]          [1][0][0][0]         [1][0][0][1]
  //                                          Headerinfo on Plate type        all plates    one plate     rows on plate      row property      wellsDictionary key

  //  ** Write the layers that belong to the preparatory steps, putting solvents into the solvent containers, making solutions
  for (key in locationsOnSourcePlates) {                                     //locationsOnSourcePlates dictionary with the plate type as key (e.g. 96_0.8 ) and an array of plates filled with compounds filled in there by function amendSourcePlate (and afterwards optimizeSourcePlate)
    if (locationsOnSourcePlates[key][1].length == 0) { continue; }        // if there's no plate of this category, go to the next one
    for (let sourcePlate = 0; sourcePlate < locationsOnSourcePlates[key][1].length; sourcePlate++) {                            // go through all the plates of this category
      for (var sourcePlateRow = 0; sourcePlateRow < locationsOnSourcePlates[key][1][sourcePlate].length; sourcePlateRow++) {    // go through all the filled rows of the plate in question
        var LSLayer = XmlService.createElement('LSLayer');                                                                     // Create a dosing layer
        LSLayers.addContent(LSLayer);
        var layerIndexElement = XmlService.createElement('Index').setText(layerIndex);  // write the index of this layer
        layerIndex++;
        LSLayer.addContent(layerIndexElement);
        var LSMaps = XmlService.createElement('LSMaps');  // under this label, all the dosing maps for this compound are housed
        LSLayer.addContent(LSMaps);
        var LSSourceMap = XmlService.createElement('LSSourceMap');  // under this label, all the source maps for this compound are housed
        LSMaps.addContent(LSSourceMap);
        LSSourceMap.addContent(XmlService.createElement('MapType').setText("Uniform"));
        LSSourceMap.addContent(XmlService.createElement('Unit').setText(513).setAttribute("Name", "ul"));
        LSSourceMap.addContent(XmlService.createElement('MappedUnit').setText(513).setAttribute("Name", "ul"));
        LSSourceMap.addContent(XmlService.createElement('Description')
          .setText('Add ' + parseFloat(overage * locationsOnSourcePlates[key][1][sourcePlate][sourcePlateRow][0]).toFixed(2) + " ul " + locationsOnSourcePlates[key][1][sourcePlate][sourcePlateRow][1] + " to " + key + " mL " + (sourcePlate + 1) + " (" + rowToLetterAssignment[sourcePlateRow + 1] + "1:" + rowToLetterAssignment[sourcePlateRow + 1] + "6) in ul"));
        LSSourceMap.addContent(XmlService.createElement('Annotation').setAttribute("Null", "1"));
        LSSourceMap.addContent(XmlService.createElement('LibraryName').setText(key + " mL " + (sourcePlate + 1)));
        var Source = XmlService.createElement('Source');
        LSSourceMap.addContent(Source);
        Source.addContent(XmlService.createElement('Name').setText(locationsOnSourcePlates[key][1][sourcePlate][sourcePlateRow][1]));     //Component name and level
        Source.addContent(XmlService.createElement('SourceArrayName'));
        Source.addContent(XmlService.createElement('SourcePos'));
        var LSCalcParameters = XmlService.createElement('LSCalcParameters');
        LSSourceMap.addContent(LSCalcParameters);
        var LSCalcParameter = XmlService.createElement('LSCalcParameter').setAttribute("Name", "Value");
        LSCalcParameters.addContent(LSCalcParameter);
        LSCalcParameter.addContent(XmlService.createElement('Flags').setText(0));
        LSCalcParameter.addContent(XmlService.createElement('ID').setText(1));
        LSCalcParameter.addContent(XmlService.createElement('Value').setText(parseFloat(overage * locationsOnSourcePlates[key][1][sourcePlate][sourcePlateRow][0]).toFixed(2)).setAttribute("Type", "5").setAttribute("DI", "0"));
        //LSSourceMap.addContent(XmlService.createElement('Tags').setAttribute("Name", "Tags"));

        var Tags = XmlService.createElement('Tags').setAttribute("Name", "Tags");
        LSSourceMap.addContent(Tags);
        if (postDesignCreatorFlag) {
          Tags.addContent(XmlService.createElement('Tag'));
          Tags.addContent(XmlService.createElement('Tag').setText("SkipMap"));         // added in design creator, so that the layer is not executed by the robot
        }

        var LSMapAmounts = XmlService.createElement('LSMapAmounts');
        LSSourceMap.addContent(LSMapAmounts);
        for (var sourceWell = 1; sourceWell < 7; sourceWell++) {
          var LSMapAmount = XmlService.createElement('LSMapAmount');
          LSMapAmounts.addContent(LSMapAmount);
          LSMapAmount.addContent(XmlService.createElement('Row').setText(sourcePlateRow + 1));
          LSMapAmount.addContent(XmlService.createElement('Column').setText(sourceWell));
          LSMapAmount.addContent(XmlService.createElement('Value').setText(parseFloat(overage * locationsOnSourcePlates[key][1][sourcePlate][sourcePlateRow][0]).toFixed(2)));
          LSMapAmount.addContent(XmlService.createElement('MappedValue').setText(parseFloat(overage * locationsOnSourcePlates[key][1][sourcePlate][sourcePlateRow][0]).toFixed(2)));
        }
      }
    }
  }

  //  ** Dose the compounds found in wellsDictionary to the respective wells ( one compound = one layer, if several levels of a compound are present, these go into separate layers)

  for (key in wellsDictionary) {

    // wellsDictionary[key][0] = Header Information
    //            wellsDictionary[data[row][3]+data[row][4]] = [[data[row][3], data[row][4], data[row][16], data[row][18],data[row][19],data[row][20] ],[]];
    //                                                             component name   limit/level       comp name       limit/level    dose as        concentr      unit        solvent

    // wellsDictionary[key][1] = Dosing Information   [row, column, coordinate, volume/mass]
    // wellsDictionary[key][2] = Dosing Boundaries    [firstRow, firstColumn, firstCoordinate,lastRow, lastColumn, lastCoordinate]


    if (wellsDictionary[key][0][6] === true) {  //if the compound is preplated as a solution and then evaporated, it's treated as a solid and thus no dosing layer from a source plate to the reaction plate is created 
      wellsDictionary[key][0][2] = "evaporated Solution";
    }

    if (wellsDictionary[key][0][1] === true || wellsDictionary[key][0][1] === false) {  //this is true for the starting materials which are flagged as limited sm or not 
      wellsDictionary[key][0][1] = "";
    }
    //wellsDictionary[key][0][1];


    let LSLayer = XmlService.createElement('LSLayer');  // Create a dosing layer
    LSLayers.addContent(LSLayer);
    let layerIndexElement = XmlService.createElement('Index').setText(layerIndex);  // write the index of this layer
    layerIndex++;
    LSLayer.addContent(layerIndexElement);
    let LSMaps = XmlService.createElement('LSMaps');  // under this label, all the dosing maps for this compound are housed
    LSLayer.addContent(LSMaps);

    switch (wellsDictionary[key][0][2]) {  //the maps look different depending on the dosing mode (solids come from no source plate, liquids from an array on a source plate defined in locationsOnSourcePlates
      case "Solid":
      case "NonDose Solid":
      case "evaporated Solution":
      case "no6Tip":
        var dosingRanges = "";
        for (let dosingRange = 2; dosingRange < wellsDictionary[key].length; dosingRange++) {
          if (wellsDictionary[key][dosingRange][2] == wellsDictionary[key][dosingRange][5]) {
            dosingRanges = dosingRanges + "(" + wellsDictionary[key][dosingRange][2] + ")";
          }
          else { dosingRanges = dosingRanges + "(" + wellsDictionary[key][dosingRange][2] + ":" + wellsDictionary[key][dosingRange][5] + ")"; }
        }
        let LSSourceMap = XmlService.createElement('LSSourceMap');  // under this label, all the dosing maps for this compound are housed
        LSMaps.addContent(LSSourceMap);
        LSSourceMap.addContent(XmlService.createElement('MapType').setText("Uniform"));
        LSSourceMap.addContent(XmlService.createElement('Unit').setText(257).setAttribute("Name", "mg"));
        LSSourceMap.addContent(XmlService.createElement('MappedUnit').setText(257).setAttribute("Name", "mg"));
        LSSourceMap.addContent(XmlService.createElement('Description')
          .setText('Add ' + wellsDictionary[key][1][0][3] + " mg " + wellsDictionary[key][0][0] + wellsDictionary[key][0][1] + " to " + nameReactionPlate + " " + dosingRanges + " in mg"));   // Bug: only considers the first continuous range, string also wrong if individual vials are to be filled (A1)(A3)(A7) instead of (A1:A9)
        LSSourceMap.addContent(XmlService.createElement('Annotation').setAttribute("Null", "1"));                                                                                      // not clear whether this bug has any bearing
        LSSourceMap.addContent(XmlService.createElement('LibraryName').setText(nameReactionPlate));
        let Source = XmlService.createElement('Source');
        LSSourceMap.addContent(Source);
        Source.addContent(XmlService.createElement('Name').setText(wellsDictionary[key][0][0] + wellsDictionary[key][0][1]));
        Source.addContent(XmlService.createElement('SourceArrayName'));
        Source.addContent(XmlService.createElement('SourcePos'));
        let LSCalcParameters = XmlService.createElement('LSCalcParameters');
        LSSourceMap.addContent(LSCalcParameters);
        let LSCalcParameter = XmlService.createElement('LSCalcParameter').setAttribute("Name", "Value");
        LSCalcParameters.addContent(LSCalcParameter);
        LSCalcParameter.addContent(XmlService.createElement('Flags').setText(0));
        LSCalcParameter.addContent(XmlService.createElement('ID').setText(1));
        LSCalcParameter.addContent(XmlService.createElement('Value').setText(wellsDictionary[key][1][0][3]).setAttribute("Type", "5").setAttribute("DI", "0"));
        let Tags = XmlService.createElement('Tags').setAttribute("Name", "Tags");
        LSSourceMap.addContent(Tags);

        if (postDesignCreatorFlag) {
          Tags.addContent(XmlService.createElement('Tag'));
          Tags.addContent(XmlService.createElement('Tag').setText("SkipMap"));
        }
        let LSMapAmounts = XmlService.createElement('LSMapAmounts');
        LSSourceMap.addContent(LSMapAmounts);
        for (var well = 0; well < wellsDictionary[key][1].length; well++) {
          let LSMapAmount = XmlService.createElement('LSMapAmount');
          LSMapAmounts.addContent(LSMapAmount);
          LSMapAmount.addContent(XmlService.createElement('Row').setText(wellsDictionary[key][1][well][0]));
          LSMapAmount.addContent(XmlService.createElement('Column').setText(wellsDictionary[key][1][well][1]));
          LSMapAmount.addContent(XmlService.createElement('Value').setText(wellsDictionary[key][1][well][3]));
          LSMapAmount.addContent(XmlService.createElement('MappedValue').setText(parseFloat(wellsDictionary[key][1][well][3]).toFixed(2)));
        }
        break;

      case "Liquid":
      case "Solution": {

        //  platesOnDeck = [['reaction plate',8,12,projectName,1], ['sampling LCMS' ,8,12,projectName,2]];  

        //  locationsOnSourcePlates   plate type : [   [rowCount,colCount, maxVol], [                        [                  [                       vialVolume,            ComponentName+Level   ]]]]
        //                                         [0]   [0][0]                      [1]                      [1][0]            [1][0][0]               [1][0][0][0]         [1][0][0][1]
        //                                          Headerinfo on Plate type        all plates of this type   first plate      first row on plate      row properties      wellsDictionary key

        //  sourceTargetPlateLocation   [sourcePlateType, plate+1, row+1, dosingRanges, indivDosingRangesStrings, indivDosingRanges]

        var sourceTargetPlateLocation = retrieveLocations(key, locationsOnSourcePlates, wellsDictionary[key], filledRows, filledColumns);     //figure out the location of that compound on the source plate and on the reaction plate
        var sourcePlateName = sourceTargetPlateLocation[0] + " mL " + sourceTargetPlateLocation[1];
        let sourcePlateRow = sourceTargetPlateLocation[2];
        dosingRanges = sourceTargetPlateLocation[3];    // e.g. (D1:D12)(H1:H12) or (A1:D12)
        var indivDosingRangesStrings = sourceTargetPlateLocation[4];   // e.g. (D1)(D7)(H1)(H7) or (A1:D1)(A7:D7)    as array elements
        var indivDosingRanges = sourceTargetPlateLocation[5];                // [ [[ row, column] ... for all vials that are part of the indivDosingRangesString], next indiv dosing String
        var colCountOfPlateType = sourceTargetPlateLocation[6];           // 12 for 96, 8 for 48, 6 for 24...

        var LSArrayMap = XmlService.createElement('LSArrayMap').setAttribute("Version", "2");
        LSMaps.addContent(LSArrayMap);
        LSArrayMap.addContent(XmlService.createElement('MapType').setText("Uniform"));
        LSArrayMap.addContent(XmlService.createElement('Unit').setText(513).setAttribute("Name", "ul"));
        LSArrayMap.addContent(XmlService.createElement('MappedUnit').setText(513).setAttribute("Name", "ul"));
        LSArrayMap.addContent(XmlService.createElement('Description')
          .setText('Add ' + wellsDictionary[key][1][0][3] + " ul " + sourcePlateName + " (" + rowToLetterAssignment[sourcePlateRow] + "1:" + rowToLetterAssignment[sourcePlateRow] + "6) to " + platesOnDeck[0][0] + " " + dosingRanges + " in ul"));
        LSArrayMap.addContent(XmlService.createElement('LibraryName').setText(platesOnDeck[0][0]));
        LSArrayMap.addContent(XmlService.createElement('SourceArrayName').setText(sourcePlateName));
        let LSCalcParameters = XmlService.createElement('LSCalcParameters');
        LSArrayMap.addContent(LSCalcParameters);

        let LSCalcParameter = XmlService.createElement('LSCalcParameter').setAttribute("Name", "Value");
        LSCalcParameters.addContent(LSCalcParameter);
        LSCalcParameter.addContent(XmlService.createElement('Flags').setText(0));
        LSCalcParameter.addContent(XmlService.createElement('ID').setText(1));
        LSCalcParameter.addContent(XmlService.createElement('Value').setText(parseFloat(wellsDictionary[key][1][0][3]).toFixed(2)).setAttribute("Type", "5").setAttribute("DI", "0"));
        //LSSourceMap.addContent(XmlService.createElement('Tags').setAttribute("Name", "Tags"));
        let Tags = XmlService.createElement('Tags').setAttribute("Name", "Tags");
        LSArrayMap.addContent(Tags);
        if (postDesignCreatorFlag) {
          Tags.addContent(XmlService.createElement('Tag').setText("4Tip"));            // Tags added in Design Creator
          Tags.addContent(XmlService.createElement('Tag').setText("H6Tip"));
          Tags.addContent(XmlService.createElement('Tag').setText("LookAhead"));

        }
        LSMaps = XmlService.createElement('LSMaps');  // under this label, all the dosing maps for this row are housed
        LSArrayMap.addContent(LSMaps);

        for (var indivDosingRange = 0; indivDosingRange < indivDosingRanges.length; indivDosingRange++) {
          let LSSourceMap = XmlService.createElement('LSSourceMap'); // Write a source map for each well
          LSMaps.addContent(LSSourceMap);
          LSSourceMap.addContent(XmlService.createElement('MapType').setText("Uniform"));
          LSSourceMap.addContent(XmlService.createElement('Unit').setText(513).setAttribute("Name", "ul"));
          LSSourceMap.addContent(XmlService.createElement('MappedUnit').setText(513).setAttribute("Name", "ul"));
          LSSourceMap.addContent(XmlService.createElement('Description')
            .setText('Add ' + wellsDictionary[key][1][0][3] + " ul " + sourcePlateName + " (" + rowToLetterAssignment[sourcePlateRow] + (indivDosingRange + 1) + ") to " + platesOnDeck[0][0] + " " + indivDosingRangesStrings[indivDosingRange] + " in ul"));
          LSSourceMap.addContent(XmlService.createElement('Annotation').setAttribute("Null", "1"));
          LSSourceMap.addContent(XmlService.createElement('LibraryName').setText(platesOnDeck[0][0]));   // target library is the reaction plate
          let Source = XmlService.createElement('Source');
          LSSourceMap.addContent(Source);
          Source.addContent(XmlService.createElement('Name').setText(sourcePlateName + " (" + rowToLetterAssignment[sourcePlateRow] + (indivDosingRange + 1) + ")"));
          Source.addContent(XmlService.createElement('SourceArrayName').setText(sourcePlateName));
          Source.addContent(XmlService.createElement('SourcePos').setText(colCountOfPlateType * (sourcePlateRow - 1) + indivDosingRange + 1));     //well number on the source plate
          let LSCalcParameters = XmlService.createElement('LSCalcParameters');
          LSSourceMap.addContent(LSCalcParameters);
          let LSCalcParameter = XmlService.createElement('LSCalcParameter').setAttribute("Name", "Value");
          LSCalcParameters.addContent(LSCalcParameter);
          LSCalcParameter.addContent(XmlService.createElement('Flags').setText(0));
          LSCalcParameter.addContent(XmlService.createElement('ID').setText(1));
          LSCalcParameter.addContent(XmlService.createElement('Value').setText(parseFloat(wellsDictionary[key][1][0][3]).toFixed(2)).setAttribute("Type", "5").setAttribute("DI", "0"));
          let Tags = XmlService.createElement('Tags').setAttribute("Name", "Tags");
          LSSourceMap.addContent(Tags);

          if (postDesignCreatorFlag) {
            Tags.addContent(XmlService.createElement('Tag').setText("4Tip"));            // Tags added in Design Creator
            Tags.addContent(XmlService.createElement('Tag').setText("H6Tip"));
            Tags.addContent(XmlService.createElement('Tag').setText("LookAhead"));
          }

          let LSMapAmounts = XmlService.createElement('LSMapAmounts');
          LSSourceMap.addContent(LSMapAmounts);

          for (let well = 0; well < indivDosingRanges[indivDosingRange].length; well++) {               // Bug : correct for shift on plate due being partially filled.
            let LSMapAmount = XmlService.createElement('LSMapAmount');
            LSMapAmounts.addContent(LSMapAmount);
            LSMapAmount.addContent(XmlService.createElement('Row').setText(indivDosingRanges[indivDosingRange][well][0]));
            LSMapAmount.addContent(XmlService.createElement('Column').setText(indivDosingRanges[indivDosingRange][well][1]));
            LSMapAmount.addContent(XmlService.createElement('Value').setText(wellsDictionary[key][1][0][3]));
            LSMapAmount.addContent(XmlService.createElement('MappedValue').setText(parseFloat(wellsDictionary[key][1][0][3]).toFixed(2)));
          }
        }
        break;
      }
    }
  }

  // ** Build the layers for heating, stirring, waiting, cooling

  if (postDesignCreatorFlag) {

    var actionLayersParameters = [['StartReactionTimer', 'undefined', 0, 0, ["Processing"]],
    ['StirRate', "rpm", 3585, stirrerRpm, ["Processing"]],
    ['HeatingTemp', 'degC', 1281, temperature, ["Processing", "Wait"]],
    ['Delay', 'h', 770, reactionTime, ["Processing"]],
    ['HeatingTemp', 'degC', 1281, 21, ["Processing", "Wait"]]];    //Not clear, if the Wait tag is correct or whether it prevents progression

    for (var processingParameter = 0; processingParameter < actionLayersParameters.length; processingParameter++) {

      let LSLayer = XmlService.createElement('LSLayer');  // Create a dosing layer
      LSLayers.addContent(LSLayer);
      let layerIndexElement = XmlService.createElement('Index').setText(layerIndex);  // write the index of this layer
      layerIndex++;
      LSLayer.addContent(layerIndexElement);
      let LSMaps = XmlService.createElement('LSMaps');  // under this label, all the dosing maps for this compound are housed
      LSLayer.addContent(LSMaps);

      var LSParameterMap = XmlService.createElement('LSParameterMap');
      LSMaps.addContent(LSParameterMap);
      LSParameterMap.addContent(XmlService.createElement('MapType').setText("Uniform"));
      LSParameterMap.addContent(XmlService.createElement('Unit').setText(actionLayersParameters[processingParameter][2]).setAttribute("Name", actionLayersParameters[processingParameter][1]));
      if (actionLayersParameters[processingParameter][1] == "undefined") {
        LSParameterMap.addContent(XmlService.createElement('Description')
          .setText('Set ' + actionLayersParameters[processingParameter][0] + " to " + actionLayersParameters[processingParameter][3] + "; " + platesOnDeck[0][0] + " (A1:H12)"));

      } else {
        LSParameterMap.addContent(XmlService.createElement('Description')
          .setText('Set ' + actionLayersParameters[processingParameter][0] + " to " + actionLayersParameters[processingParameter][3] + " " + actionLayersParameters[processingParameter][1] + "; " + platesOnDeck[0][0] + " (A1:H12)"));
      }
      LSParameterMap.addContent(XmlService.createElement('Annotation').setAttribute("Null", "1"));
      LSParameterMap.addContent(XmlService.createElement('LibraryName').setText(platesOnDeck[0][0]));
      LSParameterMap.addContent(XmlService.createElement('ParameterName').setText(actionLayersParameters[processingParameter][0]));
      let LSCalcParameters = XmlService.createElement('LSCalcParameters');
      LSParameterMap.addContent(LSCalcParameters);
      let LSCalcParameter = XmlService.createElement('LSCalcParameter').setAttribute("Name", "Value");
      LSCalcParameters.addContent(LSCalcParameter);
      LSCalcParameter.addContent(XmlService.createElement('Flags').setText(0));
      LSCalcParameter.addContent(XmlService.createElement('ID').setText(0));
      LSCalcParameter.addContent(XmlService.createElement('Value').setText(actionLayersParameters[processingParameter][3]).setAttribute("Type", "5").setAttribute("DI", "0"));

      let Tags = XmlService.createElement('Tags').setAttribute("Name", "Tags");
      LSParameterMap.addContent(Tags);
      if (postDesignCreatorFlag) {
        for (var tag = 0; tag < actionLayersParameters[processingParameter][4].length; tag++) {
          Tags.addContent(XmlService.createElement('Tag').setText(actionLayersParameters[processingParameter][4][tag]));
        }
      }

      let LSMapAmounts = XmlService.createElement('LSMapAmounts');
      LSParameterMap.addContent(LSMapAmounts);

      for (let row = 0; row < 8; row++) {
        for (let well = 0; well < 12; well++) {
          let LSMapAmount = XmlService.createElement('LSMapAmount');
          LSMapAmounts.addContent(LSMapAmount);
          LSMapAmount.addContent(XmlService.createElement('Row').setText((row + 1)));
          LSMapAmount.addContent(XmlService.createElement('Column').setText((well + 1)));
          LSMapAmount.addContent(XmlService.createElement('Value').setText(actionLayersParameters[processingParameter][3]));
          LSMapAmount.addContent(XmlService.createElement('MappedValue').setText(actionLayersParameters[processingParameter][3]));
        }
      }
    }
  }

  // ** Build the layers for the transfer of diluent to the sampling plate

  for (key in samplingDictionary) {

    let LSLayer = XmlService.createElement('LSLayer');  // Create a dosing layer
    LSLayers.addContent(LSLayer);
    let layerIndexElement = XmlService.createElement('Index').setText(layerIndex);  // write the index of this layer
    layerIndex++;
    LSLayer.addContent(layerIndexElement);
    let LSMaps = XmlService.createElement('LSMaps');  // under this label, all the dosing maps for this compound are housed
    LSLayer.addContent(LSMaps);

    // samplingDictionary[key][0] = Header Information
    //            samplingDictionary[data[row][3]+data[row][4]] = [[data[row][3], data[row][4], data[row][16], data[row][18],data[row][19],data[row][20] ],[]];
    //                                                             component name   limit/level       comp name       limit/level    dose as        concentr      unit        solvent

    // samplingDictionary[key][1] = Dosing Information   [row, column, coordinate, volume/mass]
    // samplingDictionary[key][2] = Dosing Boundaries    [firstRow, firstColumn, firstCoordinate,lastRow, lastColumn, lastCoordinate]
    //  var platesOnDeck = [['reaction plate',8,12,projectName,1], ['sampling LCMS' ,8,12,projectName,2]];  var sampleVolume = 10     layerIndex

    //  locationsOnSourcePlates   plate type : [   [rowCount,colCount, maxVol], [        [               [                vialVolume, ComponentName+Level            ]]]]
    //                                         [0]   [0][0]                      [1]        [1][0]       [1][0][0]          [1][0][0][0]         [1][0][0][1]
    //                                          Headerinfo on Plate type        all plates    one plate     rows on plate      row property      samplingDictionary key

    //  sourceTargetPlateLocation   [sourcePlateType, plate+1, row+1, dosingRanges, indivDosingRangesStrings, indivDosingRanges]


    let sourceTargetPlateLocation = retrieveLocations(key, locationsOnSourcePlates, samplingDictionary[key], filledRows, filledColumns);     //figure out the location of that compound on the source plate and on the reaction plate
    let sourcePlateName = sourceTargetPlateLocation[0] + " mL " + sourceTargetPlateLocation[1];
    let sourcePlateRow = sourceTargetPlateLocation[2];
    let dosingRanges = sourceTargetPlateLocation[3];    // e.g. (D1:D12)(H1:H12) or (A1:D12)
    let indivDosingRangesStrings = sourceTargetPlateLocation[4];   // e.g. (D1)(D7)(H1)(H7) or (A1:D1)(A7:D7)    as array elements
    let indivDosingRanges = sourceTargetPlateLocation[5];                // [ [[ row, column] ... for all vials that are part of the indivDosingRangesString], next indiv dosing String
    let colCountOfPlateType = sourceTargetPlateLocation[6];           // 12 for 96, 8 for 48, 6 for 24...

    let LSArrayMap = XmlService.createElement('LSArrayMap').setAttribute("Version", "2");
    LSMaps.addContent(LSArrayMap);
    LSArrayMap.addContent(XmlService.createElement('MapType').setText("Uniform"));
    LSArrayMap.addContent(XmlService.createElement('Unit').setText(513).setAttribute("Name", "ul"));
    LSArrayMap.addContent(XmlService.createElement('MappedUnit').setText(513).setAttribute("Name", "ul"));
    LSArrayMap.addContent(XmlService.createElement('Description')
      .setText('Add ' + samplingDictionary[key][1][0][3] + " ul " + sourcePlateName + "(" + rowToLetterAssignment[sourcePlateRow] + "1:" + rowToLetterAssignment[sourcePlateRow] + "6) to " + platesOnDeck[1][0] + " " + dosingRanges + " in ul"));
    LSArrayMap.addContent(XmlService.createElement('LibraryName').setText(platesOnDeck[1][0]));
    LSArrayMap.addContent(XmlService.createElement('SourceArrayName').setText(sourcePlateName));
    let LSCalcParameters = XmlService.createElement('LSCalcParameters');
    LSArrayMap.addContent(LSCalcParameters);
    let Tags = XmlService.createElement('Tags').setAttribute("Name", "Tags");
    LSArrayMap.addContent(Tags);
    if (postDesignCreatorFlag) {
      Tags.addContent(XmlService.createElement('Tag').setText("4Tip"));
      Tags.addContent(XmlService.createElement('Tag').setText("H6Tip"));
      Tags.addContent(XmlService.createElement('Tag').setText("LookAhead"));     // those two tags don't make much sense for >500 uL diluent volume

    }
    let LSCalcParameter = XmlService.createElement('LSCalcParameter').setAttribute("Name", "Value");
    LSCalcParameters.addContent(LSCalcParameter);
    LSCalcParameter.addContent(XmlService.createElement('Flags').setText(0));
    LSCalcParameter.addContent(XmlService.createElement('ID').setText(1));
    LSCalcParameter.addContent(XmlService.createElement('Value').setText(parseFloat(samplingDictionary[key][1][0][3]).toFixed(2)).setAttribute("Type", "5").setAttribute("DI", "0"));
    //  LSSourceMap.addContent(XmlService.createElement('Tags').setAttribute("Name", "Tags"));


    LSMaps = XmlService.createElement('LSMaps');  // under this label, all the dosing maps for this row are housed
    LSArrayMap.addContent(LSMaps);

    for (let indivDosingRange = 0; indivDosingRange < indivDosingRanges.length; indivDosingRange++) {
      let LSSourceMap = XmlService.createElement('LSSourceMap'); // Write a source map for each well
      LSMaps.addContent(LSSourceMap);
      LSSourceMap.addContent(XmlService.createElement('MapType').setText("Uniform"));
      LSSourceMap.addContent(XmlService.createElement('Unit').setText(513).setAttribute("Name", "ul"));
      LSSourceMap.addContent(XmlService.createElement('MappedUnit').setText(513).setAttribute("Name", "ul"));
      LSSourceMap.addContent(XmlService.createElement('Description')
        .setText('Add ' + samplingDictionary[key][1][0][3] + " ul " + sourcePlateName + " (" + rowToLetterAssignment[sourcePlateRow] + (indivDosingRange + 1) + ") to " + platesOnDeck[1][0] + " " + indivDosingRangesStrings[indivDosingRange] + " in ul"));
      LSSourceMap.addContent(XmlService.createElement('Annotation').setAttribute("Null", "1"));
      LSSourceMap.addContent(XmlService.createElement('LibraryName').setText(platesOnDeck[1][0]));   // target library is the reaction plate
      let Source = XmlService.createElement('Source');
      LSSourceMap.addContent(Source);
      Source.addContent(XmlService.createElement('Name').setText(sourcePlateName + " (" + rowToLetterAssignment[sourcePlateRow] + (indivDosingRange + 1) + ")"));
      Source.addContent(XmlService.createElement('SourceArrayName').setText(sourcePlateName));
      Source.addContent(XmlService.createElement('SourcePos').setText(colCountOfPlateType * (sourcePlateRow - 1) + indivDosingRange + 1));     //well number on the source plate
      LSCalcParameters = XmlService.createElement('LSCalcParameters');
      LSSourceMap.addContent(LSCalcParameters);
      let LSCalcParameter = XmlService.createElement('LSCalcParameter').setAttribute("Name", "Value");
      LSCalcParameters.addContent(LSCalcParameter);
      LSCalcParameter.addContent(XmlService.createElement('Flags').setText(0));
      LSCalcParameter.addContent(XmlService.createElement('ID').setText(1));
      LSCalcParameter.addContent(XmlService.createElement('Value').setText(parseFloat(samplingDictionary[key][1][0][3]).toFixed(2)).setAttribute("Type", "5").setAttribute("DI", "0"));
      //  LSSourceMap.addContent(XmlService.createElement('Tags').setAttribute("Name", "Tags"));
      let Tags = XmlService.createElement('Tags').setAttribute("Name", "Tags");
      LSSourceMap.addContent(Tags);

      if (postDesignCreatorFlag) {
        Tags.addContent(XmlService.createElement('Tag').setText("4Tip"));
        Tags.addContent(XmlService.createElement('Tag').setText("H6Tip"));
        Tags.addContent(XmlService.createElement('Tag').setText("LookAhead"));     // those two tags don't make much sense for >500 uL diluent volume

      }

      let LSMapAmounts = XmlService.createElement('LSMapAmounts');
      LSSourceMap.addContent(LSMapAmounts);

      for (let well = 0; well < indivDosingRanges[indivDosingRange].length; well++) {
        let LSMapAmount = XmlService.createElement('LSMapAmount');
        LSMapAmounts.addContent(LSMapAmount);
        LSMapAmount.addContent(XmlService.createElement('Row').setText(indivDosingRanges[indivDosingRange][well][0]));
        LSMapAmount.addContent(XmlService.createElement('Column').setText(indivDosingRanges[indivDosingRange][well][1]));
        LSMapAmount.addContent(XmlService.createElement('Value').setText(samplingDictionary[key][1][0][3]));
        LSMapAmount.addContent(XmlService.createElement('MappedValue').setText(parseFloat(samplingDictionary[key][1][0][3]).toFixed(2)));
      }
    }
  }

  // ** Build the layers for sampling

  for (var plateRow = 1 + rowShift; plateRow < 1 + rowShift + filledRows; plateRow++) {
    let LSLayer = XmlService.createElement('LSLayer');  // Create a dosing layer for this sampling row
    LSLayers.addContent(LSLayer);
    let layerIndexElement = XmlService.createElement('Index').setText(layerIndex);  // write the index of this layer
    layerIndex++;
    LSLayer.addContent(layerIndexElement);
    let LSMaps = XmlService.createElement('LSMaps');  // under this label, all the dosing maps for this compound are housed
    LSLayer.addContent(LSMaps);
    let LSArrayMap = XmlService.createElement('LSArrayMap').setAttribute("Version", "1");
    LSMaps.addContent(LSArrayMap);
    LSArrayMap.addContent(XmlService.createElement('MapType').setText("Uniform"));
    LSArrayMap.addContent(XmlService.createElement('Unit').setText(513).setAttribute("Name", "ul"));
    LSArrayMap.addContent(XmlService.createElement('MappedUnit').setText(513).setAttribute("Name", "ul"));
    LSArrayMap.addContent(XmlService.createElement('Description')         // Bug: Needs to be adapted to plate sizes < 96
      .setText('Add ' + sampleVolume + " ul " + platesOnDeck[0][0] + " (" + rowToLetterAssignment[plateRow] + coordinateFirstWell[1] + ":" + rowToLetterAssignment[plateRow] + coordinateLastWell[1] + ") to " + platesOnDeck[1][0] + " (" + rowToLetterAssignment[plateRow] + coordinateFirstWell[1] + ":" + rowToLetterAssignment[plateRow] + coordinateLastWell[1] + ") in ul"));
    LSArrayMap.addContent(XmlService.createElement('LibraryName').setText(platesOnDeck[1][0]));
    LSArrayMap.addContent(XmlService.createElement('SourceArrayName').setText(platesOnDeck[0][0]));
    let LSCalcParameters = XmlService.createElement('LSCalcParameters');
    LSArrayMap.addContent(LSCalcParameters);
    let LSCalcParameter = XmlService.createElement('LSCalcParameter').setAttribute("Name", "Value");
    LSCalcParameters.addContent(LSCalcParameter);
    LSCalcParameter.addContent(XmlService.createElement('Flags').setText(0));
    LSCalcParameter.addContent(XmlService.createElement('ID').setText(1));
    LSCalcParameter.addContent(XmlService.createElement('Value').setText(sampleVolume).setAttribute("Type", "5").setAttribute("DI", "0"));


    let Tags = XmlService.createElement('Tags').setAttribute("Name", "Tags");
    LSArrayMap.addContent(Tags);
    if (postDesignCreatorFlag) {
      Tags.addContent(XmlService.createElement('Tag').setText("4Tip"));
      Tags.addContent(XmlService.createElement('Tag').setText("H6Tip"));
    }
    LSMaps = XmlService.createElement('LSMaps');  // under this label, all the dosing maps for this row are housed
    LSArrayMap.addContent(LSMaps);


    for (var plateColumn = 1 + columnShift; plateColumn < 1 + columnShift + filledColumns; plateColumn++) {  //  This has to be adapted for plate sizes of less than 96

      let LSSourceMap = XmlService.createElement('LSSourceMap'); // Write a source map for each well
      LSMaps.addContent(LSSourceMap);
      LSSourceMap.addContent(XmlService.createElement('MapType').setText("Uniform"));
      LSSourceMap.addContent(XmlService.createElement('Unit').setText(513).setAttribute("Name", "ul"));
      LSSourceMap.addContent(XmlService.createElement('MappedUnit').setText(513).setAttribute("Name", "ul"));
      LSSourceMap.addContent(XmlService.createElement('Description')
        .setText('Add ' + sampleVolume + " ul " + platesOnDeck[0][0] + " (" + rowToLetterAssignment[plateRow] + plateColumn + ") to " + platesOnDeck[1][0] + " (" + rowToLetterAssignment[plateRow] + plateColumn + ") in ul"));
      LSSourceMap.addContent(XmlService.createElement('Annotation').setAttribute("Null", "1"));
      LSSourceMap.addContent(XmlService.createElement('LibraryName').setText(platesOnDeck[1][0]));
      let Source = XmlService.createElement('Source');
      LSSourceMap.addContent(Source);
      Source.addContent(XmlService.createElement('Name').setText(platesOnDeck[0][0] + " (" + rowToLetterAssignment[plateRow] + plateColumn + ")"));
      Source.addContent(XmlService.createElement('SourceArrayName').setText(platesOnDeck[0][0]));
      Source.addContent(XmlService.createElement('SourcePos').setText((plateRow - 1) * platesOnDeck[0][2] + plateColumn));
      let LSCalcParameters = XmlService.createElement('LSCalcParameters');
      LSSourceMap.addContent(LSCalcParameters);
      let LSCalcParameter = XmlService.createElement('LSCalcParameter').setAttribute("Name", "Value");
      LSCalcParameters.addContent(LSCalcParameter);
      LSCalcParameter.addContent(XmlService.createElement('Flags').setText(0));
      LSCalcParameter.addContent(XmlService.createElement('ID').setText(1));
      LSCalcParameter.addContent(XmlService.createElement('Value').setText(sampleVolume).setAttribute("Type", "5").setAttribute("DI", "0"));
      //  LSSourceMap.addContent(XmlService.createElement('Tags').setAttribute("Name", "Tags"));

      let Tags = XmlService.createElement('Tags').setAttribute("Name", "Tags");
      LSSourceMap.addContent(Tags);
      if (postDesignCreatorFlag) {
        Tags.addContent(XmlService.createElement('Tag').setText("4Tip"));
        Tags.addContent(XmlService.createElement('Tag').setText("H6Tip"));
      }

      let LSMapAmounts = XmlService.createElement('LSMapAmounts');
      LSSourceMap.addContent(LSMapAmounts);

      let LSMapAmount = XmlService.createElement('LSMapAmount');
      LSMapAmounts.addContent(LSMapAmount);
      LSMapAmount.addContent(XmlService.createElement('Row').setText(plateRow));
      LSMapAmount.addContent(XmlService.createElement('Column').setText(plateColumn));
      LSMapAmount.addContent(XmlService.createElement('Value').setText(sampleVolume));
      LSMapAmount.addContent(XmlService.createElement('MappedValue').setText(sampleVolume));
    }
  }

  //These elements contain unknown information, likely related to how the data is displayed in Library Studio (window partitions, which plate is selected...)
  root.addContent(XmlService.createElement('ViewMgrData').setText('<LSViewMgr><ActiveLayer>1</ActiveLayer><ActiveMap>1</ActiveMap><SelectionMode>0</SelectionMode><ChartType>0</ChartType><ActivePlate>' + platesOnDeck[1][0] + '</ActivePlate><CompositionType>1</CompositionType><PercentType>2</PercentType><ShowWindowState>0</ShowWindowState><SizeByWhat>TotalVolume</SizeByWhat><ScaleByDesign>0</ScaleByDesign><ShapeFactor>0.</ShapeFactor><StructureType>All</StructureType><StructureColumns>0</StructureColumns><TabIndex>0</TabIndex><SplitterHoriz>0.67</SplitterHoriz><SplitterVert>0.2</SplitterVert><SplitterFavor>0.2</SplitterFavor><InvisibleSources/><LibraryViewCollections><LSLibraryViews><LibraryUID>1</LibraryUID><LSLibraryView><LibraryUID>1</LibraryUID><ActiveLayer>1</ActiveLayer><TopWindowPosition>16</TopWindowPosition><LeftWindowPosition>16</LeftWindowPosition><WindowWidth>377</WindowWidth><WindowHeight>277</WindowHeight><InvisibleSources/></LSLibraryView><LSLibraryView><LibraryUID>1</LibraryUID><ActiveLayer>1</ActiveLayer><TopWindowPosition>-6</TopWindowPosition><LeftWindowPosition>-13</LeftWindowPosition><WindowWidth>887</WindowWidth><WindowHeight>619</WindowHeight><InvisibleSources/></LSLibraryView></LSLibraryViews><LSLibraryViews><LibraryUID>9</LibraryUID><LSLibraryView><LibraryUID>9</LibraryUID><ActiveLayer>1</ActiveLayer><TopWindowPosition>16</TopWindowPosition><LeftWindowPosition>409</LeftWindowPosition><WindowWidth>377</WindowWidth><WindowHeight>277</WindowHeight><InvisibleSources/></LSLibraryView><LSLibraryView><LibraryUID>9</LibraryUID><ActiveLayer>1</ActiveLayer><TopWindowPosition>10</TopWindowPosition><LeftWindowPosition>10</LeftWindowPosition><WindowWidth>300</WindowWidth><WindowHeight>200</WindowHeight><InvisibleSources/></LSLibraryView></LSLibraryViews></LibraryViewCollections><FieldOrganizer><Fields/></FieldOrganizer></LSViewMgr><!--p-->'));
  root.addContent(XmlService.createElement('CustomData').setText('<CustomDataColl/><!--p-->'));

  var document = XmlService.createDocument(root);
  var xml = XmlService.getCompactFormat().setOmitDeclaration(true).format(document);


  return [xml, locationsOnSourcePlates];
}

// **********************************************************************************************************************



//  locationsOnSourcePlates   plate type : [   [rowCount,colCount, maxVol], [        [               [                vialVolume, ComponentName+Level            ]]]]
//                                         [0]   [0][0]                      [1]        [1][0]       [1][0][0]          [1][0][0][0]         [1][0][0][1]
//                                          Headerinfo on Plate type        all plates    one plate     rows on plate      row property      wellsDictionary key

/**
 * FileGenerator: Amends the source plates to be put on Junior
 * part of the LEA-xml generation  
 * 
 * @param {Number} overage factor >1 determining how much more liquid should be put on the source plates.
 * @param {Array} chemicalsAndMixtures Array of all chemicals and mixtures present on the plate
 * @param {Object} locationsOnSourcePlates id of the folder in which the folder is to be created.
 * @return {Object} returns the amended locationsOnSourcePlates dictionary.
 */
function amendSourcePlates(overage, chemicalsAndMixtures, locationsOnSourcePlates) {
  var coord = 9;
  var plateIndex96 = locationsOnSourcePlates["96_0.8"][1].length - 1;        // address of the last plate containing source rows (material), will be 0 for plate 1, 1 for 2 etc, 
  var plateIndex48 = locationsOnSourcePlates["48_2.0"][1].length - 1;
  var plateIndex244 = locationsOnSourcePlates["24_4.0"][1].length - 1;
  var plateIndex248 = locationsOnSourcePlates["24_8.0"][1].length - 1;

  var compoundLevel = '';
  if (chemicalsAndMixtures[12] === true || chemicalsAndMixtures[12] === false) { compoundLevel = ''; } else { compoundLevel = chemicalsAndMixtures[12]; }    //Since Level and limit share the same column, the true/false emanating from the tickbox settings can come through and be written 

  if (chemicalsAndMixtures[3] == "Solution") { coord = 10; } // volume of interest for solutions is stored in index 10, the one for liquids in slot 9


  if (overage * chemicalsAndMixtures[coord] / 6 < locationsOnSourcePlates["96_0.8"][0][2] * 1000) {  // true, if 1/6  of the total liquid volume times the overage is less than 800 uL, the maximum volume of a source vial on this type of plate

    if (locationsOnSourcePlates["96_0.8"][1][plateIndex96].length < locationsOnSourcePlates["96_0.8"][0][0]) { // true, if there are 7 or less rows filled on the plate, i.e. there's space on the plate for this liquid

      locationsOnSourcePlates["96_0.8"][1][plateIndex96].push([chemicalsAndMixtures[coord] / 6, chemicalsAndMixtures[11] + compoundLevel]);  //Add the volume per vial and the name of the solution to that plate
    } else {
      locationsOnSourcePlates["96_0.8"][1].push([]);   //If the plate was full, create a new plate
      locationsOnSourcePlates["96_0.8"][1][plateIndex96 + 1].push([chemicalsAndMixtures[coord] / 6, chemicalsAndMixtures[11] + compoundLevel]);  //Add the volume per vial and the name of the solution to the new plate
    }
  } else if (overage * chemicalsAndMixtures[coord] / 6 < locationsOnSourcePlates["48_2.0"][0][2] * 1000) {

    if (locationsOnSourcePlates["48_2.0"][1][plateIndex48].length < locationsOnSourcePlates["48_2.0"][0][0]) { // true, if there are 5 or less rows filled on the plate, i.e. there's space on the plate for this liquid
      locationsOnSourcePlates["48_2.0"][1][plateIndex48].push([chemicalsAndMixtures[coord] / 6, chemicalsAndMixtures[11] + compoundLevel]);  //Add the volume per vial and the name of the solution to that plate
    } else {
      locationsOnSourcePlates["48_2.0"][1].push([]);   //create a new plate
      locationsOnSourcePlates["48_2.0"][1][plateIndex48 + 1].push([chemicalsAndMixtures[coord] / 6, chemicalsAndMixtures[11] + compoundLevel]);  //Add the volume per vial and the name of the solution to the new plate
    }
  } else if (overage * chemicalsAndMixtures[coord] / 6 < locationsOnSourcePlates["24_4.0"][0][2] * 1000) {
    if (locationsOnSourcePlates["24_4.0"][1][plateIndex244].length < locationsOnSourcePlates["24_4.0"][0][0]) { // true, if there are 3 or less rows filled on the plate, i.e. there's space on the plate for this liquid
      locationsOnSourcePlates["24_4.0"][1][plateIndex244].push([chemicalsAndMixtures[coord] / 6, chemicalsAndMixtures[11] + compoundLevel]);  //Add the volume per vial and the name of the solution to that plate
    } else {
      locationsOnSourcePlates["24_4.0"][1].push([]);   //create a new plate
      locationsOnSourcePlates["24_4.0"][1][plateIndex244 + 1].push([chemicalsAndMixtures[coord] / 6, chemicalsAndMixtures[11] + compoundLevel]);  //Add the volume per vial and the name of the solution to the new plate
    }
  } else if (1.1 * chemicalsAndMixtures[coord] / 6 < locationsOnSourcePlates["24_8.0"][0][2] * 1000) {    // For large volumes, a small overage is sufficient, otherwise the script will fail in corner cases where the overage would lead to volumes >8 mL which are too high and not needed for large volumes 
    if (locationsOnSourcePlates["24_8.0"][1][plateIndex248].length < locationsOnSourcePlates["24_8.0"][0][0]) { // true, if there are 3 or less rows filled on the plate, i.e. there's space on the plate for this liquid
      locationsOnSourcePlates["24_8.0"][1][plateIndex248].push([chemicalsAndMixtures[coord] / 6, chemicalsAndMixtures[11] + compoundLevel]);  //Add the volume per vial and the name of the solution to that plate
    } else {
      locationsOnSourcePlates["24_8.0"][1].push([]);   //create a new plate
      locationsOnSourcePlates["24_8.0"][1][plateIndex248 + 1].push([chemicalsAndMixtures[coord] / 6, chemicalsAndMixtures[11] + compoundLevel]);  //Add the volume per vial and the name of the solution to the new plate
    }
  }
  return locationsOnSourcePlates;
}

/**
 * FileGenerator: looks at partially filled source plates and combines two partially filled plates, if combining them on the larger volume plate leads to a reduction in the overall number of plates
 * part of the LEA-xml generation  
 * 
 * @param {Object} locationsOnSourcePlates id of the folder in which the folder is to be created.
 * @return {Object} returns the optimized locationsOnSourcePlates dictionary.
 */
function optimizeSourcePlates(locationsOnSourcePlates) {

  var plateIndex96 = locationsOnSourcePlates["96_0.8"][1].length - 1;        //  address for the last plate containing source rows, will be 0 for plate 1, 1 for 2 etc,
  var plateIndex48 = locationsOnSourcePlates["48_2.0"][1].length - 1;
  var plateIndex244 = locationsOnSourcePlates["24_4.0"][1].length - 1;
  var plateIndex248 = locationsOnSourcePlates["24_8.0"][1].length - 1;
  var pushHappened = 0; // will be set to 1, if a push from 96 to 48 happened, 2 for push from 48 to 24 4.0, 3 for push from 24 4.0 to 24 8.0, this prevents material from being pushed up more than one plate size

  // ** In this cascade of if-statements, the last partially filled plate of the category in question is compared with the plate type one size higher to see whether by pushing the content of the smaller plate up, a plate can be elimanted

  if (pushHappened == 0 && locationsOnSourcePlates["96_0.8"][1][plateIndex96].length + locationsOnSourcePlates["48_2.0"][1][plateIndex48].length < 7 && locationsOnSourcePlates["96_0.8"][1][plateIndex96].length * locationsOnSourcePlates["48_2.0"][1][plateIndex48].length > 0) {   // true, if the sum of rows filled on the sending and receiving plate is not more than the row count of the receiving plate and if both the sending and receiving plate contain at least one filled row. 
    for (var plateRow = 0; plateRow < locationsOnSourcePlates["96_0.8"][1][plateIndex96].length; plateRow++) { // go through all the rows of the 96-well plate and copy them onto the 48-well plate
      var item = locationsOnSourcePlates["96_0.8"][1][plateIndex96][plateRow];
      locationsOnSourcePlates["48_2.0"][1][plateIndex48].push(item);
    }
    locationsOnSourcePlates["96_0.8"][1].pop(); //remove the last 96-well plate 
    pushHappened = 1;

  } else if (pushHappened < 1 && locationsOnSourcePlates["24_4.0"][1][plateIndex244].length + locationsOnSourcePlates["48_2.0"][1][plateIndex48].length < 5 && locationsOnSourcePlates["48_2.0"][1][plateIndex48].length * locationsOnSourcePlates["24_4.0"][1][plateIndex244].length > 0) {
    for (let plateRow = 0; plateRow < locationsOnSourcePlates["48_2.0"][1][plateIndex48].length; plateRow++) {
      let item = locationsOnSourcePlates["48_2.0"][1][plateIndex48][plateRow];
      locationsOnSourcePlates["24_4.0"][1][plateIndex244].push(item);

    }
    locationsOnSourcePlates["48_2.0"][1].pop();
    pushHappened = 2;

  } else if (pushHappened < 2 && locationsOnSourcePlates["24_4.0"][1][plateIndex244].length + locationsOnSourcePlates["24_8.0"][1][plateIndex248].length < 5 && locationsOnSourcePlates["24_8.0"][1][plateIndex248].length * locationsOnSourcePlates["24_4.0"][1][plateIndex244].length > 0) {
    for (let plateRow = 0; plateRow < locationsOnSourcePlates["24_4.0"][1][plateIndex244].length; plateRow++) {
      let item = locationsOnSourcePlates["24_4.0"][1][plateIndex244][plateRow];
      locationsOnSourcePlates["24_8.0"][1][plateIndex248].push(item);
    }
    locationsOnSourcePlates["24_4.0"][1].pop();
    pushHappened = 3;
  }
  // if a push happened from 96 to 48, a second push is possible from 244 to 248. It should also be possible recursively, but I couldn't get the locationsOnSourcePlates to be returned to the original call :( This will only be a problem, if a lot more source plate types are added.
  if (pushHappened == 1 && locationsOnSourcePlates["24_4.0"][1][plateIndex244].length + locationsOnSourcePlates["24_8.0"][1][plateIndex248].length < 5 && locationsOnSourcePlates["24_8.0"][1][plateIndex248].length * locationsOnSourcePlates["24_4.0"][1][plateIndex244].length > 0) {
    for (let plateRow = 0; plateRow < locationsOnSourcePlates["24_4.0"][1][plateIndex244].length; plateRow++) {
      let item = locationsOnSourcePlates["24_4.0"][1][plateIndex244][plateRow];
      locationsOnSourcePlates["24_8.0"][1][plateIndex248].push(item);
    }
    locationsOnSourcePlates["24_4.0"][1].pop();
    pushHappened = 3;
  }

  for (let key in locationsOnSourcePlates) {   // remove empty plates
    if (locationsOnSourcePlates[key][1].length == 0) { continue; }
    if (locationsOnSourcePlates[key][1][0].length == 0) { locationsOnSourcePlates[key][1].pop(); }
  }

  return locationsOnSourcePlates;
}


// wellsDictionary[key][0] = Header Information
// wellsDictionary[data[row][3]+data[row][4]] = [[data[row][3], data[row][4], data[row][16], data[row][18],data[row][19],data[row][20] ],[]];
//                  component name   limit/level       comp name       limit/level    dose as        concentr      unit        solvent

// wellsDictionary[key][1] = Dosing Information   [[row, column, coordinate, volume/mass]...]


// wellsDictionary[key][2] = Dosing Boundaries    [firstRow, firstColumn, firstCoordinate,lastRow, lastColumn, lastCoordinate]...


//  var platesOnDeck = [['reaction plate',8,12,projectName,1], ['sampling LCMS' ,8,12,projectName,2]];  var sampleVolume = 10     layerIndex
//  locationsOnSourcePlates   plate type : [   [rowCount,colCount, maxVol], [        [               [                vialVolume, ComponentName+Level            ]]]]
//                                         [0]   [0][0]                      [1]        [1][0]       [1][0][0]          [1][0][0][0]         [1][0][0][1]
//                                          Headerinfo on Plate type        all plates    one plate     rows on plate      row property      wellsDictionary key

/**
 * FileGenerator: Part of the Lea XML creation process. 
 * 
 * @param {String} key 
 * @param {Object} locationsOnSourcePlates 
 * @param {Object} wellsDictionaryBoundaries 
 * @param {Number} filledRows 
 * @param {Number} filledColumns 
 */
function retrieveLocations(key, locationsOnSourcePlates, wellsDictionaryBoundaries, filledRows, filledColumns) {
  // wellsDictionaryBoundaries is equivalent to wellsDictionary[key]
  var rowToLetterAssignment = { 1: "A", 2: "B", 3: "C", 4: "D", 5: "E", 6: "F", 7: "G", 8: "H" };
  var dosingRanges = "";
  var indivDosingRangesStrings = [];
  var indivDosingRanges = [];

  // In this part, the function goes through the ranges, where the individual compounds should be located on the reaction plate and constructs the peculiar strings needed for LSLayers that are associated with dosing these compounds to the reaction plate

  for (var dosingRange = 2; dosingRange < wellsDictionaryBoundaries.length; dosingRange++) {      // go through the continuous dosing range of the individual component (key = componentName + Level), starting at 2 because [0] contains header information and [1] the individual wells
    if (wellsDictionaryBoundaries[dosingRange][2] == wellsDictionaryBoundaries[dosingRange][5]) {                                              // if it's a single well (first coordinate = last coordinate), the string looks like this: (A3)
      dosingRanges = dosingRanges + "(" + wellsDictionaryBoundaries[dosingRange][2] + ")";
    }
    else { dosingRanges = dosingRanges + "(" + wellsDictionaryBoundaries[dosingRange][2] + ":" + wellsDictionaryBoundaries[dosingRange][5] + ")"; }   // if it's a range, this component of the string looks like this: (A1:D8)
  }

  /*        Examples for the outputs of these functions
  
dosingRanges (A1:A12)
indivDosingRangesStrings [(A1)(A7), (A2)(A8), (A3)(A9), (A4)(A10), (A5)(A11), (A6)(A12)]
indivDosingRanges [[[1.0, 1.0], [1.0, 7.0]], [[1.0, 2.0], [1.0, 8.0]], [[1.0, 3.0], [1.0, 9.0]], [[1.0, 4.0], [1.0, 10.0]], [[1.0, 5.0], [1.0, 11.0]], [[1.0, 6.0], [1.0, 12.0]]]

dosingRanges (A1:H6)
indivDosingRangesStrings [(A1:H1), (A2:H2), (A3:H3), (A4:H4), (A5:H5), (A6:H6)]
indivDosingRanges [[[1.0, 1.0], [2.0, 1.0], [3.0, 1.0], [4.0, 1.0], [5.0, 1.0], [6.0, 1.0], [7.0, 1.0], [8.0, 1.0]], [[1.0, 2.0], [2.0, 2.0], [3.0, 2.0], [4.0, 2.0], [5.0, 2.0], [6.0, 2.0], [7.0, 2.0], [8.0, 2.0]], [[1.0, 3.0], [2.0, 3.0], [3.0, 3.0], [4.0, 3.0], [5.0, 3.0], [6.0, 3.0], [7.0, 3.0], [8.0, 3.0]], [[1.0, 4.0], [2.0, 4.0], [3.0, 4.0], [4.0, 4.0], [5.0, 4.0], [6.0, 4.0], [7.0, 4.0], [8.0, 4.0]], [[1.0, 5.0], [2.0, 5.0], [3.0, 5.0], [4.0, 5.0], [5.0, 5.0], [6.0, 5.0], [7.0, 5.0], [8.0, 5.0]], [[1.0, 6.0], [2.0, 6.0], [3.0, 6.0], [4.0, 6.0], [5.0, 6.0], [6.0, 6.0], [7.0, 6.0], [8.0, 6.0]]]

dosingRanges (A1:E12)
indivDosingRangesStrings [(A1:E1)(A7:E7), (A2:E2)(A8:E8), (A3:E3)(A9:E9), (A4:E4)(A10:E10), (A5:E5)(A11:E11), (A6:E6)(A12:E12)]
indivDosingRanges [[[1.0, 1.0], [2.0, 1.0], [3.0, 1.0], [4.0, 1.0], [5.0, 1.0], [1.0, 7.0], [2.0, 7.0], [3.0, 7.0], [4.0, 7.0], [5.0, 7.0]], [[1.0, 2.0], [2.0, 2.0], [3.0, 2.0], [4.0, 2.0], [5.0, 2.0], [1.0, 8.0], [2.0, 8.0], [3.0, 8.0], [4.0, 8.0], [5.0, 8.0]], [[1.0, 3.0], [2.0, 3.0], [3.0, 3.0], [4.0, 3.0], [5.0, 3.0], [1.0, 9.0], [2.0, 9.0], [3.0, 9.0], [4.0, 9.0], [5.0, 9.0]], [[1.0, 4.0], [2.0, 4.0], [3.0, 4.0], [4.0, 4.0], [5.0, 4.0], [1.0, 10.0], [2.0, 10.0], [3.0, 10.0], [4.0, 10.0], [5.0, 10.0]], [[1.0, 5.0], [2.0, 5.0], [3.0, 5.0], [4.0, 5.0], [5.0, 5.0], [1.0, 11.0], [2.0, 11.0], [3.0, 11.0], [4.0, 11.0], [5.0, 11.0]], [[1.0, 6.0], [2.0, 6.0], [3.0, 6.0], [4.0, 6.0], [5.0, 6.0], [1.0, 12.0], [2.0, 12.0], [3.0, 12.0], [4.0, 12.0], [5.0, 12.0]]]

dosingRanges (A1:A12)(E1:E12)
indivDosingRangesStrings [(A1)(A7)(E1)(E7), (A2)(A8)(E2)(E8), (A3)(A9)(E3)(E9), (A4)(A10)(E4)(E10), (A5)(A11)(E5)(E11), (A6)(A12)(E6)(E12)]
indivDosingRanges [[[1.0, 1.0], [1.0, 7.0], [5.0, 1.0], [5.0, 7.0]], [[1.0, 2.0], [1.0, 8.0], [5.0, 2.0], [5.0, 8.0]], [[1.0, 3.0], [1.0, 9.0], [5.0, 3.0], [5.0, 9.0]], [[1.0, 4.0], [1.0, 10.0], [5.0, 4.0], [5.0, 10.0]], [[1.0, 5.0], [1.0, 11.0], [5.0, 5.0], [5.0, 11.0]], [[1.0, 6.0], [1.0, 12.0], [5.0, 6.0], [5.0, 12.0]]]

*/

  for (var counter = 0; counter < 6; counter++) { // there are 6 vials in one row
    indivDosingRangesStrings.push("");
    indivDosingRanges.push([]);
    for (let dosingRange = 2; dosingRange < wellsDictionaryBoundaries.length; dosingRange++) {

      var coordinateFirstWell = [wellsDictionaryBoundaries[dosingRange][0], wellsDictionaryBoundaries[dosingRange][1]];
      var coordinateLastWell = [wellsDictionaryBoundaries[dosingRange][3], wellsDictionaryBoundaries[dosingRange][4]];

      var dosingRangeHeight = wellsDictionaryBoundaries[dosingRange][3] - wellsDictionaryBoundaries[dosingRange][0]; //how many rows are in this dosing range
      for (var sixColumnBlock = 0; sixColumnBlock < Math.floor((coordinateLastWell[1] - coordinateFirstWell[1] + 1) / 6); sixColumnBlock++) {  // there can be either one or two 6-column blocks in each dosing range
        if (dosingRangeHeight == 0) {
          indivDosingRangesStrings[counter] += "(" + rowToLetterAssignment[wellsDictionaryBoundaries[dosingRange][0]] + (counter + coordinateFirstWell[1] + 6 * sixColumnBlock) + ")";
          indivDosingRanges[counter].push([wellsDictionaryBoundaries[dosingRange][0], (counter + coordinateFirstWell[1] + 6 * sixColumnBlock)]);
        } else {  // (A1:D1)(A7:D7)
          indivDosingRangesStrings[counter] += "(" + rowToLetterAssignment[wellsDictionaryBoundaries[dosingRange][0]] + (counter + coordinateFirstWell[1] + 6 * sixColumnBlock) + ":" + rowToLetterAssignment[wellsDictionaryBoundaries[dosingRange][0] + dosingRangeHeight] + (counter + coordinateFirstWell[1] + 6 * sixColumnBlock) + ")";
          for (var height = 0; height < dosingRangeHeight + 1; height++) {
            indivDosingRanges[counter].push([wellsDictionaryBoundaries[dosingRange][0] + height, (counter + coordinateFirstWell[1] + 6 * sixColumnBlock)]);
          }
        }
      }
    }
  }
  // This part of the function looks through the whole locationsOnSourcePlates dictionary in search for the key of wellsDictionary that is currently being handled  
  for (let sourcePlateType in locationsOnSourcePlates) {
    for (var plate = 0; plate < locationsOnSourcePlates[sourcePlateType][1].length; plate++) {
      for (var row = 0; row < locationsOnSourcePlates[sourcePlateType][1][plate].length; row++) {
        if (locationsOnSourcePlates[sourcePlateType][1][plate][row][1] == key) {
          return [sourcePlateType, plate + 1, row + 1, dosingRanges, indivDosingRangesStrings, indivDosingRanges, locationsOnSourcePlates[sourcePlateType][0][1]];
        }
      }
    }
  }
}

//  ***** 

var WellsDictForDb = {}; // This dictionary will be filled with the data to be written to the wells-table of the database

/**
 * FileGenerator: This is the big one! Saves the current plate and writes the input files as well as the data to the different gSheets - connected to the SavePlate Button 
 */
function savePlate() {


  // Stores the information on plate ingredients and well ingredients which will then be pushed into the respective sheets
  var plateIngredientsArray = [];
  var plateIngredientsDictionary = {};
  // plateIngredientsDictionary[data[row][3]+'_'+data[row][7]+'_'+data[row][18]+'_'+data[row][19]+'_'+data[row][20]] = [data[row][1],data[row][10], data[row][11],data[row][16],data[row][18],data[row][19],data[row][20],data[row][22],    0,         0,           0,       data[row][3]  , , data[row][4],      data[row][24]    ] // new entry to dictionary is created with component name as key and array of component role, MW, density, Dose as,  Concentration, Unit, Solvent, Solution Density and trailing zeros for mass, volume liquid, volume solution as value,
  //                            Comp name        Batch ID          Concentration     Unit               Solvent            0            1             2             3               4            5         6               7            8,         9,          10                11            12                13
  //                                                                                                                      Role              MW            density      dose as         conc          unit      solvent     sol density      mass    liquid vol      sol. vol      Comp Name       Limit/Level       to be evaporated?


  // Stores the new batches and solutions, so they can be written to the Batch db and Solutions sheets
  var newBatches = [];
  var newSolutionsArray = [];
  //Dictionaries with the new Column and Row Labels later to be written to the ColRowLabels Sheet in the database
  var colLabelsDict = {};
  var rowLabelsDict = {};
  //transformed versions of colLabelsDict and rowLabelsDict that are in the correct format for writing to the database
  var rowLabelsTableDict = {};
  var colLabelsTableDict = {};


  var wellsDictionary = {}; //contains information on each compound that is dosed for writing it into the Quantos XML file


  // quantosHeads contains information about solid starting materials that will be dosed using Quantos Chronect
  var quantosHeads = [["", "Analysis Method", "Dosing Head Tray", "Dosing Head Pos.", "Substance", "Lot ID", "Filled Quantity [mg]", "Expiration Date", "Retest Date", "Dose Limit", "Tap before Dosing?", "Intensity [%]", "Duration [s]"]]; //later filled with dosing head information to write to Quantos heads
  var solidCounter = 0;   //counts the number of newly registered solids
  // newBatchLabels contains information used for example for printing labels
  var newBatchLabels = [["Component Role", "Component ID", "Component Name", "CAS", "Roche-No.", "Molecular Formula", "Batch ID", "Producer", "Assay", "MW", "Dosing Head ID or density", "Date"]];

  // Initializes values for the limiting starting material and not limiting starting material: used only, if column 12 is used for control reactions
  var componentIdlimitingSM = 0;
  var componentIdAltSM2 = '';
  var componentIdAltLimitingSM = '';

  // dictionary that is used to temporarily store the row numbers in data of all components that are the same component and may have different batches

  var sameComponentDict = {};
  var sameOrDifferentFlag = "";

  //Initialize containers used for saving the sh
  var saveData = [[], [], [], [], []]; // This array contains sub-arrays with the different Regions of the File Generator plate that need to be saved [[Column R, preserve formulas in rows 19,20],[value, if not formula S7:U10],[value, if not formula Y2:Z123],[value, if not formula AD2:AI123],[AP2:AP123]]
  var valuesR2R50 = [];
  var valuesS7U10 = [];
  var valuesV2V6 = [];
  var valuesY2Z123 = [];
  var valuesAD2AN123 = [];
  var valuesAP2AP123 = [];

  //connect to the different sheets
  SpreadsheetApp.getActiveSpreadsheet().toast('Connecting to Sheets', 'Status', 3);
  var fileGeneratorSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FileGenerator");
  var batchDbSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Batch DB");
  var batchDbSheetDataRange = batchDbSheet.getDataRange();
  var batchDbSheetContent = batchDbSheetDataRange.getValues();
  var componentDbSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Component DB");
  var solutionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Solutions");
  var solutionsSheetDataRange = solutionsSheet.getDataRange();
  var solutionsSheetContent = solutionsSheetDataRange.getValues();
  var plateIngredientsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PlateIngredients");
  //var wellsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Wells");
  var plateBuilderHelperSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PlateBuilderHelper");
  var platesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Plates");
  var submitRequestSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Submit Request");
  var submitRequestSheetDataRange = submitRequestSheet.getDataRange().getValues();
  var sheetsWithPlateData = [platesSheet, plateIngredientsSheet];   // Relevant for the detection and overwriting of known plates

  const gSheetFileId = SpreadsheetApp.getActiveSpreadsheet().getId();

  //read the main component table
  SpreadsheetApp.getActiveSpreadsheet().toast('Reading Data', 'Status', 3);
  var data = fileGeneratorSheet.getRange("R2:AP133").getValues();

  if (data[10][0].substr(0, 24) == "Experimental Procedure: ") { // In case the Experimental Procedure Header is still present in the cell R12, remove it from the string.
    data[10][0] = data[10][0].substr(24);
  }

  //lab-specific settings from R39 down:

  //RoSL

  const toleranceMode = data[39][0];    //the allowed error in the solid weighing, minusplus meaning the error can be in both directions, zeroplus means that at least the amount of compound described should be in the vial.
  const generateJuniorFile = data[40][0]; //True or false, determines whether a Junior file should be generated
  const generateQuantosFile = data[41][0]; //True or false, determines whether Quantos input files should be generated
  const generateMasslynxSequence = data[42][0]; //True or false, determines whether sequence files for measuring samples on Waters Masslynx should be generated
  const generateChemstationSequences = data[43][0]; //True or false, determines whether sequence files for measuring samples on Agilent Chemstation should be generated
  const useFrontDoor = data[44][0]; //True or false, determines whether the balance front door is used or not. 

  // in case the broad screening checkbox in PlateBuilder is ticked (Value mirrored in FileGenerator!R15), then also read B2:M9
  var broadscreenlayout = [];
  if (data[13][0] == true) {
    broadScreenLayout = fileGeneratorSheet.getRange("B2:M9").getValues();
  }

  var formulas = fileGeneratorSheet.getRange("R2:AP123").getFormulas();

  //abort, if data is missing which is the case if R50 contains a value >0
  if (data[98][0] > 0) {
    Browser.msgBox("There is information missing to save this plate.");
    return;
  }

  if (data[99][0] != "all Reagent Types present") { // alerts the user, if one or more reagent categories present on the previous plate or reaction category isn't present on the current plate
    var popupMessage = "";
    if (data[1][0] > 1) {
      popupMessage = data[0][0] + "_" + data[1][0] + " contains at least one substance category not present on this plate. Are you sure you want to continue nevertheless?";
    } else {
      popupMessage = data[11][0] + "s contain at least one substance category not present on this plate. Are you sure you want to continue nevertheless?";
    }
    var answer = Browser.msgBox("Confirmation needed", popupMessage, Browser.Buttons.YES_NO);
    if (answer == "no") {
      return;
    }
  }
  var plateStatus = data[97][0];

  // go through the sheets that contain plate data and, if the plate exists already, erase all rows belonging to that plate
  if (data[38][0] === false && plateStatus === "Known Plate") {
    // Make sure the user really wants to overwrite the plate. 
    var response = Browser.msgBox("Confirmation needed", 'Plate ' + data[0][0] + '_' + data[1][0] + ' is known and would be overwritten, if you press yes. Are you sure you want to do that?', Browser.Buttons.YES_NO);
    if (response == "yes") {
      Logger.log('The user clicked "Yes."');
    } else {
      Logger.log('The user clicked "No".');
      return;
    }
    for (var plateDataSheet = 0; plateDataSheet < sheetsWithPlateData.length; plateDataSheet++) {

      var found = removeThenSetNewVals(sheetsWithPlateData[plateDataSheet], data[0][0] + "_" + data[1][0]);
      if (found == 0) { break; }  // The plate doesn't exist on the first sheet, so there's no point checking the others

    }
  }

  // get the number of the last row that contains data
  var batchDbLastRow = batchDbSheet.getLastRow();
  var componentDbLastRow = componentDbSheet.getLastRow();
  var solutionsLastRow = solutionsSheet.getLastRow();
  var plateIngredientsLastRow = plateIngredientsSheet.getLastRow();
  //var wellsSheetLastRow = wellsSheet.getLastRow()
  var platesSheetLastRow = platesSheet.getLastRow();
  var batchDbRowCounter = batchDbLastRow; // the first column in the batchDB sheet is a row counter used to present the user always with the five latest batches that were registered


  //Reads the components that are chosen for screening and writes them into two arrays
  var componentsInRows = plateBuilderHelperSheet.getRange("T6:V13").getValues();        //could be optimized to read only one arry and splice it into the individual component (roles) arrays
  var componentsInColumns = plateBuilderHelperSheet.getRange("T26:V37").getValues();
  var columnRoles = fileGeneratorSheet.getRange("N11:N13").getValues();
  var rowRoles = fileGeneratorSheet.getRange("O10:Q10").getValues();

  // these will hold the formatted strings to be used in the presentation and the cheatsheet
  var formattedComponentsInRows = [];
  var formattedComponentsInColumns = [];

  var alternativeReactionPartnersArray = [];
  if (componentsInColumns.length * componentsInRows.length == 96 && data[12][0] === true) { // only the case if control reactions are activated and the plate is filled. 
    alternativeReactionPartnersArray = fileGeneratorSheet.getRange("P17:R18").getValues(); // contains the starting materials and their substitutes for the control reactions
    if (alternativeReactionPartnersArray[0][2] == '' || alternativeReactionPartnersArray[1][2] == '') {
      Browser.msgBox("At least one Alternative reaction partner not specified. Please correct or deactivate control reactions.");
      return;
    }
  }

  //get the list of known batchIDs, componentIDs, componentNames, solutionIDs and molecular formulas of starting materials, (side) product(s)
  var batchIDs = batchDbSheet.getRange(2, 7, batchDbLastRow - 1).getValues();
  var componentIDs = componentDbSheet.getRange(2, 2, componentDbLastRow - 1).getValues();
  var componentMWs = componentDbSheet.getRange(2, 4, componentDbLastRow - 1).getValues();
  var componentNames = componentDbSheet.getRange(2, 3, componentDbLastRow - 1).getValues();
  var solutionIDs = solutionsSheet.getRange(2, 1, solutionsLastRow - 1).getValues();
  var molecularFormulasString = "";
  //The molecular formulas of all starting materials and products are found in AA124:AB133 which is equivalent to data[122][9] to data[131][10]
  for (let row = 122; row < 132; row++) {
    if (data[row][10].length > 0) molecularFormulasString += data[row][10] + '\t'; //molecular formula
  }
  //removes the inner brackets from these five arrays, so that the indexOf function can be used
  componentNames = [].concat.apply([], componentNames);
  componentIDs = [].concat.apply([], componentIDs);
  componentMWs = [].concat.apply([], componentMWs);
  batchIDs = [].concat.apply([], batchIDs);
  solutionIDs = [].concat.apply([], solutionIDs);

  var numberFilledRows = 0;
  if (data[13][0] == true) { // if broad screen is active
    componentsInRows = [];  //empty the array
    for (row = 0; row < broadScreenLayout.length; row++) {
      componentsInRows.push([broadScreenLayout[row][5], broadScreenLayout[row][5]]); // take column 6 of the broadscreenlayout since it's guaranteed to have the right number of rows (mostly 8)
    }

  }


  for (var arrayRow = 0; arrayRow < 8; arrayRow++) { if (componentsInRows[arrayRow][0].length > 1) { numberFilledRows++; } } // counts the number of filled rows in the array
  //removes the number of empty rows from the end of the array which is defined as the difference between the number of rows in the array and the number of filled rows
  if (numberFilledRows - componentsInRows.length < 0) { componentsInRows = componentsInRows.slice(0, numberFilledRows - componentsInRows.length); }  //only executed when there is at least 1 empty row, since slice returns an empty array, when the second argument is also 0

  //now do the same thing for the array containing the information on the filled columns
  numberFilledRows = 0;
  for (arrayRow = 0; arrayRow < 12; arrayRow++) { if (componentsInColumns[arrayRow][0].length > 0) { numberFilledRows++; } }
  if (numberFilledRows - componentsInColumns.length < 0) { componentsInColumns = componentsInColumns.slice(0, numberFilledRows - componentsInColumns.length); }

  // these will hold the formatted strings to be used in the presentation and maybe the cheatsheet
  formattedComponentsInRows = componentsInRows;
  formattedComponentsInColumns = componentsInColumns;

  //How many rows and columns are available to fill, initialized value: 96-well plate with 1.0 mL vials
  var rowsOnPlate = 8;
  var columnsOnPlate = 12;

  var vialOption = "96 wells 1 mL gold";           // describes the type of vial used on Quantos, gold = analytical sales plates or green unchained labs, black would be another option which we don't use typically. 
  switch (data[27][0]) {                           // depending on the vial volume chosen, this distinguishes between 1 and 1.2 mL vials. 
    case 1:
      vialOption = "96 wells 1 mL gold";
      break;
    case 1.2:
      vialOption = "96 wells 1.2 mL gold";
      break;
    case 2:
      if (componentsInRows.length < 7 && componentsInColumns.length < 9) {
        vialOption = "48 wells 2 mL gold";
        rowsOnPlate = 6;
        columnsOnPlate = 8;
      } else {
        vialOption = "96 wells 1 mL gold";    //Even if a vial is selected that doesn't exist for 96-well plates, it reverts to 1.0 mL vials and warns the user
        Browser.msgBox("2 mL vials are only available on a 48-well plate. Standard 1 mL vials are selected instead for the purpose of building the Quantos input file.");
      }
      break;

    default:
      vialOption = "96 wells 1 mL gold";    //Even if a vial is selected that doesn't exist for 96-well plates, it reverts to 1.0 mL vials and warns the user
      if (componentsInRows.length > 6 || componentsInColumns.length > 8) {
        Browser.msgBox("There are no " + data[27][0] + " mL vials for 96-well plates. Standard 1 mL vials are selected instead for the purpose of building the Quantos input file.");
      } break;
  }

  if (componentsInRows.length < 5 && componentsInColumns.length < 7) {    // a 24-well plate could be used unless 2 mL vials were selected in which case a 8x6 plate has to be used.
    rowsOnPlate = 4;
    columnsOnPlate = 6;

    switch (data[27][0]) {   // set the parameter for Quantos depending on the selected vial volume in cell R29
      case 1:
        vialOption = "24 wells 1 mL gold";
        break;
      case 1.2:
        vialOption = "24 wells 1.2 mL gold";
        break;
      case 2:  // for 2 mL vials only the 8x6 plate is available
        rowsOnPlate = 6;
        columnsOnPlate = 8;
        vialOption = "48 wells 2 mL gold";
        break;
      case 4:
        vialOption = "24 wells 4 mL gold";
        break;
      case 8:
        vialOption = "24 wells 8 mL gold";
        break;
    }
  } else if (componentsInColumns.length < 7 && data[27][0] < 2 && data[13][0] == false) {     // a 96-well plate is used in which only the middle 6 columns are filled, neccessary because Junior can only use 6-tip at the edge of a plate ( first column = 1 or 7). Only to be used if broad screen is inactive.
    //In LEA Studio, a plate with 6 columns and 8 rows is set up with the same pitch as a 96-well plate 

    columnsOnPlate = 6;
  }


  var quantosXmlArray = []; //later filled with dosing information to generate the Quantos csl-file


  var rowToLetterAssignment = { 1: "A", 2: "B", 3: "C", 4: "D", 5: "E", 6: "F", 7: "G", 8: "H" };
  var wellsData = [[], [], [], []]; // container used to shuffle data back and forth between savePlate and writeWellsData
  //wellsData[0] = Array later to be written to the Wells-Sheet
  //wellsData[1] = non-dosing solids, appears not to be used anywhere
  //wellsData[2] = plateIngredientsDictionary, will be used to amend the PlateIngredients Sheet and also for generating the Junior input file
  //wellsData[3] = wellsDictionary for Junior, also used in QuantosXlsx creation
  //wellsData[4] = solid-dosing information, used to generate Quantos input file



  //** Calculate the number of rows and columns that the first well should be shifted in - If a plate is not completely full, the filled wells should be centered on the plate, as stirring and sealing is better there.

  const columnShift = Math.round((columnsOnPlate - componentsInColumns.length) / 2 + 0.01);   //0.01 added to make clear that it should round up
  const rowShift = Math.round((rowsOnPlate - componentsInRows.length) / 2 - 0.01);       //0.01 subtracted to make clear that it should round down
  const firstFilledColumn = 1 + columnShift;
  const lastFilledColumn = firstFilledColumn + componentsInColumns.length;
  const firstFilledRow = 1 + rowShift;
  const lastFilledRow = firstFilledRow + componentsInRows.length;

  SpreadsheetApp.getActiveSpreadsheet().toast('Processing Data', 'Status', 3);

  var sample1LCMS = [];
  var sample2LCMS = [];
  var sample3LCMS = [];
  var sample4LCMS = [];
  var samplesAgilentC8 = ["Location\tSample Name\tMethod Name\tInj/Location\tSample Type\tInj Volume", "P1F1" + '\t' + "blank" + '\t' + "FAST-XB BEH C8-T" + '\t1\t1'];
  var samplesAgilentC18 = ["Location\tSample Name\tMethod Name\tInj/Location\tSample Type\tInj Volume", "P1F1" + '\t' + "blank" + '\t' + "FAST-XB BEH C18-T" + '\t1\t1'];

  if (generateMasslynxSequence == true || generateChemstationSequences == true) { // only generate the LCMS and Agilent lists, if either checkbox is clicked

    var agilentInjectionVolume = 0.3;

    //******** Generate the LCMS and Agilent sample list
    for (let plateColumn = firstFilledColumn; plateColumn < lastFilledColumn; plateColumn++) { // go through the components in the columns of the plate            
      // go through a second loop in which the filled rows are cycled through  
      for (var wellsInColumn = firstFilledRow; wellsInColumn < lastFilledRow; wellsInColumn++) {   // The correction factor accounts for the placement of the filled vials in the center of the plate, if <96 wells are filled.              
        // For the occupied wells amend the samplesLCMS array with the position number (A1 = 1 and then up to 96, right then down) 
        sample1LCMS.push(((wellsInColumn - 1) * 12 + plateColumn) + '\t' + data[0][0] + '_' + data[1][0] + '_1_xhyC_' + rowToLetterAssignment[wellsInColumn] + plateColumn + '\t' + molecularFormulasString);
        sample2LCMS.push(((wellsInColumn - 1) * 12 + plateColumn) + '\t' + data[0][0] + '_' + data[1][0] + '_2_xhyC_' + rowToLetterAssignment[wellsInColumn] + plateColumn + '\t' + molecularFormulasString);
        sample3LCMS.push(((wellsInColumn - 1) * 12 + plateColumn) + '\t' + data[0][0] + '_' + data[1][0] + '_3_xhyC_' + rowToLetterAssignment[wellsInColumn] + plateColumn + '\t' + molecularFormulasString);
        sample4LCMS.push(((wellsInColumn - 1) * 12 + plateColumn) + '\t' + data[0][0] + '_' + data[1][0] + '_4_xhyC_' + rowToLetterAssignment[wellsInColumn] + plateColumn + '\t' + molecularFormulasString);



        if (wellsInColumn == firstFilledRow) { // add a blank injection for every new row on the plate to clean the HPLC column
          samplesAgilentC18.push("P1F1" + '\t' + "blank" + '\t' + "FAST-XB BEH C18-T" + '\t1\t1');
          samplesAgilentC8.push("P1F1" + '\t' + "blank" + '\t' + "FAST-XB BEH C8-T" + '\t1\t1');
        } // append the samples to the arrays for both HPLC methods in use
        samplesAgilentC18.push("P2" + rowToLetterAssignment[wellsInColumn] + plateColumn + '\t' + data[0][0] + '_' + data[1][0] + '_' + rowToLetterAssignment[wellsInColumn] + plateColumn + '\t' + "FAST-XB BEH C18-T" + '\t1\t' + agilentInjectionVolume);
        samplesAgilentC8.push("P2" + rowToLetterAssignment[wellsInColumn] + plateColumn + '\t' + data[0][0] + '_' + data[1][0] + '_' + rowToLetterAssignment[wellsInColumn] + plateColumn + '\t' + "FAST-XB BEH C8-T" + '\t1\t' + agilentInjectionVolume);
      }
    }
  }


  //****************** Main Loop *************** Go through the compound table and for each compound register Batches (parent compound and solvent) as well as solvent IDs. Then figure out in which wells the component in question should be put.

  for (var row = 0; row < 122; row++) { // Below cell S123 the section with the side products of this reaction starts
    if (data[row][1] == "" && row > 69) { continue; } // skip this row if the Reaction role is empty and if the loop is beyond row 70 (R2:R70 needs to be captured...) .
    if (data[row][4] === true) { componentIdlimitingSM = data[row][2]; } //records the component ID of the limiting starting material to be written to the Plates Sheet later
    if (data[row][13] > 0 || data[row][15] > 0) {//only true, if the mass or volume is > 0. 
      //These replace '-' with an empty string in the array, making data analysis easier afterwards
      if (data[row][12] == '-') { data[row][12] = ''; } // Equivalents
      if (data[row][13] == '-') { data[row][13] = ''; } // Mass (mg)
      if (data[row][14] == '-') { data[row][14] = ''; } // Solvent (mL/g)
      if (data[row][15] == '-') { data[row][15] = ''; } // Volume (uL)
      //************ Register NEW Batch IDs
      if (batchIDs.indexOf(data[row][2] + "_" + data[row][7]) == -1 && data[row][7] != "Not registered" && data[row][1] != "Other Variable") { //The condition above is only true, if the Batch ID ( data[row][2] +"_"+ data[row][7] ) is not found in the list of registered Batch IDs and if batch information is entered and if there is a component ID present
        batchIDs.push(data[row][2] + "_" + data[row][7]); //The freshly registered batch ID is added to the list of known IDs in case a compound is registered twice (e.g. as solvent to prepare a solution and as solvent in the reaction table)
        batchDbRowCounter++;
        newBatches.push([batchDbRowCounter, data[row][2], data[row][7], data[row][8], data[row][9], data[row][3], data[row][2] + "_" + data[row][7]]);

        if ((data[row][16] == "Solid" || data[row][16] == "NonDose Solid") && data[row][1] != 'Product') {
          solidCounter++;
          newBatchLabels.push([
            (data[row][1]).toString().replace(',', ':'),
            data[row][2],
            data[row][3].toString().replace(',', ':'),
            "",
            "",
            "",
            data[row][7].toString().replace(',', ':'),
            data[row][8].toString().replace(',', ':'),
            data[row][9] * 100 + "%", data[row][10],
            data[row][2] + "@" + String(data[row][7].toString().replace(',', ':')).substring(String(data[row][7]).length - 14),
            Utilities.formatDate(new Date(), 'GMT+1', "MMM d ''yy")]);
        } else if (data[row][1] != 'Product') {
          newBatchLabels.push([(data[row][1]).toString().replace(',', ':'),
          data[row][2],
          data[row][3].toString().replace(',', ':'),
            "",
            "",
            "",
          data[row][7].toString().replace(',', ':'),
          data[row][8].toString().replace(',', ':'),
          data[row][9] * 100 + "%", data[row][10], data[row][2] + "@" + String(data[row][7].toString().replace(',', ':')).substring(String(data[row][7]).length - 14),
          Utilities.formatDate(new Date(), 'GMT+1', "MMM d ''yy")]);
        }

      } else if (data[row][7] == "Not registered" && data[row][1] != "Other Variable") { console.log(data[row][3] + " is not registered"); } //Decision to be taken what to do with unregistered batches, safest way is to catch it and stop execution of the script.

      if (data[row][16] == "Solution") { //check whether the batch of solvent specified is registered in Batch DB and register it in case it isn't           
        if (batchIDs.indexOf(componentIDs[componentNames.indexOf(data[row][20])] + '_' + data[row][21]) == -1) {
          var producer = Browser.inputBox("What's the manufacturer of " + data[row][20] + ", batch: " + data[row][21], "e.g. Sigma-Aldrich", Browser.Buttons.OK);
          batchIDs.push(componentIDs[componentNames.indexOf(data[row][20])] + "_" + data[row][21]); //The freshly registered batch ID is added to the list of known IDs in case a compound is registered twice (e.g. as solvent to prepare a solution and as solvent in the reaction table)
          batchDbRowCounter++;
          newBatches.push([batchDbRowCounter, componentIDs[componentNames.indexOf(data[row][20])], data[row][21], producer, 1, data[row][20], componentIDs[componentNames.indexOf(data[row][20])] + "_" + data[row][21]]);
        } // ***********  Register NEW Solution IDs
        if (solutionIDs.indexOf(data[row][2] + "_" + data[row][7] + "_" + data[row][20] + "_" + data[row][21] + "_" + data[row][18] + "_" + data[row][19]) == -1) { // check whether the solution ID as defined by the information in columns AJ to AM exists and register a new solution if it doesn't
          solutionIDs.push(data[row][2] + "_" + data[row][7] + "_" + data[row][20] + "_" + data[row][21] + "_" + data[row][18] + "_" + data[row][19]);
          newSolutionsArray.push([data[row][2] + "_" + data[row][7] + "_" + data[row][20] + "_" + data[row][21] + "_" + data[row][18] + "_" + data[row][19], data[row][2], data[row][7], componentIDs[componentNames.indexOf(data[row][20])], data[row][21], data[row][18], data[row][19], data[row][22]]);
        }
      }
      if (data[row][1] != 'Product') { // Register the compound as Plate and Well Ingredient, if it's not the product. This check appears to be superfluous:  data[row][1] != '' && 
        var compoundLevel; // will hold information on the compound level
        //if (row > 9) { data[row][3] = String(data[row][3]).slice(0, -1) }    //The component names in rows 12 onward (Screening components) contain an extra invisible character at the end that has to be removed for further processing, has been fixed directly in the gSheet
        if (data[row][4] != true && data[row][4] != false) { compoundLevel = data[row][4]; } else { compoundLevel = ''; }
        if (Object.keys(wellsDictionary).indexOf(data[row][3] + compoundLevel) == -1) {  //If the combintion of component and level is not present in the WellsDictionary, add an entry for it. It's unlikely but cannot be excluded that the same compound appears twice on a plate. In the even more unlikely case that once it's dosed as solution and once not, this may cause problems.
          //console.log(data[row][7]);
          wellsDictionary[data[row][3] + compoundLevel] = [[data[row][3], data[row][4], data[row][16], data[row][18], data[row][19], data[row][20], data[row][24], data[row][2] + "@" + (data[row][7]).toString().slice(-14)], []];
          //            component name   limit/level       comp name       limit/level    dose as        concentr      unit        solvent      to be evaporated?          dosehead ID
          wellsData[3] = wellsDictionary;
        }
        if (Object.keys(plateIngredientsDictionary).indexOf(data[row][3] + '_' + data[row][4] + '_' + data[row][18] + '_' + data[row][19] + '_' + data[row][20]) == -1) { //This combination of compound, batch, concentration, unit and solvent doesn't exist yet in the plateIngredientsDictionary
          plateIngredientsDictionary[data[row][3] + '_' + data[row][4] + '_' + data[row][18] + '_' + data[row][19] + '_' + data[row][20]] = [data[row][1], data[row][10], data[row][11], data[row][16], data[row][18], data[row][19], data[row][20], data[row][22], 0, 0, 0, data[row][3], data[row][4], data[row][24]]; // new entry to dictionary is created with component name as key and array of component role, MW, density, Dose as,  Concentration, Unit, Solvent, Solution Density and trailing zeros for mass, volume liquid, volume solution as value as well as the component name
          wellsData[2] = plateIngredientsDictionary;
          // plateIngredientsDictionary[data[row][3]+'_'+data[row][7]+'_'+data[row][18]+'_'+data[row][19]+'_'+data[row][20]] = [data[row][1],data[row][10], data[row][11],data[row][16],data[row][18],data[row][19],data[row][20],data[row][22],    0,         0,           0,       data[row][3]  , , data[row][4],      data[row][24]    ] // new entry to dictionary is created with component name as key and array of component role, MW, density, Dose as,  Concentration, Unit, Solvent, Solution Density and trailing zeros for mass, volume liquid, volume solution as value,
          //                            Comp name        Batch ID          Concentration     Unit               Solvent            0            1             2             3               4            5         6               7            8,         9,          10                11            12                13
          //                                                                                                                      Role              MW            density      dose as         conc          unit      solvent     sol density      mass    liquid vol      sol. vol      Comp Name       Limit/Level       to be evaporated?
          if (data[row][16] == "Solution" && Object.keys(plateIngredientsDictionary).indexOf(data[row][20] + '_' + '_' + '_' + '_') == -1) {
            // make an entry to the dictionary for the solvent unless it exists already as a pure compound            
            plateIngredientsDictionary[data[row][20] + '_' + '_' + '_' + '_'] = ["Solvents, all", componentMWs[componentNames.indexOf(data[row][20])], data[row][11], data[row][16], data[row][18], data[row][19], data[row][20], data[row][22], 0, 0, 0, data[row][20], data[row][4], data[row][24]]; // new entry to dictionary is created with component name as key and array of component role, MW, density, Dose as,  Concentration, Unit, Solvent, Solution Density and trailing zeros for mass, volume liquid, volume solution as value as well as the component name
            wellsData[2] = plateIngredientsDictionary;
          }
        } else if (data[row][16] != "Solution" && plateIngredientsDictionary[data[row][3] + '_' + data[row][4] + '_' + data[row][18] + '_' + data[row][19] + '_' + data[row][20]][3] == "Solution") { // In case the component exists already as a solvent, but then appears as a pure component later, the dictionary needs to be updated. Otherwise, the information of the pure compound is not captured which leads to problems in generating the Junior XML. 
          var tempMass = plateIngredientsDictionary[data[row][3] + '_' + data[row][4] + '_' + data[row][18] + '_' + data[row][19] + '_' + data[row][20]][8];
          var tempLiqVol = plateIngredientsDictionary[data[row][3] + '_' + data[row][4] + '_' + data[row][18] + '_' + data[row][19] + '_' + data[row][20]][9];
          var tempSolVol = plateIngredientsDictionary[data[row][3] + '_' + data[row][4] + '_' + data[row][18] + '_' + data[row][19] + '_' + data[row][20]][10];
          plateIngredientsDictionary[data[row][3] + '_' + data[row][4] + '_' + data[row][18] + '_' + data[row][19] + '_' + data[row][20]] = [data[row][1], data[row][10], data[row][11], data[row][16], data[row][18], data[row][19], data[row][20], data[row][22], tempMass, tempLiqVol, tempSolVol, data[row][3], data[row][4], data[row][24]];
          wellsData[2] = plateIngredientsDictionary;
        }
        if (data[row][16] == "Solution") { // Different fields are written to the plate ingredients sheet depending on whether the compound is dosed as a solution or not 
          plateIngredientsArray.push([data[0][0] + '_' + data[1][0],
          data[row][2] + '_' + data[row][7] + '_' + data[row][4] + '_' + data[row][1], // This should also include the solvent ID, if it's a solution
          data[row][2],
          data[row][7],
          data[row][4],
          data[row][1],
          data[row][13], //intended mass
          data[row][2] + "_" + data[row][7] + "_" + data[row][20] + "_" + data[row][21] + "_" + data[row][18] + "_" + data[row][19], // Solution ID
          data[row][24],
          data[row][23],
          data[row][12]]);
        } else {
          plateIngredientsArray.push([data[0][0] + '_' + data[1][0],
          data[row][2] + '_' + data[row][7] + '_' + data[row][4] + '_' + data[row][1],
          data[row][2],
          data[row][7],
          data[row][4],
          data[row][1],
          data[row][13],
            '',
          data[row][16],    // to be evaporated, normally empty, but now used also to signify how the component is dosed
          data[row][15],
          data[row][12]]);
        }
      }
      //*****************    Check in which wells the ingredient in question should be present *****************
      var plateColumn = 0;
      if (row < 9 && data[row][1] != 'Product') { //True, for rows 2 to 11 , where the starting materials and  fixed components are located which are put into all wells unless control reactions are active. This check appears to be superfluous: && data[row][3].length > 1
        if (componentsInColumns.length * componentsInRows.length < 96 || data[12][0] === false) { // Control reactions aren't active or the plate has less than 96 wells which means the component in question is put into all active wells      
          for (let wellsInColumn = firstFilledRow; wellsInColumn < lastFilledRow; wellsInColumn++) {
            for (plateColumn = firstFilledColumn; plateColumn < lastFilledColumn; plateColumn++) { // go through the components in the columns of the plate            
              // go through a second loop in which the ingredient that was found in the column will be written to all wells of the corresponding column on the plate which are populated (known from the length of componentsInColumns and componentsInRows)  
              // The correction factor accounts for the placement of the filled vials in the center of the plate, if <96 wells are filled.              
              wellsData = writeWellsData(quantosXmlArray, data[row], wellsData, rowToLetterAssignment[wellsInColumn] + (plateColumn), data[0][0], data[1][0], vialOption);
            }
          }
        } else { // this is entered, if there are 96 wells on the plate and control reactions are activated, i.e. column 12 is treated differently, with some wells containing alternative reaction partners or no catalyst

          for (let wellsInColumn = 1; wellsInColumn < rowsOnPlate + 1; wellsInColumn++) {  // go down each column
            for (plateColumn = 1; plateColumn < columnsOnPlate + 1; plateColumn++) {  // go to the right in each row   

              if (plateColumn < 12) { //true for the first 11 columns, same as if control reactions weren't active                                
                wellsData = writeWellsData(quantosXmlArray, data[row], wellsData, rowToLetterAssignment[wellsInColumn] + plateColumn, data[0][0], data[1][0], vialOption);
              } else if (data[row][1] != 'Starting Material') { // true for the fixed components, i.e. internal standard in column 12                
                wellsData = writeWellsData(quantosXmlArray, data[row], wellsData, rowToLetterAssignment[wellsInColumn] + plateColumn, data[0][0], data[1][0], vialOption);
              } else { // true, if the data row being read is a starting material
                switch (wellsInColumn) { // depending on the position of the column in the loop, alternative reaction partners are added to the componentsInRows array, so they can later (in data row 10 and 11)  be found and added to the wellIngredientsArray
                  case 1: //control reactions, both starting materials are substituted
                  case 8:
                    if (row == 0) { // Would otherwise be executed for every starting material, only needs to be done once and it's known that row 0 contains a starting material
                      componentsInRows[0].push(alternativeReactionPartnersArray[0][2]);
                      componentsInRows[0].push(alternativeReactionPartnersArray[1][2]);
                      componentsInRows[7].push(alternativeReactionPartnersArray[0][2]);
                      componentsInRows[7].push(alternativeReactionPartnersArray[1][2]);
                    } break;
                  case 2:
                  case 7: // no catalyst present, run through normally wrt starting materials
                    if (row == 0) {
                      componentsInRows[1].push("Georg Wuitschik fecit."); // Even though no changes are made, a dummy entry is placed into the array, since something is expected in componentsInRows[1][3] and [4]
                      componentsInRows[6].push("Georg Wuitschik fecit.");
                    }
                    wellsData = writeWellsData(quantosXmlArray, data[row], wellsData, rowToLetterAssignment[wellsInColumn] + plateColumn, data[0][0], data[1][0], vialOption);
                    break;
                  case 3:
                  case 5: // limiting starting material is substituted
                    if (data[row][4] === true) {
                      if (row == 0) {
                        componentsInRows[2].push(alternativeReactionPartnersArray[0][2]);
                        componentsInRows[2].push(alternativeReactionPartnersArray[1][0]);
                        componentsInRows[4].push(alternativeReactionPartnersArray[0][2]);
                        componentsInRows[4].push(alternativeReactionPartnersArray[1][0]);
                        componentIdAltLimitingSM = alternativeReactionPartnersArray[0][2];
                        componentIdAltSM2 = alternativeReactionPartnersArray[1][2];
                      } else { // row must be 1, since control reactions are only active for 2 starting materials and those are ordered the same way in the table. 
                        componentsInRows[2].push(alternativeReactionPartnersArray[0][0]);
                        componentsInRows[2].push(alternativeReactionPartnersArray[1][2]);
                        componentsInRows[4].push(alternativeReactionPartnersArray[0][0]);
                        componentsInRows[4].push(alternativeReactionPartnersArray[1][2]);
                      }
                    }
                    break;
                  case 4: //The other starting material is substituted
                  case 6:
                    if (data[row][4] === true) {
                      if (row == 0) {
                        componentsInRows[3].push(alternativeReactionPartnersArray[0][0]);
                        componentsInRows[3].push(alternativeReactionPartnersArray[1][2]);
                        componentsInRows[5].push(alternativeReactionPartnersArray[0][0]);
                        componentsInRows[5].push(alternativeReactionPartnersArray[1][2]);
                      } else { // row must be 1, since control reactions are only active for 2 starting materials and those are ordered the same way in the table. 
                        componentsInRows[3].push(alternativeReactionPartnersArray[0][2]);
                        componentsInRows[3].push(alternativeReactionPartnersArray[1][0]);
                        componentsInRows[5].push(alternativeReactionPartnersArray[0][2]);
                        componentsInRows[5].push(alternativeReactionPartnersArray[1][0]);
                      }
                    }
                    break;
                }// end of Switch clause and case where column 12 is written and the row in question contains a starting material
                // We are now in column 12 and working our way down the rows looking into the componentsInRow array whether in 
                if (data[row][3] == componentsInRows[wellsInColumn - 1][3] || data[row][3] == componentsInRows[wellsInColumn - 1][4]) { //writes the starting materials to the well ingredients array that aren't substituted
                  wellsData = writeWellsData(quantosXmlArray, data[row], wellsData, rowToLetterAssignment[wellsInColumn] + plateColumn, data[0][0], data[1][0], vialOption);
                }
              }
            }
          }
        } // end of else statement that contains code executed when there are 96 wells filled and control reactions box is checked
      } // ******* end of the part, where starting materials and fixed components are handled

      if (row > 9) { // true for screening components
        if (data[13][0] == true) { // if broad screen is activated
          // go through the rows of broadScreenLayout and check whether the component in question is present
          for (var plateRow = 0; plateRow < 8; plateRow++) {
            // go through each column
            for (var wellsInRow = firstFilledColumn - 1; wellsInRow < lastFilledColumn; wellsInRow++) {
              if ((data[row][5] == broadScreenLayout[plateRow][wellsInRow] && data[row][1] == rowRoles[0][0])) {
                wellsData = writeWellsData(quantosXmlArray, data[row], wellsData, rowToLetterAssignment[plateRow + 1] + (wellsInRow + 1), data[0][0], data[1][0], vialOption);
              }
            }
          }
        } else {  // entered if it's a normal screen. Now look for components in the normal rows
          for (let plateRow = 0; plateRow < componentsInRows.length; plateRow++) { // check whether the component in question is present in any of the rows on the plate and whether the reaction role is the same as in the row heading (allows to have same compound present on a plate in different categories, e.g. AcOH as additive and as solvent)            
            if ((data[row][5] == componentsInRows[plateRow][0] && data[row][1] == rowRoles[0][0]) ||
              (data[row][5] == componentsInRows[plateRow][1] && data[row][1] == rowRoles[0][1]) ||
              (data[row][5] == componentsInRows[plateRow][2] && data[row][1] == rowRoles[0][2])) { // compares the undivided Component Name (ComponentName + Level) against what is chosen for screening in the rows
              for (let wellsInRow = firstFilledColumn; wellsInRow < lastFilledColumn; wellsInRow++) { // go through a second loop in which the ingredient that was found in the row will be written to all wells of the corresponding row on the plate which are populated (known from the length of componentsInColumns and componentsInRows)  
                if (componentsInColumns.length * componentsInRows.length < 96 || data[12][0] === false || wellsInRow < 12 || (((plateRow - 1) % 5) != 0) || data[row][1].toString().indexOf("Catalyst") == -1) { // only not true for catalysts in H2 and H7, if control reactions are enabled and the plate has 96 filled wells. In that special case, a catalyst is not added, since it's a control reaction.              
                  wellsData = writeWellsData(quantosXmlArray, data[row], wellsData, rowToLetterAssignment[plateRow + firstFilledRow] + wellsInRow, data[0][0], data[1][0], vialOption);
                }
              }
            }
          }
        }
        for (let plateColumn = 0; plateColumn < componentsInColumns.length; plateColumn++) {   // check whether the component in question is present in any of the columns on the plate and whether the reaction role is the same as in the column heading (allows to have same compound present on a plate in different categories, e.g. AcOH as additive and as solvent)   
          if ((data[row][5] == componentsInColumns[plateColumn][0] && data[row][1] == columnRoles[0]) ||
            (data[row][5] == componentsInColumns[plateColumn][1] && data[row][1] == columnRoles[1]) ||
            (data[row][5] == componentsInColumns[plateColumn][2] && data[row][1] == columnRoles[2])) { // go through a second loop in which the ingredient that was found in the column will be written to all wells of the corresponding column on the plate which are populated (known from the length of componentsInColumns and componentsInRows)  
            for (let wellsInColumn = firstFilledRow; wellsInColumn < lastFilledRow; wellsInColumn++) {
              if (componentsInColumns.length * componentsInRows.length < 96 || data[12][0] != true || plateColumn < 11 || data[row][1].toString().indexOf("Catalyst") == -1 || (((wellsInColumn - 2) % 5) != 0)) { //// only not true for catalysts in H2 and H7, if control reactions are enabled and the plate has 96 filled wells                   
                wellsData = writeWellsData(quantosXmlArray, data[row], wellsData, rowToLetterAssignment[wellsInColumn] + (plateColumn + firstFilledColumn), data[0][0], data[1][0], vialOption);
              }
            }
          }
        } // end of loop through the columns
        if (row < 12 && componentsInColumns.length * componentsInRows.length == 96 && data[12][0] === true) {// this is entered, if there are 96 wells on the plate and control reactions are activated, and it only enters the clause for row = 10 and 11 in the data table which corresponds to the rows that contain the alternative reaction partners
          plateColumn = 12; // we are only interested in column 12, since this is where the alternative reaction partners will be used. 
          for (let wellsInColumn = 1; wellsInColumn < rowsOnPlate + 1; wellsInColumn++) {    //// We are working our way down the rows looking into the componentsInRow array whether we can find an alternative reaction partner in the 3rd and 4th element of each row         
            if (data[row][3] == componentsInRows[wellsInColumn - 1][3] || data[row][3] == componentsInRows[wellsInColumn - 1][4]) { // checks against the two additional columns added to the componentsInRows array added in the Switch Statement
              wellsData = writeWellsData(quantosXmlArray, data[row], wellsData, rowToLetterAssignment[wellsInColumn] + plateColumn, data[0][0], data[1][0], vialOption);
            }
          }
        }
      } // end of if clause row > 9
    } else if (data[row][1] == "Other Variable") { // if another variable like temperature or pressure is registered, an entry for the PlateIngredientsSheet needs to be generated although
      let tempOrPressure = '';
      if (!isNaN(data[row][12])) {
        tempOrPressure = data[row][12];
      }

      plateIngredientsArray.push([data[0][0] + '_' + data[1][0],
      data[row][2] + '_' + data[row][7] + '_' + data[row][4] + '_' + data[row][1],
      data[row][2],
      data[row][7],
      data[row][4],
      data[row][3],    //replace "Other Variable" with the actual name of the variable like Temperature or Pressure
      data[row][13],
        '',
        '',
      data[row][15],
        tempOrPressure
      ]);
    }
    // End of if clause checking whether a component ID is present in this row
    //  [[Column R, preserve formulas in rows 19,20],[value, if not formula S7:U10],[value, if not formula Y2:Z123],[value, if not formula AD2:AI123],[AP2:AP123]]

    //** Fill the arrays that are later converted to text and stored in the plates sheet, so that the plate can be loaded again. 

    for (var column = 0; column < data[row].length; column++) {  // Go through all columns of the row in question and put the content into different arrays depending on the column number
      switch (column) {
        case 0:
          switch (row) {
            case 17:
            case 18:
              valuesR2R50.push(['=R' + row + '&" "']);
              break;
            default:
              if (row < 49) { valuesR2R50.push([data[row][column]]); } //only push the first 50 rows into that array, as the rest is empty and since before here the compound table ended at R50
              break;
          }
          break;
        case 1:
          if (data[row][column] != '' && row > 4 && row < 9) { valuesS7U10.push(['']); } else if (row > 4 && row < 9) { valuesS7U10.push([data[row][column]]); } //formulas[row][column]
          break;
        case 2:
        case 3:
          if (formulas[row][column] != '' && row > 4 && row < 9) { valuesS7U10[row - 5].push(''); } else if (row > 4 && row < 9) { valuesS7U10[row - 5].push(data[row][column]); }
          break;
        case 4:
          if (formulas[row][column] != '' && row < 5) { valuesV2V6.push(['']); } else if (row < 5) { valuesV2V6.push([data[row][column]]); }
          break;
        case 7:
          if (data[row][column] == '') { valuesY2Z123.push(['']); } else { valuesY2Z123.push([data[row][column]]); }
          break;
        case 8:
          if (data[row][column] == '') { valuesY2Z123[row].push(''); } else { valuesY2Z123[row].push(data[row][column]); }
          break;
        case 12:
          if (formulas[row][column] != '') { valuesAD2AN123.push(['']); } else { valuesAD2AN123.push([data[row][column]]); }
          break;
        case 13:
        case 14:
        case 15:
        case 16:
          if (formulas[row][column] != '') { valuesAD2AN123[row].push(''); } else { valuesAD2AN123[row].push(data[row][column]); }
          break;
        case 17:
        case 18:
        case 19:
        case 20:
        case 21:
        case 22:
          if (data[row][column] == '') { valuesAD2AN123[row].push(''); } else { valuesAD2AN123[row].push(data[row][column]); }
          break;
        case 24:
          valuesAP2AP123.push([data[row][column]]);
          break;
      }
    }
  } //End of loop through the rows

  //go through the plateIngredientsDictionary and subtract for every liquid or solution component the total amount of volume needed from the respective values in batchDB- and solutions-sheets
  var volumeFoundFlag = 0; //if the exact combination of component ID and batch ID is found, this is set to 1
  for (var key in plateIngredientsDictionary) {
    volumeFoundFlag = 0;
    switch (plateIngredientsDictionary[key][3]) { //contains whether the component is dosed as a liquid, solid or solution
      case "Solid": //the solid inventory is automatically updated based on the residual amount of solid in the dosing heads
        break;
      case "Liquid": //   liquids are found in the batchDB
        for (row = 1; row < batchDbSheetContent.length; row++) { // go through the batchDB and check whether a liquid bottle exists for this component name and whether a volume is entered
          if (plateIngredientsDictionary[key][11] == batchDbSheetContent[row][5] && // check if the Component Name is the same
            batchDbSheetContent[row][8] != "" &&  // check if a volume is entered for this batch
            isNaN(batchDbSheetContent[row][8]) === false // also make sure it's a number 
          ) {
            volumeFoundFlag = 1;
            //console.log(batchDbSheetContent[row][5] + " was found and batch " + key.split("_")[1] + ". An amount of " + plateIngredientsDictionary[key][9] + "will be deducted from the current fill value of " +  batchDbSheetContent[row][8])
            batchDbSheetContent[row][8] = Math.round(10 * (batchDbSheetContent[row][8] - 1.2 * plateIngredientsDictionary[key][9] / 1000)) / 10; // deduct the total volume used on this plate together with 20% overage and round to 0.1 mL accuracy
            break;
          }
        }
        if (volumeFoundFlag == 0) { // if no line for this component ID containing a volume  is found, go through the sheet again and only look for a component name match
          for (row = 1; row < batchDbSheetContent.length; row++) {
            if (plateIngredientsDictionary[key][11] == batchDbSheetContent[row][5] // check only if the Component Name is the same
            ) {
              volumeFoundFlag = 1;
              //console.log(batchDbSheetContent[row][5] + " was found, but no volume information. Starting volume of 0 is assumed, please add actual volume.")
              batchDbSheetContent[row][7] = batchDbSheetContent[row][7] + batchDbSheetContent[row][8];
              batchDbSheetContent[row][8] = 0; //set it to 0 so that the following subtraction works
              batchDbSheetContent[row][8] = Math.round(batchDbSheetContent[row][8] - 1.2 * plateIngredientsDictionary[key][9] / 1000);
              break;
            }
          }

        }
        if (volumeFoundFlag == 0) { console.log(plateIngredientsDictionary[key][11] + " was not found in the BatchDB-sheet, volume needed for this plate was not deducted."); }
        break;
      case "Solution": // solutions are found in the Solutions sheet

        for (row = 1; row < solutionsSheetContent.length; row++) { // go through the batchDB and check whether a liquid bottle exists for this component name and whether a volume is entered
          if (plateIngredientsDictionary[key][11] == solutionsSheetContent[row][10] && // check if the Component Name is the same
            plateIngredientsDictionary[key][4] == solutionsSheetContent[row][5] &&   // check if the concentration matches
            plateIngredientsDictionary[key][5] == solutionsSheetContent[row][6] &&    // check if the unit matches
            plateIngredientsDictionary[key][6] == solutionsSheetContent[row][11] &&   // check if the solvent matches
            solutionsSheetContent[row][9] != "" &&  // check if a volume is entered for this batch
            isNaN(solutionsSheetContent[row][9]) === false // also make sure it's a number 
          ) {
            volumeFoundFlag = 1;
            console.log(solutionsSheetContent[row][0] + " was found and batch " + key.split("_")[1] + ". An amount of " + plateIngredientsDictionary[key][10] + "will be deducted from the current fill value of " + solutionsSheetContent[row][9]);
            solutionsSheetContent[row][9] = Math.round(solutionsSheetContent[row][9] - 1.2 * plateIngredientsDictionary[key][10] / 1000); // deduct the total volume used on this plate together with 20% overage
            break;
          }
        }
        if (volumeFoundFlag == 0) { // if no line for this component ID containing a volume  is found, go through the sheet again and only look for a component name match
          for (row = 1; row < solutionsSheetContent.length; row++) {
            if (plateIngredientsDictionary[key][11] == solutionsSheetContent[row][5] && // check only if the Component Name is the same
              plateIngredientsDictionary[key][4] == solutionsSheetContent[row][5] &&   // check if the concentration matches
              plateIngredientsDictionary[key][5] == solutionsSheetContent[row][6] &&    // check if the unit matches
              plateIngredientsDictionary[key][6] == solutionsSheetContent[row][11]    // check if the solvent matches
            ) {
              volumeFoundFlag = 1;
              //Browser.msgBox
              console.log(solutionsSheetContent[row][0] + " was found, but no volume information. Starting volume of 0 is assumed, please add actual volume.");
              solutionsSheetContent[row][8] = solutionsSheetContent[row][8] + solutionsSheetContent[row][9];
              solutionsSheetContent[row][9] = 0; //set it to 0 so that the following subtraction works
              solutionsSheetContent[row][9] = Math.round(solutionsSheetContent[row][9] - 1.2 * plateIngredientsDictionary[key][10] / 1000);
              break;
            }
          }

        }
        if (volumeFoundFlag == 0) { console.log(plateIngredientsDictionary[key][11] + " was not found in the solutions-sheet, volume needed for this plate was not deducted."); }
        break;

    }
  }

  for (row = 0; row < solutionsSheetContent.length; row++) {//remove the two last columns 
    solutionsSheetContent[row].pop();
    solutionsSheetContent[row].pop();
  }

  for (row = 10; row < 122; row++) {// go through the data once more to build the formatted column/row labels needed for the presentation; 
    //                 data[row][5] == componentsInRows[plateRow][0] && data[row][1] == rowRoles[0][0]) ||
    //                (data[row][5] == componentsInRows[plateRow][1] && data[row][1] == rowRoles[0][1]) ||
    //                (data[row][5] == componentsInRows[plateRow][2] && data[row][1] == rowRoles[0][2]
    if (data[row][1] == "") break;

    //Figure out whether on this plate, different batches of the same compound are compared with each other. In this case, the information on which batch is used is to be included in the string generated for the presenation and ColRowLabels.

    //Check, if column V ( data[row][4] ) contains an L as first letter and if data[row][6] already contains information on whether the same or different batch is used.
    //IF the same compound is used and it wasn't dealt with already, then go through the rest of data and look if at least one the other rows containing this compound is using a different batch.
    if (data[row][4].length > 0 && String(data[row][6]).substring(0, 5) != "Batch") {
      if (!(data[row][2] in sameComponentDict)) { //for every Component ID that has levels, a dictionary is created with the Component ID as key.
        sameComponentDict[data[row][2]] = {};
      }
      sameComponentDict[data[row][2]][row] = data[row][7];
      for (var secondaryRow = row + 1; secondaryRow < 122; secondaryRow++) {
        if (data[secondaryRow][2] == data[row][2]) {
          sameComponentDict[data[row][2]][secondaryRow] = data[secondaryRow][7]; // set the name of the batch as value of the row in which the compound is found in data if the component ID matches
        }
      }


      //Set data[row][6] for all rows of that compound to "Batch: same" or "Batch: different" depending on what is the case. 
      if (Object.values(sameComponentDict[data[row][2]]).every((val, i, arr) => val === arr[0])) {   //inspired by: https://stackoverflow.com/questions/14832603/check-if-all-values-of-array-are-equal 
        sameOrDifferentFlag = "same";
      } else {
        sameOrDifferentFlag = "different";
      }
      for (let key in sameComponentDict[data[row][2]]) {
        data[key][6] = "Batch: " + sameOrDifferentFlag;
      }
    }


    for (var rowComponentsRow = 0; rowComponentsRow < componentsInRows.length; rowComponentsRow++) {// go through componentsInRows and check where the compound is present
      if (!(rowToLetterAssignment[firstFilledRow + rowComponentsRow] in rowLabelsDict)) {// generate a new entry in the dictionary for this row unless it exists already
        rowLabelsDict[rowToLetterAssignment[firstFilledRow + rowComponentsRow]] = {};
        rowLabelsDict[rowToLetterAssignment[firstFilledRow + rowComponentsRow]]["RowLabel"] = rowToLetterAssignment[firstFilledRow + rowComponentsRow] + ": ";    // i.e. A: 
      }

      for (var rowComponentsColumn = 0; rowComponentsColumn < componentsInRows[0].length; rowComponentsColumn++) {
        if (data[row][5] == componentsInRows[rowComponentsRow][rowComponentsColumn] && data[row][1] == rowRoles[0][rowComponentsColumn]) { //if Component name and Type match replace the string with a formatted version
          // generate the formatted string and write it into the corresponding cell
          formattedComponentsInRows[rowComponentsRow][rowComponentsColumn] = generateFormattedPlateIngredient(data[row]); //function found in ColRowLabels
          rowLabelsDict[rowToLetterAssignment[firstFilledRow + rowComponentsRow]]["Layer_" + (rowComponentsColumn + 1) + "_PlateIngredientId"] = data[row][2] + "_" + data[row][7] + "_" + data[row][4] + "_" + data[row][1];
          rowLabelsDict[rowToLetterAssignment[firstFilledRow + rowComponentsRow]]["RowLabel"] += formattedComponentsInRows[rowComponentsRow][rowComponentsColumn].replace(/'/g, "''") + "/ "; //some compound names contain single quotes which need to be converted to double single quotes for writing to the SQL-DB
        }
      }
    }


    for (let rowComponentsColumn = 0; rowComponentsColumn < componentsInColumns.length; rowComponentsColumn++) {// go through componentsInColumns and check where the compound is present
      if (!(firstFilledColumn + rowComponentsColumn in colLabelsDict)) {// generate a new entry in the dictionary for this row unless it exists already
        colLabelsDict[firstFilledColumn + rowComponentsColumn] = {};
        colLabelsDict[firstFilledColumn + rowComponentsColumn]["ColumnLabel"] = firstFilledColumn + rowComponentsColumn + ": ";    // i.e. 1: 
      }

      for (var colComponentsColumn = 0; colComponentsColumn < componentsInColumns[0].length; colComponentsColumn++) {
        if (data[row][5] == componentsInColumns[rowComponentsColumn][colComponentsColumn] && data[row][1] == columnRoles[colComponentsColumn]) { //if Component name and Type match replace the string with a formatted version
          // generate the formatted string and write it into the corresponding cell
          formattedComponentsInColumns[rowComponentsColumn][colComponentsColumn] = generateFormattedPlateIngredient(data[row]); //function found in ColRowLabels
          colLabelsDict[firstFilledColumn + rowComponentsColumn]["Layer_" + (colComponentsColumn + 1) + "_PlateIngredientId"] = data[row][2] + "_" + data[row][7] + "_" + data[row][4] + "_" + data[row][1];
          colLabelsDict[firstFilledColumn + rowComponentsColumn]["ColumnLabel"] += formattedComponentsInColumns[rowComponentsColumn][colComponentsColumn].replace(/'/g, "''") + "/ "; //some compound names contain single quotes which need to be converted to double single quotes for writing to the SQL-DB
        }
      }
    }
  }

  // Now that the formatted strings are generated for each variable component on the plate, go through both dictionaries and bring them into the format that can be sent to the database
  var counter = 1;
  for (let key in colLabelsDict) {
    colLabelsTableDict[counter] = {
      ID: {
        ELN_ID: data[0][0],
        PLATENUMBER: data[1][0],
        PlateColumn: key
      },
      DATA: colLabelsDict[key]

    };
    counter++;
  }
  counter = 1;

  for (let key in rowLabelsDict) {
    rowLabelsTableDict[counter] = {
      ID: {
        ELN_ID: data[0][0],
        PLATENUMBER: data[1][0],
        PlateRow: key
      },
      DATA: rowLabelsDict[key]

    };
    counter++;
  }
  /*var dictdata = {
      '1': {
        ID: { Coordinate: 'C1', ELN_ID: "ELN032036-303", PLATENUMBER: 22, Component_ID: 999, Batch_ID: "adfadfe", ComponentRole: "Acids, all" },
        DATA: { PlateID: 'ELN029554-013_1', PlateIngredientID: '143444_RI90520202_false_Starting Material', ActualMass: 21.44, limSM_or_Level: "false", DosingTimestamp: 'adfadfaf' }
      } 
    };*/


  // ***********************    Write the data
  SpreadsheetApp.getActiveSpreadsheet().toast('Writing Data', 'Status', 3);
  var date = new Date();
  var user = "wuitschg"; //Browser.inputBox("Please provide your user name");
  const projectName = data[6][0];
  const stepName = data[7][0];
  const reactionType = data[11][0];
  const platePurpose = data[9][0] + " Procedure: " + data[10][0];
  const diluentVolume = data[25][0];
  const sampleVolume = data[26][0];

  // ** Write the input files for LCMS, Quantos and Junior

  // var quantosCSVString = wellsData[1].join( "\r\n");   // This is a relic of the time when the Quantos input file was a csv. 

  //deprecated, now all data is sorted into folders according to file-type and not organized by ELN anymore.
  //var folderID = createFolder(globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["PROJECTdATAfOLDERiD"], data[0][0] + "_" + data[5][5] + "_" + data[6][5] + "_" + data[7][5]) //create a new folder with the ELN-ID as name on gDrive under HTS Docs > Robot Input Files unless it exists already (which it does in this case since it's generated when a new reaction is registered in Submit Request)

  var lcmsFolder = DriveApp.getFolderById(globalVariableDict[gSheetFileId]["LCMSfOLDERiD"]);
  var quantosFolder = DriveApp.getFolderById(globalVariableDict[gSheetFileId]["QUANTOSfOLDERiD"]);
  var juniorFolder = DriveApp.getFolderById(globalVariableDict[gSheetFileId]["JUNIORfOLDERiD"]);
  var agilentFolder = DriveApp.getFolderById(globalVariableDict[gSheetFileId]["AGILENTfOLDERiD"]);
  //var cheatsheetsFolder = DriveApp.getFolderById(globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["CHEATSHEETSfOLDERiD"] );

  //var quantosCSVFile = quantosFolder.createFile(data[0][0] + "_" + data[1][0] + ".csv", quantosCSVString) // writes the csv input file for Quantos
  if (rowsOnPlate != 6) { generateQuantosXlsx(wellsData[3], globalVariableDict[gSheetFileId]["QUANTOSfOLDERiD"], data[0][0] + "_" + data[1][0] + "_Quantos", rowsOnPlate, columnsOnPlate); }

  if (generateMasslynxSequence == true) {   //only write the LCMS-sequences, if the corresponding checkbox is clicked. 

    var sample1LCMSstring = sample1LCMS.join("\r\n");
    var sample2LCMSstring = sample2LCMS.join("\r\n");
    var sample3LCMSstring = sample3LCMS.join("\r\n");
    var sample4LCMSstring = sample4LCMS.join("\r\n");

    var sample1LCMSfile = lcmsFolder.createFile("LCMS " + data[0][0] + " Plate " + data[1][0] + " IPC 1.txt", sample1LCMSstring); // writes the input file for the LCMS
    var sample2LCMSfile = lcmsFolder.createFile("LCMS " + data[0][0] + " Plate " + data[1][0] + " IPC  2.txt", sample2LCMSstring); // writes the input file for the LCMS
    var sample3LCMSfile = lcmsFolder.createFile("LCMS " + data[0][0] + " Plate " + data[1][0] + " IPC   3.txt", sample3LCMSstring); // writes the input file for the LCMS
    var sample4LCMSfile = lcmsFolder.createFile("LCMS " + data[0][0] + " Plate " + data[1][0] + " IPC    4.txt", sample4LCMSstring); // writes the input file for the LCMS
  }
  if (generateChemstationSequences == true) {//only write the Chemstation-sequences, if the corresponding checkbox is clicked.
    var samplesAgilentC18string = samplesAgilentC18.join("\r\n");
    var samplesAgilentC8string = samplesAgilentC8.join("\r\n");
    var samplesAgilentC18file = agilentFolder.createFile("Agilent C18 " + data[0][0] + "_" + data[1][0] + ".txt", samplesAgilentC18string); // writes the input file for the HPLC using the C18 method
    var samplesAgilentC8file = agilentFolder.createFile("Agilent C8 " + data[0][0] + "_" + data[1][0] + ".txt", samplesAgilentC8string);    // writes the input file for the HPLC using the C8 method
  }
  var LeaXML = [];

  if (generateJuniorFile == true) { // only generate a Lea file if the corresponding checkbox in R42 is checked (normally it isn't since this function is rarely used and needs quite some time to run)

    if (data[12][0] === true || rowsOnPlate == 6) { // The corresponding liquid dosing layers will not be compatible with the 6Tip and thus the dosing has to be performed manually.
      Browser.msgBox("If either starting material or alternative starting material is a liquid or solution, the corresponding dosing steps won't work in LEA Studio and thus the dosings have to be performed manually. The corresponding dosing steps must be set to 'Skip' in Design Creator!");
    } else {
      try {
        LeaXML = createLeaXml(wellsData[2], wellsData[3], projectName, stepName, reactionType, platePurpose, diluentVolume, sampleVolume, data[0][0], user, data[1][0], componentsInRows.length, componentsInColumns.length, data[29][0], data[30][0], data[31][0], data[36][0], data[37][0], rowShift, columnShift, rowsOnPlate, columnsOnPlate);
      } catch (err) {
        SpreadsheetApp.getActiveSpreadsheet().toast('Error ' + err + ' triggered creating LEA input file, proceeding with rest of script without writing File', 'Error', 3);
        LeaXML = [];
      } finally {

        if (LeaXML.length > 0) { // If the wrong number of columns is present, an empty string is returned and no files should be written
          var juniorXmlFile = juniorFolder.createFile('Junior xml ' + data[0][0] + "_" + data[1][0] + ".xml", LeaXML[0]);              // only used for debugging purposes
          if (data[37][0] === false) {
            var juniorLsrFile = juniorFolder.createFile('Junior lsr ' + data[0][0] + "_" + data[1][0] + ".lsr", LeaXML[0]);
          } else {
            let juniorLsrFile = juniorFolder.createFile('Junior lsr incl processing ' + data[0][0] + "_" + data[1][0] + ".lsr", LeaXML[0]);
          }

        }
      }
    }
  }

  var cheatSheetFileUrl = backupPlate(globalVariableDict[gSheetFileId]["CHEATSHEETSfOLDERiD"], data[0][0] + "_Plates", data[1][0], LeaXML[1], data[36][0], data);   //puts a copy of the FileGenerator Sheet into a separate Google Sheet into the same folder containing a copy of the FileGenerator sheet including all the source plates used
  //LeaXML[1] contains the locationsOnSourcePlates dictionary which allows to put layout and content of the source plates into the backup copy of the FileGenerator sheet

  // Add the Cheatsheet Url to the Submit-Request Slide:

  for (row = 1; row < submitRequestSheetDataRange.length; row++) {
    if (submitRequestSheetDataRange[row][0] == data[0][0]) {
      submitRequestSheet.getRange(row + 1, 11).setValue(cheatSheetFileUrl); // row in array is 0-indexed, row in getRange is 1-indexed
      break;
    }
  }

  //Amend the presentation
  var presentationFileId = createPresentation(globalVariableDict[gSheetFileId]["PRESENTATIONfOLDERiD"], data[0][0] + " Results");
  try {
    fillPresentationTemplate({}, data, cheatSheetFileUrl, presentationFileId, formattedComponentsInColumns, formattedComponentsInRows);
  } catch (error) {

    webhookChatMessage(data[0][0], data[5][5], data[7][5], 'Error amending presentation, check if a gslide presentation according to the new template exists for this ELN-ID: ' + error, data[1][0], cheatSheetFileUrl, "https://docs.google.com/open?id=" + presentationFileId);
    //SpreadsheetApp.getActiveSpreadsheet().toast('Error amending presentation, check if a gslide presentation according to the new template exists for this ELN-ID: ' + error, 'Status');
  }

  var variableFileNameComponent = data[0][0] + "_" + data[1][0] + "_";
  if (newBatchLabels.length > 1) { // generate a labels file for new batches entered for components added in FileGenerator. 
    var newBatchLabelsstring = newBatchLabels.join("\r\n");
    for (var item = 1; item < newBatchLabels.length; item++) {
      variableFileNameComponent = variableFileNameComponent + "_" + newBatchLabels[item][2];
    }
    var pTouchFolder = DriveApp.getFolderById(globalVariableDict[gSheetFileId]["PtOUCHlABELSfOLDERiD"]);
    var newBatchLabelsfile = pTouchFolder.createFile(variableFileNameComponent + ".csv", newBatchLabelsstring); // writes the Batch Label csv for P-Touch
  }

  if (generateQuantosFile == true) {   // Only generate Quantos input files, if the corresponding box in R43 is ticked. 
    if (quantosHeads.length > 1) {  // only write the xml for new heads, if there are solids in the list of compounds to be registered. 
      var quantosHeadsXML = createQuantosXml(quantosHeads);
      var quantosHeadsFile = quantosFolder.createFile('Quantos solids ' + variableFileNameComponent + ".csl", quantosHeadsXML);     // contains all solids for writing dosing heads
    }
    quantosXmlArray = wellsData[4];
    // Add the two header lines. 
    var preDoseTapping = 6;      // length in seconds of the predose tapping
    var preDoseStrength = 40;           // strength in % of maximum tapping strength
    var useSideDoors = "False";
    quantosXmlArray.unshift([1, "C:\\Users\\Public\\Documents\\Chronos\\Methods\\Set Config.cam", vialOption, preDoseTapping, preDoseStrength, toleranceMode, "Quantos", useFrontDoor, useSideDoors, "", "", "", "", "", "", "", ""]);
    quantosXmlArray.unshift(["", "Analysis Method", "Dosing Tray Type", "PreDose Tap Duration [s]", "PreDose Tap Intensity [%]", "Tolerance Mode", "Device", "Use Front Door?", "Use Side Doors?", "Substance Name", "Dosing Vial Tray", "Dosing Vial Pos.", "Dosing Vial Pos. [Axx]", "Amount [mg]", "Tolerance [%]", "Sample ID", "Comment"]);
    var quantosXML = createQuantosXml(quantosXmlArray);
    var quantosXmlFile = quantosFolder.createFile(data[0][0] + "_" + data[1][0] + ".xml", quantosXML);     // only used for debugging purposes (can be opened in Excel, but only differs in the file ending)
    var quantosCslFile = quantosFolder.createFile(data[0][0] + "_" + data[1][0] + ".csl", quantosXML);
  }
  //Write the new batchDB and solutions sheets with volumes of solvents and solutions needed for this plate subtracted from the available stock
  solutionsSheet.getRange(1, 1, solutionsSheetContent.length, solutionsSheetContent[0].length).setValues(solutionsSheetContent);
  batchDbSheet.getRange(1, 1, batchDbSheetContent.length, batchDbSheetContent[0].length).setValues(batchDbSheetContent);

  // ** Write wellsData, plateData and plateIngredients to the respective sheets

  if (data[38][0] === true) { return; }   // if the tickbox is selected that only the robot files should be written, then the script is stopped at this point and no data is written to the plate ingredients, wells, plates, batch and solution sheets


  // write the information in WellsDictForDb to the wells table in the cloud database:

  var connector = new mssql_jdbc_api(   // connect to the database

    globalVariableDict[gSheetFileId]["DBsERVERiP"],
    globalVariableDict[gSheetFileId]["DBpORT"],
    globalVariableDict[gSheetFileId]["DBnAME"],
    globalVariableDict[gSheetFileId]["DBuSERNAME"],
    globalVariableDict[gSheetFileId]["DBpASSWORD"]);

  if (data[38][0] === false && plateStatus === "Known Plate") {  // If the plate already exists, then delete all entries for this plate from the table.
    connector.execute("DELETE FROM " + globalVariableDict[gSheetFileId]["WELLStABLEnAME"] + " where ELN_ID ='" + data[0][0] + "' and PLATENUMBER = " + data[1][0]);
    connector.execute("DELETE FROM " + globalVariableDict[gSheetFileId]["COLlABELStABLEnAME"] + " where ELN_ID ='" + data[0][0] + "' and PLATENUMBER = " + data[1][0]);
    connector.execute("DELETE FROM " + globalVariableDict[gSheetFileId]["ROWlABELStABLEnAME"] + " where ELN_ID ='" + data[0][0] + "' and PLATENUMBER = " + data[1][0]);
  }





  if (newSolutionsArray.length > 0) {  // If there are new Solutions defined on the sheet, write them to the Solutions Sheet    
    solutionsSheet.getRange(solutionsLastRow + 1, 1, newSolutionsArray.length, newSolutionsArray[0].length).setValues(newSolutionsArray);
  }

  if (newBatches.length > 0) {
    batchDbSheet.getRange(batchDbLastRow + 1, 1, newBatches.length, newBatches[0].length).setValues(newBatches);
  }


  savePlateDesign("Plates");
  var plateData = [[data[0][0] + "_" + data[1][0], data[0][0], data[1][0], date, user, componentIdlimitingSM, componentIdAltLimitingSM, componentIdAltSM2, data[9][0] + "; Procedure: " + data[10][0], data[29][0], data[30][0], data[31][0], "", arrayToText(valuesR2R50), arrayToText(valuesS7U10), arrayToText(valuesV2V6), arrayToText(valuesY2Z123), arrayToText(valuesAD2AN123), arrayToText(valuesAP2AP123)]];
  platesSheet.getRange(platesSheetLastRow + 1, 1, 1, 19).setValues(plateData);
  plateIngredientsSheet.getRange(plateIngredientsLastRow + 1, 1, plateIngredientsArray.length, 11).setValues(plateIngredientsArray);


  if (Object.keys(colLabelsTableDict).length > 0) { // only write if there are lines to be written (almost always the case)

    connector.insertDataDictionary(globalVariableDict[gSheetFileId]["COLlABELStABLEnAME"], colLabelsTableDict); // write the new lines
  }

  if (Object.keys(rowLabelsTableDict).length > 0) { // only write if there are lines to be written (almost always the case)

    connector.insertDataDictionary(globalVariableDict[gSheetFileId]["ROWlABELStABLEnAME"], rowLabelsTableDict); // write the new lines
  }

  if (Object.keys(WellsDictForDb).length > 0) { // only write if there are lines to be written (almost always the case)

    connector.insertDataDictionary(globalVariableDict[gSheetFileId]["WELLStABLEnAME"], WellsDictForDb); // write the new lines
  }
  connector.disconnect();
  SpreadsheetApp.getActiveSpreadsheet().toast('Finished', 'Status');
  SpreadsheetApp.flush(); // should help if people forget to push the button that follows. 
  //function in Submit Request.gs: send chat message 
  webhookChatMessage(data[0][0], data[5][5], data[7][5], data[6][5], data[1][0], cheatSheetFileUrl, "https://docs.google.com/open?id=" + presentationFileId);


}

//**********************

/**
 * FileGenerator: Load a plate defined by ELN-ID and platenumber and writes it back to PlateBuilder and FileGenerator. Connected to the Load Plate button. 
 */
function loadPlate() {

  //connect to the different sheets 
  var fileGeneratorSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FileGenerator");
  var platesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Plates");
  var platesSheetLastRow = platesSheet.getLastRow();

  var selectedPlate = fileGeneratorSheet.getRange("R2:R3").getValues(); //contains the ELN ID and the plate number

  var plateIDs = platesSheet.getRange(2, 1, platesSheetLastRow).getValues();   //list of all known plate IDs
  plateIDs = [].concat.apply([], plateIDs); //removes the inner brackets, so indexOf function can be used on the array

  //find out which plate should be loaded

  var rowNumSelectedPlate = plateIDs.indexOf(selectedPlate[0] + '_' + selectedPlate[1]) + 2;  //looks in the list of plate IDs and returns the index where the Plate ID in question is located (2 is added to bring in phase with row count in the Platesheet)

  if (rowNumSelectedPlate == 1) { //true if selected plate is not found on the plates sheet
    Browser.msgBox(selectedPlate[0] + '_' + selectedPlate[1] + " was not found. Use the drop-down in cell R3 to select an existing plate.");
    return;
  }

  //load corresponding plate design to set the plate builder sheet, the parameter "Plates" tells the function in which columns the information is stored, since the same function is also used to load standard plate designs
  loadPlateDesign("Plates", rowNumSelectedPlate);
  SpreadsheetApp.flush();   // appears to be necessary for the lookup of Molecular Weights in column AB to work properly.
  //reset fileGenerator sheet, so that all formulas that may have been overwritten by the user are restored. 
  newFileGeneratorSheetReset();

  // load plate data

  var plateData = platesSheet.getRange(rowNumSelectedPlate, 14, 1, 6).getValues(); // get the data that was changed by the user specifically for this plate after ELN-ID and plate design is loaded (e.g. equivalents, control reactions, dosing as solution...)
  var colApData = textToArray(plateData[0][5]);
  var contentR2R50 = textToArray(plateData[0][0]);


  fileGeneratorSheet.getRange(2, 18, contentR2R50.length, 1).setValues(contentR2R50);    //contains all the settings
  var S7U10data = textToArray(plateData[0][1]);
  for (var col = 0; col < S7U10data[0].length; col++) { //cycles through the array and only writes an element to the respective posiiton in S7:U10, if the element is not '' (prevents overwriting of formulas)
    for (var row = 0; row < S7U10data.length; row++) {
      if (S7U10data[row][col] != '') { fileGeneratorSheet.getRange(row + 7, col + 19).setValue(S7U10data[row][col]); }
    }
  }
  fileGeneratorSheet.getRange("R8:R9").clearContent();  // clearing and then flushing appears to be neccessary - otherwise a #Ref error may appear in the sheet (probably as a result of gScript bundling operations together)
  fileGeneratorSheet.getRange("R13:R15").clearContent();
  SpreadsheetApp.flush();
  fileGeneratorSheet.getRange("R8:R9").setFormulas([['=iferror(VLOOKUP(FileGenerator!R2,\'HTE-Requests\'!A2:D,2,false),"not found")'], ['=iferror(VLOOKUP(FileGenerator!R2,\'HTE-Requests\'!A2:D,3,false),"not found")']]);   // those two cells need to contain formulas, as they're used to retrieve project name and step name from the HTE-requests sheet
  fileGeneratorSheet.getRange("R13:R15").setFormulas([['=VLOOKUP(R2,\'HTE-Requests\'!A2:D,4,0)'], ['false'], ['=PlateBuilder!C16']]);  // lookup of reaction type, checkbox and broad screen checkbox value from PlateBuilder
  var savedBatchIdsAndProducers = textToArray(plateData[0][3]);
  fileGeneratorSheet.getRange("V2:V6").setValues(textToArray(plateData[0][2])); //Limit checkboxes
  fileGeneratorSheet.getRange(2, 25, savedBatchIdsAndProducers.length, savedBatchIdsAndProducers[0].length).setValues(savedBatchIdsAndProducers); //Batch Ids and Producers, has to be of variable length due to the change from 48 max components to 121

  var AD2AN123data = textToArray(plateData[0][4]);

  for (col = 0; col < AD2AN123data[0].length; col++) { //cycles through the array and only writes an element to the respective posiiton in AD2:AN123, if the element is not '' (prevents overwriting of formulas)
    for (let row = 0; row < AD2AN123data.length; row++) {
      if (AD2AN123data[row][col] != '') { fileGeneratorSheet.getRange(row + 2, col + 30).setValue(AD2AN123data[row][col]); }
    }
  }
  fileGeneratorSheet.getRange(2, 42, colApData.length, colApData[0].length).setValues(colApData); //sets the "to be evaporated" checkboxes to the right values
  fileGeneratorSheet.getRange("R50").clearContent();
  SpreadsheetApp.flush();
  fileGeneratorSheet.getRange("R50").setFormula('=(FileGenerator!R2="")+(FileGenerator!R8="not found")+(FileGenerator!R9="not found")+(FileGenerator!R11="")+(FileGenerator!R13="not found")+(FileGenerator!R14=TRUE)*((FileGenerator!R17="")+(FileGenerator!R18=""))+(FileGenerator!AE2="Enter Mass")+(FileGenerator!AE3="Enter Mass")+(FileGenerator!AE4="Enter Mass")+(FileGenerator!AE5="Enter Mass")+(FileGenerator!AD2="Enter value")+(FileGenerator!AD3="Enter value")+(FileGenerator!AD4="Enter value")+(FileGenerator!AD5="Enter value")+sum(FileGenerator!AS2:AS123)');


}


/**
 * FileGenerator: Deprecated, replaced by newFileGeneratorSheetReset function: this function resets all fields that can be edited by the user and fills in the correct formulas or values.
 
function resetFileGenerator() {   // This function restores all the formulas and values that may have been changed by the user

  //generate empty arrays which will later be populated in a loop and filled ones for the simple cases

  var formulaR2toR3 = [[''], ['=countif(Plates!B2:B,R2)+1']];
  var contentR17to18 = [[''], ['']];
  var contentR14 = [['false']];
  var contentR11 = [['']];
  var contentR23to33 = [[0.08], [10], [3], [3], [800], [5], [''], [''], [400], ["Please fill in"], ["Please fill in"]];
  var contentS7toU10 = [['=if(U7="","",DropdownTables!L3)', '=if(U7="","",index(\'Component DB\'!B:C,match(U7,\'Component DB\'!C:C,0),1))', ''],
  ['=if(U8="","",DropdownTables!L10)', '=if(U8="","",index(\'Component DB\'!B:C,match(U8,\'Component DB\'!C:C,0),1))', ''],
  ['=if(U9="","",DropdownTables!L17)', '=if(U9="","",index(\'Component DB\'!B:C,match(U9,\'Component DB\'!C:C,0),1))', ''],
  ['=if(U10="","",DropdownTables!L24)', '=if(U10="","",index(\'Component DB\'!B:C,match(U10,\'Component DB\'!C:C,0),1))', '']];
  var contentV2to6 = [['true'], ['false'], ['false'], ['false'], ['false']];
  var formulasYAA = [];
  var formulasACDEFG = [];
  var formulasAHAO = [];
  var contentA2AP = [];


  //loops from (row) 2 to 123 and fills the array with formulas that need to be in the respective cells

  for (var i = 2; i < 124; i++) {

    // formulas for columns Y and Z: Batch ID, Producer

    formulasYAA.push(['=if(T' + i + '="","",tempTables!$A$' + (15 + 6 * i) + ')',
    '=if(T' + i + '=\"\",\"\",iferror( vlookup(T' + i + ',\'Batch DB\'!B:D,3,0),\"Not registered\"))',
    '=if(T' + i + '="","",if(iferror( vlookup(T' + i + ',\'Batch DB\'!B:E,4,0),1) = "",0.999, iferror( vlookup(T' + i + ',\'Batch DB\'!B:E,4,0),1)))']);

    // =if(T2="","",if(iferror( vlookup(T2,'Batch DB'!B:E,4,0),1)="",0.999,iferror(vlookup(T2,'Batch DB'!B:E,4,0),1)))

    // Rows 2 to 6 contain the starting materials and product and thus the formulas differ here from the rest of the sheet for columns AD to AF: equivalents, mass, solvenent mL/g

    if (i < 7) {
      formulasACDEFG.push(['=if(T' + i + '="","", VLOOKUP(T' + i + ',\'Component DB\'!B:J,8,false))',
      '=if(T' + i + '="","",if(V' + i + '*1+(S' + i + '="Product"),1,"Enter value"))',
      '=if((V' + i + ')*(T' + i + '<>""), "Enter Mass",if((AD' + i + '<>"")*isnumber(AD' + i + ')*(sumproduct($V$2:$V$6)=1),$W$2/$X$2*AD' + i + '*AB' + i + '/AA' + i + ',""))',
      '=if(T' + i + '="","","-")',
      '=if(T' + i + '="", "",if(regexmatch(S' + i + ',"Solvent"),AF' + i + '*$W$2,if((isnumber(AC' + i + '))*(isnumber(AE' + i + ')),AE' + i + '/AC' + i + ',"-")))'
      ]);
    } else {
      formulasACDEFG.push(['=if(T' + i + '="","", VLOOKUP(T' + i + ',\'Component DB\'!B:J,8,false))',
      '=if(T' + i + '="","",if(regexmatch(S' + i + ',"Solvent"),"-",if(regexmatch(S' + i + ',"Catalyst"),$R$23,if(regexmatch(S' + i + ',"Acid"),$R$25,if(regexmatch(S' + i + ',"Base"),$R$25,if(regexmatch(S' + i + ',"Coupling Reagent"),$R$26,"Enter Equiv"))))))',
      '=if(T' + i + '="","",if((AD' + i + '<>"")*isnumber(AD' + i + ')*(AB' + i + '<>"")*isnumber(AB' + i + ')*($X$2<>0),$W$2/$X$2*AD' + i + '*AB' + i + '/AA' + i + ',""))',
      '=if(T' + i + '="","",iferror(find("Solvent",S' + i + ')*$R$24,"-"))',
      '=if(T' + i + '="", "",if(regexmatch(S' + i + ',"Solvent"),AF' + i + '*$W$2,if(isnumber(AC' + i + '),AE' + i + '/AC' + i + ',"-")))'
      ]);
    }

    //formulas for columns AH to AP: Dose as, Solution ID, Concentration, Unit (Arrayformula in concentration column covering both concentration and unit column), Solvent, Solvent Batch ID, Density, Volume, Evaporate before?

    formulasAHAO.push(['=if(T' + i + '="","",if(isnumber(AC' + i + '),"Liquid","Solid"))',
    '=if(AH' + i + '<> "Solution","",tempTables!$C$' + (15 + 6 * i) + ")",
    '=if((AI' + i + '<>"")*(AI' + i + '<>"Not registered"),iferror(query(Solutions!A2:G, "Select F, G where A = \'"&AI' + i + '&"\' limit 1",),{"define concentration","Unit?"}),{"",""})',
      '',
    '=if((AI' + i + '<>"")*(AI' + i + '<>"Not registered"),iferror(vlookup(vlookup(AI' + i + ',Solutions!A:D,4,0),\'Component DB\'!B2:C,2,0),"Solvent?"),"")',
    '=if(AH' + i + '<> "Solution","",tempTables!$B$' + (15 + 6 * i) + ")",
    '=if((AI' + i + '<>"")*(AI' + i + '<>"Not registered"),iferror(vlookup(AI' + i + ',Solutions!A:H,8,0),"Density?"),"")',
    '=if($T' + i + '="","", if($AH' + i + '="Solution",iferror(switch(AK' + i + ',"M",AD' + i + '*$W$4/AJ' + i + '*1000, "%w/w",AE' + i + '*100/AJ' + i + '/AN' + i + ', "g/L", AE' + i + '/AJ' + i + '*1000),"not defined"),""))'
    ]);
  }
  //connect to the FileGenerator sheet and write the arrays. 
  var FileGeneratorSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FileGenerator");
  FileGeneratorSheet.getRange("R2:R3").setFormulas(formulaR2toR3);
  FileGeneratorSheet.getRange(2, 25, formulasYAA.length, 3).setFormulas(formulasYAA);
  FileGeneratorSheet.getRange(2, 29, formulasACDEFG.length, 5).setFormulas(formulasACDEFG);
  FileGeneratorSheet.getRange(2, 34, formulasAHAO.length, 8).setFormulas(formulasAHAO);
  FileGeneratorSheet.getRange("AP2:AP123").setValue('false');
  FileGeneratorSheet.getRange("V2:V6").setValues(contentV2to6);
  FileGeneratorSheet.getRange("R11").setValues(contentR11);
  FileGeneratorSheet.getRange("R17:R18").setValues(contentR17to18);
  FileGeneratorSheet.getRange("R23:R33").setValues(contentR23to33);
  FileGeneratorSheet.getRange("R14").setValue('false');
  FileGeneratorSheet.getRange("R15").setFormula('=PlateBuilder!C16');
  FileGeneratorSheet.getRange("R8:R9").clearContent();
  SpreadsheetApp.flush();
  FileGeneratorSheet.getRange("R8:R9").setFormulas([['=iferror(VLOOKUP(FileGenerator!R2,\'HTE-Requests\'!A2:D,2,false),"not found")'], ['=iferror(VLOOKUP(FileGenerator!R2,\'HTE-Requests\'!A2:D,3,false),"not found")']]);   // those two cells need to contain formulas, as they're used to retrieve project name and step name from the HTE-requests sheet
  FileGeneratorSheet.getRange("S7:U10").setValues(contentS7toU10);
  FileGeneratorSheet.getRange("R39:R40").setValue('false');
  FileGeneratorSheet.getRange("R38").setValue(1.5);
  FileGeneratorSheet.getRange("R42").setValue('false');
  FileGeneratorSheet.getRange("R41").setValue('MinusPlus');


  //It's unknown, why explicit references to the FileGenerator sheet are important for this formula and the ones in R8 and R9, but it would generate a #Ref error from May 19 onward, if e.g. R2 was used. 

  FileGeneratorSheet.getRange("R50").clearContent();
  SpreadsheetApp.flush();
  FileGeneratorSheet.getRange("R50").setFormula('=(FileGenerator!R2="")+(FileGenerator!R8="not found")+(FileGenerator!R9="not found")+(FileGenerator!R11="")+(FileGenerator!R13="not found")+(FileGenerator!R14=TRUE)*((FileGenerator!R17="")+(FileGenerator!R18=""))+(FileGenerator!AE2="Enter Mass")+(FileGenerator!AE3="Enter Mass")+(FileGenerator!AE4="Enter Mass")+(FileGenerator!AE5="Enter Mass")+(FileGenerator!AD2="Enter value")+(FileGenerator!AD3="Enter value")+(FileGenerator!AD4="Enter value")+(FileGenerator!AD5="Enter value")+sum(FileGenerator!AS2:AS123)');
  FileGeneratorSheet.getRange("R29").setFormula('=if($W$2*max(AF2:AF123)+max(AO2:AO123)>800,4,1)');
}*/

/**
 * FileGenerator: This resets the lookup tables in the tempTables Sheet and the conditional formatting in File Generator. It is normally not used, only if columns/rows were deleted by the user.
 */
function resetFileGeneratorSheet() {
  var tempTablesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("tempTables");
  var FileGeneratorSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FileGenerator");
  var conditionalFormatRules = FileGeneratorSheet.getConditionalFormatRules();

  for (var i = -1; i < 121; i++) {

    //Writes the formulas and cell formatting in sheet tempTables column A for the queries and data validations that belong to column T in sheet FileGenerator. 
    //The formula in column A retrieves the 5 latest batches registered for the component ID in FileGenerator column T.
    tempTablesSheet.getRange(32 + 6 * i, 1).setValue("Dropdown for FileGenerator T" + (i + 3) + ":");
    tempTablesSheet.getRange(33 + 6 * i, 1).setFormula('=if(FileGenerator!T' + (i + 3) + '="","",iferror(query(\'Batch DB\'!A2:C, "Select C where B = "&FileGenerator!T' + (i + 3) + '&" order by A desc limit 5"),"Not registered"))');

    tempTablesSheet.getRange(32 + 6 * i, 1).setBackground('#ffffff');
    tempTablesSheet.getRange(33 + 6 * i, 1).setBackground('#fff2cc');
    tempTablesSheet.getRange(34 + 6 * i, 1, 4, 1).setBackground('#d9ead3');

    FileGeneratorSheet.getRange(i + 3, 25).setDataValidation(SpreadsheetApp.newDataValidation()
      .setAllowInvalid(true)
      .requireValueInRange(FileGeneratorSheet.getRange('tempTables!$A$' + (33 + 6 * i) + ':$A$' + (37 + 6 * i)), false)
      .build());

    FileGeneratorSheet.getRange(i + 3, 25).setFormula("=tempTables!$A$" + (33 + 6 * i));


    //Writes the formulas and cell formatting in sheet tempTables column B for the queries and data validations that belong to column AM in sheet FileGenerator.
    //The formula in column B retrieves the 5 latest batches registered for the component ID in FileGenerator column AM.
    tempTablesSheet.getRange(32 + 6 * i, 2).setValue("Dropdown for FileGenerator AM" + (i + 3) + ":");

    tempTablesSheet.getRange(33 + 6 * i, 2).setFormula('=if(FileGenerator!AL' + (i + 3) + '="","",iferror(query(\'Batch DB\'!A2:F, "Select C where F = "&FileGenerator!AL' + (i + 3) + '&" order by A desc limit 5"),"Not registered"))');
    tempTablesSheet.getRange(32 + 6 * i, 2).setBackground('#ffffff');
    tempTablesSheet.getRange(33 + 6 * i, 2).setBackground('#fff2cc');
    tempTablesSheet.getRange(34 + 6 * i, 2, 4, 1).setBackground('#d9ead3');

    FileGeneratorSheet.getRange(i + 3, 39).setDataValidation(SpreadsheetApp.newDataValidation()
      .setAllowInvalid(true)
      .requireValueInRange(FileGeneratorSheet.getRange('tempTables!$B$' + (33 + 6 * i) + ':$B$' + (37 + 6 * i)), false)
      .build());

    FileGeneratorSheet.getRange(i + 3, 39).setFormula("=tempTables!$B$" + (33 + 6 * i));

    //Writes the formulas and cell formatting in sheet tempTables column C for the queries and data validations that belong to column AI in sheet FileGenerator. 
    //The formula in column C retrieves up to 5 solution IDs registered for the component ID in FileGenerator column T.

    tempTablesSheet.getRange(32 + 6 * i, 3).setValue("Dropdown for FileGenerator AI" + (i + 3) + ":");

    tempTablesSheet.getRange(33 + 6 * i, 3).setFormula('=if(FileGenerator!AH' + (i + 3) + ' <> "Solution","",iferror( query(Solutions!A2:C,"Select A where (B = "& FileGenerator!T' + (i + 3) + '&" and C = \'" &FileGenerator!Y' + (i + 3) + ' & "\' ) limit 5"), "Solution not registered"))');


    //'=if(FileGenerator!AH'+(i+3)+' <> "Solution","",iferror( query(Solutions!A2:H,"Select A where (B = "& FileGenerator!T'+(i+3)+'&" and C = '" &FileGenerator!Y'+(i+3)+' & "' and F = "&FileGenerator!AJ'+(i+3)+'&" and G = '" &FileGenerator!AK'+(i+3)+' & "' and H = "&FileGenerator!AN'+(i+3)+' & ") limit 5"), "Solution not registered"))'

    tempTablesSheet.getRange(32 + 6 * i, 3).setBackground('#ffffff');
    tempTablesSheet.getRange(33 + 6 * i, 3).setBackground('#fff2cc');
    tempTablesSheet.getRange(34 + 6 * i, 3, 4, 1).setBackground('#d9ead3');

    FileGeneratorSheet.getRange(i + 3, 35).setDataValidation(SpreadsheetApp.newDataValidation()
      .setAllowInvalid(true)
      .requireValueInRange(FileGeneratorSheet.getRange('tempTables!$C$' + (33 + 6 * i) + ':$C$' + (37 + 6 * i)), false)
      .build());

    FileGeneratorSheet.getRange(i + 3, 35).setFormula("=tempTables!$C$" + (33 + 6 * i));


    //This code generates conditional formatting rules which color the respective cell in column Y, AM or AI light green, if a certain cell in tempTables is not empty. 
    // This is only the case, if there is more than one Batch/Solution ID registered for the Component ID in question.
    conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
      .setRanges([FileGeneratorSheet.getRange('Y' + (i + 3))])
      .whenFormulaSatisfied('=(indirect("tempTables!A' + (34 + 6 * i) + '")<>"")')
      .setBackground('#d9ead3')
      .build());
    FileGeneratorSheet.setConditionalFormatRules(conditionalFormatRules);

    conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
      .setRanges([FileGeneratorSheet.getRange('AM' + (i + 3))])
      .whenFormulaSatisfied('=(indirect("tempTables!B' + (34 + 6 * i) + '")<>"")')
      .setBackground('#d9ead3')
      .build());
    FileGeneratorSheet.setConditionalFormatRules(conditionalFormatRules);


    conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
      .setRanges([FileGeneratorSheet.getRange('AI' + (i + 3))])
      .whenFormulaSatisfied('=(indirect("tempTables!C' + (34 + 6 * i) + '")<>"")')
      .setBackground('#d9ead3')
      .build());
    FileGeneratorSheet.setConditionalFormatRules(conditionalFormatRules);
  }
}

/**
 * FileGenerator: Creates or amends the given folder with the given filename with all the information found in File Generator plus info on all liquids to be dosed manually and - if applicable - the layout of the sourceplates for Junior
 * 
 * @param {String} folderID id of the folder in which the gsheet is to be created.
 * @param {String} elnId ELN-ID.
 * @param {Number} plateNumber id of the folder in which the gsheet is to be created.
 * @param {Object} locationsOnSourcePlates dictionary containing the layout of the different source plates in case the lea input file needs to be created.
 * @param {Number} overage factor >1 determining how much more liquid should be put on the source plates.
 * @param {Array} data content of the FileGenerator sheet read by the saveFile function.
 * @return {String} URL of the newly created gsheet.
 */
function backupPlate(folderId, elnId, plateNumber, locationsOnSourcePlates, overage, data) {
  folderId = globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["CHEATSHEETSfOLDERiD"];
  // adapted from: https://stackoverflow.com/questions/25106580/copy-value-and-format-from-a-sheet-to-a-new-google-spreadsheet-document 
  var rowToLetterAssignment = { 1: "A", 2: "B", 3: "C", 4: "D", 5: "E", 6: "F", 7: "G", 8: "H" };   //translates row numbers into the corresponding letter
  var source = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FileGenerator");
  var sourceName = source.getSheetName();
  var sValues = source.getDataRange().getValues();
  var cheatSheetFileId = createSpreadsheet(globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["CHEATSHEETSfOLDERiD"], elnId);
  var destination = SpreadsheetApp.openById(cheatSheetFileId);   // inputs are the folder ID where the file should be located and the name of the Spreadsheet
  var cheatSheetUrl = destination.getUrl();
  source.copyTo(destination);
  var destinationSheet = destination.getSheetByName('Copy of ' + sourceName);
  if (destinationSheet) {
    destinationSheet.getRange(1, 1, sValues.length, sValues[0].length).setDataValidation(null);
    destinationSheet.getRange(1, 1, sValues.length, sValues[0].length).setValues(sValues);// overwrite all formulas that the copyTo preserved
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast('There was an error generating the cheatsheet using English as System Language, trying German now.', 'Error', 3);
    destinationSheet = destination.getSheetByName('Kopie von ' + sourceName);
  }

  if (destinationSheet) {
    destinationSheet.getRange(1, 1, sValues.length, sValues[0].length).setDataValidation(null);
    destinationSheet.getRange(1, 1, sValues.length, sValues[0].length).setValues(sValues);// overwrite all formulas that the copyTo preserved
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast('There was an error generating the cheatsheet.', 'Error', 3);
  }

  var itt = destination.getSheetByName('Plate ' + plateNumber);

  if (itt) {  // if the sheet exists already, it'll be deleted
    destination.deleteSheet(itt);
  }
  destinationSheet.setName('Plate ' + plateNumber);

  var sheet1 = destination.getSheetByName('Sheet1');
  if (sheet1) {  // when the spreadsheet is freshly created, it contains a starting sheet which is deleted by this code
    destination.deleteSheet(sheet1);
  }
  // When dosing manually, liquid components often are forgotten, thus cell O11 is used to put an overview of all liquids and solutions found on the plate. 
  var manuallyDosedComponents = "";
  var targetAmount = "";
  var solutionParameters = "";
  var solventOfInterest = "";
  for (var row = 0; row < data.length; row++) {
    if (data[row][16] != "Liquid" && data[row][16] != "Solution" && data[row][16] != "NonDose Solid") { continue; }
    if (row < 10) { data[row][5] = data[row][3]; } // For starting materials data[row][5] contains non-relevant information, but for plate components the combination of component name and level which is what we want. 

    switch (data[row][16]) {
      case "Liquid":
        targetAmount = parseFloat(data[row][15]).toFixed(1) + " uL ";
        solutionParameters = "";
        solventOfInterest = "";
        break;
      case "Solution":
        targetAmount = parseFloat(data[row][23]).toFixed(1) + " uL ";
        solutionParameters = "(" + data[row][18] + data[row][19];
        if (data[row][20] == "water") { solventOfInterest = ", aq)"; } else { solventOfInterest = " in " + data[row][20] + ")"; }
        break;
      case "NonDose Solid":
        targetAmount = parseFloat(data[row][13]).toFixed(1) + " mg ";
        solutionParameters = " TO BE DOSED MANUALLY!";
        solventOfInterest = "";
        break;
    }


    manuallyDosedComponents = manuallyDosedComponents + data[row][5] + solutionParameters + solventOfInterest + ": " + targetAmount + String.fromCharCode(10);
  }
  destinationSheet.getRange('O11').setValue(manuallyDosedComponents)
    .setVerticalAlignment('middle');
  destinationSheet.getRange('O10').setFontColor('#000000')
    .setFontWeight('bold')
    .setFontStyle('italic')
    .setValue("Components to be dosed manually:");
  destinationSheet.getRange('P10').clearContent();


  //  locationsOnSourcePlates   plate type : [   [rowCount,colCount, maxVol], [        [               [                vialVolume, ComponentName+Level            ]]]]
  //                                         [0]   [0][0]                      [1]        [1][0]       [1][0][0]          [1][0][0][0]         [1][0][0][1]
  //                                          Headerinfo on Plate type        all plates    one plate     rows on plate      row property      wellsDictionary key

  //  Put the source plates below the reaction plate, so it's documented and can be used when running the reaction

  var startingRow = 24;
  var currentRow = startingRow; // Row number where the first Source plate begins
  var sourcePlateContent = [];  // holds the physical represenation of the source plate in question as an array that can be written to the sheet later.
  var sourcePlateCount = 0;     // counts the total number of source plates with something on them
  var maxColumnCount = 0;       // what is the sourcePlate with the largest number of columns, determines how wide the array has to be that is 

  for (let key in locationsOnSourcePlates) {
    if (locationsOnSourcePlates[key][1].length == 0) { continue; }
    if (locationsOnSourcePlates[key][0][1] > maxColumnCount) { maxColumnCount = locationsOnSourcePlates[key][0][1]; }
  }

  for (let key in locationsOnSourcePlates) {                                     //locationsOnSourcePlates dictionary with the plate type as key (e.g. 96_0.8 ) and an array of plates filled with compounds filled in there by function amendSourcePlate (and afterwards optimizeSourcePlate)
    if (locationsOnSourcePlates[key][1].length == 0) { continue; }                                                                         // true, if there's at least one plate for this category, i.e. the number of rows on the first plate is greater than zero.

    for (var sourcePlate = 0; sourcePlate < locationsOnSourcePlates[key][1].length; sourcePlate++) {                            // go through all the plates of this category
      sourcePlateContent.push(["", key + " mL " + (sourcePlate + 1)]);
      for (var residualColumnsInRow = 2; residualColumnsInRow < (1 * maxColumnCount + 2); residualColumnsInRow++) {
        sourcePlateContent[currentRow - startingRow].push("");   //fill up the rest of the line with empty cells
      }
      currentRow++;
      for (var sourcePlateRow = 0; sourcePlateRow < locationsOnSourcePlates[key][0][0]; sourcePlateRow++) {
        sourcePlateContent.push([rowToLetterAssignment[sourcePlateRow + 1]]);
        for (var columnsInRow = 1; columnsInRow < (1 * maxColumnCount + 2); columnsInRow++) {
          //for (var sourcePlateRow = 0; sourcePlateRow < locationsOnSourcePlates[key][1][sourcePlate].length; sourcePlateRow++){    // go through all the filled rows of the plate in question
          if (sourcePlateRow < locationsOnSourcePlates[key][1][sourcePlate].length && columnsInRow < 7) {
            sourcePlateContent[currentRow - startingRow].push("◯");
          } else if (sourcePlateRow < locationsOnSourcePlates[key][1][sourcePlate].length && columnsInRow < (1 * locationsOnSourcePlates[key][0][1] + 1)) {
            sourcePlateContent[currentRow - startingRow].push(" ");
          } else if (sourcePlateRow < locationsOnSourcePlates[key][1][sourcePlate].length && columnsInRow < (1 * locationsOnSourcePlates[key][0][1] + 2)) {
            sourcePlateContent[currentRow - startingRow].push(parseFloat(overage * locationsOnSourcePlates[key][1][sourcePlate][sourcePlateRow][0]).toFixed(0) + " ul " + locationsOnSourcePlates[key][1][sourcePlate][sourcePlateRow][1] + " per Vial (incl. " + (overage - 1) * 100 + "% overage)");
          } else if (columnsInRow < (1 * locationsOnSourcePlates[key][0][1] + 1)) {
            sourcePlateContent[currentRow - startingRow].push(" ");
          } else if (columnsInRow < (1 * locationsOnSourcePlates[key][0][1] + 2)) {
            sourcePlateContent[currentRow - startingRow].push("empty");
          } else if (columnsInRow < (1 * maxColumnCount + 2)) {
            sourcePlateContent[currentRow - startingRow].push("");
          }
        }
        currentRow++;
      }
    }
    sourcePlateCount++;
    sourcePlateContent.push(["", ""]);
    for (let residualColumnsInRow = 2; residualColumnsInRow < (1 * maxColumnCount + 2); residualColumnsInRow++) {
      sourcePlateContent[currentRow - startingRow].push("");   //fill up the rest of the line with empty cells
    }
    currentRow++;
  }

  if (sourcePlateContent.length > 0) {  //if no junior file is written, then this array will be empty
    destinationSheet.getRange(startingRow, 1, sourcePlateContent.length, sourcePlateContent[0].length).setValues(sourcePlateContent);
  }


  // Finally, the column containing Solvent (mL/g) is replaced with the dosing Head IDs, since at this stage the dosing head ID is more important and the Solv mL/g can be derived from the liquid volume 
  destinationSheet.getRange('AF1:AF123').clear({ contentsOnly: true, skipFilteredRows: true });
  destinationSheet.getRange('AF2').setFormula('=ArrayFormula(if(T2:T123="","",if(isnumber(AC2:AC123),"This is a liquid",T2:T123&"@"&if(Y2:Y123 = "", "Not registered",right(Y2:Y123,14)))))');
  destinationSheet.getRange('AF1').setValue('Dosing Head ID');
  destinationSheet.autoResizeColumns(32, 1);
  destinationSheet.clearConditionalFormatRules(); //delete all Conditional Formatting, since many of the BatchIDs will be red and displease Vera's eyes
  // Now rebuild the conditional formatting in the topleft corner's plateview that indicates which components are liquids
  var conditionalFormatRules = destinationSheet.getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
    .setRanges([destinationSheet.getRange('O2:Q9')])
    .whenFormulaSatisfied('=(VLOOKUP(if(left(O2,1) =" ",mid(O2,2,len(O2)-1),O2),$W$12:$AH$123,12,false)="Liquid")+(VLOOKUP(if(left(O2,1) =" ",mid(O2,2,len(O2)-1),O2),$W$12:$AH$123,12,false)="Solution")')
    .setBackground('#EAD1DC')
    .build());
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
    .setRanges([destinationSheet.getRange('B11:M13')])
    .whenFormulaSatisfied('=(VLOOKUP(if(left(B11,1) =" ",mid(B11,2,len(B11)-1),B11),$W$12:$AH$123,12,false)="Liquid")+(VLOOKUP(if(left(B11,1) =" ",mid(B11,2,len(B11)-1),B11),$W$12:$AH$123,12,false)="Solution")')
    .setBackground('#EAD1DC')
    .build());

  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
    .setRanges([destinationSheet.getRange('B2:M13')])
    .whenFormulaSatisfied('=(VLOOKUP(if(left(B2,1) =" ",mid(B2,2,len(B2)-1),B2),$W$12:$AH$50,12,false)="Liquid")+(VLOOKUP(if(left(B2,1) =" ",mid(B2,2,len(B2)-1),B2),$W$12:$AH$50,12,false)="Solution")')
    .setBackground('#ead1dc')
    .build());

  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
    .setRanges([destinationSheet.getRange('B2:M9')])
    .whenFormulaSatisfied('=(($O2<>" ") + ($R$15 = true))*(B$11<>" ")*(($O2<>"") + ($R$15 = true))*(B$11<>"")')
    .setBackground('#B7E1CD')
    .build());

  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
    .setRanges([destinationSheet.getRange('B2:M9')])
    .whenFormulaSatisfied('=(($O2=" ")+($O2=""))+((B$11=" ")+(B$11=""))')
    .setBackground('#FFF2CC')
    .build());
  destinationSheet.setConditionalFormatRules(conditionalFormatRules);

  if (destinationSheet.getRange("R15").getValue == "TRUE") { // merge the cells belonging to the broad screen component

    var layout = destinationSheet.getRange("B2:M9").getValues();
    var periodCounter = 0;
    var lengthOfPeriod = [];
    var contentOfLastCell = layout[0][0];
    var contentOfCurrentCell = layout[0][0];
    var endOfRow = 0;

    for (row = 0; row < layout.length; row++) {
      for (var column = 0; column < layout[0].length; column++) {

        contentOfCurrentCell = layout[row][column];

        if (contentOfCurrentCell != contentOfLastCell || endOfRow == 1) {
          lengthOfPeriod.push([periodCounter, contentOfCurrentCell, row + 1, column + 1]);
          if (periodCounter > 1) {
            if (endOfRow == 0) { destinationSheet.getRange(2, column - periodCounter + 2, 8, periodCounter).mergeAcross(); } else {
              destinationSheet.getRange(2, 12 - periodCounter + 2, 8, periodCounter).mergeAcross();
            }

          }
          periodCounter = 0;
        }
        periodCounter++;
        contentOfLastCell = contentOfCurrentCell;
        endOfRow = 0;
      }
      endOfRow = 1;
    }


  }

  return cheatSheetUrl;

}

/**
 * FileGenerator: Creates a new Google sheet in the given folder with the given filename (analogous to the createSpreadsheet-function ) unless it already exists.
 * copied from https://yagisanatode.com/2018/07/08/google-apps-script-how-to-create-folders-in-directories-with-driveapp/ and amended to create Spreadsheet, not folder
 * 
 * @param {String} folderID id of the folder in which the gsheet is to be created.
 * @param {String} fileName name of the new gsheet.
 * @return {String} fileID of the newly created gsheet.
 */
function createSpreadsheet(folderID = globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["CHEATSHEETSfOLDERiD"], fileName = "") {
  var parentFolder = DriveApp.getFolderById(folderID);
  var filesInFolder = parentFolder.getFiles();
  var doesntExist = true;
  var newFile = '';

  // Check if folder already exists.
  while (filesInFolder.hasNext()) {
    var file = filesInFolder.next();

    //If the name exists return the id of the folder
    if (file.getName() === fileName) {
      doesntExist = false;
      newFile = file;
      return newFile.getId();
    }
  }
  //If the name doesn't exist, then create a new file
  if (doesntExist == true) {
    //If the file doesn't exist

    var resource = {
      title: fileName,
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [{ id: folderID }]
    };
    var fileJson;
    try {
      fileJson = Drive.Files.insert(resource, null, { supportsAllDrives: true });    //Introduced, because it sometimes crashes with this error: API call to drive.files.insert failed with error: Internal Error
    }
    catch (err) {
      Logger.log(err);
      sleep(1000);
      fileJson = Drive.Files.insert(resource, null, { supportsAllDrives: true });  // {supportsAllDrives: true} needed for generating files on a Google Shared Drive https://qa.ostack.cn/qa/?qa=852949/
    }
    finally {
      var fileId = fileJson.id;

      return fileId;
    }
  }
}

// copied from https://yagisanatode.com/2018/07/08/google-apps-script-how-to-create-folders-in-directories-with-driveapp/
//Create folder unless it already exists
/**
 * FileGenerator: Creates a new subfolder in the given folder with the given foldername (analogous to the createSpreadsheet-function ) unless it already exists.
 * copied from https://yagisanatode.com/2018/07/08/google-apps-script-how-to-create-folders-in-directories-with-driveapp/ 
 * 
 * @param {String} folderID id of the folder in which the folder is to be created.
 * @param {String} fileName name of the new folder.
 * @return {String} fileID of the newly created folder.
 */
function createFolder(folderID, folderName) {
  var parentFolder = DriveApp.getFolderById(folderID);
  var subFolders = parentFolder.getFolders();
  var doesntExist = true;
  var newFolder = '';

  // Check if folder already exists.
  while (subFolders.hasNext()) {
    var folder = subFolders.next();

    //If the name exists return the id of the folder
    if (folder.getName() === folderName) {
      doesntExist = false;
      newFolder = folder;
      return newFolder.getId();
    }
  }
  //If the name doesn't exist, then create a new folder
  if (doesntExist == true) {
    //If the file doesn't exist
    newFolder = parentFolder.createFolder(folderName);
    return newFolder.getId();
  }
}

/**
 * FileGenerator: remove all rows from a given gSheet that match the string plateToDelete. Function is called from savePlate() in cases, where a plate needs to be overwritten
 * adapted from: https://yagisanatode.com/2019/06/12/google-apps-script-delete-rows-based-on-a-columns-cell-value-in-google-sheet/
 * 
 * @param {Object} sheet the sheet in which the given lines need to be deleted.
 * @param {String} plateToDelete name of the plate that needs to be deleted.
 * @return {Number} 0 or 1 depending on whether the plate was found or not.
 */
function removeThenSetNewVals(sheet, plateToDelete) {

  var dataRange = sheet.getDataRange();
  var rangeVals = dataRange.getValues();
  var found = 0;
  var newRangeVals = [];

  for (var i = 0; i < rangeVals.length; i++) {
    if (rangeVals[i][0] != plateToDelete) {

      newRangeVals.push(rangeVals[i]);

    } else { found = 1; }
  }
  if (found == 0) { return found; }
  dataRange.clearContent();

  var newRange = sheet.getRange(1, 1, newRangeVals.length, newRangeVals[0].length);
  newRange.setValues(newRangeVals);
  return found;
}

/**
 * FileGenerator: This function is called from savePlate when it iterates through all the ingredients to be put on the plate and for every coordinate in which the component needs to end up. 
 * It accepts what is currently known about the plate and amends it with what needs to be added for the given compound / vial. 
 * wellsData is an array that contains everything needed to write all the output files: wellsData[0] contains all the information that was later written to the wells sheet, [1] is the data going to the Quantos input csv and is not filled anymore
 * @param {Array} quantosXmlArray contains info on all the solids that need to be dosed on Quantos.
 * @param {Array} data one line containing all the info on one reaction component.
 * @param {Array} wellsData array that will later be written to the Wells Sheet.
 * @param {String} plateCoordinate Coordinate of the vial in question.
 * @param {String} ELNiD ELN-ID.
 * @param {Number} plateNumber number of the plate.
 * @param {String} vialOption what kind of vial is on the plate.
 * 
 * @return {Array} wellsdata Array, now amended with the information taken from the line in the file generator sheet and vial coordinate in question.
 */
function writeWellsData(quantosXmlArray, data, wellsData, plateCoordinate, ELNiD, plateNumber, vialOption) {

  // wellsData acts as a data shuttle between the two functions in the absence of global variables: 
  // wellsData[0] contained all the information that as later written to the wells sheet, 
  // wellsdata[1] is the data going to the Quantos input csv and is not filled anymore
  var wellsDictionary = wellsData[3];  // this dictionary contains the wells data to be written to the junior xml
  var plateIngredientsDictionary = wellsData[2]; // this dictionary contains all the ingredients used on the plate and sums up the mass, volume and solution volume for use in the junior xml

  var firstRowNumber = 0;                 // Initialize the variable used to store the first well of the dosing range: It's the top left corner of the dosing range
  var firstColNumber = 0;
  var lastRowNumber = 0;                  // Initialize the variables used to compare the current plate coordinates with the coordinate of the last well of a continuous range
  var lastColNumber = 0;
  var notPartOfRange = 0;
  var numberOfKeysInWellsDictForDb = Object.keys(WellsDictForDb).length;
  WellsDictForDb[numberOfKeysInWellsDictForDb + 1] = {};        // contains later one line of data

  //depending on the dosing mode, the field for mass or volume may be empty which trips up the parseFloat(x).toFixed rounding function, thus the distincion of two cases for these variables
  var solutionVolume;
  var liquidVolume;
  var mass;
  if (String(data[23]).length > 0) { solutionVolume = parseFloat(data[23]).toFixed(2); } else { solutionVolume = ''; }
  if (String(data[15]).length > 0) { liquidVolume = parseFloat(data[15]).toFixed(2); } else { liquidVolume = ''; }
  if (String(data[13]).length > 0) { mass = parseFloat(data[13]).toFixed(2); } else { mass = parseFloat(data[15] * data[11]).toFixed(2); }

  var rowNumber = 0;
  var colNumber = Number(String(plateCoordinate).substr(1));   // column number of the current well
  switch (String(plateCoordinate).substr(0, 1)) {                 // row number of the current well (this back translation of the column literal into the row number is awkward, but saves adding these as input for the function)
    case "A":
      rowNumber = 1;
      break;
    case "B":
      rowNumber = 2;
      break;
    case "C":
      rowNumber = 3;
      break;
    case "D":
      rowNumber = 4;
      break;
    case "E":
      rowNumber = 5;
      break;
    case "F":
      rowNumber = 6;
      break;
    case "G":
      rowNumber = 7;
      break;
    case "H":
      rowNumber = 8;
      break;
  }

  //**********   Determine the starting and end-wells of continuous sequences for the Junior xml headers in the LSLibraries section

  var tempData = data[4];
  if (data[4] === true || data[4] === false) {    // This is a workaround, since data[4] if overwritten will influence the state of data[row][4] in the savePlate function
    data[4] = '';
  }

  if (wellsDictionary[data[3] + data[4]].length == 2) {                                                                                    // true, if the well in question is the first well of a certain compound
    wellsDictionary[data[3] + data[4]].push([rowNumber, colNumber, plateCoordinate, rowNumber, colNumber, plateCoordinate]);     // Add the coordinate as a compartment to the array which is returned by the dictionary
  } else {

    for (var dosingRange = 2; dosingRange < wellsDictionary[data[3] + data[4]].length; dosingRange++) { //Go through all the independent continuous dosing ranges found so far for this compound and check whether the current well belongs to it

      firstRowNumber = wellsDictionary[data[3] + data[4]][dosingRange][0];                 // the top left corner of the dosing range
      firstColNumber = wellsDictionary[data[3] + data[4]][dosingRange][1];

      lastRowNumber = wellsDictionary[data[3] + data[4]][dosingRange][3];                  // The last coordinate found for the continuous dosing range in question 
      lastColNumber = wellsDictionary[data[3] + data[4]][dosingRange][4];

      // A well belongs to an existing dosing range, if the columnnumber is the same as the column number of the first well of the sequence and the difference of the rownumber and the row number of the current last well of the sequence is not bigger than 1

      if ((firstColNumber == colNumber) && (rowNumber - lastRowNumber == 1) || (rowNumber == lastRowNumber) && (colNumber - lastColNumber == 1) || (rowNumber == firstRowNumber) && (colNumber - lastColNumber == 1) || (rowNumber - lastRowNumber == 1) && (colNumber == lastColNumber)) { // true, if the current well is part of a known continuous range

        notPartOfRange = 0;
        wellsDictionary[data[3] + data[4]][dosingRange][3] = rowNumber;                        // current row and column numbers as well as coordinates replace the one from the last well, leaving the starting values alone
        wellsDictionary[data[3] + data[4]][dosingRange][4] = colNumber;                        // --> This way, the complete sequence will be built. 
        wellsDictionary[data[3] + data[4]][dosingRange][5] = plateCoordinate;
        break;     //   the current well has found a home, the loop is stopped

      } else {                                                                                                             // Current well is not part of a continuous sequence of wells 
        notPartOfRange = 1;
      }
    }
    if (notPartOfRange == 1) { // if, after checking all the known ranges, the current well wasn't assigned to a range, a new range is created
      wellsDictionary[data[3] + data[4]].push([rowNumber, colNumber, plateCoordinate, rowNumber, colNumber, plateCoordinate]);    // add the current coordinate as new sub-array to the dictionary that will from now on be iterated over
    }
  }
  var componentRole = data[1];
  var batchId = "undefined";
  if (data[7].length > 0) batchId = data[7]; // batchID is a key column in the wells table on the database. Thus, if empty in the gSheet, it is set to "undefined"

  if (componentRole == "Starting Material") {   //Spotfire differentiates between limiting and other SM, thus this distinction is made.
    if (tempData == true) {
      componentRole = "lim SM";
    } else {
      componentRole = "other SM";
    }
  }

  switch (data[16]) {    // depending on the dosing type, different data is written to the sub-arrays of wellsData (partially represented by dictionaries to make it more easily readable
    case "Solution":

      WellsDictForDb[numberOfKeysInWellsDictForDb + 1]["ID"] = { // will contain all the data fields that are primary keys in the wells data table: ELN_ID, PLATENUMBER, Component_ID, Batch_ID, ComponentRole, Coordinate 

        ELN_ID: ELNiD,
        PLATENUMBER: plateNumber,
        Component_ID: data[2],
        Batch_ID: batchId,
        ComponentRole: componentRole,
        Coordinate: plateCoordinate
      };

      // will contain all the data fields that are not primary keys in the wells data table: PlateID and PlateIngredientID (to be removed in the future, because now redundant), limSM_or_Level, ActualMass, ActualVolume, DosingTimeStamp 
      WellsDictForDb[numberOfKeysInWellsDictForDb + 1]["DATA"] = {
        //PlateID: ELNiD + '_' + plateNumber,
        //PlateIngredientID: data[2] + '_' + data[7] + '_' + tempData + '_' + data[1],
        limSM_or_Level: tempData,
        ActualMass: mass,
        ActualVolume: solutionVolume
      };


      wellsDictionary[data[3] + data[4]][1].push([rowNumber, colNumber, plateCoordinate, parseFloat(data[23]).toFixed(2)]);      // add info about the well coordinate and how much of the solution goes in there

      //wellsData[0].push([ELNiD + '_' + plateNumber, data[2] + '_' + data[7] + '_' + tempData + '_' + data[1], plateCoordinate, mass, solutionVolume]);
      plateIngredientsDictionary[data[3] + '_' + tempData + '_' + data[18] + '_' + data[19] + '_' + data[20]][8] = parseFloat(plateIngredientsDictionary[data[3] + '_' + tempData + '_' + data[18] + '_' + data[19] + '_' + data[20]][8]) + parseFloat(mass);
      if (liquidVolume != '') { plateIngredientsDictionary[data[3] + '_' + tempData + '_' + data[18] + '_' + data[19] + '_' + data[20]][9] = parseFloat(plateIngredientsDictionary[data[3] + '_' + tempData + '_' + data[18] + '_' + data[19] + '_' + data[20]][9]) + parseFloat(liquidVolume); }  // adds the liquid volume of the current well to what was already in there from previous wells of the same compound
      plateIngredientsDictionary[data[3] + '_' + tempData + '_' + data[18] + '_' + data[19] + '_' + data[20]][10] = parseFloat(plateIngredientsDictionary[data[3] + '_' + tempData + '_' + data[18] + '_' + data[19] + '_' + data[20]][10]) + parseFloat(solutionVolume);   // adds the solution volume of the current well to what was already in there from previous wells of the same compound
      break;
    case "Solid":
    case "NonDose Solid": // This case is only different in that these solids should not go into the Quantos xml-Array and should be highlighted on the Cheatsheet, so that it's not forgotten. 
      wellsDictionary[data[3] + data[4]][1].push([rowNumber, colNumber, plateCoordinate, mass]);       // add info about the well coordinate and how much of the solid goes in there

      WellsDictForDb[numberOfKeysInWellsDictForDb + 1]["ID"] = { // will contain all the data fields that are primary keys in the wells data table: ELN_ID, PLATENUMBER, Component_ID, Batch_ID, ComponentRole, Coordinate 

        ELN_ID: ELNiD,
        PLATENUMBER: plateNumber,
        Component_ID: data[2],
        Batch_ID: batchId,
        ComponentRole: componentRole,
        Coordinate: plateCoordinate
      };

      // will contain all the data fields that are not primary keys in the wells data table: PlateID and PlateIngredientID (to be removed in the future, because now redundant), limSM_or_Level, ActualMass, ActualVolume, DosingTimeStamp 
      WellsDictForDb[numberOfKeysInWellsDictForDb + 1]["DATA"] = {
        //PlateID: ELNiD + '_' + plateNumber,
        //PlateIngredientID: data[2] + '_' + data[7] + '_' + tempData + '_' + data[1],
        limSM_or_Level: tempData,
        ActualMass: mass
      };

      //wellsData[0].push([ELNiD + '_' + plateNumber, data[2] + '_' + data[7] + '_' + tempData + '_' + data[1], plateCoordinate, mass, liquidVolume]);
      wellsData[1].push(wellsData[1].length + ";" + "C:\\Users\\Public\\Documents\\Chronos\\Methods\\Dosing Method.cam" + ";" + data[2] + "@" + batchId.toString().slice(-14) + ";" + "Tray 2" + ";" + wellsData[1].length + ";" + plateCoordinate + ";" + mass + ";" + 10);    // Quantos CSV only relevant for solid dosing - replaced by xml format

      if (data[16] == "Solid") {
        quantosXmlArray.push([quantosXmlArray.length + 2, "C:\\Users\\Public\\Documents\\Chronos\\Methods\\Dosing Method.cam", "", "", "", "", "Quantos", "", "", data[2] + "@" + batchId.toString().slice(-14), "Tray7", "", plateCoordinate, mass, parseFloat(7 / (parseFloat(mass) + 0.03) + 3).toFixed(0), ELNiD + '_' + plateNumber + " - Initial - " + vialOption, data[3]]); // will be used to generate the csl input file, tray 7 is chosen instead of a valid tray, because it forces the user to specify a tray and prevents erroneous dosing into tray 1.
      }
      plateIngredientsDictionary[data[3] + '_' + tempData + '_' + data[18] + '_' + data[19] + '_' + data[20]][8] = parseFloat(plateIngredientsDictionary[data[3] + '_' + tempData + '_' + data[18] + '_' + data[19] + '_' + data[20]][8]) + parseFloat(mass);             // adds the mass of the current well to what was already in there from previous wells of the same compound
      break;
    case "Liquid":
      wellsDictionary[data[3] + data[4]][1].push([rowNumber, colNumber, plateCoordinate, parseFloat(data[15]).toFixed(2)]);       // add info about the well coordinate and how much of the liquid goes in there

      WellsDictForDb[numberOfKeysInWellsDictForDb + 1]["ID"] = { // will contain all the data fields that are primary keys in the wells data table: ELN_ID, PLATENUMBER, Component_ID, Batch_ID, ComponentRole, Coordinate 

        ELN_ID: ELNiD,
        PLATENUMBER: plateNumber,
        Component_ID: data[2],
        Batch_ID: batchId,
        ComponentRole: componentRole,
        Coordinate: plateCoordinate
      };

      // will contain all the data fields that are not primary keys in the wells data table: PlateID and PlateIngredientID (to be removed in the future, because now redundant), limSM_or_Level, ActualMass, ActualVolume, DosingTimeStamp 
      WellsDictForDb[numberOfKeysInWellsDictForDb + 1]["DATA"] = {
        limSM_or_Level: tempData,
        ActualMass: mass,
        ActualVolume: liquidVolume
      };

      //wellsData[0].push([ELNiD + '_' + plateNumber, data[2] + '_' + data[7] + '_' + tempData + '_' + data[1], plateCoordinate, mass, liquidVolume]);  // add info about the well coordinate and how much of the liquid goes in there
      plateIngredientsDictionary[data[3] + '_' + tempData + '_' + data[18] + '_' + data[19] + '_' + data[20]][8] = parseFloat(plateIngredientsDictionary[data[3] + '_' + tempData + '_' + data[18] + '_' + data[19] + '_' + data[20]][8]) + parseFloat(mass);              // adds the mass of the current well to what was already in there from previous wells of the same compound
      plateIngredientsDictionary[data[3] + '_' + tempData + '_' + data[18] + '_' + data[19] + '_' + data[20]][9] = parseFloat(plateIngredientsDictionary[data[3] + '_' + tempData + '_' + data[18] + '_' + data[19] + '_' + data[20]][9]) + parseFloat(liquidVolume);     // adds the liquid volume of the current well to what was already in there from previous wells of the same compound
      break;
  }

  if (tempData === true || tempData === false) {    // if 
    data[4] = tempData;
  }

  wellsData[2] = plateIngredientsDictionary;
  wellsData[3] = wellsDictionary;
  wellsData[4] = quantosXmlArray;

  return wellsData;
}

/**
 * FileGenerator: resets the fileGenerator by copying a range from the SheetBackups Sheet, connected to the "Reset Sheet" Button
 * 
 */
function newFileGeneratorSheetReset() {   //executed when pressing the "Reset Sheet" Button
  // based on https://stackoverflow.com/questions/73331715/copy-a-range-to-another-spreadsheet-with-data-validations-formats-etc

  //Replicate rangeA on rangeB

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let source = ss.getSheetByName("SheetBackups");
  let destination = ss.getSheetByName("FileGenerator");
  destination.clearConditionalFormatRules();
  let rangeA = source.getRange("A155:AP287");
  let rangeB = destination.getRange("A1");

  rangeA.copyTo(rangeB);


}

