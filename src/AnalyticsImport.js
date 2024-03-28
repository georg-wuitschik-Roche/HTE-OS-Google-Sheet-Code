/*jshint sub:true*/
/**
 * Resets the AnalyticsImport Sheet, doesn't take parameters and returns nothing.
 *   
 */
function resetAnalyticsImportSheet() {
  var spreadsheet = SpreadsheetApp.getActive();

  spreadsheet.getRange('A1:L1').setDataValidation(SpreadsheetApp.newDataValidation()
    .setAllowInvalid(false)
    .requireValueInRange(spreadsheet.getRange('$AA$2:$AA$10'), true)
    .build());

  spreadsheet.getRange('A:L').clearContent();
  spreadsheet.getRange('X2:X17').clearContent();
  spreadsheet.getRange('R8:R10').clearContent();
  spreadsheet.getRange('R2:R8').setFormulas([
    ['=iferror(match(Q2,$A$1:$L$1,false), "not assigned yet")'],
    ['=iferror(match(Q3,$A$1:$L$1,false), "not assigned yet")'],
    ['=iferror(match(Q4,$A$1:$L$1,false), "not assigned yet")'],
    ['=iferror(match(Q5,$A$1:$L$1,false), "not assigned yet")'],
    ['=iferror(match(Q6,$A$1:$L$1,false), "not assigned yet or not present")'],
    ['=if(R2 = "not assigned yet",, index(split(indirect("R2C"&R2,false),"_"),,1)&"_"&index(split(indirect("R2C"&R2,false),"_"),,2))'],
    ['=if(R6 = "not assigned yet or not present",, indirect("R2C"&R6,false))']
  ]);
}

/**
 * reads the content of the AnalyticsImport Datasheet and writes it to the Google Cloud Database.
 * 
 */
function readExternalAnalyticsData() {

  var peakLabelDict = {};
  // Dictionaries holding the information for sampleData and peakData:
  var sampleDataDict = {};
  var peakDataDict = {};
  var analyticalDataTable = [];
  var peakLabelLookupTable = [];
  var sampleNameComponents = [];
  var sampleNameValue = "";
  var uniqueSampleNames = {};
  var peakLabelValue = "";
  var area_BPValue = 0;
  var area_TotalValue = 0;
  var areaValue = 0;
  var compoundColumnValue = "";
  var retentionTimeValue = 0;
  var methodValue = "";
  var sampleCounter = 0;


  /* structure of sampleDataDict = {
      '1': {
        ID: {ELN_ID:'ELN032036-148', SAMPLE_ID: 'ELN032036-148_1_2_12.5h21C_B11_MethodID'}, 
        DATA:{PlateIngredientID:'143_RI90520202_false_Starting Material', Coordinate:'A1', ActualMass: 21.43,  DosingTimestamp:''}}
    };*/

  /* structure of peakDataDict = {
    '1': {
      ID: {ELN_ID:'ELN032036-148', SAMPLE_ID: 'ELN032036-148_1_2_12.5h21C_B11_MethodID', PEAK_ID: 1, PEAK_REF: 1}, 
      DATA:{PlateIngredientID:'143_RI90520202_false_Starting Material', Coordinate:'A1', ActualMass: 21.43,  DosingTimestamp:''}}
  };*/

  // connect to the sheet

  var analyticsImportSheet = SpreadsheetApp.getActive();

  // read columns A-N (contains the analytical data and R1:X17 (contains the assignment of compound name/category to the peak labels and the position of the required columns in columsn A-L)

  analyticalDataTable = analyticsImportSheet.getRange("A2:N").getValues();
  peakLabelLookupTable = analyticsImportSheet.getRange("R2:X17").getDisplayValues();

  if (peakLabelLookupTable[9][0] != "All good, but make sure PeakLabels are assigned!") {
    Browser.msgBox("Nice Try! Not everything filled in correctly. Aborting Script now.");
    return;
  }
  const gSheetFileId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var connector = new mssql_jdbc_api(   // connect to the database

    globalVariableDict[gSheetFileId]["DBsERVERiP"],
    globalVariableDict[gSheetFileId]["DBpORT"],
    globalVariableDict[gSheetFileId]["DBnAME"],
    globalVariableDict[gSheetFileId]["DBuSERNAME"],
    globalVariableDict[gSheetFileId]["DBpASSWORD"]);

  // fill the peakLabelDict with key/value pairs in this fashion PeakLabel : ComponentRole or SideProductName using the data from R1:X17.

  for (var row = 0; row < peakLabelLookupTable.length; row++) {
    if (peakLabelLookupTable[row][6] == "") { // if no PeakLabel is assigned, then skip that line.
      continue;
    }
    switch (peakLabelLookupTable[row][2]) {
      case "Side Product":
        peakLabelDict[peakLabelLookupTable[row][6]] = peakLabelLookupTable[row][1]; // Assign the component name as value, if it's a side product
        break;
      default:
        peakLabelDict[peakLabelLookupTable[row][6]] = peakLabelLookupTable[row][2]; // In all other cases use the component category (e.g. ligand, Internal standard, lim SM)
        break;
    }
  }

  var sampleNameColumn = peakLabelLookupTable[0][0];
  var retentionTimeColumn = peakLabelLookupTable[1][0];
  var areaColumn = peakLabelLookupTable[2][0];
  var peakLabelColumn = peakLabelLookupTable[3][0];
  var methodColumn = peakLabelLookupTable[4][0];
  var methodId = peakLabelLookupTable[6][0];
  var dateMeasured = peakLabelLookupTable[7][0];                      //The date information in R8 ends up in the DATE_FIELD in sampledata.
  var detectionModeAnalyticalDevice = peakLabelLookupTable[8][0];    //The detection mode / analytical device information ends up in the DESCRIPTION field in sampledata.



  // go through the lines of the analytical data and build the dictionaries containing the data to be written to sample- and peakdata

  for (row = 0; row < analyticalDataTable.length; row++) {
    if (String(analyticalDataTable[row][0]) == "") { // once an empty value shows up, it must be the end of the table, thus stop looping.
      break;
    }
    if (analyticalDataTable[row][areaColumn - 1] == "") continue; // if there's no area given, skip this line
    sampleNameValue = analyticalDataTable[row][sampleNameColumn - 1];//Column containing the Samplename is used to fill columns Sample_ID, ELN_ID, Platenumber, Samplenumber, RXNConditions, Platerow and Platecolumn in sampledata.

    if (!(sampleNameValue in uniqueSampleNames)) {//generate new entries in the sampleDataDictionary only if a new sample is detected.
      sampleCounter++;
      sampleDataDict[sampleCounter] = {};
      sampleDataDict[sampleCounter]["ID"] = {};
      sampleDataDict[sampleCounter]["DATA"] = {};
      uniqueSampleNames[sampleNameValue] = 1;   // Add the sampleNameValue to the dictionary to avoid duplicates in sampledata and set the row counter to 1.

      //Column containing the Samplename is used to fill columns Sample_ID, ELN_ID, Platenumber, Samplenumber, RXNConditions, Platerow and Platecolumn in sampledata.
      sampleNameComponents = sampleNameValue.split("_");


      // The method column is used to calculate the value in R7. If no method column is designated and the value in R7 overwritten, then the value in R7 is used instead (That's for cases in which no method column is present.). The method information ends up in the field INLETMETHOD in sampledata. 
      if (methodColumn == "not assigned yet or not present") { // The method column may not be present in all cases or the user may want to override it with the content of R7
        methodValue = methodId;
      } else { //when a method column is assigned, then use the content of the respective cell
        methodValue = analyticalDataTable[row][methodColumn - 1];
      }

      sampleDataDict[sampleCounter]["ID"]["ELN_ID"] = sampleNameComponents[0];
      sampleDataDict[sampleCounter]["ID"]["SAMPLE_ID"] = sampleNameValue + "_" + methodValue;
      sampleDataDict[sampleCounter]["ID"]["SAMPLE_ID"] = sampleDataDict[sampleCounter]["ID"]["SAMPLE_ID"].substr(0, 50);
      sampleDataDict[sampleCounter]["DATA"]["DATE_FIELD"] = dateMeasured;
      sampleDataDict[sampleCounter]["DATA"]["INLETMETHOD"] = methodValue;
      sampleDataDict[sampleCounter]["DATA"]["USERNAME"] = "HTE_LAB";
      sampleDataDict[sampleCounter]["DATA"]["PLATENUMBER"] = sampleNameComponents[1];
      sampleDataDict[sampleCounter]["DATA"]["SAMPLENUMBER"] = sampleNameComponents[2];
      sampleDataDict[sampleCounter]["DATA"]["RXNCONDITIONS"] = sampleNameComponents[3];
      sampleDataDict[sampleCounter]["DATA"]["PLATEROW"] = sampleNameComponents[4].substring(0, 1);
      sampleDataDict[sampleCounter]["DATA"]["PLATECOLUMN"] = sampleNameComponents[4].substring(1);


    }

    //In peakdata, Samplename is only used to fill ELN_ID and Sample_ID, Sample_ID is a combination of SampleName and Method ID in order to avoid collisions in cases where the same sample is measured using different methods. The SAMPLE_ID field is limited to 50 characters, so it needs to be truncated. 


    retentionTimeValue = analyticalDataTable[row][retentionTimeColumn - 1];

    //The three columns containing area information (2 hidden in columns M and N) are used to fill the corresponding Abs_Area, Area_BP and Area_Total columns in peakdata.
    areaValue = analyticalDataTable[row][areaColumn - 1];
    area_BPValue = analyticalDataTable[row][12]; // calculated autmatically and located in columns M and N
    area_TotalValue = analyticalDataTable[row][13];
    peakLabelValue = analyticalDataTable[row][peakLabelColumn - 1];
    // The Peak Label column is translated into the respective values from the peakLabelDict and the information ends up in the COMPOUND column in peakdata.
    if (peakLabelValue in peakLabelDict) {
      compoundColumnValue = peakLabelDict[peakLabelValue];
    } else {
      compoundColumnValue = "NULL";
    }

    peakDataDict[row + 1] = {};
    peakDataDict[row + 1]["ID"] = {};
    peakDataDict[row + 1]["DATA"] = {};

    peakDataDict[row + 1]["ID"]["ELN_ID"] = sampleNameComponents[0];
    peakDataDict[row + 1]["ID"]["SAMPLE_ID"] = sampleNameValue + "_" + methodValue;
    peakDataDict[row + 1]["ID"]["SAMPLE_ID"] = peakDataDict[row + 1]["ID"]["SAMPLE_ID"].substr(0, 50);
    peakDataDict[row + 1]["ID"]["PEAK_ID"] = uniqueSampleNames[sampleNameValue];
    peakDataDict[row + 1]["ID"]["PEAK_REF"] = uniqueSampleNames[sampleNameValue];
    uniqueSampleNames[sampleNameValue] = uniqueSampleNames[sampleNameValue] + 1;

    peakDataDict[row + 1]["DATA"]["DESCRIPTION"] = detectionModeAnalyticalDevice;
    peakDataDict[row + 1]["DATA"]["ABS_AREA"] = areaValue;
    peakDataDict[row + 1]["DATA"]["AREA_BP"] = area_BPValue;
    peakDataDict[row + 1]["DATA"]["AREA_TOTAL"] = area_TotalValue;
    peakDataDict[row + 1]["DATA"]["TIME_FIELD"] = parseFloat(String(retentionTimeValue).replace(/'/g, "."));  //some weird programs use ' as decimal point instead of .
    peakDataDict[row + 1]["DATA"]["COMPOUND"] = compoundColumnValue;
  }

  // check whether the database already contains these samples by searching in sampledata for all samples with this truncated sampleID (Sample ID without coordinate) and detection type (Description).
  // If the same number of samples as present in the current report exist already, use update, if there are none, use insert. If the number is different, warn the user and ask how to proceed. If yes, delete and update. 

  //console.log(sampleDataDict);
  //console.log(peakDataDict);

  connector.executeQuery("SELECT * FROM dbo.sampledata  WHERE ELN_ID = '" + sampleDataDict[1]["ID"]["ELN_ID"] + "' and PLATENUMBER = " + sampleDataDict[1]["DATA"]["PLATENUMBER"] + " and SAMPLENUMBER = " + sampleDataDict[1]["DATA"]["SAMPLENUMBER"] + " and INLETMETHOD = '" + sampleDataDict[1]["DATA"]["INLETMETHOD"] + "'");

  var resultsQueryLength = connector.getResultsAsArray().length;

  if (resultsQueryLength > 0) {
    if (resultsQueryLength == sampleCounter) {
      var response = Browser.msgBox("There are already " + sampleCounter + " samples in the database with the same key. Do you want to delete the previous samples and write the new ones?",
        Browser.Buttons.YES_NO);
      if (response == Browser.Buttons.YES) {
        connector.executeQuery("DELETE FROM dbo.sampledata  WHERE ELN_ID = '" + sampleDataDict[1]["ID"]["ELN_ID"] + "' and PLATENUMBER = " + sampleDataDict[1]["DATA"]["PLATENUMBER"] + " and SAMPLENUMBER = " + sampleDataDict[1]["DATA"]["SAMPLENUMBER"] + " and INLETMETHOD = '" + sampleDataDict[1]["DATA"]["INLETMETHOD"] + "'");
        connector.insertDataDictionary(globalVariableDict[gSheetFileId]["SAMPLEDATAtABLEnAME"], sampleDataDict); // write the new lines
        connector.insertDataDictionary(globalVariableDict[gSheetFileId]["PEAKDATAtABLEnAME"], peakDataDict); // write the new lines
      }
    } else {
      var response = Browser.msgBox("There are " + sampleCounter + " samples to import, but the database contains already " + connector.getResultsAsArray().length + " samples with the same key. Do you want to delete the previous samples and write the new ones?",
        Browser.Buttons.YES_NO);
      if (response == Browser.Buttons.YES) {
        connector.executeQuery("DELETE FROM dbo.sampledata  WHERE ELN_ID = '" + sampleDataDict[1]["ID"]["ELN_ID"] + "' and PLATENUMBER = " + sampleDataDict[1]["DATA"]["PLATENUMBER"] + " and SAMPLENUMBER = " + sampleDataDict[1]["DATA"]["SAMPLENUMBER"] + " and INLETMETHOD = '" + sampleDataDict[1]["DATA"]["INLETMETHOD"] + "'");
        connector.insertDataDictionary(globalVariableDict[gSheetFileId]["SAMPLEDATAtABLEnAME"], sampleDataDict); // write the new lines
        connector.insertDataDictionary(globalVariableDict[gSheetFileId]["PEAKDATAtABLEnAME"], peakDataDict); // write the new lines

      } // otherwise nothing to be done
    }
  } else {
    connector.insertDataDictionary(globalVariableDict[gSheetFileId]["SAMPLEDATAtABLEnAME"], sampleDataDict); // write the new lines
    connector.insertDataDictionary(globalVariableDict[gSheetFileId]["PEAKDATAtABLEnAME"], peakDataDict); // write the new lines
  }

  //console.log(connector.insertDataDictionary("dbo.sampledatatest", sampleDataDict));
  //console.log(connector.insertDataDictionary("dbo.peakdatatest", peakDataDict));

  connector.disconnect();

}
