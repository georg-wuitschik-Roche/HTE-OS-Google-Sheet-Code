// Abhijeet Chopra
// via https://gist.github.com/abhijeetchopra/99a11fb6016a70287112

/**
 * Macros: Google Apps Script to make copies of Google Sheet in specified destination folder, used to generate a backup of the HTE-Platform every week.
 */
function makeCopy() {

  // generates the timestamp and stores in variable formattedDate as year-month-date hour-minute-second
  var formattedDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd' 'HH:mm:ss");

  // gets the name of the original file and appends the word "copy" followed by the timestamp stored in formattedDate
  var name = SpreadsheetApp.getActiveSpreadsheet().getName() + " Copy " + formattedDate;

  // gets the destination folder by their ID. REPLACE xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx with your folder's ID that you can get by opening the folder in Google Drive and checking the URL in the browser's address bar
  var destination = DriveApp.getFolderById(globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["BACKUPFOLDERID"]);

  // gets the current Google Sheet file
  var file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());

  // makes copy of "file" with "name" at the "destination"
  file.makeCopy(name, destination);
}



/**
 * Macros: used to repair the dropdowns and conditional formattings in the tempTables and FileGenerator sheets
 */
function writeDataValidations() {
  var tempTablesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("tempTables");
  var FileGeneratorSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FileGenerator");
  var conditionalFormatRules = FileGeneratorSheet.getConditionalFormatRules();

  for (var i = -1; i < 48; i++) {

    //Writes the formulas and cell formatting in sheet tempTables column A for the queries and data validations that belong to column T in sheet FileGenerator. 
    //The formula in column A retrieves the 5 latest batches registered for the component ID in FileGenerator column T.
    /*tempTablesSheet.getRange(32+6*i, 1).setValue("Dropdown for FileGenerator T"+(i+3)+":");
    tempTablesSheet.getRange(33+6*i, 1).setFormula('=if(FileGenerator!T'+(i+3)+'="","",iferror(query(\'Batch DB\'!A2:C, "Select C where B = "&FileGenerator!T'+(i+3)+'&" order by A desc limit 5"),"Not registered"))');
                                                 
    tempTablesSheet.getRange(32+6*i, 1).setBackground('#ffffff');
    tempTablesSheet.getRange(33+6*i, 1).setBackground('#fff2cc');
    tempTablesSheet.getRange(34+6*i, 1,4,1).setBackground('#d9ead3');
    
    FileGeneratorSheet.getRange(i+3,25).setDataValidation(SpreadsheetApp.newDataValidation()
      .setAllowInvalid(true)
      .requireValueInRange(FileGeneratorSheet.getRange('tempTables!$A$'+(33+6*i)+':$A$'+(37+6*i)), false)
      .build());
    
    FileGeneratorSheet.getRange(i+3,25).setFormula("=tempTables!$A$"+(33+6*i))
    
    
    //Writes the formulas and cell formatting in sheet tempTables column B for the queries and data validations that belong to column AM in sheet FileGenerator.
    //The formula in column B retrieves the 5 latest batches registered for the component ID in FileGenerator column AM.
    tempTablesSheet.getRange(32+6*i, 2).setValue("Dropdown for FileGenerator AM"+(i+3)+":");
    
    tempTablesSheet.getRange(33+6*i, 2).setFormula('=if(FileGenerator!AL'+(i+3)+'="","",iferror(query(Placeholder!A2:F, "Select C where F = \'"&FileGenerator!AL'+(i+3)+'&"\' order by A desc limit 5"),"Not registered"))');
    tempTablesSheet.getRange(32+6*i, 2).setBackground('#ffffff');
    tempTablesSheet.getRange(33+6*i, 2).setBackground('#fff2cc');
    tempTablesSheet.getRange(34+6*i, 2,4,1).setBackground('#d9ead3');
    
    FileGeneratorSheet.getRange(i+3,39).setDataValidation(SpreadsheetApp.newDataValidation()
      .setAllowInvalid(true)
      .requireValueInRange(FileGeneratorSheet.getRange('tempTables!$B$'+(33+6*i)+':$B$'+(37+6*i)), false)
      .build());
    
    FileGeneratorSheet.getRange(i+3,39).setFormula("=tempTables!$B$"+(33+6*i))
   
    //Writes the formulas and cell formatting in sheet tempTables column C for the queries and data validations that belong to column AI in sheet FileGenerator. 
    //The formula in column C retrieves up to 5 solution IDs registered for the component ID in FileGenerator column T.
    
    tempTablesSheet.getRange(32+6*i, 3).setValue("Dropdown for FileGenerator AI"+(i+3)+":");
    tempTablesSheet.getRange(33+6*i, 3).setFormula('=if(FileGenerator!AH'+(i+3)+' <> "Solution","",iferror( query(Solutions!A2:C,"Select A where (B = "& FileGenerator!T'+(i+3)+'&" and C = \'" &FileGenerator!Y'+(i+3)+' & "\') limit 5")))');
                                                   
    tempTablesSheet.getRange(32+6*i, 3).setBackground('#ffffff');
    tempTablesSheet.getRange(33+6*i, 3).setBackground('#fff2cc');
    tempTablesSheet.getRange(34+6*i, 3,4,1).setBackground('#d9ead3');
    
    FileGeneratorSheet.getRange(i+3,35).setDataValidation(SpreadsheetApp.newDataValidation()
      .setAllowInvalid(true)
      .requireValueInRange(FileGeneratorSheet.getRange('tempTables!$C$'+(33+6*i)+':$C$'+(37+6*i)), false)
      .build());
    
    FileGeneratorSheet.getRange(i+3,35).setFormula("=tempTables!$C$"+(33+6*i)) 
   
  */
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
 * Count the number of cells of the active spreadsheet, used to find out how close the sheet is to the cell limit.
 *
 * @param {A1} input Used to force recalculation.
 * @customfuncion
 */
function cellsCount() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var cells = 0;
  sheets.forEach(function (sheet) {
    cells = cells + sheet.getMaxRows() * sheet.getMaxColumns();
    console.log(sheet.getName() + ": " + sheet.getMaxRows() * sheet.getMaxColumns() + " cells");
  });
  console.log("Total number of cells is " + cells);
}

// helper function to create folder structure and generate the content of global Variables.gs
function createFoldersAndGlobalVariablesString() {

  var globalVariablesString = 'var globalVariableDict =    { "platformSheetId": {\n        // This is the id of the root folder named "HTE Platform". All the other folders this tool is working with are in this folder.\n        PROJECTdATAfOLDERiD: "pROJECTdATAfOLDERiD",\n        // This folder named "P-Touch Labels" contains all the csv-files for printing P-Touch labels and csl-files for writing heads that are generated by the Registration sheet when registering a new batch/chemical.\n        PtOUCHlABELSfOLDERiD: "ptOUCHlABELSfOLDERiD",\n        //Folder in which all presentations are located\n        PRESENTATIONfOLDERiD: "pRESENTATIONfOLDERiD",\n        //Folder containing the Cheatsheets:\n        CHEATSHEETSfOLDERiD: "cHEATSHEETSfOLDERiD",\n        //Folder containing the LCMS sequence files\n        LCMSfOLDERiD: "lCMSfOLDERiD",\n        // Folder containing the Agilent sequence files\n        AGILENTfOLDERiD: "aGILENTfOLDERiD",\n        // Folder containing the Junior input files\n        JUNIORfOLDERiD: "jUNIORfOLDERiD",\n        // Folder containing Excel-Files with data from cELN:\n        CELNfOLDERiD: "cELNfOLDERiD",\n        // Folder where backup copies of the Google Sheet end up that are created automatically using makeCopy (remember to setup a daily trigger to run the backup)\n        BACKUPFOLDERID: "your backup folder ID goes here",\n          // Google Slides Presentation Template ID:\n        PRESENTATIONtEMPLATEiD: "replace with yours",\n        // Link to Spotfire (without ELNID and final, you need to replace it with your link. It is used to create a ELN-ID specific link for the gSheet presentation\n        SPOTFIRElINK: "https://<<Your Spotfire Server>>/spotfire/wp/OpenAnalysis?file=<<ID of your Spotfire File>>&wavid=0&configurationBlock=ELNSelect%3D%22",\n        //Folder containing the Quantos input files\n        QUANTOSfOLDERiD: "qUANTOSfOLDERiD",\n      //file id of the Quantos hotel xml-file which is synchronized using a batch file onto gDrive\n      // on the computer running Quantos Chronect, you need to locate the file called "Quantos_DosingHeadRacks.xml" and sync it to Google Drive using the \n      // backup function of the Google-Drive client. Share that file in gDrive with everybody who needs to work with the sheet.\n       QUANTOShOTELfILEiD: "put id of the xml-file here",        // This folder named "Quantos Dosing Results" contains a synced version of the dosing results by Quantos. It is filled by a robocopy from the Quantos Computer and then synced using Google File Stream.\n        QUANTOSdOSINGrESULTSiD: "replace with yours",\n        //This is the folder named "Quantos Correction Dosings" containing the correction dosing files generated in the correction sheet.\n        QUANTOScORRECTIONdOSINGiD: "qUANTOScORRECTIONdOSINGiD",\n        //URL of the Webhook used to send ChatMessages when plates are designed or reactions are registered:\n        CHATMESSAGEwEBHOOK: "replace with yours",\n        //Webhooks needed for communicating the status of Quantos Chronect to the users:\n        HEADeMPTYqUANTOShEARTBEAT: "replace with yours",\n        ISSUEwEBHOOKqUANTOShEARTBEAT: "replace with yours",	\n        PROGRESSwEBHOOK	: "replace with yours",\n        SUCCESSwEBHOOKqUANTOShEARTBEAT: "replace with yours",\n        //These are needed to connect to the database for writing and reading data that was so far stored in the gSheet in tables like Wells...\n        DBsERVERiP: "xxx.xxx.xxx.xxx",\n        DBnAME: "name of your database",\n        DBuSERNAME: "username to access the database", // needs to have read/write privileges\n        DBpASSWORD: "password to access the database",\n        DBpORT: 1433, // standard port to access the MS SQL server, may have to change that one\n        WELLStABLEnAME: "dbo.wells_prod",  // if you change the names of the data tables, you will also have to modify them in Spotfire\n        COLlABELStABLEnAME: "dbo.ColLabels_prod",\n        ROWlABELStABLEnAME: "dbo.RowLabels_prod",\n        SAMPLEDATAtABLEnAME: "dbo.sampledata",\n        PEAKDATAtABLEnAME: "dbo.peakdata",\n    }\n};\n// Constant for all versions of the sheet\n// Link to the API for converting smiles images, inchi etc \nconst LINKtOfASTaPI = "<<link to your API>>";\n// API Password for retrieving structures\nconst FASTaPIkEY = "Password needed to access the Chemical Translator API";    //needs to be set for your instance of the Chemical Translator API\n// In order to search for CAS-numbers, you need to get your own custom (free) Google Search Engine: https://developers.google.com/custom-search/docs/tutorial/introduction\nconst ApiKey = "<<Your API Key>>";\nconst SearchEngineID = "<<Your Search Engine ID>>";\n';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  const getStartedSheet = ss.getSheetByName("How to get started");
  var platformSheetId = ss.getId();
  var file = DriveApp.getFileById(platformSheetId);
  var folders = file.getParents();
  var idOfCurrentFolder = "";
  while (folders.hasNext()) {
    idOfCurrentFolder = folders.next().getId();
    console.log(idOfCurrentFolder);
  }

  var parentFolder = DriveApp.getFolderById(idOfCurrentFolder);

  var pROJECTdATAfOLDERiD = idOfCurrentFolder;
  var ptOUCHlABELSfOLDERiD = parentFolder.createFolder("pTouch Labels").getId();
  var pRESENTATIONfOLDERiD = parentFolder.createFolder("Presentations").getId();
  var cHEATSHEETSfOLDERiD = parentFolder.createFolder("Cheatsheets").getId();
  var lCMSfOLDERiD = parentFolder.createFolder("LCMS Sequence Files").getId();
  var aGILENTfOLDERiD = parentFolder.createFolder("Agilent Sequence Files").getId();
  var jUNIORfOLDERiD = parentFolder.createFolder("Junior Input Files").getId();
  var cELNfOLDERiD = parentFolder.createFolder("cELN imports").getId();
  var qUANTOSfOLDERiD = parentFolder.createFolder("Quantos Input Files").getId();
  var qUANTOScORRECTIONdOSINGiD = parentFolder.createFolder("Quantos Correction Dosing Files").getId();


  globalVariablesString = globalVariablesString
    .replace("platformSheetId", platformSheetId)
    .replace("pROJECTdATAfOLDERiD", pROJECTdATAfOLDERiD)
    .replace("ptOUCHlABELSfOLDERiD", ptOUCHlABELSfOLDERiD)
    .replace("pRESENTATIONfOLDERiD", pRESENTATIONfOLDERiD)
    .replace("cHEATSHEETSfOLDERiD", cHEATSHEETSfOLDERiD)
    .replace("lCMSfOLDERiD", lCMSfOLDERiD)
    .replace("aGILENTfOLDERiD", aGILENTfOLDERiD)
    .replace("jUNIORfOLDERiD", jUNIORfOLDERiD)
    .replace("cELNfOLDERiD", cELNfOLDERiD)
    .replace("qUANTOScORRECTIONdOSINGiD", qUANTOScORRECTIONdOSINGiD)
    .replace("qUANTOSfOLDERiD", qUANTOSfOLDERiD);


  console.log(globalVariablesString);
  getStartedSheet.getRange("F4").setValue(globalVariablesString);


}
