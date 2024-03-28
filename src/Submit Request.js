/*jshint sub:true*/
// this function takes the content of the Excel-file imported from cELN and writes it into the Submit Request Sheet

/**
 * Submit Request: import the data contained in the Excel file exported from the ELN.
 * @param {Array} cElnFileContent cElnFileContent is an array that contains the file content as an array as first element and the file name as second element.
 */
function importCelnContent(cElnFileContent) {

  var reactionComponents = {};  // This dictionary will contain all the information extracted from the Excel sheet
  var emptyRowCounter = 0;      // is increased by one once an empty row is encountered, used to figure out when the product section begins (first empty row found) and when the end of the table is reached (second empty row)
  var inchiData = ["", ""];
  var formula = "";

  // connect to the Submit Request Sheet
  var submitRequestSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Submit Request");

  // read the second line of the cElnFileContent and put it into a dictionary that is constructed in analogy to the one used in newExperimentSubmit

  reactionComponents[0] =        // all other components are positive integers starting with 1
  {
    "StepName": cElnFileContent[1][2],
    "ProjectName": cElnFileContent[1][7],
    "ElnId": cElnFileContent[1][1],
    "Theme": cElnFileContent[1][5],
    "intElnId": cElnFileContent[1][0]
  };

  // read the content from line 5 onward which contains the starting material information and separated by an empty line the (side) product table in analogy to newExperimentSubmit

  for (var row = 4; row < cElnFileContent.length; row++) {
    if (cElnFileContent[row][0] == "") {
      emptyRowCounter++; //either end of the reactant section or the product section

      if (emptyRowCounter > 1) { break; } else {// exit the loop after the end of the product table or skip the empty line and the header line after the end of the reactant table
        row++; // first line of the product section contains headers which are skipped this way
        continue;
      }
    }
    // depending on the state of the empty row counter decide whether the row in question contains a reactant or (side)product
    switch (emptyRowCounter) {    // if 0, then in the starting materials table, if 1 in the (side)products table
      case 0:          //starting materials
        if (cElnFileContent[row].length > 27 && cElnFileContent[row][28].split("InChIKey=").length > 1) {   // before the Excel-files contained formula and inchi(key), the array containing the data was only 26 columns broad, also it needs to be made sure that both inchi and inchi key are present in column 28. 

          formula = cElnFileContent[row][27];
          inchiData = cElnFileContent[row][28].split("InChIKey=");
        } else {
          formula = "add manually for now";
          inchiData = ["Generation failed, add manually :(", "Generation failed, add manually :("];
        }
        if (cElnFileContent[row][22] == 0) { cElnFileContent[row][22] = "-"; }                                                               // If no density is provided, then a 0 will appear in the Excel File which needs to be converted to a "-" (the gsheet will treat compounds with a "-" as solids afterwards)
        if (cElnFileContent[row][23] == 0) { cElnFileContent[row][23] = ""; } else { cElnFileContent[row][22] = parseFloat(cElnFileContent[row][22]) / 100; } // If no purity is provided, then a 0 will appear in the Excel File which needs to be converted to an empty string (the gsheet will treat an empty string as 99.9% purity during registration) apparently the Zero gets often imported into the Sheet?! To be corrected! 0% purity is dangerous!
        reactionComponents[2 * (row - 3) - 0] =    // Use the column number (0-indexed) into which the corresponding starting material should go as key (Starting Material 1 --> 1, Starting Material 2 --> 3...)
        {
          "InchiKey": inchiData[1],
          "Inchi": inchiData[0],
          "ComponentName": "= iferror(vlookup(" + columnToLetter(2 * (row - 3) - 0) + "2,'Component DB'!$A:$C, 3,false),\"Please set\")",                             // cElnFileContent[row][4] before, but IUPAC name and mostly too long
          "RocheNumber": cElnFileContent[row][15],
          "Smiles": cElnFileContent[row][26].replace(/\|.+\|/, ""), // Some Smiles strings contain a racemic or stereoflag (e.g.  |r,&1:2,25,41| or  |r,&1:1|) at the end which is upsetting Spotfire. This regex removes these 
          "CAS": cElnFileContent[row][18],
          "MW": cElnFileContent[row][19],
          "Density": cElnFileContent[row][22],
          "MolecularFormula": formula,
          "BatchId": cElnFileContent[row][17],
          "Producer": cElnFileContent[row][16],
          "Assay": cElnFileContent[row][23] / 100
        };
        break;
      case 1: // (side)products
        if (cElnFileContent[row][16].split("InChIKey=").length > 1) { inchiData = cElnFileContent[row][16].split("InChIKey="); } else { inchiData = ["Generation failed, add manually :(", "Generation failed, add manually :("]; }
        reactionComponents[2 * (cElnFileContent[row][1]) + 9] =    // Use the column number (0-indexed) into which the corresponding product should go as key (Product --> 10, Side Product 1 --> 12...)
        {
          "InchiKey": inchiData[1],
          "Inchi": inchiData[0],
          "ComponentName": "= iferror(vlookup(" + columnToLetter(2 * (cElnFileContent[row][1]) + 9) + "2,'Component DB'!$A:$C, 3,false),\"Please set\")", // cElnFileContent[row][6] before, but IUPAC name and mostly too long
          "RocheNumber": "",
          "Smiles": cElnFileContent[row][15].replace(/\|.+\|/, ""), // Some Smiles strings contain a racemic or stereoflag (e.g.  |r,&1:2,25,41| or  |r,&1:1|) at the end which is upsetting Spotfire. This regex removes these 
          "CAS": "",
          "MW": cElnFileContent[row][5],
          "Density": "-",
          "MolecularFormula": cElnFileContent[row][7],
          "BatchId": "",
          "Producer": "",
          "Assay": ""
        };
        break;
    }

  }

  //push the data of the combined dictionary into the Submit Request Sheet
  let reactingFunctionalGroupProduct = "";    // needs to be "-" for the product, so the user doesn't have to enter it manually. 
  for (var key in reactionComponents) {
    if (key > 0) { //starting materials and (side)products
      if (key == 11) { reactingFunctionalGroupProduct = "-"; } else { reactingFunctionalGroupProduct = ""; };
      submitRequestSheet.getRange(3, key, 8, 1).setValues([
        //[reactionComponents[key].InchiKey],   //now calculated using the API
        [reactionComponents[key].ComponentName],
        [reactionComponents[key].RocheNumber],
        ['=hyperlink("http://smiles-ds.marathon.bahpc.roche.com:3020/depict/500/500/"&encodeurl("' + reactionComponents[key].Smiles + '"),"' + reactionComponents[key].Smiles + '")'],
        ['= if(' + columnToLetter(key) + '5<>"", IMPORTDATA("' + LINKtOfASTaPI + '/smiles-to-inchi-plain/' + FASTaPIkEY + '/"& encodeurl(' + columnToLetter(key) + '5),"ç"),iferror(vlookup(' + columnToLetter(key) + '2,\'Component DB\'!$A$2:$K,8,false),if(E7="","",hyperlink("https://www.google.com/search?q="&substitute(' + columnToLetter(key) + '7, " ","+")&"+inchi",iferror(IMPORTXML("https://cactus.nci.nih.gov/chemical/structure/"&' + columnToLetter(key) + '7&"/inchi/xml","//item[@id=\'1\']"),"Not found")))))'], //[reactionComponents[key].Inchi],
        [reactionComponents[key].CAS],
        [reactionComponents[key].MW],
        [reactionComponents[key].Density],
        [reactingFunctionalGroupProduct] //,
        //[reactionComponents[key].MolecularFormula],
      ]);

      if (key < 9 && reactionComponents[key].BatchId != "") { // write Batch/Producer/Assay only if it's a starting material (key <9) and both batch ID and Producer are not empty strings
        submitRequestSheet.getRange(14, key, 3, 1).setValues([
          [reactionComponents[key].BatchId],
          [reactionComponents[key].Producer],
          [reactionComponents[key].Assay]
        ]);
      }
    } else {   // key = 0, that is other information like ELN-ID etc needs to be written.
      submitRequestSheet.getRange(20, 2).setValue(reactionComponents[0].ElnId);
      submitRequestSheet.getRange(2, 10).setValue(reactionComponents[0].StepName);
      submitRequestSheet.getRange(21, 10).setValue(reactionComponents[0].ProjectName);
      submitRequestSheet.getRange(21, 13).setValue(reactionComponents[0].Theme);
      submitRequestSheet.getRange(23, 13).setValue(reactionComponents[0].intElnId);
    }
  }
}

/**
 * Submit Request: displays the html-form for uploading the Excel-file exported from the ELN.
 * based on https://www.youtube.com/watch?v=U9JFn30b-PY
 */
function showDialogue() {                                                     //for show the msg to the user
  var template = HtmlService.createTemplateFromFile("fileUploadDialogue").evaluate();   //file html
  SpreadsheetApp.getUi().showModalDialog(template, "File Upload");           //Show to user, add title
}

/**
 * Submit Request: This function takes the file info received from the fileupload dialogue, triggers the import and returns the URL to be displayed in the upload dialogue
 * based on https://www.youtube.com/watch?v=U9JFn30b-PY
 * @param {String} data Sample text.
 * @param {String} name Sample text.
 * @param {String} type Sample text.
 * @return {String} file URL of the imported Excel file.
 */
function uploadFilesToGoogleDrive(data, name, type) {                          //function to call on front side
  var datafile = Utilities.base64Decode(data);                               //decode data from Base64
  var blob2 = Utilities.newBlob(datafile, type, name);                      //create a new blob with decode data, name, type
  var folder = DriveApp.getFolderById(globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["CELNfOLDERiD"]); //Get folder of destination for file (final user need access before execution)
  var newFile = folder.createFile(blob2);                                   //Create new file (property of final user)

  var rowData = [                                                          //for print results
    newFile.getName(),
    newFile.getId(),
    newFile.getUrl(),
    newFile.getSize(),
    newFile.getDateCreated()
  ];

  var fileContent = getAndAppendData(newFile.getId());    // function is located in Correction.gs, an array is returned that contains the file content as an array as first element and the file name as second element. 


  importCelnContent(fileContent[0]);

  return newFile.getUrl();                                                   //Return URL
}
/**
 * Submit Request: register a new experiment and calls the function new experiment submit, handing over the number of the column at which the import should commence. Connected to the corresponding button.
 */
function buttonExperimentSubmission() {
  newExperimentSubmit(1);
}
/**
 * Submit Request: register new sideproducts and calls the function new experiment submit, handing over the number of the column at which the import should commence. Connected to the corresponding button.
 */
function buttonRegisterSideProduct() {
  newExperimentSubmit(12);
}
/**
 * Submit Request: register either a new experiment or new side products, indirectly connected to the corresponding buttons.
 */
function newExperimentSubmit(startingColumn = 1) {
  // connect to sheets
  var submitRequestSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Submit Request");

  var componentDbSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Component DB");
  var hteRequestSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HTE-Requests");
  var componentRolesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Component Roles");
  var smProdSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SMs/Prods");
  var batchDbSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Batch DB");

  var componentDbLastRow = componentDbSheet.getLastRow();
  var componentIDs = componentDbSheet.getRange(2, 2, componentDbLastRow - 1).getValues();
  componentIDs = [].concat.apply([], componentIDs);

  // variables used
  var newBatches = [];
  var newComponents = [];
  var newRoles = [];
  var newSmsProds = [];
  var densityOrHeadId = "";


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

  var newBatchLabels = [[
    "Component Role",
    "Component ID",
    "Component Name",
    "CAS",
    "Roche-No.",
    "Molecular Formula",
    "Batch ID",
    "Producer",
    "Assay",
    "MW",
    "Dosing Head ID or density",
    "Date"]];   // later filled with information for the p-Touch Labels
  var componentRoleFlag = 0;
  var componentIdOfKey = 0;
  var reactionComponents = {};

  //find out the what the last row is of the relevant sheets
  var hteRequestlastRow = hteRequestSheet.getLastRow();
  var knownElnIds = hteRequestSheet.getRange(2, 1, hteRequestlastRow - 1, 1).getValues();
  knownElnIds = [].concat.apply([], knownElnIds);                                            // remove inner brackets, so that indexOf function can be used to figure out if an ELN-ID is known or not
  var lastRowSMsProds = smProdSheet.getLastRow();
  var lastRowRoles = componentRolesSheet.getLastRow();
  var lastRowBatch = batchDbSheet.getLastRow();
  var componentDblastRow = componentDbSheet.getLastRow();

  var lastRegisteredComponentId = parseInt(componentDbSheet.getRange(componentDblastRow, 2).getValue()); // this assumes that the last filled cell in column B in the Component DB sheet containst the component ID with the highest number 


  // read relevant Submit request sheet data
  var Content = submitRequestSheet.getRange(1, 1, 25, 21).getValues();

  if (knownElnIds.indexOf(Content[19][1]) != -1 && startingColumn == 1) { // only true, if the ELNID exists already
    Browser.msgBox(Content[19][1] + " already exists in the HTE-Requests Sheet. Please correct. Aborting script now.");
    return;
  }

  // Check if everything is filled in correctly by checking whether cell B22 contains "Hit that button"
  if ((Content[21][1] != "Hit that Button!" && startingColumn == 1) || (Content[23][16] != "Ready!" && startingColumn == 12)) {
    Browser.msgBox("Nice try! Not all fields are filled out correctly.");
    return;
  }


  //get/create folder IDs
  //create a new folder with the ELN-ID, Project-, Step- and Customer-name as name on gDrive under HTS Docs > Robot Input Files unless it exists already

  //depreciated, data is now sorted 
  //var folderID = createFolder(globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["PROJECTdATAfOLDERiD"], Content[19][1] + "_" + Content[20][9] + "_" + Content[1][9] + "_" + Content[22][9])
  //var supportingDocsFolderId = createFolder(folderID, "Supporting Docs-" + String(Content[19][1]).substring(Content[19][1].length - 3, Content[19][1].length) + "_" + Content[20][9] + "_" + Content[1][9] + "_" + Content[22][9]); // creates the supporting docs folder - not used anymore since presentations are stored centrally
  var quantosFolder = DriveApp.getFolderById(globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["QUANTOSfOLDERiD"]);
  var batchLabelsFolder = DriveApp.getFolderById(globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["PtOUCHlABELSfOLDERiD"]);

  // read the Component Roles data, needed in order to avoid assigning duplicate component roles in case a reaction component shows up again in a different reaction
  var componentRolesSheetData = componentRolesSheet.getRange(2, 1, lastRowRoles, 3).getValues();

  // go through the Content array and fill the data into a dictionary
  for (var column = startingColumn; column < Content[0].length; column++) {
    if (Content[10][column].length < 2 && column != 9) { continue; } // rejects empty columns, happens only if there's a molecular formula with less than two letters or a check column with one 🐇

    switch (column) {
      case 1: case 3: case 5: case 7:          //starting materials
        reactionComponents[Content[2][column] + "_sm"] =    // Component Name serves as Dictionary key, assumes that it's unique within one reaction
        {
          "InfoType": "Starting Material",
          "InchiKey": Content[1][column],
          "ComponentName": Content[2][column],
          "RocheNumber": Content[3][column],
          "Smiles": Content[4][column].replace(/\|.+\|/, ""), // Some Smiles strings contain a racemic or stereoflag (e.g.  |r,&1:2,25,41| or  |r,&1:1|) at the end which is upsetting Spotfire. This regex removes these 
          "Inchi": Content[5][column],
          "CAS": Content[6][column],
          "MW": Content[7][column],
          "Density": Content[8][column],
          "ReactingFG": Content[9][column],
          "MolecularFormula": Content[10][column],
          "CompoundStatus": Content[12][column], // "Compound found in DB" if it's an existing compound in Component DB,"New Compound in DB" if not
          "BatchId": Content[13][column],
          "Producer": Content[14][column],
          "Assay": Content[15][column],
          "BatchStatus": Content[16][column],          // Is the entered batch known ("Batch found in DB") already or not ("").
          "CurrentComponentId": Content[17][column],  // If it's a known compound, Content[17][column] will contain its component ID, if not it's empty.
          "SideReactionType": Content[3][9]      // for starting materials and product, this field contains the intended reaction    
        };
        break;
      case 9:                                         // general Info
        //contains Step Name [0], Reaction Type [1], Customer Guidance [2], Project [3] Customer Department [4], Customer Name [5], Submission Date [6], ELN-Number [7], Theme
        reactionComponents["otherInfo"] =        // assumes that no component will ever be named "otherInfo"
        {
          "StepName": Content[1][column],
          "ReactionType": Content[3][column],
          "CustomerGuidance": Content[7][column],
          "ProjectName": Content[20][column],
          "Department": Content[21][column],
          "Customer": Content[22][column],
          "SubmissionDate": Content[23][column],
          "ElnId": Content[19][1],
          "Theme": Content[20][12],
          "intElnId": Content[22][12]

        };
        break;
      case 10:         //product
        reactionComponents[Content[2][column] + "_prod"] = // Component Name serves as Dictionary key, assumes that it's unique within one reaction
        {
          "InfoType": "Product",
          "InchiKey": Content[1][column],
          "ComponentName": Content[2][column],
          "RocheNumber": Content[3][column],
          "Smiles": Content[4][column].replace(/\|.+\|/, ""), // Some Smiles strings contain a racemic or stereoflag (e.g.  |r,&1:2,25,41| or  |r,&1:1|) at the end which is upsetting Spotfire. This regex removes these 
          "Inchi": Content[5][column],
          "CAS": Content[6][column],
          "MW": Content[7][column],
          "Density": Content[8][column],
          "ReactingFG": "-",
          "MolecularFormula": Content[10][column],
          "CompoundStatus": Content[12][column], // "Compound found in DB" if it's an existing compound in Component DB,"New Compound in DB" if not
          "BatchId": Content[13][column],
          "Producer": Content[14][column],
          "Assay": Content[15][column],
          "BatchStatus": Content[16][column],           // Is the entered batch known ("Batch found in DB") already or not (""). 
          "CurrentComponentId": Content[17][column],   // If it's a known compound, Content[17][column] will contain its component ID, if not it's empty.
          "SideReactionType": Content[3][9]       // for starting materials and product, this field contains the intended reaction   
        };
        if (Content[2][column] + "_sm" in reactionComponents) {
          reactionComponents[Content[2][column] + "_prod"].CompoundStatus = "Same as one of the starting materials";
        }
        break;
      case 12: case 14: case 16: case 18: case 20:   // side products
        reactionComponents[Content[2][column] + "_sideprod"] =
        {
          "InfoType": "Side Product",
          "InchiKey": Content[1][column],
          "ComponentName": Content[2][column],
          "RocheNumber": Content[3][column],
          "Smiles": Content[4][column].replace(/\|.+\|/, ""), // Some Smiles strings contain a racemic or stereoflag (e.g.  |r,&1:2,25,41| or  |r,&1:1|) at the end which is upsetting Spotfire. This regex removes these 
          "Inchi": Content[5][column],
          "CAS": Content[6][column],
          "MW": Content[7][column],
          "Density": Content[8][column],
          "ReactingFG": Content[9][column],
          "MolecularFormula": Content[10][column],
          "CompoundStatus": Content[12][column], // "Compound found in DB" if it's an existing compound in Component DB,"New Compound in DB" if not
          "SideReactionType": Content[13][column],
          "CurrentComponentId": Content[17][column]
        };
        if (Content[2][column] + "_sm" in reactionComponents) {
          reactionComponents[Content[2][column] + "_prod"].CompoundStatus = "Same as one of the starting materials";
        }
        break;
    }
  }

  if (startingColumn == 12) {   // in case a sideProduct is registered, the user has to choose in cell Q19 the ELN-ID for which the reaction should be registered
    reactionComponents["otherInfo"] = { "ElnId": Content[18][16] };
  }

  // sort data into arrays ready for writing to the different sheets or into files

  for (var key in reactionComponents) {


    if (key == "otherInfo") { continue; }
    if (reactionComponents[key].CompoundStatus == "New Compound in DB") {  //if it's a new component

      lastRegisteredComponentId++;
      componentIdOfKey = lastRegisteredComponentId;
      reactionComponents[key].CurrentComponentId = componentIdOfKey;
      if (reactionComponents[key].InfoType != "Side Product") { //Side products are not registered in the roles sheet, since they wouldn't be used as inputs in plate builder
        newRoles.push([lastRegisteredComponentId, reactionComponents[key].ComponentName, "SM/Prod"]);
        componentRolesSheetData.push([lastRegisteredComponentId, reactionComponents[key].ComponentName, "SM/Prod"]);
      }
      newComponents.push([
        reactionComponents[key].InchiKey,
        componentIdOfKey,
        reactionComponents[key].ComponentName,
        reactionComponents[key].MW,
        reactionComponents[key].CAS,
        reactionComponents[key].RocheNumber,
        reactionComponents[key].Smiles,
        reactionComponents[key].Inchi,
        reactionComponents[key].Density,
        '',
        reactionComponents[key].MolecularFormula]);

    } else {// if the compound is not new
      if (reactionComponents[key].CompoundStatus == "Same as one of the starting materials") {
        reactionComponents[key].CurrentComponentId = reactionComponents[reactionComponents[key].ComponentName + "_sm"].CurrentComponentId;
      }
      componentIdOfKey = reactionComponents[key].CurrentComponentId;
      if (reactionComponents[key].InfoType != "Side Product") { //Side Products are not written to the Component Roles sheet, since they won't be used in PlateBuilder
        for (var row = 0; row < componentRolesSheetData.length; row++) { //look whether the compound is already registered as SM/Prod in Component Roles to avoid duplicates, since it's not new
          if (componentRolesSheetData[row][0] == componentIdOfKey && componentRolesSheetData[row][2] == "SM/Prod") {
            componentRoleFlag = 1; // set the flag to prevent writing a duplicate role
            break;
          }
        }
        if (componentRoleFlag == 0) { //if it's 1, the role exists for this compound already 
          componentRolesSheetData.push([reactionComponents[key].CurrentComponentId, reactionComponents[key].ComponentName, "SM/Prod"]);
          newRoles.push([reactionComponents[key].CurrentComponentId, reactionComponents[key].ComponentName, "SM/Prod"]);
        }
        componentRoleFlag = 0; // resets the flag to 0
      }
    }

    tmpName = reactionComponents[key].ComponentName;


    if ((isNaN(parseFloat(reactionComponents[key].Density)) == true || reactionComponents[key].Density == "") &&
      reactionComponents[key].InfoType == "Starting Material") {  // only true, if it's a solid and a Starting material (Products and side products don't get dosed on Quantos)
      quantosHeads.push([
        quantosHeads.length,
        "C:\\Users\\Public\\Documents\\Chronos\\Methods\\HeadWrite in Sequence.cam",
        "Heads",
        quantosHeads.length,
        componentIdOfKey + "@" + String(reactionComponents[key].BatchId).substring(String(reactionComponents[key].BatchId).length - 14),
        tmpName,
        6666,     //filling amount, to be corrected by the user
        "",
        "",
        999,   // dosing limit, i.e how many times the head can be dosed, 999 is the maximum
        "True", // Tap before dosing
        40,     // Tapping intensity
        2]);     // Tapping duration

    }
    newSmsProds.push([
      reactionComponents["otherInfo"].ElnId,
      tmpName,
      componentIdOfKey,
      reactionComponents[key].InfoType,
      reactionComponents[key].ReactingFG,
      reactionComponents[key].SideReactionType]);

    if (reactionComponents[key].BatchId != "" &&
      reactionComponents[key].Producer != "" &&
      reactionComponents[key].InfoType != "Side Product") { // if it's a new compound and both BatchID and Producer and if it's not a side product were provided, put this batch into the BatchDB
      newBatches.push([lastRowBatch + 1,
        componentIdOfKey,
      reactionComponents[key].BatchId,
      reactionComponents[key].Producer,
      reactionComponents[key].Assay,
        tmpName,
      componentIdOfKey + "_" + reactionComponents[key].BatchId // Batch Key
      ]);
      if (isNaN(parseFloat(reactionComponents[key].Density)) == true || reactionComponents[key].Density == "") {  // only true, if it's a solid
        densityOrHeadId = componentIdOfKey + "@" + String(reactionComponents[key].BatchId).substring(String(reactionComponents[key].BatchId).length - 14);  //Dosing Head ID
      } else { // compound is a liquid
        densityOrHeadId = reactionComponents[key].Density; //density
      }
      newBatchLabels.push([
        reactionComponents[key].InfoType,    //Starting Material or Product
        componentIdOfKey,
        (tmpName).toString().replace(',', ';'),    // , is the separator in the csv, thus it needs to be replaced if present
        reactionComponents[key].CAS,
        reactionComponents[key].RocheNumber,
        reactionComponents[key].MolecularFormula,
        reactionComponents[key].BatchId,
        reactionComponents[key].Producer,
        reactionComponents[key].Assay * 100 + "%",
        reactionComponents[key].MW,
        densityOrHeadId,    //Dosing Head ID for solids, density for liquids
        Utilities.formatDate(new Date(), 'GMT+1', "MMM d ''yy")]);
    }
  }



  // ************** Write data to the different sheets and fill in the Google Slides presentation *******************



  if (startingColumn == 1) { // only if a new experiment is registered, data should be written to the HTE-Requests sheet, the batch DB sheet and the Component Roles sheet. 
    //Also, no files need to be written if side products are registered. 

    var presentationLink = fillPresentationTemplate(reactionComponents, [], "none");  //function located in FileGenerator.gs

    var cheatSheetFileId = createSpreadsheet(globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["CHEATSHEETSfOLDERiD"], reactionComponents["otherInfo"].ElnId + "_Plates");

    // write data to HTE-Requests Sheet
    hteRequestSheet.getRange(hteRequestlastRow + 1, 1, 1, 12).setValues([[
      reactionComponents["otherInfo"].ElnId,
      reactionComponents["otherInfo"].ProjectName,
      reactionComponents["otherInfo"].StepName,
      reactionComponents["otherInfo"].ReactionType,
      reactionComponents["otherInfo"].Customer,
      reactionComponents["otherInfo"].Department,
      reactionComponents["otherInfo"].CustomerGuidance,
      reactionComponents["otherInfo"].SubmissionDate,
      reactionComponents["otherInfo"].Theme,
      presentationLink,
      DriveApp.getFileById(cheatSheetFileId).getUrl(),
      reactionComponents["otherInfo"].intElnId]]);

    SpreadsheetApp.flush();
    hteRequestSheet.getFilter().sort(1, false); //reset the sorting of the filter, so the new line correctly ends up at the top. 
    // send out the mail to get a card created in our HTE-Trello Board
    webhookChatMessage(reactionComponents["otherInfo"].ElnId, reactionComponents["otherInfo"].ProjectName, reactionComponents["otherInfo"].Customer, reactionComponents["otherInfo"].StepName, 0, "", presentationLink);

    // write data to Batch DB Sheet
    if (newBatches.length > 0) { batchDbSheet.getRange(lastRowBatch + 1, 1, newBatches.length, newBatches[0].length).setValues(newBatches); }
    // write data to Component Roles Sheet
    if (newRoles.length > 0) { componentRolesSheet.getRange(lastRowRoles + 1, 1, newRoles.length, newRoles[0].length).setValues(newRoles); }

    if (newBatchLabels.length > 1) {// only create a file if there's more data present than the header
      var newBatchLabelsstring = newBatchLabels.join("\r\n");
      var newBatchLabelsfile = batchLabelsFolder.createFile("P-touch SMs " + String(Content[19][1]).substring(String(Content[19][1]).length - 4) + ".csv", newBatchLabelsstring); // writes the Batch Label csv for P-Touch
    }

    if (quantosHeads.length > 1) {  // only create a file if there's more data present than the header 
      var quantosHeadsXML = createQuantosXml(quantosHeads);  // converts the quantosHeads array into a Quantos digestible XML-string  
      var quantosHeadsFile = quantosFolder.createFile('Quantos head data ' + Content[19][1] + ".csl", quantosHeadsXML);     // contains all solids for writing dosing heads
    }
  } else { //side products are registered and are now added to the presentation
    addNewSideProductToPresentation(reactionComponents);
  }

  // write data to Component DB Sheet
  if (newComponents.length > 0) { componentDbSheet.getRange(componentDblastRow + 1, 1, newComponents.length, newComponents[0].length).setValues(newComponents); }

  // write data to SMs/Prods Sheet
  if (newSmsProds.length > 0) { smProdSheet.getRange(lastRowSMsProds + 1, 1, newSmsProds.length, newSmsProds[0].length).setValues(newSmsProds); }

  var response = Browser.msgBox("Do you want to reset the Sheet?", Browser.Buttons.YES_NO);
  if (response == "yes") {
    newSubmitRequestSheetReset();
  }

}

/**
 * Submit Request: backup the sheet content using the array to text method.
 *  - Not used anymore since, reset now works using copy / paste.
 */
function backupSheetContent() {    // used to backup the content of the sheet in case formulas change 

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Submit Request");
  var Content = sheet.getRange(1, 1, 25, 21).getValues();
  var Formulas = sheet.getRange(1, 1, 25, 21).getFormulas();
  var backupContent = [];
  var backupRange = sheet.getRange("W1:W25");

  for (var row = 0; row < Content.length; row++) {

    for (var column = 0; column < Content[0].length; column++) {
      if (Formulas[row][column] != '') { Content[row][column] = Formulas[row][column]; }
    }

    backupContent.push([arrayToText([Content[row]])]);
  }
  backupRange.setValues(backupContent);
}

/**
 * Submit Request: resets the sheet content using the text to array method.
 *  - Not used anymore since, reset now works using copy / paste.
 */
function resetSubmissionSheet() {  //resets the sheet, connected to the corresponding button, not used anymore, replaced by newSubmitRequestSheetReset

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Submit Request");
  var contentRange = sheet.getRange(1, 1, 25, 21);
  var contentBefore = contentRange.getValues();
  var contentAfter = [];
  var backupContent = sheet.getRange("W1:W25").getValues();

  for (var row = 0; row < backupContent.length; row++) {

    contentAfter.push(textToArray(backupContent[row][0])[0]);
  }
  for (row = 0; row < contentAfter.length; row++) {

    for (var column = 0; column < contentAfter[0].length; column++) {
      if (contentAfter[row][column] != contentBefore[row][column]) {
        if (String(contentAfter[row][column]).substring(0, 1) != '=') {
          sheet.getRange(row + 1, column + 1).setValue(contentAfter[row][column]);
        } else { sheet.getRange(row + 1, column + 1).setFormula(contentAfter[row][column]); }

      }
    }

    //contentAfter.push(textToArray(backupContent[row][0])[0]);
  }



  //contentRange.setValues(contentAfter);
}


/**
 * Submit Request: Sends messages to a Google chat room when a new experiment or plate is registered, called from savePlate and newExperimentSubmit
 *  
 */
function webhookChatMessage(ElnId = "ELNxxxxx-999", ProjectName = "Test Project", Customer = "Mr Cork Ring", StepName = "Test Step", plate = 2, cheatsheetLink = "", presentationLink = "") {
  var message = "";
  var webhookURL = globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["CHATMESSAGEwEBHOOK"];
  switch (plate) {
    case -1:
      message = {
        'text': "Problem while saving " + ElnId + ", this is the error message: " + StepName,
      };
      break;
    case 0:
      message = {
        'text': ElnId + " " + StepName + " for " + Customer + " in " + ProjectName + " was just registered. The force is strong in you, Team RoSL! 🗿🎉",
      };
      break;
    default:
      message = {
        'text': "Plate -" + ElnId + "_" + plate + " for " + ProjectName + " was just designed. One plate closer, Team RoSL! 🤖🥳",
      };
      break;
  }

  var options = {
    'method': 'post',
    'contentType': 'application/json',
    // Convert the JavaScript object to a JSON string.
    'payload': JSON.stringify(message)
  };
  UrlFetchApp.fetch(webhookURL, options);
}


/**
* Submit Request: Column to Letter, converts the column number into the corresponding character, e.g. 1 --> A, 4 --> D
* from StackOverflow: http://stackoverflow.com/questions/21229180/convert-column-index-into-corresponding-column-letter
*/
function columnToLetter(column = 26) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Submit Request: resets the Submit Request by copying a range from the SheetBackups Sheet, connected to the "Reset Sheet" Button
 * 
 */
function newSubmitRequestSheetReset() {   //executed when pressing the "Reset Sheet" Button
  // based on https://stackoverflow.com/questions/73331715/copy-a-range-to-another-spreadsheet-with-data-validations-formats-etc

  //Replicate rangeA on rangeB

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let source = ss.getSheetByName("SheetBackups");
  let destination = ss.getSheetByName("Submit Request");

  let rangeA = source.getRange("A1:V53");
  let rangeB = destination.getRange("A1");

  rangeA.copyTo(rangeB);


}
