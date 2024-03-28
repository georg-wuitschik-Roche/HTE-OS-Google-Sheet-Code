/*jshint sub:true*/
/**
 * Registration: This function is used to register new compounds in the registration sheet. It is triggered by the "Register Stuff" Button
 * 
 */
function registration() {
  //connect to the different sheets 
  var registrationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registration");
  var batchDbSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Batch DB");
  var componentDbSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Component DB");
  var solutionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Solutions");
  var componentRolesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Component Roles");


  // get the number of the last row that contains data
  var batchDbLastRow = batchDbSheet.getLastRow();
  var componentDbLastRow = componentDbSheet.getLastRow();
  var solutionsLastRow = solutionsSheet.getLastRow();
  var componentRolesLastRow = componentRolesSheet.getLastRow();

  //get the list of known batchIDs, componentIDs, componentNames, solutionIDs and molecular formulas of starting materials, (side) product(s)
  var batchIDs = batchDbSheet.getRange(2, 7, batchDbLastRow - 1).getValues();
  var componentIDs = componentDbSheet.getRange(2, 2, componentDbLastRow - 1).getValues();
  var componentNames = componentDbSheet.getRange(2, 3, componentDbLastRow - 1).getValues();
  //var solutionIDs = solutionsSheet.getRange(2, 1, solutionsLastRow - 1).getValues()
  var INCHIkeys = componentDbSheet.getRange(2, 1, componentDbLastRow - 1).getValues();

  //removes the inner brackets from these five arrays, so that the indexOf function can be used
  componentNames = [].concat.apply([], componentNames);
  componentIDs = [].concat.apply([], componentIDs);
  batchIDs = [].concat.apply([], batchIDs);
  //solutionIDs = [].concat.apply([], solutionIDs)
  INCHIkeys = [].concat.apply([], INCHIkeys);

  // create a counter for the batchID last row that starts with the initial last row
  var batchDbLastRowCounter = batchDbLastRow;
  var componentDbLastRowCounter = componentIDs[componentIDs.length - 1]; // this is the last known component ID, ensures that the right component ID is selected even if rows are deleted in the Component IDs sheet. 

  //create empty arrays for the new components, batches and solutions:
  var newComponents = [];
  var newBatches = [];
  var newSolutions = [];
  var newComponentRoles = [];

  var newBatchLabels = [["Component Role", "Component ID", "Component Name", "CAS", "Roche-No.", "Molecular Formula", "Batch ID", "Producer", "Assay", "MW", "Dosing Head ID or density", "Date"]];
  var newSolutionLabels = [["Solute ID", "Solute Name", "Solute Batch ID", "Solvent Name", "Solvent Batch ID", "Concentration", "Unit", "Solution Density", "Date"]];

  var quantosHeads = [["", "Analysis Method", "Dosing Head Tray", "Dosing Head Pos.", "Substance", "Lot ID", "Filled Quantity [mg]", "Expiration Date", "Retest Date", "Dose Limit", "Tap Before Dosing?", "Intensity [%]", "Duration [s]"]]; //later filled with dosing head information to write to Quantos heads


  // Read the main component table
  // 0 Component Role 1	Cmpt ID 2	Component Name 3 !!!RCDB ID!!! 4 CAS	5 Roche No	6 Molecular Formula	7 Batch ID	8 Producer	9 Assay
  // 10 MW	11 Density	12 Inchi-Key	13 Smiles	14 Inchi	15 Quantos Head ID 16 Solution	17 Conc.	18 Unit	19 Solvent	20 Solvent Batch ID	21 Density	22 Bottle/Head Type and Size	23 Amount (mL/mg)	24 Location Reserve Bottle
  var data = registrationSheet.getRange("B2:Y21").getValues();

  for (var row = 0; row < data.length; row++) {  //go through the data

    if (data[row][2] == "") { break; } // exits the loop, if no component name is found
    if (data[row][6] == "Find Formula on Google") { data[row][6] = "N/A"; }
    data[row][13] = data[row][13].replace(/\|.+\|/, ""); // Some Smiles strings contain a racemic or stereoflag (e.g.  |r,&1:2,25,41| or  |r,&1:1|) at the end which is upsetting Spotfire. This regex removes these 


    if (data[row][1] == "New" && componentNames.indexOf(data[row][2]) == -1 && INCHIkeys.indexOf(data[row][12]) == -1) {

      //Register the compound, add 5 empty columns and the RCDB ID
      newComponents.push([data[row][12], componentDbLastRowCounter + 1, data[row][2], data[row][10], data[row][4], data[row][5], data[row][13], data[row][14], data[row][11], "", data[row][6], "", "", "", "", "", data[row][3]]);
      //Write entry for ComponentRoles
      newComponentRoles.push(generateComponentRoles(componentDbLastRowCounter + 1, data[row][2], data[row][0])); //determination of component ID as a function of the row is not good and needs to be adjusted to what is done in HTE Submit
      componentNames.push(data[row][2]); //add the component name to the list of known component names
      INCHIkeys.push(data[row][12]);  //add the InchiKey to the list of known InchiKeys
      componentIDs.push(componentDbLastRowCounter + 1);

      // Put the new component ID into the data array, so that the labels contain the ID and not "New"
      if (String(data[row][7]).length > 14) {
        data[row][15] = (componentDbLastRowCounter + 1) + "@" + String(data[row][7]).substring(String(data[row][7]).length - 14);
      } else {
        data[row][15] = (componentDbLastRowCounter + 1) + "@" + String(data[row][7]);
      }
      data[row][1] = componentDbLastRowCounter + 1;

      componentDbLastRowCounter++;

      // Register compound batches, if the compound is new
      if (batchIDs.indexOf((componentDbLastRowCounter) + "_" + data[row][7]) == -1 && data[row][7] != "Not registered") {   //This check will almost always be true, since the compound is new and thus it's super unlikely that for the new compound ID a batch already exists.

        //Register the batch: Rowcounter (used to determine the order of registration in FileGenerator), Component ID, Batch ID, Producer, Assay, Component Name, Batch Key
        newBatches.push([batchDbLastRowCounter + 1, componentDbLastRowCounter, data[row][7], data[row][8], data[row][9], data[row][2], (componentDbLastRowCounter) + "_" + data[row][7], "", data[row][23]]);
        batchIDs.push((componentDbLastRowCounter) + "_" + data[row][7]);
        batchDbLastRowCounter++;
      }

      if (data[row][16] == true) { // Create a label for the solution, if the corresponding checkbox is checked
        newSolutionLabels.push([componentDbLastRowCounter, data[row][2], data[row][7], data[row][19], data[row][20], data[row][17], data[row][18], data[row][21], Utilities.formatDate(new Date(), 'GMT+1', "MMM d ''yy")]);
        // create an entry for the solutions table
        if (data[row][19] != "" &&
          data[row][20] != "" &&
          isNaN(data[row][17]) == false &&
          data[row][18] != "") {
          //Register the solution: Solution ID, Component ID Solute, Batch ID Solute, Component ID Solvent, Batch ID Solvent, Concentration, Unit, Density
          newSolutions.push(
            [(componentDbLastRowCounter) + "_" + data[row][7] + "_" + data[row][19] + "_" + data[row][20] + "_" + data[row][17] + "_" + data[row][18],
              componentDbLastRowCounter,
            data[row][7],
            componentIDs[componentNames.indexOf(data[row][19])],
            data[row][20],
            data[row][17],
            data[row][18],
            data[row][21],
            data[row][22],
            data[row][23]]);
          //solutionIDs.push((componentDbLastRowCounter - 1) + "_" + data[row][7] + "_" + data[row][19] + "_" + data[row][20] + "_" + data[row][17] + "_" + data[row][18]);
        }
      }

      // For the cases below the compound is already registered. data[row][2] can be used as name, as it is already matched to the name in the ComponentDB
    } else if (batchIDs.indexOf(data[row][1] + "_" + data[row][7]) == -1 && data[row][7] != "Not registered") { // this is the case, where the compound is known and the if-clause checks whether the batch specified exists already and if not registers it.
      newBatches.push([batchDbLastRowCounter + 1, data[row][1], data[row][7], data[row][8], data[row][9], data[row][2], data[row][1] + "_" + data[row][7], "", data[row][23]]);
      batchIDs.push(data[row][1] + "_" + data[row][7]);
      batchDbLastRowCounter++;

      if (data[row][16] == true) { // Create a label for the solution, if the corresponding checkbox is checked
        newSolutionLabels.push([data[row][1], data[row][2], data[row][7], data[row][19], data[row][20], data[row][17], data[row][18], data[row][21], Utilities.formatDate(new Date(), 'GMT+1', "MMM d ''yy")]);
        // does the same thing for solutions
        if (data[row][19] != "" &&
          data[row][20] != "" &&
          isNaN(data[row][17]) == false &&
          data[row][18] != "") {
          //Register the solution: Solution ID, Component ID Solute, Batch ID Solute, Component ID Solvent, Batch ID Solvent, Concentration, Unit, Density
          newSolutions.push(
            [data[row][1] + "_" + data[row][7] + "_" + data[row][19] + "_" + data[row][20] + "_" + data[row][17] + "_" + data[row][18],
            data[row][1],
            data[row][7],
            componentIDs[componentNames.indexOf(data[row][19])],
            data[row][20],
            data[row][17],
            data[row][18],
            data[row][21],
            data[row][22],
            data[row][23]]);
          //solutionIDs.push((componentDbLastRowCounter - 1) + "_" + data[row][7] + "_" + data[row][19] + "_" + data[row][20] + "_" + data[row][17] + "_" + data[row][18]);
        }
      }
    } else {
      Browser.msgBox("Batch " + data[row][7] + " of " + data[row][16] + " exists already in the database. A new solution for this batch will be regardless now, if you activated the corresponding checkbox in column Q.");
      if (data[row][16] == true) { // Create a label for the solution, if the corresponding checkbox is checked
        newSolutionLabels.push([data[row][1], data[row][2], data[row][7], data[row][19], data[row][20], data[row][17], data[row][18], data[row][21], Utilities.formatDate(new Date(), 'GMT+1', "MMM d ''yy")]);
        // does the same thing for solutions
        if (data[row][19] != "" &&
          data[row][20] != "" &&
          isNaN(data[row][17]) == false &&
          data[row][18] != "") {
          //Register the solution: Solution ID, Component ID Solute, Batch ID Solute, Component ID Solvent, Batch ID Solvent, Concentration, Unit, Density
          newSolutions.push(
            [data[row][1] + "_" + data[row][7] + "_" + data[row][19] + "_" + data[row][20] + "_" + data[row][17] + "_" + data[row][18],
            data[row][1],
            data[row][7],
            componentIDs[componentNames.indexOf(data[row][19])],
            data[row][20],
            data[row][17],
            data[row][18],
            data[row][21],
            data[row][22],
            data[row][23]]);
          //solutionIDs.push((componentDbLastRowCounter - 1) + "_" + data[row][7] + "_" + data[row][19] + "_" + data[row][20] + "_" + data[row][17] + "_" + data[row][18]);
        }
      }
    }


    labelRole = data[row][0];
    labelID = data[row][1];
    labelName = data[row][2];

    // Collect info for label printer and Quantos
    if (isNaN(parseFloat(data[row][11])) == true || data[row][11] == "") {   // true, if no density is given or the content of the cell is not a floating point number = component is a solid

      newBatchLabels.push([(labelRole).toString().split(',').join(':'), labelID, (labelName).toString().split(',').join(';'), data[row][4], data[row][5], data[row][6], data[row][7], data[row][8], data[row][9] * 100 + "%", data[row][10], data[row][15], Utilities.formatDate(new Date(), 'GMT+1', "MMM d ''yy")]);
      quantosHeads.push([quantosHeads.length, "C:\\Users\\Public\\Documents\\Chronos\\Methods\\HeadWrite in Sequence.cam", "Heads", quantosHeads.length, data[row][15], String(labelName).substring(0, 15), 6666, "", "", 999, "True", 40, 1]);

    } else { // if a density is provided, then it must be a liquid and doesn't have a dosing Head ID, the combination of split and join is used to replace all occurrencies of commas, as these are used as separators in the label csv.
      newBatchLabels.push([(labelRole).toString().split(',').join(':'), labelID, (labelName).toString().split(',').join(';'), data[row][4], data[row][5], data[row][6], data[row][7], data[row][8], data[row][9] * 100 + "%", data[row][10], data[row][11] + " g/mL", Utilities.formatDate(new Date(), 'GMT+1', "MMM d ''yy")]);
    }

  }

  // If newComponents have a picture in column M (index 12), put an empty string
  for (var comp = 0; comp < newComponents.length; comp++) {
    if (newComponents[comp][12] instanceof Object) newComponents[comp][12] = '';
  }

  //Write the arrays containing component, batch, solution and role data to the different sheets, if there's data in them
  if (newComponents.length > 0) { componentDbSheet.getRange(componentDbLastRow + 1, 1, newComponents.length, newComponents[0].length).setValues(newComponents); }                          // write the new components
  if (newBatches.length > 0) { batchDbSheet.getRange(batchDbLastRow + 1, 1, newBatches.length, newBatches[0].length).setValues(newBatches); }                                           //  write the new batches
  if (newSolutions.length > 0) { solutionsSheet.getRange(solutionsLastRow + 1, 1, newSolutions.length, newSolutions[0].length).setValues(newSolutions); }                               //   write the new solutions
  if (newComponentRoles.length > 0) { componentRolesSheet.getRange(componentRolesLastRow + 1, 1, newComponentRoles.length, newComponentRoles[0].length).setValues(newComponentRoles); } //    write the new component roles

  //SpreadsheetApp.flush();
  var folder = DriveApp.getFolderById(globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["PtOUCHlABELSfOLDERiD"]);   // This is the folder P-Touch in Robot Input Files
  var quantosFolder = DriveApp.getFolderById(globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["QUANTOSfOLDERiD"]);

  if (newBatchLabels.length > 1) {
    var batchLabelFileName = 'Compound IDs';
    var newBatchLabelsstring = newBatchLabels.join("\r\n");
    for (var item = 1; item < newBatchLabels.length; item++) {
      batchLabelFileName = batchLabelFileName + "_" + newBatchLabels[item][1];
    }
    var newBatchLabelsfile = folder.createFile(batchLabelFileName + ".csv", newBatchLabelsstring); // writes the Batch Label csv for P-Touch
  }

  if (newSolutionLabels.length > 1) {
    var newSolutionLabelsstring = newSolutionLabels.join("\r\n");
    var newSolutionLabelsfile = folder.createFile("Solution labels " + Utilities.formatDate(new Date(), 'GMT+1', "yyyy-MM-dd") + ".csv", newSolutionLabelsstring); // writes the Solution Label csv for P-Touch
  }

  if (quantosHeads.length > 1) {  // only write the xml, if there are solids in the list of compounds to be registered. 
    var quantosHeadsXML = createQuantosXml(quantosHeads);
    var quantosHeadsFile = quantosFolder.createFile('Quantos solids ' + Utilities.formatDate(new Date(), 'GMT+1', "yyyy-MM-dd") + ".csl", quantosHeadsXML);     // contains all solids for writing dosing heads
  }
  var response = Browser.msgBox("Do you want to reset the Sheet?", Browser.Buttons.YES_NO);
  if (response == "yes") {
    newRegistrationSheetReset();
  }
}

/**
 * Registration: generates the p-Touch files for Dosing heads and Solutions
 * 
 */
function printOldBatchLabels() {
  //connect to the different sheets 
  var registrationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registration");
  var batchDbSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Batch DB");
  var componentDbSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Component DB");
  var solutionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Solutions");
  var componentRolesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Component Roles");


  // get the number of the last row that contains data
  var batchDbLastRow = batchDbSheet.getLastRow();
  var componentDbLastRow = componentDbSheet.getLastRow();
  // var solutionsLastRow = solutionsSheet.getLastRow()
  // var componentRolesLastRow = componentRolesSheet.getLastRow()

  //get the list of known batchIDs, componentIDs, componentNames, solutionIDs and molecular formulas of starting materials, (side) product(s)
  var batchIDs = batchDbSheet.getRange(2, 7, batchDbLastRow - 1).getValues();
  var componentIDs = componentDbSheet.getRange(2, 2, componentDbLastRow - 1).getValues();
  var componentNames = componentDbSheet.getRange(2, 3, componentDbLastRow - 1).getValues();
  //var solutionIDs = solutionsSheet.getRange(2, 1, solutionsLastRow - 1).getValues()
  var INCHIkeys = componentDbSheet.getRange(2, 1, componentDbLastRow - 1).getValues();

  //removes the inner brackets from these five arrays, so that the indexOf function can be used
  componentNames = [].concat.apply([], componentNames);
  componentIDs = [].concat.apply([], componentIDs);
  batchIDs = [].concat.apply([], batchIDs);
  //solutionIDs = [].concat.apply([], solutionIDs)
  INCHIkeys = [].concat.apply([], INCHIkeys);

  // create a counter for the componentID and batchID last row that starts with the initial last row

  var componentDbLastRowCounter = componentIDs[componentIDs.length - 1]; // this is the last known component ID, ensures that the right component ID is selected even if rows are deleted in the Component IDs sheet. 
  var batchDbLastRowCounter = batchDbLastRow;

  //create empty arrays for the new components, batches and solutions:
  var newComponents = [];
  //var newBatches = [];
  //var newSolutions = [];
  //var newComponentRoles = [];

  var newBatchLabels = [["Component Role", "Component ID", "Component Name", "CAS", "Roche-No.", "Molecular Formula", "Batch ID", "Producer", "Assay", "MW", "Dosing Head ID or density", "Date"]];
  var newSolutionLabels = [["Solute ID", "Solute Name", "Solute Batch ID", "Solvent Name", "Solvent Batch ID", "Concentration", "Unit", "Solution Density", "Date"]];

  var quantosHeads = [["", "Analysis Method", "Dosing Head Tray", "Dosing Head Pos.", "Substance", "Lot ID", "Filled Quantity [mg]", "Expiration Date", "Retest Date", "Dose Limit", "Tap Before Dosing?", "Intensity [%]", "Duration [s]"]]; //later filled with dosing head information to write to Quantos heads


  //read the main component table
  var data = registrationSheet.getRange("B2:W21").getValues();

  for (var row = 0; row < data.length; row++) {  //go through the data

    if (data[row][2] == "") { break; } // exits the loop, if no component name is found
    if (data[row][6] == "Find Formula on Google") { data[row][6] = "N/A"; }



    if (data[row][1] == "New" && componentNames.indexOf(data[row][2]) == -1 && INCHIkeys.indexOf(data[row][12]) == -1) {
      //Register the compound
      newComponents.push([data[row][12], componentDbLastRowCounter + 1, data[row][2], data[row][10], data[row][4], data[row][5], data[row][13], data[row][14], data[row][11], "", data[row][6], "", "", "", "", "", data[row][3]]);
      //Write entry for ComponentRoles
      //newComponentRoles.push(generateComponentRoles(componentDbLastRowCounter + 1, data[row][2], data[row][0])) //determination of component ID as a function of the row is not good and needs to be adjusted to what is done in HTE Submit
      /* componentNames.push(data[row][2]) //add the component name to the list of known component names
      INCHIkeys.push(data[row][12])  //add the InchiKey to the list of known InchiKeys
      componentIDs.push(componentDbLastRowCounter + 1)

      // Put the new component ID into the data array, so that the labels contain the ID and not "New"
      if (String(data[row][7]).length > 14) {
        data[row][15] = (componentDbLastRowCounter + 1) + "@" + String(data[row][7]).substring(String(data[row][7]).length - 14)
      } else {
        data[row][15] = (componentDbLastRowCounter + 1) + "@" + String(data[row][7])
      }
      data[row][1] = componentDbLastRowCounter + 1

      componentDbLastRowCounter++; */


      /* Register compound batches, if the compound is new
      if (batchIDs.indexOf((componentDbLastRowCounter) + "_" + data[row][7]) == -1 && data[row][7] != "Not registered") {   //This check will almost always be true, since the compound is new and thus it's super unlikely that for the new compound ID a batch already exists.

        //Register the batch: Rowcounter (used to determine the order of registration in FileGenerator), Component ID, Batch ID, Producer, Assay, Component Name, Batch Key
        newBatches.push([batchDbLastRowCounter + 1, componentDbLastRowCounter, data[row][7], data[row][8], data[row][9], data[row][2], (componentDbLastRowCounter) + "_" + data[row][7]]);
        batchIDs.push((componentDbLastRowCounter) + "_" + data[row][7]);
        batchDbLastRowCounter++; */
    }

    if (data[row][16] == true) { // Create a label for the solution, if the corresponding checkbox is checked
      newSolutionLabels.push([componentDbLastRowCounter, data[row][2], data[row][7], data[row][19], data[row][20], data[row][17], data[row][18], data[row][21], Utilities.formatDate(new Date(), 'GMT+1', "MMM d ''yy")]);
      // create an entry for the solutions table
      /*if (data[row][19] != "" &&
        data[row][20] != "" &&
        isNaN(data[row][17]) == false &&
        data[row][18] != "") {
        //Register the solution: Solution ID, Component ID Solute, Batch ID Solute, Component ID Solvent, Batch ID Solvent, Concentration, Unit, Density
        newSolutions.push(
          [(componentDbLastRowCounter) + "_" + data[row][7] + "_" + data[row][19] + "_" + data[row][20] + "_" + data[row][17] + "_" + data[row][18],
          componentDbLastRowCounter,
          data[row][7],
          componentIDs[componentNames.indexOf(data[row][19])],
          data[row][20],
          data[row][17],
          data[row][18],
          data[row][21]]);
        //solutionIDs.push((componentDbLastRowCounter - 1) + "_" + data[row][7] + "_" + data[row][19] + "_" + data[row][20] + "_" + data[row][17] + "_" + data[row][18]);
      } */
    }

    /* } else if (batchIDs.indexOf(data[row][1] + "_" + data[row][7]) == -1 && data[row][7] != "Not registered") { // this is the case, where the compound is known and the if-clause checks whether the batch specified exists already and if not registers it.
       newBatches.push([batchDbLastRowCounter + 1, data[row][1], data[row][7], data[row][8], data[row][9], data[row][2], data[row][1] + "_" + data[row][7]]);
       batchIDs.push(data[row][1] + "_" + data[row][7]);
       batchDbLastRowCounter++;
 
       if (data[row][16] == true) { // Create a label for the solution, if the corresponding checkbox is checked
         newSolutionLabels.push([data[row][1], data[row][2], data[row][7], data[row][19], data[row][20], data[row][17], data[row][18], data[row][21], Utilities.formatDate(new Date(), 'GMT+1', "MMM d ''yy")]);
         // does the same thing for solutions
          if (data[row][19] != "" &&
           data[row][20] != "" &&
           isNaN(data[row][17]) == false &&
           data[row][18] != "") {
           //Register the solution: Solution ID, Component ID Solute, Batch ID Solute, Component ID Solvent, Batch ID Solvent, Concentration, Unit, Density
           newSolutions.push(
             [data[row][1] + "_" + data[row][7] + "_" + data[row][19] + "_" + data[row][20] + "_" + data[row][17] + "_" + data[row][18],
             data[row][1],
             data[row][7],
             componentIDs[componentNames.indexOf(data[row][19])],
             data[row][20],
             data[row][17],
             data[row][18],
             data[row][21]]);
           //solutionIDs.push((componentDbLastRowCounter - 1) + "_" + data[row][7] + "_" + data[row][19] + "_" + data[row][20] + "_" + data[row][17] + "_" + data[row][18]);
         }
       
       } 
     } else {
       Browser.msgBox("Batch " + data[row][7] + " of " + data[row][16] + " exists already in the database. A new solution for this batch will be regardless now, if you activated the corresponding checkbox in column Q.")
       if (data[row][16] == true) { // Create a label for the solution, if the corresponding checkbox is checked
         newSolutionLabels.push([data[row][1], data[row][2], data[row][7], data[row][19], data[row][20], data[row][17], data[row][18], data[row][21], Utilities.formatDate(new Date(), 'GMT+1', "MMM d ''yy")]);
         // does the same thing for solutions
         if (data[row][19] != "" &&
           data[row][20] != "" &&
           isNaN(data[row][17]) == false &&
           data[row][18] != "") {
           //Register the solution: Solution ID, Component ID Solute, Batch ID Solute, Component ID Solvent, Batch ID Solvent, Concentration, Unit, Density
           newSolutions.push(
             [data[row][1] + "_" + data[row][7] + "_" + data[row][19] + "_" + data[row][20] + "_" + data[row][17] + "_" + data[row][18],
             data[row][1],
             data[row][7],
             componentIDs[componentNames.indexOf(data[row][19])],
             data[row][20],
             data[row][17],
             data[row][18],
             data[row][21]]);
           //solutionIDs.push((componentDbLastRowCounter - 1) + "_" + data[row][7] + "_" + data[row][19] + "_" + data[row][20] + "_" + data[row][17] + "_" + data[row][18]);
         }
       }
     }*/


    // Collect info for label printer and Quantos
    if (isNaN(parseFloat(data[row][11])) == true || data[row][11] == "") {   // true, if no density is given or the content of the cell is not a floating point number = component is a solid

      newBatchLabels.push([(data[row][0]).toString().split(',').join(':'), data[row][1], (data[row][2]).toString().split(',').join(';'), data[row][4], data[row][5], data[row][6], data[row][7], data[row][8], data[row][9] * 100 + "%", data[row][10], data[row][15], Utilities.formatDate(new Date(), 'GMT+1', "MMM d ''yy")]); //the combination of split and join is used to replace all occurrencies of commas, as these are used as separators in the label csv.
      quantosHeads.push([quantosHeads.length, "C:\\Users\\Public\\Documents\\Chronos\\Methods\\HeadWrite in Sequence.cam", "Heads", quantosHeads.length, data[row][15], String(data[row][2]).substring(0, 15), 6666, "", "", 999, "True", 40, 1]);

    } else { // if a density is provided, then it must be a liquid and doesn't have a dosing Head ID, the combination of split and join is used to replace all occurrencies of commas, as these are used as separators in the label csv.
      newBatchLabels.push([(data[row][0]).toString().split(',').join(':'), data[row][1], (data[row][2]).toString().split(',').join(';'), data[row][4], data[row][5], data[row][6], data[row][7], data[row][8], data[row][9] * 100 + "%", data[row][10], data[row][11] + " g/mL", Utilities.formatDate(new Date(), 'GMT+1', "MMM d ''yy")]);
    }


  }

  //Write the arrays containing component, batch, solution and role data to the different sheets, if there's data in them
  /* if (newComponents.length > 0) { componentDbSheet.getRange(componentDbLastRow + 1, 1, newComponents.length, newComponents[0].length).setValues(newComponents) };                          // write the new components
   if (newBatches.length > 0) { batchDbSheet.getRange(batchDbLastRow + 1, 1, newBatches.length, newBatches[0].length).setValues(newBatches) };                                             //  write the new batches
   if (newSolutions.length > 0) { solutionsSheet.getRange(solutionsLastRow + 1, 1, newSolutions.length, newSolutions[0].length).setValues(newSolutions) };                                //   write the new solutions
   if (newComponentRoles.length > 0) { componentRolesSheet.getRange(componentRolesLastRow + 1, 1, newComponentRoles.length, newComponentRoles[0].length).setValues(newComponentRoles) }; //    write the new component roles
 
   //SpreadsheetApp.flush(); */
  var folder = DriveApp.getFolderById(globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["PtOUCHlABELSfOLDERiD"]);   // This is the folder P-Touch in Robot Input Files
  var quantosFolder = DriveApp.getFolderById(globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["QUANTOSfOLDERiD"]);

  if (newBatchLabels.length > 1) {
    var batchLabelFileName = 'Re-Label';
    var newBatchLabelsstring = newBatchLabels.join("\r\n");
    for (var item = 1; item < newBatchLabels.length; item++) {
      batchLabelFileName = batchLabelFileName + "_" + newBatchLabels[item][1];
    }
    var newBatchLabelsfile = folder.createFile(batchLabelFileName + ".csv", newBatchLabelsstring); // writes the Batch Label csv for P-Touch
  }

  if (newSolutionLabels.length > 1) {
    var newSolutionLabelsstring = newSolutionLabels.join("\r\n");
    var newSolutionLabelsfile = folder.createFile("Re-Solution labels " + Utilities.formatDate(new Date(), 'GMT+1', "yyyy-MM-dd") + ".csv", newSolutionLabelsstring); // writes the Solution Label csv for P-Touch
  }

  if (quantosHeads.length > 1) {  // only write the xml, if there are solids in the list of compounds to be registered. 
    var quantosHeadsXML = createQuantosXml(quantosHeads);
    var quantosHeadsFile = quantosFolder.createFile('Quantos solids ' + Utilities.formatDate(new Date(), 'GMT+1', "yyyy-MM-dd") + ".csl", quantosHeadsXML);     // contains all solids for writing dosing heads
  }
  var response = Browser.msgBox("Do you want to reset the Sheet?", Browser.Buttons.YES_NO);
  if (response == "yes") {
    newRegistrationSheetReset();
  }
}
/**
 * Registration: register new reference spectra
 * 
 */
function registerAnalytics() {

  // Connect to the sheets involved

  var registrationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registration");
  var analyticsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cmpnd_Ref_Analytics");


  // get the number of the last row that contains data
  var analyticsSheetLastRow = analyticsSheet.getLastRow();

  // Read the data

  var data = registrationSheet.getRange("B24:L46").getValues();


  // Write the data to the Cmpnd_Ref_Analytics sheet

  analyticsSheet.getRange(analyticsSheetLastRow + 1, 1, data.length, 11).setValues(data);
  var response = Browser.msgBox("Do you want to reset the Sheet?", Browser.Buttons.YES_NO);
  if (response == "yes") {
    newRegistrationSheetReset();
  }
}


/**
 * Registration: generates the p-Touch files for Dosing heads and Solutions
 * @param {Number} ComponentID Component ID.
 * @param {String} ComponentName Component Name.
 * @param {String} ComponentRole Role of the compound to be registered.
 * @return {Array} one line of the array to be written to the Component Roles sheet.
 * 
 */
function generateComponentRoles(ComponentID, ComponentName, ComponentRole) {
  switch (String(ComponentRole).substring(0, 5)) { //The component roles sheet contains a super category for catalysts, acids, bases and solvents that allows the user to choose either from the sub- or super-category
    case "Catal":
      return [ComponentID, ComponentName, ComponentRole, "Catalysts, all"];

    case "Acids":
      return [ComponentID, ComponentName, ComponentRole, "Acids, all"];

    case "Bases":
      return [ComponentID, ComponentName, ComponentRole, "Bases, all"];

    case "Solve":
      return [ComponentID, ComponentName, ComponentRole, "Solvents, all"];

    case "Photo":
      return [ComponentID, ComponentName, ComponentRole, "Photocatalysts, all"];

    default:
      return [ComponentID, ComponentName, ComponentRole, ""];
  }
}

/**
 * Registration: Takes the whole sheet and compresses all content and formulas into text that is stored in Z1:Z45 with one cell containing all the data for the whole row.
 * Not needed anymore, since backup/restore is now being handled using copy/paste & NOT updated for format with RCDB
 */
function backupRegistrationSheetContent() { // Takes the whole sheet and compresses all content and formulas into text that is stored in Z1:Z45 with one cell containing all the data for the whole row.

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registration");
  var Content = sheet.getRange(1, 1, 46, 24).getValues();
  var Formulas = sheet.getRange(1, 1, 46, 24).getFormulas();
  var backupContent = [];
  var backupRange = sheet.getRange("Z1:Z46");

  for (var row = 0; row < Content.length; row++) { // go through all rows

    for (var column = 0; column < Content[0].length; column++) { // go through all cells of each row
      if (Formulas[row][column] != '') { Content[row][column] = Formulas[row][column]; }      // If there's a formula present in this row, save the formula and not the value of the cell
    }

    backupContent.push([arrayToText([Content[row]])]);  //once one row is finished, turn the content into a string and append it to an array.
  }
  backupRange.setValues(backupContent);  // write the content of the array
}

/**
 * Registration: Take the data from Z1:Z45 and write it back to the sheet.
 * not used anymore, superseded by newRegistrationSheetReset, NOT updated for format with RCDB
 
function resetRegistrationSheet() {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registration");
  var contentRange = sheet.getRange(1, 1, 45, 24);
  var contentBefore = contentRange.getValues();
  var contentAfter = [];
  var backupContent = sheet.getRange("Z1:Z45").getValues();

  for (var row = 0; row < backupContent.length; row++) {    // go through each row of the content in column W

    contentAfter.push(textToArray(backupContent[row][0])[0]);    // Take the content from the corresponding cell in column W and upack it
  }
  for (row = 0; row < contentAfter.length; row++) { // go through all the rows of the unpacked array

    for (var column = 0; column < contentAfter[0].length; column++) {  // go through each cell of this row of the unpacked array
      if (contentAfter[row][column] != contentBefore[row][column]) {   // if there's a difference between what is already there and what is in the backup content
        if (String(contentAfter[row][column]).substring(0, 1) != '=') {  // if the backup content doesn't contain a formula, write the value    
          sheet.getRange(row + 1, column + 1).setValue(contentAfter[row][column]);
        } else { sheet.getRange(row + 1, column + 1).setFormula(contentAfter[row][column]); } // if it contains a formula, write the formula 

      }
    }
  }
} 

*/

/**
 * Registration: This resets the lookup tables in the Registration Sheet.
 * 
 */
function writeBatchLookupFormulas() {
  var registrationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registration");


  for (var i = 0; i < 10; i++) {
    for (var col = 1; col < 3; col++) {

      registrationSheet.getRange(48 + 5 * i, (col - 1) * 2 + 1).setValue("Dropdown for Registration I" + (i + 2 + (col - 1) * 10) + ":");
      registrationSheet.getRange(48 + 5 * i, (col - 1) * 2 + 2).setFormula('=if($I' + (i + 2 + (col - 1) * 10) + '="","",iferror(query(\'Batch DB\'!$A$2:$D, "Select C where B = "&$C' + (i + 2 + (col - 1) * 10) + '&" order by A desc limit 5"),"Not registered"))');

      // set background colors light yellow for the cell containing the formula and light green for the cells that may be filled by the formula
      registrationSheet.getRange(48 + 5 * i, 2 * col).setBackground('#fff2cc');
      registrationSheet.getRange(49 + 5 * i, 2 * col, 4, 1).setBackground('#d9ead3');


      if (col == 1) {
        registrationSheet.getRange(i + 2 + (col - 1) * 10, 9).setDataValidation(SpreadsheetApp.newDataValidation()
          .setAllowInvalid(true)
          .requireValueInRange(registrationSheet.getRange('Registration!$B$' + (48 + 5 * i) + ':$B$' + (52 + 5 * i)), false)
          .build());

        registrationSheet.getRange(i + 2 + (col - 1) * 10, 9).setFormula("=Registration!$B$" + (48 + 5 * i));
      } else {
        registrationSheet.getRange(i + 2 + (col - 1) * 10, 9).setDataValidation(SpreadsheetApp.newDataValidation()
          .setAllowInvalid(true)
          .requireValueInRange(registrationSheet.getRange('Registration!$D$' + (48 + 5 * i) + ':$D$' + (52 + 5 * i)), false)
          .build());

        registrationSheet.getRange(i + 2 + (col - 1) * 10, 9).setFormula("=Registration!$D$" + (48 + 5 * i));
      }



    }
  }
}

/**
 * Registration: resets the Registration Sheet by copying a range from the SheetBackups Sheet, connected to the "Reset Sheet" Button
 * 
 */
function newRegistrationSheetReset() {   //executed when pressing the "Reset Sheet" Button
  // based on https://stackoverflow.com/questions/73331715/copy-a-range-to-another-spreadsheet-with-data-validations-formats-etc

  //Replicate rangeA on rangeB

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let source = ss.getSheetByName("SheetBackups");
  let destination = ss.getSheetByName("Registration");

  let rangeA = source.getRange("A56:Y152");
  let rangeB = destination.getRange("A1");

  rangeA.copyTo(rangeB);


}
