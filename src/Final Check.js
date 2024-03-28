/**
 * Final Check: returns an array with the number of undosed solids not dosed so far for the plateIDs handed over. 
 * @param {Array} last10PlateIds array of PlateIds to be checked
 * @param {Array} checkBoxValue dummyParameter to force re-calculation of the formula on change
 * @return {Array} Array of PlateID, number of missing dosings, Component ID and BatchID
 * @customfunction
 */
function getSolidsOfSelectedPlates(last10PlateIds, checkBoxValue = 1) {

    // setup the variables
    var elnId = "";
    var plateNumber = 0;


    //select concat(ELN_ID,'_', PLATENUMBER) as PlateID, Coordinate, Component_ID, Batch_ID, ActualMass, DosingTimestamp from wells_prod where ActualVolume is null and ELN_ID = 'ELN032036-374' and PLATENUMBER = 1 and (

    var sqlString = "select STRING_AGG(concat(ELN_ID,'_', PLATENUMBER, '$$', Coordinate, '$$',Component_ID, '$$',Batch_ID, '$$',ActualMass, '$$',DosingTimestamp), '£') WITHIN GROUP (ORDER BY ELN_ID ASC, PLATENUMBER ASC, Component_ID ASC, Batch_ID ASC) from wells_prod where  ActualVolume is null and (";
    var selectedPlateIds = [];
    for (var row = 0; row < last10PlateIds.length; row++) { // generate the part of the sql String that specifies which plates to look for

        [elnId, plateNumber] = last10PlateIds[row][0].split("_");
        if (elnId.length > 0) {
            sqlString += "(ELN_ID = '" + elnId + "' and PLATENUMBER = " + plateNumber + ")";
        } else {
            sqlString = sqlString.substring(0, sqlString.length - 4);
            break;
        }
        if (row < last10PlateIds.length - 1) sqlString += " or ";
    }
    sqlString += " )";

    // connect to the database
    const idOfCurrentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet().getId();
    var connector = new mssql_jdbc_api(   // connect to the database
        globalVariableDict[idOfCurrentSpreadsheet]["DBsERVERiP"],
        globalVariableDict[idOfCurrentSpreadsheet]["DBpORT"],
        globalVariableDict[idOfCurrentSpreadsheet]["DBnAME"],
        globalVariableDict[idOfCurrentSpreadsheet]["DBuSERNAME"],
        globalVariableDict[idOfCurrentSpreadsheet]["DBpASSWORD"]);

    // send the queries
    connector.setBatchSize(1000);
    connector.executeQuery(sqlString);
    var sqlResultsString = connector.getResultsAsArray(); // Result is a long string in a 2D-array with one member: [ [ 'ELN032036-371_1$$A1$$1$$00389584$$17.32$$£ELN032036-371_1$$A1$$37$$MKCG3978$$1.33$$£ELN032036-371_1$$A1$$265$$BCCB0369$$0.4651$$Dosing Completed_£ELN032036-371_1$$A1$$1243$$42298$$0.0002$$Dosing Completed_; PowderflowError_Dosing Status: PowderflowError - Sample Data Error: NotAllowedAtTheMoment£ELN032036-371_1$$A1$$1439$$A0398525$$0.4151$$Dosing Completed_£ELN032036-371_1$$A1$$2148$$No ... ']]
    sqlResultsString = sqlResultsString[0][0]; // get the string out

    const sqlResultsTablefirstSplit = sqlResultsString.split("£"); // split the string into individual lines
    var sqlResultsTable = [];
    for (var row = 0; row < sqlResultsTablefirstSplit.length; row++) {
        sqlResultsTable.push(sqlResultsTablefirstSplit[row].split("$$")); // split the individual lines into columns to get the actual table.
    }

    connector.disconnect();
    return sqlResultsTable;
}

/**
 * Final Check: Toggles a checkbox in B139 used to trigger an update of getSolidsOfSelectedPlates. 
 */
function toggleCheckbox() {
    var spreadsheet = SpreadsheetApp.getActive();
    const currentState = spreadsheet.getRange('B139').getValue();
    if (currentState) spreadsheet.getRange('B139').setValue('FALSE');
    else spreadsheet.getRange('B139').setValue('TRUE');
};

/**
 * Final Check: creates the inital and correction dosing files for the plates selected in Final Check
 */
function generateQuantosFinalCheckXml() {
    // quantosXmlArray.push([quantosXmlArray.length + 2, "C:\\Users\\Public\\Documents\\Chronos\\Methods\\Dosing Method.cam", "", "", "", "", "Quantos", "", "", data[2] + "@" + batchId.toString().slice(-14), "Tray7", "", plateCoordinate, mass, parseFloat(7 / (parseFloat(mass) + 0.03) + 3).toFixed(0), ELNiD + '_' + plateNumber + " - Initial - " + vialOption, data[3]]); // will be used to generate the csl input file, tray 7 is chosen instead of a valid tray, because it forces the user to specify a tray and prevents erroneous dosing into tray 1.
    // quantosXmlArray.unshift([1, "C:\\Users\\Public\\Documents\\Chronos\\Methods\\Set Config.cam", vialOption, preDoseTapping, preDoseStrength, toleranceMode, "Quantos", useFrontDoor, useSideDoors, "", "", "", "", "", "", "", ""]);
    // quantosXmlArray.unshift(["", "Analysis Method", "Dosing Tray Type", "PreDose Tap Duration [s]", "PreDose Tap Intensity [%]", "Tolerance Mode", "Device", "Use Front Door?", "Use Side Doors?", "Substance Name", "Dosing Vial Tray", "Dosing Vial Pos.", "Dosing Vial Pos. [Axx]", "Amount [mg]", "Tolerance [%]", "Sample ID", "Comment"]);

    // Read in date from columns D to N in the Final Check sheet:
    // Include?	Plate ID	Coordinate	Component ID	Batch ID	Mass retrieved from db	Dosing Timestamp	Role	intended Mass	Mass dosed so far	Deviation from desired value
    //    0        1             2         3              4               5                    6                  7          8                9                    10
    const finalCheckSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Final Check");
    const lastFilledRow = finalCheckSheet.getRange('B138').getValue();
    if (lastFilledRow == 0) return;
    const finalCheckData = finalCheckSheet.getRange("D2:N" + lastFilledRow).getValues();
    const numberOfVialsOnPlatesArray = finalCheckSheet.getRange("A2:C11").getValues();
    var quantosXmlDictionary = {};

    var vialOptionDictionary = {
        96: "96 wells 1 mL gold",
        48: "48 wells 2 mL gold",
        24: "24 wells 1 mL gold"
    };


    var quantosXmlString = "";
    const initialsFolder = DriveApp.getFolderById(globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["QUANTOSfOLDERiD"]);
    const correctionsFolder = DriveApp.getFolderById(globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["QUANTOScORRECTIONdOSINGiD"]);
    const toleranceMode = "MinusPlus";
    const preDoseTapping = 6;      // length in seconds of the predose tapping
    const preDoseStrength = 40;           // strength in % of maximum tapping strength
    const useFrontDoor = "True";
    const useSideDoors = "False";

    //go through the numberOfVialsOnPlatesArray and setup the quantosXmlDictionary with the possible keys

    for (var row = 0; row < numberOfVialsOnPlatesArray.length; row++) {
        if (numberOfVialsOnPlatesArray[row][0] == false) continue;
        quantosXmlDictionary[numberOfVialsOnPlatesArray[row][1]] = {};
        quantosXmlDictionary[numberOfVialsOnPlatesArray[row][1]]["initial"] = [];
        quantosXmlDictionary[numberOfVialsOnPlatesArray[row][1]]["correction"] = [];
        if (parseInt(numberOfVialsOnPlatesArray[row][2]) in vialOptionDictionary) {
            quantosXmlDictionary[numberOfVialsOnPlatesArray[row][1]]["vialOption"] = vialOptionDictionary[parseInt(numberOfVialsOnPlatesArray[row][2])]; //all full plates get sorted into categories asssuming 1 mL vials, since there is no way of knowing
        } else {
            console.log(numberOfVialsOnPlatesArray[row][0]);
            const response = Browser.inputBox(numberOfVialsOnPlatesArray[row][1] + " may only be partially filled. Enter the number of vials the full plate would have ( 24, 48 or 96 ):");
            if (response in vialOptionDictionary) {
                quantosXmlDictionary[numberOfVialsOnPlatesArray[row][1]]["vialOption"] = vialOptionDictionary[response]; //all full plates get sorted into categories asssuming 1 mL vials, since there is no way of knowing
            } else {
                Browser.msgBox(response + " is not a recognized plate size. Accepted answers are 24, 28 and 96. Aborting script now.");
                return;
            }
        }
    }

    // Go through the array and distribute the content into a dictionary with the different plates as keys in a format that can be sent to createQuantosXml 
    for (row = 0; row < finalCheckData.length; row++) {
        if (finalCheckData[row][0] == false) continue;    // only include rows, where the checkbox is checked
        if (finalCheckData[row][1] in quantosXmlDictionary) {
            if (finalCheckData[row][6] == "") {

                // attach the info for this line to the array of initial dosings
                quantosXmlDictionary[finalCheckData[row][1]]["initial"].push([
                    quantosXmlDictionary[finalCheckData[row][1]]["initial"].length + 2,
                    "C:\\Users\\Public\\Documents\\Chronos\\Methods\\Dosing Method.cam", "", "", "", "",
                    "Quantos", "", "",
                    finalCheckData[row][3] + "@" + finalCheckData[row][4].toString().slice(-14),
                    "Tray7", "",
                    finalCheckData[row][2],
                    parseFloat(finalCheckData[row][8]).toFixed(3),
                    parseFloat(7 / (parseFloat(finalCheckData[row][8] - finalCheckData[row][9]) + 0.03) + 3).toFixed(0),
                    finalCheckData[row][1] + " - Initial - " + quantosXmlDictionary[finalCheckData[row][1]]["vialOption"],
                    finalCheckData[row][3] + "@" + finalCheckData[row][4].toString().slice(-14)
                ]

                );
            } else {   // there is already dosing information present, so it must be a correction dosing
                // attach the info for this line to the array of correction dosings
                quantosXmlDictionary[finalCheckData[row][1]]["correction"].push([
                    quantosXmlDictionary[finalCheckData[row][1]]["initial"].length + 2,
                    "C:\\Users\\Public\\Documents\\Chronos\\Methods\\Dosing Method.cam", "", "", "", "",
                    "Quantos", "", "",
                    finalCheckData[row][3] + "@" + finalCheckData[row][4].toString().slice(-14),
                    "Tray7", "",
                    finalCheckData[row][2],
                    parseFloat(finalCheckData[row][8] - finalCheckData[row][9]).toFixed(3),
                    parseFloat(7 / (parseFloat(finalCheckData[row][8] - finalCheckData[row][9]) + 0.03) + 3).toFixed(0),
                    finalCheckData[row][1] + " - Correction - " + quantosXmlDictionary[finalCheckData[row][1]]["vialOption"],
                    finalCheckData[row][3] + "@" + finalCheckData[row][4].toString().slice(-14)
                ]);
            }
        }
    }

    // Go through the final dictionary and generate the Quantos input files 

    for (var key in quantosXmlDictionary) {
        // check the length of the arrays for initial and correction dosings. If they contain entries, generate the corresponding dosing files. 
        if (quantosXmlDictionary[key]["initial"].length > 0) {

            //Add the header lines
            quantosXmlDictionary[key]["initial"].unshift([1, "C:\\Users\\Public\\Documents\\Chronos\\Methods\\Set Config.cam", quantosXmlDictionary[key]["vialOption"], preDoseTapping, preDoseStrength, toleranceMode, "Quantos", useFrontDoor, useSideDoors, "", "", "", "", "", "", "", ""]);
            quantosXmlDictionary[key]["initial"].unshift(["", "Analysis Method", "Dosing Tray Type", "PreDose Tap Duration [s]", "PreDose Tap Intensity [%]", "Tolerance Mode", "Device", "Use Front Door?", "Use Side Doors?", "Substance Name", "Dosing Vial Tray", "Dosing Vial Pos.", "Dosing Vial Pos. [Axx]", "Amount [mg]", "Tolerance [%]", "Sample ID", "Comment"]);

            quantosXmlString = createQuantosXml(quantosXmlDictionary[key]["initial"]);
            initialsFolder.createFile(key + "_Final_Check_Ini.xml", quantosXmlString);     // only used for debugging purposes
            initialsFolder.createFile(key + "_Final_Check_Ini.csl", quantosXmlString);

        }
        if (quantosXmlDictionary[key]["correction"].length > 0) {
            //Add the header lines
            quantosXmlDictionary[key]["correction"].unshift([1, "C:\\Users\\Public\\Documents\\Chronos\\Methods\\Set Config.cam", quantosXmlDictionary[key]["vialOption"], preDoseTapping, preDoseStrength, toleranceMode, "Quantos", useFrontDoor, useSideDoors, "", "", "", "", "", "", "", ""]);
            quantosXmlDictionary[key]["correction"].unshift(["", "Analysis Method", "Dosing Tray Type", "PreDose Tap Duration [s]", "PreDose Tap Intensity [%]", "Tolerance Mode", "Device", "Use Front Door?", "Use Side Doors?", "Substance Name", "Dosing Vial Tray", "Dosing Vial Pos.", "Dosing Vial Pos. [Axx]", "Amount [mg]", "Tolerance [%]", "Sample ID", "Comment"]);

            quantosXmlString = createQuantosXml(quantosXmlDictionary[key]["correction"]);
            correctionsFolder.createFile(key + "_Final_Check_Corr.xml", quantosXmlString);     // only used for debugging purposes
            correctionsFolder.createFile(key + "_Final_Check_Corr.csl", quantosXmlString);

        }
        finalCheckSheet.getRange("D2:D").setValue("true"); // make sure all checkboxes are checked. 
    }
};