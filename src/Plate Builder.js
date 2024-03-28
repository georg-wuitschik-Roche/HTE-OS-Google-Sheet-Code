
/**
 * Platebuilder: Switch columns / rows in PlateBuilder, connected to the "Switch" button.
 * 
 */
function transposePlates() {  //this function switches the components in rows and columns in PlateBuilder

  //connect to the sheet

  var plateBuilderSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PlateBuilder");

  // read the two arrays

  var oldColumnValues = plateBuilderSheet.getRange("D1:R3").getValues();
  var oldRowValues = plateBuilderSheet.getRange("A4:C14").getValues();

  // add for additional columns into the new columns array to adapt to the max number of 12 columns

  oldRowValues.splice(10, 0, ['', '', '']);
  oldRowValues.splice(10, 0, ['', '', '']);
  oldRowValues.splice(10, 0, ['', '', '']);
  oldRowValues.splice(10, 0, ['', '', '']);

  // transpose the arrays

  var newRowValues = transposeArray(oldColumnValues);
  var newColumnValues = transposeArray(oldRowValues);

  // strip the four columns of reagents not fitting into max 8 lines available for rows

  newRowValues.splice(10, 4);

  console.log(newRowValues);

  // write back both arrays

  plateBuilderSheet.getRange("D1:R3").setValues(newColumnValues);
  plateBuilderSheet.getRange("A4:C14").setValues(newRowValues);

}
/**
 * Platebuilder: transposes a 2-D array which is needed in transposePlates().
 * @param {Array} a 2-D array to be transposed.
 */
function transposeArray(a) { // taken from https://stackoverflow.com/questions/4492678/swap-rows-with-columns-transposition-of-a-matrix-in-javascript
  return Object.keys(a[0]).map(function (c) {
    return a.map(function (r) { return r[c]; });
  });
}
/**
 * Platebuilder: only calls the loadPlateDesign function, connected to the "LoadDesign" button.
 */
function loadDesignButton() {//This function is called by the "Load Plate Design" Button in Plate Builder
  loadPlateDesign("Standard Designs", 0);
}

/**
 * Platebuilder: only calls the savePlateDesign function, connected to the "SaveDesign" button.
 */
function saveDesignButton() { //This function is called by the "Save Plate Design" Button in Plate Builder
  //It's set up this way because the savePlateDesign function is also used to save wholes plates in File Generator
  savePlateDesign("Standard Designs");
}

/**
 * Platebuilder: Loads a plate design either from the Standard Designs sheet or from the plates sheet
 * @param {String} sourceSheet Name of the source sheet from which a plate design should be loaded.
 * @param {Number} rowNumber number of the row where the plate design is found in the sheet from which the plate design is loaded.
 * 
 */
function loadPlateDesign(sourceSheet, rowNumber) {
  // connect to the two sheets
  var writeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PlateBuilder");
  var readSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sourceSheet);
  var cellData = [];
  var dataRow;
  // get the number of the row that contains the data

  var columnOffset = 0;

  switch (sourceSheet) {
    case "Standard Designs":
      columnOffset = 0;
      dataRow = writeSheet.getRange("T1").getValue(); // This cell contains the number of the row, where the selected design is saved in the Sheet Standard Designs or "New Design", if cell A2 in Plate Builder contains an unknown design name
      break;
    case "Plates":
      columnOffset = 19;
      dataRow = rowNumber;
      break;
  }

  if (dataRow == "New Design") { //Script is aborted, if the design is unknown
    Browser.msgBox("This Design is unknown.");
    return;
  }
  //Get the data and write it afterwards to the PlateBuilderSheet 

  cellData = readSheet.getRange(dataRow, 2 + columnOffset, 1, 7).getValues();

  if (cellData[0][4] != "") {  // true, if it's a broad design
    var componentFoundFlag = 0;
    var broadScreenComponents = textToArray(cellData[0][4] + cellData[0][5] + cellData[0][6]);
    writeSheet.getRange("C16").setValue(true); // activate the checkbox
    writeSheet.getRange("C4:C5").setValues([[broadScreenComponents[0][1]], [0]]); // type of component and number components (must be 0, since it's a broad screen)
    broadScreenComponents.shift(); //remove the first line, since it won't be needed anymore
    SpreadsheetApp.flush();
    Utilities.sleep(5000);
    var currentState = writeSheet.getRange("E18:F").getValues();
    for (var row = 0; row < currentState.length; row++) {
      currentState[row][0] = false;
      if (currentState[row][1] == "") {
        currentState.length = row; // if an empty cell is detected, stop and remove the rest of the rows
        break;
      }

      //Browser.msgBox(currentState)
    }
    console.log(currentState);
    for (var broadScreenComponentsRow = 0; broadScreenComponentsRow < broadScreenComponents.length; broadScreenComponentsRow++) { //go through each line of the Components to be loaded and check if there the tickbox is checked
      componentFoundFlag = 0;
      if (broadScreenComponents[broadScreenComponentsRow][0] == "true") {
        for (var currentStateRow = 0; currentStateRow < currentState.length; currentStateRow++) {
          if (currentState[currentStateRow][1] == broadScreenComponents[broadScreenComponentsRow][1]) { //if the component names are identical, then set it to true
            currentState[currentStateRow][0] = true;
            componentFoundFlag = 1; //signal that the component was found
            break;
          }
        }
        if (componentFoundFlag == 0) {
          Browser.msgBox(broadScreenComponents[broadScreenComponentsRow][1] + " was not found. Has maybe the Component Name changed?");
          console.log(broadScreenComponents[broadScreenComponentsRow][1]);
        }
      }

    }
    for (let currentStateRow = 0; currentStateRow < currentState.length; currentStateRow++) {
      currentState[currentStateRow].pop(); //the component name is not needed anymore and thus removed
    }

    writeSheet.getRange(18, 5, currentState.length, 1).setValues(currentState); // set the checkboxes accordingly
    writeSheet.getRange("F6:Q13").setFontSize(8); // set the fontsize of the screening plate,since the onEdit function is not triggered when a change is made using a script
    writeSheet.getRange("X6:AI13").setFontSize(8);
  } else if (writeSheet.getRange("C16").isChecked()) { //the plate loaded is a not a broad screen, but the broad screen tickbox is active
    writeSheet.getRange("C16").setValue(false); //untick the checkbox
    writeSheet.getRange("F6:Q13").setFontSize(27); // set the fontsize of the screening plate,since the onEdit function is not triggered when a change is made using a script
    writeSheet.getRange("X6:AI13").setFontSize(27);
  }
  writeSheet.getRange("A4:C4").setValues([textToArray(cellData[0][2])[0]]);  // one needs to write the reagent types for the rows first separately, otherwise there's a conflict with data validation
  SpreadsheetApp.flush(); //Without flushing the source ranges for the data validation aren't updated and thus the script fails.

  writeSheet.getRange("D1:R3").setValues(textToArray(cellData[0][0])); // We don't know why the script doesn't fail in this case, since it's the exact equivalent to the case above...
  writeSheet.getRange("X4:HZ4").setValues(textToArray(cellData[0][1]));
  writeSheet.getRange("V6:V150").setValues(textToArray(cellData[0][3]));
  writeSheet.getRange("A4:C14").setValues(textToArray(cellData[0][2]));

}
/**
 * Platebuilder: Saves a plate design either either to the Standard Designs sheet or to the plates sheet
 * @param {String} targetSheet Name of the source sheet from which a plate design should be loaded.
 * 
 */
function savePlateDesign(targetSheet) {
  // connect to the two sheets
  var readSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PlateBuilder");
  var writeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheet);
  var columnOffset = 0;
  var dataRow;

  switch (targetSheet) {
    case "Standard Designs":
      columnOffset = 0;
      dataRow = readSheet.getRange("T1").getValue();
      break;
    case "Plates":
      columnOffset = 19;
      break;
  }


  if (dataRow != "New Design" && targetSheet == "Standard Designs") {
    Browser.msgBox("The name of this Plate Design exists already.");
    return;
  }

  // get the number of the last row that contains data
  var lastRow = writeSheet.getLastRow();

  //get the design Name that was entered
  var designName = readSheet.getRange("A2").getValue();

  // get the data in columns and rows as well as the state of the corresponding switches

  var columnsData = readSheet.getRange("D1:R3").getValues();
  var columnsDataString = "";

  var columnsSwitches = readSheet.getRange("X4:HZ4").getValues();
  var columnsSwitchesString = "";

  var rowsData = readSheet.getRange("A4:C14").getValues();
  var rowsDataString = "";

  var rowsSwitches = readSheet.getRange("V6:V150").getValues();
  var rowsSwitchesString = "";

  var broadDesignFlag = readSheet.getRange("C16").getValue();

  if (broadDesignFlag == true) {

    var broadScreenComponents = readSheet.getRange("E17:F").getValues(); //contains the category to be screened as well as a list of category members and which were selected. 
    for (var row = 1; row < broadScreenComponents.length; row++) {
      if (broadScreenComponents[row][1] == "") {
        broadScreenComponents.length = row; // if an empty cell is detected, stop and remove the rest of the rows
        break;
      }
    }
    var broadScreenComponentsString = arrayToText(broadScreenComponents);

    if (broadScreenComponentsString.length < 50001) {
      writeSheet.getRange(lastRow + 1, 6 + columnOffset).setValue(broadScreenComponentsString);
    } else if (broadScreenComponentsString.length < 100001) {
      writeSheet.getRange(lastRow + 1, 6 + columnOffset).setValue(broadScreenComponentsString.substring(0, 50000));
      writeSheet.getRange(lastRow + 1, 7 + columnOffset).setValue(broadScreenComponentsString.substring(50000, 100000));
    } else {
      writeSheet.getRange(lastRow + 1, 6 + columnOffset).setValue(broadScreenComponentsString.substring(0, 50000));
      writeSheet.getRange(lastRow + 1, 7 + columnOffset).setValue(broadScreenComponentsString.substring(50000, 100000));
      writeSheet.getRange(lastRow + 1, 8 + columnOffset).setValue(broadScreenComponentsString.substring(100000, 150000));
    }

  }

  //Write the name of the design in column 1
  writeSheet.getRange(lastRow + 1, 1 + columnOffset).setValue(designName);

  // convert the array D1:R3 into a format that is later easily readable when loading the design 
  columnsDataString = arrayToText(columnsData);
  //Write the column Data to column 2
  writeSheet.getRange(lastRow + 1, 2 + columnOffset).setValue(columnsDataString);

  columnsSwitchesString = arrayToText(columnsSwitches);
  writeSheet.getRange(lastRow + 1, 3 + columnOffset).setValue(columnsSwitchesString);

  rowsDataString = arrayToText(rowsData);
  writeSheet.getRange(lastRow + 1, 4 + columnOffset).setValue(rowsDataString);

  rowsSwitchesString = arrayToText(rowsSwitches);
  writeSheet.getRange(lastRow + 1, 5 + columnOffset).setValue(rowsSwitchesString);
}


/**
 * Platebuilder: This helper function converts arrays into a notation that can then be converted back to an array by Google Apps Script when loaded.
 * Using the in-built function will remove all the brackets from the array.
 * @param {Array} array Array to be converted into text by combining the individual members.
 * @return {String} Ouput text.
 */
function arrayToText(array) {
  var arrayString = "";
  for (var i = 0; i < array.length; i++) {
    arrayString = arrayString + "[" + array[i].join("-!-") + "],";

  }
  arrayString = arrayString.slice(1, -2);
  return arrayString;
}


/**
 * Platebuilder: This function takes the text found in the cell and converts it back into an array that can be written back to a range in the PlateBuilder Sheet
 * adapted from https://stackoverflow.com/questions/15783169/convert-string-to-matrix-array
 * @param {String} inputString string to be split up into individual members of an array which is then returned.
 * @return {Array} Ouput array.
 */
function textToArray(inputString) {
  var temp = inputString.split('],[');
  //Logger.log(temp)
  var row = [];
  var len = temp.length;

  for (n = 0; n < len; ++n) {
    row.push(temp[n].split('-!-'));

  }
  return row;
}

/**
 * Platebuilder: This function waits for an edit event in the plate builder sheet to happen and updates the corresponding formulas in the different layers
 */
function onEdit(e) {

  var editedSheet = e.source.getSheetName();
  // check if the edit happened on PlateBuilder, exit function if not.
  if (editedSheet != "PlateBuilder") return;
  var coordinate = e.range.getA1Notation();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PlateBuilder");

  // if the broad screen checkbox is clicked and the value of the checkbox now is true, then set the number of variations to 0, remove whatever is below, set levels to 1 and change the font size so whatever is chosen becomes legible, reverse the operation if it is unchecked.

  if (coordinate == 'C16') {
    if (sheet.getRange("C16").isChecked()) {
      sheet.getRange("F6:Q13").setFontSize(8);
      sheet.getRange("X6:AI13").setFontSize(8);
      sheet.getRange("A5:C14").setValues([[0, 0, 0], ["", "", ""], ["", "", ""], ["", "", ""], ["", "", ""], ["", "", ""], ["", "", ""], ["", "", ""], ["", "", ""], [1, 1, 1]]);
    } else {
      sheet.getRange("F6:Q14").setFontSize(27);
      sheet.getRange("X6:AI13").setFontSize(27);
      sheet.getRange("A5:C14").setValues([[1, 1, 1], ["", "", ""], ["", "", ""], ["", "", ""], ["", "", ""], ["", "", ""], ["", "", ""], ["", "", ""], ["", "", ""], [1, 1, 1]]);
    }
  }

  //load the PlateBuilderHelper sheet and two ranges therein which indicate if any of the formulas in F1:Q3 or A6:C13 were overwritten.
  var helperSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PlateBuilderHelper");
  var helperSheetData = helperSheet.getRange(38, 34, 2, 3).getValues(); //contains formulas indicating whether any of the formulas in Platebuilder!F1:Q3 and A6:C13 were changed
  var helperColIndValues = helperSheetData[1];
  var helperLineIndValues = helperSheetData[0];


  // if all cells contain formulas, each of the cells must be 1, thus their product as well and then the function is exited, because there was no change to the formulas in F1:Q3 and A6:C13
  if (helperColIndValues[0][0] * helperColIndValues[0][1] * helperColIndValues[0][2] * helperLineIndValues[0][0] * helperLineIndValues[0][1] * helperLineIndValues[0][2] == 1) return;



  //The next three if-statements check whether any of the formulas in the respective rows has been changed and whether the change that triggered onEdit resulted from cells D1:D3
  // If that's the case, the formulas in the respective area are restored.
  // The if's are only triggered, if the category of compound to be screened in the respective row is changed. Thus, amendments to individual components or additions are preserved both in that row, but also in others.
  if (helperColIndValues[0] < 1 && coordinate == 'D1') {
    //restore the formulas in the first line
    sheet.getRange(1, 6, 1, 12).setValues([[
      '=if($E$1>0,PlateBuilderHelper!Y31,"")',
      '=if($E$1>0,PlateBuilderHelper!Z31,"")',
      '=if($E$1>0,PlateBuilderHelper!AA31,"")',
      '=if($E$1>0,PlateBuilderHelper!AB31,"")',
      '=if($E$1>0,PlateBuilderHelper!AC31,"")',
      '=if($E$1>0,PlateBuilderHelper!AD31,"")',
      '=if($E$1>0,PlateBuilderHelper!AE31,"")',
      '=if($E$1>0,PlateBuilderHelper!AF31,"")',
      '=if($E$1>0,PlateBuilderHelper!AG31,"")',
      '=if($E$1>0,PlateBuilderHelper!AH31,"")',
      '=if($E$1>0,PlateBuilderHelper!AI31,"")',
      '=if($E$1>0,PlateBuilderHelper!AJ31,"")'
    ]]);
  }

  if (helperColIndValues[1] < 1 && coordinate == 'D2') {
    //restore the formulas in the second line
    sheet.getRange(2, 6, 1, 12).setValues([[
      '=if($E$2>0,PlateBuilderHelper!Y32,"")',
      '=if($E$2>0,PlateBuilderHelper!Z32,"")',
      '=if($E$2>0,PlateBuilderHelper!AA32,"")',
      '=if($E$2>0,PlateBuilderHelper!AB32,"")',
      '=if($E$2>0,PlateBuilderHelper!AC32,"")',
      '=if($E$2>0,PlateBuilderHelper!AD32,"")',
      '=if($E$2>0,PlateBuilderHelper!AE32,"")',
      '=if($E$2>0,PlateBuilderHelper!AF32,"")',
      '=if($E$2>0,PlateBuilderHelper!AG32,"")',
      '=if($E$2>0,PlateBuilderHelper!AH32,"")',
      '=if($E$2>0,PlateBuilderHelper!AI32,"")',
      '=if($E$2>0,PlateBuilderHelper!AJ32,"")'
    ]]);
  }

  if (helperColIndValues[2] < 1 && coordinate == 'D3') {
    //Logger.log("Col3 triggered")

    //restore the formulas in the third line
    sheet.getRange(3, 6, 1, 12).setValues([[
      '=if($E$3>0,PlateBuilderHelper!Y33,"")',
      '=if($E$3>0,PlateBuilderHelper!Z33,"")',
      '=if($E$3>0,PlateBuilderHelper!AA33,"")',
      '=if($E$3>0,PlateBuilderHelper!AB33,"")',
      '=if($E$3>0,PlateBuilderHelper!AC33,"")',
      '=if($E$3>0,PlateBuilderHelper!AD33,"")',
      '=if($E$3>0,PlateBuilderHelper!AE33,"")',
      '=if($E$3>0,PlateBuilderHelper!AF33,"")',
      '=if($E$3>0,PlateBuilderHelper!AG33,"")',
      '=if($E$3>0,PlateBuilderHelper!AH33,"")',
      '=if($E$3>0,PlateBuilderHelper!AI33,"")',
      '=if($E$3>0,PlateBuilderHelper!AJ33,"")'
    ]]);
  }

  //The next three if-statements check whether any of the formulas in the respective rows has been changed and whether the change that triggered onEdit resulted from cells A4:C4
  // If that's the case, the formulas in the respective area are restored.
  if (helperLineIndValues[0] < 1 && coordinate == 'A4') {

    //restore the formulas in the first column
    sheet.getRange(6, 1, 8, 1).setValues([
      ['=if($A$5>0, PlateBuilderHelper!Y42, "")'],
      ['=if($A$5>0, PlateBuilderHelper!Y43, "")'],
      ['=if($A$5>0, PlateBuilderHelper!Y44, "")'],
      ['=if($A$5>0, PlateBuilderHelper!Y45, "")'],
      ['=if($A$5>0, PlateBuilderHelper!Y46, "")'],
      ['=if($A$5>0, PlateBuilderHelper!Y47, "")'],
      ['=if($A$5>0, PlateBuilderHelper!Y48, "")'],
      ['=if($A$5>0, PlateBuilderHelper!Y49, "")']
    ]);
  }

  if (helperLineIndValues[1] < 1 && coordinate == 'B4') {
    //restore the formulas in the second column
    sheet.getRange(6, 2, 8, 1).setValues([
      ['=if($B$5>0, PlateBuilderHelper!Z42, "")'],
      ['=if($B$5>0, PlateBuilderHelper!Z43, "")'],
      ['=if($B$5>0, PlateBuilderHelper!Z44, "")'],
      ['=if($B$5>0, PlateBuilderHelper!Z45, "")'],
      ['=if($B$5>0, PlateBuilderHelper!Z46, "")'],
      ['=if($B$5>0, PlateBuilderHelper!Z47, "")'],
      ['=if($B$5>0, PlateBuilderHelper!Z48, "")'],
      ['=if($B$5>0, PlateBuilderHelper!Z49, "")']
    ]);

  }

  if (helperLineIndValues[2] < 1 && coordinate == 'C4') {

    //restore the formulas in the third column
    sheet.getRange(6, 3, 8, 1).setValues([
      ['=if($C$5>0, PlateBuilderHelper!AA42, "")'],
      ['=if($C$5>0, PlateBuilderHelper!AA43, "")'],
      ['=if($C$5>0, PlateBuilderHelper!AA44, "")'],
      ['=if($C$5>0, PlateBuilderHelper!AA45, "")'],
      ['=if($C$5>0, PlateBuilderHelper!AA46, "")'],
      ['=if($C$5>0, PlateBuilderHelper!AA47, "")'],
      ['=if($C$5>0, PlateBuilderHelper!AA48, "")'],
      ['=if($C$5>0, PlateBuilderHelper!AA49, "")']
    ]);

  }
}


