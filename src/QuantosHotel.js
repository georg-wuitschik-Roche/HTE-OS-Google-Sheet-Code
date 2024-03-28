/**
 * QuantosHotel:  readQuantosHotel (Quantoshotel isn't called directly, since this also opens the possibility to read hotels from other solid dosing robots)
 */
function readHotel() {

  readQuantosHotel();
}


/**
 * QuantosHotel: reads the contents of the Quantos Hotel Config file in HTS Docs/Project Data/Quantos Hotel Config and puts the important fields into the QuantosHotel Sheet
 */
function readQuantosHotel() {


  var quantosHotelSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SolidHotel");
  var batchDbSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Batch DB");
  var batchDbSheetDataRange = batchDbSheet.getDataRange();
  var batchDbSheetContent = batchDbSheetDataRange.getValues();
  var arrayOfBatchKeys = []; //so we can check later if a batchkey already exists
  var changeFlag = 0; //a change to 1 indicates that a change has occurred compared to what was last written to the batch db
  var headNotFoundFlag = 0;
  var quantosHotelSheetContent = [];
  var quantosHeadName = "";
  var remainingQuantityInHead = "";
  var headType = ""; // Contains the type of dosing head 
  var headInformation = {}; // Dictionary later holding the information extracted from the xml with the head position as the key
  var quantosHotelXmlFile = DriveApp.getFileById(globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["QUANTOShOTELfILEiD"]);
  var quantosHotelXmlString = quantosHotelXmlFile.getBlob().getDataAsString();  //get the content of the xml-file
  //var lastUpdated = quantosHotelXmlFile.getLastUpdated();

  var document = XmlService.parse(quantosHotelXmlString);
  var root = document.getRootElement();
  var indivHead = root.getChild("Rack").getChild("Heads").getChildren(); //  now at the level of the individual heads

  for (var row = 1; row < batchDbSheetContent.length; row++) {
    arrayOfBatchKeys.push(batchDbSheetContent[row][6]);
  }


  for (head = 0; head < indivHead.length; head++) { // go through all the Heads and append the content to the headInfo dictionary with the head position as key
    var headInfo = indivHead[head].getChild("HeadInfo");
    if (headInfo.getAttributes().length > 0) {
      headInformation[indivHead[head].getChild("Position").getText()] = ["-",
        "-",
        "-",
        "-",
        "-"];

    }
    else {
      quantosHeadName = headInfo.getChild("Substance").getText();
      remainingQuantityInHead = headInfo.getChild("Rem._quantity").getText();
      headType = headInfo.getChild("Head_type").getText();
      headInformation[indivHead[head].getChild("Position").getText()] =
        [quantosHeadName,
          headInfo.getChild("Lot_ID").getText(),
          headInfo.getChild("Rem._dosages").getText(),
          headInfo.getChild("Dose_limit").getText(),
          remainingQuantityInHead];
      headNotFoundFlag = 0; // resets the flag before checking batchDB if the next head is found
      for (row = 0; row < batchDbSheetContent.length; row++) { // go through the content of the batch DB sheet
        if (quantosHeadName == batchDbSheetContent[row][1] + "@" + (batchDbSheetContent[row][2]).toString().slice(-14)) {
          if (remainingQuantityInHead != (batchDbSheetContent[row][8]).toString()) { // If the Lot_ID of the current head matches the one constructed from Component ID and truncated Batch ID and if the remaining quantity of material recorded in batch DB different from what is in the head, then update the amount of material if it's different 
            changeFlag = 1;
            batchDbSheetContent[row][8] = remainingQuantityInHead;
            switch (headType) {
              case "QH002-CNMW":
                headType = " small plastic";
                break;
              default:
                headType = " non-plastic: " + headType;
            }

            batchDbSheetContent[row][7] = (batchDbSheetContent[row][7]).toString().split("-- ")[0] + "-- " + headType; //preserves the text in front of the --
            headNotFoundFlag = 1;

            break;
          }
          headNotFoundFlag = 1;
          changeFlag = 1;
        }
      }
      //console.log(quantosHeadName + " : " + headNotFoundFlag);
      if (headNotFoundFlag == 0 && quantosHeadName.includes("@") && arrayOfBatchKeys.includes(quantosHeadName.split("@")[0] + "_" + quantosHeadName.split("@")[1]) == false) { // Add a new line to the batchDB in case the head doesn't exist yet.
        batchDbSheetContent.push([batchDbSheetContent.length + 1,
        quantosHeadName.split("@")[0],
        quantosHeadName.split("@")[1],
          "Please add producer",
          "",
        "=VLOOKUP(B" + (batchDbSheetContent.length + 1) + ",'Component DB'!B2:C,2,false)",
        quantosHeadName.split("@")[0] + "_" + quantosHeadName.split("@")[1],
        "-- " + headType,
          remainingQuantityInHead]);
      }

    }
  }
  for (let key in headInformation) { // go through all keys of the dictionary and append the content to the array of heads to be written
    quantosHotelSheetContent.push([key, headInformation[key][0], headInformation[key][1], headInformation[key][2], headInformation[key][3], headInformation[key][4]]);
  }
  if (changeFlag > 0) {// the batch DB sheet is quite large. It should only be written if there's a change. 
    console.log("write");
    batchDbSheet.getRange(1, 1, batchDbSheetContent.length, batchDbSheetContent[0].length).setValues(batchDbSheetContent);
  }
  quantosHotelSheet.getRange(2, 1, quantosHotelSheetContent.length, quantosHotelSheetContent[0].length).setValues(quantosHotelSheetContent);


}
