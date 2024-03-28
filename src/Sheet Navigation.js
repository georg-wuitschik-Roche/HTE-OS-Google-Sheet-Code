//@OnlyCurrentDoc, adapted from https://spreadsheet.dev/navigation-menu-in-google-sheets

// Use a onOpen() simple trigger to create
// a custom menu.

/**
 * Sheet Navigation: used to construct the custom menu of the gSheet.
 */
function onOpen() {
  var dataTables = SpreadsheetApp.getUi().createMenu("Data Tables")
    .addItem("HTE-Requests", "requests")
    .addItem("Component DB", "componentDb")
    .addItem("Batch DB", "batchDb")
    .addItem("SolidHotel", "solidHotel")
    .addItem("SMs/Prods", "smsProds")
    .addItem("Component Roles", "componentRoles")
    .addItem("Cmpnd_Ref_Analytics", "cmpndRefAnalytics")
    .addItem("Solutions", "solutions")
    .addItem("Plates", "plates")
    .addItem("PlateIngredients", "plateIngredients")
    .addItem("Standard Designs", "standardDesigns");

  var helperTables = SpreadsheetApp.getUi().createMenu("Helper Tables")
    .addItem("DropdownTables", "dropdownTables")
    .addItem("tempTables", "tempTables")
    .addItem("PlateBuilderHelper", "plateBuilderHelper");

  SpreadsheetApp.getUi().createMenu("SheetNavi")
    .addSubMenu(dataTables)
    .addSubMenu(helperTables)
    .addItem("Submit Request", "submitRequest")
    .addItem("PlateBuilder", "PlateBuilder")
    .addItem("FileGenerator", "FileGenerator")
    .addItem("Registration", "Registration")
    .addItem("Correction", "Correction")
    .addItem("Hotelplanner", "Hotelplanner")
    .addToUi();
}

// Activate the sheet named sheetName in the spreadsheet. 
/**
 * Sheet Navigation: used to construct the custom menu of the gSheet.
 */
function setActiveSpreadsheet(sheetName) {
  SpreadsheetApp.getActive().getSheetByName(sheetName).activate();
}

// One function per menu item.
// One of these functions will be called when users select the
// corresponding menu item from the navigation menu.
// The function then activates the sheet that the user selected.

/**
 * Sheet Navigation: used to construct the custom menu of the gSheet.
 */
function submitRequest() {
  setActiveSpreadsheet("Submit Request");
}
/**
 * Sheet Navigation: used to construct the custom menu of the gSheet.
 */
function PlateBuilder() {
  setActiveSpreadsheet("PlateBuilder");
}
/**
 * Sheet Navigation: used to construct the custom menu of the gSheet.
 */
function FileGenerator() {
  setActiveSpreadsheet("FileGenerator");
}
/**
 * Sheet Navigation: used to construct the custom menu of the gSheet.
 */
function Registration() {
  setActiveSpreadsheet("Registration");
}
/**
 * Sheet Navigation: used to construct the custom menu of the gSheet.
 */
function Correction() {
  setActiveSpreadsheet("Correction");
}
/**
 * Sheet Navigation: used to construct the custom menu of the gSheet.
 */
function Hotelplanner() {
  setActiveSpreadsheet("Hotelplanner");
}
/**
 * Sheet Navigation: used to construct the custom menu of the gSheet.
 */
function requests() { setActiveSpreadsheet('HTE-Requests'); }
/**
 * Sheet Navigation: used to construct the custom menu of the gSheet.
 */
function componentDb() { setActiveSpreadsheet('Component DB'); }
/**
 * Sheet Navigation: used to construct the custom menu of the gSheet.
 */
function batchDb() { setActiveSpreadsheet('Batch DB'); }
/**
 * Sheet Navigation: used to construct the custom menu of the gSheet.
 */
function solidHotel() { setActiveSpreadsheet('SolidHotel'); }
/**
 * Sheet Navigation: used to construct the custom menu of the gSheet.
 */
function smsProds() { setActiveSpreadsheet('SMs/Prods'); }
/**
 * Sheet Navigation: used to construct the custom menu of the gSheet.
 */
function componentRoles() { setActiveSpreadsheet('Component Roles'); }
/**
 * Sheet Navigation: used to construct the custom menu of the gSheet.
 */
function cmpndRefAnalytics() { setActiveSpreadsheet('Cmpnd_Ref_Analytics'); }
/**
 * Sheet Navigation: used to construct the custom menu of the gSheet.
 */
function solutions() { setActiveSpreadsheet('Solutions'); }
/**
 * Sheet Navigation: used to construct the custom menu of the gSheet.
 */
function plates() { setActiveSpreadsheet('Plates'); }
/**
 * Sheet Navigation: used to construct the custom menu of the gSheet.
 */
function plateIngredients() { setActiveSpreadsheet('PlateIngredients'); }
/**
 * Sheet Navigation: used to construct the custom menu of the gSheet.
 */
function wells() { setActiveSpreadsheet('Wells'); }
/**
 * Sheet Navigation: used to construct the custom menu of the gSheet.
 */
function colRowLabels() { setActiveSpreadsheet('ColRowLabels'); }
/**
 * Sheet Navigation: used to construct the custom menu of the gSheet.
 */
function standardDesigns() { setActiveSpreadsheet('Standard Designs'); }
/**
 * Sheet Navigation: used to construct the custom menu of the gSheet.
 */
function dropdownTables() { setActiveSpreadsheet('DropdownTables'); }
/**
 * Sheet Navigation: used to construct the custom menu of the gSheet.
 */
function tempTables() { setActiveSpreadsheet('tempTables'); }

function plateBuilderHelper() { setActiveSpreadsheet('PlateBuilderHelper'); }
