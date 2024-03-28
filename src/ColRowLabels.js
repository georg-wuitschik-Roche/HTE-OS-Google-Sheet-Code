/*jshint sub:true*/

/**
 * ColRowLabels: Generates the formatted plate ingredient to be used in the results presentation
 * @param {Array} data Line from the PlateIngredients Sheet
 * @return {String} Formatted Plate Ingredient as String depending on the type of compound.
 */
function generateFormattedPlateIngredient(data) {
  if (data[12] === 0 || data[14] === 0) return ""; // if, both equiv and ml/g are zero. This happens if there are more than two levels of a compound, one of which is 0. 
  var formattedPlateIngredient = "not specified";
  var componentName = data[3] + " (";
  if (data[6] == "Batch: different") { // in the rare event, when the same component is present with different batches, the batch ID needs to be part of the string in the presentation
    componentName += "Batch: " + data[7] + ", ";
  }
  var solutionString = "not specified";

  if (data[16] == "Solution") {//if dosed as a solution
    if (data[20] == "water" || data[20] == ", aq") {
      solutionString = ", " + data[18] + data[19] + ", aq";
    } else {
      solutionString = ", " + data[18] + data[19] + " in " + data[20];
    }
  } else {//component is not dosed as a solution
    solutionString = "";
  }

  switch (String(data[1]).substring(0, 8)) {
    case "Catalyst":
    case "Ligands":
    case "Additive":
    case "Internal":
      if ((parseFloat(data[12]) * 1000).toFixed(0) % 10 > 0) { //the mod10 of number of eq * 1000 will be bigger than 0, if y in x.ymol% is bigger than 0 and thus worth reporting
        formattedPlateIngredient = componentName + ((parseFloat(data[12]) * 100).toFixed(1)) + "mol%)";
      } else { formattedPlateIngredient = componentName + ((parseFloat(data[12]) * 100).toFixed(0)) + "mol%" + solutionString + ")"; } //in most cases it'll be a integer number of mol% without the need to report the .0
      break;
    case "Solvents":
      formattedPlateIngredient = componentName + data[14] + "vol" + solutionString + ")";
      break;
    case "Other Va": // Component Name contains all there is to be known about freetext variables like temperature, order of addition etc
      if (isNaN(data[12])) {   // typically (Enter Quantity) and thus not a number
        formattedPlateIngredient = componentName.slice(0, -2); // remove the " (" again
      } else if (String(data[3]).substring(1, 4) == "emp") {  //  covers Temp, Temperature, temperature...
        formattedPlateIngredient = componentName + (parseFloat(data[12])).toFixed(0) + " °C" + ")";
      } else if (String(data[3]).substring(1, 5) == "ress") {  //  covers Pressure and pressure and press...
        formattedPlateIngredient = componentName + (parseFloat(data[12])).toFixed(0) + " bar" + ")";
      }
      break;
    default:
      if (data[14] === "-" || data[14] === "") {
        if (parseFloat(data[12]).toFixed(0) === 0) {
          formattedPlateIngredient = componentName + (parseFloat(data[12])) + " eq" + solutionString + ")";
        } else {
          formattedPlateIngredient = componentName + (parseFloat(data[12]).toFixed(1)) + " eq" + solutionString + ")";
        }
      } else {
        formattedPlateIngredient = componentName + data[14] + "vol" + solutionString + ")";
      }
      break;
  }

  return formattedPlateIngredient;
}

