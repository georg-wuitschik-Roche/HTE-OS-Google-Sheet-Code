/*jshint sub:true*/
/**
 * Generate Results Presentation: Creates a new gSlide presenation in the given folder with the given filename (analogous to the createSpreadsheet-function ).
 * @param {String} folderID id of the folder in which the presenation is to be created.
 * @param {String} fileName name of the new presentation.
 * @return {String} fileID of the newly created presentation.
 */
function createPresentation(folderID, fileName) {
    var parentFolder = DriveApp.getFolderById(folderID);
    var filesInFolder = parentFolder.getFiles();
    var doesntExist = true;
    var newFile = '';

    // Check if file already exists.
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

        var template = DriveApp.getFileById(globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["PRESENTATIONtEMPLATEiD"]);
        var copy = template.makeCopy();
        copy.setName(fileName);
        var fileId = copy.getId();
        //move the copy to the experiment folder, not necessary, since the template is in the same folder, also wouldn't work with Shared Drive
        //Drive.Files.update({ parents: [{ id: folderID }] }, fileId); // source: https://tanaikech.github.io/2019/11/20/moving-file-to-specific-folder-using-google-apps-script/


        return fileId;
    }

}

/**
 * Add sideproducts to be registered to all relevant slides of the results presentation.
 * @param {Object} reactionComponents Dictionary containing all information on side products to be registered.
 */
function addNewSideProductToPresentation(reactionComponents) {
    //connect to the presentation to be modified

    var fileIdPresentation = createPresentation(globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["PRESENTATIONfOLDERiD"], reactionComponents["otherInfo"].ElnId + " Results");

    var presentation = SlidesApp.openById(fileIdPresentation);
    var slideHeight = presentation.getPageHeight();

    //get the individual slides to manipulate them
    var slides = presentation.getSlides();
    var tables = slides[1].getTables();
    var slideCounter = 0;
    var relevantImageCounter = 0;

    //go through all the slides and amend the tables of Overall conclusion (slide 2) and the individual Plate Conclusions (unknown positions, since the number of slides per plate is unpredictable) 

    //Amend the table of the Overall Conclusion slide 

    tables = slides[1].getTables();
    var images = slides[1].getImages();
    var reactionEquation = tables[tables.length - 1];
    var colCount = reactionEquation.getNumColumns();
    var numberOfExistingStructuresOnSlide = colCount / 2;    // the number of structures is always equal to half the number of table columns, additional columns contain + signs, the reaction arrow and the row headers
    var groups = slides[1].getGroups();

    for (var currentImageCount = 0; currentImageCount < images.length; currentImageCount++) {
        if (images[currentImageCount].getTop() < slideHeight / 7.0 + 50 && images[currentImageCount].getTop() > slideHeight / 7.0 - 50) {//true, if the top of the picture in question is in a band +-50 points of the originally set height of slideHeight/7.

            var imageDescription = amendStructureScalingAndPosition(images[relevantImageCounter], (colCount) / 2, Object.keys(reactionComponents).length - 1, relevantImageCounter);
            //console.log(imageDescription)
            if (imageDescription == "Product" && groups.length > 0) {
                groups[0].setLeft(images[relevantImageCounter].getLeft() - groups[0].getWidth() + 30);    //sets the reaction arrow/description to the correct x-position. 
            }
            relevantImageCounter++;
        }
    }

    for (var key in reactionComponents) {
        if (key == "otherInfo") { continue; }
        // add two columns, the first one containing a "+" sign and the second one the side product data
        reactionEquation.appendColumn();
        colCount++;
        addImageToSlide(presentation, slides, reactionComponents, key, colCount, 25, 1, numberOfExistingStructuresOnSlide);

        reactionEquation.getColumn(colCount - 1).getCell(1).getText().setText("+");
        reactionEquation.appendColumn();
        colCount++;
        reactionEquation.getColumn(colCount - 1).getCell(0).getText().setText(reactionComponents[key].ComponentName);
        reactionEquation.getColumn(colCount - 1).getCell(1).getText().setText(
            getRetentionTimePrediction(reactionComponents[key].Smiles)
        ).getTextStyle().setForegroundColor('#c40000');
        reactionEquation.getColumn(colCount - 1).getCell(2).getText().setText(parseFloat(reactionComponents[key].MW).toFixed(2));
        reactionEquation.getColumn(colCount - 1).getCell(3).getText().setText(reactionComponents[key].SideReactionType);


    }

    //Iterate through all the slides and all shapes of each slide: If on any given slide there is a shape with the text "Cheatsheet" it must be a plate conclusion slide with a table that needs amending

    // code adapted from https://gist.github.com/atwellpub/9e217c49e840e3d9709dfbf847b5fa62
    slides.forEach(function (slide) {
        var shapes = (slide.getShapes());
        shapes.forEach(function (shape) {
            var text = shape.getText();
            // If a wash-replacement of Cheatsheet to Cheatsheet is successful, then it's a plate conclusion slide
            var m = text.replaceAllText('Plate Conclusion Slide', 'Plate Conclusion Slide');
            if (m > 0) {
                var groups = slide.getGroups();
                tables = slide.getTables();
                var images = slide.getImages();
                reactionEquation = tables[tables.length - 1]; // normally there's only one table 
                colCount = reactionEquation.getNumColumns();
                numberOfExistingStructuresOnSlide = colCount / 2;
                console.log("numberOfExistingStructuresOnSlide: ", numberOfExistingStructuresOnSlide);
                relevantImageCounter = 0;
                for (var currentImageCount = 0; currentImageCount < images.length; currentImageCount++) {
                    if (images[currentImageCount].getTop() < slideHeight / 7.0 + 50 && images[currentImageCount].getTop() > slideHeight / 7.0 - 50) {//true, if the top of the picture in question is in a band +-50 points of the originally set height of slideHeight/7.

                        var imageDescription = amendStructureScalingAndPosition(images[relevantImageCounter], (colCount) / 2, Object.keys(reactionComponents).length - 1, relevantImageCounter);
                        console.log(imageDescription);
                        if (imageDescription == "Product" && groups.length > 0) { // adjusts the position of the reaction arrow/description so that it's right side is approximately flush with the left side of the product structure image.
                            groups[0].setLeft(images[relevantImageCounter].getLeft() - groups[0].getWidth() + 30);
                        }
                        relevantImageCounter++;
                    }
                }

                // now add the additional columns/pictures for the new sideproducts. 

                for (var key in reactionComponents) { //go through the dictionary of newly registered side products and append them to the table
                    if (key == "otherInfo") { continue; }
                    // add two columns, the first one containing a "+" sign and the second one the side product data
                    reactionEquation.appendColumn();
                    colCount++;
                    addImageToSlide(presentation, slides, reactionComponents, key, colCount, 55, slideCounter, numberOfExistingStructuresOnSlide);

                    reactionEquation.getColumn(colCount - 1).getCell(1).getText().setText("+");
                    reactionEquation.appendColumn();
                    colCount++;
                    reactionEquation.getColumn(colCount - 1).getCell(0).getText().setText("-");
                    reactionEquation.getColumn(colCount - 1).getCell(1).getText().setText(reactionComponents[key].ComponentName);
                    reactionEquation.getColumn(colCount - 1).getCell(2).getText().setText(parseFloat(reactionComponents[key].MW).toFixed(2));
                    reactionEquation.getColumn(colCount - 1).getCell(3).getText().setText(reactionComponents[key].SideReactionType);

                }
            }
        });
        slideCounter++;
    });

}


/**
 * Generate Results Presentation: This function adds an image of a given molecule based on its smile to the slide, shrinks and positions it correctly. It uses the FAST API for generating the png from the Smiles string.
 * @param {Object} presentation Presentation object, i.e. the results presentation.
 * @param {Array} slides collection of all the slide objects in the presentation in the form of an array.
 * @param {Object} reactionComponents Dictionary containing all the informations on the different starting materials and (side)products.
 * @param {String} key Component name, key of the reactionComponents dictionary. 
 * @param {Number} colCount number of the component table column at which the new component will reside.
 * @param {Number} xShift x-axis shift in pixels accounting for the space needed for the reaction equation (will be 0 for starting materials and 55 for (side)products).
 * @param {Number} slideNumber 0-based count of the slide number the new image should reside at.
 * @param {Number} numberOfExistingStructuresOnSlide number of structures already present on the slide, basically a running counter.
 * @return {Number} left edge of the added image, used in case of the product to determine the x-position of the reaction equation.
 */
function addImageToSlide(presentation, slides, reactionComponents, key, colCount, xShift, slideNumber, numberOfExistingStructuresOnSlide = 0) {
    try {
        var image = slides[slideNumber].insertImage(LINKtOfASTaPI + '/smiles-to-image/' + FASTaPIkEY + "/" + encodeURIComponent(reactionComponents[key].Smiles) + "?width=400&height=300&format=png");  //Google Cloud Run-based Python FastAPI deployed on the PTD Google Cloud Project PRJ-RSC-PTD-DATALAK-SB-001 as roslfastapi
        image.setDescription(reactionComponents[key].InfoType); // set the alternative text of the picture so that it can be identified later when adjusting the position of the reaction arrow/conditions group. 

        image.scaleWidth(((720 - 55) / (Object.keys(reactionComponents).length + numberOfExistingStructuresOnSlide) - 5) / 300);// width of the presentation is 720, extra space for reaction conditions is 55 and the spacing between two compounds is 5, 300 is the width of the picture when 400 pixels width resolution is selected
        image.scaleHeight((
            (720 - 55) /
            (Object.keys(reactionComponents).length + numberOfExistingStructuresOnSlide) - 5) / 300);
        var imgWidth = image.getWidth();
        var imgHeight = image.getHeight();
        var pageHeight = presentation.getPageHeight();

        var newX = 20 + xShift + (colCount / 2) * ((720 - 55) / (Object.keys(reactionComponents).length + numberOfExistingStructuresOnSlide));    //Start 20 from the left and then add the available space for one picture ((720-55)/Object.keys(reactionComponents).length) multiplied with the number of the compound at hand ((colCount)/2) 

        var newY = pageHeight / 7.0;
        //console.log(key, " New Image position and dimensions: ", newX, imgWidth, imgHeight)
        image.setLeft(newX).setTop(newY);
        if (reactionComponents[key].InfoType == "Product") {
            var groups = slides[slideNumber].getGroups();
            if (groups.length > 0) {
                groups[0].setLeft(image.getLeft() - groups[0].getWidth() + 30);
            }
        }
        image.sendToBack();
        return image.getLeft();//used in case of the product to determine the x-position of the reaction equation. 
    } catch (error) {
        console.log(key, error);
    }
}


/**
 * Generate Results Presentation: this function is needed to re-size and -position the pictures of reaction components when adding additional side products. All current structures need to be smaller and shifted to the left to make space for the new side products. For that, the idea is to go through all images of a given slide and iterate through all that are within a certain distance from the top (high enough not to interfere with screenshots added by the user.). 
 * At the point when this function is called it is already established that the given slide is either the overall conclusion slide or the plate conclusion slide
 * @param {Object} image image object, the image that needs to be re-sized / -positioned.
 * @param {Number} numberOfStructuresNow number of structures already present before addition of the new side product.
 * @param {Number} numberOfNewSideProducts number of new side products to be added.
 * @param {Number} relevantImageCounter number of structures already present on the slide, basically a running counter.
 * @return {String} Description of the image, contains information on whether the image in question is a starting material or (side)product. If product, then it's left edge is used to set the right edge of the reaction equation.
 */
function amendStructureScalingAndPosition(image, numberOfStructuresNow, numberOfNewSideProducts, relevantImageCounter) {
    var oldWidth = image.getWidth();
    console.log("ScaleFactor: " + (((720 - 55) / (numberOfStructuresNow + numberOfNewSideProducts) - 5) / ((720 - 55) / numberOfStructuresNow - 5)));
    image.scaleWidth(((720 - 55) / (numberOfStructuresNow + numberOfNewSideProducts) - 5) / ((720 - 55) / numberOfStructuresNow - 5));   //new scaling factor divided by the old one
    image.scaleHeight(((720 - 55) / (numberOfStructuresNow + numberOfNewSideProducts) - 5) / ((720 - 55) / numberOfStructuresNow - 5));   //new scaling factor divided by the old one
    console.log("left x after scaling: " + image.getLeft());
    var newX = image.getLeft() - (numberOfStructuresNow - relevantImageCounter) / (numberOfStructuresNow + numberOfNewSideProducts) * (oldWidth) * numberOfNewSideProducts;  //subtract from the current left edge x-position the amount that the picture needs to be shifted to the left, important: loop starts with the rightmost picture which is why it's (numberOfStructuresNow - relevantImageCounter), using oldwidth as the basis for calculating how much space needs to be saved is too much (since only newwidth*numberofnewsideproducts needs to be saved, but it looks good on the slide setting the new side products slightly apart and making it clear that they were added later. )
    image.setLeft(newX);
    image.sendToBack();
    return image.getDescription();
}

// This function takes the presentation template found in the "Data" subfolder, makes a copy of it and fills it with information about the reaction and the current plate
/**
 * Generate Results Presentation: this function is needed to re-size and -position the pictures of reaction components when adding additional side products. All current structures need to be smaller and shifted to the left to make space for the new side products. For that, the idea is to go through all images of a given slide and iterate through all that are within a certain distance from the top (high enough not to interfere with screenshots added by the user.). 
 * At the point when this function is called it is already established that the given slide is either the overall conclusion slide or the plate conclusion slide
 * @param {Object} reactionComponents dictionary containing all information on the reaction components with the component name as key.
 * @param {Array} data Content of the FileGenerator sheet, only handed over when the function is called by the SavePlate function, empty when called when submitting a new request.
 * @param {String} cheatSheetFileUrl url of the cheatsheet, to be put as link on the overview slide of each plate.
 * @param {String} fileIdPresentation fileid of the presentation in case it already exists.
 * @param {Array} formattedComponentsInColumns array containing the formatted strings of all components in the COLUMNS of the plate in question.
 * @param {Array} formattedComponentsInRows array containing the formatted strings of all components in the ROWS of the plate in question.
 * @return {String} URL leading to the presentation.
 */
function fillPresentationTemplate(reactionComponents = {}, data = [], cheatSheetFileUrl = "https://docs.google.com/spreadsheets/d/1k7px3RY3axXPqwf-gY3iMnTNSiWDTEE-7rnv154HJ-k/edit#gid=1470374422", fileIdPresentation = "", formattedComponentsInColumns = [], formattedComponentsInRows = []) {

    //debugging part:
    //var fileGeneratorSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FileGenerator");
    //var data = fileGeneratorSheet.getRange("R2:AP133").getValues();
    var destinationFolderId = globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["PRESENTATIONfOLDERiD"]; //createFolder(globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["PROJECTdATAfOLDERiD"], data[0][0] + "_" + data[5][5] + "_" + data[6][5] + "_" + data[7][5]) //create a new folder with the ELN-ID as name on gDrive under HTS Docs > Robot Input Files unless it exists already 
    var commonComponents = ""; // will contain all information on components present in all vials 
    var combinedData = {}; // stores information on starting materials and products in a dictionary that is constructed like reactionComponents from Submit Request.
    var reactionEquation = []; //will hold the table below the reaction equation
    var colCount = 0; //number of columns in the table below the reaction equation
    if (data.length > 0) { // if a plate is read into the presentation
        for (var row = 0; row < 5; row++) {
            if (data[row][1] == "") continue;
            combinedData[data[row][3]] = {
                "eqScale": data[row][12] + " eq/ " + parseFloat(data[row][13]).toFixed(1) + " mg",
                "MW": parseFloat(data[row][10]).toFixed(2),
                "dosedAsReactionType": data[row][16],
                "InfoType": data[row][1],
                "Smiles": data[row + 122][6] //needed to get structures, neccessary Smiles are at the bottom of the table. 
            };
        }

        for (row = 5; row < 9; row++) {
            if (data[row][1] == "") continue;
            if (commonComponents.length > 0) commonComponents += ", ";
            commonComponents += generateFormattedPlateIngredient(data[row]);
        }

        for (row = 122; row < 132; row++) {
            if (data[row][1] == "") continue;
            combinedData[data[row][1]] = {
                "eqScale": "-",
                "MW": parseFloat(data[row][3]).toFixed(2),
                "dosedAsReactionType": data[row][2],
                "InfoType": "Side Product",
                "Smiles": data[row][4]
            };
        }
    }
    if (fileIdPresentation == "") {
        fileIdPresentation = createPresentation(globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["PRESENTATIONfOLDERiD"], reactionComponents["otherInfo"].ElnId + " Results");
    }
    var presentation = SlidesApp.openById(fileIdPresentation);
    var presentationLink = presentation.getUrl();

    //get the individual slides to manipulate them
    var slides = presentation.getSlides();
    var tables = slides[1].getTables();
    var indefiniteArticle = "";
    if (cheatSheetFileUrl == "none") { //This is the case if a new experiment is registered, as the cheatsheet is only created when a plate is registered  

        // Modify the placeholder text on the title slide
        var reactionTypeFirstLetter = reactionComponents["otherInfo"].ReactionType[0];
        if (reactionTypeFirstLetter.match(/[AEOUIaeiou]/g) == null) { //If the reaction type starts with a vowel, use "An" otherwise "A" as indefinite article
            indefiniteArticle = ": A ";
        }
        else { indefiniteArticle = ": An "; }
        slides[0].replaceAllText("{{Presentation title}}", reactionComponents["otherInfo"].ProjectName + ": " + reactionComponents["otherInfo"].StepName);
        slides[0].replaceAllText("{{Presentation subtitle}}", reactionComponents["otherInfo"].ElnId + indefiniteArticle + reactionComponents["otherInfo"].ReactionType + "-Screen for " + reactionComponents["otherInfo"].Customer);
        slides[0].replaceAllText("{{DATE}}", Utilities.formatDate(new Date(), "CET", 'MMMM dd, yyyy'));

        // Set the links to Spotfire and the Cheatsheet on the Overall Conclusion slide

        slides[1].getShapes().forEach(function (shape) { //adapted from: https://stackoverflow.com/questions/63551731/google-app-script-to-create-link-in-google-slides
            var text = shape.getText();

            var m = text.replaceAllText('{{Link to Spotfire}}', 'Spotfire Link');
            if (m > 0) text.getTextStyle().setLinkUrl(globalVariableDict[SpreadsheetApp.getActiveSpreadsheet().getId()]["SPOTFIRElINK"] + reactionComponents["otherInfo"].ElnId + "%22;");


            //var n = text.replaceAllText('{{Link to Cheatsheet}}','CheatSheet');
            //if (n>0) text.getTextStyle().setLinkUrl(cheatSheetFileUrl);

            //Insert link to folder in Windows Explorer and/or chemical drawing containing the reaction equation

        }
        );

        //Append columns to an existing table  containing Component Name (as links to Windream), retention time in LCMS and molecular weight for starting materials and products 

        tables = slides[1].getTables();
        reactionEquation = tables[0];
        colCount = reactionEquation.getNumColumns();
        // go through all reaction components and write Component name, Smiles, MW, Batch ID and Producer
        for (var key in reactionComponents) {
            if (key == "otherInfo") { continue; }
            switch (reactionComponents[key].InfoType) { //The table is filled with different data depending on the nature of the component in question
                case "Starting Material":

                    // add two columns except for the first starting material, the first column containing a "+" sign and the second one the data
                    if (colCount > 1) {
                        reactionEquation.appendColumn();
                        colCount++;
                        reactionEquation.getColumn(colCount - 1).getCell(1).getText().setText("+");
                    }
                    addImageToSlide(presentation, slides, reactionComponents, key, colCount, 0, 1, 0);
                    reactionEquation.appendColumn();
                    colCount++;
                    reactionEquation.getColumn(colCount - 1).getCell(0).getText().setText(reactionComponents[key].ComponentName);
                    reactionEquation.getColumn(colCount - 1).getCell(1).getText().setText(
                        getRetentionTimePrediction(reactionComponents[key].Smiles)
                    ).getTextStyle().setForegroundColor('#c40000');
                    reactionEquation.getColumn(colCount - 1).getCell(2).getText().setText(parseFloat(reactionComponents[key].MW).toFixed(2));
                    reactionEquation.getColumn(colCount - 1).getCell(3).getText().setText(reactionComponents[key].BatchId + " \\ " + reactionComponents[key].Producer);



                    break;
                case "Product":
                    //add two columns, the first one only containing an arrow, the second one the product data
                    reactionEquation.appendColumn();
                    colCount++;
                    addImageToSlide(presentation, slides, reactionComponents, key, colCount, 55, 1, 0);


                    reactionEquation.getColumn(colCount - 1).getCell(1).getText().setText(String.fromCharCode(8680));  // right Arrow character code
                    reactionEquation.appendColumn();
                    colCount++;
                    reactionEquation.getColumn(colCount - 1).getCell(0).getText().setText(reactionComponents[key].ComponentName);
                    reactionEquation.getColumn(colCount - 1).getCell(1).getText().setText(
                        getRetentionTimePrediction(reactionComponents[key].Smiles)
                    ).getTextStyle().setForegroundColor('#c40000');
                    reactionEquation.getColumn(colCount - 1).getCell(2).getText().setText(parseFloat(reactionComponents[key].MW).toFixed(2));
                    reactionEquation.getColumn(colCount - 1).getCell(3).getText().setText(reactionComponents[key].BatchId + " \\ " + reactionComponents[key].Producer);



                    break;
                case "Side Product":
                    // add two columns, the first one containing a "+" sign and the second one the side product data
                    reactionEquation.appendColumn();
                    colCount++;
                    addImageToSlide(presentation, slides, reactionComponents, key, colCount, 55, 1, 0);

                    reactionEquation.getColumn(colCount - 1).getCell(1).getText().setText("+");
                    reactionEquation.appendColumn();
                    colCount++;
                    reactionEquation.getColumn(colCount - 1).getCell(0).getText().setText(reactionComponents[key].ComponentName);
                    reactionEquation.getColumn(colCount - 1).getCell(1).getText().setText(
                        getRetentionTimePrediction(reactionComponents[key].Smiles)
                    ).getTextStyle().setForegroundColor('#c40000');
                    reactionEquation.getColumn(colCount - 1).getCell(2).getText().setText(parseFloat(reactionComponents[key].MW).toFixed(2));
                    reactionEquation.getColumn(colCount - 1).getCell(3).getText().setText(reactionComponents[key].SideReactionType);



                    break;
            }
        }
    } else { // This is the case when the function is called during plate creation (savePlate() function from FileGenerator.gs) in which data contains the stuff on the FileGenerator Sheet

        //duplicate the template slides at the back, adapt generic phrases to the plate at hand and move the slides below slide 2 (Overall Conclusion). This way, the newest plate is always on top. 
        for (var i = 1; i < 4; i++) {
            slides[slides.length - i].duplicate();
            slides[slides.length - i].replaceAllText("Plate x", "Plate " + data[1][0]);
            slides[slides.length - i].replaceAllText("Placeholder", "Plate Conclusion Slide"); // Side products that are registered afterwards are directed to the plate conclusion slide using this invisible textbox in the top right corner. Changing the name avoids new side products being added to the template as well. 


            slides[slides.length - i].move(2);
        }
        slides = presentation.getSlides();
        // insert Plate Goal / Setup
        slides[2].replaceAllText("{{Plate Goal}}", data[9][0]);   //Cell R11 in File Generator  

        slides[2].replaceAllText("{{Plate Goal/Setup}}", data[10][0]);    // Cell R12

        slides[2].replaceAllText("{{All Vials contain:}}", "All Vials contain: " + commonComponents);


        slides[2].getShapes().forEach(function (shape) { //adapted from: https://stackoverflow.com/questions/63551731/google-app-script-to-create-link-in-google-slides
            var text = shape.getText();
            var n = text.replaceAllText('{{Link to Cheatsheet}}', 'CheatSheet');
            if (n > 0) text.getTextStyle().setLinkUrl(cheatSheetFileUrl);


            //Insert link to folder in Windows Explorer and/or chemical drawing containing the reaction equation


        });

        //Add columns to the plate conclusion table

        tables = slides[2].getTables();
        reactionEquation = tables[0];
        colCount = reactionEquation.getNumColumns();
        for (let key in combinedData) {

            switch (combinedData[key].InfoType) { //The table is filled with different data depending on the nature of the component in question
                case "Starting Material":
                    // add two columns except for the first starting material, the first column containing a "+" sign and the second one the data
                    if (colCount > 1) {
                        reactionEquation.appendColumn();
                        colCount++;
                        reactionEquation.getColumn(colCount - 1).getCell(1).getText().setText("+");
                    }
                    addImageToSlide(presentation, slides, combinedData, key, colCount, 0, 2, 1);   //number of components on slide (last input) is set to 1 to compensate for the lack of otherInfo key present in reactioncomponents but not in combined data. 

                    reactionEquation.appendColumn();
                    colCount++;
                    reactionEquation.getColumn(colCount - 1).getCell(0).getText().setText(combinedData[key].eqScale);
                    reactionEquation.getColumn(colCount - 1).getCell(1).getText().setText(key);
                    reactionEquation.getColumn(colCount - 1).getCell(2).getText().setText(parseFloat(combinedData[key].MW).toFixed(2));
                    reactionEquation.getColumn(colCount - 1).getCell(3).getText().setText(combinedData[key].dosedAsReactionType);

                    break;
                case "Product":

                    //add two columns, the first one only containing an arrow, the second one the product data
                    reactionEquation.appendColumn();
                    colCount++;
                    addImageToSlide(presentation, slides, combinedData, key, colCount, 55, 2, 1);
                    reactionEquation.getColumn(colCount - 1).getCell(1).getText().setText(String.fromCharCode(8680));  // right Arrow character code
                    reactionEquation.appendColumn();
                    colCount++;
                    reactionEquation.getColumn(colCount - 1).getCell(0).getText().setText(combinedData[key].eqScale);
                    reactionEquation.getColumn(colCount - 1).getCell(1).getText().setText(key);
                    reactionEquation.getColumn(colCount - 1).getCell(2).getText().setText(parseFloat(combinedData[key].MW).toFixed(2));
                    reactionEquation.getColumn(colCount - 1).getCell(3).getText().setText("-"); // product isn't dosed
                    break;
                case "Side Product":

                    // add two columns, the first one containing a "+" sign and the second one the side product data
                    reactionEquation.appendColumn();
                    colCount++;
                    addImageToSlide(presentation, slides, combinedData, key, colCount, 55, 2, 1);

                    reactionEquation.getColumn(colCount - 1).getCell(1).getText().setText("+");
                    reactionEquation.appendColumn();
                    colCount++;
                    reactionEquation.getColumn(colCount - 1).getCell(0).getText().setText(combinedData[key].dosedAsReactionType);
                    reactionEquation.getColumn(colCount - 1).getCell(1).getText().setText(key);
                    reactionEquation.getColumn(colCount - 1).getCell(2).getText().setText(parseFloat(combinedData[key].MW).toFixed(2));
                    reactionEquation.getColumn(colCount - 1).getCell(3).getText().setText(combinedData[key].eqScale);
                    break;
            }
        }

        // Amend tables on the slides for qualitative and quantitative analysis (slides[3] and [4])

        var tablesOnQualSlide = slides[3].getTables();
        var qualRowLabelsTable = tablesOnQualSlide[1];
        var qualColumnLabelsTable = tablesOnQualSlide[0];
        var tablesOnQuantSlide = slides[4].getTables();
        var quantRowLabelsTable = tablesOnQuantSlide[1];
        var quantColumnLabelsTable = tablesOnQuantSlide[0];


        for (let row = 0; row < formattedComponentsInRows.length; row++) {
            //write the content of formattedComponentsInRows to the table and append rows starting from row 2
            if (row > 0) {
                qualRowLabelsTable.appendRow();
                quantRowLabelsTable.appendRow();
            }
            for (let column = 0; column < formattedComponentsInRows[0].length; column++) {
                qualRowLabelsTable.getRow(row).getCell(column).getText().setText(formattedComponentsInRows[row][2 - column]);
                quantRowLabelsTable.getRow(row).getCell(column).getText().setText(formattedComponentsInRows[row][2 - column]);
            }
        }

        for (let row = 0; row < formattedComponentsInColumns.length; row++) {
            //write the content of formattedComponentsInColumns to the other table and append rows starting from row 2
            if (row > 0) {
                qualColumnLabelsTable.appendColumn();
                quantColumnLabelsTable.appendColumn();
            }
            for (let column = 0; column < formattedComponentsInColumns[0].length; column++) {
                qualColumnLabelsTable.getColumn(row).getCell(column).getText().setText(formattedComponentsInColumns[row][column]);
                quantColumnLabelsTable.getColumn(row).getCell(column).getText().setText(formattedComponentsInColumns[row][column]);
            }
        }

    }
    return presentationLink;
}

/**
 * Generate Results Presentation: this function calls the retention time prediction API prepared by Torsten Schindler and Pascal Zimmerli and returns the predicted retention time for a given smiles string.
 * @param {String} smiles smiles string of the compound in question.
 * @return {Number} predicted retention time in minutes rounded to two decimal places
 */
function getRetentionTimePrediction(smiles = 'COC1=CC(N(C2=CC(OC)=C(OC)C(OC)=C2)C3=NC=CO3)=CC(OC)=C1OC') {
    /*try {
        var url = "<<link to a retention time prediction API you might have>>/" + encodeURIComponent(smiles);
        var response = UrlFetchApp.fetch(url, { 'muteHttpExceptions': true });
        var json = response.getContentText();
        var data = JSON.parse(json);
        Utilities.sleep(1700); //Wait for 1.7 sec to prevent overloading the API: longer or shorter leads to problems
        return parseFloat(data.rt_predicted).toFixed(2);
    } catch (err) {
        return err;
    }*/
    return "no Retention Time predicted";

}