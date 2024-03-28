/*jshint sub:true*/

/**
 * heartbeat: sends a message to a webhook in a Google chatspace
 *  @param {String} message - message to be sent
 *  @param {String} webhookLink - URL of the webhook in the chatspace to send the message to
 *  @returns {Object} Object all the info about the last dosing found in the text file.  
 */
function sendChatMessage(message = "This is a test.", webhookLink) {
    const payload = JSON.stringify({ text: message });
    const options = {
        method: 'POST',
        contentType: 'application/json',
        payload: payload,
    };
    UrlFetchApp.fetch(webhookLink, options);
}


/**
 * heartbeat: triggered every minute to look for new dosings and dosing files.
 */
function identifyLastDosingSequence() {

    const sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();

    const headEmptyQuantosHeartbeat = globalVariableDict[sheetId]["HEADeMPTYqUANTOShEARTBEAT"];
    const issueWebhookQuantosHeartbeat = globalVariableDict[sheetId]["ISSUEwEBHOOKqUANTOShEARTBEAT"];
    const progressWebhook = globalVariableDict[sheetId]["PROGRESSwEBHOOK"];
    const successWebhookQuantosHeartbeat = globalVariableDict[sheetId]["SUCCESSwEBHOOKqUANTOShEARTBEAT"];

    const now = new Date();

    const cutoffTime = new Date(now.getTime() - 7 * 1000 * 60);  // If the last change of the file is less than 4 minutes old, we assume that the dosing is still ongoing
    const timeOut = new Date(now.getTime() - 8.03 * 1000 * 60); // Time before which a change to the dosing file is regarded as stale to avoid double messaging.
    const stopConsiderTime = new Date(now.getTime() - 9.99 * 1000 * 60); // Time before which a change to the dosing file is regarded as stale to avoid double messaging.



    directoryArray = getFileAndFolderIds();  //get all text files in dosing results ordered by newest modified date on top

    const fileName = directoryArray[0][0];
    const fileNameCore = fileName.substring(0, fileName.length - 27); //everything before _Runlog_
    const fileId = directoryArray[0][1];
    const lastModified = directoryArray[0][3];

    if (lastModified < stopConsiderTime) return; //stop, if the last dosing is more than 9.99 min ago


    var [lastDosingInfo, dosingBeforeLastDosingInfo] = parseDosingResults(directoryArray[0][1]);
    console.log("latest dosing file: " + directoryArray[0]);
    console.log(lastDosingInfo.lastDosingAsText);

    const scriptCache = CacheService.getScriptCache();
    const latestDosingFromCache = scriptCache.get('latestDosing'); // get the current status from cache

    console.log("Cache: " + latestDosingFromCache);
    const latestDosingFromCacheDict = JSON.parse(latestDosingFromCache);

    if (lastDosingInfo.lastDosingStatus) {
        var cacheContent =
        {
            fileName: fileName,
            fileNameCore: fileNameCore,
            fileId: fileId,
            lastModified: lastModified,
            lastDosingHead: lastDosingInfo.lastDosingHead,
            lastVial: lastDosingInfo.lastVial,
            lastDosingStatus: lastDosingInfo.lastDosingStatus
        };
    } else {
        var cacheContent =
        {
            fileName: fileName,
            fileNameCore: fileNameCore,
            fileId: fileId,
            lastModified: lastModified,
            lastDosingHead: lastDosingInfo.lastDosingHead,
            lastVial: lastDosingInfo.lastVial
        };
    }


    scriptCache.put('latestDosing', JSON.stringify(cacheContent), 600);

    if (JSON.stringify(cacheContent) == latestDosingFromCache) {
        if (lastModified > timeOut && lastModified < cutoffTime) { // no dosing is ongoing and between 6 and 7 min have passed since the last one (avoids double reporting) , need to differentiate when the last dosing stopped and whether it happened at the end or beforehand

            if (lastDosingInfo.lastDosingStatus) {
                if (lastDosingInfo.lastDosingStatus[3] == lastDosingInfo.lastDosingStatus[2]) { // True, if the last dosing noted in the file was the last dosing of the sequence and if there's no newer dosing file.  

                    sendChatMessage(lastDosingInfo.lastPlateId + " finished regularly and no new dosing sequence started. Thank you for including it in your prayers." + " " + getDosingFileFromRunlog(fileName), successWebhookQuantosHeartbeat);
                    const suspiciousDosings = generateDosingResultsSummary(directoryArray[1][1]);
                    if (suspiciousDosings.length > 1) {
                        sendChatMessage(suspiciousDosings, progressWebhook);
                    }
                    console.log(lastDosingInfo.lastPlateId + " finished regularly. Thank you for including it in your prayers!");
                } else {

                    if (lastDosingInfo.lastDosingStatus[1] == lastDosingInfo.lastDosingStatus[0]) { // True, if it's the last dosing of this dosing head..  
                        sendChatMessage("ALERT: " + lastDosingInfo.lastPlateId + " finished, but no new dosings detected after " + lastDosingInfo.lastDosingStatus[2] + " out of " + lastDosingInfo.lastDosingStatus[3] + ". It was the last dosing of " + lastDosingInfo.lastDosingHead + " in this sequence. Here the info on the last dosing:\r\n" + lastDosingInfo.lastDosingAsText + " " + getDosingFileFromRunlog(fileName), issueWebhookQuantosHeartbeat);
                        console.log("ALERT: " + lastDosingInfo.lastPlateId + " finished, but no new dosings detected after " + lastDosingInfo.lastDosingStatus[2] + " out of " + lastDosingInfo.lastDosingStatus[3] + ". It was the last dosing of " + lastDosingInfo.lastDosingHead + " in this sequence. Here the info on the last dosing:\r\n" + lastDosingInfo.lastDosingAsText + " " + getDosingFileFromRunlog(fileName));
                    } else { //It stopped prematurely before all dosings of the current dosing head were finished.
                        sendChatMessage("ALERT: " + lastDosingInfo.lastPlateId + " finished, but no new dosings detected after " + lastDosingInfo.lastDosingStatus[2] + " out of " + lastDosingInfo.lastDosingStatus[3] + ". It stopped after " + lastDosingInfo.lastDosingStatus[0] + " out of " + lastDosingInfo.lastDosingStatus[1] + " dosings of " + lastDosingInfo.lastDosingHead + " in this sequence. Here the info on the last dosing:\r\n" + lastDosingInfo.lastDosingAsText + " " + getDosingFileFromRunlog(fileName), issueWebhookQuantosHeartbeat);
                        console.log("ALERT: " + lastDosingInfo.lastPlateId + " finished, but no new dosings detected after " + lastDosingInfo.lastDosingStatus[2] + " out of " + lastDosingInfo.lastDosingStatus[3] + ". It stopped after " + lastDosingInfo.lastDosingStatus[0] + " out of " + lastDosingInfo.lastDosingStatus[1] + " dosings of " + lastDosingInfo.lastDosingHead + " in this sequence. Here the info on the last dosing:\r\n" + lastDosingInfo.lastDosingAsText + " " + getDosingFileFromRunlog(fileName));
                    }
                }
            } else {
                sendChatMessage(lastDosingInfo.lastPlateId + " finished. Here the info on the last dosing:\r\n" + lastDosingInfo.lastDosingAsText, progressWebhook);
                console.log(lastDosingInfo.lastPlateId + " finished. Here the info on the last dosing:\r\n" + lastDosingInfo.lastDosingAsText);
            }
        }
        return;
    } //The dosing is either still ongoing or finished less than 10 min ago. But the results must have been reported already. 


    // Deal with newly started schedules, i.e. when there are no dosings more recent than 10 min and the cache is empty.



    if (!latestDosingFromCache) {
        sendChatMessage("Quantos has just started work on " + fileNameCore + " as the first dosing sequence of a maybe larger schedule. May the god of robots have mercy with all of them. Here the dosing results file as it builds: " + getDosingFileFromRunlog(fileName), progressWebhook);
        console.log("Quantos has just started work on " + fileNameCore + " as the first dosing sequence of a maybe larger schedule. May the god of robots have mercy with all of them.");
        return;
    }


    // This point is only reached, if the latest dosing happened less than 10 min ago and if it wasn't reported yet. 

    if (fileId != latestDosingFromCacheDict.fileId) { // A new dosing has started after another one had finished within the last 10 min.

        if (latestDosingFromCacheDict.lastDosingStatus) {  // If the previous dosing included a status marker
            if (latestDosingFromCacheDict.lastDosingStatus[3] == latestDosingFromCacheDict.lastDosingStatus[2]) { // True, if the last dosing noted in the previous dosing file was the last dosing of the sequence and if there's no newer dosing file.  
                sendChatMessage(latestDosingFromCacheDict.fileNameCore + " finished regularly and " + fileNameCore + " started. Thank you for including both in your prayers." + " " + getDosingFileFromRunlog(fileName), successWebhookQuantosHeartbeat);
                const suspiciousDosings = generateDosingResultsSummary(directoryArray[1][1]);
                if (suspiciousDosings.length > 1) {
                    sendChatMessage(suspiciousDosings, progressWebhook);
                }
                console.log(latestDosingFromCacheDict.fileNameCore + " finished regularly and " + fileNameCore + " started. Thank you for including both in your prayers." + " " + getDosingFileFromRunlog(latestDosingFromCacheDict.fileName));
            } else {
                if (latestDosingFromCacheDict.lastDosingStatus[1] == latestDosingFromCacheDict.lastDosingStatus[0]) { // True, if it's the last dosing of this dosing head..  
                    sendChatMessage("ALERT: " + latestDosingFromCacheDict.fileNameCore + " finished, but no new dosings detected after " + latestDosingFromCacheDict.lastDosingStatus[2] + " out of " + latestDosingFromCacheDict.lastDosingStatus[3] + ". It was the last dosing of " + latestDosingFromCacheDict.lastDosingHead + " in this sequence." + " " + getDosingFileFromRunlog(latestDosingFromCacheDict.fileName), issueWebhookQuantosHeartbeat);
                    console.log("ALERT: " + latestDosingFromCacheDict.fileNameCore + " finished, but no new dosings detected after " + latestDosingFromCacheDict.lastDosingStatus[2] + " out of " + latestDosingFromCacheDict.lastDosingStatus[3] + ". It was the last dosing of " + latestDosingFromCacheDict.lastDosingHead + " in this sequence." + " " + getDosingFileFromRunlog(latestDosingFromCacheDict.fileName));
                } else { //It stopped prematurely before all dosings of the current dosing head were finished.
                    sendChatMessage("ALERT: " + latestDosingFromCacheDict.fileNameCore + " finished, but no new dosings detected after " + latestDosingFromCacheDict.lastDosingStatus[2] + " out of " + latestDosingFromCacheDict.lastDosingStatus[3] + ". It stopped after " + latestDosingFromCacheDict.lastDosingStatus[0] + " out of " + latestDosingFromCacheDict.lastDosingStatus[1] + " dosings of " + latestDosingFromCacheDict.lastDosingHead + " in this sequence." + " " + getDosingFileFromRunlog(latestDosingFromCacheDict.fileName), issueWebhookQuantosHeartbeat);
                    console.log("ALERT: " + latestDosingFromCacheDict.fileNameCore + " finished, but no new dosings detected after " + latestDosingFromCacheDict.lastDosingStatus[2] + " out of " + latestDosingFromCacheDict.lastDosingStatus[3] + ". It stopped after " + latestDosingFromCacheDict.lastDosingStatus[0] + " out of " + latestDosingFromCacheDict.lastDosingStatus[1] + " dosings of " + latestDosingFromCacheDict.lastDosingHead + " in this sequence." + " " + getDosingFileFromRunlog(latestDosingFromCacheDict.fileName));
                }
            }
        } else {
            sendChatMessage(latestDosingFromCacheDict.fileNameCore + " finished. The file didn't contain statusdata, thus that's all we know." + " " + getDosingFileFromRunlog(latestDosingFromCacheDict.fileName), progressWebhook);
            console.log(latestDosingFromCacheDict.fileNameCore + " finished. The file didn't contain statusdata, thus that's all we know." + " " + getDosingFileFromRunlog(latestDosingFromCacheDict.fileName));

            const suspiciousDosings = generateDosingResultsSummary(directoryArray[1][1]);
            if (suspiciousDosings.length > 1) {
                sendChatMessage(suspiciousDosings, progressWebhook);
            }
        }
        sendChatMessage(fileNameCore + " has just commenced dosing. May the god of robots have mercy with it. You'll see the dosing progress reflected with a few mins delay in this dosing file: " + getDosingFileFromRunlog(fileName), progressWebhook);
        console.log(fileNameCore + " has just commenced dosing. May the god of robots have mercy with it.");
        return;
    } else {   // the file id hasn't changed, i.e. no new dosing sequence has started

        if (lastDosingInfo.lastDosingStatus && lastDosingInfo.lastDosingStatus[3] - lastDosingInfo.lastDosingStatus[2] == 2) { //Report the number of dosings left, if less than 3 dosings are left in this sequence. 
            sendChatMessage("Only TWO dosings  left in " + fileNameCore + ".", progressWebhook);
            console.log("Fewer than " + (lastDosingInfo.lastDosingStatus[3] - lastDosingInfo.lastDosingStatus[2] + 1) + " dosings  left in " + lastDosingInfo.lastPlateId);
        }

        //Report issue in Quantos Heartbeat, if the current dosing and the one before that undershot the target weight by more than 90%. 

        if (parseFloat(lastDosingInfo.lastDifferencePct.split(" ")[0]) < -90 && parseFloat(dosingBeforeLastDosingInfo.lastDifferencePct.split(" ")[0]) < -90) {
            sendChatMessage("The actual weight of " + lastDosingInfo.lastDosingHead + " in " + lastDosingInfo.lastVial.split(" - ")[1] + " is " + lastDosingInfo.lastDifferencePct.split(" ")[0] + "% below the " + lastDosingInfo.lastTargetDose + " target.", headEmptyQuantosHeartbeat);
            console.log("The actual weight of " + lastDosingInfo.lastDosingHead + " in " + lastDosingInfo.lastVial.split(" - ")[1] + " is " + lastDosingInfo.lastDifferencePct.split(" ")[0] + "% below the " + lastDosingInfo.lastTargetDose + " target.");
        }
    }
}

/**
 * heartbeat: Looks in the dosing results folder returns 
 *  @param {String/null} fileId - id of the textfile to be parsed
 *  @returns {String/null} String containing all the problematic dosings as a table.  
 */
function generateDosingResultsSummary(fileId = "1EpA_ZMDQKukdTvfnNMnayyiKHUid-EcM") {


    const fileContent = DriveApp.getFileById(fileId).getAs("text/plain").getDataAsString();
    const contentArray = fileContent.split('\r\n----------------------------------------------------------\r\n');  //separates the dosings
    var individaulDosingsAsArray = [];
    for (let row = 0; row < contentArray.length; row++) {
        individaulDosingsAsArray.push(contentArray[row].split('\r\n'));  //split each sample into individual lines
    }
    var problematicDosings = ["Head ID      Vial   Target Deviation         Error                 List of dosings with problems."];
    var deviationFromTarget = 0;
    var dosingHeadId = "";
    var vial = "";
    var errorMessage = "";
    var targetValue = 0;
    for (let row = 0; row < individaulDosingsAsArray.length - 2; row++) {
        deviationFromTarget = individaulDosingsAsArray[row][9].split('Diff%:           ')[1].split(" ")[0]; // e.g. Diff%:           -2.58 %
        dosingHeadId = individaulDosingsAsArray[row][2].split('Substance:    ')[1];
        vial = individaulDosingsAsArray[row][3].split('Vial:              ')[1].split(" - ")[1];   //e.g. Vial:              Tray5:10 - A10
        targetValue = individaulDosingsAsArray[row][5].split('Target Dose: ')[1];
        if (individaulDosingsAsArray[row][11].length > 5) {
            errorMessage = individaulDosingsAsArray[row][11].split('Error:            ')[1]; //e.g. Error:            Dosing Status: PowderflowError - Sample Data Error: NotAllowedAtTheMoment}
        } else {
            errorMessage = "";
        }

        if (errorMessage.length > 0 ||
            (parseFloat(deviationFromTarget) < -30 && parseFloat(targetValue.split(" ")[0]) > 1) ||
            (parseFloat(deviationFromTarget) < -50 && parseFloat(targetValue.split(" ")[0]) < 1)) { // if there's an error or if the dosing is too low
            problematicDosings.push(dosingHeadId + "\t" + vial + "\t" + targetValue + "\t" + deviationFromTarget + "\t" + errorMessage);
        }

    }
    if (problematicDosings.length > 1) {
        return problematicDosings.join("\r\n");
    };
}

/**
 * heartbeat: Looks in the dosing results folder and checks on the progress for dosings o
 *  @param {String/null} fileId - id of the textfile to be parsed
 *  @returns {Object} Object all the info about the last dosing found in the text file.  
 */
function parseDosingResults(fileId = "1EpA_ZMDQKukdTvfnNMnayyiKHUid-EcM") {
    const fileContent = DriveApp.getFileById(fileId).getAs("text/plain").getDataAsString();
    const contentArray = fileContent.split('\r\n----------------------------------------------------------\r\n');  //separates the dosings
    const lastDosingAsArray = contentArray[contentArray.length - 2].split('\r\n');  //split the last sample into individual lines


    var lastDosingInfo = {};   //object that'll store the parsed data of the latest dosing


    lastDosingInfo.lastDosingAsArray = lastDosingAsArray;
    lastDosingInfo.lastDosingAsText = contentArray[contentArray.length - 2];

    lastDosingInfo.lastTimeStamp = lastDosingAsArray[0].split('\t----------------------------------------------------------')[0];
    lastDosingInfo.lastDosingHeadHotelPosition = lastDosingAsArray[1].split('Head:           Heads:')[1];
    lastDosingInfo.lastDosingHead = lastDosingAsArray[2].split('Substance:    ')[1];
    lastDosingInfo.lastVial = lastDosingAsArray[3].split('Vial:              ')[1];
    lastDosingInfo.lastSampleId = lastDosingAsArray[4].split('Sample ID:     ')[1];
    lastDosingInfo.sampleIdArray = lastDosingInfo.lastSampleId.split(" - ");
    lastDosingInfo.lastPlateId = lastDosingInfo.sampleIdArray[0];
    lastDosingInfo.iniVsCorr = lastDosingInfo.sampleIdArray[1];
    lastDosingInfo.lastPlateType = lastDosingInfo.sampleIdArray[2];
    if (lastDosingInfo.sampleIdArray.length > 3) lastDosingInfo.lastDosingStatus = lastDosingInfo.sampleIdArray[3].split('/');
    lastDosingInfo.lastTargetDose = lastDosingAsArray[5].split('Target Dose: ')[1];
    lastDosingInfo.lastActualWeight = lastDosingAsArray[7].split('Act. Weight:  ')[1];
    lastDosingInfo.lastDifferenceMg = lastDosingAsArray[8].split('Difference:    ')[1];
    lastDosingInfo.lastDifferencePct = lastDosingAsArray[9].split('Diff%:           ')[1];
    lastDosingInfo.lastValidity = lastDosingAsArray[10].split('Validity:         ')[1];


    if (contentArray.length - 2 > 0) { // if there is more than one dosing present in the file
        var dosingBeforeLastDosingInfo = {}; //object that'll store the parsed data of the dosing before the latest dosing
        const dosingBeforeLastDosingAsArray = contentArray[contentArray.length - 3].split('\r\n');  //split the last sample into individual lines


        dosingBeforeLastDosingInfo.dosingBeforeLastDosingAsArray = dosingBeforeLastDosingAsArray;
        dosingBeforeLastDosingInfo.lastDosingAsText = contentArray[contentArray.length - 3];

        dosingBeforeLastDosingInfo.lastTimeStamp = dosingBeforeLastDosingAsArray[0].split('\t----------------------------------------------------------')[0];
        dosingBeforeLastDosingInfo.lastDosingHeadHotelPosition = dosingBeforeLastDosingAsArray[1].split('Head:           Heads:')[1];
        dosingBeforeLastDosingInfo.lastDosingHead = dosingBeforeLastDosingAsArray[2].split('Substance:    ')[1];
        dosingBeforeLastDosingInfo.lastVial = dosingBeforeLastDosingAsArray[3].split('Vial:              ')[1];
        dosingBeforeLastDosingInfo.lastSampleId = dosingBeforeLastDosingAsArray[4].split('Sample ID:     ')[1];
        dosingBeforeLastDosingInfo.sampleIdArray = dosingBeforeLastDosingInfo.lastSampleId.split(" - ");
        dosingBeforeLastDosingInfo.lastPlateId = dosingBeforeLastDosingInfo.sampleIdArray[0];
        dosingBeforeLastDosingInfo.iniVsCorr = dosingBeforeLastDosingInfo.sampleIdArray[1];
        dosingBeforeLastDosingInfo.lastPlateType = dosingBeforeLastDosingInfo.sampleIdArray[2];
        if (dosingBeforeLastDosingInfo.sampleIdArray.length > 3) dosingBeforeLastDosingInfo.lastDosingStatus = dosingBeforeLastDosingInfo.sampleIdArray[3].split('/');
        dosingBeforeLastDosingInfo.lastTargetDose = dosingBeforeLastDosingAsArray[5].split('Target Dose: ')[1];
        dosingBeforeLastDosingInfo.lastActualWeight = dosingBeforeLastDosingAsArray[7].split('Act. Weight:  ')[1];
        dosingBeforeLastDosingInfo.lastDifferenceMg = dosingBeforeLastDosingAsArray[8].split('Difference:    ')[1];
        dosingBeforeLastDosingInfo.lastDifferencePct = dosingBeforeLastDosingAsArray[9].split('Diff%:           ')[1];
        dosingBeforeLastDosingInfo.lastValidity = dosingBeforeLastDosingAsArray[10].split('Validity:         ')[1];

    }

    return [lastDosingInfo, dosingBeforeLastDosingInfo];
}
/**
 * Main run file that retrieves all text files that were modified less than 5 days ago in a parent folder. from: https://yagisanatode.com/list-all-files-and-folders-in-a-folders-directory-tree-in-google-drive-apps-script/
 */
function getFileAndFolderIds() {

    // The main parent folder.
    const rootId = "1-4mR8IRWpbXziYCGKW1l4QMJ-7IIhdVI";

    let folderData = {};
    let directoryArray = [];
    let folders = [
        {
            name: Drive.Files.get(rootId, { 'fields': 'title' }).title,
            id: rootId
        }
    ];

    // Iterate over each folder in the folders array. 
    while (folders.length) {

        folderData = getItemsForFolderArray(folders); // Calls the Drive API to retrieve all items. 
        let items = folderData.items; // Extracts the items array from the folder data object.


        //Converts items array to the format we need for our Google Sheet and creates a folder array of new child folders to search.
        const itemArrays = createFileArrays(items, folders); // {childFoldersIds, spreadsheetFormatted}
        folders = itemArrays.childFolderIds; // Adds newly found child folders to the folder array.
        directoryArray = directoryArray.concat(itemArrays.spreadsheetFormatted);
    };
    directoryArray.sort(function (a, b) { // the file that was modified the latest ends up on top.
        return a[4] - b[4];
    });


    return directoryArray;
};

/**
 * Retrieves the current list of items from the selected folders from the Drive API.
 * 
 * @see FIELDS {@link https://developers.google.com/drive/api/guides/fields-parameter}
 * 
 * Fields can be modified to your preference here. You can nest fields by using brackets. 
 * @param {Array<Object>} folders - all parent folders to query [{id, name}]
 * @param {String|null} pageToken - The page token should there be more items to retrieve or null if not. 
 * @returns {Object} Object containing a page token and an array of found items of files and 
 * folders {items<array>, nextpageToken}. 
 */
function getItemsForFolderArray(folders, pageToken) {

    const queryString = createQueryString(folders);

    const payload =
    {
        'q': queryString,
        'fields': `items(id, title, mimeType, parents(id), modifiedDate, createdDate), nextPageToken`,
        'supportsAllDrives': true
    };

    if (pageToken) payload.pageToken = pageToken;

    return Drive.Files.list(payload);
};

/**
 * Generates the query string 
 * 
 * called from getCurrentDirectory()
 * 
 * @see QUERY {@link https://developers.google.com/drive/api/guides/ref-search-terms}
 * 
 * @param {Array<Object>} folders - Array of objects containg [{id, name}]
 * @returns {String} The query string for the Drive API request. 
 */
function createQueryString(folders) {

    const MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
    const now = new Date();
    const cutoffDate = new Date(now.getTime() - 5 * MILLIS_PER_DAY);
    const formatedDate = Utilities.formatDate(cutoffDate, "GMT", 'yyyy-MM-dd');

    let queryString = "";

    // If just one folder no need to add brackets.
    if (folders.length === 1) {
        queryString = `'${folders[0].id}' in parents `;
        // Iterate through each folder and create the query for each. 
    } else {
        queryString = `(`;
        folders.forEach((folder, idx) => {
            queryString += (idx === folders.length - 1) ? `'${folder.id}' in parents ` : `'${folder.id}' in parents OR `;
        });
        queryString += `) `;
    }

    // Add any extra queries here. You might add a list of file types. 
    queryString += `AND trashed=false AND (mimeType = "text/plain" OR mimeType = "application/vnd.google-apps.folder") AND modifiedDate > "` + formatedDate + `"`;
    //Excel files:   AND (mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" OR mimeType = "application/vnd.google-apps.folder") 
    // text files:  AND (mimeType = "text/plain" OR mimeType = "application/vnd.google-apps.folder")


    return queryString;
};

/**
 * Iterates through all found items and creates two arrays:
 * 1) spreadsheetFormatted - A 2d array to be added to the selected sheet tab. 
 * conatins [[image, file title, file id, parent name, parent id, file mimeType]]
 * 2) childFolderIds - used to update the folder variable [{id, name}]
 * @param {Array<Object>} folderArray - Array of objects containg [{id, title, mimeType, parents[{id}]}] 
 * @param {Array<Object>} folders - Array of objects containg [{id, name}] 
 * @returns {Object}  
 * 
 */
function createFileArrays(folderArray, folders) {

    let fileArrays = {
        spreadsheetFormatted: [],
        childFolderIds: []
    };

    // Iterate over each found item. 
    folderArray.forEach(file => {

        const isFolder = file.mimeType === "application/vnd.google-apps.folder"; // Is current file a folder?


        //## For shreadsheetFormatted ##

        const fileData = [
            file.title,               // File Name
            file.id,                  // File ID
            new Date(file.createdDate),
            new Date(file.modifiedDate)
        ];
        if (!isFolder) fileArrays.spreadsheetFormatted = fileArrays.spreadsheetFormatted.concat([fileData]);


        //## For childFolderIds ##
        if (isFolder) {
            fileArrays.childFolderIds = fileArrays.childFolderIds.concat([
                {
                    name: file.title,
                    id: file.id
                }
            ]);
        }
    });

    return fileArrays;
};


/*
* ****getDosingFileFromRunlog***  //
*
*param 1: File Name of the runlog text file
*
*returns: url of the first file matching this name (in our case there can only be one)
.
*/

function getDosingFileFromRunlog(fileName = "418_1 ini 2 from 129 onward_Runlog_20230815_074423.txt") {

    fileName = fileName.substring(0, fileName.length - 27) + fileName.substring(fileName.length - 20, fileName.length - 3) + "xlsx"; // cuts out the Runlog_ and replaces the ending in order to arrive at the correct filename
    var files = DriveApp.getFilesByName(fileName);

    while (files.hasNext()) {
        var file = files.next();
        return file.getUrl();

    };


};

//https://docs.google.com/spreadsheets/d/1LXUiCE_au_lEoRGRtngDAF1gIpLOdYo1?rtpof=true&usp=drive_fs