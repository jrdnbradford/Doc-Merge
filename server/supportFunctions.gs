function onOpen() {
    SpreadsheetApp.getUi().createMenu(appTitle)
        .addItem("Run " + appTitle, "serveIndex")
        .addItem("Help", "serveHelp")
        .addItem("License", "serveLicense")
        .addToUi();
}


function validateData(data) {
    // Parameter data is array of arrays
    // Function trims whitespace and checks for incomplete rows
    data.forEach((row) => {
        const mergeStatus = row[row.length - 1];
        for (let i = 0; i < row.length - 1; i++) {
            if (typeof row[i] == "string") {
                row[i] = row[i].trim();
            }
            if (row[i] === "") {
                if (mergeStatus != "Incomplete Data") {
                    row[row.length - 1] = "Incomplete Data";
                }
                return;
            }
        }
        if (mergeStatus == "Incomplete Data") {
            row[row.length - 1] = "";
        }
    });
    Logger.log("Data validation complete");
}


function getUserCache(keys) {
    // Used by client side to get cached properties
    return userCache.getAll(keys);
}


function getEmailIndex(headers) {
    // Returns index of email column
    for (let i = 0; i < headers.length; i++) {
        if (typeof headers[i] == "string") {
            const includesEmail = headers[i].replace(" ", "").toLowerCase().includes("email");
            if (includesEmail) {
                return headers.indexOf(headers[i]);
            }
        }
    }
}


function getHeaderVals() {
    // Adds merge status column
    // Get, cache, and return header values from active sheet
    const lastCol = activeSheet.getLastColumn();
    const lastColVal = activeSheet.getRange(1, lastCol).getValue();
    if (lastColVal != "Merge Status") {
        const mergeCell = activeSheet.getRange(1, lastCol + 1);
        mergeCell.setValue("Merge Status");
    }

    let headerVals = activeSheet.getSheetValues(1, 1, 1, activeSheet.getLastColumn())[0];
    for (let i = 0; i < headerVals.length; i++) {
        if (typeof headerVals[i] == "string") {
            headerVals[i] = headerVals[i].trim();
        }
    }
    let headerValStr = headerVals.join("|");
    userCache.put("HEADER_VALUES", headerValStr);
    headerVals.pop(); // Minus Merge Status
    return headerVals;
}


function getMergeSetup(url) {
    // Parameter url is provided by user and used to get template doc info
    // Runs on input event
    if (url == "") {
        return {url: ""};
    }
    let doc;
    try {
        doc = DocumentApp.openByUrl(url);
    } catch(e) {
        try {
            doc = DocumentApp.openById(url);
        } catch(e) {
            return {url: null};
        }
    }
    const docName = doc.getName();
    const docUrl = doc.getUrl();
    const folderName = getMergeFolderName(docName + " Merge");
    return {name: docName, url: docUrl, folderName: folderName};
}


function getMergeFolderName(name) {
    // Is this necessary?
    // Parameter name is a string
    // Check if folder already exists, cache and return folder name
    let fileIndex = 0;
    const originalName = name;
    const driveRoot = DriveApp.getRootFolder();
    let identicallyNamedFolders = driveRoot.getFoldersByName(name);
    while (identicallyNamedFolders.hasNext()) {
        fileIndex++;
        name = originalName + " (" + fileIndex + ")";
        identicallyNamedFolders = driveRoot.getFoldersByName(name);
    }
    return name;
}


function getReplacementTasks(template, headers) {
    // Parameter template is a Google Doc
    // Parameter headers is an array of strings
    // Used by docMerge.gs to perform only necessary text replacements
    // for each section. Returns an object that specifies indices of
    // data needing replacement.
    let tasks = {header: [], body: [], footer: []};
    let sections = [template.getHeader(), template.getBody(), template.getFooter()]
        .filter(function(section) {
            return section != null;
    });

    sections.forEach((section) => {
        let sectionType = section.getType();
        for (let i = 0; i < headers.length - 1; i++) {
            if (section.findText("<<" + headers[i] + ">>")) {
                if (sectionType == "HEADER_SECTION") {
                    tasks.header.push(i);
                }
                else if (sectionType == "BODY_SECTION") {
                    tasks.body.push(i);
                }
                else if (sectionType == "FOOTER_SECTION") {
                    tasks.footer.push(i);
                }
            }
        }
    });
    return tasks;
}