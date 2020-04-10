function docMerge(folderName, mergeType, fileNameHeader, templateDocUrl, shareFile) {
    let start = new Date();

    if (logTroubleShootingInfo) {
        Logger.log(
            "Drive folder name: " + folderName + "\n" +
            "Merge type: " + mergeType + "\n" +
            "File name header: " + fileNameHeader + "\n" +
            "Template Doc URL: " + templateDocUrl + "\n" +
            "Share file: " + shareFile
        );
    }

    try { // Get template Doc as Document and File objects
        var templateDoc = DocumentApp.openByUrl(templateDocUrl);
        var templateFile = DriveApp.getFileById(templateDoc.getId());
    } catch(e) {
        Logger.log(e);
        userCache.put("ERROR", e);
        serveError();
        return;
    }

    // Get header and data values
    let lastRow = activeSheet.getLastRow();
    let lastCol = activeSheet.getLastColumn();
    let dataVals = activeSheet.getSheetValues(2, 1, lastRow - 1, lastCol);
    let headerVals = userCache.get("HEADER_VALUES").split("|");
    if (logTroubleShootingInfo) {
        Logger.log(
            "Header: " + headerVals + "\n" +
            "Number of rows: " + dataVals.length
        );
    }

    validateData(dataVals); // Validate rows

    try { // Create directory structure
        var mergeFolder = DriveApp.createFolder(folderName)
                                  .setDescription(mergeFolderDescription);
        var tempDocFolder = (mergeType == "PDF")
                            ? mergeFolder.createFolder("Temp Docs")
                                         .setDescription(tempDocFolderDescription)
                            : undefined;
    } catch(e) {
        Logger.log(e);
        userCache.put("ERROR", e);
        serveError();
        return;
    }
    let mergeFolderUrl = mergeFolder.getUrl();

    // Setup for main loop
    let [runtimeExceeded, totalMerged, totalIncomplete, totalErrors, rowIndex] = [false, 0, 0, 0, 1];
    let fileNameIndex = headerVals.indexOf(fileNameHeader);
    let emailIndex = shareFile ? getEmailIndex(headerVals) : undefined;
    let replacementTasks = getReplacementTasks(templateDoc, headerVals);

    for (const row of dataVals) { // Create new Docs/PDFs
        rowIndex++;

        let mergeStatusCell = activeSheet.getRange(rowIndex, lastCol);

        let mergeStatus = row[row.length - 1]; // Skip certain rows
        if (mergeStatus === "Complete" || mergeStatus === "Incomplete Data") {
            if (mergeStatusCell.getValue() != mergeStatus) {
                mergeStatusCell.setValue(mergeStatus);
                SpreadsheetApp.flush();
            }
            if (mergeStatus === "Incomplete Data") {
                totalIncomplete++;
            }
            continue;
        }

        let fileName = row[fileNameIndex];
        try { // Copy template Doc
            var docCopyFile = (mergeType == "PDF")
                            ? templateFile.makeCopy(tempDocFolder)
                            : templateFile.makeCopy(mergeFolder);
            var newDoc = DocumentApp.openById(docCopyFile.getId()).setName(fileName);
        } catch(e) {
            Logger.log(e);
            totalErrors++;
            mergeStatusCell.setValue("Error");
            SpreadsheetApp.flush();
            continue;
        }

        // Perform replacement in copied template Doc
        for (const section in replacementTasks) {
            if (section == "header") {
                var replacementSection = newDoc.getHeader();
            } else if (section == "body") {
                var replacementSection = newDoc.getBody();
            } else if (section == "footer") {
                var replacementSection = newDoc.getFooter();
            }
            let dataIndices = replacementTasks[section];
            dataIndices.forEach((index) => {
                let toBeReplaced = "<<" + headerVals[index] + ">>";
                let replacement = row[index];
                replacementSection.replaceText(toBeReplaced, replacement);
            });
        }
        newDoc.saveAndClose();

        if (mergeType == "PDF") {
            try { // Create PDF from each new Doc
                let pdfBlob = newDoc.getAs("application/pdf");
                var pdfFile = mergeFolder.createFile(pdfBlob);
            } catch(e) {
                Logger.log(e);
                totalErrors++;
                mergeStatusCell.setValue("Error");
                SpreadsheetApp.flush();
            }
        }

        if (shareFile) { // Share file
            let emailAddress = row[emailIndex];
            let fileToShare = (mergeType == "PDF") ? pdfFile : newDoc;
            fileToShare.addViewer(emailAddress); // If Doc, doesn't email
        }

        // Finish
        mergeStatusCell.setValue("Complete");
        SpreadsheetApp.flush();
        totalMerged++;

        // Break if getting close to script runtime limitation
        let runningTime = (new Date() - start) / 1000; // Seconds
        if (runningTime >= scriptRuntimeLimit) {
            runtimeExceeded = true;
            ss.toast(runtimeExceededToastMsg, runtimeExceededToastTitle, -1);
            break;
        }
    } // End loop

    if (mergeType == "PDF") { // Trash Temp Docs folder
        tempDocFolder.setTrashed(true);
        if (logTroubleShootingInfo) {
            Logger.log("Temp Docs folder trashed");
        }
    }

    let end = new Date(); // Time completion
    let finishTime = end - start; // MS
    let minutes = Math.round((finishTime / 60000) * 100) / 100;

    // Cache data needed for complete.html
    userCache.putAll({ // Requires strings
        "FOLDER_NAME": folderName,
        "FOLDER_URL": mergeFolderUrl,
        "MINUTES": minutes.toString(),
        "TOTAL ROWS": dataVals.length.toString(),
        "TOTAL_MERGED": totalMerged.toString(),
        "TOTAL_INCOMPLETE": totalIncomplete.toString(),
        "TOTAL_ERRORS": totalErrors.toString(),
        "RUNTIME_EXCEEDED": runtimeExceeded.toString()
    }, 30);

    serveComplete();

    if (logTroubleShootingInfo) {
        Logger.log(
            "Runtime exceeded: " + runtimeExceeded + "\n" +
            "Merge folder URL: " + mergeFolderUrl + "\n" +
            "Total rows: " + dataVals.length + "\n" +
            "Total merged: " + totalMerged + "\n" +
            "Total incomplete: " + totalIncomplete + "\n" +
            "Total errors: " + totalErrors + "\n" +
            "Completion time in minutes: " + minutes
        );
    }
}