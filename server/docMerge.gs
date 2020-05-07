function docMerge(folderName, mergeType, fileNameHeader, templateDocUrl, shareFile) {
    serveRunning()
    const start = new Date();

    if (logTroubleShootingInfo) {
        Logger.log(
            "Drive folder name: " + folderName +
            "\nMerge type: " + mergeType +
            "\nFile name header: " + fileNameHeader +
            "\nTemplate Doc URL: " + templateDocUrl +
            "\nShare file: " + shareFile
        );
    }

    let templateDoc;
    let templateFile;
    try { // Get template Doc as Document and File objects
        templateDoc = DocumentApp.openByUrl(templateDocUrl);
        templateFile = DriveApp.getFileById(templateDoc.getId());
    } catch(e) {
        Logger.log(e);
        userCache.put("ERROR", e);
        serveError();
        return;
    }

    // Get header and data values
    const lastRow = activeSheet.getLastRow();
    const lastCol = activeSheet.getLastColumn();
    const headerVals = userCache.get("HEADER_VALUES").split("|");
    let dataVals = activeSheet.getSheetValues(2, 1, lastRow - 1, lastCol);
    if (logTroubleShootingInfo) {
        Logger.log(
            "Header: " + headerVals +
            "\nNumber of rows: " + dataVals.length
        );
    }

    validateData(dataVals); // Validate rows

    let mergeFolder;
    let tempDocFolder;
    try { // Create directory structure
        mergeFolder = DriveApp.createFolder(folderName)
                              .setDescription(mergeFolderDescription);
        tempDocFolder = (mergeType == "PDF")
                        ? mergeFolder.createFolder("Temp Docs")
                                     .setDescription(tempDocFolderDescription)
                        : undefined;
    } catch(e) {
        Logger.log(e);
        userCache.put("ERROR", e);
        serveError();
        return;
    }
    const mergeFolderUrl = mergeFolder.getUrl();

    // Setup for main loop
    let [runtimeExceeded, totalMerged, totalIncomplete, totalErrors, rowIndex] = [false, 0, 0, 0, 1];
    const fileNameIndex = headerVals.indexOf(fileNameHeader);
    const emailIndex = shareFile ? getEmailIndex(headerVals) : undefined;
    const replacementTasks = getReplacementTasks(templateDoc, headerVals);

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

        let docCopyFile;
        let newDoc;
        let fileName = row[fileNameIndex];
        try { // Copy template Doc
            docCopyFile = (mergeType == "PDF")
                          ? templateFile.makeCopy(tempDocFolder)
                          : templateFile.makeCopy(mergeFolder);
            newDoc = DocumentApp.openById(docCopyFile.getId()).setName(fileName);
        } catch(e) {
            Logger.log(e);
            totalErrors++;
            mergeStatusCell.setValue("Error");
            SpreadsheetApp.flush();
            continue;
        }

        // Perform replacement in copied template Doc
        for (const section in replacementTasks) {
            let replacementSection;
            if (section == "header") {
                replacementSection = newDoc.getHeader();
            } else if (section == "body") {
                replacementSection = newDoc.getBody();
            } else if (section == "footer") {
                replacementSection = newDoc.getFooter();
            }
            let dataIndices = replacementTasks[section];
            dataIndices.forEach((index) => {
                let toBeReplaced = "<<" + headerVals[index] + ">>";
                let replacement = row[index];
                replacementSection.replaceText(toBeReplaced, replacement);
            });
        }
        newDoc.saveAndClose();

        let pdfFile;
        if (mergeType == "PDF") {
            try { // Create PDF from each new Doc
                let pdfBlob = newDoc.getAs("application/pdf");
                pdfFile = mergeFolder.createFile(pdfBlob);
            } catch(e) {
                Logger.log(e);
                totalErrors++;
                mergeStatusCell.setValue("Error");
                SpreadsheetApp.flush();
                continue;
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

    const end = new Date(); // Time completion
    const finishTime = end - start; // MS
    const minutes = Math.round((finishTime / 60000) * 100) / 100;

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
            "Runtime exceeded: " + runtimeExceeded +
            "\nMerge folder URL: " + mergeFolderUrl +
            "\nTotal rows: " + dataVals.length +
            "\nTotal merged: " + totalMerged +
            "\nTotal incomplete: " + totalIncomplete +
            "\nTotal errors: " + totalErrors +
            "\nCompletion time in minutes: " + minutes
        );
    }
}