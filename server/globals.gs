// Variables
const ss = SpreadsheetApp.getActiveSpreadsheet();
const activeSheet = SpreadsheetApp.getActiveSheet();
const userCache = CacheService.getUserCache();

// Configuration
const logTroubleShootingInfo = true;

const appTitle = "Doc Merge";

const scriptRuntimeLimit = 1740; // Seconds = 29 Minutes
const runtimeExceededToastTitle = "Runtime Exceeded";
const runtimeExceededToastMsg = appTitle + " will start where it left off if you run it again";

const mergeFolderDescription = "Created via Google Apps Script by " + appTitle;
const tempDocFolderDescription = "Created by " + appTitle + " for holding Docs";