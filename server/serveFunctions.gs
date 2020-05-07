function serveIndex() {
    let html = HtmlService.createTemplateFromFile("index")
                  .evaluate().setWidth(350).setHeight(450);

    SpreadsheetApp.getUi().showModalDialog(html, appTitle);
}


function serveRunning() {
    let html = HtmlService.createTemplateFromFile("running")
                  .evaluate().setWidth(350).setHeight(200);

    SpreadsheetApp.getUi().showModalDialog(html, appTitle);
}


function serveHelp() {
    let html = HtmlService.createTemplateFromFile("help")
                  .evaluate().setWidth(500).setHeight(600);

    SpreadsheetApp.getUi().showModalDialog(html, appTitle + " Help");
}


function serveComplete() {
    let html = HtmlService.createTemplateFromFile("complete")
                  .evaluate().setWidth(400).setHeight(300);

    SpreadsheetApp.getUi().showModalDialog(html, appTitle + " Complete");
}


function serveError() {
    let html = HtmlService.createTemplateFromFile("error")
                  .evaluate().setWidth(500).setHeight(300);

    SpreadsheetApp.getUi().showModalDialog(html, appTitle + " Error");
}


function serveLicense() {
    let html = HtmlService.createTemplateFromFile("license")
                  .evaluate().setWidth(500).setHeight(300);

    SpreadsheetApp.getUi().showModalDialog(html, appTitle + " License");
}