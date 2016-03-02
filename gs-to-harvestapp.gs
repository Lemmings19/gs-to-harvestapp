// Set up the Google Spreadsheets dropdown menu
function onOpen() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuEntries = [{name: "Create Harvest Tasks", functionName: "upload"},
    {name: "Authenticate with Harvest", functionName: "startService"},
    {name: "Generate redirect URI", functionName: "generateRedirectURI"}];
    ss.addMenu("Harvest", menuEntries);
}

// Get/update the control values.
function checkControlValues(requireClientId, requireClientSecret) {
    var col = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Controls").getRange("B3:B24").getValues();

    var sheetName = col[4][0].toString().trim();
    if (sheetName == "") {
        return "No sheet selected. Update Controls sheet.";
    } else if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)) {
        return "Sheet not found. Update Controls sheet.";
    }
    ScriptProperties.setProperty("sheetName", sheetName);

    var projectId = col[13][0].toString().trim();
    if (projectId == "") {
        return "No project selected. Update Controls sheet."
    }
    ScriptProperties.setProperty("projectId", projectId);

    var clientId = col[20][0].toString().trim();
    if (requireClientId && clientId == "") {
        return "Client ID not found. Update Controls sheet."
    }
    ScriptProperties.setProperty("clientId", clientId);

    var clientSecret = col[21][0].toString().trim();
    if (clientSecret && clientId == "") {
        return "Client secret not found. Update Controls sheet."
    }
    ScriptProperties.setProperty("clientSecret", clientSecret);

    return "";
}

// Commit spreadsheet cells to Harvest
function upload() {
    var statusCol = 0;
    var epicCol = statusCol + 3;
    var titleCol = statusCol + 5;

    var startRow = 10;

    var startTime = new Date();
    Logger.log("Started Harvest export at:" + startTime);
    var error = checkControlValues(true, true);
    if (error != "") {
        Browser.msgBox("ERROR:Values in the Controls sheet have not been set. Please fix the following error:\n " + error);
        return;
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ScriptProperties.getProperty("sheetName"));

    var successCount = 0;
    var partialCount = 0;
    var rows = sheet.getDataRange().getValues();

    for (var i = startRow ; i < rows.length ; i++) {
        var currentRow = rows[i];

        // Only process if there is a card title and an epic.
        if (currentRow[titleCol].trim() != "" && currentRow[epicCol] != "") {
            r = i + 1;

            var status = currentRow[statusCol];

            currentTime = new Date();

            Logger.log("Row " + r + ":" + currentTime);

            if (currentTime.valueOf() - startTime.valueOf() >= 330000) { // 5.5 minutes - scripts time out at 6 minutes
                Browser.msgBox("NOTICE: Script was about to time out so upload has been terminated gracefully ." + successCount + " tasks were uploaded successfully.");
                return;
            } else if (status == ".") { // Row already processed.
                Logger.log("Ignoring row " + r + ". Status column indicates already imported.");
            } else if (status == "x") {
                Browser.msgBox("ERROR: Row " + r + " indicates that it was partially created the last time this script was run. Ask the developer about what may have caused this.");
                return;
            } else if (status == "") { // Status cell empty. Import row.

                var statusCell = sheet.getRange(r, statusCol + 1, 1, 1);

                // Indicate that this row has begun importing.
                statusCell.setValue("x");

                partialCount++;

                createProjectTask(currentRow[titleCol]);

                // Indicate that this row has been exported.
                statusCell.setValue(".");

                SpreadsheetApp.flush();
                partialCount --;
                successCount ++;
            }
        }
    }

    Browser.msgBox( successCount + " Harvest tasks items were uploaded successfully.");

    return;
}

// Puts the project key in the Controls sheet, for user convenience.
function generateRedirectURI() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Controls");

    var projectKeyCell = sheet.getRange(19, 2, 1, 1);

    projectKeyCell.setValue("https://script.google.com/macros/d/" + ScriptApp.getProjectKey() + "/usercallback");
}

// Creates a task and assigns it to a project.
// This is a good example for how all REST messaging will work.
function createProjectTask(title) {
    var harvestService = getHarvestService();

    var headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Authorization": "Bearer " + harvestService.getAccessToken()
    }

    var payload = {
        "task": {
            "name": title,
            "billable_by_default": true,
            "is_default": false,
            "default_hourly_rate": 0,
            "deactivated": false
        }
    }

    payload = JSON.stringify(payload);

    var options = {
        "method": "post",
        "headers": headers,
        "payload": payload
    }

    var url = "https://joelynch.harvestapp.com/projects/" + ScriptProperties.getProperty("projectId") + "/task_assignments/add_with_create_new_task";

    var response = UrlFetchApp.fetch(url, options);
}

// Begin OAuth2 with GUI.
function startService() {
    var error = checkControlValues(true, true);
    if (error != "") {
        Browser.msgBox("ERROR:Values in the Controls sheet have not been set. Please fix the following error:\n " + error);
        return;
    }

    var harvestService = getHarvestService();
    if (!harvestService.hasAccess()) {
        var authorizationUrl = harvestService.getAuthorizationUrl();
        var template = HtmlService.createTemplate('Click <a href="' + authorizationUrl + '" target="_blank">here</a> to authenticate.');
        var page = template.evaluate();
        SpreadsheetApp.getUi().showSidebar(page);
        return authorizationUrl;
    } else {
        return null;
    }
}

// Using the OAuth2 library, configure authentication variables.
function getHarvestService() {
    return OAuth2.createService('harvest')
        .setAuthorizationBaseUrl('https://api.harvestapp.com/oauth2/authorize')
        .setTokenUrl('https://api.harvestapp.com/oauth2/token')
        .setClientId(ScriptProperties.getProperty("clientId"))
        .setClientSecret(ScriptProperties.getProperty("clientSecret"))
        .setCallbackFunction('authCallback')
        .setPropertyStore(PropertiesService.getUserProperties())
        .setTokenFormat(OAuth2.TOKEN_FORMAT.JSON);

}

// More OAuth2 stuff.
function authCallback(request) {
    var harvestService = getHarvestService();
    var isAuthorized = harvestService.handleCallback(request);
    if (isAuthorized) {
        return HtmlService.createHtmlOutput('Success! You can close this tab.');
    } else {
        return HtmlService.createHtmlOutput('Denied. You can close this tab');
    }
}

// More OAuth2 stuff.
function clearService(){
    OAuth2.createService('harvest')
        .setPropertyStore(PropertiesService.getUserProperties())
        .reset();
}
