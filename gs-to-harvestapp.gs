// Set up the Google Spreadsheets dropdown menu
function onOpen() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuEntries = [{name: "Add tasks and hours to Harvest", functionName: "scanSheet"},
    {name: "Add hours to Harvest", functionName: "updateProjectTaskHours"},
    {name: "Authenticate with Harvest", functionName: "startService"},
    {name: "Generate redirect URI", functionName: "generateRedirectURI"}];
    ss.addMenu("Harvest", menuEntries);
}

// Get/update the control values.
function checkControlValues() {
    var col = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Controls").getRange("C1:C29").getValues();

    var sheetName = col[27][0].toString().trim();
    if (sheetName == "") {
        return "No sheet selected. Update Controls sheet.";
    } else if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)) {
        return "Sheet not found. Update Controls sheet.";
    }
    ScriptProperties.setProperty("sheetName", sheetName);

    var subdomain = col[28][0].toString().trim();
    if (subdomain == "") {
        return "No subdomain selected. Update Controls sheet."
    }
    ScriptProperties.setProperty("baseUrl", "https://" + subdomain + ".harvestapp.com");

    var projectId = col[16][0].toString().trim();
    if (projectId == "") {
        return "No project selected. Update Controls sheet."
    }
    ScriptProperties.setProperty("projectId", projectId);

    var clientId = col[23][0].toString().trim();
    if (clientId == "") {
        return "Client ID not found. Update Controls sheet."
    }
    ScriptProperties.setProperty("clientId", clientId);

    var clientSecret = col[24][0].toString().trim();
    if (clientSecret == "") {
        return "Client secret not found. Update Controls sheet."
    }
    ScriptProperties.setProperty("clientSecret", clientSecret);

    return "";
}

// Loop through the main sheet and upload stories to Harvest.
// If upload is toggled to false, get story hours and return them. Do not upload.
function scanSheet(upload) {
    upload = (typeof upload == 'undefined' ? true : upload);

    var statusCol = 0;
    var epicCol = statusCol + 3;
    var titleCol = statusCol + 5;
    var hoursCol = statusCol + 12;

    var startRow = 10;

    var hours = [];

    var startTime = new Date();
    Logger.log("Started scanning sheet for Harvest " + (upload ? "uploads" : "hour updates") + " at:" + startTime);
    var error = checkControlValues();
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

            currentTime = new Date();

            Logger.log("Row " + r + ":" + currentTime);

            if (currentTime.valueOf() - startTime.valueOf() >= 330000) { // 5.5 minutes - scripts time out at 6 minutes
                Browser.msgBox("NOTICE: Script was about to time out so it has been terminated gracefully . " + (!upload ? successCount + " tasks were uploaded successfully." : ""));
                return;
            }

            if (upload) {
                var status = currentRow[statusCol];

                if (status == ".") { // Row already processed.
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
            } else {
                if (currentRow[hoursCol] > 0) {
                    hours.push([currentRow[titleCol], currentRow[hoursCol]]);
                } else {
                    Logger.log("Ignoring row " + r + ". No hours found.");
                }
            }
        }
    }

    if (!upload) {

        return hours;
    } else {
        Browser.msgBox( successCount + " Harvest tasks were uploaded successfully. Adding hours...");
        updateProjectTaskHours();
        return;
    }
}

// Puts the project key in the Controls sheet, for user convenience.
function generateRedirectURI() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Controls");

    var projectKeyCell = sheet.getRange(20, 3, 1, 1);

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

    var url = ScriptProperties.getProperty("baseUrl") + "/projects/" + ScriptProperties.getProperty("projectId") + "/task_assignments/add_with_create_new_task/";

    var response = UrlFetchApp.fetch(url, options);
}

// Applies hour budgets in sheet to tasks which do not yet have hour budgets.
/**
 * Logic flow:
 * - get all tasks for a project
 * - store task id's in an array
 * - don't store tasks with hour budgets
 * - for each task without hours lookup the task's name, store that in the array
 * - don't store tasks which are defaults
 * - loop through the sheet, match story names, put their hour values in the array
 * - send an update for each 'task assignment' that we have hours for and but harvest doesn't
 */
function updateProjectTaskHours() {
    var error = checkControlValues();
    if (error != "") {
        Browser.msgBox("ERROR:Values in the Controls sheet have not been set. Please fix the following error:\n " + error);
        return;
    }

    var successCount = 0;

    var result = JSON.parse(getProjectTasks());

    var tasks = [];
    // Get the ID's for the tasks we're interested in.
    for (var i = 0; i < result.length; i++) {
        // Get tasks that do not have hours assigned.
        if (result[i].task_assignment.budget == null || result[i].task_assignment.budget == 0) {
            tasks.push([result[i].task_assignment.id, result[i].task_assignment.task_id]);
        }
    }

    var filteredTasks = [];
    // Get the names for the tasks we're interested in.
    for (var i = 0; i < tasks.length; i++) {
        var taskDetails = JSON.parse(getTaskDetails(tasks[i][1]));
        // Only look at non-default tasks.
        if (taskDetails.task.is_default == false) {
            filteredTasks.push([tasks[i][0], tasks[i][1], taskDetails.task.name]);
        }
    }

    var toUpload = [];
    var hours = scanSheet(false);
    // Loop through the sheet, get hours for the tasks we've filtered.
    for (var i = 0; i < hours.length; i++) {
        for (var j = 0; j < filteredTasks.length; j++) {
            // If the names match
            if (filteredTasks[j][2] == hours[i][0]) {
                toUpload.push([filteredTasks[j][0], hours[i][1]]);
            }
        }
    }

    // Upload new task budgets for tasks which do not already have budgets.
    for (i = 0; i < toUpload.length; i++) {
        var response = putProjectTaskHours(toUpload[i][0], toUpload[i][1]);
        successCount++;
    }

    Browser.msgBox(successCount + " Harvest tasks had hours added. ");
}

// Returns all tasks that exist in a project.
function getProjectTasks() {
    var url = ScriptProperties.getProperty("baseUrl") + "/projects/" + ScriptProperties.getProperty("projectId") + "/task_assignments/";
    return restGet(url);
}

// Returns all details for a task.
function getTaskDetails(taskId) {
    var url = ScriptProperties.getProperty("baseUrl") + "/tasks/" + taskId + "/";
    return restGet(url);
}

function putProjectTaskHours(assignmentId, hours) {
    var harvestService = getHarvestService();

    var headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Authorization": "Bearer " + harvestService.getAccessToken()
    }

    var payload = {
        "task_assignment": {
            "budget": hours
        }
    }

    payload = JSON.stringify(payload);

    var options = {
        "method": "put",
        "headers": headers,
        "payload": payload
    }

    var url = ScriptProperties.getProperty("baseUrl") + "/projects/" + ScriptProperties.getProperty("projectId") + "/task_assignments/" + assignmentId + "/";
 
    var response = UrlFetchApp.fetch(url, options);
}

// Performs a GET on the url you pass in.
function restGet(url) {
    var harvestService = getHarvestService();

    var headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Authorization": "Bearer " + harvestService.getAccessToken()
    }

    var options = {
        "method": "get",
        "headers": headers
    }

    var response = UrlFetchApp.fetch(url, options);

    return response;
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
