/**
 * status_email.gs v2 - Google app script to send status email and do related operations.
 *
 * Author: Ramo
 */



/**
 * On invocation triggers an email for the preconfigured settings.
 * Mail will be sent if and only if 'run' field in config.json is true and today
 * is a working day.
 * Actions: 1. Disable write access to everyone for avoiding any race condition.
 *          2. Copy the Today sheet contents to Daily status main spreadsheet.
 *          3. Send email to configured users with the status document.
 *          4. Reset the Today status spreadsheet for next day use.
 *          5. Enable write access to everyone.
 */
function sendStatusEmail() {
  var configId = '<configId>';  // Caution: This need to be set before running the script.
  var config = getStatusConfig(configId);

  var pre_process = (config) => {
    var todayFile = DriveApp.getFileById(config.todaySheetId);
    var editors = todayFile.getEditors();
    editors.forEach(user => todayFile.removeEditor(user));
    config.store.editors = editors;
    
    var statusSheet = SpreadsheetApp.openById(config.statusSheetId);
    var todaySheet = SpreadsheetApp.openById(config.todaySheetId);
    todaySheet.getSheets()[0].copyTo(statusSheet);
    statusSheet.getSheets()[statusSheet.getSheets().length - 1].activate();
    statusSheet.moveActiveSheet(1);
    var mdy = splitDateToMDYArray(config.today);
    var todaySheetName = config.sprintId + '_' + mdy[0] + mdy[1];
    statusSheet.getActiveSheet().setName(todaySheetName);
  };

  var post_process = (config) => {
    clearCellsInColor(config.todaySheetId, '#ffffff');
    var todayFile = DriveApp.getFileById(config.todaySheetId);
    config.store.editors.forEach(user => todayFile.addEditor(user));
  };

  sendEmail(config, pre_process, post_process);
}

