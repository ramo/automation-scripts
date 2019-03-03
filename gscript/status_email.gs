/**
 * status_email.gs  - Google app script to send status email and do related operations.
 *
 * Author: Ramo
 */



/**
 * On invocation triggers an email for the preconfigured settings.
 * Mail will be sent if and only if 'run' field in config.json is true and today
 * is a working day.
 * Actions: 1. Send email to configured users with the status document.
 *          2. Take backup of the status document.
 *          3. Reset the status document for next day use.
 */
function sendEmail() {

  // Load configuration
  var confId = '<confId>';  // Caution: This need to be set before running the script.
  var config = getConfig(confId);

  if (!config.run) {
    Logger.log('Configuration run is not enabled. Hence not running the sendEmail()');
    return;
  }

  var today = new Date();
  if (!isWorkingDay(today, config.holidays)) {
    Logger.log('Not a working day. Hence not running the sendEmail()');
    return;
  }
    
  var ta = splitDate(today);
  var reportName = config.reportPrefix + '-' + ta.join('-') + '.pdf';  

  // prepare the pdf report from shared document
  var doc = DocumentApp.openById(config.mailDocId);  
  var report = doc.getBlob().copyBlob();
  report.setName(reportName);
  
  // prepare the mail stuffs
  var subject = 'Status Report - ' + ta.join('/');
  var to = config.to.join(',');
  
  var options = new Object();
  options.cc = config.cc.join(',');
  options.attachments = [report];
  
  // Send email
  GmailApp.sendEmail(to, subject, config.mailBody, options);

  // Take a backup of the status document.
  takeBackup(config.mailDocId);

  // reset the document for nextday purpose
  var docBody = doc.getBody();
  emptyDocument(docBody);
  var tomorrow = nextWorkingDay(today, config.holidays);
  var tma = splitDate(tomorrow);
  var docHeader = 'Status Report - ' + tma.join('/');
  docBody.insertParagraph(0, docHeader)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .editAsText().setBold(true);
  docBody.appendParagraph('\n\n--- Remove me and start writing your status from here. ---');
  doc.saveAndClose();
}

// This function required to emtpy the document,
// if it is not ended with carriage return. 
function emptyDocument(body) {
  body.appendParagraph('');// to be sure to delete the last paragraph in case it doesn't end with a cr/lf
  while (body.getNumChildren() > 1) {
    body.removeChild(body.getChild(0));
  }
}


/**
 *
 * Checks if the given date is a working day.
 * Current behavior: Given date is a working day if
 *              - Input date is a week day
 *              - Input date is not a holiday
 * Note: holiday is decided from the given holidays set
 *       as the input.
 * 
 */
function isWorkingDay(date, holidays) {
  var day = date.getDay();
  
  // checking for weekday
  if (day < 1 || day > 5) {
    return false;
  }
  
  var sp = splitDate(date);
  // checking for holiday or not
  return holidays.indexOf(sp[0] + '-' + sp[1]) == -1;
}



/**
 * Returns the next working day for any given date.
 * Depends on isWorkingDay() function. 
 */
function nextWorkingDay(date, holidays) {
  var cd = date;
  do {
    cd = nextDay(cd);
  } while(!isWorkingDay(cd, holidays));
  return cd;
}


/**
 *
 * Returns the next day for the given date.
 */
function nextDay(date) {
  var nd = new Date(date);
  nd.setDate(date.getDate() + 1);
  return nd;
}


/**
 * utility method to format the given number
 * as 2 digit number string.
 */
function format0(n) {
  if (n < 10) {
    n = '0' + n;
  }
  return n;
}


/**
 * utility method to split the date object into 
 * array of month, day, year.
 */
function splitDate(date) {
  var y = date.getFullYear() - 2000;
  var m = format0(date.getMonth() + 1);
  var d = format0(date.getDate());
  return [m, d, y];
}

/**
 * Take backup of the given document idenfied by the docId.
 * Creates a backup under folder named Backup. 
 * If the folder is not present, new one is created otherwise
 * existing one is used.
 */
function takeBackup(docId) {
  var file = DriveApp.getFileById(docId);
  var sd = file.getParents().next();
  var ddfp = sd.getFoldersByName('Backup');
  var dd = ddfp.hasNext() ? ddfp.next() : sd.createFolder('Backup');
  file.makeCopy(dd);
}


/**
 * Loads the config.json identified by the confId.
 * File should be a proper json file for the loading to 
 * be successful. Otherwise error will be thrown.
 */
function getConfig(confId) {

/**
 * Sample config.json
 *  {
 *    "run": true,
 *    "mailDocId": "<docId>",
 *    "holidays": [
 *      "05-01",
 *      "06-05",
 *      "08-15",
 *      "09-02",
 *      "10-02",
 *      "10-08",
 *      "10-28",
 *      "12-25"
 *    ],
 *    "mailBody": "Some mail contents",
 *    "to": [
 *      "<email1>",
 *      "<email2>"
 *    ],
 *    "cc": [
 *      "<email3>",
 *      "<email4>"
 *    ],
 *    "reprtPrefix": "MyTeamStatusReport"
 *  }
 * 
 */

  var confFile = DriveApp.getFileById(confId);
  var config = JSON.parse(confFile.getAs('application/json').getDataAsString('UTF-8'));
  return config;
}
