/**
 * common_lib.gs  - Common library methods to be used in status email Google app script.
 *
 * Author: Ramo
 */


 /**
 * On invocation triggers an email for the preconfigured settings.
 * Mail will be sent if and only if 'run' field in config is true and today
 * is a working day.
 * Actions: 1. Do preprocess by calling the callback.
 *          2. Send email to configured users with the status document.
 *          3. Do postprocess by calling the callback.
 */
function sendEmail(config, preprocess, postprocess) {
  if (!config.emergency) {
    if (!config.run) {
      Logger.log('Configuration run is not enabled. Hence not running the sendEmail()');
      return;
    }
    
    if (!isWorkingDay(config.today, config.holidays)) {
      Logger.log('Not a working day. Hence not running the sendEmail()');
      return;
    }  
  }

  preprocess && preprocess(config);
  var mdy = splitDateToMDYArray(config.today);
  var subject = config.mail.subjectPrefix + ' - ' + mdy.join('/');
  var to = config.mail.to.join(',');
  var options = new Object();
  options.cc = config.mail.cc.join(',');
  options.htmlBody = config.mail.htmlBody;
  options.attachments = config.store.attachments;
  GmailApp.sendEmail(to, subject, config.mail.htmlBody, options);
  postprocess && postprocess(config);
}



/**
 * 
 * Empty a document by removing all its contents. 
 * This method handles scenario in which the last 
 * paragraph doesn't end with the cr/lf also.
 */
function clearDocument(docId) {
  var doc = DocumentApp.openById(docId);
  var docBody = doc.getBody();
  docBody.appendParagraph(''); // to be sure to delete the last paragraph in case it doesn't end with a cr/lf
  while (docBody.getNumChildren() > 1) {
    docBody.removeChild(docBody.getChild(0));
  }
  doc.saveAndClose(); // TODO Need to confirm this step.
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
  
  var sp = splitDateToMDYArray(date);
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
function splitDateToMDYArray(date) {
  var y = date.getFullYear() - 2000;
  var m = format0(date.getMonth() + 1);
  var d = format0(date.getDate());
  return [m, d, y];
}



/**
 * Take backup of the given GDrive file idenfied by the fileId.
 * Creates a backup under folder named Backup. 
 * If the folder is not present, new one is created otherwise
 * existing one is used.
 */
function takeBackup(fileId) {
  var file = DriveApp.getFileById(fileId);
  var sd = file.getParents().next();
  var ddfp = sd.getFoldersByName('Backup');
  var dd = ddfp.hasNext() ? ddfp.next() : sd.createFolder('Backup');
  file.makeCopy(dd);
}


/**
 * Helper method to clear cells in a sheet with 
 * specific color background.
 */
function clearCellsInColor(sheetId, color) {
  var ss = SpreadsheetApp.openById(sheetId);
  ss.getSheets().forEach(sheet => {
    for (var row = 1; row <= sheet.getLastRow(); row++) {
      for (var col = 1; col <= sheet.getLastColumn(); col++) {
        var cell = sheet.getRange(row, col);
        var bg = cell.getBackground();
        if (bg === color) {
          cell.clearContent();
        }
      }
    }
  });
}

function addEditorSilently(fileId, userEmail) {
  Drive.Permissions.insert(
   {
     'role': 'writer',
     'type': 'user',
     'value': userEmail
   },
   fileId,
   {
     'sendNotificationEmails': 'false'
   });
}


/**
 * Loads the config.json identified by the confId.
 * File should be a proper json file for the loading to 
 * be successful. Otherwise error will be thrown.
 */
function getConfig(configId) {
  var confFile = DriveApp.getFileById(configId);
  var config = JSON.parse(confFile.getAs('application/json').getDataAsString('UTF-8'));
  return config;
}

/**
 * Loads the config.json identified by the confId.
 * File should be a proper json file for the loading to 
 * be successful. Otherwise error will be thrown.
 * Along with the configuration in the config.json,
 * some additional status related stuffs also performed. 
 */
function getStatusConfig(configId) {
  var config = getConfig(configId);
  config.today = config.date ? new Date(config.date) : new Date();
  config.store = new Object();
  return config;
}

