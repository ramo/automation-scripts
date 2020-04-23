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
function sendStatusEmail() {
  var configId = '<configId>';  // Caution: This need to be set before running the script.
  var config = getStatusConfig(configId);

  var preprocess = (config) => {
    var doc = DocumentApp.openById(config.mailDocId);  
    var reportPdf = doc.getBlob().copyBlob();
    var mdy = splitDateToMDYArray(config.today);
    var reportName = config.reportPrefix + '-' + mdy.join('-') + '.pdf';  
    reportPdf.setName(reportName);
    config.store.attachments = [reportPdf];
  };

  var postprocess = (config) => {
    takeBackup(config.mailDocId);
    resetDocument(config);
  };

  sendEmail(config, preprocess, postprocess);
}


function resetDocument(config) {
  clearDocument(config.mailDocId);
  var doc = DocumentApp.openById(config.mailDocId);  
  var docBody = doc.getBody();
  var mdy = splitDateToMDYArray(nextWorkingDay(config.today, config.holidays));
  var docHeader = config.mail.subjectPrefix + ' - ' + mdy.join('/');
  docBody.insertParagraph(0, docHeader)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .editAsText().setBold(true);
  docBody.appendParagraph('\n\n--- Remove me and start writing your status from here. ---');
  doc.saveAndClose();
}







