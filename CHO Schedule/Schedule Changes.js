//This function changes the formatting of completed schedule changes and notifies stakeholders 
function markAsResolved() {
 
  //add heuristics, like make sure it's a good email address, make sure that it isn't already green, maybe a way to re-open the request?
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveRange().getRow();
  var recipient = sheet.getRange(row, 2).getValue();
  var supervisorCol = 10;

  var numCols = sheet.getLastColumn();
  var requestRange = sheet.getRange(row, 1, 1, numCols);
  
  // Displays a dialog box with a message, input field, and "Yes" and "No" buttons. The user can
  // also close the dialog by clicking the close button in its title bar.
  var response = ui.prompt('Who approved this? (Select \"No\" if request not approved or cancelled)', ui.ButtonSet.YES_NO);
  
  // Process the user's response.
  if (response.getSelectedButton() == ui.Button.YES) {
    sheet.getRange(row, supervisorCol).setValue(response.getResponseText());
    requestRange.setBackground("#00ff00"); //green 
  } else if (response.getSelectedButton() == ui.Button.NO) {
    requestRange.setBackground("#ff0000"); //red
  } else {
    Logger.log('The user clicked the close button in the dialog\'s title bar.');
    return;
  }
  
  if (recipient.length <= 0) return; 
  
  var time = sheet.getRange(row, 1).getValue();
  
  var minutes = time.getMinutes().toString();
  var seconds = time.getSeconds().toString();
  if (time.getMinutes() < 10) minutes = "0" + minutes;
  if (time.getSeconds() < 10) seconds = "0" + seconds;
  
  var requestCode = (time.getMonth()+1).toString() + "/" + time.getDate().toString() + 
    "/" + time.getYear().toString() + " " + time.getHours().toString() + ":" + 
      minutes + ":" + seconds;   
  
  var subject = "Resolved: CHO " + sheet.getRange(row, 3).getValue() + "-" + sheet.getRange(row, 4).getValue() + " schedule change (" + requestCode + ")";   
  var message = "Your schedule change request has been processed.";
  
  
  MailApp.sendEmail(recipient, subject, message); 
  
  requestRange.setWrap(false);
  Logger.log("completed markAsResolved");
}


//This function formats new schedule change requests. An accompanying trigger must be created to ensure it runs
function formatNewRequest() {
  Logger.log("got called!");
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("Schedule Changes");
  
  var requestRow = sheet.getLastRow();

  var newRequestRange = sheet.getRange(requestRow, 1, 1, sheet.getLastColumn());
  
  newRequestRange.setBackground("#ff9900");
  var userRange = sheet.getRange(requestRow, 2);
  userRange.setFormula("=HYPERLINK(\"" + userRange.getValue() + "\")");
  newRequestRange.setWrap(true);
  newRequestRange.setVerticalAlignment("top");
  Logger.log("formatted correctly!"); 
}
