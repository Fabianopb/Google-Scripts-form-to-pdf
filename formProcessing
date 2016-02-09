//*** Verify if there is a trigger "onFormSubmit" set up in Resources > Current project's trigger ***//

//Global variables
var SOURCE_TEMPLATE = "1RdoqaMSqCRMYh4oUpTArXDX4jpmnqVnu-_QoNOpGvGY"; var CUSTOMER_SPREADSHEET = "15Fk6GWBcvJCe59oFXd7YOLTnmY6m0jJ6G1dpD1sM-xI"; var TARGET_FOLDER = "0B4l6bSAKJm8db25DaWJCS2xtdm8";

var WILD_CARD_BASIC_INFO = ["#NAME", "#BIRTH", "#EMAIL", "#PHONE", "#SSN"];
var WILD_CARD_TRIP_INFO = ["#DESTINATION", "#DEPARTURE", "#ARRIVAL", "#PURPOSE"]
var WILD_CARD_TRANSP_TYPE = ["#T1", "#T2", "#T3", "#T4", "#T5", "#T6"];
var WILD_CARD_TICKETS_ALLOWANCE = ["#PROJECT", "#TICKETS_DESCRIPTION", "#TICKETS_PRICE", "#ALLOWANCE_DAYS", "#DAILY_ALLOWANCE"];
var WILD_CARD_ACCOMMODATION = ["#ACCOMMODATION", "#A1", "#A2"];
var WILD_CARD_PARTICIPATION = ["#PARTICIPATION", "#P1", "#P2"];
var WILD_CARD_OTHER = ["#OTHER_COSTS", "#OTHER_TOTAL"];
var WILD_CARD_CALCULATED = ["#ALLOWANCE_TOTAL", "#TOTAL_COSTS"];

var TRANSP_TYPE_ARRAY = ["train", "plane", "bus", "university car", "own car", "other"];
var KEYWORD = "directly"

//*** THIS IS THE MAIN FUNCTION ***//
function generatesTravelPlan() {
  
  //Connect to the spreadsheet, get the last row and retrieve the data
  var sheet = SpreadsheetApp.openById(CUSTOMER_SPREADSHEET).getSheets()[0];
  var lastRow = sheet.getLastRow();
  var timeStamp = sheet.getRange(lastRow, 1).getValues();
  var basicInfo = sheet.getRange(lastRow, 2, 1, WILD_CARD_BASIC_INFO.length).getValues();
  var tripInfo  = sheet.getRange(lastRow, 7, 1, WILD_CARD_TRIP_INFO.length).getValues();
  var transpInfo = sheet.getRange(lastRow, 11).getValues();
  var ticketsAllowanceInfo = sheet.getRange(lastRow, 12, 1, WILD_CARD_TICKETS_ALLOWANCE.length).getValues();
  var accommodationInfo = sheet.getRange(lastRow, 17, 1, 2).getValues();
  var participationInfo = sheet.getRange(lastRow, 19, 1, 2).getValues();
  var otherInfo  = sheet.getRange(lastRow, 20, 1, WILD_CARD_OTHER.length).getValues();
  
  //Transform timestamps to shorter formats
  timeStamp = timeStamp.toString().substring(4, 23);
  basicInfo[0][1] = basicInfo[0][1].toString().substring(4, 15);
  tripInfo[0][1] = tripInfo[0][1].toString().substring(0, 15);
  tripInfo[0][2] = tripInfo[0][2].toString().substring(0, 15);
  
  //Calculate the total daily allowance and the total costs
  var totalDailyAllowance = ticketsAllowanceInfo[0][3] * ticketsAllowanceInfo[0][4];
  var totalCosts = ticketsAllowanceInfo[0][2] + totalDailyAllowance + accommodationInfo[0][0] + participationInfo[0][0] + otherInfo[0][0];
  
  //Create a new document based on the template
  var template = DriveApp.getFileById(SOURCE_TEMPLATE);
  var newFile = template.makeCopy(basicInfo[0][0] + " " + timeStamp, DriveApp.getFolderById(TARGET_FOLDER));
  
  //Get Id of the new document, open it and get its body
  var targetId = newFile.getId();
  var targetDoc = DocumentApp.openById(targetId);
  var targetBody = targetDoc.getBody();
  
  //Replace the wild cards in the new document with the user data: BASIC INFO
  for(var i = 0; i < WILD_CARD_BASIC_INFO.length; i++) {
    targetBody.replaceText(WILD_CARD_BASIC_INFO[i], basicInfo[0][i]);
  }
  
  //Replace the wild cards in the new document with the user data: TRIP INFO
  for(var i = 0; i < WILD_CARD_TRIP_INFO.length; i++) {
    targetBody.replaceText(WILD_CARD_TRIP_INFO[i], tripInfo[0][i]);
  }
  
  //Replace the wild cards in the new document with the user data: TRANSPORTATION INFO
  for(var i = 0; i < WILD_CARD_TRANSP_TYPE.length; i++) {
    if(transpInfo.toString().indexOf(TRANSP_TYPE_ARRAY[i]) > -1) {
      targetBody.replaceText(WILD_CARD_TRANSP_TYPE[i], "X");
    }
    else {
      targetBody.replaceText(WILD_CARD_TRANSP_TYPE[i], "");
    }
  }
  
  //Replace the wild cards in the new document with the user data: TICKETS AND ALLOWANCE INFO
  for(var i = 0; i < WILD_CARD_TICKETS_ALLOWANCE.length; i++) {
    targetBody.replaceText(WILD_CARD_TICKETS_ALLOWANCE[i], ticketsAllowanceInfo[0][i]);
  }
  
  //Replace the wild cards in the new document with the user data: ACCOMMODATION INFO
  targetBody.replaceText(WILD_CARD_ACCOMMODATION[0], accommodationInfo[0][0]);
  if(accommodationInfo[0][1].toString() == "") {
    targetBody.replaceText(WILD_CARD_ACCOMMODATION[1], "");
    targetBody.replaceText(WILD_CARD_ACCOMMODATION[2], "");
  }
  else if(accommodationInfo[0][1].indexOf(KEYWORD) > -1) {
    targetBody.replaceText(WILD_CARD_ACCOMMODATION[1], "X");
    targetBody.replaceText(WILD_CARD_ACCOMMODATION[2], "");
  }
  else {
    targetBody.replaceText(WILD_CARD_ACCOMMODATION[1], "");
    targetBody.replaceText(WILD_CARD_ACCOMMODATION[2], "X");
  }
  
  //Replace the wild cards in the new document with the user data: PARTICIPATION INFO
  targetBody.replaceText(WILD_CARD_PARTICIPATION[0], participationInfo[0][0]);
  if(participationInfo[0][1].toString() == "") {
    targetBody.replaceText(WILD_CARD_PARTICIPATION[1], "");
    targetBody.replaceText(WILD_CARD_PARTICIPATION[2], "");
  }
  if(participationInfo[0][1].indexOf(KEYWORD) > -1) {
    targetBody.replaceText(WILD_CARD_PARTICIPATION[1], "X");
    targetBody.replaceText(WILD_CARD_PARTICIPATION[2], "");
  }
  else {
    targetBody.replaceText(WILD_CARD_PARTICIPATION[1], "");
    targetBody.replaceText(WILD_CARD_PARTICIPATION[2], "X");
  }
  
  //Replace the wild cards in the new document with the user data: OTHER COSTS INFO
  for(var i = 0; i < WILD_CARD_OTHER.length; i++) {
    targetBody.replaceText(WILD_CARD_OTHER[i], otherInfo[0][i]);
  }
  
  //Replace the wild cards in the new document with the calculated data: TOTAL DAILY ALLOWANCE and TOTAL COSTS
  targetBody.replaceText(WILD_CARD_CALCULATED[0], totalDailyAllowance);
  targetBody.replaceText(WILD_CARD_CALCULATED[1], totalCosts);
  
  //Save and close so the pdf can be generated based on the last version
  targetDoc.saveAndClose();
  
  //Generates pdf and send to the user's e-mail
  var pdfDocument = DriveApp.getFileById(targetId).getAs("application/pdf");
  MailApp.sendEmail(basicInfo[0][2], "Your travel plan", "Hello, here is your travel plan! Please print, sign and return it to the IDBM office. You can also scan and send it via e-mail (please use a scanner or scanning app, do not take a picture).\n\nWe will try to get back to you in no more than 2 working days. Do not reply to this message, if you have questions contact directly the IDBM office.\n\nCheers,\nIDBM", {attachments: pdfDocument});
  MailApp.sendEmail("fabiano.brito@aalto.fi, anna-mari.saari@aalto.fi", "IDBM new travel plan", '', { htmlBody: basicInfo[0][0].toString() + ' has just created a travel plan! You can check it from <a href="https://drive.google.com/drive/u/0/folders/0B4l6bSAKJm8db25DaWJCS2xtdm8">here</a>.' });
  
  //Delete the information from the spreadsheet
  sheet.deleteRow(lastRow);
  
}
