////This code will take information from the CAD Charlie Board, create an Event in Google Calendar (WHOS???), and send an email via Gmail to the customer, BCCing the Charlie and Salesperson///

//Global Variables
  //Sheet
  formattedDate = Utilities.formatDate(new Date, "GMT", "MM-dd-yyyy");
  sheets = SpreadsheetApp.getActiveSpreadsheet();
  sheet = SpreadsheetApp.getActiveSheet();
  cellRow = sheet.getActiveCell().getRow();
  cellCol = sheet.getActiveCell().getColumn();
  cellURL = "https://docs.google.com/spreadsheets/d/[SPREADSHEET KEY]/edit#gid=0&range="+cellRow+":"+cellRow;
  startRow = cellRow;
  //Data
  messageDesc = sheet.getRange(cellRow, 27).getValue();
  messagePhase = messageDesc.substr(5,5);
  emailAddress = sheet.getRange(cellRow, 26).getValue();
  emailSent = sheet.getRange(cellRow, 28).getValue();
  emailReplyTo = sheet.getRange(cellRow, 29).getValue();
  emailSalesBCC = sheet.getRange(cellRow, 30).getValue();
  eventIdRetrieve = sheet.getRange(cellRow, 25).getValue();
  eventColour = 11
  

//This creates the Script Menu at the top of the Google Sheet
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Script Menu')
      .addItem('Update Calendar', 'calendarUpdate')
      .addToUi();
}

//This creates an Event in Google Calendar (WHOS???)
function calendarUpdate() {
  //Gets sheet and range details
  //var sheets = SpreadsheetApp.getActiveSpreadsheet();
  //var sheet = SpreadsheetApp.getActiveSheet();
  //var cellRow = sheet.getActiveCell().getRow();
  //var cellCol = sheet.getActiveCell().getColumn();
  //var cellURL = "https://docs.google.com/spreadsheets/d/[SPREADSHEET KEY]/edit#gid=0&range="+cellRow+":"+cellRow;
  
  //Gets SOR details for the Event and checks if it already exists
  var eventIdRetrieve = sheets.getRangeByName("eventID").getCell(cellRow, 1).getValue(); 
  var sor = sheets.getRangeByName("sor").getCell(cellRow, 1).getValue();
  var customer = sheets.getRangeByName("customer").getCell(cellRow, 1).getValue();
  var jobType = sheets.getRangeByName("jobType").getCell(cellRow, 1).getValue();
  var charlie = sheets.getRangeByName("charlie").getCell(cellRow, 1).getValue();
  var noRacks = sheets.getRangeByName("noRacks").getCell(cellRow, 1).getValue();
  var desc = sheets.getRangeByName("desc").getCell(cellRow, 1).getValue();
  //Creates UI
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
     'Please confirm',
     'Do you want to update the calendar with the selected job?',
      ui.ButtonSet.YES_NO);
  
  //If the event has no ID
  if (eventIdRetrieve == ""){
    //Creates Event variables
    if (result == ui.Button.YES) {
      var targetDate = sheets.getRangeByName("targetDate").getCell(cellRow, 1).getValue();
      var startDate = new Date(targetDate.getTime() + 28800000);
      var endDate = new Date(startDate.getTime() + 3600000);
      var targetDateF = Utilities.formatDate(new Date(startDate), "GMT", "dd-MM-yyyy");   
    //Used for naming
    if(sheets.getRangeByName("nonStock").getCell(cellRow, 1).getValue() == 1){
      var nonStock = "Yes"
      }else{
        var nonStock = "No"
        }
    
    if(jobType == "Tube Build"){
      var jobType = "TB";
    }
      else if (jobType == "Alu Build"){
        var jobType = "AB"
        }
      else if (jobType == "Batch SOR"){
        var jobType = "TB"
        }
      else if (jobType == "Duraflow Kit"){
        var jobType = "DK"
        }
      else if (jobType == "FOC Freebie"){
        var jobType = "TB"
        }
      else if (jobType == "LS2 Build"){
        var jobType = "LB"
        }
      else if (jobType == "LS2 Kit"){
        var jobType = "LK"
        }
      else if (jobType == "Stage Build"){
        var jobType = "SB"
        }
      else if (jobType == "Tube Kit"){
        var jobType = "TK"
        }
    else{
      var mistake = ui.alert(
        "That ain't no build!",
        "What you talking 'bout Willis?!?!\nDo you want to cancel?",
      ui.ButtonSet.YES_NO);
    }
    
    if(mistake == ui.Button.NO || mistake == undefined){
      ui.alert('Target Date: ' + targetDateF + '\n' + 'Event Title: [NC]' + sor + ' - ' + customer + ' - ' + noRacks + ' - ' + charlie + ' [' + jobType + ']\nDescription: ' + desc + ' (Non Stock: ' + nonStock + ')');
      calendarInput(cellURL,sor,customer,noRacks,charlie,jobType,desc,nonStock, targetDate);
    }
    
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Come back when you are ready');
  }
  }
  else {
    //Sends email if the job is already in the calendar
    var title = '[NC] ' + sor + ' - ' + customer + ' - ' + noRacks + ' - ' + charlie + ' [' + jobType + ']';
    var resultEmail = ui.alert('This job is already on the calendar\nDo you want to send an email and/or event update?', ui.ButtonSet.YES_NO);
    var description = desc + ' (Non Stock: ' + nonStock + ')';
    var emailSubject = sor + ' - ' + desc + ' - ' + noRacks;
    if (resultEmail == ui.Button.YES) {
      sendEmails2(cellRow, emailSubject);
    }
  }
}

//Function actually creates the Event
function calendarInput(cellURL, sor,customer,noRacks,charlie,jobType,desc,nonStock, targetDate, startDate, endDate, emailSubject){
  var title = '[NC] ' + sor + ' - ' + customer + ' - ' + noRacks + ' - ' + charlie + ' [' + jobType + ']';
  var description = desc + ' (Non Stock: ' + nonStock + ')' + '\n\n' + cellURL + '\n';
  var startDate = new Date(targetDate.getTime() + 28800000);
  var endDate = new Date(startDate.getTime() + 3600000);
  var event = CalendarApp.getCalendarById('[CALENDAR ID]').createEvent(title,
     new Date(startDate),
     new Date(endDate),
     {description: description,
      guests: emailSalesBCC});
  var eventID = event.getId();
  var emailSubject = sor + ' - ' + desc + ' - ' + noRacks;
  returnID(eventID, emailSubject);
}

function returnID(eventID, emailSubject, sheets, cellRow){
  //lists calendar event ID in spreadsheet, stops duplication
  //var sheets = SpreadsheetApp.getActiveSpreadsheet();
  //var sheet = SpreadsheetApp.getActiveSheet();
  //var cellRow = sheet.getActiveCell().getRow();
  
  var eventIdReturn = sheets.getRangeByName("eventID").getCell(cellRow, 1).setValue(eventID);
  
  //send email to customer - first time Calendar event is added
  //sendEmails2(cellRow, emailSubject);
}

//Updates Event with time stamped description
function calendarEventUpdate(calendarUpdateYES){
  var getDesc = CalendarApp.getCalendarById('[CALENDAR ID]@group.calendar.google.com').getEventSeriesById(eventIdRetrieve).getDescription();
  var setDesc = getDesc + "\n" + formattedDate + " - Status: " + messageDesc;
  var getTitle = CalendarApp.getCalendarById('[SPREADSHEET KEY]@group.calendar.google.com').getEventSeriesById(eventIdRetrieve).getTitle();
  
  if (messagePhase == 'Produ'){
    messagePhase = '[Prod]'
    eventColour = 10
    }else if (messagePhase == 'arget'){
    messagePhase = '[Admin]'
    eventColour = 3
    }
    else if (messagePhase == 'elive'){
    messagePhase = '[CAD]'
    eventColour = 6
    }
    else if (messagePhase == 'Quali'){
    messagePhase = '[Qual]'
    eventColour = 9
    }
    else if (messagePhase == 'ransp'){
    messagePhase = '[Tran]'
    eventColour = 10
    }
    
   var setTitle = messagePhase + getTitle;
  
    CalendarApp.getCalendarById('[SPREADSHEET KEY]@group.calendar.google.com').getEventSeriesById(eventIdRetrieve).setDescription(setDesc); //Sets new description
    CalendarApp.getCalendarById('[SPREADSHEET KEY]@group.calendar.google.com').getEventSeriesById(eventIdRetrieve).setColor(eventColour); //Sets new colour
    CalendarApp.getCalendarById('[SPREADSHEET KEY]@group.calendar.google.com').getEventSeriesById(eventIdRetrieve).setTitle(setTitle) //Sets new title
}

// This constant is written in column Z for rows for which an email
// has been sent successfully.
var EMAIL_SENT = "EMAIL_SENT";

function sendEmails2(cellRow, emailSubject) {
  //var sheet = SpreadsheetApp.getActiveSheet();
  //var startRow = cellRow;  // First row of data to process
  
  var numRows = 1;   // Number of rows to process
  // Fetch the range of cells Z:AD
  var dataRange = sheet.getRange(startRow, 26, numRows, 30)
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  var ui = SpreadsheetApp.getUi();
  
  //gathers info and sends email
  //for (var i = 0; i < data.length; ++i) {
    //var row = data[i];
    //var emailAddress = row[0];  //First column
    //var message = "This is a brief update to let you know that your order has been sent " + row[1];       // Second column
    //var emailSent = row[2];     // Third column
    //var emailReplyTo = row[3];    //Fourth column
    //var emailSalesBCC = row[4];     //Fifth column
    var message = "This is a brief update to let you know that your order has been sent " + messageDesc;
    if (emailSent != EMAIL_SENT) {  // Prevents sending duplicates
      var subject = emailSubject;
      MailApp.sendEmail(emailAddress, subject, message, {bcc:emailReplyTo + "," + emailSalesBCC, replyTo:emailReplyTo});
      sheet.getRange(startRow, 28).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
      ui.alert('Your email has been sent to ' + emailAddress + "\nBCC:\n" + emailReplyTo + "\n" + emailSalesBCC);
      calendarEventUpdate();
    }else{
      ui.alert('The EMAIL_SENT marker is full. If you wish to send an email update, please clear EMAIL_SENT and retry');
      calendarEventUpdate();
  }
//}
}
