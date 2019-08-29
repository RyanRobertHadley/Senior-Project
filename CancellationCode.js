var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheets()[0];
var reservation_Cal = CalendarApp.getCalendarById("m36fuue0q2o9nh5qhlk1epcurs@group.calendar.google.com");
var lastRow = sheet.getLastRow();
var lastColumn = sheet.getLastColumn();
var HTMLString= "";

//Purpose: Retrieve the event id from the spreadsheet, determine if the event requested exists.
//Assumptions: The form has been submitted, and the last field in column 2 represents the latest submission.
//Pre-Condtitions: The event id submitted will be stored in the form, no HTML page has been chosen.
//Post-Conditions: The event id will be identified, and the appropriate HTML page will be chosen.
//Parameters: NA
function doGet(){
  var eventID = sheet.getRange(lastRow, 2).getValue();
  var event = reservation_Cal.getEventById(eventID); 

  if(event == null){
    HTML_Page = HtmlService.createTemplateFromFile('Cancellation_Rejection').evaluate();   
  }
  else{   
    HTML_Page = HtmlService.createTemplateFromFile('Cancellation_Confirmation').evaluate(); 
  }  
  return HTML_Page
}

//Purpose: Removes the event from the reservation calendar according to the event id submitted by the user.
//Assumptions: The Cancellation_Confirmation HTML page has been selected, and the cancellation has been confirmed in the doGet() function.
//Pre-Condtitions: Event will exist in the reservation calendar.
//Post-Conditions: Event will be removed from the reservation calendar.
//Parameters: The events Id as it was taken from the form submission, and the event that is to be removed.
function deleteEvent(eventID, event){
  var deleted = false;
  if(event != null){
    reservation_Cal.getEventById(eventID).deleteEvent();
    deleted = true;
  }
  return deleted;
}

//Purpose: Create a message for the user by assigning a specific value to the HTMLString variable, which will be displayed.
//Assumptions: The doGet() function, and the HTML page has been chosen.
//Pre-Condtitions: The user has been shown an HTML page.
//Post-Conditions: The status of the cancellation will be shown.
//Parameters: NA
function createOutput(){
  var eventID = sheet.getRange(lastRow, 2).getValue();
  var event = reservation_Cal.getEventById(eventID); 
  deleted = deleteEvent(eventID, event);
  
  if(deleted == false){   
    HTMLString = "The reservation does not exist or it has already been cancelled. Please check the calendar.";    
  }  
  else{
    HTMLString = "Your reservation has been cancelled."
  }  
  return HTMLString;
}
  