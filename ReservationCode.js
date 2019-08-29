var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheets()[0];
var lastRow = sheet.getLastRow();
var lastColumn = sheet.getLastColumn();
var events_Cal = CalendarApp.getCalendarById("3857iklcc8sfojr3ee7d87kp40@group.calendar.google.com");
var reservation_Cal = CalendarApp.getCalendarById("m36fuue0q2o9nh5qhlk1epcurs@group.calendar.google.com");
var HTMLString="";
  
//Purpose: Retrieve the last submitted form information.
//Assumptions: There is a form that has been submitted, and it's information is present at the bottom row of the Google Spreadsheet.
//Pre-Condtitions: The form information regarding a reservation is present in the spreadsheet, but has yet to be collected by the app.
//Post-Conditions: The form information will be collected from the spreadsheet.
//Parameters: NA
function getSubmission(){
  this.timestamp = sheet.getRange(lastRow, 1).getValue();
  this.name = sheet.getRange(lastRow, 2).getValue();
  this.date = sheet.getRange(lastRow, 3).getValue();
  this.time = sheet.getRange(lastRow, 4).getValue();
  this.date.setHours(this.time.getHours());
  this.date.setMinutes(this.time.getMinutes());
  //This is significant in that the advising meetings will be restricted to one hour intervals.
  this.standardEndTime = new Date(this.date);
  this.standardEndTime.setHours(this.time.getHours() + 1);
  return this
}

//Purpose: Call the isOpen(), isRoomAvailable(),notWeekend(), and isAvailable() functions to determine what conflicts exist with the requested reservation.
//Assumptions: NA
//Pre-Condtitions: The status of the reservation is not yet determined.
//Post-Conditions: The status of the reservation will be determined to be conflict-free or unavailable. 
//Parameters: NA
function checkAllConflicts(){
  var request = getSubmission();
  var isRoomAvailable = true;
  if(isOpen(request, isRoomAvailable) && notWeekend(request, isRoomAvailable) &&notPast(request,isRoomAvailable) && isAvailable(request,isRoomAvailable)){
    isRoomAvailable = true;
  }
  else{
    isRoomAvailable = false;
  }
  //Logger.log(sheet.getRange(lastRow, 5).getValue());
  return isRoomAvailable;
}

//Purpose: Determine whether or not the requested time slot of the form submission adheres to the hours that the facility is open.
//Assumptions: NA
//Pre-Condtitions: Status of the requested time interval is not yet determined.
//Post-Conditions: Status of the requested time interval will either be confirmed or rejected by the returned boolean.
//Parameters: request - variable representing the current results of getSubmission,the most recent form data. isOpen - boolean that will represent the results of the time interval query.
function isOpen(request, isOpen){
  var time = ((request.date).toTimeString()).slice(0,5); //returns the full time, which we splice (09:00:00 -> 09:00)
  var hours = Number(time.slice(0,2));
  var date = (request.date).toDateString();
  var day = date.substring(0,3);
  //the center is only open from 9-5
  if(day === "Fri"){
    if(hours < 9.0 || hours > 15){
      isOpen = false;
      sheet.getRange(lastRow, 8).setValue('closed');
    }
  }
  else if(day === "Thur"){
    if(hours < 9.0 || hours > 17) {
      isOpen = false;
      sheet.getRange(lastRow, 8).setValue('closed');
    }
  }
  else if(hours < 9.0 || hours > 16) {
    isOpen = false;
    sheet.getRange(lastRow, 8).setValue('closed');    
  }
  //Logger.log(isOpen);
  return isOpen;
}

//Purpose: Determine whether or not the requested day value of the form submission adheres to the days that the facility is open.
//Assumptions: NA
//Pre-Condtitions: Status of the requested day is not yet determined.
//Post-Conditions: Status of the requested day will either be confirmed or rejected by the returned boolean.
//Parameters: request - variable representing the current results of getSubmission,the most recent form data. notWeekend - boolean that will represent the results of the available day query.
function notWeekend(request,notWeekend){
  var date = (request.date).toDateString();
  var day = date.substring(0,3);
  if(day === "Sun" || day === "Sat"){
    notWeekend = false;
    sheet.getRange(lastRow, 8).setValue('weekend');
  }
  return notWeekend;
}

//Purpose: Determine whether or not the requested day value of the form submission adheres to the days that the facility is open.
//Assumptions: NA
//Pre-Condtitions: Status of the requested day is not yet determined.
//Post-Conditions: Status of the requested day will either be confirmed or rejected by the returned boolean.
//Parameters: request - variable representing the current results of getSubmission,the most recent form data. notWeekend - boolean that will represent the results of the available day query.
function notPast(request,notPast){
  var currentDate = new Date();
  var requestDate = request.date;
  if ( currentDate.valueOf() > requestDate.valueOf()){
    notPast = false;
    sheet.getRange(lastRow, 8).setValue('past');
  }
  return notPast;
}

//Purpose: Determine whether there is a room that is available for the time and day requested by the user.
//Assumptions: The value for the location of the reservations is stored in the 5th column of the spreadsheet data.
//Pre-Condtitions: Status of the requested reservation is not yet determined.
//Post-Conditions: Status of the requested reservation will either be confirmed or rejected by the returned boolean. In addition, the value for the reserved room will be stored for the most recent submission.
//Parameters: request - variable representing the current results of getSubmission,the most recent form data. isRoomAvailable - boolean that will represent the results of the available room query.
function isAvailable(request,isRoomAvailable){
  var event_Conflicts = events_Cal.getEvents(request.date, request.standardEndTime); 
  var trinity1_Conflicts = reservation_Cal.getEvents(request.date, request.standardEndTime, {search: 'Trinity1'});
  var trinity2_Conflicts = reservation_Cal.getEvents(request.date, request.standardEndTime, {search: 'Trinity2'});
  var ignatius1_Conflicts = reservation_Cal.getEvents(request.date, request.standardEndTime, {search: 'Ignatius1'});
  var resurrection1_Conflicts = reservation_Cal.getEvents(request.date, request.standardEndTime, {search: 'Resurrection1'});
  
  if (trinity1_Conflicts.length < 1){
    sheet.getRange(lastRow, 5).setValue('Trinity1');
  } 
  else if (trinity2_Conflicts.length < 1){
    sheet.getRange(lastRow, 5).setValue('Trinity2');
  }    
  else if (ignatius1_Conflicts.length < 1){
    sheet.getRange(lastRow, 5).setValue('Ignatius1');
  }    
  else if(event_Conflicts.length < 1 && resurrection1_Conflicts.length < 1){
      sheet.getRange(lastRow, 5).setValue('Resurrection1');
  }
  else{
    isRoomAvailable = false;
    sheet.getRange(lastRow, 8).setValue('full');
  }
  Logger.log(isRoomAvailable);
  return isRoomAvailable;
}

//Purpose: Sends the updateCalendar() function the information from the last row of the spreadsheet, after it is confirmed that there are no conflict. Also stores the event id in the spreadsheet in case of a cancellation
//Assumptions: There is a form that has been submitted, and it's information is present at the bottom row of the Google Spreadsheet. 
//Assumptions: The requested reservation has been checked for conflicts by the checkAllConflicts() function, and the room reserved determined by the isRoomAvailable() function.
//Pre-Condtitions: The data from the form submission has been verified, and is waiting to be delivered to the calendar.
//Post-Conditions: The data will be sent to the updateCalendar() function, as well as the new room value taken from the spreadsheet.
//Parameters: NA.
function reserveRoom(){
  var request = getSubmission(); 
  var roomValue = sheet.getRange(lastRow, 5).getValue();
  var event_id = updateCalendar(request, roomValue);
  sheet.getRange(lastRow, 6).setValue(event_id);
}

//Purpose: Updates the calendar with the data stored in the spreadsheet using the calendar.createEvent() function(pre-defined).
//Assumptions: The data from the spreadsheet have been verified to be the latest submission, and there are no conflicts.
//Pre-Condtitions: The data from the form submission has been recieved, the conflicts determined to be non-existent, and the room value saved to the spreadsheet. 
//Post-Conditions: The requested reservation will be stored in the calendar. This includes the date, time, room, and name of the user.
//Parameters: The results of the getSubmission function, this is also needed to be called again for a currently unkown reason.
function updateCalendar(request, location){
  var request = getSubmission();
  var roomReserve = reservation_Cal.createEvent(
    request.name,
    request.date,
    standardEndTime,
    {location: location});
  event_id = roomReserve.getId(); //put in google sheet
  event_id = event_id.substr(0, event_id.indexOf("@"));
  if(location == "Trinity1"){
    roomReserve.setColor(CalendarApp.EventColor.ORANGE);
  }
  if(location == "Trinity2"){
    roomReserve.setColor(CalendarApp.EventColor.YELLOW);
  }
  if(location == "Igantius1"){
    roomReserve.setColor(CalendarApp.EventColor.MAUVE);
  }
  return event_id;
}

//Purpose: Will call the reserveRoom() function only if the checkAllConflicts() function returns true.
//Assumptions: The checkAllConflicts() function will accurately determine the availability of the reservation request.
//Pre-Condtitions: Request has been submitted but has not been verified or reserved.
//Post-Conditions: Request will either be completed or rejected.
//Parameters: NA
function isConfirmed(){
  var request = getSubmission();
  var isRoomAvailable = checkAllConflicts();
  var reservationConfirmed = Boolean(false);
  if (isRoomAvailable){
    reserveRoom();
    reservationConfirmed = Boolean(true);
  }
  return reservationConfirmed;
}

//Purpose: Adjust the HTMLString variable according to the apropriate results of the reservation request the time value will be formatted by the createTime_Output function.
//Purpose: This function is ultimately used by one of the two HTML pages to display the proper information on the correct page.
//Assumptions: The request has been either confirmed or denied, and if confirmed, the time,date,and room have been stored in the spreadsheet.
//Pre-Condtitions: The HTMLString variable is a null value.
//Post-Conditions: The HTMLString variable will be set to the appropriate results of the reservation, and returned.
//Parameters: NA
function createOutput(){
  var request = getSubmission();
  var assignedRoom = sheet.getRange(lastRow, 5).getValue();
  var time = request.date.toLocaleTimeString();
  var date = (request.date).toDateString();               //returns the date
  var errorCode = sheet.getRange(lastRow, 8).getValue(); //weekend,full,past, or closed
  if (assignedRoom != ""){
    HTMLString = 
      "<br/><strong>Name: </strong>" + request.name +
        "<br/><strong>Date: </strong>" + date +
          "<br/><strong>Time: </strong>" + time +
            "<br/><strong>Location: </strong>" + assignedRoom +
              "<br/><strong>ID: </strong>" + sheet.getRange(lastRow, 6).getValue();
  }
  else {
    HTMLString ="There are no rooms available at this time and date." +
      "<br/><strong>Date: </strong>" + date +
        "<br/><strong>Time: </strong>" + time;
    if (errorCode == 'weekend'){
      HTMLString = HTMLString + '<p> A weekend date was entered.</p>';
    }
    else if (errorCode == 'full'){
      HTMLString = HTMLString + '<p> All three available rooms are full at this time. You may contact the center to see if an extra room is open.</p>';
    }
    else if (errorCode == 'past'){
      HTMLString = HTMLString + '<p> A date was entered that has already passed.</p>';
    }
    else if (errorCode == 'closed'){
      HTMLString = HTMLString + '<p> A time was entered after-hours. The center is open from: <br/> 9-5pm Mon-Wed.| 9-6pm Thurs.| 9-4pm Fri.</p>';
    }
  }
 // Logger.log(assignedRoom);
    return HTMLString;   
}

//Purpose: Adjust the HTMLString variable according to the apropriate results of the reservation request the time value will be formatted by the createTime_Output function.
//Assumptions: The request has been either confirmed or denied, and if confirmed, the time,date,and room have been stored in the spreadsheet.
//Pre-Condtitions: The time value for the requested date is formatted incorrectly for our user to view
//Post-Conditions: A new time value will be produced.
function createTime_Output(request){
  var time = ((request.date).toTimeString()).slice(0,5); //returns the full time, which we splice (09:00:00 -> 09:00)
  if(Number(time.slice(0,2)) < 12) {
    time = time + " AM";
  }
  else
  {
    var t = Number(time.slice(0,2))-12;
    time = String(t) + " PM";
    //time = time + " PM";
  }
  return time;
}

//Purpose: The primary trigger of the app. Calls the isConfirmed() function in order to determine the status of the reservation request, then displays an HTML page from one of the two pre-made templates.
//Assumptions: The functions that are called in the isConfirmed() function will produce accurate results.
//Pre-Condtitions: The web app has been loaded.
//Post-Conditions: This function will execute upon the app loading. The results of the reservation request will be displayed to the user from one of the HTML templates, and the HTMLString variable.
//doPost is a defined trigger in Google App Scripts
function doGet(){
  var reservationConfirmed = isConfirmed(); 
  if (reservationConfirmed == true){
    HTML_Page= HtmlService.createTemplateFromFile('Frontend_Confirmation').evaluate();
  }
  else if(reservationConfirmed == false){
    HTML_Page= HtmlService.createTemplateFromFile('Frontend_Rejection').evaluate();
  }
  //cleanSheet();
  return HTML_Page;
}
