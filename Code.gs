function pushToCalendar() {
  //User settings
  var calId = "byronscyoung@gmail.com";
  var reminder = 30; //Set reminder time to 30 min
  var sendInvitation = true;

  if (getUser() != calId) {
    var html = HtmlService.createHtmlOutputFromFile('cookie').setWidth(500).setHeight(500);
    SpreadsheetApp.getUi().alert("Sorry! You do not have permission to Sync the Calendar.");
    SpreadsheetApp.getUi().showModalDialog(html, 'Here\'s a cookie instead.');
    return 0;
  }

  //Spreadsheet Variables
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assignments');
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(3, 2, lastRow, 8);
  var headrows = 0;
  var data = range.getValues();

  //Calendar Variables
  var cal = CalendarApp.getCalendarById(calId);
  var calEvents = cal.getEvents(new Date('1/1/1970'), new Date('1/1/2030'));

//  //Create Calendar Events
  for (var i = 0; i < data.length-2; i++) {
    if (i < headrows) continue;

    var row = data[i];
    var date = Date(row[0]);
    var title = row[1];
    var tstart = new Date(row[0]);
    var tstop = new Date(row[0]);

    //Set Starting Hours from two cells
    tstart.setHours(new Date(row[4]).getHours());
    tstart.setMinutes(new Date(row[4]).getMinutes());
    tstart.setSeconds(new Date(row[4]).getSeconds());

    //Set Ending Hours
    if(row[5] != ""){
      tstop.setHours(new Date(row[5]).getHours());
      tstop.setMinutes(new Date(row[5]).getMinutes());
      tstop.setSeconds(new Date(row[5]).getSeconds());
    }
    else { //Add one hour to original start time if there is no end time set
      tstop = new Date(tstart.getTime() + 60*60000);
    }

    var loc = row[2];
    var desc = row[3];
    var id = row[7];

    // Check if event already exists
    try {
      var event = getEvent(calEvents, row[7]);
    }
    catch (e) {
      // do nothing - we just want to avoid the exception when event doesn't exist
    }
    if(!event){
      var newEvent = cal.createEvent(title, tstart, tstop,{description: desc, sendInvites: sendInvitation, location: loc});
      newEvent.addPopupReminder(reminder);
      inviteGuests(newEvent, row[6]);
      row[7] = newEvent.getId();
    }
    else {
      //If event title, description, or start time and end time is different then update it
      if(event.getTitle() != title || event.getDescription() != desc || event.getStartTime() != tstart || calEvents[0].getEndTime() != tstop || event.getLocation() != loc){
        event.setTitle(title);
        event.setDescription(desc);
        event.setTime(tstart, tstop);
        event.setLocation(loc);
      }
    }
    debugger;
  }
  // Record all event IDs to spreadsheet
  range.setValues(data);
}

//Events are created depending on which user is running the script
function getUser(){
  return Session.getActiveUser().getEmail();
}

function getEvent(calEventsId, rowId){
  for (var i = 0; i < calEventsId.length; i++){
    if (calEventsId[i].getId() == rowId) {
      return calEventsId[i];
    }
  }
  return null;
}

function inviteGuests(event, rowData){
  var guestSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Participants');
  var guestLastRow = guestSheet.getLastRow();
  var guestRange = guestSheet.getDataRange();
  var guestData = guestRange.getValues();
  var guestHeader = 2;

  names = rowData.split(', ');

  if (names[0].equals("Everyone") == true) {
    for (var i = 0; i < guestData.length; i++) {
      if (i < guestHeader) continue;
      event.addGuest(guestData[i][1]);
    }
  }
  else {
    var j = 0;

    for (var i = 0; i < guestData.length; i++) {
      if (i < guestHeader) continue;
      if(names[j] == guestData[i][0]){
        event.addGuest(guestData[i][1]);
        j++;
      }
    }
  }
}

//Fill Data Validation With multiple selections
function chooseParticipants() {
  var html = HtmlService.createTemplateFromFile('dialog').evaluate();
  SpreadsheetApp.getUi().showModalDialog(html, "Choose Participants");
}
var valid = function(){
  try{
    return SpreadsheetApp.getActiveRange().getDataValidation().getCriteriaValues()[0].getValues();
  }catch(e){
    return null
  }
}
function fillCell(e){
  var s = [];
  for(var i in e){
    if(i.substr(0, 2) == 'ch') s.push(e[i]);
  }
  if(s.length) SpreadsheetApp.getActiveRange().setValue(s.join(', '));
}

//On Open Add Menu
function onOpen() {
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  menuEntries.push({name: "Sync Calendar", functionName: "pushToCalendar"});
  menuEntries.push({name: "Choose Participants", functionName: "chooseParticipants"});
  activeSheet.addMenu("Starlord", menuEntries);
}
