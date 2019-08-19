/**
Creating the UI menu item for launching the script
**/

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Calendar')
      .addItem('Update Calendar', 'updateCalendar')
      .addToUi()
}

/**
The beatiful function for migrating the data from spreadsheet to calendar
**/

function updateCalendar() {
  
  
  /**
  Opening the calendar
  **/
  
  var calendarId = "one99.pl_mdesm69uckc4ig8br3n09m21is@group.calendar.google.com";
  var eventCal = CalendarApp.getCalendarById(calendarId);
  
  /**
  Pulling data from spreadsheet using the utils function
  **/
  
  function getLastDataRow(sheet) {
    var lastRow = sheet.getLastRow();
    var range = sheet.getRange("A" + lastRow);
    if (range.getValue() !== "") {
      return lastRow;
    } 
    else {
      return range.getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
    }              
  };
  var spreadsheet = SpreadsheetApp.getActiveSheet();
  var range = spreadsheet.getRange('A4:D'+ getLastDataRow(spreadsheet));
  var signups = range.getValues();
  var changes = [];
  
  /**
  A loop for creating events in Calendar for all people
  **/
  
  for (x=0; x<signups.length; x++) {
    var vacation = signups[x];
    var employee = vacation[0];
    var type = vacation[1];
    var startTime = vacation[2];
    var endTime = vacation[3];
    var record = employee + " (" + type + ")";
    endTime.setDate(endTime.getDate() + 1);
    /**
    Condition for coloring based on the type
    
    Coloring:
    Pale Blue ("1")
    Pale Green ("2")
    Mauve ("3")
    Pale Red ("4")
    Yellow ("5")
    Orange ("6")
    Cyan ("7")
    Gray ("8")
    Blue ("9")
    Green ("10")
    Red ("11")
    **/
    
    var eventColor = "9"
    if (type == "Urlop"){
      eventColor = "11"
    }
    
    /**
    Checking if there are any existing events for the employee within the given date range
    **/
    var eventCal = CalendarApp.getCalendarById("one99.pl_mdesm69uckc4ig8br3n09m21is@group.calendar.google.com");
    var alreadyExists = eventCal.getEvents(startTime, endTime, {search: employee})
    /**
    Creating new event
    **/
    if (alreadyExists.length <= 0) {
      var newEndTime = endTime
      eventCal.createAllDayEvent(record, startTime, endTime).setColor(eventColor);     
      newEndTime.setDate(endTime.getDate() - 1);
      /**
      Adding a record of the changes made to alert array
      **/
      if (startTime.getTime() === newEndTime.getTime()) {
        changes.push(record+ " " + startTime.toLocaleDateString("en-US"));
        }
      else {
          changes.push(record+ " " + startTime.toLocaleDateString("en-US")+ " - "+newEndTime.toLocaleDateString("en-US"));
        }
      }
    }
    /**
    Displaying the alert array with changes that have been made
    **/  
  var ui = SpreadsheetApp.getUi();
  if (changes.length>0){
    ui.alert("Following changes have been added: "+ "\n" + changes.join("\n"));  
    }
  else {
    ui.alert("No new changes :(")
    }
   }











