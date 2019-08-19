function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Test It')
      .addItem('Update Calendar', 'updateCalendar')
      .addToUi()
};
function updateCalendar() {
  var calendarId = "one99.pl_mdesm69uckc4ig8br3n09m21is@group.calendar.google.com";
  var eventCal = CalendarApp.getCalendarById(calendarId);

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
  var allRows = range.getValues();
  var row = 0;
  var signups = [];
  var changes = [];
  
  for (row=0;row<allRows.length;row++){
      if (!allRows[row].join("")){
          signups.push(allRows[row]);
      };
  };
  
  for (x=0; x<signups.length; x++) {
      var vacation = signups[x];
      var employee = vacation[0];
      var type = vacation[1];
      var startTime = vacation[2];
      var endTime = vacation[3];
      var record = employee + " (" + type + ")";
      var eventColor = "9";

      var eventCal = CalendarApp.getCalendarById(calendarId);
      var alreadyExists = eventCal.getEvents(startTime, endTime, {search: employee});

      if (type == "Urlop"){
          eventColor = "11"
      };

      if (alreadyExists.length <= 0) {
          if (startTime.getTime() === endTime.getTime()) {
              eventCal.createAllDayEvent(record, startTime).setColor(eventColor);
              changes.push(record+ " " + startTime.toLocaleDateString("en-US"));
          }  
          else {
              eventCal.createAllDayEvent(record, startTime, endTime).setColor(eventColor)
              changes.push(record+ " " + startTime.toLocaleDateString("en-US")+ " - "+endTime.toLocaleDateString("en-US"))
          }      
      }
  };

  var ui = SpreadsheetApp.getUi();
  if (changes.length>0){
      ui.alert("Following changes have been added: "+ "\n" + changes.join("\n"));  
  }
  else {
      ui.alert("No new changes :(")
  };
};