function pushToCalendar() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var lastRow = sheet.getLastRow()

  var calendar = CalendarApp.getCalendarById('CALENDAR_ID_GOES_HERE')

  var range = sheet.getRange(2, 1, lastRow, 16)
  var values = range.getValues()
 
  // cell numbers from spreadsheet, subtracted 1 (0-based array)
  var assessmentCells = [1, 4, 7, 10, 13]
 
  // loop over 'values'/submissions
  for (var i = 0; i < values.length; i++) {
  
    var course = values[i][0]
  
    // Loops over assessments 1-5 in the row/value
    for (var z = 0; z < assessmentCells.length; z++) {

      var name = values[i][assessmentCells[z]]
      var date = values[i][assessmentCells[z]+1]
      var comments = values[i][assessmentCells[z]+2]
  
      // Skips blanks
      if (name.length > 0) {
         
        var newEventTitle = course + ' - ' + name;
        var newEvent = calendar.createAllDayEvent(newEventTitle, date);
        
        var newEventId = newEvent.getId();

        var desc = newEvent.setDescription(comments)
        
        Logger.log(course + name)
        Logger.log(newEventId)
        }
      
      }
    }
}


function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];  
  menuEntries.push({name: "Update Calendar", functionName: "pushToCalendar"}); 
  sheet.addMenu("Push", menuEntries);  
}