function pushToDoc() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var lastRow = sheet.getLastRow()

  var template = 'TEMPLATE_ID_GOES_HERE'

  var range = sheet.getRange(1, 1, 2, 16)
  var values = range.getValues()

  var doc = createDuplicateDocument(template, values[1][0] + " - " + values[1][1])

  pushTitleContent(doc)
  pushAcademicContent(doc)
  pushAssessments(doc)
  pushClassSchedule(doc)
  
}

function pushTitleContent(doc, values) {
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange(1, 1, 2, 16)
  var values = range.getValues()

  for (var i = 0; i < values[0].length; i ++) {
      
      var header = values[0][i]
      var cell = values[1][i]
      
      if (header.length > 0 && cell.length > 0) {      
        replaceParagraph(doc, "//" + header + "//", cell)
    }
  }
}

function pushAcademicContent(doc) {

  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange(6, 1, 2, 16)
  var values = range.getValues()
  
  for (var i = 0; i < values[0].length; i ++) {
      
      var header = values[0][i]
      var cell = values[1][i]
      
      if (header.length > 0 && cell.length > 0) {      
        replaceParagraph(doc, "//" + header + "//", cell)
    }
  }
}


function pushAssessments(doc) {

  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange(11, 1, 4, 16)
  var values = range.getValues()

  var text = ""
  var sums = ""

  // for each assignment (row)
  for (i = 1; i < values.length; i ++) {

    // for each column (category)
    for (j = 0; j < values[0].length; j ++) {
      
      var header = values[0][j]
      var cell = values[i][j]
      
      if (header.length > 0 && cell.length > 0) {
        text = text + "\r" + header + cell
        }
      }

    sums = sums + values[i][0] + ": " + values[i][1] + "\r"
    
    text = text + "\r"

    }

  replaceParagraph(doc, "//Assignments//", text)
  
  
  
  replaceParagraph(doc, "//Assessment Sum//", sums)
  
}


function pushClassSchedule(doc) {

  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange(23, 1, 11, 16)
  var values = range.getValues()
  
  var total_text = ""
  
  // loop over each class #, write content
  for (i = 1; i < values.length; i ++) {
  
    var text = ""
  
    // loop over each column (element)
    for (j = 1; j < values[0].length; j ++) {

      var header = values[0][j]
      var cell = values[i][j]
      
      if (header.length > 0 && cell.length > 0) {
        text = text + header + ": " + cell + "\r"
      }
    }

    if (text.length > 0) {
        total_text = total_text + "\r" + values[i][0] + "\r" + text
      }
  }
  replaceParagraph(doc, "//Class Schedule//", total_text)
}


function createDuplicateDocument(sourceId, name) {
    var source = DriveApp.getFileById(sourceId);
    var newFile = source.makeCopy(name);
    var targetFolder = DriveApp.getFolderById('GOOGLE_DRIVE_FOLDER_ID_GOES_HERE');
    targetFolder.addFile(newFile);
    
    return DocumentApp.openById(newFile.getId());
}

function replaceParagraph(doc, keyword, newText) {
  var ps = doc.getParagraphs();
  for(var i=0; i<ps.length; i++) {
    var p = ps[i];
    var text = p.getText();

    if(text.indexOf(keyword) >= 0) {
      p.setText(newText);
      p.setBold(false);
    }
  } 
}


function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];  
  menuEntries.push({name: "Push to Doc", functionName: "pushToDoc"}); 
  sheet.addMenu("Push", menuEntries);  
}