function func()
{
  // Create pupil spreadsheet and copy sheet with marks
  var newDoc = SpreadsheetApp.create("copy", 2, 50);
  var source = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var dest   = SpreadsheetApp.openById(newDoc.getId());
  source.copyTo(dest); 
  
  // Copy formatting
  var destSheets = dest.getSheets();
  destSheets[1].clearContents();
  destSheets[1].getRange("A1:CC2").copyTo(destSheets[0].getRange("A1:CC2"));
  
  dest.deleteSheet(destSheets[1]);
  
  // Set content
  destSheets[0].setName("Marks");
  destSheets[0].getRange("A1:A1").setFormula("=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1AYy515x2Pt2CRVlqi_BPxtULZZVNRlaQhXFjEwm1BtQ/edit#gid=0\";\"CopySheet!A1:CC1\")");
  destSheets[0].getRange("A2:A2").setFormula("=IMPORTRANGE(\"https://docs.google.com/spreadsheets/d/1AYy515x2Pt2CRVlqi_BPxtULZZVNRlaQhXFjEwm1BtQ/edit#gid=0\";\"CopySheet!A2:CC2\")");
  
  // Share pupil's sheet
  var file = DriveApp.getFileById(dest.getId());
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  // destSheets[0].getRange("E1").setValue(file.getUrl());
}
