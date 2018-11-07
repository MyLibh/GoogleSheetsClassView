var NUM_OF_FIXED_LINES = 2; // Too many uses of 'NUM_OF_FIXED_LINES + 1'

/*********************************************************************************************************************************************************************************************************/
function ProceedPupil(line, source, url)
{
    var email = source.getRange("A" + line + ":A" + line).getValue(); // Pupil's email
    var name  = source.getRange("B" + line + ":B" + line).getValue(); // Pupil's name
    var sheet = source.getName(); // sheet???

    // Create pupil spreadsheet and copy sheet with marks
    var newDoc = SpreadsheetApp.create(name, (NUM_OF_FIXED_LINES + 1), 100);
    var dest   = SpreadsheetApp.openById(newDoc.getId());
    source.copyTo(dest);

    var destSheets = dest.getSheets();

    // Save formatting
    destSheets[1].clearContents();
    destSheets[1].getRange("A1:CC" + (NUM_OF_FIXED_LINE + 1)).copyTo(destSheets[0].getRange("A1:CC" + (NUM_OF_FIXED_LINES + 1)));

    dest.deleteSheet(destSheets[1]);

    // Set content
    destSheets[0].setName("Marks"); // Set list name

    destSheets[0].getRange("A1:A1").setFormula("=IMPORTRANGE(\"" + url + "\";\"" + sheet + "!B1:CC" + NUM_OF_FIXED_LINES + "\")");
    destSheets[0].getRange("A" + (NUM_OF_FIXED_LINES + 1) + ":A" + (NUM_OF_FIXED_LINES + 1) + "").setFormula("=IMPORTRANGE(\"" + url + "\";\"" + sheet + "!B" + line + ":CC" + line + "\")");

    // Share pupil's sheet
    var file = DriveApp.getFileById(dest.getId());
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
}

/*********************************************************************************************************************************************************************************************************/
function ProceedClass(var classSheet)
{
    // for each of 'line' call 'ProceedPupil()'
}

/*********************************************************************************************************************************************************************************************************/
function Main()
{
    var source = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var url    = SpreadsheetApp.getActiveSpreadsheet().getUrl(); // Make global?

    // for each of 'getSheets() call 'ProceedClass()'

    ProceedPupil(3, source, url);
}
