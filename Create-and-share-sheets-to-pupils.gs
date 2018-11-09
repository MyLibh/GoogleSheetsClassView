/********** User ***********************************************************************************************************************************************************************************************/

var NUM_OF_LINES_IN_HEADER       = 2;
var MAX_NUM_OF_LINES_TO_PROCEED  = 50;
var MAX_NUM_OF_ROWS_TO_PROCEED  = 20;
var START_LINE_OF_SECOND_GROOP   = 30;
var LIST_WITH_STUDENT_MARKS_NAME = "Marks";

/********** Technical ***********************************************************************************************************************************************************************************************/

var NUM_OF_LINES_TO_COPY         = NUM_OF_LINES_IN_HEADER + 1;
var MAIN_SHEET_LINK              = SpreadsheetApp.getActiveSpreadsheet().getUrl();

/*********************************************************************************************************************************************************************************************************/
function IsEmail(email)
{
    //return (Text)email.editAsText().findText("@");
}

/*********************************************************************************************************************************************************************************************************/
function ProceedStudent(row, classSheet, headerInd)
{
    var className = classSheet.getName();

    var email    = classSheet.getRange("A" + row + ":A" + row).getValue(); // Student's email
    var filename = classSheet.getRange("B" + row + ":B" + row).getValue(); // Student's filename

    // Create pupil spreadsheet
    var classFolder = DriveApp.getFoldersByName(className).next();
    var studentFile = SpreadsheetApp.create("_" + filename, NUM_OF_LINES_TO_COPY, 100);
    var copyFile    = DriveApp.getFileById(studentFile.getId());

    classFolder.addFile(copyFile);
    DriveApp.getRootFolder().removeFile(copyFile);

    // Copy formatting
    var studentSpreadsheet = SpreadsheetApp.openById(studentFile.getId());
    classSheet.copyTo(studentSpreadsheet);

    var studentSheets = studentSpreadsheet.getSheets();

    studentSheets[1].getRange("A:CC").copyTo(studentSheets[0].getRange("A:CC"), {formatOnly:true});
    for(var i = 1; i < MAX_NUM_OF_ROWS_TO_PROCEED; ++i)
      studentSheets[0].setColumnWidth(i, studentSheets[1].getColumnWidth(i+1));

    studentSpreadsheet.deleteSheet(studentSheets[1]);

    // Set content
    studentSheets[0].setName(LIST_WITH_STUDENT_MARKS_NAME); // Set list name

    studentSheets[0].getRange("A1:A1").setFormula("=IMPORTRANGE(\"" + MAIN_SHEET_LINK + "\";\"" + className + "!B" + headerInd + ":CC" + (NUM_OF_LINES_IN_HEADER+headerInd-1) + "\")");
    studentSheets[0].getRange("A" + NUM_OF_LINES_TO_COPY + ":A" + NUM_OF_LINES_TO_COPY + "").setFormula("=IMPORTRANGE(\"" + MAIN_SHEET_LINK + "\";\"" + className + "!B" + row + ":CC" + row + "\")");

    // Share student's sheet
    var file = DriveApp.getFileById(studentSpreadsheet.getId());
    file.addViewer(email);
}

/*********************************************************************************************************************************************************************************************************/
function ProceedClass(classSheet)
{
    DriveApp.getRootFolder().createFolder(classSheet.getName());

    for(var row = NUM_OF_LINES_TO_COPY; row < START_LINE_OF_SECOND_GROOP; ++row)
    {
      if(IsEmail(classSheet.getRange("A" + row + ":A" + row).getValue()))
          ProceedStudent(row, classSheet, 1);
    }

    for(var row = START_LINE_OF_SECOND_GROOP; row <= MAX_NUM_OF_LINES_TO_PROCEED; ++row)
    {
      if(IsEmail(classSheet.getRange("A" + row + ":A" + row).getValue()))
          ProceedStudent(row, classSheet, START_LINE_OF_SECOND_GROOP);
    }
}

/*********************************************************************************************************************************************************************************************************/
function Main()
{
    var source = SpreadsheetApp.getActiveSpreadsheet().getSheets();

    /*var numberOfClasses = SpreadsheetApp.getActiveSpreadsheet().getNumSheets();
    for(var class = 0; class < numberOfClasses; ++class)
      ProceedClass(source[class]);*/

    ProceedStudent(4, source[0], 1);
}
