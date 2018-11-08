// USER CONSTANTS
var NUM_OF_LINES_IN_HEADER       = 1;
var NUM_OF_LINES_TO_COPY         = NUM_OF_LINES_IN_HEADER + 1;
var MAX_NUM_OF_PUPILS            = 14;
var MAIN_SHEET_LINK              = "https://docs.google.com/spreadsheets/d/1AYy515x2Pt2CRVlqi_BPxtULZZVNRlaQhXFjEwm1BtQ/edit#gid=0";
var LIST_WITH_STUDENT_MARKS_NAME = "Marks";

/*********************************************************************************************************************************************************************************************************/
function ProceedStudent(row, classSheet)
{
    var className = classSheet.getName();
  
    var email    = classSheet.getRange("A" + row + ":A" + row).getValue(); // Student's email
    var filename = classSheet.getRange("B" + row + ":B" + row).getValue(); // Student's filename

    // Create pupil spreadsheet 
    var classFolder = DriveApp.getFoldersByName(className).next();
    var studentFile = SpreadsheetApp.create("hui" + filename, NUM_OF_LINES_TO_COPY, 100);
    var copyFile    = DriveApp.getFileById(studentFile.getId());
  
    classFolder.addFile(copyFile);
    DriveApp.getRootFolder().removeFile(copyFile);
    
    var studentSpreadsheet = SpreadsheetApp.openById(studentFile.getId());
    classSheet.copyTo(studentSpreadsheet);
    
    // Copy formatting
    var studentSheets = studentSpreadsheet.getSheets();
  
    studentSheets[1].clearContents();
    //studentSheets[1]/*.getRange("A1:CC" + NUM_OF_LINES_TO_COPY)*/.copyTo(studentSheets[0]/*.getRange("A1:CC" + NUM_OF_LINES_TO_COPY)*/);
  
    // Set content
    studentSheets[1].setName(LIST_WITH_STUDENT_MARKS_NAME); // Set list name

    studentSheets[1].getRange("A1:A1").setFormula("=IMPORTRANGE(\"" + MAIN_SHEET_LINK + "\";\"" + className + "!B1:CC" + NUM_OF_LINES_IN_HEADER + "\")");
    studentSheets[1].getRange("A" + NUM_OF_LINES_TO_COPY + ":A" + NUM_OF_LINES_TO_COPY + "").setFormula("=IMPORTRANGE(\"" + MAIN_SHEET_LINK + "\";\"" + className + "!B" + row + ":CC" + row + "\")");

    // Clear
    
    studentSheets[1].deleteRows(NUM_OF_LINES_TO_COPY + 1, 50);
    studentSpreadsheet.deleteSheet(studentSheets[0]);
  
    // Share student's sheet
    var file = DriveApp.getFileById(studentSpreadsheet.getId());
    file.addViewer(email);
}

/*********************************************************************************************************************************************************************************************************/
function ProceedClass(classSheet)
{
    DriveApp.getRootFolder().createFolder(classSheet.getName());
  
    for(var row = NUM_OF_LINES_IN_HEADER; row < NUM_OF_LINES_IN_HEADER + MAX_NUM_OF_PUPILS; ++row)
    {
        if(classSheet.getRange("A" + row + ":A" + row).getValue() == "") // If students run out(email does not exist)
            break;
      
        ProceedStudent(NUM_OF_LINES_IN_HEADER + row, classSheet);
    }
}

/*********************************************************************************************************************************************************************************************************/
function Main()
{
    var source = SpreadsheetApp.getActiveSpreadsheet().getSheets();

    var numberOfClasses = SpreadsheetApp.getActiveSpreadsheet().getNumSheets();
    for(var class = 0; class < numberOfClasses; ++class)
      ProceedClass(source[class]);
}
