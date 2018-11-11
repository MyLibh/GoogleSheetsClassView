//====================================================================================================================================================================================
//========= User =====================================================================================================================================================================
//====================================================================================================================================================================================

var ROWS_IN_HEADER   = 2;       // Lesha napishet na perfect english
var SECOND_GROUP_ROW = 17;      // -//-
var MARKS_LIST_NAME  = "Marks"; // -//-

//====================================================================================================================================================================================
//========= Technical ================================================================================================================================================================
//====================================================================================================================================================================================

var NUM_OF_ROWS_TO_COPY = ROWS_IN_HEADER + 1;                     // Number of rows + 1 row for student's marks
var MAIN_SHEET_LINK     = SpreadsheetApp.getActiveSpreadsheet().getUrl(); // Link to the table with marks for all classes

/*
 * \brief  Main function of the script.
 */
function Main()
{
  const source = SpreadsheetApp.getActiveSpreadsheet().getSheets();

  ProcessClass(source[0]);

  /*
  var classesNum = SpreadsheetApp.getActiveSpreadsheet().getNumSheets();
  for(var class = 0; class < classesNum; ++class)
    ProcessClass(source[class]);
  */
}

/*
 * \brief  Processes each student in the class.
 *
 * \param[in]  classSheet  Table(sheet) with grades.
 */
function ProcessClass(classSheet)
{
  DriveApp.getRootFolder().createFolder(classSheet.getName());

  for(var row = NUM_OF_ROWS_TO_COPY; row < SECOND_GROUP_ROW; ++row)
    if(IsEmail(classSheet.getRange("A" + row + ":A" + row).getValue()))
      ProcessStudent(row, classSheet, 1);

  const rowsNum = classSheet.getLastRow(); // Number of rows with data
  for(var row = SECOND_GROUP_ROW + ROWS_IN_HEADER; row < rowsNum; ++row)
    if(IsEmail(classSheet.getRange("A" + row + ":A" + row).getValue()))
      ProcessStudent(row, classSheet, SECOND_GROUP_ROW);
}

/*
 * \brief  Processes student.
 *
 * \param[in]  row          Student's marks row
 * \param[in]  classSheet   Table(sheet) with grades.
 * \param[in]  groupOffset  Offset of the group
 */
function ProcessStudent(row, classSheet, groupOffset)
{
  const className   = classSheet.getName();                                   // Class
  const classFolder = DriveApp.getFoldersByName(className).next();            // Folder with name 'className'
  const filename    = classSheet.getRange("B" + row + ":B" + row).getValue(); // Filename(student's full name)
  const columnsNum  = classSheet.getLastColumn();                             // Number of columns with data

  // Create pupil spreadsheet
  {
    var studentSpreadsheet = SpreadsheetApp.create(filename, NUM_OF_ROWS_TO_COPY, columnsNum); // Student's spreadsheet
    var copyFile           = DriveApp.getFileById(studentSpreadsheet.getId());                 // Copy of 'studentSpreadsheet' int root folder

    classFolder.addFile(copyFile);
    DriveApp.getRootFolder().removeFile(copyFile);
  }

  // Copy formatting
  {
    var studentSpreadsheet = SpreadsheetApp.openById(studentSpreadsheet.getId());
    classSheet.copyTo(studentSpreadsheet);

    var studentSheets   = studentSpreadsheet.getSheets(); // Array of sheets(lists)
    var copyFormatRange = "1:" + NUM_OF_ROWS_TO_COPY;     // Format copy range
    studentSheets[1].getRange(copyFormatRange).copyTo(studentSheets[0].getRange(copyFormatRange), { formatOnly : true });
    studentSheets[0].deleteColumn(1);

    for(var i = 1; i < columnsNum; ++i)
      studentSheets[0].setColumnWidth(i, studentSheets[1].getColumnWidth(i + 1));

    studentSpreadsheet.deleteSheet(studentSheets[1]);
  }

  // Set content
  {
    // Set list name
    studentSheets[0].setName(MARKS_LIST_NAME);

    // Set header
    for(var i = 1; i <= ROWS_IN_HEADER; ++i)
    {
      var studentHeaderFormula = "=IMPORTRANGE(\"" + MAIN_SHEET_LINK + "\";\"" + className + "!B" + (i + groupOffset - 1) + ":CC" + (i + groupOffset - 1) + "\")";
      studentSheets[0].getRange("A" + i + ":A" + i).setFormula(studentHeaderFormula);
    }

    // Set marks
    var studentMarksRange   = "A" + NUM_OF_ROWS_TO_COPY + ":A" + NUM_OF_ROWS_TO_COPY;
    var studentMarksFormula = "=IMPORTRANGE(\"" + MAIN_SHEET_LINK + "\";\"" + className + "!B" + row + ":CC" + row + "\")";
    studentSheets[0].getRange(studentMarksRange).setFormula(studentMarksFormula);
  }

  // Share student's sheet
  {
    const file  = DriveApp.getFileById(studentSpreadsheet.getId());
    const email = classSheet.getRange("A" + row + ":A" + row).getValue();
    file.addViewer(email);
  }
}

/*
 * \brief  Checks if 'obj' is email
 *
 * \param[in]  obj  The object to find a match
 *
 * \return TRUE if 'obj' matches, otherwise it FALSE
 */
function IsEmail(obj)
{
  const pattern = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;

  return pattern.test(String(obj).toLowerCase());
}
