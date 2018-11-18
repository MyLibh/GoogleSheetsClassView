//====================================================================================================================================================================================
//========= User =====================================================================================================================================================================
//====================================================================================================================================================================================

var ROWS_IN_HEADER   = 2;               // Header size,                                              see https://github.com/MyLibh/GoogleSheetsClassView#s-Requirements-Header
var SECOND_GROUP_ROW = NO_SECOND_GROUP; // The line the second group starts with,                    see https://github.com/MyLibh/GoogleSheetsClassView#s-Requirements
var MARKS_LIST_NAME  = "Marks";         // Name of list in pupil's spreadsheet where marks would be, see https://github.com/MyLibh/GoogleSheetsClassView#s-Setup
var LISTS_TO_COPY    = ["Лист2"];

//====================================================================================================================================================================================
//========= Technical ================================================================================================================================================================
//====================================================================================================================================================================================

var NUM_OF_ROWS_TO_COPY      = ROWS_IN_HEADER + 1;                             // Number of rows in header and row for student's marks
var MAIN_SHEET_LINK          = SpreadsheetApp.getActiveSpreadsheet().getUrl(); // Link to the table with marks for all classes
var MAIN_SHEET_PARENT_FOLDER = GetMainSheetFolder();                           // The folder that contains the table with marks

//====================================================================================================================================================================================
//========= Flags ====================================================================================================================================================================
//====================================================================================================================================================================================

var NO_SECOND_GROUP;  // Only one group exists
var NO_LISTS_TO_COPY; // No lists for copying exist

/*
 * \brief  Main function of the script.
 */
function Main()
{
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const source     = spreadsheet.getSheets();
  const classesNum = spreadsheet.getNumSheets();
  for(var class = 0; class < classesNum; ++class)
  {
    if(LISTS_TO_COPY != NO_LISTS_TO_COPY && LISTS_TO_COPY.indexOf(source[class].getName()) != -1)
      continue;

    ProcessClass(source[class]);
  }
}

/*
 * \brief  Processes each student in the class.
 *
 * \param[in]  classSheet  Table(sheet) with grades.
 */
function ProcessClass(classSheet)
{
  if(!MAIN_SHEET_PARENT_FOLDER.getFoldersByName(classSheet.getName()).hasNext())
    MAIN_SHEET_PARENT_FOLDER.createFolder(classSheet.getName());

  const rowsNum             = classSheet.getLastRow(); // Number of rows with data
  const lastRowInFirstGroup = (SECOND_GROUP_ROW == NO_SECOND_GROUP)? rowsNum : SECOND_GROUP_ROW - 1;

  ProcessGroup(classSheet, 1, lastRowInFirstGroup);
  if (SECOND_GROUP_ROW != NO_SECOND_GROUP)
    ProcessGroup(classSheet, SECOND_GROUP_ROW, rowsNum);
}

/*
 * \brief  Processes each student in the group(in the range [firstRow, lastRow]).
 *
 * \param[in]  classSheet  Table(sheet) with grades.
 * \param[in]  firstRow    First row in the group.
 * \param[in]  lastRow     Last row in the group.
 */
function ProcessGroup(classSheet, firstRow, lastRow)
{
  for(var row = firstRow + ROWS_IN_HEADER; row <= lastRow; ++row)
    if(IsEmail(classSheet.getRange("A" + row + ":A" + row).getValue()))
      ProcessStudent(row, classSheet, firstRow);
}

/*
 * \brief  Processes student.
 *
 * \param[in]  row            Student's marks row
 * \param[in]  classSheet     Table(sheet) with grades.
 * \param[in]  firstRawGroup  First row in the group.
 */
function ProcessStudent(row, classSheet, firstRawGroup)
{
  const className   = classSheet.getName();                                   // Class
  const classFolder = DriveApp.getFoldersByName(className).next();            // Folder with name 'className'
  const filename    = classSheet.getRange("B" + row + ":B" + row).getValue(); // Filename(student's full name)
  const columnsNum  = classSheet.getLastColumn();                             // Number of columns with data

  // Create pupil spreadsheet
  {
    var folderIterator = classFolder.getFilesByName(filename);
    while(folderIterator.hasNext())
      classFolder.removeFile(folderIterator.next());

    var studentSpreadsheet = SpreadsheetApp.create(filename, NUM_OF_ROWS_TO_COPY, columnsNum); // Student's spreadsheet
    var copyFile           = DriveApp.getFileById(studentSpreadsheet.getId());                 // Copy of 'studentSpreadsheet' int root folder

    classFolder.addFile(copyFile);
    DriveApp.getRootFolder().removeFile(copyFile);
  }

  // Copy formatting
  {
    var studentSpreadsheet = SpreadsheetApp.openById(studentSpreadsheet.getId());
    classSheet.copyTo(studentSpreadsheet);

    const studentSheets   = studentSpreadsheet.getSheets(); // Array of sheets(lists)
    const copyFormatRange = "1:" + NUM_OF_ROWS_TO_COPY;     // Format copy range

    studentSheets[1].getRange(copyFormatRange).copyTo(studentSheets[0].getRange(copyFormatRange), { formatOnly : true });
    studentSheets[0].deleteColumn(1);

    for(var i = 1; i < columnsNum; ++i)
      studentSheets[0].setColumnWidth(i, studentSheets[1].getColumnWidth(i + 1));

    studentSpreadsheet.deleteSheet(studentSheets[1]);

    CopyLists(SpreadsheetApp.getActiveSpreadsheet(), studentSpreadsheet);
  }

  // Set content
  {
    // Set list name
    studentSheets[0].setName(MARKS_LIST_NAME);

    // Set header
    for(var i = 1; i <= ROWS_IN_HEADER; ++i)
    {
      var studentHeaderFormula = "=IMPORTRANGE(\"" + MAIN_SHEET_LINK + "\";\"" + className + "!B" + (i + firstRawGroup - 1) + ":CC" + (i + firstRawGroup - 1) + "\")";
      studentSheets[0].getRange("A" + i + ":A" + i).setFormula(studentHeaderFormula);
    }

    // Set marks
    var studentMarksRange   = "A" + NUM_OF_ROWS_TO_COPY + ":A" + NUM_OF_ROWS_TO_COPY;
    var studentMarksFormula = "=IMPORTRANGE(\"" + MAIN_SHEET_LINK + "\";\"" + className + "!B" + row + ":CC" + row + "\")";
    studentSheets[0].getRange(studentMarksRange).setFormula(studentMarksFormula);
  }

  ShareSheet(studentSpreadsheet.getId(), classSheet, "A" + row + ":A" + row);
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

/*
 * \brief Finds main sheet folder
 *
 * \return The folder that contains the main table
 */
function GetMainSheetFolder()
{
  var msFileId = SpreadsheetApp.getActive().getId();
  var msFile   = DriveApp.getFileById(msFileId);

  return msFile.getParents().next();
}

/*
 * \brief Shares sheet by email
 *
 * \param[in]  ssId  Spreadsheet id
 * \param[in]  src   Sheet to get email
 * \param[in]  rng   Range in 'src' with email
 */
function ShareSheet(ssId, src, rng)
{
  const file  = DriveApp.getFileById(ssId);
  const email = src.getRange(rng).getValue();
  file.addViewer(email);
}

/*
 * \brief Copies list from 'src' to 'dest'
 *
 * \param[in]  list  List for copying
 * \param[in]  src   Sheet, which contains 'list'
 * \param[in]  dest  Destination of copying
 */
function CopyList(list, src, dest)
{
   src.getSheetByName(list).copyTo(dest);

   dest.getSheets()[dest.getSheets().length - 1].setName(list);
}

/*
 * \brief Copies all lists from 'LISTS_TO_COPY' to dest sheet
 *
 * \param[in]  src   Sheet, which contains 'LISTS_TO_COPY'
 * \param[in]  dest  Destination of copying
 */
function CopyLists(src, dest)
{
  if(LISTS_TO_COPY != NO_LISTS_TO_COPY)
    for(var i = 0; i < LISTS_TO_COPY.length; ++i)
      CopyList(LISTS_TO_COPY[i], src, dest);
}
