//====================================================================================================================================================================================
//========= User =====================================================================================================================================================================
//====================================================================================================================================================================================

var ROWS_IN_HEADER       = 2;                // Header size,                                              see https://github.com/MyLibh/GoogleSheetsClassView#s-Requirements-Header
var SECOND_GROUP_ROW     = 23;               // The line the second group starts with,                    see https://github.com/MyLibh/GoogleSheetsClassView#s-Requirements
var LISTS_TO_COPY        = ["информация"];   // Array of list names
var MARKS_LIST_NAME      = "Marks";          // Name of list in pupil's spreadsheet where marks would be, see https://github.com/MyLibh/GoogleSheetsClassView#s-Setup
var STUDENTS_FOLDER_NAME = "Students";       // Name of folder with "A-D" class folders

//====================================================================================================================================================================================
//========= Technical ================================================================================================================================================================
//====================================================================================================================================================================================

var NUM_OF_ROWS_TO_COPY      = ROWS_IN_HEADER + 1;                             // Number of rows in header and row for student's marks
var MAIN_SHEET_LINK          = SpreadsheetApp.getActiveSpreadsheet().getUrl(); // Link to the table with marks for all classes
var MAIN_SHEET_PARENT_FOLDER = GetMainSheetFolder();                           // The folder that contains the table with marks
var STUDENTS_FOLDER          = TryToCreateStudentsFolder();                    // Folder with "A-D" class folders

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
  {
    var emails = String(classSheet.getRange("A" + row + ":A" + row).getValue()).split(", ");
    var hasValidEmail = false;
    for(var i = 0; i < emails.length; ++i)      
      if(IsEmail(emails[i]))
      {
         hasValidEmail = true;
         
         break;
      }

    if(hasValidEmail)
      ProcessStudent(row, classSheet, firstRow);
  }
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
  const className   = classSheet.getName();                                   
  const filename    = classSheet.getRange("B" + row + ":B" + row).getValue(); 
  const columnsNum  = classSheet.getLastColumn(); // Number of columns with data

  // Create pupil's folder and spreadsheet
  {
    var classFolder        = STUDENTS_FOLDER.getFoldersByName(className).hasNext() ? 
                               STUDENTS_FOLDER.getFoldersByName(className).next() :  
                               STUDENTS_FOLDER.createFolder(className);
    var studentFolder      = classFolder.getFoldersByName(filename).hasNext() ? 
                               classFolder.getFoldersByName(filename).next() :  
                               classFolder.createFolder(filename);                            
    var studentSpreadsheet = SpreadsheetApp.create("#" + " класс " + "####-####", NUM_OF_ROWS_TO_COPY, columnsNum); 
    var copyFile           = DriveApp.getFileById(studentSpreadsheet.getId()); // Copy of 'studentSpreadsheet'

    studentFolder.addFile(copyFile);
    DriveApp.getRootFolder().removeFile(copyFile);
  }

  // Copy formatting
  {
    var studentSpreadsheet = SpreadsheetApp.openById(studentSpreadsheet.getId());
    classSheet.copyTo(studentSpreadsheet);

    const studentSheets   = studentSpreadsheet.getSheets(); // Array of sheets(lists)
    const copyFormatRange = "1:" + NUM_OF_ROWS_TO_COPY;     // Format copy range

    studentSheets[1].getRange(copyFormatRange).copyTo(studentSheets[0].getRange(copyFormatRange), { formatOnly : true });
    
    for(var i = 0; i < columnsNum; ++i)
      studentSheets[0].setColumnWidth(i + 1, studentSheets[1].getColumnWidth(i + 1));

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
      var studentHeaderFormula = "=IMPORTRANGE(\"" + MAIN_SHEET_LINK + "\";\"" + className + "!" + (i + firstRawGroup - 1) + ":" + (i + firstRawGroup - 1) + "\")";
      studentSheets[0].getRange("A" + i + ":A" + i).setFormula(studentHeaderFormula);
    }

    // Set marks
    var studentMarksRange   = "A" + NUM_OF_ROWS_TO_COPY + ":A" + NUM_OF_ROWS_TO_COPY;
    var studentMarksFormula = "=IMPORTRANGE(\"" + MAIN_SHEET_LINK + "\";\"" + className + "!" + row + ":" + row + "\")";
    studentSheets[0].getRange(studentMarksRange).setFormula(studentMarksFormula);
  }

  var emails = String(classSheet.getRange("A" + row + ":A" + row).getValue()).split(", ");
  for(var i = 0; i < emails.length; ++i)
    if(IsEmail(emails[i]))
      ShareSheet(studentFolder.getId(), emails[i]);
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
 * \brief Creates students' folder if it does not exist
 *
 * \return The folder that will contain each student folder
 */
function TryToCreateStudentsFolder()
{
  return MAIN_SHEET_PARENT_FOLDER.getFoldersByName(STUDENTS_FOLDER_NAME).hasNext() ? 
           MAIN_SHEET_PARENT_FOLDER.getFoldersByName(STUDENTS_FOLDER_NAME).next() :  
           MAIN_SHEET_PARENT_FOLDER.createFolder(STUDENTS_FOLDER_NAME); 
}

/*
 * \brief Shares sheet by email
 *
 * \param[in]  ssId  Spreadsheet id
 * \param[in]  src   Sheet to get email
 * \param[in]  rng   Range in 'src' with email
 */
function ShareSheet(ssId, email)
{
  const file = DriveApp.getFileById(ssId);
  
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
   var сopy_list = src.getSheetByName(list)
   if(сopy_list != null)
   {
     сopy_list.copyTo(dest);

     dest.getSheets()[dest.getSheets().length - 1].setName(list);
   }
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
