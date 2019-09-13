//====================================================================================================================================================================================
//========= Flags ====================================================================================================================================================================
//====================================================================================================================================================================================

var SETUP            = 0x01; // Create students sheets(first time execution)
var UPDATE_STUDENTS  = 0x02; // Add uncreated students
var UPDATE_FORMAT    = 0x03; // Update format
var UPDATE_VIEWERS   = 0x04; // Share additional emails 

var NO_SECOND_GROUP  = null; // Only one group exists

var NO_LISTS_TO_COPY = null; // No lists for copying exist

//====================================================================================================================================================================================
//========= User =====================================================================================================================================================================
//====================================================================================================================================================================================

var SCRIPT_TARGET                  = UPDATE_STUDENTS;            // The purpose of the script

var ROWS_IN_HEADER                 = 3;                // Header size,                                              see https://github.com/MyLibh/GoogleSheetsClassView#s-Requirements-Header
var SECOND_GROUP_ROW               = 21;               // The line the second group starts with,                    see https://github.com/MyLibh/GoogleSheetsClassView#s-Requirements
var LISTS_TO_COPY                  = ["информация"];   // Array of list names

var MARKS_LIST_NAME                = "оценки";          // Name of list in pupil's spreadsheet where marks would be, see https://github.com/MyLibh/GoogleSheetsClassView#s-Setup
var STUDENTS_FOLDER_NAME           = "ведомости школьников";       // Name of folder with "A-D" class folders

var INDIVIDUAL_STUDENT_FOLDER_SUFF = ", ВТЭК"          // Addition to personal folder name

//====================================================================================================================================================================================
//========= Technical ================================================================================================================================================================
//====================================================================================================================================================================================

var NUM_OF_ROWS_TO_COPY      = ROWS_IN_HEADER + 1;                             // Number of rows in header and row for student's marks
var MAIN_SHEET_LINK          = SpreadsheetApp.getActiveSpreadsheet().getUrl(); // Link to the table with marks for all classes
var MAIN_SHEET_PARENT_FOLDER = GetMainSheetFolder();                           // The folder that contains the table with marks
var STUDENTS_FOLDER          = TryToCreateStudentsFolder();                    // Folder with "A-D" class folders

//====================================================================================================================================================================================
//========= Core =====================================================================================================================================================================
//====================================================================================================================================================================================

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
function ProcessStudent(row, classSheet, firstRowGroup)
{
  const className  = classSheet.getName();                                   
  const filename   = classSheet.getRange("B" + row + ":B" + row).getValue(); 
  const columnsNum = classSheet.getLastColumn(); // Number of columns with data

  var studentProps = CreateStudentFolderAndSpreadsheet(className, filename, columnsNum);
  if(studentProps.spreadsheet == null && SCRIPT_TARGET != UPDATE_VIEWERS)
	return; // Either UPDATE_FORMAT for the uncreated student, or UPDATE_STUDENT for the existing

  if(SCRIPT_TARGET == UPDATE_VIEWERS)
  {
    var viewers = studentProps.folder.getViewers();
    var emails = String(classSheet.getRange("A" + row + ":A" + row).getValue()).split(", ");
    
    for(var i = 0; i < emails.length; ++i)
      for(var j = 0; j < viewers.length; ++j)
        if(emails[i] != viewers[j])
          ShareSheet(studentProps.folder.getId(), emails[i]);

    return;
  }
  
  var sheets = ProcessFormat(studentProps.spreadsheet, classSheet, columnsNum);
  ProcessContent(sheets[0], className, firstRowGroup, row);

  var emails = String(classSheet.getRange("A" + row + ":A" + row).getValue()).split(", ");
  for(var i = 0; i < emails.length; ++i)
    if((!studentProps.existed || SCRIPT_TARGET == SETUP) && IsEmail(emails[i]))
      ShareSheet(studentProps.folder.getId(), emails[i]);
}

/*
 * \brief  Creates student's folder and spreadsheet
 *
 * \param[in]  className   Name of the list which contains class marks
 * \param[in]  filename    Name of the student's spreadsheet
 * \param[in]  columnsNum  Number of columns with data
 *
 * \return  Tuple with student's existed state, folder and spreadsheet
 */
function CreateStudentFolderAndSpreadsheet(className, filename, columnsNum)
{
  var classFolder   = STUDENTS_FOLDER.getFoldersByName(className).hasNext() ? 
                        STUDENTS_FOLDER.getFoldersByName(className).next() :  
                        STUDENTS_FOLDER.createFolder(className);
  var existed       = false;
  var studentFolder = classFolder.getFoldersByName(filename + INDIVIDUAL_STUDENT_FOLDER_SUFF).hasNext() ? 
                        (existed = true, 
                        classFolder.getFoldersByName(filename + INDIVIDUAL_STUDENT_FOLDER_SUFF).next()) :  
                        classFolder.createFolder(filename + INDIVIDUAL_STUDENT_FOLDER_SUFF); 
  
  var studentSsName      = SpreadsheetApp.getActiveSpreadsheet().getName();
  var studentFiles       = studentFolder.getFilesByName(studentSsName);
  var studentSpreadsheet = null;
  if(SCRIPT_TARGET == SETUP || (SCRIPT_TARGET == UPDATE_STUDENTS && !studentFiles.hasNext()))
  {
	if(studentFiles.hasNext()) 
      studentFolder.removeFile(studentFiles.next());
      
    studentSpreadsheet = SpreadsheetApp.create(studentSsName, NUM_OF_ROWS_TO_COPY, columnsNum);
    
    var copyFile = DriveApp.getFileById(studentSpreadsheet.getId()); // Copy of 'studentSpreadsheet'
  
    studentFolder.addFile(copyFile);
    DriveApp.getRootFolder().removeFile(copyFile);
  }
  else if(SCRIPT_TARGET == UPDATE_FORMAT && studentFiles.hasNext())
  {
    studentSpreadsheet = studentFiles.next();
  } 
  
  var res = 
  {
    existed:     existed,
    folder:      studentFolder,
    spreadsheet: studentSpreadsheet
  };
  
  return res;
}

/*
 * \brief  Processes list format
 *
 * \param[in]  spreadsheet  Student's spreadsheet
 * \param[in]  classSheet   Sheet with class's marks
 * \param[in]  columnsNum   Number of columns with data
 *
 * \return  Array of student's lists
 */
function ProcessFormat(spreadsheet, classSheet, columnsNum)
{
  var studentSpreadsheet = SpreadsheetApp.openById(spreadsheet.getId());
  classSheet.copyTo(studentSpreadsheet);
  
  const sheets = studentSpreadsheet.getSheets(); // Array of student's sheets(lists)
  const range  = "1:" + NUM_OF_ROWS_TO_COPY;     // Format copy range
  
  sheets[sheets.length - 1].getRange(range).copyTo(sheets[0].getRange(range), { formatOnly : true });
  
  for(var i = 0; i < columnsNum; ++i)
    sheets[0].setColumnWidth(i + 1, sheets[sheets.length - 1].getColumnWidth(i + 1));
  
  studentSpreadsheet.deleteSheet(sheets[sheets.length - 1]);
  
  ProcessLists(SpreadsheetApp.getActiveSpreadsheet(), studentSpreadsheet);
  
  return sheets;
}

/*
 * \brief  Processes all lists from 'LISTS_TO_COPY'
 *
 * \param[in]  src        Sheet, which contains 'LISTS_TO_COPY'
 * \param[in]  dest       Destination of copying
 */
function ProcessLists(src, dest, className)
{
  if(LISTS_TO_COPY != NO_LISTS_TO_COPY)
    for(var i = 0; i < LISTS_TO_COPY.length; ++i)
      ProcessList(LISTS_TO_COPY[i], src, dest);
}

/*
 * \brief  Processes list from 'src' to 'dest'
 *
 * \param[in]  list  List for processing
 * \param[in]  src   Sheet, which contains 'list'
 * \param[in]  dest  estination for process
 */
function ProcessList(list, src, dest)
{
  var list4copying = src.getSheetByName(list);
  if(list4copying != null)
  {
	if(SCRIPT_TARGET == UPDATE_FORMAT)
	  dest.deleteSheet(dest.getSheetByName(list));

	var new_list = dest.insertSheet();
	list4copying.copyTo(dest);
     
	var sheets = dest.getSheets();
	var copied_list = sheets[sheets.length - 1];
	copied_list.getRange("A:Z").copyTo(new_list.getRange("A:Z"), { formatOnly : true });

	var columnsNum = copied_list.getLastColumn();
	for(var i = 0; i < columnsNum; ++i)
	new_list.setColumnWidth(i + 1, copied_list.getColumnWidth(i + 1));

	new_list.setName(list);
	new_list.getRange("A1:A1").setFormula("=IMPORTRANGE(\"" + MAIN_SHEET_LINK + "\";\"" + list + "!A:Z\")");
     
	dest.deleteSheet(copied_list);
  }
}

/*
 * \brief  Sets content in student's sheet
 *
 * \param[in]  list           List with marks
 * \param[in]  className      Name of the list which contains class marks
 * \param[in]  firstRowGroup  First row of the group
 * \param[in]  row            Row with student's marks
 */
function ProcessContent(list, className, headerRow,  row)
{
  if(SCRIPT_TARGET == SETUP || SCRIPT_TARGET == UPDATE_STUDENTS)
  {
	list.setName(MARKS_LIST_NAME);

	// Set header
	for(var i = 1; i <= ROWS_IN_HEADER; ++i)
	{
	var headerFormula = MakeImportrangeRowFormula(className, i + headerRow - 1);
	list.getRange("A" + i + ":A" + i).setFormula(headerFormula);
	}

	// Set marks
	var marksRange   = "A" + NUM_OF_ROWS_TO_COPY + ":A" + NUM_OF_ROWS_TO_COPY;
	var marksFormula = MakeImportrangeRowFormula(className, row);
	list.getRange(marksRange).setFormula(marksFormula);
  }
}

/*
 * \brief  Checks if 'obj' is email
 *
 * \param[in]  obj  The object to find a match
 *
 * \return  TRUE if 'obj' matches, otherwise it FALSE
 */
function IsEmail(obj)
{
  const pattern = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;

  return pattern.test(String(obj).toLowerCase());
}

/*
 * \brief  Finds main sheet folder
 *
 * \return  The folder that contains the main table
 */
function GetMainSheetFolder()
{
  var msFileId      = SpreadsheetApp.getActive().getId();
  var msFileParents = DriveApp.getFileById(msFileId).getParents();

  return msFileParents.hasNext() ? msFileParents.next() : DriveApp.getRootFolder();
}

/*
 * \brief  Creates students' folder if it does not exist
 *
 * \return  The folder that will contain each student folder
 */
function TryToCreateStudentsFolder()
{
  return MAIN_SHEET_PARENT_FOLDER.getFoldersByName(STUDENTS_FOLDER_NAME).hasNext() ? 
           MAIN_SHEET_PARENT_FOLDER.getFoldersByName(STUDENTS_FOLDER_NAME).next() :  
           MAIN_SHEET_PARENT_FOLDER.createFolder(STUDENTS_FOLDER_NAME); 
}

/*
 * \brief  Shares sheet by email
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
 * \brief  Makes IMPORTRANGE for specified row
 *
 * \param[in]  className  Name of the sheet where to produce IMPORTRANGE
 * \param[in]  row        Row for import
 *
 * \return  IMPORTRANGE formula
 */
function MakeImportrangeRowFormula(className, row)
{
  return "=IMPORTRANGE(\"" + MAIN_SHEET_LINK + "\";\"" + className + "!" + row + ":" + row + "\")";
}
