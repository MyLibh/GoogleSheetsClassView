//====================================================================================================================================================================================
//========= Flags ====================================================================================================================================================================
//====================================================================================================================================================================================

var SETUP            = 0x01; // Create students sheets(first time execution)
var UPDATE_STUDENTS  = 0x02; // Add uncreated students
var UPDATE_FORMAT    = 0x03; // Update format
var UPDATE_VIEWERS   = 0x04; // Share additional emails

//====================================================================================================================================================================================
//========= User =====================================================================================================================================================================
//====================================================================================================================================================================================

var SCRIPT_TARGET;                  // The purpose of the script

var ROWS_IN_HEADER;                 // Header size
var SECOND_GROUP_ROW = null;        // The line the second group starts with
var LISTS_TO_COPY    = null;        // Array of list names

var MARKS_LIST_NAME;                // Name of list in pupil's spreadsheet where marks would be
var STUDENTS_FOLDER_NAME;           // Name of folder with "A-D" class folders

var INDIVIDUAL_STUDENT_FOLDER_SUFF; // Addition to personal folder name

//====================================================================================================================================================================================
//========= Technical ================================================================================================================================================================
//====================================================================================================================================================================================

var NUM_OF_ROWS_TO_COPY;      // Number of rows in header and row for student's marks
var MAIN_SHEET_LINK;          // Link to the table with marks for all classes
var MAIN_SHEET_PARENT_FOLDER; // The folder that contains the table with marks
var STUDENTS_FOLDER;          // Folder with "A-D" class folders

//====================================================================================================================================================================================
//========= Core =====================================================================================================================================================================
//====================================================================================================================================================================================

function start(params)
{
    params = params[""];

    SCRIPT_TARGET                  = Number(params[0]);
    STUDENTS_FOLDER_NAME           = params[1];
    MARKS_LIST_NAME                = params[2];
    INDIVIDUAL_STUDENT_FOLDER_SUFF = params[3];
    ROWS_IN_HEADER                 = Number(params[4]);

    var i = 5;
    if (params[i] == "on")
        SECOND_GROUP_ROW = Number(params[++i]);

    if (++i + 1 < params.length - 1) // -1 for "Close"
    {
        LISTS_TO_COPY = [];
        for(i++; i < params.length - 1; ++i)
            LISTS_TO_COPY.push(params[i]);
    }

    NUM_OF_ROWS_TO_COPY      = ROWS_IN_HEADER + 1;
    MAIN_SHEET_LINK          = SpreadsheetApp.getActiveSpreadsheet().getUrl(); 
    MAIN_SHEET_PARENT_FOLDER = GetMainSheetFolder();                         
    STUDENTS_FOLDER          = TryToCreateStudentsFolder();                    
    
    Logger.log("target: ", SCRIPT_TARGET);
    Logger.log("students folder name:", STUDENTS_FOLDER_NAME);
    Logger.log("marks list name: ", MARKS_LIST_NAME);
    Logger.log("student folder stuff: ", INDIVIDUAL_STUDENT_FOLDER_SUFF);
    Logger.log("rows in header: ", ROWS_IN_HEADER);
    Logger.log("second group row: ", SECOND_GROUP_ROW);
    Logger.log("lists to copy: ", LISTS_TO_COPY);
    Logger.log("rows to copy: ", NUM_OF_ROWS_TO_COPY);
    Logger.log("main sheet link: ", MAIN_SHEET_LINK);
    Logger.log("main sheet parent folder: ", MAIN_SHEET_PARENT_FOLDER);
    Logger.log("students folder: ", STUDENTS_FOLDER);

    Main();
}

function onOpen()
{
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Script')
      .addItem('Sidebar', 'showSidebar')
      .addToUi();
}

function showSidebar()
{
    var html = HtmlService
        .createTemplateFromFile('sidebar')
        .evaluate()
        .setTitle('Меню скрипта')
        .setWidth(200);

    SpreadsheetApp.getUi().showSidebar(html);
}

/*
 * \brief  Main function of the script.
 */
function Main()
{
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const source     = spreadsheet.getSheets();
  const classesNum = spreadsheet.getNumSheets();
  for(var classIdx = 0; classIdx < classesNum; ++classIdx)
  {
    if(LISTS_TO_COPY != null && LISTS_TO_COPY.indexOf(source[classIdx].getName()) != -1)
      continue;

    ProcessClass(source[classIdx]);
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
  const lastRowInFirstGroup = (SECOND_GROUP_ROW == null)? rowsNum : SECOND_GROUP_ROW - 1;

  ProcessGroup(classSheet, 1, lastRowInFirstGroup);
  if (SECOND_GROUP_ROW != null)
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
	  return; // Either UPDATE_FORMAT for the uncreated student

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
  else if (SCRIPT_TARGET == UPDATE_STUDENTS && studentProps.existed)
  {
    ProcessContent(SpreadsheetApp.openById(studentProps.spreadsheet.getId()).getSheets()[0], className, firstRowGroup, row);

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
      studentFiles.next().setTrashed(true);

    studentSpreadsheet = SpreadsheetApp.create(studentSsName, NUM_OF_ROWS_TO_COPY, columnsNum);
    
    DriveApp.getFileById(studentSpreadsheet.getId()).moveTo(studentFolder);
  }
  else if((SCRIPT_TARGET == UPDATE_FORMAT || SCRIPT_TARGET == UPDATE_STUDENTS) && studentFiles.hasNext())
  {
    studentSpreadsheet = studentFiles.next();
    existed = true;
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
  if(LISTS_TO_COPY != null)
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
