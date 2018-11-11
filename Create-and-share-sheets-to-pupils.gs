/********** User ***********************************************************************************************************************************************************************************************/ 
var NUM_OF_LINES_IN_HEADER       = 2; 
var START_LINE_OF_SECOND_GROUP   = 17; 
var LIST_WITH_STUDENT_MARKS_NAME = "Marks"; 

//====================================================================================================================================================================================
//========= Technical ===========================================================================================================================================================================
//====================================================================================================================================================================================

var NUM_OF_ROWS_TO_COPY = NUM_OF_LINES_IN_HEADER + 1; 
var MAIN_SHEET_LINK     = SpreadsheetApp.getActiveSpreadsheet().getUrl(); 

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

/*********************************************************************************************************************************************************************************************************/ 
/*
 * \brief  Processes student.
 *
 * \param[in]  row          Student's marks row
 * \param[in]  classSheet   Table(sheet) with grades.
 * \param[in]  groupOffset  Offset of the group
 */
function ProceedStudent(row, classSheet, headerInd) 
{ 
  // Create pupil spreadsheet 
  var className = classSheet.getName(); 
  var filename  = classSheet.getRange("B" + row + ":B" + row).getValue(); // Student's filename 
  
  var columnsNum = classSheet.getLastColumn();
  
  var classFolder = DriveApp.getFoldersByName(className).next(); 
  var studentFile = SpreadsheetApp.create("_" + filename, NUM_OF_ROWS_TO_COPY, columnsNum); 
  var copyFile    = DriveApp.getFileById(studentFile.getId()); 
  
  classFolder.addFile(copyFile); 
  DriveApp.getRootFolder().removeFile(copyFile); 
  
  // Copy formatting
  { 
    var studentSpreadsheet = SpreadsheetApp.openById(studentFile.getId()); 
    classSheet.copyTo(studentSpreadsheet); 
    
    var studentSheets = studentSpreadsheet.getSheets(); 
    
    studentSheets[1].getRange("1:" + NUM_OF_ROWS_TO_COPY).copyTo(studentSheets[0].getRange("1:" + NUM_OF_ROWS_TO_COPY), {formatOnly:true}); 
    studentSheets[0].deleteColumn(1);
    
    for(var i = 1; i < columnsNum; ++i) 
      studentSheets[0].setColumnWidth(i, studentSheets[1].getColumnWidth(i + 1)); 
    
    studentSpreadsheet.deleteSheet(studentSheets[1]); 
  }
  
  // Set content
  {
    // Set list name 
    studentSheets[0].setName(LIST_WITH_STUDENT_MARKS_NAME);
    
    // Set header
    for(var i = 1; i <= NUM_OF_LINES_IN_HEADER; ++i)
      studentSheets[0].getRange("A" + i + ":A" + i).setFormula("=IMPORTRANGE(\"" + MAIN_SHEET_LINK + "\";\"" + className + "!B" + (i+headerInd-1) + ":CC" + (i+headerInd-1) + "\")"); 
  
    // Set marks
    studentSheets[0].getRange("A" + NUM_OF_ROWS_TO_COPY + ":A" + NUM_OF_ROWS_TO_COPY + "").setFormula("=IMPORTRANGE(\"" + MAIN_SHEET_LINK + "\";\"" + className + "!B" + row + ":CC" + row + "\")"); 
  } 
    
  // Share student's sheet
  {
    var file  = DriveApp.getFileById(studentSpreadsheet.getId());       // Student's sheet
    var email = classSheet.getRange("A" + row + ":A" + row).getValue(); // Student's email 
    file.addViewer(email);
  }
} 

/*********************************************************************************************************************************************************************************************************/ 
/*
 * \brief  Processes each student in the class.
 *
 * \param[in]  classSheet  Table(sheet) with grades.
 */
function ProceedClass(classSheet) 
{ 
  DriveApp.getRootFolder().createFolder(classSheet.getName()); 
  
  for(var row = NUM_OF_ROWS_TO_COPY; row < START_LINE_OF_SECOND_GROUP; ++row) 
    if(IsEmail(classSheet.getRange("A" + row + ":A" + row).getValue())) 
      ProceedStudent(row, classSheet, 1); 
  
  var rowsNum = classSheet.getLastRow();
  for(var row = START_LINE_OF_SECOND_GROUP; row <= rowsNum; ++row) 
    if(IsEmail(classSheet.getRange("A" + row + ":A" + row).getValue())) 
      ProceedStudent(row, classSheet, START_LINE_OF_SECOND_GROUP); 
} 

/*********************************************************************************************************************************************************************************************************/ 
/*
 * \brief  Main function of the script.
 */
function Main() 
{ 
  var source = SpreadsheetApp.getActiveSpreadsheet().getSheets(); 
  
  //var numberOfClasses = SpreadsheetApp.getActiveSpreadsheet().getNumSheets(); 
  //for(var class = 0; class < numberOfClasses; ++class) 
  ProceedClass(source[0]); 
}
