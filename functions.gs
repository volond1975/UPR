function CopySpreadsheet(spreadsheetId, newName) {
  SpreadsheetApp.openById(spreadsheetId).copy(newName);
}

function ExportSpreadsheetToPdf(spreadsheetId, pdfFilename) {
  var spreadsheetFile = DriveApp.getFileById(spreadsheetId);
  pdfFilename = pdfFilename || spreadsheetFile.getName() + ".pdf";
  DriveApp.getRootFolder().createFile(spreadsheetFile.getAs('application/pdf')).setName(pdfFilename);
}

function GetSpreadsheet(spreadsheetId) {
  var spreadsheetId=spreadsheetId||SpreadsheetApp.getActiveSpreadsheet().getId();
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  Logger.log(spreadsheetId)
  return spreadsheet;
}

function GetSheet(spreadsheetId, sheetName) {
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName(sheetName);
  return sheet;
}
/**
 * Returns sheet by gid
 *
 * @param {number} gid The gid of the sheet.
 * @return {GoogleAppsScript.Spreadsheet.Sheet} sheet Лист Таблицы
 * @customfunction
 */


function GetSheetByGid(ss,gid){
  gid = +gid || 0;
  var res_ = undefined;
  var sheets_ = ss.getSheets();
  for(var i = sheets_.length; i--; ){
    if(sheets_[i].getSheetId() === gid){
      res_ = sheets_[i];
      break;
    }
  }
  return res_;
}















function CheckAndFixColumnCount(sheet, count) {
  var lastColumnPosition = sheet.getMaxColumns();
  if (count > lastColumnPosition) {
    sheet.insertColumnsAfter(lastColumnPosition, count - lastColumnPosition);
  }
}

function CheckAndFixRowCount(sheet, count) {
  var lastRowPosition = sheet.getMaxRows();
  if (count > lastRowPosition) {
    sheet.insertRowsAfter(lastRowPosition, count - lastRowPosition);
  }
}

function GetRange(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn,NamedRange) {
  var sheet = GetSheet(spreadsheetId, sheetName);
  CheckAndFixRowCount(sheet, toRow);
  CheckAndFixColumnCount(sheet, toColumn);

  return sheet.getRange(fromRow, fromColumn, toRow - fromRow + 1, toColumn - fromColumn + 1);
}

function CheckEquality(table, faultValue) {
  if (typeof faultValue === "undefined") {
    faultValue = "varies";
  }

  var topLeft = table[0][0];
  for (var i in table) {
    for (var j in table[i]) {
      if (topLeft != table[i][j]) {
        return faultValue;
      }
    }
  }

  return topLeft;
}

function GetColumnWidth(spreadsheetId, sheetName, fromPosition, toPosition) {
  var sheet = getSheet(spreadsheetId, sheetName);
  checkAndFixColumnCount(sheet, toPosition);

  var columnWidth = sheet.getColumnWidth(fromPosition);
  for (var i = fromPosition; i <= toPosition; i++) {
    var width = sheet.getColumnWidth(i);
    if (width != columnWidth) {
      return -1;
    }
  }
  return columnWidth;
}

function SetColumnWidth(spreadsheetId, sheetName, fromPosition, toPosition, width) {
  var sheet = getSheet(spreadsheetId, sheetName);
  checkAndFixColumnCount(sheet, toPosition);
  
  for (var i = fromPosition; i <= toPosition; i++) {
    sheet.setColumnWidth(i, width);
  }
}

function GetRowHeight(spreadsheetId, sheetName, fromPosition, toPosition) {
  var sheet = getSheet(spreadsheetId, sheetName);
  checkAndFixRowCount(sheet, toPosition);

  var rowHeight = sheet.getRowHeight(fromPosition);
  for (var i = fromPosition; i <= toPosition; i++) {
    var height = sheet.getRowHeight(i);
    if (height != rowHeight) {
      return -1;
    }
  }
  return rowHeight;
}

function GetRowHeight(spreadsheetId, sheetName, fromPosition, toPosition, height) {
  var sheet = getSheet(spreadsheetId, sheetName);
  checkAndFixRowCount(sheet, toPosition);

  for (var i = fromPosition; i <= toPosition; i++) {
    sheet.setRowHeight(i, height);
  }
}

function GetBorders(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn, top, left, bottom, right, vertical, horizontal, htmlColor, style) {
  var borderStyle = SpreadsheetApp.BorderStyle.SOLID;
  if (style == "xlDash") {
    borderStyle = SpreadsheetApp.BorderStyle.DASHED;
  } else if (style == "xlDot") {
    borderStyle = SpreadsheetApp.BorderStyle.DOTTED;
  }
  getRange.apply(this, arguments).setBorder(top, left, bottom, right, vertical, horizontal, htmlColor, borderStyle);
}

function getBackgroundColor(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn) {
  return checkEquality(getRange.apply(this, arguments).getBackgrounds());
}

function setBackgroundColor(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn, htmlColor) {
  getRange.apply(this, arguments).setBackground(htmlColor);
}

function GetFontColor(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn) {
  return checkEquality(getRange.apply(this, arguments).getFontColors());
}

function SetFontColor(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn, htmlColor) {
  getRange.apply(this, arguments).setFontColor(htmlColor);
}

function GetFontWeight(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn) {
  return checkEquality(getRange.apply(this, arguments).getFontWeights());
}

function sSetFontWeight(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn, weight) {
  getRange.apply(this, arguments).setFontWeight(weight);
}

function GetFontStyle(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn) {
  return checkEquality(getRange.apply(this, arguments).getFontStyles());
}

function SetFontStyle(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn, style) {
  getRange.apply(this, arguments).setFontStyle(style);
}

function GetFontLine(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn) {
  return checkEquality(getRange.apply(this, arguments).getFontLines());
}

function SetFontLine(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn, line) {
  getRange.apply(this, arguments).setFontLine(line);
}

function GetFontSize(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn) {
  return checkEquality(getRange.apply(this, arguments).getFontSizes());
}

function SetFontSize(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn, size) {
  getRange.apply(this, arguments).setFontSize(size);
}

function GetHorizontalAlignment(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn) {
  return checkEquality(getRange.apply(this, arguments).getHorizontalAlignments());
}

function SetHorizontalAlignment(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn, alignment) {
  getRange.apply(this, arguments).setHorizontalAlignment(alignment);
}

function GetVerticalAlignment(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn) {
  return checkEquality(getRange.apply(this, arguments).getVerticalAlignments());
}

function SetVerticalAlignment(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn, alignment) {
  getRange.apply(this, arguments).setVerticalAlignment(alignment);
}

function GetNumberFormat(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn) {
  return checkEquality(getRange.apply(this, arguments).getNumberFormats());
}

function SetNumberFormat(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn, format) {
  getRange.apply(this, arguments).setNumberFormat(format);
}

function GetWrap(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn) {
  return checkEquality(getRange.apply(this, arguments).getWraps());
}

function SetWrap(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn, wrap) {
  getRange.apply(this, arguments).setWrap(wrap);
}

function GetValue(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn) {
  var value = checkEquality(getRange.apply(this, arguments).getValues(), null);
  return value == null ? null : "" + value;
}

function SetValue(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn, value) {
  getRange.apply(this, arguments).setValue(value);
}

function GetDisplayValue(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn) {
  return checkEquality(getRange.apply(this, arguments).getDisplayValues(), null);
}

function GetFormula(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn) {
  return checkEquality(getRange.apply(this, arguments).getFormulas());
}

function SetFormula(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn, formula) {
  getRange.apply(this, arguments).setFormula(formula);
}

function GetFormulaR1C1(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn) {
  return checkEquality(getRange.apply(this, arguments).getFormulasR1C1());
}

function SetFormulaR1C1(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn, formula) {
  getRange.apply(this, arguments).setFormulaR1C1(formula);
}

function ClearRange(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn) {
  getRange.apply(this, arguments).clear();
}

function ClearRangeContent(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn) {
  getRange.apply(this, arguments).clearContent();
}

function ClearRangeFormat(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn) {
  getRange.apply(this, arguments).clearFormat();
}

function MergeRange(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn) {
  getRange.apply(this, arguments).merge();
}

function MergeRangeAcross(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn) {
  getRange.apply(this, arguments).mergeAcross();
}

function UnmergeRange(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn) {
  getRange.apply(this, arguments).breakApart();
}

function ShiftRangeToRight(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn) {
  var sheet = getRange.apply(this, arguments).getSheet();
  sheet.insertColumnsBefore(fromColumn, toColumn - fromColumn + 1);
}

function ShiftRangeDown(spreadsheetId, sheetName, fromRow, toRow, fromColumn, toColumn) {
  var sheet = getRange.apply(this, arguments).getSheet();
  sheet.insertRowsBefore(fromRow, toRow - fromRow + 1);
}

//Работа с именованными диапазонами
//========================================================================
/**
 * Returns namedRanges or namedRange
 *
 * @param {number} gid The gid of the sheet.
 * @return {GoogleAppsScript.Spreadsheet.Sheet} sheet Лист Таблицы
 * @customfunction
 */ 
function GetNamedRanges(gid,NameRange,patern,ss){
   var ss=ss||SpreadsheetApp.getActiveSpreadsheet();
   var gid=+gid||ss.getActiveSheet().getSheetId();
   var sheet=GetSheetByGid(ss,gid)
   var activeCellSheet=GetSheetByGid(ss,gid)
   
   
   
   
   var namedRanges = ss.getNamedRanges();
  
   
   var res_ = [];
  if (NameRange === undefined){
  

    
    
    
    
    
   for(var i = namedRanges.length; i--;){
//     Logger.log(sheet.getName()===namedRanges[i].getRange().getSheet().getName());
    if(sheet.getName()===namedRanges[i].getRange().getSheet().getName()){
     //var res_ = [1, 2];
//res_ = res_.concat(namedRanges[i]);
     
    res_.push(namedRanges[i]);
    
      //break;
    }
  }
  }
  else{
   for(var i = namedRanges.length; i--;){
    if(namedRanges[i].getName() === NameRange && sheet.getName()===namedRanges[i].getRange().getSheet().getName()){
      res_.push(namedRanges[i]);
 //   Logger.log(namedRanges[i]) 
    break;
    }
  }
  };

 
  //Logger.log(res_.length); 

 
  return res_;  
}

//Работа с именованными диапазонами
//========================================================================
/**
 * Returns Имена Именованых Диапазонов в книге 
 * @param {number} gid The gid of the sheet.
 * @return {GoogleAppsScript.Spreadsheet.Sheet} sheet Лист Таблицы
 * @customfunction
 */ 

function GetNamedRangesNames(gid,NameRange,patern,ss){
   var ss=ss||SpreadsheetApp.getActiveSpreadsheet();
   var gid=+gid||ss.getActiveSheet().getSheetId();
   var sheet=GetSheetByGid(ss,gid)
   var activeCellSheet=GetSheetByGid(ss,gid)
   var namedRanges = ss.getNamedRanges();
  
  
  return namedRanges.map(function(named){return named.getName()})}
/**
 * Returns Адресса в А1 нотациии Именованых Диапазонов в книге or namedRange
 *
 * @param {number} gid The gid of the sheet.
 * @return {GoogleAppsScript.Spreadsheet.Sheet} sheet Лист Таблицы
 * @customfunction
 */ 
function getNamedRangesA1Notation(gid,NameRange,patern,ss){
   var ss=GetSpreadsheet(ss);
  // var gid=+gid||ss.getActiveSheet().getSheetId();
  // var sheet=GetSheetByGid(ss,gid)
 //  var activeCellSheet=GetSheetByGid(ss,gid)
   var namedRanges = ss.getNamedRanges();
  return namedRanges.map(function(named){return named.getRange().getA1Notation()})}

 
function GetNamedRangesA1Notation(gid,NameRange,patern,ss){
   var ss=GetSpreadsheet(ss);
  // var gid=+gid||ss.getActiveSheet().getSheetId();
  // var sheet=GetSheetByGid(ss,gid)
 //  var activeCellSheet=GetSheetByGid(ss,gid)
   var namedRanges = ss.getNamedRanges();
  return namedRanges.map(function(named){return named.getRange().getA1Notation()})}
/**
 * Returns Имена и Адресса в А1 нотациии Именованых Диапазонов в книге or namedRange
 *
 * @param {number} gid The gid of the sheet.
 * @return {GoogleAppsScript.Spreadsheet.Sheet} sheet Лист Таблицы
 * @customfunction
 */ 

function GetNamedRangesNameAndA1Notation(gid,NameRange,patern,ss){
   var ss=GetSpreadsheet(ss)
   var gid=+gid||ss.getActiveSheet().getSheetId();
   var sheet=GetSheetByGid(ss,gid)
   var activeCellSheet=getSheetByGid(ss,gid)
   var namedRanges = ss.getNamedRanges();
  
  
  return namedRanges.map(function(named){return [named.getName(),named.getSheet().getName(),named.getRange().getA1Notation()]})}
//function testRanges() {
//  var range1 = getContiguousRange("C3").getA1Notation();
//  var range2 = getContiguousRange("B8").getA1Notation();
//  debugger;
//}


/**
 * Return the contiguous Range that contains the given cell.
 *
 * @param {String} cellA1 Location of a cell, in A1 notation.
 * @param {Sheet} sheet   (Optional) sheet to examine. Defaults
 *                          to "active" sheet.
 *
 * @return {Range} A Spreadsheet service Range object.
 */
function GetContiguousRange(cellA1,sheet) {
  // Check for parameters, handle defaults, throw error if required is missing
  if (arguments.length < 2) 
    sheet = SpreadsheetApp.getActiveSheet();
  if (arguments.length < 1)
    throw new Error("getContiguousRange(): missing required parameter.");
  
  // A "contiguous" range is a rectangular group of cells whose "edge" contains
  // cells with information, with all "past-edge" cells empty.
  // The range will be no larger than that given by "getDataRange()", so we can
  // use that range to limit our edge search.
  var fullRange = sheet.getDataRange();
  var data = fullRange.getValues();
  
  // The data array is 0-based, but spreadsheet rows & columns are 1-based.
  // We will make logic decisions based on rows & columns, and convert to
  // 0-based values to reference the data.
  var topLimit = fullRange.getRowIndex(); // always 1
  var leftLimit = fullRange.getColumnIndex(); // always 1
  var rightLimit = fullRange.getLastColumn();
  var bottomLimit = fullRange.getLastRow();
  
  // is there data in the target cell? If no, we're done.
  var contiguousRange = SpreadsheetApp.getActiveSheet().getRange(cellA1);
  var cellValue = contiguousRange.getValue();
  if (cellValue = "") return contiguousRange;
  
  // Define the limits of our starting dance floor
  var minRow = contiguousRange.getRow();
  var maxRow = minRow;
  var minCol = contiguousRange.getColumn();
  var maxCol = minCol;
  var chkCol, chkRow;  // For checking if the edge is clear

  // Now, expand our range in one direction at a time until we either reach
  // the Limits, or our next expansion would have no filled cells. Repeat
  // until no direction need expand.
  var expanding;
  do {
    expanding = false;
    // Move it to the left
    if (minCol > leftLimit) {
      chkCol = minCol - 1;
      for (var row = minRow; row <= maxRow; row++)  {
        if (data[row-1][chkCol-1] != "") {
          expanding = true;
          minCol = chkCol; // expand left 1 column
          break;
        }
      }
    }
    
    // Move it on up
    if (minRow > topLimit) {
      chkRow = minRow - 1;
      for (var col = minCol; col <= maxCol; col++)  {
        if (data[chkRow-1][col-1] != "") {
          expanding = true;
          minRow = chkRow; // expand up 1 row
          break;
        }
      }
    }
    
    // Move it to the right
    if (maxCol < rightLimit) {
      chkCol = maxCol + 1;
      for (var row = minRow; row <= maxRow; row++)  {
        if (data[row-1][chkCol-1] != "") {
          expanding = true;
          maxCol = chkCol; // expand right 1 column
          break;
        }
      }
    }
    
    // Then get on down
    if (maxRow < bottomLimit) {
      chkRow = maxRow + 1;
      for (var col = minCol; col <= maxCol; col++)  {
        if (data[chkRow-1][col-1] != "") {
          expanding = true;
          maxRow = chkRow; // expand down 1 row
          break;
        }
      }
    }
       
  } while (expanding);  // Lather, rinse, repeat
  
  // We've found the extent of our contiguous range - return a Range object.
  return sheet.getRange(minRow, minCol, (maxRow - minRow + 1), (maxCol - minCol + 1))
}



function RangeIntersect (R1, R2) {
  return (R1.getLastRow() >= R2.getRow()) && (R2.getLastRow() >= R1.getRow()) && (R1.getLastColumn() >= R2.getColumn()) && (R2.getLastColumn() >= R1.getColumn());
}


