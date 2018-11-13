//https://medium.com/%D0%BC%D0%B0%D0%BD%D0%B6%D0%B5%D1%82%D1%8B-%D0%B3%D0%B5%D0%B9%D0%BC-%D0%B4%D0%B8%D0%B7%D0%B0%D0%B9%D0%BD%D0%B5%D1%80%D0%B0/google-spreadsheets-cf9d49dd6f27
function depDrop_(range, sourceRange){
 var rule = SpreadsheetApp.newDataValidation().requireValueInRange(sourceRange, true).build();
 range.setDataValidation(rule);
}

function onEdit (){
 var aCell = SpreadsheetApp.getActiveSheet().getActiveCell();
 var aColumn = aCell.getColumn();
 
 if (aColumn == 1 && aCell.getValue() == '') {
 var range = SpreadsheetApp.getActiveSheet().getRange(aCell.getRow(), aColumn + 1);
 range.clearDataValidations();
 range.clearContent();
 return;
 }
 
 
 
 if (aColumn == 1 && SpreadsheetApp.getActiveSheet()){
 var range = SpreadsheetApp.getActiveSheet().getRange(aCell.getRow(), aColumn + 1);
 var sourceRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(aCell.getValue());
 depDrop_(range, sourceRange);
 }
}

/**
 * Return the contiguous Range that contains the given cell.
 * Возвращает смежный диапазон Range, содержащий заданную ячейку
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

//Если вам нужен диапазон, представляющий пересечение, вы можете использовать следующий код:
function GetIntersection(range1, range2) {
  if (range1.getSheet().getSheetId() != range2.getSheet().getSheetId()) {
    return null;
  }
  var sheet = range1.getSheet();
  var startRow = Math.max(range1.getRow(), range2.getRow());
  var endRow = Math.min(range1.getLastRow(), range2.getLastRow());
  var startColumn = Math.max(range1.getColumn(), range2.getColumn());
  var endColumn = Math.min(range1.getLastColumn(), range2.getLastColumn());
  if (startRow > endRow || startColumn > endColumn) {
    return null;
  }
  return sheet.getRange(startRow, startColumn, endRow - startRow + 1, endColumn - startColumn + 1);
}
