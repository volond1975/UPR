function GetNameRangebyA1Notation(spreadsheet,A1Notation){
//var cnf=Config('DataTable');
//var spreadsheet=cnf.spreadsheet;
var namedRanges = spreadsheet.getNamedRanges();
var A1Notation='E4:G4'
// Logger.log(named.getRange().getA1Notation()+'-'+A1Notation);
var namedRangesA1not = namedRanges.map(function(named){return named.getRange().getA1Notation()});
var namedRangesA1Name = namedRanges.map(function(named){return named.getName()});
var namedRangesA1notIndex=namedRangesA1not.indexOf(A1Notation);
Logger.log('getNameRangebyA1Notation Индекс По А1 нотации '+namedRangesA1notIndex);
//var namedRangeNAme=namedRangesA1Name[namedRangesA1notIndex]
//Logger.log('getNameRangebyA1Notation имя По А1 нотации '+namedRangesA1Name[namedRangesA1notIndex])

var res_=namedRangesA1notIndex!==-1?GetNamedRanges(namedRangesA1Name[namedRangesA1notIndex],gid)[0]:false;
Logger.log('isNamedRange '+' '+res_)
return res_
}

//nds_[0].getRange().getA1Notation==A1Notation          
//var res_= (nds_.length===0)?false:nds_[0];

// return res_ 
//} catch (error) {
//  return Logger.log(error)
//}

//}

/**
* Определение CURENT RANGE для ячейки
* Создание именованого DATATABLE_ + нормализованый (getA1Notation) диапазона на основании CURENT RANGE
* Создание DATATABLE на основании именнованого диапазона
*/
function Test_Bind_DateTable_NamedRange(){
var cnf=Config('DataTable')
var spreadsheet=cnf.spreadsheet
var sheet=cnf.activeSheet
var CurentRangeCell=GetContiguousRange("E3",sheet);
var name = "DataTable_" + CurentRangeCell.getA1Notation().replace(/[^A-Z0-9]/g,"");
var bCreate=isNamedRange(name,CurentRangeCell.getA1Notation(),cnf.gid);
Logger.log('bCreate'+name)
var namedRange=(!bCreate)?SpreadsheetApp.getActive().setNamedRange(name, CurentRangeCell):bCreate;
Logger.log(name+ ' в листе ' +cnf.gid +  ' диапазон '+ namedRange.getRange().getA1Notation())
if(namedRange.getRange().getA1Notation()!==CurentRangeCell.getA1Notation()){ 
Logger.log('переопределяем регион  '+name)
//namedRange.setRange(CurentRangeCell);
}
var DataTableNamedRange=dataTableFromArray(namedRange.getRange().getValues())
//Logger.log(DataTableNamedRange.toJSON())
}

function isNamedRange(name,A1Notation,gid){
//var name = name||"DataTable_" + CurentRangeCell.getA1Notation().replace(/[^A-Z0-9]/g,"");

//try {
var nds_=GetNamedRanges(name,gid)
Logger.log(nds_.length)
if(nds_.length===0){
//Logger.log('isNamedRange По имени '+name+' не найдено')
//var namedRanges = spreadsheet.getNamedRanges();
// Logger.log(named.getRange().getA1Notation()+'-'+A1Notation);
//var namedRangesA1not = namedRanges.map(function(named){return named.getRange().getA1Notation()});
//var namedRangesA1Name = namedRanges.map(function(named){return named.getName()});
//var namedRangesA1notIndex=namedRangesA1not.indexOf(A1Notation);
//Logger.log('isNamedRange Индекс По А1 нотации '+namedRangesA1notIndex)

//Logger.log('isNamedRange имя По А1 нотации '+namedRangesA1Name[namedRangesA1notIndex]+' не найдено')

//var res_=namedRangesA1notIndex?GetNamedRanges(namedRangesA1Name[namedRangesA1notIndex],gid)[0]:false;
var res_=GetNameRangebyA1Notation(spreadsheet,A1Notation)
Logger.log('isNamedRange '+name+' '+res_)
}
else if(nds_[0].getRange().getA1Notation===A1Notation){
var res_=nds_[0]
}
else {
var res_=false
}


/**
 * Produce a dataTable object suitable for use with Charts, from
 * an array of rows (such as you'd get from Range.getValues()).
 * Assumes labels are in row 0, and the data types in row 1 are
 * representative for the table.
 * https://gist.github.com/mogsdad/8714493 
 * @param {Array} data  Array of table rows
 *
 @ @returns {DataTable} Refer to GAS documentation
 */
function dataTableFromArray( data ) {
Logger.log(data)
 var dataTable = Charts.newDataTable();
  for (var col=0; col<data[0].length; col++) {
    var label = data[0][col];
    var firstCell = data[1][col];
    if (typeof firstCell == 'string')
      dataTable.addColumn(Charts.ColumnType.STRING, label);
    else
      dataTable.addColumn(Charts.ColumnType.NUMBER, label);
  }
  for (var row = 1; row < data.length; row++) {
    dataTable.addRow(data[row]);
  }  
  return dataTable.build();
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
function GetNamedRanges(NameRange,gid,patern,ss){
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

 
Logger.log(res_.length); 

 
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


//Пользовательская функция, возвращающая адрес A1 именованного диапазона:
/**
 * Returns sheet by gid
 * Пользовательская функция, возвращающая адрес A1 именованного диапазона:
 * @param {number} gid The gid of the sheet.
 * @return {GoogleAppsScript.Spreadsheet.Sheet} sheet Лист Таблицы
 * @customfunction
 */
function MyGetRangeByName(NameRange) {  // just a wrapper
 var sh=Config("Финансы").activeSheet
  
  
  return sh.namedRanges.map(function(named){return named.getRange().getA1Notation()})};}
  function dinamicNamedRange(){
var sheetName="Списки"
var NameRange='table_Подстановка'
var ssDirection=SpreadsheetApp.Direction
var cnf=Config(sheetName)
var spreadsheet=cnf.spreadsheet
var sheet=cnf.activeSheet
var ndRange=GetNamedRanges(cnf.gid,NameRange,sheet,spreadsheet)[0]
var nRange=ndRange.getRange()
var RangeLO=sheet.getRange(nRange.getRow(),nRange.getColumn())

//Logger.log(RangeLO.getNextDataCell(ssDirection.NEXT).getA1Notation())//Направление увеличения индексов столбцов.
var newRangeLOHeader=GetRange(Config().spreadsheetId,sheetName, nRange.getRow(), nRange.getRow(), nRange.getColumn(), RangeLO.getNextDataCell(ssDirection.NEXT).getColumn())

var minrow=nRange.getRow()
for (var i = nRange.getColumn(); i <= RangeLO.getNextDataCell(ssDirection.NEXT).getColumn(); i++) {
//Здесь мы определяем количество строк на которое мог расширится диапазон
minrow=sheet.getRange(newRangeLOHeader.getRow(),i).getNextDataCell(ssDirection.DOWN).getRow()>minrow?sheet.getRange(newRangeLOHeader.getRow(),i).getNextDataCell(ssDirection.DOWN).getRow():minrow


};
minrow=getLastRowOfRange(SpreadsheetApp.getActiveSheet().getRange(nRange.getRow(),nRange.getColumn(),sheet.getLastRow(),RangeLO.getNextDataCell(ssDirection.NEXT).getColumn()))
Logger.log(minrow)
//Logger.log(RangeLO.getNextDataCell(ssDirection.UP).getA1Notation())	//Направление убывающих индексов строк.
//Logger.log(RangeLO.getNextDataCell(ssDirection.DOWN).getA1Notation())	//Направление увеличения индексов строк.
//Logger.log(RangeLO.getNextDataCell(ssDirection.PREVIOUS).getA1Notation())	//Направление убывающих индексов столбцов.


var newRangeLO=GetRange(Config().spreadsheetId, sheetName, RangeLO.getRow(), minrow, RangeLO.getColumn(), RangeLO.getNextDataCell(ssDirection.NEXT).getColumn())


ndRange.setRange(newRangeLO).getRange()

}

/* var targetRange = SpreadsheetApp.getActiveSheet().getRange(1,1,4,3); 
var lastRowOfTargetRange = getLastRowOfRange(targetRange); 


function getLastRowOfRange (range) { 

   var rangeValues = range.getValues(); 
    var columns = range.getNumColumns(); 
    var lastRowHasContent = 0; 

    for (var i = 0; i < columns; i++) { 
    var rowsCount = 0; 
    while (rangeValues[rowsCount][i]) { 
     rowsCount++; 
    } 
    if (rowsCount > lastRowHasContent) 
     lastRowHasContent = rowsCount; 
    } 

    return lastRowHasContent; 

} 
*/