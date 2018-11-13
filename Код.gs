function onEdit(){
var sheetName="Списки"

}

function hello() {
  Logger.log("Hello, " + mondo);
}






function ActCell(){
var sheets=SpreadsheetApp.getSelection().getCurrentCell().getA1Notation()
Logger.log(sheets)

var aCell = SpreadsheetApp.getActiveSheet().getActiveCell();
//Browser.msgBox(aCell.getSheet().getName() +"!"+     aCell.getA1Notation())
Logger.log(aCell.getSheet().getName() +"!"+     aCell.getA1Notation())

var userid = Session.getActiveUser().getEmail()
Logger.log(userid)
}

function getCell() {
var cnf=Config()
var spreadsheet=cnf.spreadsheet
  var actvC = SpreadsheetApp.getActiveSpreadsheet().getName();
  //var txt_actvC = actvC.getA1Notation();
   Logger.log(actvC)
 // Logger.log(actvC + ' ' +actvC.getSheet().getName() +"!"+     actvC.getA1Notation())
var sheets=SpreadsheetApp.getActiveSpreadsheet().getSheets()

  actvC = SpreadsheetApp.getActiveSheet().getActiveCell();
  txt_actvC = actvC.getA1Notation();
  Logger.log(SpreadsheetApp.getActiveSheet().getName() +"!"+     txt_actvC);
  
  
  
  
}


function jump() {
  var sheet = SpreadsheetApp.getActiveSheet();
  Logger.log(sheet.getName())
  return sheet.setActiveRange(
    sheet.getRange(
      sheet.getDataRange().getHeight() + 1, 1)
  );
}



function GetNamedRangeByNameAndSheet(NameRange,sheet,spreadsheet){
var namedRanges = spreadsheet.getNamedRanges();
  
   
   var res_ = [];
     for(var i = namedRanges.length; i--;){
    if(namedRanges[i].getName() === NameRange && sheet.getName()===namedRanges[i].getRange().getSheet().getName()){
      res_.push(namedRanges[i]);
 
    break;
    }
  };
return res_[0]







}





















function testRanges() {
  var range1 = GetContiguousRange("C3").getA1Notation();
  //var range2 = GetContiguousRange("B8").getA1Notation();
  Logger.log(range1)
}








var spreadsheet = SpreadsheetApp.getActive();
var sheet = SpreadsheetApp.getActiveSheet();

function updateValues(){
  dataRangeSearch = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn());
  dataSearch = dataRangeSearch.getValues().reduce(function (a, b) {
    return a.concat(b);
  });;
}
updateValues();

function findValue(fieldName){
  var row = dataSearch.indexOf(fieldName);
  if (row == -1){ // Variable could not be found
    SpreadsheetApp.getUi().alert('I couldn\'t find field name "'+fieldName+'"');
    return "";
  } else {
    
    return (row+1)/sheet.getLastColumn()|0; //Return the value of the field to the right of the matched string
  }
}
function TestFindValue(){
var HeaderNameRange=getNamedRanges('903053947','table_Ф_Наличие','table_Ф_Наличие',ss)//sheet.NamedRange('table_Ф_Наличие').
Logger.log(findValue("Модель"))
Logger.log(sheet.getName())
}














