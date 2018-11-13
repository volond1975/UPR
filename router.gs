 ///Mcbr-v4SsYKJP7JMohttAZyz3TLx7pV4j
 //http://ramblings.mcpher.com/Home/excelquirks/gassnips/fiddler 
//https://mashe.hawksey.info/2015/10/google-sheets-as-a-database-authenticated-insert-with-apps-script-using-execution-api/
//https://mashe.hawksey.info/2018/02/google-apps-script-patterns-writing-rows-of-data-to-google-sheets/
 
function getFiddlerByNamedRange(NameRange,gid,sheetname,ss){
  
var cnf=Config(sheetname)
var spreadsheet=cnf.spreadsheet 
var sheet=cnf.activeSheet
var DataTable=GetNamedRanges(NameRange,gid,sheetname,ss)
Logger.log(NameRange+' '+DataTable)
//Logger.log(DataTable[0].getRange())
//var sheet=DataTable[0].getRange().getSheet()
// sheet.activate()

var fiddlerDataTable = new cUseful.Fiddler(sheet)
 return fiddlerDataTable.setValues(DataTable[0].getRange().getValues())
}

function addTimeShtamp(fiddler,name){
  fiddler
  .insertColumn (name , fiddler.getHeaders()[0])
  .mapColumn(name, function (value,properties) {
    return new Date().getTime();
  });
  return fiddler
}
function addIdKeyFiddler(fiddler,defaultKEY){ 
 fiddler.insertColumn ('_id') 
 fiddler.mapRows(function (row,properties) {
 var defaultArr=  defaultKEY.map(function(i){return row[i] })
 row._id +=defaultArr.join("_")
                 //row._id += (row["Модель"] + '_'+row.Запах + '_' +row.Место );
    return row;
  });
   return fiddler
 }
 function Test1(){
  var DataTable=GetNamedRanges('DataTable_Поступление','0')}
 function Test(){
 //var DataTable=GetNamedRanges('DataTable_Поступление','0')
// var DataTable_F=GetNamedRanges('DataTable_Ф_Поступление','903053947')
 // Logger.log(nr[0].getRange().getA1Notation())
 // Logger.log(nrt[0].getRange().getA1Notation())
//  Logger.log( GetIntersection(nr[0].getRange(), nrt[0].getRange()).getValues())
//  Logger.log(getNamedRangeHeader('DataTable_DataTable','1637697477').getValues()) 
//  Logger.log(getNamedRangeDataRange('DataTable_DataTable','1637697477').getValues()) 
//  Logger.log(getNamedRangeNameColumn('DataTable_DataTable','1637697477','Подстановка').getValues())
//  Logger.log(getNamedRangeNameColumnDataRange('DataTable_DataTable','1637697477','Подстановка').getValues())
 //Logger.log(getNamedRangeRows('DataTable_DataTable','1637697477','2').getValues())
// Logger.log(getNamedRangeRowsFindValue('DataTable_DataTable','1637697477','Подстановка','Новый').getValues())
 //Таблица для вставки
// var fiddlerDataTable = new cUseful.Fiddler()
// fiddlerDataTable.setValues(DataTable[0].getRange().getValues())
 
 var defaultKEY=["Место","Модель","Запах"] 
   //Форма
// var fiddlerDataTable_F = new cUseful.Fiddler()
// fiddlerDataTable_F.setValues(DataTable_F[0].getRange().getValues())
 var masterFiddler=getFiddlerByNamedRange('DataTable_Поступление','0','Поступление')
  
 addIdKeyFiddler(masterFiddler,defaultKEY)
 var updateFiddler =addTimeShtamp(getFiddlerByNamedRange('DataTable_Ф_Поступление','903053947','[Ф] Поступление'),"Дата" )
  addIdKeyFiddler(updateFiddler,defaultKEY)
 var dataDataTable = updateFiddler.getData();
  // Logger.log(dataDataTable)  
   
   

 

   
    // now we drive the update off the update inserting as required
  updateFiddler.getData().forEach (function (updateRow) {
    
    // get matches on this key
    var matches = masterFiddler.selectRows ("_id" , function (value) {
      
      return value === updateRow._id;
    });
      Logger.log('matches')
      Logger.log(matches)
    // it's an existing item
    if (matches.length) {
    
      // assume its okay to have duplicate keys and update them all for this example
      matches.forEach (function (match) {
        var Count=masterFiddler.getData()[match].Количество
        Logger.log('before insert Count '+masterFiddler.getData()[match].Количество)
        masterFiddler.getData()[match] = updateRow;
         masterFiddler.getData()[match].Количество=masterFiddler.getData()[match].Количество+Count
          Logger.log('after insert Count '+masterFiddler.getData()[match].Количество)
      });
    }
    // its an update
    else {
      // insert 1 row at end
        Logger.log('insert 1 row at end')
      //masterFiddler.insertRows ( null , 1 , updateRow);
      masterFiddler.insertRows ( )
    }
    
  });
 
  // might want to sort the master by key now
 masterFiddler.setData(masterFiddler.sort ("_id"));
  masterFiddler.filterColumns (function (name,properties) {
    return name !== '_id';
  });
 Logger.log(masterFiddler.getData())    
   
  // and write the updated data
var  nr=GetNamedRanges('DataTable_Поступление','0')[0]
var sheet=nr.getSheet()
// sheet.activate()
// masterFiddler.getRange(GetNamedRanges('DataTable_Поступление','0')[0].getRange())
//  .setValues(masterFiddler.createValues())  
   
 masterFiddler.dumpValues (sheet);

//  var dataDataTable_F = masterFiddler.getData();
 //  Logger.log(dataDataTable_F)   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
/* var Count=dataDataTable_F[0].Количество 
//Logger.log(Count) 
fiddlerDataTable_F.filterColumns (function (name,properties) {
    return name !== "Количество" ;
  });
// Logger.log(fiddlerDataTable_F.getData()) 
 
  fiddlerDataTable_F.mapRows(function (row,properties) {
    row.Модель += ( '_'+row.Запах + '(' +row.Место+')' );
    return row;
  });
   
 
 
fiddlerDataTable_F.mapHeaders (function(name,properties) {
    return name.replace('Модель','_id');
  });
// Logger.log(fiddlerDataTable_F.getData())
// var Count=dataDataTable_F[0].Количество
  // var defaultKEY={'Модель','Запах','Место'}
fiddlerDataTable_F.filterColumns (function (name,properties) {
    return name === '_id' ;
    }); 
 
 fiddlerDataTable_F.insertColumn ('timeshtamp','_id') 
 fiddlerDataTable_F.insertColumn ('Количество') 
// Logger.log(fiddlerDataTable_F.getData())
//data[0].Подстановка = 'Port Moresby area'
//Logger.log(data)
 //fiddler
 // .getRange(DataTable.getRange())
 // .setValues( fiddler.createValues())

*/
 }

function updater () {
  
  var MASTERID = "1181bwZspoKoP98o4KuzO0S11IsvE59qCwiw4la9kL4o";
  var UPDATEID = MASTERID;
  var MASTERNAME = "master";
  var UPDATENAME = "updates";
  
  // get all the master and update data
  //var masterFiddler = new cUseful.Fiddler (SpreadsheetApp.openById(MASTERID).getSheetByName(MASTERNAME));
 // var updateFiddler = new cUseful.Fiddler (SpreadsheetApp.openById(UPDATEID).getSheetByName(UPDATENAME));
  var masterFiddler=getFiddlerByNamedRange('DataTable_Ф_Поступление','903053947')  
  var updateFiddler=getFiddlerByNamedRange('DataTable_Поступление','0')
  
  // now we drive the update off the update inserting as required
  updateFiddler.getData().forEach (function (updateRow) {
    
    // get matches on this key
    var matches = masterFiddler.selectRows ("_id" , function (value) {
      return value === updateRow._id;
    });
    
    // it's an existing item
    if (matches.length) {
      // assume its okay to have duplicate keys and update them all for this example
      matches.forEach (function (match) {
        masterFiddler.getData()[match] = updateRow;
      });
    }
    // its an update
    else {
      // insert 1 row at end
      masterFiddler.insertRows ( null , 1 , updateRow);
    }
    
  });
 
  // might want to sort the master by key now
  masterFiddler.setData(masterFiddler.sort ("_id"));
  
  // and write the updated data
  masterFiddler.dumpValues ();

}













 
function showFiddler (fiddlerObject , outputRange) {
  
  // clear and write result 
  outputRange
  .getSheet()
  .clearContents();
  
  fiddlerObject
  .getRange(outputRange)
  .setValues(fiddlerObject.createValues());
}




function getNamedRangeHeader(NameRange,gid,patern,ss){
 nr=GetNamedRanges('DataTable_DataTable','1637697477')
 sheet= getNamedRangeSheet(nr[0])
 var fromData=nr[0].getRange()
 var fromCellRow=nr[0].getRange().getRow()
 var fromCellCol=nr[0].getRange().getColumn()
 var fromColsN = fromData.getNumColumns() 
 var fromRowsN = fromData.getNumRows() 
 var fromRow1 = sheet.getRange(fromCellRow, fromCellCol, 1, fromColsN); 
 return fromRow1
}








function getNamedRangeDataRange(NameRange,gid,patern,ss){
 nr=GetNamedRanges('DataTable_DataTable','1637697477')
 sheet= getNamedRangeSheet(nr[0])
 var fromData=nr[0].getRange()
 var fromCellRow=nr[0].getRange().getRow()
 var fromCellCol=nr[0].getRange().getColumn()
 var fromColsN = fromData.getNumColumns() 
 var fromRowsN = fromData.getNumRows() 
 var fromRows = sheet.getRange(fromCellRow+1, fromCellCol, fromRowsN, fromColsN); 
 return fromRows
}

function getNamedRangeRows(NameRange,gid,Rows,ss){
 var splRows=Rows.split(':')
 Logger.log(splRows.length)
 nr=GetNamedRanges('DataTable_DataTable','1637697477')
 sheet= getNamedRangeSheet(nr[0])
 var fromData=nr[0].getRange()
 var fromCellRow=nr[0].getRange().getRow()
 var fromCellCol=nr[0].getRange().getColumn()
 var fromColsN = fromData.getNumColumns() 
 var fromRowsN = fromData.getNumRows() 
 if (splRows.length===2){
   st=+splRows[0]+fromCellRow
   end=(+splRows[1]-st)+fromCellRow+st
   Logger.log('From-'+st+'; To'+( end))
   var fromRows = sheet.getRange(st, fromCellCol, end-st+1, fromColsN);  
 }
  else{
     st=+Rows+fromCellRow
    var fromRows = sheet.getRange(st, fromCellCol, 1, fromColsN); }
 return fromRows
}




function getNamedRangeRowsFindValue(NameRange,gid,fieldName,Value,ss){
  updateValues(getNamedRangeNameColumn(NameRange,gid,fieldName,ss))
  var row = dataSearch.indexOf(Value);
  if (row == -1){ // Variable could not be found
    SpreadsheetApp.getUi().alert('I couldn\'t find field name "'+fieldName+'"');
    return "";
  } else {
    
    return getNamedRangeRows(NameRange,gid,(row)+'',ss); //Return the value of the field to the right of the matched string
  }
}






function getNamedRangeNameColumn(NameRange,gid,NameColumn,ss){
 nr=GetNamedRanges('DataTable_DataTable','1637697477')
 nrheader=getNamedRangeHeader('DataTable_DataTable','1637697477')
 indexCol=nrheader.getValues()[0].indexOf(NameColumn)
 sheet= getNamedRangeSheet(nr[0])
 var fromData=nr[0].getRange()
 var fromCellRow=nr[0].getRange().getRow()
 var fromCellCol=nr[0].getRange().getColumn()
 var fromColsN = fromData.getNumColumns() 
 var fromRowsN = fromData.getNumRows() 
 var fromCol = sheet.getRange(fromCellRow, fromCellCol+indexCol, fromRowsN, 1); 
 return fromCol
}

function getNamedRangeNameColumnDataRange(NameRange,gid,NameColumn,ss){
 nr=GetNamedRanges('DataTable_DataTable','1637697477')
 nrheader=getNamedRangeHeader('DataTable_DataTable','1637697477')
 indexCol=nrheader.getValues()[0].indexOf(NameColumn)
 sheet= getNamedRangeSheet(nr[0])
 var fromData=nr[0].getRange()
 var fromCellRow=nr[0].getRange().getRow()
 var fromCellCol=nr[0].getRange().getColumn()
 var fromColsN = fromData.getNumColumns() 
 var fromRowsN = fromData.getNumRows() 
 var fromCol = sheet.getRange(fromCellRow+1, fromCellCol+indexCol, fromRowsN, 1); 
 return fromCol
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
//  ss.setActiveSheet(sheet)
   
   var res_ = [];
  if (NameRange === undefined){
     for(var i = namedRanges.length; i--;){

    if(sheet.getName()===namedRanges[i].getRange().getSheet().getName()){
    
    res_.push(namedRanges[i]);
    
    
    }
  }
 Logger.log(res_)  
  }
  else{
   for(var i = namedRanges.length; i--;){
    if(namedRanges[i].getName() === NameRange && sheet.getName()===namedRanges[i].getRange().getSheet().getName()){
      res_.push(namedRanges[i]);

    break;
    }
  }
  };

//Logger.log(res_.length); 

  return res_;  
}

function getNamedRangesRangeFullA1Not(ss) {  
var ss=ss||SpreadsheetApp.getActiveSpreadsheet();

   var namedRanges = ss.getNamedRanges();
 
  return namedRanges.map(function(named){return getNamedRangeRangeFullA1Not(named)})};
  










 function getNamedRangeRangeA1Not(namedRange,bNormalize){
 if (bNormalize === undefined){var bNormalize=false};
 var nr= (bNormalize)?namedRange.getRange().getA1Notation().replace(/[^A-Z0-9]/g,""):namedRange.getRange().getA1Notation()
 return nr
 } ;
 
 function getNamedRangeRangeFullA1Not(namedRange,bNormalize){
 if (bNormalize === undefined){var bNormalize=false};
 
 return namedRange.getRange().getSheet().getName()+'!'+ getNamedRangeRangeA1Not(namedRange,bNormalize);
 } 
 
 function getNamedRangeRange(namedRange){
 return namedRange.getRange()
 } 
 
 function getNamedRangeSheet(namedRange){
 return namedRange.getRange().getSheet()
 } 
  
 function getNamedRangeSheetName(namedRange){
 return namedRange.getRange().getSheet().getName()
 }
function updateValues(dataRangeSearch){
  //dataRangeSearch = activeSheet.getRange(1,1,activeSheet.getLastRow());
  dataSearch = dataRangeSearch.getValues().reduce(function (a, b) {
    return a.concat(b);
  });;
}

