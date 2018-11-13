function PriceListRepository() { 
var activeSheet = Config().activeSheet
return {
get: function(id) {
return id
? activeSheet.getRange(id).getValues()
: activeSheet.getRange('a1:b')
.getValues()
.filter(function(row) { return !!row[0] })
},
post: function(data) {
var dataArr = data.split(',')
activeSheet
.appendRow(dataArr)
return 'OK'
}
}
}
