function Config(activeSheetName) {
  
  
  
  spreadsheetId = '14vVgGon3Xs2dtx9wi9iqmPn0VUKYaevwCddcVFd2VKA'
  spreadsheet = SpreadsheetApp.openById(this.spreadsheetId) 
 
  activeSheetName=activeSheetName||GetSheetByGid(spreadsheet,0).getName()
  activeSheet = this.spreadsheet.getSheetByName(activeSheetName)
  gid=''+activeSheet.getSheetId()
  
 
  return {
    spreadsheetId: spreadsheetId,
    spreadsheet: spreadsheet,
    activeSheet: activeSheet,
    gid:gid
  }
}
function testSession(){
Logger.log(MySession.getActiveUser())
}
function MySession(){
User=Session.getActiveUser();
UserKey=Session.getTemporaryActiveUserKey();
EffectiveUser=Session.getEffectiveUser();
return {
User:User,
UserKey:UserKey,
EffectiveUser:EffectiveUser
}



}