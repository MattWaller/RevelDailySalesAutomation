function saveAsSpreadsheet(){ 
  var currentDate = new Date();
  var date = currentDate.getDate();
  var month = currentDate.getMonth(); //Be careful! January is 0 not 1
  var year = currentDate.getFullYear();
  var renameFile = "OperationsReport-" +(date - 1) + "-" +(month + 1) + "-" + year;  
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var destFolder = DriveApp.getFolderById("FOLDER_ID"); 
  DriveApp.getFileById(sheet.getId()).makeCopy(renameFile, destFolder); 
}
