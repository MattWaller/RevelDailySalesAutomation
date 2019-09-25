function ImportWeeklyDrop() {
  var app = SpreadsheetApp;
  var as = app.getActiveSpreadsheet().getSheetByName("CashDropBackup");
  var clear = as.getRange("A:Z").clearContent();
  var import = as.getRange(1,1).setValue('=importrange("https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit","WeeklyCash!A:Z")');
  var data = as.getRange("A:Z").getValues();
  
  var copy = as.getRange("A:Z").setValues(data);
  
  
}
