function reset() {
  var s=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Operation Report RAW");
  s.getRange("A:Z").clearContent();
  
  var s2=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ingredient Summary RAW");
  s2.getRange("A:Z").clearContent();
  
  
}
