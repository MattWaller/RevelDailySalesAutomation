function importCSVFromGoogleDrive() {
  // import from Operations CSV
    var file = DriveApp.getFilesByName("OperationsReport.csv").next();
    var csvData = Utilities.parseCsv(file.getBlob().getDataAsString());
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Operation Report RAW");
    sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);

  // import from Ingredient CSV
    var file2 = DriveApp.getFilesByName("IngredientsReport.csv").next();
    var csvData2 = Utilities.parseCsv(file2.getBlob().getDataAsString());
    var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ingredient Summary RAW");
    sheet2.getRange(1, 1, csvData2.length, csvData2[0].length).setValues(csvData2);

}
