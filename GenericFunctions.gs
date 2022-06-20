var GOOGLEDRIVEFOLDERID = 'The ID of the Google Drive Folder you wish for the CSV copy to be saved.';
var DAYSTOKEEPBACKUPS = 20;

function createCSV() {
  // You can define mutliple calls to the saveSheetAsCSV function within this createCSV function as long as the sheets exist within this Google Sheets.
  saveSheetAsCSV("nameofcsvyouaresaving.csv","nameofsheetyouaregettingcsvdatafrom");
  //saveSheetAsCSV("nameofcsvyouaresaving.csv","nameofsheetyouaregettingcsvdatafrom");
  //saveSheetAsCSV("nameofcsvyouaresaving.csv","nameofsheetyouaregettingcsvdatafrom");
  //saveSheetAsCSV("nameofcsvyouaresaving.csv","nameofsheetyouaregettingcsvdatafrom");
}

function saveSheetAsCSV(fileName,sheetName) {
  var folder = DriveApp.getFolderById(GOOGLEDRIVEFOLDERID);
  var sheetData = parseSheetToCSV(sheetName);
  // Delete previous version of file(s) with same name
  var files = folder.getFilesByName(fileName);
  while(files.hasNext()) {
    var file = files.next();
    file.setTrashed(true);
  }
  folder.createFile(fileName,sheetData);
}

function parseSheetToCSV(sheetName) {
  var csvData = "";
  var sheetRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getDataRange();
  var sheetValues = sheetRange.getValues();
  
  for(var row = 0; row < sheetValues.length; row++) {
    var currentRow = "";
    for(var column = 0; column < sheetValues[0].length; column++) {
      var currentValue = "";

      // If value has a comma and doesn't already have quotations, make sure to surround with quotations
      if((sheetValues[row][column].toString().indexOf(',') != -1) && (sheetValues[row][column].toString().indexOf('"') != 0)) {
        currentValue = "\"" + sheetValues[row][column] + "\"";
      }else{
        currentValue = sheetValues[row][column];
      }

      // If this is the last column then append a new line tag, otherwise just append a comma
      if(column < sheetValues[0].length-1) {
        currentRow += currentValue + ",";
      }else{
        currentRow += currentValue + "\r\n";
      }
    }
    csvData += currentRow;
  }
  
  return csvData;
}
