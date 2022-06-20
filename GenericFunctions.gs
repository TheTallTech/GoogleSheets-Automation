var GOOGLEDRIVEFOLDERID = '1zriR1Ebo6d-gTbsRZPak5RMWBPd61YPj';

function createCSV() {
  saveSheetAsCSV("Student_Device_Assignments.csv","Master Device Assignment");
}

function saveSheetAsCSV(fileName,sheetName) {
  var folder = DriveApp.getFolderById(GOOGLEDRIVEFOLDERID);
  var sheetData = parseSheetToCSV(sheetName);
  // Delete previous version of file(s) with same name
  var files = folder.getFilesByName(fileName);
  while(files.hasNext()) {
    var file = files.next();
    //folder.removeFile(file);
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

function backupSheet(sheetToBackup) {
  //var sheetToBackup = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Testing Master');
  //Get todays date and search all sheets for backup sheets, any that are older than 7 days delete
  //Create backup using todays date
  deleteOldBackups(20);
  var currentBackup = createDatedSheet('Backup');
  // Copy sheet data to backup sheet
  let maxRows = sheetToBackup.getLastRow();
  let maxCol = sheetToBackup.getLastColumn();
  let originSheetData = sheetToBackup.getRange(1,1,maxRows,maxCol).getValues();
  currentBackup.getRange(1,1,maxRows,maxCol).setValues(originSheetData);
  currentBackup.hideSheet();
}

function deleteOldBackups(numDaysToKeep) {
  var numDaysMilliseconds = numDaysToKeep * 24 * 60 * 60 * 1000;
  const sheetName = 'Backup'
  var activeSS = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = activeSS.getSheets();
  for(let i = 0; i < sheets.length; i++) {
    let sheet = sheets[i].getName();
    //console.log('Sheet name is: ' + sheet);
    if(sheet.includes(sheetName)) {
      // Check backup sheet date and if older than 7 days delete
      let backupDateString = sheet.substring(7,17);
      let backupDate = getDateFromString(backupDateString).getTime();
      //console.log('Backup date string: ' + backupDateString);
      let currentDate = new Date().getTime();
      //console.log('Sheet to check is: ' + sheet);
      //console.log('Backup date milliseconds: ' + backupDate);
      //console.log('Current date milliseconds: ' + currentDate);
      //console.log('Number of days to keep milliseconds: ' + numDaysMilliseconds);
      var calculation = currentDate - backupDate;
      //console.log('Current - backup milliseconds: ' + calculation);
      if((currentDate - backupDate) >= numDaysMilliseconds) {
        //console.log('Sheet to delete is: ' + sheet);
        activeSS.deleteSheet(sheets[i]);
      }
    }
  }
}

function getDateFromString(stringDate) {
  let monthInt = parseInt(stringDate.substring(0,2));
  let dayInt = parseInt(stringDate.substring(3,5));
  let yearInt = parseInt(stringDate.substring(6));
  //console.log("Month " + monthInt);
  //console.log("Day " + dayInt);
  //console.log("Year " + yearInt);
  var newDateObject = new Date(yearInt,monthInt-1,dayInt);
  //console.log("Date object: " + newDateObject);
  return newDateObject;
}

function formatDate(dateToFormat) {
  var day = dateToFormat.getDate();
  var month = dateToFormat.getMonth()+1;
  var year = dateToFormat.getFullYear();

  // Create string following MM-DD-YYYY format
  if(month < 10) {month = '0' + month}
  if(day < 10) {day = '0' + day}
  var dateString = month + '-' + day + '-' + year;

  return dateString;
}

function createDatedSheet(sheetNameHalf) {
  //var sheetNameHalf = 'Backup';
  var currentDate = new Date();
  var newSheetName = sheetNameHalf + ' ' + formatDate(currentDate);
  var dynamicSheetName = newSheetName;
  //SpreadsheetApp.getUi().alert('Sheet to create will have the name: ' + newSheetName);
  var sheetExists = false;
  var dailyCounter = 0;

  do{
    dynamicSheetName = newSheetName + ' ' + dailyCounter;
    if(checkForExistingSheet(dynamicSheetName)) {
      dailyCounter++;
      sheetExists = true;
    }else{
      sheetExists = false;
    }
  }while(sheetExists);
  
  var createdSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(dynamicSheetName);

  return createdSheet;
}

function checkForExistingSheet(sheetName) {
  var sheetExists = false;
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for(let i = 0; i < sheets.length; i++) {
      let sheet = sheets[i].getName();
      if(sheet == sheetName) {
        sheetExists = true;
        break;
      }
    }
    return sheetExists;
}
