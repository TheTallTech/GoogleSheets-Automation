//Global variable of how many days to keep backups.
var DAYSTOKEEPBACKUPS = 20;

function backupMultipleSheets() {
  backupSheet('Name of sheet to backup.');
  //backupSheet('Name of sheet to backup.');
  //backupSheet('Name of sheet to backup.');
  //backupSheet('Name of sheet to backup.');
}

function backupSheet(sheetToBackup) {
  //Get todays date and search all sheets for backup sheets, any that are older than number of days global variable then delete
  //Create backup using todays date
  deleteOldBackups(DAYSTOKEEPBACKUPS);
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
    if(sheet.includes(sheetName)) {
      // Check backup sheet date and if older than number of days global variable then delete
      let backupDateString = sheet.substring(7,17);
      let backupDate = getDateFromString(backupDateString).getTime();
      let currentDate = new Date().getTime();
      var calculation = currentDate - backupDate;
      if((currentDate - backupDate) >= numDaysMilliseconds) {
        activeSS.deleteSheet(sheets[i]);
      }
    }
  }
}

function getDateFromString(stringDate) {
  let monthInt = parseInt(stringDate.substring(0,2));
  let dayInt = parseInt(stringDate.substring(3,5));
  let yearInt = parseInt(stringDate.substring(6));
  var newDateObject = new Date(yearInt,monthInt-1,dayInt);
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
  var currentDate = new Date();
  var newSheetName = sheetNameHalf + ' ' + formatDate(currentDate);
  var dynamicSheetName = newSheetName;
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
