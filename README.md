# GoogleSheets-To-CSV
Google Apps Script to save a sheet as a csv and download it automatically using Google Drive for Windows.

Steps to setup:
1. Go to the Google Apps Script for the Google Sheet you wish to save as a CSV file and create a new blank script file. Delete all auto-generated content within the file.
2. Copy and paste the contents of GenericFunctions.gs into the new script file you just created and save the file.
3. Change the two global variables at the top of the script (Folder ID and Days to Keep Backups).
4. Specify the CSV name and the Google Sheet name in the CreateCSV function. Save your changes then run the createCSV function as it will ask for folder permission the first time.
5. 
