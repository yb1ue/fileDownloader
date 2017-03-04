// Add menu
function onOpen() 
{
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var entries = [ {
        name : "Run", functionName : "main" 
    }];
    spreadsheet.addMenu("** Mnual Run **", entries);
};

// Main
function main() {
  var targetURLs = getTargetURLs();
  downloadContents(targetURLs);
  clearSheet();
}

// Get file URLs from spreadsheet and store into array.
function getTargetURLs(){
  var targetURLs =[];
  var myRange = getLatestRage();
  var myValues = myRange.getValues();
  for(var Array_i = 0, Array_l = myValues.length; Array_i < Array_l; Array_i++){
    targetURLs.push(myValues[Array_i][0]);
  }
  Logger.log(targetURLs);
  return targetURLs;
}

// Download files listed in the array to Drive root directory.
function downloadContents(targetURLs){
  var rootFolder = DriveApp.getRootFolder();
  for(var Array_i = 0, Array_l = targetURLs.length; Array_i < Array_l; Array_i++){
      var response = UrlFetchApp.fetch(targetURLs[Array_i]);
      var fileBlob = response.getBlob();
      rootFolder.createFile(fileBlob);
  }
}

// Clear the sheet.
function clearSheet(){
  var myRange = getLatestRage();
  myRange.clear();
}

// Get latest spreadsheet range that contain a value (exclude header).
function getLatestRage() {
  var mySS = SpreadsheetApp.getActiveSpreadsheet();
  var mySheet = mySS.getSheetByName("List");
  var startRow = 2;
  var startCol = 1;
  var endRow = mySheet.getLastRow() - startRow + 1;
  var endCol = mySheet.getLastColumn() - startCol + 1;
  var myRange = mySheet.getRange(startRow, startCol, endRow, endCol)
  return myRange;
}
