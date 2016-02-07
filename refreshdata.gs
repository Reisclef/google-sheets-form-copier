/**
 * If there are more response lines than info lines. Add the missing formulae based on the 2nd row in Information (first row of data)
 */
function refreshData(){
  var lastResponse = getLastFormEntry();
  var lastInfo = getLastInformationRow();
  if (lastResponse > lastInfo){
    var copiedRow = getRowFormulae();
    for (var r = lastInfo; r < lastResponse; r++){
      var infoSheet = activateSheet("Target");
      var rowToChange = r + 1;
      copiedRow.copyTo(infoSheet.getRange(rowToChange + ":" + rowToChange));
    }
    Browser.msgBox(lastResponse - lastInfo + " rows added.",Browser.Buttons.OK);
  }
  else{
    Browser.msgBox("Nothing to refresh.",Browser.Buttons.OK);
  }
}

function getSheetNames(){
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var sheetNames = new Array();
  for (var s = 0; s < sheets.length; s++){
    sheetNames[s] = sheets[s].getName();
  }
  return sheetNames;
}

/**
 * Get the total number of form responses.
 */
function getLastFormEntry(){
  var responseSheet = activateSheet("Source");
  return responseSheet.getLastRow();
}

/**
 * Get the total number of calculation rows.
 */
function getLastInformationRow(){
  var infoSheet = activateSheet("Target");
  return infoSheet.getLastRow();
}

/**
 * Copy the first row of formulae the Information sheet (to avoid possible manual amendments)
 */
function getRowFormulae(row) {
  var infoSheet = activateSheet("Target");
  var range = infoSheet.getRange("2:2");
  return range;
}

/**
 * Activate the sheet based on the argument string value
 */
function activateSheet(sheet){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(sheet); 
}

/**
 * Function to add option to refresh data from form responses
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Refresh calculations', functionName: 'refreshData'}
  ];
  spreadsheet.addMenu('Custom Options', menuItems);
}

