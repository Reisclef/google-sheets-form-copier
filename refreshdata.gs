/**
 * If there are more response lines than info lines. Add the missing formulae based on the 2nd row in Information (first row of data)
 */
function refreshData(){
  var lastResponse = getLastFormEntry();
  var lastCalc = getLastInformationRow();
  if (lastResponse > lastCalc){
    var prefs = getPreferences();
    var copiedRow = getRowFormulae(prefs['rowToCopy']);
    for (var r = lastCalc; r < lastResponse; r++){
      var infoSheet = activateSheet(prefs['targetSheet']);
      var rowToChange = r + 1;
      copiedRow.copyTo(infoSheet.getRange(rowToChange + ":" + rowToChange));
    }
    Browser.msgBox(lastResponse - lastCalc + " rows added.",Browser.Buttons.OK);
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
  var prefs = getPreferences();
  var responseSheet = activateSheet(prefs['sourceSheet']);
  return responseSheet.getLastRow();
}

/**
 * Get the total number of calculation rows.
 */
function getLastInformationRow(){
  var prefs = getPreferences();
  var infoSheet = activateSheet(prefs['targetSheet']);
  return infoSheet.getLastRow();
}

/**
 * Copy the first row of formulae the Information sheet (to avoid possible manual amendments)
 */
function getRowFormulae(row) {
  var prefs = getPreferences();
  var infoSheet = activateSheet(prefs['targetSheet']);
  var range = infoSheet.getRange(row + ":" + row );
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
 * Function to add menu options for refreshing, and setting options.
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Refresh calculations', functionName: 'refreshData'},
    {name: 'Settings', functionName: 'displayOptions'}
  ];
  spreadsheet.addMenu('Custom Options', menuItems);
}

 /* 
  * Display the sidebar
  */
function displayOptions(){
  var html = HtmlService.createTemplateFromFile('sidebar').evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Form calculation Settings')
      .setWidth(300);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(html);
}

 /* 
  * Return user's preferences for source and target spreadsheet
  */
function getPreferences() {
  var userProperties = PropertiesService.getUserProperties();
  var prefs = {
    sourceSheet: userProperties.getProperty('sourceSheet'),
    targetSheet: userProperties.getProperty('targetSheet'),
    rowToCopy: userProperties.getProperty('rowToCopy'),
  };
  return prefs;
}

 /* 
  * Set the user's preferences based on the function arguments
  */
function setPreferences(source,target,row) {
  var userProperties = PropertiesService.getUserProperties()
  userProperties.setProperty('sourceSheet', source);
  userProperties.setProperty('targetSheet', target);
  userProperties.setProperty('rowToCopy', row);
}
