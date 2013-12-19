/* Other functions used by the Google Doc: */


/**
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var sheet = SpreadsheetApp.getActive();
  var entries = [
    {
      name : "Process Unsent Reports",
      functionName : "fileReports"
    },
    {
      name : "Generate New Report",
      functionName : "nostraDaemon"
    },
    {
      name : "Clear Report",
      functionName : "clearReport"
    },
    {
      name : "Mail Current Report",
      functionName : "mailPDF"
    }
  ];
  sheet.addMenu("Forecast Menu", entries);  // Adds a custom menu to the active spreadsheet
};

/**
 * Retrieves all the rows in the active spreadsheet that contain data and logs the
 * values for each row.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function readRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  for (var i = 0; i <= numRows - 1; i++) {
    var row = values[i];
    Logger.log(row);
  }
};


/**
 * Shorthand for finding empty/zero values:
 */
function isNothing(X)
{
  return (X == 0 || X == "" || X == null);
}


/**
 * DEZERO: Blank out cells with zeroes, and round other values
 */
function deZero(rCells,arValues)
{
  var x=0, y=0;
  
  for(x=0; x<arValues.length; x++)
  {
    for(y=0; y<arValues[x].length; y++)
    {
      arValues[x][y] = ( isNothing(arValues[x][y]) ) ? "" : Math.round(arValues[x][y]);
    }
  }
  
  rCells.setValues( arValues );
}

