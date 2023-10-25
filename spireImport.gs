function importSpireData(messages,source) 
{
  var sheetName = source + "Import";
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  // constants
  
  if (sheet === null) {
      let errormessage = "Unable to process data from source: " + source + ". Sheet not found."
      console.error(errormessage);
      return new Error(errormessage);
  }

  var data = [];
  for (let i = 0; i < messages.length; i++) {
    var msg = _messages[j].getPlainBody().toString().split("Hello")[1].split(" ");
    data.push([msg[16].slice(0,10),msg[12]]);
  }
  
  if (data.length > 0) {
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1,1,data.length, data[0].length).setValues(data); 
  }
}
