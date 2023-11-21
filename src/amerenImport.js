function importAmerenData(messages,source) 
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
    try {
      var msg = messages[i].getPlainBody().toString().split("*");
      var duedate = msg[10].split(" ")[1].slice(0,10).trim();
      var amount = "$" + msg[8].trim();
      data.push([duedate,amount]);
    } catch (e) {
      let errormessage = "Unable to process email. Data processing resulted in error: " + e;
      console.error(errormessage);
      return new Error(e);
    }
  }
  
  if (data.length > 0) {
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1,1,data.length, data[0].length).setValues(data); 
  }
}