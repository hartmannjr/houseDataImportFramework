function importSpireData(_messages,_source) 
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(_source+"Import");
  switch (sheet) 
  {
    case null:
      let errormessage = "Unable to process data from source: "+_source+". Sheet not found."
      console.error(errormessage);
      return new Error(errormessage);
    default:
      break;
  }

  for(var j = 0; j < _messages.length; j++)
  {
    var msg = _messages[j].getPlainBody().toString().split("Hello")[1].split(" ");
    var data = [[msg[16].slice(0,10),msg[12]]];

    sheet.getRange(sheet.getLastRow()+1,1,data.length, data[0].length).setValues(data);
  }
}
