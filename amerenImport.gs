function importAmerenData(_messages,_source) 
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(_source+"Import");
  
  if (sheet === null) {
      let errormessage = "Unable to process data from source: "+_source+". Sheet not found."
      console.error(errormessage);
      return new Error(errormessage);
  }
  for(var i = 0; i < _messages.length; i++)
  {
    var msg = _messages[i].getPlainBody().toString().split("*");
    var duedate = msg[10].split(" ")[1].slice(0,10);
    var data = [[duedate.trim(),"$"+msg[8].trim()]];

    sheet.getRange(sheet.getLastRow()+1,1,data.length, data[0].length).setValues(data);
  }
}
