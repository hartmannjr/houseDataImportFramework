function importHomeDepotData(messages,source)
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
  for (let i = 0; i < messages.length; i++)
  {
    var msg = messages[i].getPlainBody().toString().split("[image: The Home Depot]")[1];
    var layers = msg.split("\n");

    var transnum = layers[14].slice(0,20).trim();
    var orderdate = layers[14].slice(21).trim();

    let regex =/\s+\d{2}\/\d{2}\/\d{2}\s+/;
    var beef = msg.split(regex)[1].split("SUBTOTAL");

    var footer = beef[1].split("\n");
    var subtotal = footer[0].trim();
    var total = footer[4].split("$")[1].trim();
    var ratio = total/subtotal;

    var items = beef[0].split(/\d{4}\S\d{3}\S\d{3}/);

    var description = [];
    var totalPrice = [];
    var qty = [];
    var pricePer = [];
    
    for (let j = 1; j < items.length; j++){
      var item = items[j].split("\r");
      if (item.length > 6){
        var _description = item[2].trim();
        var _pretotalPrice = item[4].split(" ");
        var _totalPrice = _pretotalPrice[_pretotalPrice.length -1];
        var _qty = item[4].split("@")[0].trim();
        var _pricePer = item[4].split("@")[1].split(" ")[0];
      } else {
        var _description = item[0].split("<")[0].trim();
        var _totalPrice = item[0].split(">")[1].trim();
        var _qty = 1;
        var _pricePer = item[0].split(">")[1].trim();
      }
      description.push(_description);
      totalPrice.push(Number(_totalPrice));
      qty.push(Number(_qty));
      pricePer.push(Number(_pricePer));
    }

    for (let j = 0; j < description.length; j++){
      var newTotal = totalPrice[j]*ratio;
      data.push([transnum,orderdate,description[j],totalPrice[j],qty[j],pricePer[j],newTotal])
    }
  }

  if (data.length > 0) {
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1,1,data.length, data[0].length).setValues(data);
  }
}
