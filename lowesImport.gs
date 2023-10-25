function importLowesData(messages,source)
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
    var msg = messages[i].getPlainBody().toString();
    var layers = msg.split("*")[6].split("\n");

    var transnum = layers[5].split(" ")[3].trim();
    var orderdate = layers[6].split(" ")[3].trim() + " " + layers[6].split(" ")[4].trim();
    
    var beef = msg.split("Price")[1].split("Total")[0];
    var beeflayers = beef.split("*");

    var subtotal = msg.split("*")[14].split(" ")[1].trim();
    var total = msg.split("*")[20].split("\n")[2].split(" ")[1].trim();
    var ratio = total/subtotal;

    var items = beeflayers[1].trim().split("\n");

    var description = [];
    var totalPrice = [];
    var qty = [];
    var pricePer = [];

    for (var j = 0; j < items.length; j += 9)
    {
      description.push(items[j].trim());
      totalPrice.push(Number(items[j+3].split(" ")[1].trim()));
      qty.push(Number(items[j+7].split(" ")[0]));
      pricePer.push(Number(items[j+7].split(" ")[2].trim()))
      //newTotal.push(totalPrice[j]*ratio);
    }

    for (var k = 0; k < description.length; k++)
    {
      var newTotal = totalPrice[k]*ratio;
      data.push([transnum,orderdate,description[k],totalPrice[k],qty[k],pricePer[k],newTotal]);
    
    }
  }

  if (data.length > 0) {
  var lastRow = sheet.getLastRow();
  sheet.getRange(sheet.getLastRow()+1,1,data.length, data[0].length).setValues(data);
  }
}
