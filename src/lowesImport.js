function importLowesData(messages,source)
{
  //var source = "Lowes"; // change "template" to your source, has to match exactly with the label. remove line after debugging
  var sheetName = source + "Import";
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  // constants
  
  if (sheet === null) {
      let errormessage = "Unable to process data from source: " + source + ". Sheet not found."
      console.error(errormessage);
      return new Error(errormessage);
  }

  /*
  // start removal after debugging
  var label = GmailApp.getUserLabelByName("Imports/" + source + "Imports");
  var threads = label.getThreads();

  for (let i = 0; i < threads.length; i++)
  {
    var messages = threads[i].getMessages();

  // remove line 15 to here after debugging. master.gs will handle this.
  */

  var data = [];
  for (let i = 0; i < messages.length; i++)
  {
    try {
      var msg = messages[i].getPlainBody().toString();

      var transnum = msg.split("Transaction #")[1].split(" ")[2];
      var orderdate = msg.split("Order Date")[1].split(" ")[2] + " " + msg.split("Order Date")[1].split(" ")[3];
      
      var beef = msg.split("Price")[1].split("Invoice")[0];

      var subtotal = msg.split("Invoice")[1].split("$")[1].trim();
      var total = msg.split("Total")[2].split(" ")[2];
      var ratio = total/subtotal;

      //console.log(beef);
      var descs = beef.split("$");

      var description = [];
      var totalPrice = [];
      var qty = [];
      var pricePer = [];

      for (let j = 0; j < items.length; j += 9)
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
    } catch (e) {
      let errormessage = "Unable to process email. Data processing resulted in error: " + e;
      console.error(errormessage);
      return new Error(e);
    }
  }

  if (data.length > 0) {
  var lastRow = sheet.getLastRow();
  sheet.getRange(sheet.getLastRow()+1,1,data.length, data[0].length).setValues(data);
  }
  //} // <- remove after debugging
}