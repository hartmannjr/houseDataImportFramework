function run() {
  var errorlog = [];
  
  var sources = ["Ameren","Lowes","Spire","HomeDepot"]; // Step 1: add Source here
  var dataImportHandlers = {
    "Ameren"    : importAmerenData,
    "Lowes"     : importLowesData,
    "Spire"     : importSpireData,
    "HomeDepot" : importHomeDepotData,
    // Step 2: define data handler
    // Step 3: create new class and method for data handler
  }
  for (i = 0; i < sources.length; i++)
  {
    var result = myGetMessages(sources[i],dataImportHandlers);

    if (result instanceof Error) {
      errorlog.push(result.message);
    }
  }
  
  if (errorlog.length > 0) {
    sendErrorEmail(errorlog);
  }
}

function myGetMessages(_source, _dataImportHandler) {
  var label = GmailApp.getUserLabelByName("Imports/"+_source+"Imports");
  if (label === null) {
    let errormessage = "Unable to process emails from source: " + _source + ". Label not found."
    console.error(errormessage);
    return new Error(errormessage);
  } else {
    var threads = label.getThreads();
  }
  
  for (var i = 0; i < threads.length; i++)
  {
    var messages = threads[i].getMessages();
    var dataHandler = _dataImportHandler[_source];

    if (dataHandler) {
      var handlerResult = dataHandler(messages, _source);

      if (handlerResult instanceof Error) {
        messages.push(false);
        return new Error(handlerResult);

      } else {
        messages.push(true);
      }

    } else {
      let errormessage = "Unable to process data from source: " + _source + ". Data handler not defined.";
      console.error(errormessage);
      messages.push(false);
      return new Error(errormessage);
    }
    
    if(messages[messages.length - 1] === true)
    {
      threads[i].removeLabel(label);
      
      var newLabel = GmailApp.getUserLabelByName("Imports/Imported");
      threads[i].addLabel(newLabel);
    }
  }
}


function sendErrorEmail(_errorlog) {
  var recipient = Session.getEffectiveUser().getEmail();
  var subject = "Error results from House Data Import Job"
  var body = "The follow error(s) occured while running the House Data Import Job:\n\n"; 
  for (var i = 0; i < _errorlog.length; i++)
  {
    body += _errorlog[i] + "\n";
  }

  MailApp.sendEmail(recipient, subject, body);
}
