function run() {
  var errorlog = [];

  var sources = ["Ameren","Lowes","Spire","HomeDepot"]; // Step 1: add Source here
  var importLabels = {};
  var importedLabel = GmailApp.getUserLabelByName("Imports/Imported");
  
  if (importedLabel === null){
    let errormessage = "Imports/Imported label not found.";
    console.error = (errormessage);
    errorlog.push(errormessage);
    return;
  }

  for (let i = 0; i < sources.length; i++) {
    importLabels[sources[i]] = GmailApp.getUserLabelByName("Imports/" + sources[i] + "Imports");
  }
  
  for (let i = 0; i < sources.length; i++)
  {
    var result = myGetMessages(sources[i], importLabels[sources[i]], importedLabel);

    if (result instanceof Error) {
      errorlog.push(result.message);
    }
  }
  
  if (errorlog.length > 0) {
    var job = "House Data Import ";
    var subject = "Error results from " + job;
    sendErrorEmail(errorlog, subject, job);
  }
}

function getDataHandler(source) {
  var handlers = {
    "Ameren"    : importAmerenData,
    "Lowes"     : importLowesData,
    "Spire"     : importSpireData,
    "HomeDepot" : importHomeDepotData,
    // Step 2: define data handler
    // Step 3: create new class and method for data handler
  };
  return handlers[source];
}

function myGetMessages(source, label, importedLabel) {
  if (label === null) {
    let errormessage = "Unable to process emails from source: " + source + ". Label not found."
    console.error(errormessage);
    return new Error(errormessage);
  }
  var threads = label.getThreads();
  var handler = getDataHandler(source);

  if (!handler) {
    let errormessage = "Unable to process data from source: " + source + ". Data handler not defined.";
    console.error(errormessage);
    return new Error(errormessage);
  }
  
  threads.forEach(function(thread) {
    var messages = thread.getMessages();
    var handlerResult = handler(messages, source);
    
    if (handlerResult instanceof Error) {
      return new Error(handlerResult);
    } else {
      thread.removeLabel(label);
      thread.addLabel(importedLabel);
    }
  });
}