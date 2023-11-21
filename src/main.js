// for functions not assigned to a specific sheet, which may be called from various places
function sendErrorEmail(errorlog,subject,job) {
  var currentTime = new Date();
  var formattedTime = Utilities.formatDate(currentTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var recipient = Session.getEffectiveUser().getEmail();
  //var subject = "Error results from House Data Import Job";
  var body = "Job Time: " + formattedTime + 
  "\nThe follow error";
  if (errorlog > 1) {body += "s"}
  body += " occured while running the " + job + "Job:\n\n"; 
  for (let i = 0; i < errorlog.length; i++)
  {
    body += errorlog[i] + "\n";
  }

  MailApp.sendEmail(recipient, subject + "Job", body);
}

