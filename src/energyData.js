function editEnergyData(){
  var errorLog = [];
  var types = ["Electricity","Gas"];

  types.forEach(function(type){
    var dataRange;
    if (type === "Electricity") {
      
    }
    var result = polyfit(type,dataRange);

    if (result instanceof Error) {
      errorLog.push(result.message);
    }
  });
  if (errorlog.length > 0) {
    var job = "Energy Data Edit ";
    var subject = "Error results from " + job;
    sendErrorEmail(errorLog, subject, job);
  }
}