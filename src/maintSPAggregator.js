function aggregateData() {
  var sources = ["Lowes","HomeDepot"]; // add amazon, need importer, sheet

  var maintdata = [];
  var spdata = [];

  sources.forEach(function(source){
    var sourcesheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(source+"Import");
    var importc = sourcesheet.getRange("K:K").getValues();
    var lastRow;
    for (let i = importc.length - 1; i >= 0; i--) {
      var row = importc[i];
      if (row[0] === ''){
        lastRow = i; //this is the first open cell
      } // end of if
    } // end of for

    var sheetdata = sourcesheet.getRange(lastRow + 1,1,sourcesheet.getLastRow() - lastRow,sourcesheet.getLastColumn()).getValues();

    var categorizedItems = {};
    sheetdata.forEach(function(item, i) {

      var categoryKey = item[7];
      if (categoryKey === "") {
        let errormessage = "Unable to process item " + item[2] + ". Category Key is not defined.";
        console.log(errormessage);
        return;
      }
      var areaKey = item[8];
      if (areaKey === "") {
        let errormessage = "Unable to process item " + item[2] + ". Area Key is not defined.";
        console.log(errormessage);
        return;
      }
      var specifyKey = item[9];
      if (specifyKey === "") {
        let errormessage = "Unable to process item " + item[2] + ". Specify Key is not defined.";
        console.log(errormessage);
        return;
      }

      var comboKey = categoryKey + '_' + areaKey + '_' + specifyKey;

      if (!categorizedItems[comboKey]) {
        categorizedItems[comboKey] = [];
      }
      categorizedItems[comboKey].push(item);
      sourcesheet.getRange(lastRow + 1 + i,11).setValue(["True"]);
    });

    var categorizedTransNum = {};
    for (const itemKey in categorizedItems) {
      categorizedItems[itemKey].forEach(function(item) {
        var perTransNum = item[0] + '_' + itemKey;

        if (!categorizedTransNum[perTransNum]) {
          categorizedTransNum[perTransNum] = [];
        }
        categorizedTransNum[perTransNum].push(item);
      });
    }
    
    for (const transKey in categorizedTransNum) {
      var date;
      var category;
      var area;
      var specify;
      var totalCost = 0;
      var transNum;
      var specifications = "";

      categorizedTransNum[transKey].forEach(function(item) {
        date = item[1];
        category = item[7];
        area = item[8];
        specify = item[9];
        totalCost += item[6];
        transNum = item[0];
        if (specifications.length > 0) {specifications += "\n"}
        specifications += item[2];
      });

      
      if (category === "Maintenance") {
        maintdata.push([date, area, specify, totalCost, source + " " + transNum, "", specifications]);
      } else {
        spdata.push([date, area, specify, totalCost, source + " " + transNum, "", specifications]);
      }
    } // end of for of categorizedTransNums
  }); // end of forEach of source
  
  if (maintdata.length > 0) {
    var destinationsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Maintenance");
    var range = destinationsheet.getRange("A:G").getValues()
    let thisLastRow;
    for (let i = range.length - 1; i >= 0; i--) {
      var row = range[i];
      if (row[0] === ''){
        thisLastRow = i + 1; //this is the first open cell
      } // end of if
    }
    destinationsheet.getRange(thisLastRow,1,maintdata.length,maintdata[0].length).setValues(maintdata);
  }
  if (spdata.length > 0) {
    var destinationsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Special Purchases");
    var range = destinationsheet.getRange("A:G").getValues()
    let thatLastRow;
    for (let i = range.length - 1; i >= 0; i--) {
      var row = range[i];
      if (row[0] === ''){
        thatLastRow = i + 1; //this is the first open cell
      } // end of if
    }
    destinationsheet.getRange(thatLastRow,1,spdata.length,spdata[0].length).setValues(spdata);
  }
}
