function Coloring() {
  var arkusz = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = arkusz.getSheets();
  var firstSheet = sheets[0];
  for (var i = 2; i <= 27; i++) {
    var max = null; // Inicjujemy max dla każdego wiersza
    var min = 400; // Inicjujemy min dla każdego wiersza

    var range = firstSheet.getRange("B" + i + ":E" + i);
    var values = range.getValues()[0];

    for (var j = 0; j < values.length; j++) {
      if (values[j] !== ""){
        var normalized = values[j].toString().replace(/[^0-9,.]/g, '').replace(",",".").trim();
        console.log("Po replace:", normalized);
        if (!isNaN(normalized)) { // Sprawdzamy czy znormalizowana wartość jest liczbą
          if (normalized > max) {
            max = normalized;
          }
          if (normalized < min) {
            min = normalized;
          }
        }
      }
      
    }

    var diff = max - min;
    firstSheet.getRange("J" + i).setValue(diff);

    max = null;
    min = 400;

    var range2 = firstSheet.getRange("F" + i + ":I" + i);
    var values2 = range2.getValues()[0];

    for (var j = 0; j < values2.length; j++) {
      if (values2[j] !== ""){
        var normalized2 = values2[j].toString().replace(/[^0-9,.]/g, '').replace(",",".").trim();
        console.log("Po replace:", normalized);
        if (!isNaN(normalized2)) { // Sprawdzamy czy znormalizowana wartość jest liczbą
          if (normalized2 > max) {
            max = normalized2;
          }
          if (normalized2 < min) {
            min = normalized2;
          }
        }
      }
    }

    var diff2 = max - min;
    firstSheet.getRange("K" + i).setValue(diff2);
  }
  for (var l = 0; l < 2; l++){
    colorizeCells((l===0) ? "J" : "K");
  }
}

function colorizeCells(collumn) {
  var arkusz = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = arkusz.getSheets();
  var firstSheet = sheets[0];
  
  for (var i = 2; i <= 27; i++) {
    var range = firstSheet.getRange(collumn + i);
    var value = range.getValue();
    
    if (value > 0 && value < 0.31) {
      range.setBackground("green");
    } else if (value >= 0.31 && value < 0.5) {
      range.setBackground("yellow");
    } else if (value >= 0.5) {
      range.setBackground("red");
    }
  }
}
