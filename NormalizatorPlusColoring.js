function processRange(range) {
  var max = null;
  var min = 400;

  var values = range.getValues()[0];

  for (var j = 0; j < values.length; j++) {
    if (values[j] !== "") {
      var normalized = values[j].toString().replace(/[^0-9,.:]/g, '').replace(",", ".").trim();
      if (normalized.charAt(normalized.length - 3) === ":") {
        normalized = normalized.replace(":", ".");
      }
      if (normalized.charAt(normalized.length - 1) === ":") {
        normalized = normalized.replace(":", "");
      }
      normalized.replace(/\:$/, "");
      if (normalized.charAt(1) === ":") {
        normalized = parseFloat(normalized.substring(2)) + (60 * parseInt(normalized.charAt(0)));
      }
      console.log("Po replace:", normalized);
      if (!isNaN(normalized) && normalized !== "") {
        if (normalized > max) {
          max = normalized;
        }
        if (normalized < min) {
          min = normalized;
        }
      }
    }
  }
  return {
    max: max,
    min: min
  };
}

function Coloring() {
  var arkusz = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = arkusz.getSheets();
  var firstSheet = sheets[0];

  for (var i = 2; i <= 27; i++) {
    var range1 = firstSheet.getRange("B" + i + ":E" + i);
    var result1 = processRange(range1);

    var range2 = firstSheet.getRange("F" + i + ":I" + i);
    var result2 = processRange(range2);

    var diff1 = result1.max - result1.min;
    var diff2 = result2.max - result2.min;

    if (diff1 === -400) {
      diff1 = "";
    }
    if (diff2 === -400) {
      diff2 = "";
    }

    firstSheet.getRange("J" + i).setValue(diff1);
    firstSheet.getRange("K" + i).setValue(diff2);
  }

  for (var l = 0; l < 2; l++) {
    colorizeCells((l === 0) ? "J" : "K");
  }
}

function colorizeCells(column) {
  var arkusz = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = arkusz.getSheets();
  var firstSheet = sheets[0];

  for (var i = 2; i <= 27; i++) {
    var range = firstSheet.getRange(column + i);
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
