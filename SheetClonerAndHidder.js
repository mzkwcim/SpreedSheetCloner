function kopiujUkryjZakladkiIUsunZawartosc() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = sheet.getSheets();

  // Kopia najbardziej wysuniętej w lewo zakładki
  var sheetToCopy = sheets[0];
  var newSheet = sheetToCopy.copyTo(sheet);

  // Usuń zawartość komórek z zakresu B2:Z33 na nowej zakładce
  var rangeToRemove = newSheet.getRange('B2:Z33');
  rangeToRemove.clearContent();

  // Umieść nową zakładkę najbardziej z lewej strony
  newSheet.setName(GetNewSheetName());
  sheet.setActiveSheet(newSheet);
  sheet.moveActiveSheet(1);

  // Ukryj wszystkie zakładki, zaczynając od drugiej
  for (var i = 0; i < sheets.length; i++) {
    sheets[i].hideSheet();
  }
}

function GetNewSheetName(){
  var now = new Date();
  var day = now.getDate();
  var month = now.getMonth() + 1;
  var dayOfWeek = now.getDay();
  var year = now.getFullYear();
  if (month.toString().length < 2){
    month = "0" + month
  } else if (day.toString().length < 2){
    day = "0" + day
  }
  if (dayOfWeek === 2){
    return "ANC " + day + "." + month + "." + year + "r."
  }
  if (dayOfWeek === 5){
    return "AEC2 " + day + "." + month + "." + year + "r."
  }
}

