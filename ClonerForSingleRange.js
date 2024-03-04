function duplicateSheetAndRename() {
  // Pobierz aktywny arkusz
  var aktywnyArkusz = SpreadsheetApp.getActiveSpreadsheet();

  // Pobierz wszystkie arkusze w arkuszu głównym
  var arkusze = aktywnyArkusz.getSheets();

  // Znajdź arkusz do skopiowania (najbardziej wysunięty na prawo)
  var arkuszDoKopiowania = arkusze[arkusze.length - 1];

  // Sprawdź, czy arkusz do skopiowania istnieje
  if (!arkuszDoKopiowania) {
    Logger.log("Brak arkusza do skopiowania.");
    return;
  }

  // Skopiuj arkusz do nowego arkusza
  var nowyArkusz = arkuszDoKopiowania.copyTo(aktywnyArkusz);

  // Ustaw nazwę nowego arkusza na podstawie aktualnej godziny
  var now = new Date();
  var month = now.getMonth() + 1;
  var nextMonth = month;
  var monday = now.getDate();
  var sunday = monday + 6;
  var year = now.getFullYear();
  var nextYear = year;

  if (month === 2 && sunday > 28) {
    if ((year % 4 === 0 && year % 100 !== 0) || year % 400 === 0) {
      // Rok przestępny
      sunday -= 29;
    } else {
      // Rok zwykły
      sunday -= 28;
    }
    nextMonth += 1;
  } else if (month === 4 || month === 6 || month === 9 || month === 11) {
    if (sunday > 30) {
      sunday -= 30;
      nextMonth += 1;
    }
  } else if (sunday > 31) {
    sunday -= 31;
    nextMonth += 1;
  }

  // Poprawki dla formatowania jednocyfrowych dni, miesięcy i przyszłego dnia
  if (sunday.toString().length < 2) {
    sunday = "0" + sunday;
  }
  if (monday.toString().length < 2) {
    monday = "0" + monday;
  }
  if (month.toString().length < 2) {
    month = "0" + month;
  }
  if (nextMonth.toString().length < 2) {
    nextMonth = "0" + nextMonth;
  }

  if (nextMonth > 12) {
    nextMonth = "01";
    nextYear += 1;
    var nowaNazwa = monday + "." + month +"." + year + "-" + sunday + "." + nextMonth + "." + nextYear + "r.";
    Rename(nowyArkusz, nowaNazwa);
    clearValuesInRange(nowyArkusz, 'B2:I30');
  }else{
    var nowaNazwa = monday + "." + month + "-" + sunday + "." + nextMonth + "." + year + "r.";
    Rename(nowyArkusz, nowaNazwa);
    clearValuesInRange(nowyArkusz, 'B2:I30');
  }
  ukryjDrugiWierszWeWszystkichZakladkach();
  
}

function ukryjDrugiWierszWeWszystkichZakladkach() {
  var skoroszyt = SpreadsheetApp.getActiveSpreadsheet();
  var iloscZakladek = skoroszyt.getSheets().length;
  skoroszyt.getSheets()[iloscZakladek-1].hideRows(13);
  skoroszyt.getSheets()[iloscZakladek-1].hideRows(27);
  skoroszyt.getSheets()[iloscZakladek-1].hideRows(29);
}

function Rename(tab, name) {
  var sheetNames = tab.getParent().getSheets().map(sheet => sheet.getName());

  function checkName() {
    if (sheetNames.includes(name)) {
      name = name + '*';
      return checkName();
    }
    tab.setName(name);
  }

  checkName(name);
}

function clearValuesInRange(sheet, range) {
  var clearRange = sheet.getRange(range);
  clearRange.clearContent();
}

