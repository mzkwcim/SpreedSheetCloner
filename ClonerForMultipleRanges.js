function duplicateSheetAndRename() {

  var aktywnyArkusz = SpreadsheetApp.getActiveSpreadsheet();

  var arkusze = aktywnyArkusz.getSheets();

  var arkuszDoKopiowania = arkusze[arkusze.length - 1];

  if (!arkuszDoKopiowania) {
    Logger.log("Brak arkusza do skopiowania.");
    return;
  }

  // Skopiuj arkusz do nowego arkusza
  var nowyArkusz = arkuszDoKopiowania.copyTo(aktywnyArkusz);


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
  var zakresy = [
  'B2:J2',
  'B5:J5',
  'B8:J8',
  'B11:J11',
  'B14:J14',
  'B17:J17',
  'B20:J20',
  'B23:J23'
  ];
  if (nextMonth > 12) {
    nextMonth = "01";
    nextYear += 1;
    var nowaNazwa = monday + "." + month +"." + year + "-" + sunday + "." + nextMonth + "." + nextYear + "r.";
    Rename(nowyArkusz, nowaNazwa);
    clearMultipleRanges(nowyArkusz, zakresy);
  }else{
    var nowaNazwa = monday + "." + month + "-" + sunday + "." + nextMonth + "." + year + "r.";
    Rename(nowyArkusz, nowaNazwa);
    clearMultipleRanges(nowyArkusz, zakresy);
  }
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

function clearMultipleRanges(sheet, ranges) {
  ranges.forEach(function(range) {
    var clearRange = sheet.getRange(range);
    clearRange.clearContent();
  });
}
