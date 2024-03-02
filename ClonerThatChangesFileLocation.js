function createFolderAndCopyFile() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  var folderId = DriveApp.getFileById(activeSpreadsheet.getId()).getParents().next().getId();

  var now = new Date();
  var month = now.getMonth() + 1;
  var nextMonth = month;
  var day = now.getDate();
  var nextDay = day + 6
  var year = now.getFullYear()

  if (((year % 4 === 0 && year % 100 !== 0) || year % 400 === 0) && (month === 2 && nextDay > 29)) {
    nextMonth += 1
  } else if ((month === 2 && nextDay > 28)) {
    nextMonth += 1
  } else if ((month === 4 || month === 6 || month === 9 || month === 11) && nextDay > 30) {
    nextMonth += 1;
  } else if ((month === 1 || month === 3 || month === 5 || month === 7 || month === 8 || month === 10 || month === 12) && nextDay > 31) {
    nextMonth += 1;
  }
  
  if (nextMonth > 12) {
    nextMonth -= 12;
    year += 1;
  }

  if (month !== nextMonth) {
    var parentFolder = DriveApp.getFolderById(folderId).getParents().next();

    var newFolderName = FolderName();
    var newFolder = parentFolder.createFolder(newFolderName);

    Logger.log("Utworzono nowy folder: " + newFolderName);

    var newFolderId = newFolder.getId();

    var activeSpreadsheetBlob = DriveApp.getFileById(activeSpreadsheet.getId()).getBlob();

    var copiedFileBlob = newFolder.createFile(activeSpreadsheetBlob).getBlob();
    var copiedSpreadsheet = SpreadsheetApp.open(copiedFileBlob);

    RangeRemover(copiedSpreadsheet);
    Logger.log("Skopiowano plik do nowego folderu: " + copiedSpreadsheet.getName());
  } else {
    var copiedFile = DriveApp.getFileById(activeSpreadsheet.getId()).makeCopy(FileName(), parentFolder);
    var copiedSpreadsheet = SpreadsheetApp.open(copiedFile);

    RangeRemover(copiedSpreadsheet);
    Logger.log("Skopiowano plik do obecnego folderu: " + copiedSpreadsheet.getName());
  }
}
function getDateInfo() {
  var now = new Date();
  var month = now.getMonth() + 1;
  var nextMonth = month;
  var monday = now.getDate();
  var sunday = monday + 6;
  var year = now.getFullYear();
  var nextYear = year;

  if (month === 2 && sunday > 28) {
    if (((year % 4 === 0 && year % 100 !== 0) || year % 400 === 0) && sunday > 29) {
      // Rok przestępny
      sunday -= 29;
    } else {
      // Rok zwykły
      sunday -= 28;
    }
    nextMonth += 1;
  } else if ((month === 4 || month === 6 || month === 9 || month === 11) && sunday > 30) {
    if (sunday > 30) {
      sunday -= 30;
      nextMonth += 1;
    }
  } else if (sunday > 31) {
    sunday -= 31;
    nextMonth += 1;
  }

  return {
    month: month,
    nextMonth: nextMonth,
    monday: monday,
    sunday: sunday,
    year: year,
    nextYear: nextYear
  };
}

function FolderName() {
  var dateInfo = getDateInfo();

  if (dateInfo.nextMonth > 12) {
    dateInfo.nextMonth = "01";
    dateInfo.nextYear += 1;
    var nowaNazwa = dateInfo.monday + "." + dateInfo.month + "." + dateInfo.year + "-" + dateInfo.sunday + "." + dateInfo.nextMonth + "." + dateInfo.nextYear + "r.";
    return nowaNazwa;
  } else {
    var nowaNazwa = Months(dateInfo.nextMonth) + " " + dateInfo.nextYear;
    return nowaNazwa;
  }
}

function FileName() {
  var dateInfo = getDateInfo();

  if (dateInfo.sunday > 28) {
    if (((dateInfo.year % 4 === 0 && dateInfo.year % 100 !== 0) || dateInfo.year % 400 === 0) && dateInfo.sunday > 29) {
      // Rok przestępny
      dateInfo.sunday -= 29;
    } else {
      // Rok zwykły
      dateInfo.sunday -= 28;
    }
    dateInfo.nextMonth += 1;
  } else if ((dateInfo.month === 4 || dateInfo.month === 6 || dateInfo.month === 9 || dateInfo.month === 11) && dateInfo.sunday > 30) {
    if (dateInfo.sunday > 30) {
      dateInfo.sunday -= 30;
      dateInfo.nextMonth += 1;
    }
  } else if (dateInfo.sunday > 31) {
    dateInfo.sunday -= 31;
    dateInfo.nextMonth += 1;
  }

  if (dateInfo.sunday.toString().length < 2) {
    dateInfo.sunday = "0" + dateInfo.sunday;
  }
  if (dateInfo.monday.toString().length < 2) {
    dateInfo.monday = "0" + dateInfo.monday;
  }
  if (dateInfo.month.toString().length < 2) {
    dateInfo.month = "0" + dateInfo.month;
  }
  if (dateInfo.nextMonth.toString().length < 2) {
    dateInfo.nextMonth = "0" + dateInfo.nextMonth;
  }

  if (dateInfo.nextMonth > 12) {
    dateInfo.nextMonth = "01";
    dateInfo.nextYear += 1;
    var nowaNazwa = dateInfo.monday + "." + dateInfo.month + "." + dateInfo.year + "-" + dateInfo.sunday + "." + dateInfo.nextMonth + "." + dateInfo.nextYear + "r.";
    return nowaNazwa;
  } else {
    var nowaNazwa = dateInfo.monday + "." + dateInfo.month + "-" + dateInfo.sunday + "." + dateInfo.nextMonth + "." + dateInfo.year + "r.";
    return nowaNazwa;
  }
}

function Months(nextMonth){
  switch (nextMonth) {
    case 1:
      return "Styczeń";
    case 2:
      return "Luty";
    case 3:
      return "Marzec";
    case 4:
      return "Kwiecień";
    case 5:
      return "Maj";
    case 6:
      return "Czerwiec";
    case 7:
      return "Lipiec";
    case 8:
      return "Sierpień";
    case 9:
      return "Wrzesień";
    case 10:
      return "Październik";
    case 11:
      return "Listopad";
    case 12:
      return "Grudzień";
    default:
      return "Nieznany miesiąc";
  }
}

function RangeRemover(copiedSpreadsheet) {
  var copiedFileSheets = copiedSpreadsheet.getSheets();
  for (var i = 0; i < copiedFileSheets.length; i++) {
    var sheet = copiedFileSheets[i];
    
    // Usuń zawartość komórek w zakresie od B2:J33
    var range = sheet.getRange("B2:J33");
    var values = range.getValues();
    
    for (var row = 0; row < values.length; row++) {
      for (var col = 0; col < values[row].length; col++) {
        values[row][col] = '';
      }
    }
    range.setValues(values);
    Logger.log("Usunięto zawartość zakresu od B2:J33 w zakładce: " + sheet.getName());
  }
}

