/** Google Apps Script Web App endpoint for MARKET ORACLE. */
function doGet() {
  try {
    var book = SpreadsheetApp.getActiveSpreadsheet();
    var result = {
      recommendations: readSheet_(book, "Recommendations", true),
      allocation: readSheet_(book, "Allocation"),
      alerts: readSheet_(book, "Alerts"),
      logs: readSheet_(book, "Logs"),
      updatedAt: new Date().toISOString(),
    };
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ error: error.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function readSheet_(book, name, required) {
  var sheet = book.getSheetByName(name);
  if (!sheet) {
    if (required) throw new Error('Required sheet "' + name + '" was not found');
    return [];
  }
  var values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];
  var headers = values.shift().map(function (header) { return String(header).trim(); });
  return values.filter(function (row) {
    return row.some(function (cell) { return String(cell).trim() !== ""; });
  }).map(function (row) {
    return headers.reduce(function (record, header, index) {
      if (header) record[header] = row[index];
      return record;
    }, {});
  });
}
