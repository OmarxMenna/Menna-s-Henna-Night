const SPREADSHEET_ID = "11eJduYGhpW3me445S54dp-CGT7ugaGH9gSZmJsg4bKw";
const SHEET_NAME = "RSVPs";
const HEADERS = ["Timestamp", "Name", "Answer"];

function doGet(e) {
  const guests = getGuests_();
  const payload = JSON.stringify({ guests });
  const callback = e.parameter.callback;

  if (callback) {
    return ContentService
      .createTextOutput(`${callback}(${payload});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents || "{}");
  const name = String(data.name || "").trim();
  const answer = String(data.answer || "").trim();

  if (!name || !answer) {
    return json_({ ok: false, error: "Missing name or answer." });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getSheet_();
    const rows = sheet.getDataRange().getValues();
    const nextRow = [new Date(), name, answer];
    const existingIndex = rows.findIndex((row, index) => {
      return index > 0 && String(row[1]).trim().toLowerCase() === name.toLowerCase();
    });

    if (existingIndex >= 0) {
      sheet.getRange(existingIndex + 1, 1, 1, nextRow.length).setValues([nextRow]);
    } else {
      sheet.appendRow(nextRow);
    }

    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function getGuests_() {
  const sheet = getSheet_();
  const rows = sheet.getDataRange().getValues().slice(1);

  return rows
    .filter((row) => row[1])
    .map((row) => ({
      name: String(row[1]),
      answer: String(row[2])
    }));
}

function getSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const needsHeaders = HEADERS.some((header, index) => firstRow[index] !== header);

  if (needsHeaders) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }

  return sheet;
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
