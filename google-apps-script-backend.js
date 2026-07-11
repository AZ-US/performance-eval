const SHEET_NAME = 'records';
const SPREADSHEET_ID = '1lZPkM8gHSKKh31szuB2OFkSf5cjoMTnJoGcpnIiB9WY';

function doPost(e) {
  try {
    const payload = parsePayload_(e);
    const record = payload.record || payload;
    upsertRecord_(record);
    return json_({ok: true, savedAt: new Date().toISOString(), id: record.id});
  } catch (err) {
    return json_({ok: false, error: String(err)});
  }
}

function doGet(e) {
  const records = getRecords_();
  const payload = {ok: true, records};
  const callback = e.parameter.callback;
  if (callback) {
    return ContentService
      .createTextOutput(`${callback}(${JSON.stringify(payload)})`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(payload);
}

function parsePayload_(e) {
  const raw = e && e.postData && e.postData.contents ? e.postData.contents : '';
  if (raw) {
    try {
      return JSON.parse(raw);
    } catch (err) {
      // Continue to form payload fallback below.
    }
  }
  const formPayload = e && e.parameter && e.parameter.payload ? e.parameter.payload : '';
  if (formPayload) return JSON.parse(formPayload);
  throw new Error('Missing payload');
}

function upsertRecord_(record) {
  if (!record || !record.id) throw new Error('Missing record.id');
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const row = [
    record.id,
    record.employeeName || '',
    record.period || '',
    record.monthlyDueTotal || 0,
    record.monthlyPaidTotal || 0,
    record.quarterScore || 0,
    record.quarterGrade || '',
    record.savedAt || new Date().toISOString(),
    JSON.stringify(record)
  ];
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === record.id) {
      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return;
    }
  }
  sheet.appendRow(row);
}

function getRecords_() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const records = [];
  for (let i = 1; i < values.length; i++) {
    const raw = values[i][8];
    if (!raw) continue;
    try {
      records.push(JSON.parse(raw));
    } catch (err) {
      // Ignore malformed historical rows.
    }
  }
  return records;
}

function getSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'id',
      'employeeName',
      'period',
      'monthlyDueTotal',
      'monthlyPaidTotal',
      'quarterScore',
      'quarterGrade',
      'savedAt',
      'recordJson'
    ]);
  }
  return sheet;
}

function getSpreadsheet_() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Missing active spreadsheet');
  return ss;
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
