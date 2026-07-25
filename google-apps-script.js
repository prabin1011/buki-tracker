const SHEET_NAME = "Buki Records";

const HEADERS = [
  "Timestamp",
  "Record Type",
  "Unique Key",
  "ID",
  "Date",
  "Client ID",
  "Client",
  "Phone",
  "Naulo Stick",
  "Cig",
  "Naulo Stick Price",
  "Cig Price",
  "Amount",
  "Payment Status",
  "Payment Method",
  "Opening Naulo Stick",
  "Prepared Naulo Stick",
  "Sold Naulo Stick",
  "Opening Naulo Stick Remaining",
  "Prepared Naulo Stick Remaining",
  "Total Naulo Stick Remaining",
  "Opening Cig Pieces",
  "Cig Packets Bought",
  "Cig Pieces Bought",
  "Sold Cig",
  "Opening Cig Pieces Remaining",
  "Bought Cig Pieces Remaining",
  "Total Cig Pieces Remaining",
  "Sales",
  "Paid",
  "Due",
  "Notes",
];

function getSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);
  removeEmptyExtraSheets(spreadsheet);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    return sheet;
  }

  let currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), HEADERS.length)).getValues()[0];
  if (!currentHeaders.includes("Unique Key") && currentHeaders.includes("Record Type")) {
    const recordTypeColumn = currentHeaders.indexOf("Record Type") + 1;
    sheet.insertColumnAfter(recordTypeColumn);
    sheet.getRange(1, recordTypeColumn + 1).setValue("Unique Key");
    currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), HEADERS.length)).getValues()[0];
  }
  if (!currentHeaders.includes("Cig Pieces Bought") && currentHeaders.includes("Cig Packets Bought")) {
    const packetsColumn = currentHeaders.indexOf("Cig Packets Bought") + 1;
    sheet.insertColumnAfter(packetsColumn);
    sheet.getRange(1, packetsColumn + 1).setValue("Cig Pieces Bought");
    currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), HEADERS.length)).getValues()[0];
  }

  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  backfillUniqueKeys(sheet);
  return sheet;
}

function removeEmptyExtraSheets(spreadsheet) {
  const sheets = spreadsheet.getSheets();
  if (sheets.length <= 1) return;
  sheets.forEach((sheet) => {
    if (sheet.getName() === SHEET_NAME) return;
    if (sheet.getLastRow() === 0 && sheet.getLastColumn() === 0) {
      spreadsheet.deleteSheet(sheet);
    }
  });
}

function toRecord(row) {
  return HEADERS.reduce((record, header, index) => {
    record[header] = row[index] ?? "";
    return record;
  }, {});
}

function getRecords() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  return sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues().map(toRecord);
}

function makeUniqueKey(record) {
  const type = record["Record Type"];
  if (type === "Client") return `Client:${record["Client ID"] || record.ID}`;
  if (type === "Transaction") return `Transaction:${record.ID}`;
  if (type === "Daily Summary") return `Daily Summary:${record.Date}`;
  return `${type}:${record.ID || record.Date || Utilities.getUuid()}`;
}

function backfillUniqueKeys(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  const uniqueKeyColumn = HEADERS.indexOf("Unique Key") + 1;
  const range = sheet.getRange(2, 1, lastRow - 1, HEADERS.length);
  const rows = range.getValues();
  const updates = rows.map((row) => {
    const record = toRecord(row);
    return [record["Unique Key"] || makeUniqueKey(record)];
  });
  sheet.getRange(2, uniqueKeyColumn, updates.length, 1).setValues(updates);
}

function findRowByKey(sheet, uniqueKey) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return -1;
  const keyColumn = HEADERS.indexOf("Unique Key") + 1;
  const keys = sheet.getRange(2, keyColumn, lastRow - 1, 1).getValues();
  const index = keys.findIndex((row) => row[0] === uniqueKey);
  return index === -1 ? -1 : index + 2;
}

function upsertRecord(record) {
  const sheet = getSheet();
  const row = HEADERS.map((header) => record[header] ?? "");
  const existingRow = findRowByKey(sheet, record["Unique Key"]);

  if (existingRow === -1) {
    sheet.appendRow(row);
  } else {
    sheet.getRange(existingRow, 1, 1, HEADERS.length).setValues([row]);
  }
}

function sendJson(data, callback) {
  const json = JSON.stringify(data);
  if (callback) {
    return ContentService.createTextOutput(`${callback}(${json});`).setMimeType(
      ContentService.MimeType.JAVASCRIPT
    );
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return sendJson({ ok: true, records: getRecords() }, e.parameter.callback);
}

function doPost(e) {
  const payload = JSON.parse(e.postData.contents || "{}");
  const timestamp = new Date();

  if (payload.type === "client") {
    const c = payload.client;
    upsertRecord({
      "Timestamp": timestamp,
      "Record Type": "Client",
      "Unique Key": `Client:${c.id}`,
      "ID": c.id,
      "Date": c.joined,
      "Client ID": c.id,
      "Client": c.name,
      "Phone": c.phone,
      "Notes": c.notes,
    });
  }

  if (payload.type === "transaction") {
    const t = payload.transaction;
    upsertRecord({
      "Timestamp": timestamp,
      "Record Type": "Transaction",
      "Unique Key": `Transaction:${t.id}`,
      "ID": t.id,
      "Date": t.date,
      "Client ID": t.clientId,
      "Client": t.clientName,
      "Naulo Stick": t.nauloStick,
      "Cig": t.cig,
      "Naulo Stick Price": t.nauloStickPrice,
      "Cig Price": t.cigPrice,
      "Amount": t.amount,
      "Payment Status": t.paid ? "Paid" : "Due",
      "Payment Method": t.method,
      "Notes": t.notes,
    });
  }

  if (payload.type === "daily_summary") {
    const s = payload.summary;
    upsertRecord({
      "Timestamp": timestamp,
      "Record Type": "Daily Summary",
      "Unique Key": `Daily Summary:${s.date}`,
      "Date": s.date,
      "Opening Naulo Stick": s.openingNauloStick,
      "Prepared Naulo Stick": s.nauloStickPrepared,
      "Sold Naulo Stick": s.soldNauloStick,
      "Opening Naulo Stick Remaining": s.openingNauloStickRemaining,
      "Prepared Naulo Stick Remaining": s.preparedNauloStickRemaining,
      "Total Naulo Stick Remaining": s.remainingNauloStick,
      "Opening Cig Pieces": s.openingCigPieces || s.openingCig,
      "Cig Packets Bought": s.cigPacketsBought,
      "Cig Pieces Bought": s.cigPiecesBought,
      "Sold Cig": s.soldCig,
      "Opening Cig Pieces Remaining": s.openingCigRemaining,
      "Bought Cig Pieces Remaining": s.preparedCigRemaining,
      "Total Cig Pieces Remaining": s.remainingCig,
      "Sales": s.sales,
      "Paid": s.paid,
      "Due": s.due,
      "Notes": s.notes,
    });
  }

  return sendJson({ ok: true });
}
