const SHEET_NAME = "Buki Records";

const HEADERS = [
  "Timestamp",
  "Record Type",
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
  "Opening Cig",
  "Prepared Cig",
  "Sold Cig",
  "Opening Cig Remaining",
  "Prepared Cig Remaining",
  "Total Cig Remaining",
  "Sales",
  "Paid",
  "Due",
  "Notes",
];

function getSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function appendRecord(record) {
  const sheet = getSheet();
  sheet.appendRow(HEADERS.map((header) => record[header] ?? ""));
}

function doPost(e) {
  const payload = JSON.parse(e.postData.contents || "{}");
  const timestamp = new Date();

  if (payload.type === "client") {
    const c = payload.client;
    appendRecord({
      "Timestamp": timestamp,
      "Record Type": "Client",
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
    appendRecord({
      "Timestamp": timestamp,
      "Record Type": "Transaction",
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
    appendRecord({
      "Timestamp": timestamp,
      "Record Type": "Daily Summary",
      "Date": s.date,
      "Opening Naulo Stick": s.openingNauloStick,
      "Prepared Naulo Stick": s.nauloStickPrepared,
      "Sold Naulo Stick": s.soldNauloStick,
      "Opening Naulo Stick Remaining": s.openingNauloStickRemaining,
      "Prepared Naulo Stick Remaining": s.preparedNauloStickRemaining,
      "Total Naulo Stick Remaining": s.remainingNauloStick,
      "Opening Cig": s.openingCig,
      "Prepared Cig": s.cigPrepared,
      "Sold Cig": s.soldCig,
      "Opening Cig Remaining": s.openingCigRemaining,
      "Prepared Cig Remaining": s.preparedCigRemaining,
      "Total Cig Remaining": s.remainingCig,
      "Sales": s.sales,
      "Paid": s.paid,
      "Due": s.due,
      "Notes": s.notes,
    });
  }

  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(
    ContentService.MimeType.JSON
  );
}
