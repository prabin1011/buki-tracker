const SHEETS = {
  clients: "Clients",
  transactions: "Transactions",
  dailySummary: "Daily Summary",
};

function ensureSheet(name, headers) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.appendRow(headers);
  return sheet;
}

function doPost(e) {
  const payload = JSON.parse(e.postData.contents || "{}");

  if (payload.type === "client") {
    const c = payload.client;
    ensureSheet(SHEETS.clients, ["ID", "Name", "Phone", "Notes", "Joined"]).appendRow([
      c.id,
      c.name,
      c.phone,
      c.notes,
      c.joined,
    ]);
  }

  if (payload.type === "transaction") {
    const t = payload.transaction;
    ensureSheet(SHEETS.transactions, [
      "ID",
      "Date",
      "ClientID",
      "Client",
      "Naulo Stick",
      "Cig",
      "Naulo Stick Price",
      "Cig Price",
      "Amount",
      "Payment Status",
      "Payment Method",
      "Notes",
    ]).appendRow([
      t.id,
      t.date,
      t.clientId,
      t.clientName,
      t.nauloStick,
      t.cig,
      t.nauloStickPrice,
      t.cigPrice,
      t.amount,
      t.paid ? "Paid" : "Due",
      t.method,
      t.notes,
    ]);
  }

  if (payload.type === "daily_summary") {
    const s = payload.summary;
    ensureSheet(SHEETS.dailySummary, [
      "Date",
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
    ]).appendRow([
      s.date,
      s.openingNauloStick,
      s.nauloStickPrepared,
      s.soldNauloStick,
      s.openingNauloStickRemaining,
      s.preparedNauloStickRemaining,
      s.remainingNauloStick,
      s.openingCig,
      s.cigPrepared,
      s.soldCig,
      s.openingCigRemaining,
      s.preparedCigRemaining,
      s.remainingCig,
      s.sales,
      s.paid,
      s.due,
      s.notes,
    ]);
  }

  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(
    ContentService.MimeType.JSON
  );
}
