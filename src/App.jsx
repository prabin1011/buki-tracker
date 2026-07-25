import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

const APP_NAME = "Buki Tracker";
const CURRENCY = "NPR";
const STORE_KEY = "buki-tracker-v1";
const SHEETS_URL = import.meta.env.VITE_GOOGLE_SHEET_WEBAPP_URL || "";
const PAYMENT_METHODS = ["eSewa", "Online", "Hard cash"];
const PRODUCTS = [
  { key: "nauloStick", label: "Naulo Stick", unit: "stick", defaultPrice: 200 },
  { key: "cig", label: "Cig", unit: "cig", defaultPrice: 30 },
];

function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

function today() {
  return new Date().toISOString().slice(0, 10);
}

function money(n) {
  return `${CURRENCY} ${Number(n || 0).toLocaleString()}`;
}

function num(n) {
  return Number(n || 0);
}

function fmtDate(date) {
  if (!date) return "-";
  return new Date(`${date}T00:00:00`).toLocaleDateString("en-US", {
    month: "short",
    day: "numeric",
    year: "numeric",
  });
}

function normalizeName(name) {
  return String(name || "").trim().toLowerCase().replace(/\s+/g, " ");
}

function normalizePhone(phone) {
  return String(phone || "").replace(/\D/g, "");
}

function normalizeExcelDate(value) {
  if (!value) return today();
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
    }
  }
  const text = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) return parsed.toISOString().slice(0, 10);
  return today();
}

function emptyStock(date = today()) {
  return {
    date,
    nauloStickPrepared: 0,
    cigPrepared: 0,
    openingNauloStick: 0,
    openingCig: 0,
    notes: "",
  };
}

function defaultForm() {
  return {
    clientId: "",
    nauloStick: 0,
    cig: 0,
    nauloStickPrice: PRODUCTS[0].defaultPrice,
    cigPrice: PRODUCTS[1].defaultPrice,
    paid: true,
    method: PAYMENT_METHODS[0],
    notes: "",
  };
}

function loadInitialData() {
  try {
    const raw = localStorage.getItem(STORE_KEY);
    if (!raw) return null;
    const data = JSON.parse(raw);
    return {
      clients: Array.isArray(data.clients) ? data.clients : [],
      transactions: Array.isArray(data.transactions) ? data.transactions : [],
      stockByDate: data.stockByDate && typeof data.stockByDate === "object" ? data.stockByDate : {},
    };
  } catch {
    return null;
  }
}

function buildWorkbook(clients, transactions, stockRows) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      ["ID", "Name", "Phone", "Notes", "Joined"],
      ...clients.map((c) => [c.id, c.name, c.phone, c.notes, c.joined]),
    ]),
    "Clients"
  );
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      [
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
      ],
      ...transactions.map((t) => [
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
        t.paid ? t.method : "",
        t.notes,
      ]),
    ]),
    "Transactions"
  );
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      [
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
        "Due",
        "Paid",
        "Notes",
      ],
      ...stockRows.map((r) => [
        r.date,
        r.openingNauloStick,
        r.nauloStickPrepared,
        r.soldNauloStick,
        r.openingNauloStickRemaining,
        r.preparedNauloStickRemaining,
        r.remainingNauloStick,
        r.openingCig,
        r.cigPrepared,
        r.soldCig,
        r.openingCigRemaining,
        r.preparedCigRemaining,
        r.remainingCig,
        r.sales,
        r.due,
        r.paid,
        r.notes,
      ]),
    ]),
    "Daily Summary"
  );
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      ["Setting", "Value"],
      ["app_name", APP_NAME],
      ["currency", CURRENCY],
      ["products", PRODUCTS.map((p) => p.label).join(", ")],
    ]),
    "Settings"
  );
  return wb;
}

async function syncToGoogleSheet(payload) {
  if (!SHEETS_URL) return { skipped: true };
  const response = await fetch(SHEETS_URL, {
    method: "POST",
    mode: "no-cors",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify(payload),
  });
  return { skipped: false, response };
}

const inputStyle = {
  width: "100%",
  border: "1px solid #d5d7db",
  borderRadius: 6,
  padding: "9px 10px",
  fontSize: 14,
  fontFamily: "inherit",
  background: "#fff",
};

const buttonStyle = {
  border: "none",
  borderRadius: 6,
  padding: "9px 12px",
  fontFamily: "inherit",
  fontWeight: 800,
  cursor: "pointer",
};

function Field({ label, children }) {
  return (
    <label style={{ display: "block" }}>
      <span style={{ display: "block", fontSize: 11, fontWeight: 900, textTransform: "uppercase", color: "#667085", marginBottom: 5 }}>
        {label}
      </span>
      {children}
    </label>
  );
}

function Modal({ title, children, onClose }) {
  return (
    <div
      onClick={onClose}
      style={{
        position: "fixed",
        inset: 0,
        zIndex: 100,
        background: "rgba(15,23,42,.45)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        padding: 16,
      }}
    >
      <div onClick={(e) => e.stopPropagation()} style={{ width: 520, maxWidth: "100%", background: "#fff", borderRadius: 8, border: "1px solid #1f2937" }}>
        <div style={{ padding: "14px 16px", borderBottom: "1px solid #eaecf0", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <strong>{title}</strong>
          <button onClick={onClose} style={{ ...buttonStyle, background: "transparent", padding: 4, color: "#667085", fontSize: 20 }}>
            x
          </button>
        </div>
        <div style={{ padding: 16 }}>{children}</div>
      </div>
    </div>
  );
}

function Stat({ label, value, tone = "dark" }) {
  const color = tone === "bad" ? "#b42318" : tone === "good" ? "#067647" : "#101828";
  return (
    <div style={{ padding: "12px 14px", background: "#fff", border: "1px solid #eaecf0", borderRadius: 8 }}>
      <div style={{ fontSize: 11, color: "#667085", textTransform: "uppercase", fontWeight: 900 }}>{label}</div>
      <div style={{ marginTop: 4, fontSize: 20, fontWeight: 900, color }}>{value}</div>
    </div>
  );
}

function NumberInput({ value, onChange, ...props }) {
  const [draft, setDraft] = useState(String(num(value)));

  useEffect(() => {
    setDraft(String(num(value)));
  }, [value]);

  function commit(nextDraft) {
    const clean = nextDraft.trim();
    onChange(clean === "" ? 0 : Math.max(0, Number(clean) || 0));
  }

  return (
    <input
      {...props}
      type="number"
      min="0"
      value={draft}
      onChange={(e) => {
        setDraft(e.target.value);
        if (e.target.value !== "") commit(e.target.value);
      }}
      onBlur={() => {
        if (draft === "") {
          setDraft("0");
          onChange(0);
        } else {
          commit(draft);
        }
      }}
    />
  );
}

export default function App() {
  const initial = loadInitialData();
  const [clients, setClients] = useState(initial?.clients || []);
  const [transactions, setTransactions] = useState(initial?.transactions || []);
  const [stockByDate, setStockByDate] = useState(initial?.stockByDate || {});
  const [date, setDate] = useState(today());
  const [clientModal, setClientModal] = useState(false);
  const [txForm, setTxForm] = useState(defaultForm());
  const [clientForm, setClientForm] = useState({ name: "", phone: "", notes: "" });
  const [message, setMessage] = useState("");
  const [syncState, setSyncState] = useState(SHEETS_URL ? "Google Sheets sync on" : "Excel/local mode");
  const [loadedFileName, setLoadedFileName] = useState("");

  const stock = stockByDate[date] || emptyStock(date);

  useEffect(() => {
    localStorage.setItem(STORE_KEY, JSON.stringify({ clients, transactions, stockByDate }));
  }, [clients, transactions, stockByDate]);

  const dayTransactions = useMemo(() => transactions.filter((t) => t.date === date), [transactions, date]);

  const daily = useMemo(() => {
    const soldNauloStick = dayTransactions.reduce((sum, t) => sum + num(t.nauloStick), 0);
    const soldCig = dayTransactions.reduce((sum, t) => sum + num(t.cig), 0);
    const sales = dayTransactions.reduce((sum, t) => sum + num(t.amount), 0);
    const paid = dayTransactions.filter((t) => t.paid).reduce((sum, t) => sum + num(t.amount), 0);
    const due = sales - paid;
    const openingNauloStickUsed = Math.min(num(stock.openingNauloStick), soldNauloStick);
    const preparedNauloStickUsed = Math.max(0, soldNauloStick - openingNauloStickUsed);
    const openingCigUsed = Math.min(num(stock.openingCig), soldCig);
    const preparedCigUsed = Math.max(0, soldCig - openingCigUsed);
    const openingNauloStickRemaining = num(stock.openingNauloStick) - openingNauloStickUsed;
    const preparedNauloStickRemaining = num(stock.nauloStickPrepared) - preparedNauloStickUsed;
    const openingCigRemaining = num(stock.openingCig) - openingCigUsed;
    const preparedCigRemaining = num(stock.cigPrepared) - preparedCigUsed;
    return {
      soldNauloStick,
      soldCig,
      sales,
      paid,
      due,
      openingNauloStickUsed,
      preparedNauloStickUsed,
      openingCigUsed,
      preparedCigUsed,
      openingNauloStickRemaining,
      preparedNauloStickRemaining,
      openingCigRemaining,
      preparedCigRemaining,
      remainingNauloStick: openingNauloStickRemaining + preparedNauloStickRemaining,
      remainingCig: openingCigRemaining + preparedCigRemaining,
    };
  }, [dayTransactions, stock]);

  const allStockRows = useMemo(() => {
    const dates = Array.from(new Set([...Object.keys(stockByDate), ...transactions.map((t) => t.date)])).sort();
    return dates.map((d) => {
      const row = stockByDate[d] || emptyStock(d);
      const rows = transactions.filter((t) => t.date === d);
      const soldNauloStick = rows.reduce((sum, t) => sum + num(t.nauloStick), 0);
      const soldCig = rows.reduce((sum, t) => sum + num(t.cig), 0);
      const sales = rows.reduce((sum, t) => sum + num(t.amount), 0);
      const paid = rows.filter((t) => t.paid).reduce((sum, t) => sum + num(t.amount), 0);
      const openingNauloStickUsed = Math.min(num(row.openingNauloStick), soldNauloStick);
      const preparedNauloStickUsed = Math.max(0, soldNauloStick - openingNauloStickUsed);
      const openingCigUsed = Math.min(num(row.openingCig), soldCig);
      const preparedCigUsed = Math.max(0, soldCig - openingCigUsed);
      const openingNauloStickRemaining = num(row.openingNauloStick) - openingNauloStickUsed;
      const preparedNauloStickRemaining = num(row.nauloStickPrepared) - preparedNauloStickUsed;
      const openingCigRemaining = num(row.openingCig) - openingCigUsed;
      const preparedCigRemaining = num(row.cigPrepared) - preparedCigUsed;
      return {
        ...row,
        soldNauloStick,
        soldCig,
        sales,
        paid,
        due: sales - paid,
        openingNauloStickRemaining,
        preparedNauloStickRemaining,
        remainingNauloStick: openingNauloStickRemaining + preparedNauloStickRemaining,
        openingCigRemaining,
        preparedCigRemaining,
        remainingCig: openingCigRemaining + preparedCigRemaining,
      };
    });
  }, [stockByDate, transactions]);

  function updateStock(key, value) {
    setStockByDate((prev) => ({
      ...prev,
      [date]: { ...emptyStock(date), ...(prev[date] || {}), [key]: value },
    }));
  }

  async function addClient() {
    if (!clientForm.name.trim()) {
      setMessage("Client name is required.");
      return;
    }
    const client = { id: uid(), ...clientForm, name: clientForm.name.trim(), joined: today() };
    setClients((prev) => [...prev, client]);
    setTxForm((prev) => ({ ...prev, clientId: client.id }));
    setClientForm({ name: "", phone: "", notes: "" });
    setClientModal(false);
    setMessage(`${client.name} added.`);
    try {
      const result = await syncToGoogleSheet({ type: "client", client });
      setSyncState(result.skipped ? "Excel/local mode" : "Client synced to Google Sheets");
    } catch {
      setSyncState("Google Sheets sync failed");
    }
  }

  async function addTransaction() {
    const client = clients.find((c) => c.id === txForm.clientId);
    if (!client) {
      setMessage("Choose a client first.");
      return;
    }
    const nauloStick = Math.max(0, num(txForm.nauloStick));
    const cig = Math.max(0, num(txForm.cig));
    if (nauloStick + cig <= 0) {
      setMessage("Add at least one item.");
      return;
    }
    const amount = nauloStick * num(txForm.nauloStickPrice) + cig * num(txForm.cigPrice);
    const record = {
      id: uid(),
      date,
      clientId: client.id,
      clientName: client.name,
      nauloStick,
      cig,
      nauloStickPrice: num(txForm.nauloStickPrice),
      cigPrice: num(txForm.cigPrice),
      paid: Boolean(txForm.paid),
      method: txForm.paid ? txForm.method : "",
      amount,
      notes: txForm.notes,
    };
    setTransactions((prev) => [record, ...prev]);
    setTxForm((prev) => ({ ...defaultForm(), clientId: prev.clientId }));
    setMessage("Transaction recorded.");
    try {
      const result = await syncToGoogleSheet({ type: "transaction", transaction: record });
      setSyncState(result.skipped ? "Excel/local mode" : "Transaction synced to Google Sheets");
    } catch {
      setSyncState("Google Sheets sync failed");
    }
  }

  function togglePaid(id) {
    setTransactions((prev) =>
      prev.map((t) => (t.id === id ? { ...t, paid: !t.paid, method: !t.paid ? PAYMENT_METHODS[0] : "" } : t))
    );
  }

  async function carryForward() {
    const next = new Date(`${date}T00:00:00`);
    next.setDate(next.getDate() + 1);
    const nextDate = next.toISOString().slice(0, 10);
    setStockByDate((prev) => ({
      ...prev,
      [nextDate]: {
        ...emptyStock(nextDate),
        ...(prev[nextDate] || {}),
        openingNauloStick: num(prev[nextDate]?.openingNauloStick) + Math.max(0, daily.remainingNauloStick),
        openingCig: num(prev[nextDate]?.openingCig) + Math.max(0, daily.remainingCig),
      },
    }));
    setDate(nextDate);
    setMessage(`Remaining stock carried to ${fmtDate(nextDate)}.`);
    try {
      const result = await syncToGoogleSheet({
        type: "daily_summary",
        summary: {
          date,
          ...stock,
          ...daily,
        },
      });
      setSyncState(result.skipped ? "Excel/local mode" : "Daily summary synced to Google Sheets");
    } catch {
      setSyncState("Google Sheets sync failed");
    }
  }

  function exportExcel(saveAsNew = false) {
    const fallback = "buki_tracker.xlsx";
    const fileName = saveAsNew ? `buki_tracker_${date}.xlsx` : loadedFileName || fallback;
    XLSX.writeFile(buildWorkbook(clients, transactions, allStockRows), fileName);
    setMessage(saveAsNew ? "New Excel file exported." : `Excel saved as ${fileName}.`);
  }

  function importExcel(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array", cellDates: false });
        const clientRows = XLSX.utils.sheet_to_json(wb.Sheets.Clients || {});
        const txRows = XLSX.utils.sheet_to_json(wb.Sheets.Transactions || {});
        const dailyRows = XLSX.utils.sheet_to_json(wb.Sheets["Daily Summary"] || {});
        const nextStock = {};
        dailyRows.forEach((r) => {
          if (!r.Date) return;
          const normalizedDate = normalizeExcelDate(r.Date);
          nextStock[normalizedDate] = {
            date: normalizedDate,
            openingNauloStick: num(r["Opening Naulo Stick"]),
            nauloStickPrepared: num(r["Prepared Naulo Stick"]),
            openingCig: num(r["Opening Cig"]),
            cigPrepared: num(r["Prepared Cig"]),
            notes: r.Notes || "",
          };
        });
        const nextClients = [];
        const idMap = new Map();
        const phoneMap = new Map();
        const nameMap = new Map();
        clientRows.forEach((r) => {
          const incoming = {
            id: r.ID || uid(),
            name: String(r.Name || "").trim(),
            phone: String(r.Phone || "").trim(),
            notes: r.Notes || "",
            joined: normalizeExcelDate(r.Joined),
          };
          if (!incoming.name) return;
          const phoneKey = normalizePhone(incoming.phone);
          const nameKey = normalizeName(incoming.name);
          const existing =
            (incoming.id && idMap.get(incoming.id)) ||
            (phoneKey && phoneMap.get(phoneKey)) ||
            (nameKey && nameMap.get(nameKey));
          if (existing) {
            existing.phone = existing.phone || incoming.phone;
            existing.notes = existing.notes || incoming.notes;
            idMap.set(incoming.id, existing);
            return;
          }
          nextClients.push(incoming);
          idMap.set(incoming.id, incoming);
          if (phoneKey) phoneMap.set(phoneKey, incoming);
          if (nameKey) nameMap.set(nameKey, incoming);
        });
        const mappedTransactions = txRows.map((r) => {
          const rawClientId = r.ClientID || "";
          const rawClientName = String(r.Client || "").trim();
          const matched =
            (rawClientId && idMap.get(rawClientId)) ||
            (rawClientName && nameMap.get(normalizeName(rawClientName)));
          let client = matched;
          if (!client && rawClientName) {
            client = { id: uid(), name: rawClientName, phone: "", notes: "", joined: normalizeExcelDate(r.Date) };
            nextClients.push(client);
            idMap.set(client.id, client);
            nameMap.set(normalizeName(client.name), client);
          }
          return {
            id: r.ID || uid(),
            date: normalizeExcelDate(r.Date),
            clientId: client?.id || rawClientId,
            clientName: client?.name || rawClientName,
            nauloStick: num(r["Naulo Stick"]),
            cig: num(r.Cig),
            nauloStickPrice: num(r["Naulo Stick Price"]) || PRODUCTS[0].defaultPrice,
            cigPrice: num(r["Cig Price"]) || PRODUCTS[1].defaultPrice,
            amount:
              num(r.Amount) ||
              num(r["Naulo Stick"]) * (num(r["Naulo Stick Price"]) || PRODUCTS[0].defaultPrice) +
                num(r.Cig) * (num(r["Cig Price"]) || PRODUCTS[1].defaultPrice),
            paid: r["Payment Status"] !== "Due",
            method: r["Payment Method"] || "",
            notes: r.Notes || "",
          };
        });
        setClients(nextClients);
        setTransactions(mappedTransactions);
        setStockByDate(nextStock);
        setLoadedFileName(file.name || "buki_tracker.xlsx");
        setMessage(`${file.name || "Excel file"} loaded. Save Excel will use the same filename.`);
      } catch {
        setMessage("Could not import that Excel file.");
      } finally {
        event.target.value = "";
      }
    };
    reader.readAsArrayBuffer(file);
  }

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&display=swap');
        *{box-sizing:border-box}
        body{margin:0;background:#f3f5f7;color:#101828;font-family:Inter,Segoe UI,system-ui,sans-serif}
        button:hover{opacity:.86}
        input:focus,select:focus,textarea:focus{outline:2px solid #10182822;border-color:#101828}
        table{border-collapse:collapse;width:100%}
        th{text-align:left;font-size:11px;text-transform:uppercase;color:#667085;padding:10px;border-bottom:1px solid #eaecf0}
        td{padding:11px 10px;border-bottom:1px solid #f0f2f5;font-size:13px;vertical-align:top}
        @media(max-width:850px){.layout{grid-template-columns:1fr!important}.stats{grid-template-columns:repeat(2,1fr)!important}.wide{overflow-x:auto}.formgrid{grid-template-columns:1fr!important}}
      `}</style>

      {clientModal && (
        <Modal title="New Client" onClose={() => setClientModal(false)}>
          <div style={{ display: "grid", gap: 12 }}>
            <Field label="Client name">
              <input style={inputStyle} value={clientForm.name} onChange={(e) => setClientForm((p) => ({ ...p, name: e.target.value }))} />
            </Field>
            <Field label="Phone">
              <input style={inputStyle} value={clientForm.phone} onChange={(e) => setClientForm((p) => ({ ...p, phone: e.target.value }))} />
            </Field>
            <Field label="Notes">
              <textarea style={{ ...inputStyle, minHeight: 70, resize: "vertical" }} value={clientForm.notes} onChange={(e) => setClientForm((p) => ({ ...p, notes: e.target.value }))} />
            </Field>
            <button onClick={addClient} style={{ ...buttonStyle, background: "#101828", color: "#fff" }}>
              Add Client
            </button>
          </div>
        </Modal>
      )}

      <div style={{ maxWidth: 1220, margin: "0 auto", padding: 18 }}>
        <header style={{ display: "flex", justifyContent: "space-between", gap: 16, alignItems: "center", marginBottom: 14, flexWrap: "wrap" }}>
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <div style={{ width: 38, height: 38, borderRadius: 8, background: "#101828", color: "#fff", display: "grid", placeItems: "center", fontWeight: 900 }}>B</div>
              <div>
                <h1 style={{ margin: 0, fontSize: 24, letterSpacing: 0 }}>{APP_NAME}</h1>
                <div style={{ fontSize: 13, color: "#667085" }}>Daily client, stock, due, and payment tracker</div>
                <div style={{ fontSize: 12, color: SHEETS_URL ? "#067647" : "#667085", marginTop: 2 }}>{syncState}</div>
                {loadedFileName && <div style={{ fontSize: 12, color: "#667085", marginTop: 2 }}>Loaded workbook: {loadedFileName}</div>}
              </div>
            </div>
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <label style={{ ...buttonStyle, background: "#fff", border: "1px solid #d5d7db", color: "#101828" }}>
              Load Excel
              <input type="file" accept=".xlsx" onChange={importExcel} style={{ display: "none" }} />
            </label>
            <button onClick={() => exportExcel(false)} style={{ ...buttonStyle, background: "#101828", color: "#fff" }}>
              Save Excel
            </button>
            <button onClick={() => exportExcel(true)} style={{ ...buttonStyle, background: "#fff", color: "#101828", border: "1px solid #d5d7db" }}>
              Save As New
            </button>
          </div>
        </header>

        {message && (
          <div style={{ background: "#ecfdf3", border: "1px solid #abefc6", color: "#067647", padding: "9px 12px", borderRadius: 8, marginBottom: 12, fontWeight: 700, fontSize: 13 }}>
            {message}
          </div>
        )}

        <section style={{ background: "#fff", border: "1px solid #eaecf0", borderRadius: 8, padding: 14, marginBottom: 12 }}>
          <div style={{ display: "grid", gridTemplateColumns: "180px 1fr auto", gap: 12, alignItems: "end" }} className="formgrid">
            <Field label="Calendar date">
              <input type="date" style={inputStyle} value={date} onChange={(e) => setDate(e.target.value)} />
            </Field>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 10 }} className="formgrid">
              <Field label="Opening Naulo Stick">
                <NumberInput style={inputStyle} value={stock.openingNauloStick} onChange={(value) => updateStock("openingNauloStick", value)} />
              </Field>
              <Field label="Prepared Naulo Stick">
                <NumberInput style={inputStyle} value={stock.nauloStickPrepared} onChange={(value) => updateStock("nauloStickPrepared", value)} />
              </Field>
              <Field label="Opening Cig">
                <NumberInput style={inputStyle} value={stock.openingCig} onChange={(value) => updateStock("openingCig", value)} />
              </Field>
              <Field label="Prepared Cig">
                <NumberInput style={inputStyle} value={stock.cigPrepared} onChange={(value) => updateStock("cigPrepared", value)} />
              </Field>
            </div>
            <button onClick={carryForward} style={{ ...buttonStyle, background: "#eef4ff", color: "#3538cd", border: "1px solid #c7d7fe" }}>
              Carry to Next Day
            </button>
          </div>
        </section>

        <section className="stats" style={{ display: "grid", gridTemplateColumns: "repeat(6,1fr)", gap: 10, marginBottom: 12 }}>
          <Stat label="Stick sold" value={daily.soldNauloStick} />
          <Stat label="Cig sold" value={daily.soldCig} />
          <Stat label="Stick carry total" value={daily.remainingNauloStick} tone={daily.remainingNauloStick < 0 ? "bad" : "dark"} />
          <Stat label="Cig carry total" value={daily.remainingCig} tone={daily.remainingCig < 0 ? "bad" : "dark"} />
          <Stat label="Paid" value={money(daily.paid)} tone="good" />
          <Stat label="Due" value={money(daily.due)} tone={daily.due > 0 ? "bad" : "dark"} />
        </section>

        <section style={{ background: "#fff", border: "1px solid #eaecf0", borderRadius: 8, padding: 14, marginBottom: 12 }}>
          <div style={{ display: "flex", justifyContent: "space-between", gap: 12, alignItems: "center", flexWrap: "wrap", marginBottom: 10 }}>
            <strong>Remaining Stock</strong>
            <span style={{ fontSize: 12, color: "#667085" }}>Sales consume opening stock first, then prepared stock.</span>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 10 }} className="stats">
            <Stat label="Opening stick left" value={daily.openingNauloStickRemaining} tone={daily.openingNauloStickRemaining < 0 ? "bad" : "dark"} />
            <Stat label="Prepared stick left" value={daily.preparedNauloStickRemaining} tone={daily.preparedNauloStickRemaining < 0 ? "bad" : "dark"} />
            <Stat label="Opening cig left" value={daily.openingCigRemaining} tone={daily.openingCigRemaining < 0 ? "bad" : "dark"} />
            <Stat label="Prepared cig left" value={daily.preparedCigRemaining} tone={daily.preparedCigRemaining < 0 ? "bad" : "dark"} />
          </div>
        </section>

        <main className="layout" style={{ display: "grid", gridTemplateColumns: "360px 1fr", gap: 12, alignItems: "start" }}>
          <section style={{ background: "#fff", border: "1px solid #eaecf0", borderRadius: 8, padding: 14 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
              <strong>Record Transaction</strong>
              <button onClick={() => setClientModal(true)} style={{ ...buttonStyle, background: "#f9fafb", color: "#101828", border: "1px solid #d5d7db", padding: "7px 10px" }}>
                New Client
              </button>
            </div>
            <div style={{ display: "grid", gap: 11 }}>
              <Field label="Client">
                <select style={inputStyle} value={txForm.clientId} onChange={(e) => setTxForm((p) => ({ ...p, clientId: e.target.value }))}>
                  <option value="">Select client</option>
                  {clients.map((c) => (
                    <option key={c.id} value={c.id}>
                      {c.name}
                    </option>
                  ))}
                </select>
              </Field>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                <Field label="Naulo Stick taken">
                  <NumberInput style={inputStyle} value={txForm.nauloStick} onChange={(value) => setTxForm((p) => ({ ...p, nauloStick: value }))} />
                </Field>
                <Field label="Cig taken">
                  <NumberInput style={inputStyle} value={txForm.cig} onChange={(value) => setTxForm((p) => ({ ...p, cig: value }))} />
                </Field>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                <Field label="Stick price">
                  <NumberInput style={inputStyle} value={txForm.nauloStickPrice} onChange={(value) => setTxForm((p) => ({ ...p, nauloStickPrice: value }))} />
                </Field>
                <Field label="Cig price">
                  <NumberInput style={inputStyle} value={txForm.cigPrice} onChange={(value) => setTxForm((p) => ({ ...p, cigPrice: value }))} />
                </Field>
              </div>
              <Field label="Payment">
                <select style={inputStyle} value={txForm.paid ? "Paid" : "Due"} onChange={(e) => setTxForm((p) => ({ ...p, paid: e.target.value === "Paid" }))}>
                  <option>Paid</option>
                  <option>Due</option>
                </select>
              </Field>
              {txForm.paid && (
                <Field label="Payment method">
                  <select style={inputStyle} value={txForm.method} onChange={(e) => setTxForm((p) => ({ ...p, method: e.target.value }))}>
                    {PAYMENT_METHODS.map((m) => (
                      <option key={m}>{m}</option>
                    ))}
                  </select>
                </Field>
              )}
              <Field label="Notes">
                <textarea style={{ ...inputStyle, resize: "vertical", minHeight: 72 }} value={txForm.notes} onChange={(e) => setTxForm((p) => ({ ...p, notes: e.target.value }))} />
              </Field>
              <div style={{ background: "#f9fafb", border: "1px solid #eaecf0", borderRadius: 8, padding: 10, fontSize: 13 }}>
                Total: <strong>{money(num(txForm.nauloStick) * num(txForm.nauloStickPrice) + num(txForm.cig) * num(txForm.cigPrice))}</strong>
              </div>
              <button onClick={addTransaction} style={{ ...buttonStyle, background: "#101828", color: "#fff", padding: 12 }}>
                Record for {fmtDate(date)}
              </button>
            </div>
          </section>

          <section style={{ display: "grid", gap: 12 }}>
            <div style={{ background: "#fff", border: "1px solid #eaecf0", borderRadius: 8 }}>
              <div style={{ padding: "13px 14px", borderBottom: "1px solid #eaecf0", display: "flex", justifyContent: "space-between", gap: 10, flexWrap: "wrap" }}>
                <strong>Transactions on {fmtDate(date)}</strong>
                <span style={{ color: "#667085", fontSize: 13 }}>{dayTransactions.length} entries</span>
              </div>
              <div className="wide">
                <table>
                  <thead>
                    <tr>
                      <th>Client</th>
                      <th>Items</th>
                      <th>Amount</th>
                      <th>Payment</th>
                      <th>Notes</th>
                    </tr>
                  </thead>
                  <tbody>
                    {dayTransactions.length === 0 ? (
                      <tr>
                        <td colSpan="5" style={{ color: "#667085", textAlign: "center", padding: 28 }}>
                          No transactions recorded for this day.
                        </td>
                      </tr>
                    ) : (
                      dayTransactions.map((t) => (
                        <tr key={t.id}>
                          <td>
                            <strong>{t.clientName}</strong>
                          </td>
                          <td>
                            {t.nauloStick} Naulo Stick<br />
                            {t.cig} Cig
                          </td>
                          <td>
                            <strong>{money(t.amount)}</strong>
                          </td>
                          <td>
                            <button onClick={() => togglePaid(t.id)} style={{ ...buttonStyle, background: t.paid ? "#ecfdf3" : "#fef3f2", color: t.paid ? "#067647" : "#b42318", border: `1px solid ${t.paid ? "#abefc6" : "#fecdca"}`, padding: "5px 9px" }}>
                              {t.paid ? `Paid - ${t.method}` : "Due"}
                            </button>
                          </td>
                          <td style={{ color: "#667085" }}>{t.notes || "-"}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }} className="formgrid">
              <div style={{ background: "#fff", border: "1px solid #eaecf0", borderRadius: 8, padding: 14 }}>
                <strong>Regularity</strong>
                <div style={{ marginTop: 10, display: "grid", gap: 8 }}>
                  {clients.length === 0 ? (
                    <div style={{ color: "#667085", fontSize: 13 }}>Add clients to see who shows up regularly.</div>
                  ) : (
                    clients.map((c) => {
                      const visits = transactions.filter((t) => t.clientId === c.id).length;
                      const lastVisit = transactions.find((t) => t.clientId === c.id)?.date;
                      return (
                        <div key={c.id} style={{ display: "flex", justifyContent: "space-between", gap: 10, borderBottom: "1px solid #f0f2f5", paddingBottom: 7 }}>
                          <span>
                            <strong>{c.name}</strong>
                            <span style={{ color: "#667085", display: "block", fontSize: 12 }}>{c.phone || "No phone"}</span>
                          </span>
                          <span style={{ textAlign: "right", fontSize: 12, color: "#667085" }}>
                            {visits} visits
                            <br />
                            {lastVisit ? fmtDate(lastVisit) : "Never"}
                          </span>
                        </div>
                      );
                    })
                  )}
                </div>
              </div>

              <div style={{ background: "#fff", border: "1px solid #eaecf0", borderRadius: 8, padding: 14 }}>
                <strong>End of Day</strong>
                <div style={{ marginTop: 10, display: "grid", gap: 8, fontSize: 13 }}>
                  <div style={{ display: "flex", justifyContent: "space-between" }}>
                    <span>Total sales</span>
                    <strong>{money(daily.sales)}</strong>
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between" }}>
                    <span>Paid collected</span>
                    <strong>{money(daily.paid)}</strong>
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between" }}>
                    <span>Due remaining</span>
                    <strong>{money(daily.due)}</strong>
                  </div>
                  <div style={{ height: 1, background: "#eaecf0" }} />
                  <div style={{ display: "flex", justifyContent: "space-between" }}>
                    <span>Naulo Stick left</span>
                    <strong>{daily.remainingNauloStick}</strong>
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between" }}>
                    <span>Cig left</span>
                    <strong>{daily.remainingCig}</strong>
                  </div>
                </div>
              </div>
            </div>
          </section>
        </main>
      </div>
    </>
  );
}
