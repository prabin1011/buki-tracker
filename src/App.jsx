import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

const APP_NAME = "Buki Tracker";
const CURRENCY = "NPR";
const STORE_KEY = "buki-tracker-v1";
const SESSION_KEY = "buki-tracker-session";
const SHEETS_URL = import.meta.env.VITE_GOOGLE_SHEET_WEBAPP_URL || "";
const AUTH_USERNAME = "bookitracker@123";
const AUTH_PASSWORD = "arjun_buki_tracker";
const PAYMENT_METHODS = ["eSewa", "Online", "Hard cash"];
const PRODUCTS = [
  { key: "nauloStick", label: "Naulo Stick", unit: "stick", defaultPrice: 200 },
  { key: "cig", label: "Cig", unit: "cig", defaultPrice: 30 },
];
const CIGS_PER_PACKET = 20;

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
    cigPacketsBought: 0,
    openingNauloStick: 0,
    openingCigPieces: 0,
    notes: "",
  };
}

function cigOpeningPieces(stock) {
  return num(stock.openingCigPieces ?? stock.openingCig);
}

function cigBoughtPackets(stock) {
  return num(stock.cigPacketsBought ?? stock.cigPrepared);
}

function cigBoughtPieces(stock) {
  return cigBoughtPackets(stock) * CIGS_PER_PACKET;
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
        "Opening Cig Pieces",
        "Cig Packets Bought",
        "Cig Pieces Bought",
        "Sold Cig",
        "Opening Cig Pieces Remaining",
        "Bought Cig Pieces Remaining",
        "Total Cig Pieces Remaining",
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
        r.openingCigPieces,
        r.cigPacketsBought,
        r.cigPiecesBought,
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

function buildFilteredWorkbook(clients, transactions, stockRows, filters) {
  const matchesDate = (date) => {
    if (filters.from && date < filters.from) return false;
    if (filters.to && date > filters.to) return false;
    return true;
  };
  const matchesClient = (clientId, clientName = "") => {
    if (!filters.clientId) return true;
    return clientId === filters.clientId || normalizeName(clientName) === normalizeName(filters.clientId);
  };
  const filteredTransactions = transactions.filter(
    (t) =>
      matchesDate(t.date) &&
      matchesClient(t.clientId, t.clientName) &&
      (filters.payment === "All" || (filters.payment === "Paid" ? t.paid : !t.paid))
  );
  const filteredStockRows = stockRows.filter((row) => matchesDate(row.date));
  const clientIds = new Set(filteredTransactions.map((t) => t.clientId));
  const filteredClients = filters.recordType === "Daily Summary"
    ? []
    : clients.filter((c) => !filters.clientId || clientIds.has(c.id) || c.id === filters.clientId);

  if (filters.recordType === "Transactions") return buildWorkbook(filteredClients, filteredTransactions, []);
  if (filters.recordType === "Daily Summary") return buildWorkbook([], [], filteredStockRows);
  return buildWorkbook(filteredClients, filteredTransactions, filteredStockRows);
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

function loadGoogleSheetRecords(url) {
  return new Promise((resolve, reject) => {
    const callback = `bukiSheetCallback_${Date.now()}_${Math.random().toString(36).slice(2)}`;
    const script = document.createElement("script");
    const separator = url.includes("?") ? "&" : "?";
    const timer = window.setTimeout(() => {
      cleanup();
      reject(new Error("Google Sheet load timed out"));
    }, 15000);

    function cleanup() {
      window.clearTimeout(timer);
      delete window[callback];
      script.remove();
    }

    window[callback] = (data) => {
      cleanup();
      resolve(data.records || []);
    };

    script.onerror = () => {
      cleanup();
      reject(new Error("Could not load Google Sheet"));
    };

    script.src = `${url}${separator}callback=${callback}`;
    document.body.appendChild(script);
  });
}

function recordsToState(records) {
  const clientMap = new Map();
  const transactionMap = new Map();
  const stockMap = {};

  records.forEach((row) => {
    const type = row["Record Type"];
    if (type === "Client") {
      const id = row["Client ID"] || row.ID || uid();
      clientMap.set(id, {
        id,
        name: row.Client || "",
        phone: row.Phone || "",
        notes: row.Notes || "",
        joined: normalizeExcelDate(row.Date),
      });
    }
  });

  records.forEach((row) => {
    const type = row["Record Type"];
    if (type === "Transaction") {
      const clientId = row["Client ID"] || "";
      if (clientId && !clientMap.has(clientId) && row.Client) {
        clientMap.set(clientId, { id: clientId, name: row.Client, phone: "", notes: "", joined: normalizeExcelDate(row.Date) });
      }
      const id = row.ID || uid();
      transactionMap.set(id, {
        id,
        date: normalizeExcelDate(row.Date),
        clientId,
        clientName: row.Client || clientMap.get(clientId)?.name || "",
        nauloStick: num(row["Naulo Stick"]),
        cig: num(row.Cig),
        nauloStickPrice: num(row["Naulo Stick Price"]) || PRODUCTS[0].defaultPrice,
        cigPrice: num(row["Cig Price"]) || PRODUCTS[1].defaultPrice,
        amount:
          num(row.Amount) ||
          num(row["Naulo Stick"]) * (num(row["Naulo Stick Price"]) || PRODUCTS[0].defaultPrice) +
            num(row.Cig) * (num(row["Cig Price"]) || PRODUCTS[1].defaultPrice),
        paid: row["Payment Status"] !== "Due",
        method: row["Payment Method"] || "",
        notes: row.Notes || "",
      });
    }
    if (type === "Daily Summary") {
      const date = normalizeExcelDate(row.Date);
      stockMap[date] = {
        date,
        openingNauloStick: num(row["Opening Naulo Stick"]),
        nauloStickPrepared: num(row["Prepared Naulo Stick"]),
        openingCigPieces: num(row["Opening Cig Pieces"] || row["Opening Cig"]),
        cigPacketsBought:
          num(row["Cig Packets Bought"]) ||
          Math.ceil(num(row["Cig Pieces Bought"] || row["Prepared Cig"]) / CIGS_PER_PACKET),
        notes: row.Notes || "",
      };
    }
  });

  return {
    clients: Array.from(clientMap.values()).filter((c) => c.name),
    transactions: Array.from(transactionMap.values()).sort((a, b) => b.date.localeCompare(a.date)),
    stockByDate: stockMap,
  };
}

const inputStyle = {
  width: "100%",
  border: "1px solid #d5d7db",
  borderRadius: 8,
  padding: "10px 11px",
  fontSize: 14,
  fontFamily: "inherit",
  background: "#fff",
  color: "#101828",
};

const buttonStyle = {
  border: "none",
  borderRadius: 8,
  padding: "10px 14px",
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

function Panel({ children, style }) {
  return (
    <section style={{ background: "#fff", border: "1px solid #e4e7ec", borderRadius: 10, boxShadow: "0 1px 2px rgba(16,24,40,0.04)", ...style }}>
      {children}
    </section>
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

function LoginScreen({ onLogin }) {
  const [form, setForm] = useState({ username: "", password: "" });
  const [error, setError] = useState("");

  function submit(e) {
    e.preventDefault();
    if (form.username === AUTH_USERNAME && form.password === AUTH_PASSWORD) {
      sessionStorage.setItem(SESSION_KEY, "ok");
      onLogin();
      return;
    }
    setError("Invalid username or password.");
  }

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&display=swap');
        *{box-sizing:border-box}
        body{margin:0;background:#f4f6f8;color:#101828;font-family:Inter,Segoe UI,system-ui,sans-serif}
        input:focus{outline:2px solid #10182822;border-color:#101828}
      `}</style>
      <main style={{ minHeight: "100vh", display: "grid", placeItems: "center", padding: 18 }}>
        <form onSubmit={submit} style={{ width: 390, maxWidth: "100%", background: "#fff", border: "1px solid #e4e7ec", borderRadius: 12, padding: 22, boxShadow: "0 18px 50px rgba(16,24,40,0.08)" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 18 }}>
            <div style={{ width: 40, height: 40, borderRadius: 10, background: "#101828", color: "#fff", display: "grid", placeItems: "center", fontWeight: 900 }}>B</div>
            <div>
              <h1 style={{ margin: 0, fontSize: 22 }}>Buki Tracker</h1>
              <div style={{ fontSize: 13, color: "#667085" }}>Private access</div>
            </div>
          </div>
          <div style={{ display: "grid", gap: 12 }}>
            <Field label="Username">
              <input style={inputStyle} value={form.username} onChange={(e) => setForm((p) => ({ ...p, username: e.target.value }))} autoComplete="username" />
            </Field>
            <Field label="Password">
              <input style={inputStyle} type="password" value={form.password} onChange={(e) => setForm((p) => ({ ...p, password: e.target.value }))} autoComplete="current-password" />
            </Field>
            {error && <div style={{ background: "#fef3f2", border: "1px solid #fecdca", color: "#b42318", padding: "9px 10px", borderRadius: 8, fontSize: 13, fontWeight: 700 }}>{error}</div>}
            <button type="submit" style={{ ...buttonStyle, background: "#101828", color: "#fff", padding: 12 }}>
              Log In
            </button>
          </div>
        </form>
      </main>
    </>
  );
}

export default function App() {
  const initial = loadInitialData();
  const [authed, setAuthed] = useState(() => sessionStorage.getItem(SESSION_KEY) === "ok");
  const [clients, setClients] = useState(initial?.clients || []);
  const [transactions, setTransactions] = useState(initial?.transactions || []);
  const [stockByDate, setStockByDate] = useState(initial?.stockByDate || {});
  const [date, setDate] = useState(today());
  const [clientModal, setClientModal] = useState(false);
  const [txForm, setTxForm] = useState(defaultForm());
  const [clientForm, setClientForm] = useState({ name: "", phone: "", notes: "" });
  const [message, setMessage] = useState("");
  const [syncState, setSyncState] = useState(SHEETS_URL ? "Loading Google Sheet..." : "Local export mode");
  const [exportOpen, setExportOpen] = useState(false);
  const [exportFilters, setExportFilters] = useState({
    from: today(),
    to: today(),
    clientId: "",
    payment: "All",
    recordType: "All",
  });

  const stock = stockByDate[date] || emptyStock(date);

  function logout() {
    sessionStorage.removeItem(SESSION_KEY);
    setAuthed(false);
  }

  useEffect(() => {
    localStorage.setItem(STORE_KEY, JSON.stringify({ clients, transactions, stockByDate }));
  }, [clients, transactions, stockByDate]);

  async function refreshCloudRecords() {
    if (!SHEETS_URL) {
      setSyncState("Google Sheet URL missing");
      setMessage("Add VITE_GOOGLE_SHEET_WEBAPP_URL in Vercel production settings.");
      return;
    }
    setSyncState("Loading Google Sheet...");
    try {
      const records = await loadGoogleSheetRecords(SHEETS_URL);
      const loaded = recordsToState(records);
      setClients(loaded.clients);
      setTransactions(loaded.transactions);
      setStockByDate(loaded.stockByDate);
      setSyncState(`Google Sheet connected - ${records.length} rows loaded`);
    } catch {
      setSyncState("Google Sheet load failed - local copy active");
      setMessage("Google Sheet load failed. Redeploy Apps Script as a new Web App version and set access to Anyone.");
    }
  }

  useEffect(() => {
    if (!SHEETS_URL || !authed) return;
    let cancelled = false;
    loadGoogleSheetRecords(SHEETS_URL)
      .then((records) => {
        if (cancelled) return;
        const loaded = recordsToState(records);
        setClients(loaded.clients);
        setTransactions(loaded.transactions);
        setStockByDate(loaded.stockByDate);
        setSyncState(`Google Sheet connected - ${records.length} rows loaded`);
      })
      .catch(() => {
        if (!cancelled) {
          setSyncState("Google Sheet load failed - local copy active");
          setMessage("Google Sheet load failed. Redeploy Apps Script as a new Web App version and set access to Anyone.");
        }
      });
    return () => {
      cancelled = true;
    };
  }, [authed]);

  const dayTransactions = useMemo(() => transactions.filter((t) => t.date === date), [transactions, date]);

  const daily = useMemo(() => {
    const soldNauloStick = dayTransactions.reduce((sum, t) => sum + num(t.nauloStick), 0);
    const soldCig = dayTransactions.reduce((sum, t) => sum + num(t.cig), 0);
    const sales = dayTransactions.reduce((sum, t) => sum + num(t.amount), 0);
    const paid = dayTransactions.filter((t) => t.paid).reduce((sum, t) => sum + num(t.amount), 0);
    const due = sales - paid;
    const openingNauloStickUsed = Math.min(num(stock.openingNauloStick), soldNauloStick);
    const preparedNauloStickUsed = Math.max(0, soldNauloStick - openingNauloStickUsed);
    const openingCigUsed = Math.min(cigOpeningPieces(stock), soldCig);
    const preparedCigUsed = Math.max(0, soldCig - openingCigUsed);
    const openingNauloStickRemaining = num(stock.openingNauloStick) - openingNauloStickUsed;
    const preparedNauloStickRemaining = num(stock.nauloStickPrepared) - preparedNauloStickUsed;
    const openingCigRemaining = cigOpeningPieces(stock) - openingCigUsed;
    const preparedCigRemaining = cigBoughtPieces(stock) - preparedCigUsed;
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
      cigPiecesBought: cigBoughtPieces(stock),
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
      const openingCigUsed = Math.min(cigOpeningPieces(row), soldCig);
      const preparedCigUsed = Math.max(0, soldCig - openingCigUsed);
      const openingNauloStickRemaining = num(row.openingNauloStick) - openingNauloStickUsed;
      const preparedNauloStickRemaining = num(row.nauloStickPrepared) - preparedNauloStickUsed;
      const openingCigRemaining = cigOpeningPieces(row) - openingCigUsed;
      const preparedCigRemaining = cigBoughtPieces(row) - preparedCigUsed;
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
        openingCigPieces: cigOpeningPieces(row),
        cigPacketsBought: cigBoughtPackets(row),
        cigPiecesBought: cigBoughtPieces(row),
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
      setSyncState(result.skipped ? "Local export mode" : "Client synced to Google Sheet");
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
      setSyncState(result.skipped ? "Local export mode" : "Transaction synced to Google Sheet");
    } catch {
      setSyncState("Google Sheets sync failed");
    }
  }

  function togglePaid(id) {
    let updatedTransaction = null;
    setTransactions((prev) =>
      prev.map((t) => {
        if (t.id !== id) return t;
        updatedTransaction = { ...t, paid: !t.paid, method: !t.paid ? PAYMENT_METHODS[0] : "" };
        return updatedTransaction;
      })
    );
    if (updatedTransaction) {
      syncToGoogleSheet({ type: "transaction", transaction: updatedTransaction })
        .then((result) => setSyncState(result.skipped ? "Local export mode" : "Payment status synced to Google Sheet"))
        .catch(() => setSyncState("Google Sheet sync failed"));
    }
  }

  function currentDailySummary() {
    return {
      date,
      ...stock,
      openingCigPieces: cigOpeningPieces(stock),
      cigPacketsBought: cigBoughtPackets(stock),
      cigPiecesBought: cigBoughtPieces(stock),
      ...daily,
    };
  }

  async function saveDayStock() {
    setMessage(`Saving stock for ${fmtDate(date)}...`);
    try {
      const result = await syncToGoogleSheet({
        type: "daily_summary",
        summary: currentDailySummary(),
      });
      setSyncState(result.skipped ? "Local export mode" : "Stock day saved to Google Sheet");
      setMessage(`Stock for ${fmtDate(date)} saved.`);
    } catch {
      setSyncState("Google Sheet sync failed");
      setMessage("Could not save day stock. Check Apps Script deployment and access.");
    }
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
        openingCigPieces: cigOpeningPieces(prev[nextDate] || {}) + Math.max(0, daily.remainingCig),
      },
    }));
    setDate(nextDate);
    setMessage(`Remaining stock carried to ${fmtDate(nextDate)}.`);
    try {
      const result = await syncToGoogleSheet({
        type: "daily_summary",
        summary: currentDailySummary(),
      });
      setSyncState(result.skipped ? "Local export mode" : "Daily summary synced to Google Sheet");
    } catch {
      setSyncState("Google Sheets sync failed");
    }
  }

  async function saveExcel() {
    if (!SHEETS_URL) {
      setMessage("Google Sheet URL is not configured. Use Export Excel to download a file.");
      return;
    }
    setSyncState("Saving to Google Sheet...");
    try {
      for (const client of clients) {
        await syncToGoogleSheet({ type: "client", client });
      }
      for (const transaction of transactions) {
        await syncToGoogleSheet({ type: "transaction", transaction });
      }
      for (const summary of allStockRows) {
        await syncToGoogleSheet({ type: "daily_summary", summary });
      }
      setSyncState("Google Sheet saved");
      setMessage("Saved to the same Google Sheet.");
    } catch {
      setSyncState("Google Sheet save failed");
      setMessage("Could not save to Google Sheet. Export Excel still works.");
    }
  }

  function exportExcel() {
    const fileName = `buki_tracker_export_${exportFilters.from || "all"}_to_${exportFilters.to || "all"}.xlsx`;
    XLSX.writeFile(buildFilteredWorkbook(clients, transactions, allStockRows, exportFilters), fileName);
    setMessage(`Exported ${fileName}.`);
  }

  if (!authed) return <LoginScreen onLogin={() => setAuthed(true)} />;

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&display=swap');
        *{box-sizing:border-box}
        body{margin:0;background:#f4f6f8;color:#101828;font-family:Inter,Segoe UI,system-ui,sans-serif}
        button:hover{opacity:.86}
        input:focus,select:focus,textarea:focus{outline:2px solid #10182822;border-color:#101828}
        table{border-collapse:collapse;width:100%}
        th{text-align:left;font-size:11px;text-transform:uppercase;color:#667085;padding:10px;border-bottom:1px solid #eaecf0}
        td{padding:11px 10px;border-bottom:1px solid #f0f2f5;font-size:13px;vertical-align:top}
        .wide{overflow-x:auto}
        @media(max-width:850px){
          body{background:#f6f7f9}
          input,select,textarea{font-size:16px!important}
          .layout{grid-template-columns:1fr!important}
          .stats{grid-template-columns:repeat(2,minmax(0,1fr))!important}
          .formgrid{grid-template-columns:1fr!important}
          header{align-items:flex-start!important}
          header > div:last-child{width:100%;display:grid!important;grid-template-columns:1fr 1fr}
        }
        @media(max-width:430px){
          .stats{gap:8px!important}
        }
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

      {exportOpen && (
        <Modal title="Export Excel" onClose={() => setExportOpen(false)}>
          <div style={{ display: "grid", gap: 12 }}>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }} className="formgrid">
              <Field label="Export from">
                <input style={inputStyle} type="date" value={exportFilters.from} onChange={(e) => setExportFilters((p) => ({ ...p, from: e.target.value }))} />
              </Field>
              <Field label="Export to">
                <input style={inputStyle} type="date" value={exportFilters.to} onChange={(e) => setExportFilters((p) => ({ ...p, to: e.target.value }))} />
              </Field>
            </div>
            <Field label="Client">
              <select style={inputStyle} value={exportFilters.clientId} onChange={(e) => setExportFilters((p) => ({ ...p, clientId: e.target.value }))}>
                <option value="">All clients</option>
                {clients.map((c) => (
                  <option key={c.id} value={c.id}>
                    {c.name}
                  </option>
                ))}
              </select>
            </Field>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }} className="formgrid">
              <Field label="Payment">
                <select style={inputStyle} value={exportFilters.payment} onChange={(e) => setExportFilters((p) => ({ ...p, payment: e.target.value }))}>
                  <option>All</option>
                  <option>Paid</option>
                  <option>Due</option>
                </select>
              </Field>
              <Field label="Record type">
                <select style={inputStyle} value={exportFilters.recordType} onChange={(e) => setExportFilters((p) => ({ ...p, recordType: e.target.value }))}>
                  <option>All</option>
                  <option>Transactions</option>
                  <option>Daily Summary</option>
                </select>
              </Field>
            </div>
            <button onClick={exportExcel} style={{ ...buttonStyle, background: "#101828", color: "#fff", padding: 12 }}>
              Download Excel
            </button>
          </div>
        </Modal>
      )}

      <div style={{ maxWidth: 1240, margin: "0 auto", padding: "20px 18px 28px" }}>
        <header style={{ display: "flex", justifyContent: "space-between", gap: 16, alignItems: "center", marginBottom: 16, flexWrap: "wrap" }}>
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <div style={{ width: 40, height: 40, borderRadius: 10, background: "#101828", color: "#fff", display: "grid", placeItems: "center", fontWeight: 900 }}>B</div>
              <div>
                <h1 style={{ margin: 0, fontSize: 24, letterSpacing: 0 }}>{APP_NAME}</h1>
                <div style={{ fontSize: 13, color: "#667085" }}>Daily client, stock, due, and payment tracker</div>
                <div style={{ fontSize: 12, color: SHEETS_URL ? "#067647" : "#667085", marginTop: 2 }}>{syncState}</div>
              </div>
            </div>
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button onClick={saveExcel} style={{ ...buttonStyle, background: "#101828", color: "#fff" }}>
              Save Excel
            </button>
            <button onClick={() => setExportOpen(true)} style={{ ...buttonStyle, background: "#fff", color: "#101828", border: "1px solid #d5d7db" }}>
              Export Excel
            </button>
            <button onClick={refreshCloudRecords} style={{ ...buttonStyle, background: "#fff", color: "#067647", border: "1px solid #abefc6" }}>
              Refresh Sheet
            </button>
            <button onClick={logout} style={{ ...buttonStyle, background: "#fff", color: "#b42318", border: "1px solid #fecdca" }}>
              Log Out
            </button>
          </div>
        </header>

        {message && (
          <div style={{ background: "#ecfdf3", border: "1px solid #abefc6", color: "#067647", padding: "9px 12px", borderRadius: 8, marginBottom: 12, fontWeight: 700, fontSize: 13 }}>
            {message}
          </div>
        )}

        <Panel style={{ padding: 14, marginBottom: 12 }}>
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
                <NumberInput style={inputStyle} value={cigOpeningPieces(stock)} onChange={(value) => updateStock("openingCigPieces", value)} />
              </Field>
              <Field label="Cig packets bought">
                <NumberInput style={inputStyle} value={cigBoughtPackets(stock)} onChange={(value) => updateStock("cigPacketsBought", value)} />
              </Field>
            </div>
            <div style={{ display: "grid", gap: 8 }}>
              <button onClick={saveDayStock} style={{ ...buttonStyle, background: "#101828", color: "#fff" }}>
                Save Day Stock
              </button>
              <button onClick={carryForward} style={{ ...buttonStyle, background: "#eef4ff", color: "#3538cd", border: "1px solid #c7d7fe" }}>
                Carry Remaining
              </button>
            </div>
          </div>
          <div style={{ marginTop: 10, fontSize: 12, color: "#667085" }}>
            Opening stock is the inventory available for {fmtDate(date)}. Only remaining items carry into the next day.
          </div>
        </Panel>

        <section className="stats" style={{ display: "grid", gridTemplateColumns: "repeat(6,1fr)", gap: 10, marginBottom: 12 }}>
          <Stat label="Daily stick sold" value={daily.soldNauloStick} />
          <Stat label="Daily cig sold" value={daily.soldCig} />
          <Stat label="Stick carry total" value={daily.remainingNauloStick} tone={daily.remainingNauloStick < 0 ? "bad" : "dark"} />
          <Stat label="Cig pieces left" value={daily.remainingCig} tone={daily.remainingCig < 0 ? "bad" : "dark"} />
          <Stat label="Paid" value={money(daily.paid)} tone="good" />
          <Stat label="Due" value={money(daily.due)} tone={daily.due > 0 ? "bad" : "dark"} />
        </section>

        <Panel style={{ padding: 14, marginBottom: 12 }}>
          <div style={{ display: "flex", justifyContent: "space-between", gap: 12, alignItems: "center", flexWrap: "wrap", marginBottom: 10 }}>
            <strong>Remaining Stock</strong>
            <span style={{ fontSize: 12, color: "#667085" }}>Daily sales for {fmtDate(date)} consume opening stock first, then bought/prepared stock.</span>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 10 }} className="stats">
            <Stat label="Opening stick left" value={daily.openingNauloStickRemaining} tone={daily.openingNauloStickRemaining < 0 ? "bad" : "dark"} />
            <Stat label="Prepared stick left" value={daily.preparedNauloStickRemaining} tone={daily.preparedNauloStickRemaining < 0 ? "bad" : "dark"} />
            <Stat label="Opening cig pieces left" value={daily.openingCigRemaining} tone={daily.openingCigRemaining < 0 ? "bad" : "dark"} />
            <Stat label="Bought cig pieces left" value={daily.preparedCigRemaining} tone={daily.preparedCigRemaining < 0 ? "bad" : "dark"} />
          </div>
        </Panel>

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
                            {t.cig} Cig pieces
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
                    <span>Cig pieces left</span>
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
