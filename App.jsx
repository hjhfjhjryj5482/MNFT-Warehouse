import React, { useState, useEffect, useMemo, useCallback } from "react";
import {
  LayoutDashboard, Package, ArrowDownCircle, ArrowUpCircle, BookOpen, Truck,
  Users, Wallet, Building2, FileBarChart2, CalendarClock, Settings, Search,
  Plus, X, AlertTriangle, TrendingUp, TrendingDown, Snowflake, ChevronRight,
  Lock, Unlock, Printer, Download, CheckCircle2, XCircle, Trash2, Pencil,
  Filter, BarChart3, ChevronDown, ClipboardList, ShieldCheck, Tags, Ruler,
  Building, LogOut
} from "lucide-react";
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell, XAxis, YAxis,
  CartesianGrid, Tooltip, ResponsiveContainer, Legend
} from "recharts";
import * as XLSX from "xlsx";

/* ---------------------------------------------------------------------- */
/*  Constants & helpers                                                    */
/* ---------------------------------------------------------------------- */

const DEFAULT_CATEGORIES = ["Frozen Items", "Bakery Items", "Fast Food Raw Material", "Beverages", "Packaging", "Other"];
const DEFAULT_UNITS = ["KG", "Piece", "Pack", "Carton", "Liter", "Box", "Dozen"];
const PAY_METHODS = ["Cash", "Bank Transfer", "Cheque", "UPI/Card"];
const ROLES = ["Super Admin", "Warehouse Manager", "Accountant", "Viewer"];
const PALETTE = ["#2E6F95", "#C1832E", "#B24A3A", "#3E8266", "#7A6C9E", "#6B7480", "#0E7490", "#9D5C0D", "#5B21B6", "#B45309"];

const uid = () => (crypto.randomUUID ? crypto.randomUUID() : "id-" + Date.now() + "-" + Math.random().toString(16).slice(2));
const todayStr = () => new Date().toISOString().slice(0, 10);
const nowIso = () => new Date().toISOString();
const fmtMoney = (n) => "$" + (Number(n) || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtQty = (n) => (Number(n) || 0).toLocaleString("en-US", { maximumFractionDigits: 2 });
const fmtDate = (d) => { if (!d) return "-"; const dt = new Date(d + "T00:00:00"); return dt.toLocaleDateString("en-US", { day: "2-digit", month: "short", year: "numeric" }); };
const fmtDateTime = (iso) => { if (!iso) return "-"; const dt = new Date(iso); return dt.toLocaleString("en-US", { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" }); };
const addDays = (dateStr, n) => { const d = new Date(dateStr + "T00:00:00"); d.setDate(d.getDate() + n); return d.toISOString().slice(0, 10); };
const cx = (...a) => a.filter(Boolean).join(" ");
const hashIdx = (s, mod) => { let h = 0; for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) >>> 0; return h % mod; };
const catColor = (name) => PALETTE[hashIdx(name || "", PALETTE.length)];

function pushLog(data, user, action, detail) {
  const entry = { id: uid(), ts: nowIso(), user: user || "Unknown", action, detail };
  return [entry, ...(data.activityLog || [])].slice(0, 300);
}

function exportRowsToExcel(filename, sheetName, rows) {
  if (!rows || rows.length === 0) return;
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0, 31));
  XLSX.writeFile(wb, filename.endsWith(".xlsx") ? filename : filename + ".xlsx");
}

/* ---------------------------------------------------------------------- */
/*  Seed data                                                               */
/* ---------------------------------------------------------------------- */

function warehouseSeed(kind) {
  if (kind === "bakery-hub") {
    const suppliers = [
      { id: "sup-b1", name: "Golden Wheat Bakery Supply", contact: "+1 555-0198", address: "44 Millhouse Ave" },
      { id: "sup-b2", name: "Dairy Fresh Distributors", contact: "+1 555-0221", address: "7 Pasture Lane" },
    ];
    const branches = [
      { id: "brb-1", name: "Downtown Bakery Cafe" },
      { id: "brb-2", name: "Bakery Kiosk - Mall" },
    ];
    const items = [
      { id: "itb-1", name: "Refined Flour (Maida)", category: "Bakery Items", subCategory: "Flour", unit: "KG", minStock: 100, purchaseRate: 0.65, openingStock: 180, supplier: "sup-b1", status: "Active", notes: "" },
      { id: "itb-2", name: "Butter Block", category: "Bakery Items", subCategory: "Dairy", unit: "KG", minStock: 30, purchaseRate: 5.5, openingStock: 22, supplier: "sup-b2", status: "Active", notes: "" },
      { id: "itb-3", name: "Cake Boxes (Medium)", category: "Packaging", subCategory: "Boxes", unit: "Carton", minStock: 10, purchaseRate: 15.0, openingStock: 6, supplier: "sup-b1", status: "Active", notes: "" },
    ];
    const t = todayStr();
    const stockIns = [
      { id: uid(), date: addDays(t, -3), invoice: "BK-INV-201", supplier: "sup-b1", itemId: "itb-1", qty: 120, unit: "KG", rate: 0.65, paid: 60, method: "Cash", remarks: "", createdBy: "System", createdAt: nowIso() },
      { id: uid(), date: addDays(t, -1), invoice: "BK-INV-202", supplier: "sup-b2", itemId: "itb-2", qty: 20, unit: "KG", rate: 5.5, paid: 90, method: "Bank Transfer", remarks: "", createdBy: "System", createdAt: nowIso() },
    ];
    const stockOuts = [
      { id: uid(), date: addDays(t, -2), destination: "brb-1", itemId: "itb-1", qty: 40, unit: "KG", rate: 0.65, receiver: "N. Farooq", remarks: "", createdBy: "System", createdAt: nowIso() },
      { id: uid(), date: t, destination: "brb-2", itemId: "itb-3", qty: 4, unit: "Carton", rate: 15.0, receiver: "H. Iqbal", remarks: "", createdBy: "System", createdAt: nowIso() },
    ];
    return {
      items, suppliers, branches, stockIns, stockOuts, payments: [], closings: {},
      categories: [...DEFAULT_CATEGORIES], units: [...DEFAULT_UNITS],
      activityLog: [{ id: uid(), ts: nowIso(), user: "System", action: "Setup", detail: "Warehouse initialized with sample data." }],
      settings: { companyName: "ColdLine Foods Pvt. Ltd." },
    };
  }
  // default: cold storage warehouse
  const suppliers = [
    { id: "sup-1", name: "Nordic Cold Chain Ltd.", contact: "+1 555-0134", address: "12 Frostgate Rd, Industrial Park" },
    { id: "sup-2", name: "Golden Wheat Bakery Supply", contact: "+1 555-0198", address: "44 Millhouse Ave" },
    { id: "sup-3", name: "QuickServe Raw Materials Co.", contact: "+1 555-0177", address: "9 Harbor Yard" },
  ];
  const branches = [
    { id: "br-1", name: "Main Restaurant" },
    { id: "br-2", name: "Branch 1 - Downtown" },
    { id: "br-3", name: "Bakery Unit" },
    { id: "br-4", name: "Fast Food Unit" },
  ];
  const items = [
    { id: "it-1", name: "Frozen Chicken Breast", category: "Frozen Items", subCategory: "Poultry", unit: "KG", minStock: 100, purchaseRate: 4.2, openingStock: 240, supplier: "sup-1", status: "Active", notes: "" },
    { id: "it-2", name: "Frozen French Fries", category: "Frozen Items", subCategory: "Potato", unit: "KG", minStock: 80, purchaseRate: 1.8, openingStock: 60, supplier: "sup-1", status: "Active", notes: "" },
    { id: "it-3", name: "Refined Flour (Maida)", category: "Bakery Items", subCategory: "Flour", unit: "KG", minStock: 150, purchaseRate: 0.65, openingStock: 320, supplier: "sup-2", status: "Active", notes: "" },
    { id: "it-4", name: "Butter Block", category: "Bakery Items", subCategory: "Dairy", unit: "KG", minStock: 40, purchaseRate: 5.5, openingStock: 18, supplier: "sup-2", status: "Active", notes: "" },
    { id: "it-5", name: "Burger Buns", category: "Fast Food Raw Material", subCategory: "Bread", unit: "Pack", minStock: 100, purchaseRate: 2.1, openingStock: 210, supplier: "sup-3", status: "Active", notes: "" },
    { id: "it-6", name: "Cheddar Cheese Slices", category: "Fast Food Raw Material", subCategory: "Dairy", unit: "Pack", minStock: 60, purchaseRate: 3.4, openingStock: 40, supplier: "sup-3", status: "Active", notes: "" },
    { id: "it-7", name: "Cola Syrup Concentrate", category: "Beverages", subCategory: "Syrup", unit: "Liter", minStock: 20, purchaseRate: 12.0, openingStock: 55, supplier: "sup-3", status: "Active", notes: "" },
    { id: "it-8", name: "Takeaway Boxes (Large)", category: "Packaging", subCategory: "Boxes", unit: "Carton", minStock: 15, purchaseRate: 22.0, openingStock: 8, supplier: "sup-3", status: "Active", notes: "" },
  ];
  const t = todayStr();
  const stockIns = [
    { id: uid(), date: addDays(t, -6), invoice: "INV-1042", supplier: "sup-1", itemId: "it-1", qty: 150, unit: "KG", rate: 4.2, paid: 500, method: "Bank Transfer", remarks: "", createdBy: "System", createdAt: nowIso() },
    { id: uid(), date: addDays(t, -5), invoice: "INV-1043", supplier: "sup-2", itemId: "it-3", qty: 200, unit: "KG", rate: 0.65, paid: 130, method: "Cash", remarks: "", createdBy: "System", createdAt: nowIso() },
    { id: uid(), date: addDays(t, -4), invoice: "INV-1044", supplier: "sup-3", itemId: "it-5", qty: 300, unit: "Pack", rate: 2.1, paid: 630, method: "UPI/Card", remarks: "", createdBy: "System", createdAt: nowIso() },
    { id: uid(), date: addDays(t, -3), invoice: "INV-1045", supplier: "sup-1", itemId: "it-2", qty: 100, unit: "KG", rate: 1.8, paid: 90, method: "Cheque", remarks: "", createdBy: "System", createdAt: nowIso() },
    { id: uid(), date: addDays(t, -2), invoice: "INV-1046", supplier: "sup-3", itemId: "it-7", qty: 40, unit: "Liter", rate: 12.0, paid: 480, method: "Bank Transfer", remarks: "", createdBy: "System", createdAt: nowIso() },
    { id: uid(), date: addDays(t, -1), invoice: "INV-1047", supplier: "sup-2", itemId: "it-4", qty: 30, unit: "KG", rate: 5.5, paid: 100, method: "Cash", remarks: "", createdBy: "System", createdAt: nowIso() },
    { id: uid(), date: t, invoice: "INV-1048", supplier: "sup-1", itemId: "it-1", qty: 80, unit: "KG", rate: 4.2, paid: 200, method: "Bank Transfer", remarks: "Morning delivery", createdBy: "System", createdAt: nowIso() },
  ];
  const stockOuts = [
    { id: uid(), date: addDays(t, -6), destination: "br-1", itemId: "it-1", qty: 40, unit: "KG", rate: 4.2, receiver: "M. Alvarez", remarks: "", createdBy: "System", createdAt: nowIso() },
    { id: uid(), date: addDays(t, -5), destination: "br-4", itemId: "it-5", qty: 90, unit: "Pack", rate: 2.1, receiver: "J. Chen", remarks: "", createdBy: "System", createdAt: nowIso() },
    { id: uid(), date: addDays(t, -4), destination: "br-3", itemId: "it-3", qty: 70, unit: "KG", rate: 0.65, receiver: "P. Novak", remarks: "", createdBy: "System", createdAt: nowIso() },
    { id: uid(), date: addDays(t, -3), destination: "br-2", itemId: "it-2", qty: 35, unit: "KG", rate: 1.8, receiver: "S. Rao", remarks: "", createdBy: "System", createdAt: nowIso() },
    { id: uid(), date: addDays(t, -2), destination: "br-4", itemId: "it-7", qty: 15, unit: "Liter", rate: 12.0, receiver: "J. Chen", remarks: "", createdBy: "System", createdAt: nowIso() },
    { id: uid(), date: addDays(t, -1), destination: "br-1", itemId: "it-6", qty: 20, unit: "Pack", rate: 3.4, receiver: "M. Alvarez", remarks: "", createdBy: "System", createdAt: nowIso() },
    { id: uid(), date: t, destination: "br-3", itemId: "it-4", qty: 8, unit: "KG", rate: 5.5, receiver: "P. Novak", remarks: "Early batch", createdBy: "System", createdAt: nowIso() },
  ];
  return {
    items, suppliers, branches, stockIns, stockOuts, payments: [], closings: {},
    categories: [...DEFAULT_CATEGORIES], units: [...DEFAULT_UNITS],
    activityLog: [{ id: uid(), ts: nowIso(), user: "System", action: "Setup", detail: "Warehouse initialized with sample data." }],
    settings: { companyName: "ColdLine Foods Pvt. Ltd." },
  };
}

function emptyWarehouseData() {
  return {
    items: [], suppliers: [], branches: [], stockIns: [], stockOuts: [], payments: [], closings: {},
    categories: [...DEFAULT_CATEGORIES], units: [...DEFAULT_UNITS],
    activityLog: [{ id: uid(), ts: nowIso(), user: "System", action: "Setup", detail: "New warehouse created." }],
    settings: { companyName: "" },
  };
}

function seedRoot() {
  const warehouses = {
    "wh-1": { name: "Central Cold Storage Warehouse", owner: "Ayesha Khan", address: "Industrial Zone, Block C", ...warehouseSeed("cold-storage") },
    "wh-2": { name: "North Bakery Hub", owner: "Bilal Ahmed", address: "Millhouse Ave, Sector 4", ...warehouseSeed("bakery-hub") },
  };
  const users = [
    { id: "u-0", name: "Platform Owner", username: "owner", password: "owner123", role: "Super Admin", warehouseIds: [] },
    { id: "u-1", name: "Ayesha Khan", username: "ayesha", password: "ayesha123", role: "Super Admin", warehouseIds: ["wh-1"] },
    { id: "u-2", name: "Bilal Ahmed", username: "bilal", password: "bilal123", role: "Super Admin", warehouseIds: ["wh-2"] },
    { id: "u-3", name: "Sana Tariq", username: "sana", password: "sana123", role: "Accountant", warehouseIds: ["wh-1"] },
    { id: "u-4", name: "Guest Viewer", username: "guest", password: "guest123", role: "Viewer", warehouseIds: ["wh-1", "wh-2"] },
  ];
  return { warehouses, users };
}

/* ---------------------------------------------------------------------- */
/*  Small UI atoms                                                          */
/* ---------------------------------------------------------------------- */

function Field({ label, children, hint }) {
  return (
    <label className="block mb-3">
      <span className="block text-[11px] font-semibold tracking-wide uppercase text-slate-500 mb-1">{label}</span>
      {children}
      {hint && <span className="block text-[11px] text-slate-400 mt-1">{hint}</span>}
    </label>
  );
}

const inputCls = "w-full rounded-md border border-slate-300 bg-white px-3 py-2 text-sm text-slate-800 focus:outline-none focus:ring-2 focus:ring-[#2E6F95] focus:border-transparent placeholder:text-slate-400";

function Modal({ title, onClose, children, wide }) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4" onClick={onClose}>
      <div
        className={cx("bg-white rounded-lg shadow-2xl w-full max-h-[88vh] overflow-y-auto", wide ? "max-w-2xl" : "max-w-md")}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between px-5 py-4 border-b border-slate-200 sticky top-0 bg-white z-10">
          <h3 className="font-semibold text-slate-800 tracking-wide" style={{ fontFamily: "'Oswald', sans-serif" }}>{title}</h3>
          <button onClick={onClose} className="p-1 rounded hover:bg-slate-100 text-slate-500"><X size={18} /></button>
        </div>
        <div className="p-5">{children}</div>
      </div>
    </div>
  );
}

function Badge({ tone = "slate", children }) {
  const tones = {
    slate: "bg-slate-100 text-slate-600",
    green: "bg-emerald-50 text-emerald-700 border border-emerald-200",
    amber: "bg-amber-50 text-amber-700 border border-amber-200",
    red: "bg-red-50 text-red-700 border border-red-200",
    blue: "bg-sky-50 text-sky-700 border border-sky-200",
  };
  return <span className={cx("inline-flex items-center gap-1 px-2 py-0.5 rounded text-[11px] font-medium", tones[tone])}>{children}</span>;
}

function StatCard({ icon: Icon, label, value, sub, accent }) {
  return (
    <div className="bg-white rounded-lg border border-slate-200 p-4 relative overflow-hidden">
      <div className="absolute left-0 top-0 bottom-0 w-1" style={{ background: accent }} />
      <div className="flex items-start justify-between pl-2">
        <div>
          <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-400 mb-1">{label}</div>
          <div className="text-xl font-bold text-slate-800 tabular-nums" style={{ fontFamily: "'JetBrains Mono', monospace" }}>{value}</div>
          {sub && <div className="text-[11px] text-slate-400 mt-1">{sub}</div>}
        </div>
        <div className="w-8 h-8 rounded-md flex items-center justify-center" style={{ background: accent + "1a" }}>
          <Icon size={16} color={accent} />
        </div>
      </div>
    </div>
  );
}

function Toolbar({ children }) {
  return <div className="flex flex-wrap items-center gap-2 mb-4">{children}</div>;
}

function SearchInput({ value, onChange, placeholder }) {
  return (
    <div className="relative">
      <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
      <input value={value} onChange={(e) => onChange(e.target.value)} placeholder={placeholder || "Search..."}
        className="pl-8 pr-3 py-1.5 rounded-md border border-slate-300 text-sm w-48 focus:outline-none focus:ring-2 focus:ring-[#2E6F95]" />
    </div>
  );
}

function SelectFilter({ value, onChange, options, placeholder }) {
  return (
    <select value={value} onChange={(e) => onChange(e.target.value)}
      className="px-2.5 py-1.5 rounded-md border border-slate-300 text-sm text-slate-600 bg-white focus:outline-none focus:ring-2 focus:ring-[#2E6F95]">
      <option value="">{placeholder}</option>
      {options.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
    </select>
  );
}

function ExportButton({ onClick, label = "Export Excel" }) {
  return (
    <button onClick={onClick} className="flex items-center gap-1.5 text-sm border border-slate-300 rounded-md px-3 py-1.5 text-slate-600 hover:bg-slate-50">
      <Download size={14} /> {label}
    </button>
  );
}

function EmptyState({ icon: Icon = Package, title, sub }) {
  return (
    <div className="flex flex-col items-center justify-center py-14 text-center">
      <div className="w-12 h-12 rounded-full bg-slate-100 flex items-center justify-center mb-3"><Icon size={20} className="text-slate-400" /></div>
      <div className="text-sm font-semibold text-slate-600">{title}</div>
      {sub && <div className="text-xs text-slate-400 mt-1 max-w-xs">{sub}</div>}
    </div>
  );
}

function Toast({ toasts }) {
  return (
    <div className="fixed bottom-4 right-4 z-[100] flex flex-col gap-2 items-end">
      {toasts.map((t) => (
        <div key={t.id} className={cx("px-4 py-2.5 rounded-md shadow-lg text-sm font-medium flex items-center gap-2 animate-[fadein_.2s_ease]",
          t.tone === "error" ? "bg-red-600 text-white" : t.tone === "warn" ? "bg-amber-500 text-white" : "bg-slate-900 text-white")}>
          {t.tone === "error" ? <XCircle size={15} /> : <CheckCircle2 size={15} />} {t.text}
        </div>
      ))}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/*  Main App                                                                */
/* ---------------------------------------------------------------------- */

const NAV = [
  { key: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { section: "Inventory" },
  { key: "items", label: "Items", icon: Package },
  { key: "stockin", label: "Stock In", icon: ArrowDownCircle },
  { key: "stockout", label: "Stock Out", icon: ArrowUpCircle },
  { key: "ledger", label: "Item Ledger", icon: BookOpen },
  { section: "Suppliers" },
  { key: "suppliers", label: "Suppliers", icon: Users },
  { key: "payments", label: "Payments", icon: Wallet },
  { section: "Branches" },
  { key: "branches", label: "Restaurants / Branches", icon: Building2 },
  { section: "Reports" },
  { key: "dailyreport", label: "Daily Report", icon: FileBarChart2 },
  { key: "monthlyreport", label: "Monthly Summary", icon: BarChart3 },
  { section: "Warehouse" },
  { key: "closing", label: "Daily Closing", icon: CalendarClock },
  { section: "Admin" },
  { key: "settings", label: "Settings", icon: Settings },
  { key: "warehouses", label: "All Warehouses", icon: Building, platformOnly: true },
];

const SESSION_KEY = "wms-session-user";

export default function App() {
  const [root, setRoot] = useState(null);
  const [loaded, setLoaded] = useState(false);
  const [tab, setTab] = useState("dashboard");
  const [loggedInUserId, setLoggedInUserId] = useState(null);
  const [currentWarehouseId, setCurrentWarehouseId] = useState(null);
  const [toasts, setToasts] = useState([]);
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [loginError, setLoginError] = useState("");

  const pushToast = useCallback((text, tone = "ok") => {
    const id = uid();
    setToasts((t) => [...t, { id, text, tone }]);
    setTimeout(() => setToasts((t) => t.filter((x) => x.id !== id)), 2800);
  }, []);

  useEffect(() => {
    try {
      const raw = localStorage.getItem("wms-root-data");
      if (raw) setRoot(JSON.parse(raw));
      else setRoot(seedRoot());
    } catch (e) {
      setRoot(seedRoot());
    } finally {
      setLoaded(true);
    }
    try {
      const savedSession = window.localStorage ? window.localStorage.getItem(SESSION_KEY) : null;
      if (savedSession) setLoggedInUserId(savedSession);
    } catch (e) { /* ignore */ }
  }, []);

  const persistRoot = useCallback((next) => {
    setRoot(next);
    try { localStorage.setItem("wms-root-data", JSON.stringify(next)); }
    catch (e) { pushToast("Could not save — changes are kept for this session only.", "warn"); }
  }, [pushToast]);

  const loggedInUser = root?.users?.find((u) => u.id === loggedInUserId) || null;
  const accessibleIds = loggedInUser ? (loggedInUser.warehouseIds.length ? loggedInUser.warehouseIds : Object.keys(root.warehouses)) : [];

  useEffect(() => {
    if (loggedInUser && accessibleIds.length && (!currentWarehouseId || !accessibleIds.includes(currentWarehouseId))) {
      setCurrentWarehouseId(accessibleIds[0]);
    }
  }, [loggedInUser, currentWarehouseId, accessibleIds]);

  const handleLogin = (username, password) => {
    const u = root.users.find((x) => x.username.toLowerCase() === username.trim().toLowerCase() && x.password === password);
    if (!u) { setLoginError("Invalid username or password."); return; }
    setLoginError("");
    setLoggedInUserId(u.id);
    try { window.localStorage && window.localStorage.setItem(SESSION_KEY, u.id); } catch (e) { /* ignore */ }
  };

  const handleLogout = () => {
    setLoggedInUserId(null);
    setCurrentWarehouseId(null);
    setTab("dashboard");
    try { window.localStorage && window.localStorage.removeItem(SESSION_KEY); } catch (e) { /* ignore */ }
  };

  const setWarehouseData = useCallback((newWhData) => {
    if (!currentWarehouseId || !root) return;
    persistRoot({ ...root, warehouses: { ...root.warehouses, [currentWarehouseId]: { ...root.warehouses[currentWarehouseId], ...newWhData } } });
  }, [root, currentWarehouseId, persistRoot]);

  if (!loaded || !root) {
    return (
      <div className="w-full h-full min-h-[600px] flex items-center justify-center bg-[#F4F6F8]" style={{ fontFamily: "'Inter', sans-serif" }}>
        <div className="flex flex-col items-center gap-3">
          <Snowflake className="animate-spin text-[#2E6F95]" size={28} />
          <div className="text-sm text-slate-500">Loading warehouse data…</div>
        </div>
      </div>
    );
  }

  if (!loggedInUser) {
    return <LoginPage onLogin={handleLogin} error={loginError} />;
  }

  if (!currentWarehouseId || !root.warehouses[currentWarehouseId]) {
    return (
      <div className="w-full h-full min-h-[600px] flex items-center justify-center bg-[#F4F6F8]" style={{ fontFamily: "'Inter', sans-serif" }}>
        <div className="text-sm text-slate-500">No warehouse assigned to this account. Contact your platform owner.</div>
      </div>
    );
  }

  const whData = root.warehouses[currentWarehouseId];
  const whMeta = { id: currentWarehouseId, name: whData.name, owner: whData.owner, address: whData.address };
  const accessibleWarehouses = accessibleIds.map((id) => ({ id, name: root.warehouses[id]?.name })).filter((w) => w.name);
  const isPlatformOwner = loggedInUser.warehouseIds.length === 0;

  return (
    <ShellApp
      data={whData} setData={setWarehouseData}
      root={root} setRoot={persistRoot}
      whMeta={whMeta} currentWarehouseId={currentWarehouseId} setCurrentWarehouseId={setCurrentWarehouseId}
      accessibleWarehouses={accessibleWarehouses} isPlatformOwner={isPlatformOwner}
      tab={tab} setTab={setTab} loggedInUser={loggedInUser} onLogout={handleLogout}
      pushToast={pushToast} toasts={toasts} sidebarOpen={sidebarOpen} setSidebarOpen={setSidebarOpen}
    />
  );
}

/* ---------------------------------------------------------------------- */
/*  Login                                                                   */
/* ---------------------------------------------------------------------- */

function LoginPage({ onLogin, error }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  return (
    <div className="w-full min-h-[700px] h-full flex items-center justify-center bg-[#111823] p-4" style={{ fontFamily: "'Inter', sans-serif" }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Oswald:wght@500;600;700&display=swap'); .disp { font-family: 'Oswald', sans-serif; }`}</style>
      <div className="w-full max-w-sm bg-white rounded-xl shadow-2xl p-7">
        <div className="flex flex-col items-center mb-6">
          <div className="w-12 h-12 rounded-lg bg-[#2E6F95] flex items-center justify-center mb-3"><Snowflake size={22} color="#fff" /></div>
          <div className="text-lg font-semibold text-slate-800 disp tracking-wide">COLDLINE WMS</div>
          <div className="text-xs text-slate-400">Sign in to your warehouse</div>
        </div>
        <Field label="Username"><input className={inputCls} value={username} onChange={(e) => setUsername(e.target.value)} onKeyDown={(e) => e.key === "Enter" && onLogin(username, password)} autoFocus /></Field>
        <Field label="Password"><input type="password" className={inputCls} value={password} onChange={(e) => setPassword(e.target.value)} onKeyDown={(e) => e.key === "Enter" && onLogin(username, password)} /></Field>
        {error && <div className="text-xs text-red-600 bg-red-50 border border-red-200 rounded-md px-3 py-2 mb-3">{error}</div>}
        <button onClick={() => onLogin(username, password)} className="w-full bg-[#2E6F95] text-white text-sm font-medium py-2 rounded-md hover:bg-[#265b7c]">Sign In</button>
        <div className="mt-5 pt-4 border-t border-slate-100 text-[11px] text-slate-400 space-y-1">
          <div className="font-semibold text-slate-500 mb-1">Demo accounts</div>
          <div>owner / owner123 — Platform Owner (all warehouses)</div>
          <div>ayesha / ayesha123 — Cold Storage Warehouse</div>
          <div>bilal / bilal123 — North Bakery Hub</div>
          <div>guest / guest123 — Viewer (both warehouses)</div>
        </div>
      </div>
    </div>
  );
}

function ShellApp({ data, setData, root, setRoot, whMeta, currentWarehouseId, setCurrentWarehouseId, accessibleWarehouses, isPlatformOwner, tab, setTab, loggedInUser, onLogout, pushToast, toasts, sidebarOpen, setSidebarOpen }) {
  const itemById = useMemo(() => Object.fromEntries(data.items.map((i) => [i.id, i])), [data.items]);
  const supplierById = useMemo(() => Object.fromEntries(data.suppliers.map((s) => [s.id, s])), [data.suppliers]);
  const branchById = useMemo(() => Object.fromEntries(data.branches.map((b) => [b.id, b])), [data.branches]);
  const currentUser = loggedInUser;
  const role = currentUser.role;

  const stockAsOf = useCallback((itemId, dateInclusive) => {
    const item = itemById[itemId];
    if (!item) return 0;
    let bal = Number(item.openingStock) || 0;
    for (const s of data.stockIns) if (s.itemId === itemId && s.date <= dateInclusive) bal += Number(s.qty) || 0;
    for (const s of data.stockOuts) if (s.itemId === itemId && s.date <= dateInclusive) bal -= Number(s.qty) || 0;
    return bal;
  }, [data.stockIns, data.stockOuts, itemById]);

  const currentStock = useCallback((itemId) => stockAsOf(itemId, todayStr()), [stockAsOf]);

  const stockInPending = useCallback((si) => {
    const extra = data.payments.filter((p) => p.stockInId === si.id).reduce((a, p) => a + Number(p.amount || 0), 0);
    const total = Number(si.qty) * Number(si.rate);
    return Math.max(0, total - Number(si.paid || 0) - extra);
  }, [data.payments]);

  const ctx = { data, setData, itemById, supplierById, branchById, stockAsOf, currentStock, stockInPending, pushToast, role, currentUser, whMeta };

  const totalStockValue = useMemo(() => data.items.reduce((a, it) => a + currentStock(it.id) * Number(it.purchaseRate || 0), 0), [data.items, currentStock]);
  const lowStock = useMemo(() => data.items.filter((it) => currentStock(it.id) <= Number(it.minStock) && currentStock(it.id) > 0), [data.items, currentStock]);
  const outStock = useMemo(() => data.items.filter((it) => currentStock(it.id) <= 0), [data.items, currentStock]);
  const totalPending = useMemo(() => data.stockIns.reduce((a, si) => a + stockInPending(si), 0), [data.stockIns, stockInPending]);

  const canEdit = role !== "Viewer";
  const isAdmin = role === "Super Admin";

  const renderTab = () => {
    switch (tab) {
      case "dashboard": return <Dashboard {...ctx} totalStockValue={totalStockValue} lowStock={lowStock} outStock={outStock} totalPending={totalPending} />;
      case "items": return <ItemsPage {...ctx} canEdit={canEdit} />;
      case "stockin": return <StockInPage {...ctx} canEdit={canEdit} />;
      case "stockout": return <StockOutPage {...ctx} canEdit={canEdit} />;
      case "ledger": return <LedgerPage {...ctx} />;
      case "suppliers": return <SuppliersPage {...ctx} canEdit={canEdit} />;
      case "payments": return <PaymentsPage {...ctx} canEdit={canEdit} />;
      case "branches": return <BranchesPage {...ctx} canEdit={canEdit} />;
      case "dailyreport": return <DailyReportPage {...ctx} />;
      case "monthlyreport": return <MonthlyReportPage {...ctx} />;
      case "closing": return <ClosingPage {...ctx} totalStockValue={totalStockValue} canEdit={canEdit} />;
      case "settings": return <SettingsPage {...ctx} isAdmin={isAdmin} root={root} setRoot={setRoot} currentWarehouseId={currentWarehouseId} isPlatformOwner={isPlatformOwner} />;
      case "warehouses": return isPlatformOwner ? <WarehousesPage root={root} setRoot={setRoot} pushToast={pushToast} currentUser={currentUser} onEnterWarehouse={(id) => { setCurrentWarehouseId(id); setTab("dashboard"); }} /> : null;
      default: return null;
    }
  };

  const visibleNav = NAV.filter((n) => !n.platformOnly || isPlatformOwner);
  const activeLabel = visibleNav.find((n) => n.key === tab)?.label || "Dashboard";

  return (
    <div className="w-full min-h-[700px] h-full flex bg-[#F4F6F8] text-slate-800" style={{ fontFamily: "'Inter', sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Oswald:wght@500;600;700&family=JetBrains+Mono:wght@400;500;700&display=swap');
        * { box-sizing: border-box; }
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
        @keyframes fadein { from { opacity: 0; transform: translateY(6px);} to { opacity:1; transform:none; } }
        table { border-collapse: collapse; }
        .mono { font-family: 'JetBrains Mono', monospace; }
        .disp { font-family: 'Oswald', sans-serif; }
      `}</style>

      <aside className={cx(
        "bg-[#161E27] text-slate-300 flex flex-col shrink-0 transition-all duration-200 fixed md:static h-full z-40",
        sidebarOpen ? "w-64 left-0" : "w-64 -left-64 md:left-0 md:w-16"
      )}>
        <div className="flex items-center gap-2 px-4 h-16 border-b border-white/10 shrink-0">
          <div className="w-8 h-8 rounded bg-[#2E6F95] flex items-center justify-center shrink-0"><Snowflake size={16} color="#fff" /></div>
          <div className={cx("overflow-hidden", !sidebarOpen && "md:hidden")}>
            <div className="text-white font-semibold text-sm leading-tight disp tracking-wide">COLDLINE WMS</div>
            <div className="text-[10px] text-slate-400">Warehouse Control</div>
          </div>
        </div>
        <nav className="flex-1 overflow-y-auto py-3">
          {visibleNav.map((n, idx) => n.section ? (
            <div key={idx} className={cx("px-4 pt-4 pb-1 text-[10px] uppercase tracking-widest text-slate-500 font-semibold", !sidebarOpen && "md:hidden")}>{n.section}</div>
          ) : (
            <button key={n.key} onClick={() => { setTab(n.key); setSidebarOpen(false); }}
              className={cx("w-full flex items-center gap-3 px-4 py-2 text-sm transition-colors",
                tab === n.key ? "bg-[#2E6F95] text-white" : "text-slate-300 hover:bg-white/5")}>
              <n.icon size={16} className="shrink-0" />
              <span className={cx(!sidebarOpen && "md:hidden")}>{n.label}</span>
            </button>
          ))}
        </nav>
        <div className={cx("px-4 py-3 border-t border-white/10 text-[11px] text-slate-500", !sidebarOpen && "md:hidden")}>
          v2.0 · Multi-warehouse ready
        </div>
      </aside>

      <div className="flex-1 flex flex-col min-w-0 md:ml-16">
        <header className="h-16 bg-white border-b border-slate-200 flex items-center justify-between px-4 md:px-6 shrink-0 sticky top-0 z-30 gap-2">
          <div className="flex items-center gap-3 min-w-0">
            <button className="md:hidden p-2 rounded hover:bg-slate-100" onClick={() => setSidebarOpen((s) => !s)}><ChevronRight size={18} /></button>
            <div className="min-w-0">
              <h1 className="text-lg font-semibold text-slate-800 disp tracking-wide truncate">{activeLabel}</h1>
              <div className="text-[11px] text-slate-400">{fmtDate(todayStr())} · {whMeta.name}</div>
            </div>
          </div>
          <div className="flex items-center gap-2 md:gap-3 shrink-0">
            {(lowStock.length > 0 || outStock.length > 0) && (
              <div className="hidden sm:flex items-center gap-1 text-amber-600 bg-amber-50 border border-amber-200 rounded-full px-2.5 py-1 text-xs font-medium">
                <AlertTriangle size={13} /> {lowStock.length + outStock.length} stock alerts
              </div>
            )}
            {accessibleWarehouses.length > 1 && (
              <select value={currentWarehouseId} onChange={(e) => { setCurrentWarehouseId(e.target.value); setTab("dashboard"); }}
                className="hidden sm:block text-xs font-medium border border-slate-300 rounded-md px-2 py-1.5 bg-white text-slate-600 max-w-[150px]">
                {accessibleWarehouses.map((w) => <option key={w.id} value={w.id}>{w.name}</option>)}
              </select>
            )}
            <div className="text-right hidden md:block">
              <div className="text-xs font-medium text-slate-700 leading-tight">{currentUser.name}</div>
              <div className="text-[10px] text-slate-400 leading-tight">{role}</div>
            </div>
            <button onClick={onLogout} title="Log out" className="p-2 rounded-md border border-slate-300 text-slate-500 hover:bg-slate-50"><LogOut size={14} /></button>
          </div>
        </header>
        <main className="flex-1 overflow-y-auto p-4 md:p-6">{renderTab()}</main>
      </div>
      <Toast toasts={toasts} />
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/*  Warehouses (platform owner)                                            */
/* ---------------------------------------------------------------------- */

function WarehousesPage({ root, setRoot, pushToast, currentUser, onEnterWarehouse }) {
  const [editing, setEditing] = useState(null);
  const warehouseIds = Object.keys(root.warehouses);

  const stats = (id) => {
    const wh = root.warehouses[id];
    const stockValue = wh.items.reduce((a, it) => {
      let bal = Number(it.openingStock) || 0;
      for (const s of wh.stockIns) if (s.itemId === it.id) bal += Number(s.qty) || 0;
      for (const s of wh.stockOuts) if (s.itemId === it.id) bal -= Number(s.qty) || 0;
      return a + bal * Number(it.purchaseRate || 0);
    }, 0);
    const owners = root.users.filter((u) => u.warehouseIds.includes(id));
    return { itemCount: wh.items.length, branchCount: wh.branches.length, stockValue, users: owners };
  };

  const saveWarehouse = ({ meta, createLogin }) => {
    const isNew = !editing.id;
    const id = editing.id || ("wh-" + uid().slice(0, 8));
    const warehouses = { ...root.warehouses };
    if (isNew) {
      warehouses[id] = { name: meta.name, owner: meta.owner, address: meta.address, ...emptyWarehouseData() };
    } else {
      warehouses[id] = { ...warehouses[id], name: meta.name, owner: meta.owner, address: meta.address };
    }
    let users = root.users;
    if (isNew && createLogin && createLogin.username) {
      users = [...users, { id: uid(), name: meta.owner || meta.name, username: createLogin.username, password: createLogin.password || "changeme123", role: "Super Admin", warehouseIds: [id] }];
    }
    setRoot({ ...root, warehouses, users });
    pushToast(isNew ? "Warehouse created" : "Warehouse updated");
    setEditing(null);
  };

  return (
    <div>
      <Toolbar>
        <div className="flex-1" />
        <button onClick={() => setEditing({})} className="flex items-center gap-1.5 bg-[#2E6F95] text-white text-sm font-medium px-3 py-1.5 rounded-md"><Plus size={15} /> Add Warehouse</button>
      </Toolbar>
      <div className="grid gap-3 md:grid-cols-2">
        {warehouseIds.map((id) => {
          const wh = root.warehouses[id];
          const st = stats(id);
          return (
            <div key={id} className="bg-white rounded-lg border border-slate-200 p-4">
              <div className="flex items-start justify-between mb-2">
                <div className="flex items-center gap-2"><Building size={16} className="text-[#2E6F95]" /><span className="font-semibold text-slate-700">{wh.name}</span></div>
                <button onClick={() => setEditing({ id, name: wh.name, owner: wh.owner, address: wh.address })} className="p-1.5 rounded hover:bg-slate-100 text-slate-500"><Pencil size={14} /></button>
              </div>
              <div className="text-xs text-slate-400 mb-3">Owner: {wh.owner || "—"} · {wh.address || "No address set"}</div>
              <div className="grid grid-cols-3 gap-2 mb-3 text-center">
                <div><div className="text-[10px] uppercase text-slate-400">Items</div><div className="mono font-semibold">{st.itemCount}</div></div>
                <div><div className="text-[10px] uppercase text-slate-400">Branches</div><div className="mono font-semibold">{st.branchCount}</div></div>
                <div><div className="text-[10px] uppercase text-slate-400">Stock Value</div><div className="mono font-semibold">{fmtMoney(st.stockValue)}</div></div>
              </div>
              <div className="text-[11px] text-slate-400 mb-3">Logins: {st.users.map((u) => u.username).join(", ") || "none yet"}</div>
              <button onClick={() => onEnterWarehouse(id)} className="text-xs font-medium text-[#2E6F95] hover:underline">Enter this warehouse →</button>
            </div>
          );
        })}
      </div>
      {editing !== null && <WarehouseForm warehouse={editing} onSave={saveWarehouse} onClose={() => setEditing(null)} />}
    </div>
  );
}

function WarehouseForm({ warehouse, onSave, onClose }) {
  const [f, setF] = useState({ name: warehouse.name || "", owner: warehouse.owner || "", address: warehouse.address || "" });
  const [login, setLogin] = useState({ username: "", password: "" });
  const isNew = !warehouse.id;
  return (
    <Modal title={isNew ? "New Warehouse" : "Edit Warehouse"} onClose={onClose} wide>
      <Field label="Warehouse Name"><input className={inputCls} value={f.name} onChange={(e) => setF({ ...f, name: e.target.value })} /></Field>
      <Field label="Owner Name"><input className={inputCls} value={f.owner} onChange={(e) => setF({ ...f, owner: e.target.value })} /></Field>
      <Field label="Address"><textarea className={inputCls} rows={2} value={f.address} onChange={(e) => setF({ ...f, address: e.target.value })} /></Field>
      {isNew && (
        <div className="border-t border-slate-100 pt-3 mt-1">
          <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-500 mb-2">Owner Login (optional — can be added later in Settings)</div>
          <div className="grid grid-cols-2 gap-x-4">
            <Field label="Username"><input className={inputCls} value={login.username} onChange={(e) => setLogin({ ...login, username: e.target.value })} /></Field>
            <Field label="Password"><input className={inputCls} value={login.password} onChange={(e) => setLogin({ ...login, password: e.target.value })} placeholder="min. 6 characters" /></Field>
          </div>
        </div>
      )}
      <div className="flex justify-end gap-2 mt-2">
        <button onClick={onClose} className="px-3 py-1.5 rounded-md text-sm border border-slate-300 text-slate-600">Cancel</button>
        <button onClick={() => f.name.trim() && onSave({ meta: f, createLogin: isNew ? login : null })} className="px-3 py-1.5 rounded-md text-sm bg-[#2E6F95] text-white font-medium">Save Warehouse</button>
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------------------- */
/*  Dashboard                                                               */
/* ---------------------------------------------------------------------- */

function Dashboard({ data, stockAsOf, currentStock, totalStockValue, lowStock, outStock, totalPending, itemById }) {
  const t = todayStr();
  const yest = addDays(t, -1);

  const inToday = data.stockIns.filter((s) => s.date === t);
  const outToday = data.stockOuts.filter((s) => s.date === t);
  const openingValue = data.items.reduce((a, it) => a + stockAsOf(it.id, yest) * Number(it.purchaseRate || 0), 0);
  const inValueToday = inToday.reduce((a, s) => a + Number(s.qty) * Number(s.rate), 0);
  const outValueToday = outToday.reduce((a, s) => a + Number(s.qty) * Number(s.rate), 0);
  const purchaseToday = inValueToday;
  const paidToday = inToday.reduce((a, s) => a + Number(s.paid || 0), 0);
  const pendingToday = Math.max(0, purchaseToday - paidToday);

  const last7 = Array.from({ length: 7 }, (_, i) => addDays(t, i - 6));
  const chartData = last7.map((d) => ({
    date: fmtDate(d).slice(0, 6),
    In: data.stockIns.filter((s) => s.date === d).reduce((a, s) => a + Number(s.qty) * Number(s.rate), 0),
    Out: data.stockOuts.filter((s) => s.date === d).reduce((a, s) => a + Number(s.qty) * Number(s.rate), 0),
  }));

  const catValue = (data.categories || []).map((c) => ({
    name: c,
    value: data.items.filter((it) => it.category === c).reduce((a, it) => a + currentStock(it.id) * Number(it.purchaseRate || 0), 0),
  })).filter((c) => c.value > 0);

  const recent = [
    ...data.stockIns.map((s) => ({ ...s, kind: "in" })),
    ...data.stockOuts.map((s) => ({ ...s, kind: "out" })),
  ].sort((a, b) => (a.date < b.date ? 1 : -1)).slice(0, 8);

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
        <StatCard icon={Package} label="Opening Stock (value)" value={fmtMoney(openingValue)} accent="#2E6F95" />
        <StatCard icon={ArrowDownCircle} label="Today's Stock In" value={fmtMoney(inValueToday)} sub={`${inToday.length} entries`} accent="#3E8266" />
        <StatCard icon={ArrowUpCircle} label="Today's Stock Out" value={fmtMoney(outValueToday)} sub={`${outToday.length} dispatches`} accent="#B24A3A" />
        <StatCard icon={Wallet} label="Today's Purchase" value={fmtMoney(purchaseToday)} accent="#C1832E" />
        <StatCard icon={CheckCircle2} label="Today's Paid" value={fmtMoney(paidToday)} accent="#3E8266" />
        <StatCard icon={AlertTriangle} label="Today's Pending" value={fmtMoney(pendingToday)} accent="#B24A3A" />
        <StatCard icon={Snowflake} label="Closing Stock (value)" value={fmtMoney(openingValue + inValueToday - outValueToday)} accent="#2E6F95" />
        <StatCard icon={TrendingUp} label="Total Stock Value" value={fmtMoney(totalStockValue)} accent="#161E27" />
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
        <div className="lg:col-span-2 bg-white rounded-lg border border-slate-200 p-4">
          <h3 className="text-sm font-semibold text-slate-700 disp tracking-wide mb-3">STOCK IN VS OUT — LAST 7 DAYS</h3>
          <ResponsiveContainer width="100%" height={230}>
            <BarChart data={chartData}>
              <CartesianGrid strokeDasharray="3 3" stroke="#eef1f4" />
              <XAxis dataKey="date" tick={{ fontSize: 11 }} />
              <YAxis tick={{ fontSize: 11 }} tickFormatter={(v) => "$" + v} />
              <Tooltip formatter={(v) => fmtMoney(v)} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Bar dataKey="In" fill="#3E8266" radius={[3, 3, 0, 0]} />
              <Bar dataKey="Out" fill="#B24A3A" radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
        <div className="bg-white rounded-lg border border-slate-200 p-4">
          <h3 className="text-sm font-semibold text-slate-700 disp tracking-wide mb-3">STOCK VALUE BY CATEGORY</h3>
          <ResponsiveContainer width="100%" height={230}>
            <PieChart>
              <Pie data={catValue} dataKey="value" nameKey="name" innerRadius={45} outerRadius={80} paddingAngle={2}>
                {catValue.map((c, i) => <Cell key={i} fill={catColor(c.name)} />)}
              </Pie>
              <Tooltip formatter={(v) => fmtMoney(v)} />
            </PieChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
        <div className="bg-white rounded-lg border border-slate-200 p-4">
          <h3 className="text-sm font-semibold text-slate-700 disp tracking-wide mb-3 flex items-center gap-2"><AlertTriangle size={15} className="text-amber-600" /> LOW STOCK / OUT OF STOCK</h3>
          {lowStock.length + outStock.length === 0 ? <EmptyState icon={CheckCircle2} title="All items are well stocked" /> : (
            <div className="space-y-2 max-h-64 overflow-y-auto">
              {outStock.map((it) => (
                <div key={it.id} className="flex items-center justify-between text-sm border-b border-slate-100 pb-2">
                  <span className="text-slate-700">{it.name}</span>
                  <Badge tone="red">OUT OF STOCK</Badge>
                </div>
              ))}
              {lowStock.map((it) => (
                <div key={it.id} className="flex items-center justify-between text-sm border-b border-slate-100 pb-2">
                  <span className="text-slate-700">{it.name}</span>
                  <Badge tone="amber">{fmtQty(currentStock(it.id))} {it.unit} left</Badge>
                </div>
              ))}
            </div>
          )}
        </div>
        <div className="bg-white rounded-lg border border-slate-200 p-4">
          <h3 className="text-sm font-semibold text-slate-700 disp tracking-wide mb-3">PENDING PAYMENTS</h3>
          <div className="text-2xl font-bold text-slate-800 mono mb-1">{fmtMoney(totalPending)}</div>
          <div className="text-xs text-slate-400">Outstanding across all suppliers</div>
        </div>
        <div className="bg-white rounded-lg border border-slate-200 p-4 lg:col-span-1">
          <h3 className="text-sm font-semibold text-slate-700 disp tracking-wide mb-3">RECENT TRANSACTIONS</h3>
          <div className="space-y-2 max-h-48 overflow-y-auto">
            {recent.map((r) => (
              <div key={r.id} className="flex items-center justify-between text-xs">
                <div className="flex items-center gap-2 min-w-0">
                  {r.kind === "in" ? <ArrowDownCircle size={13} className="text-emerald-600 shrink-0" /> : <ArrowUpCircle size={13} className="text-red-500 shrink-0" />}
                  <span className="truncate text-slate-600">{itemById[r.itemId]?.name}</span>
                </div>
                <span className="text-slate-400 shrink-0 ml-2">{fmtDate(r.date).slice(0, 6)}</span>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/*  Items                                                                   */
/* ---------------------------------------------------------------------- */

function ItemsPage({ data, setData, currentStock, supplierById, pushToast, canEdit, currentUser }) {
  const [q, setQ] = useState("");
  const [cat, setCat] = useState("");
  const [editing, setEditing] = useState(null);
  const filtered = data.items.filter((it) =>
    (!q || it.name.toLowerCase().includes(q.toLowerCase())) && (!cat || it.category === cat)
  );

  const save = (form) => {
    const isNew = !editing.id;
    const items = editing.id
      ? data.items.map((i) => i.id === editing.id ? { ...editing, ...form, updatedBy: currentUser.name, updatedAt: nowIso() } : i)
      : [...data.items, { id: uid(), status: "Active", createdBy: currentUser.name, createdAt: nowIso(), ...form }];
    const activityLog = pushLog(data, currentUser.name, isNew ? "Item Created" : "Item Updated", form.name);
    setData({ ...data, items, activityLog });
    pushToast(editing.id ? "Item updated" : "Item created");
    setEditing(null);
  };

  const remove = (it) => {
    const items = data.items.map((i) => i.id === it.id ? { ...i, status: "Discontinued", updatedBy: currentUser.name, updatedAt: nowIso() } : i);
    setData({ ...data, items, activityLog: pushLog(data, currentUser.name, "Item Discontinued", it.name) });
    pushToast("Item marked as discontinued");
  };

  const exportExcel = () => {
    const rows = filtered.map((it) => ({
      Item: it.name, Category: it.category, SubCategory: it.subCategory, Unit: it.unit,
      CurrentStock: currentStock(it.id), MinLevel: it.minStock, Rate: it.purchaseRate,
      Supplier: supplierById[it.supplier]?.name || "", Status: it.status,
    }));
    exportRowsToExcel("items", "Items", rows);
  };

  return (
    <div>
      <Toolbar>
        <SearchInput value={q} onChange={setQ} placeholder="Search items..." />
        <SelectFilter value={cat} onChange={setCat} placeholder="All Categories" options={(data.categories || []).map((c) => ({ value: c, label: c }))} />
        <div className="flex-1" />
        <ExportButton onClick={exportExcel} />
        {canEdit && <button onClick={() => setEditing({})} className="flex items-center gap-1.5 bg-[#2E6F95] text-white text-sm font-medium px-3 py-1.5 rounded-md hover:bg-[#265b7c]"><Plus size={15} /> Add Item</button>}
      </Toolbar>

      {filtered.length === 0 ? <EmptyState title="No items found" sub="Try adjusting your search or add a new item to the warehouse." /> : (
        <div className="bg-white rounded-lg border border-slate-200 overflow-x-auto">
          <table className="w-full text-sm min-w-[880px]">
            <thead>
              <tr className="text-left text-[11px] uppercase tracking-wide text-slate-400 border-b border-slate-200">
                <th className="px-4 py-2.5">Item</th><th className="px-4 py-2.5">Category</th><th className="px-4 py-2.5">Unit</th>
                <th className="px-4 py-2.5 text-right">Current Stock</th><th className="px-4 py-2.5 text-right">Min Level</th>
                <th className="px-4 py-2.5 text-right">Rate</th><th className="px-4 py-2.5">Supplier</th><th className="px-4 py-2.5">Status</th>
                {canEdit && <th className="px-4 py-2.5"></th>}
              </tr>
            </thead>
            <tbody>
              {filtered.map((it) => {
                const stock = currentStock(it.id);
                const low = stock <= it.minStock && stock > 0;
                const out = stock <= 0;
                return (
                  <tr key={it.id} className="border-b border-slate-100 hover:bg-slate-50">
                    <td className="px-4 py-2.5 font-medium text-slate-700">{it.name}<div className="text-[11px] text-slate-400">{it.subCategory}</div></td>
                    <td className="px-4 py-2.5"><Badge>{it.category}</Badge></td>
                    <td className="px-4 py-2.5 text-slate-500">{it.unit}</td>
                    <td className={cx("px-4 py-2.5 text-right mono", out ? "text-red-600 font-semibold" : low ? "text-amber-600 font-semibold" : "text-slate-700")}>{fmtQty(stock)}</td>
                    <td className="px-4 py-2.5 text-right text-slate-400 mono">{fmtQty(it.minStock)}</td>
                    <td className="px-4 py-2.5 text-right mono">{fmtMoney(it.purchaseRate)}</td>
                    <td className="px-4 py-2.5 text-slate-500">{supplierById[it.supplier]?.name || "-"}</td>
                    <td className="px-4 py-2.5">
                      {out ? <Badge tone="red">OUT OF STOCK</Badge> : low ? <Badge tone="amber">LOW STOCK</Badge> : it.status === "Discontinued" ? <Badge>Discontinued</Badge> : <Badge tone="green">Active</Badge>}
                    </td>
                    {canEdit && <td className="px-4 py-2.5 text-right whitespace-nowrap">
                      <button onClick={() => setEditing(it)} className="p-1.5 rounded hover:bg-slate-100 text-slate-500"><Pencil size={14} /></button>
                      <button onClick={() => remove(it)} className="p-1.5 rounded hover:bg-slate-100 text-slate-400"><Trash2 size={14} /></button>
                    </td>}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {editing !== null && (
        <ItemForm item={editing} suppliers={data.suppliers} categories={data.categories} units={data.units} onSave={save} onClose={() => setEditing(null)} />
      )}
    </div>
  );
}

function ItemForm({ item, suppliers, categories, units, onSave, onClose }) {
  const [f, setF] = useState({
    name: item.name || "", category: item.category || categories[0], subCategory: item.subCategory || "",
    unit: item.unit || units[0], minStock: item.minStock ?? 0, purchaseRate: item.purchaseRate ?? 0,
    openingStock: item.openingStock ?? 0, supplier: item.supplier || (suppliers[0]?.id || ""), notes: item.notes || "",
  });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  return (
    <Modal title={item.id ? "Edit Item" : "New Item"} onClose={onClose} wide>
      <div className="grid grid-cols-2 gap-x-4">
        <Field label="Item Name"><input className={inputCls} value={f.name} onChange={set("name")} /></Field>
        <Field label="Sub Category"><input className={inputCls} value={f.subCategory} onChange={set("subCategory")} /></Field>
        <Field label="Category">
          <select className={inputCls} value={f.category} onChange={set("category")}>{categories.map((c) => <option key={c}>{c}</option>)}</select>
        </Field>
        <Field label="Unit">
          <select className={inputCls} value={f.unit} onChange={set("unit")}>{units.map((u) => <option key={u}>{u}</option>)}</select>
        </Field>
        <Field label="Minimum Stock Level"><input type="number" className={inputCls} value={f.minStock} onChange={set("minStock")} /></Field>
        <Field label="Purchase Rate"><input type="number" step="0.01" className={inputCls} value={f.purchaseRate} onChange={set("purchaseRate")} /></Field>
        <Field label="Opening Stock" hint="Baseline stock before any recorded transactions"><input type="number" className={inputCls} value={f.openingStock} onChange={set("openingStock")} /></Field>
        <Field label="Supplier">
          <select className={inputCls} value={f.supplier} onChange={set("supplier")}>{suppliers.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}</select>
        </Field>
      </div>
      <Field label="Notes"><textarea className={inputCls} rows={2} value={f.notes} onChange={set("notes")} /></Field>
      <div className="flex justify-end gap-2 mt-2">
        <button onClick={onClose} className="px-3 py-1.5 rounded-md text-sm border border-slate-300 text-slate-600">Cancel</button>
        <button onClick={() => f.name.trim() && onSave({ ...f, minStock: Number(f.minStock), purchaseRate: Number(f.purchaseRate), openingStock: Number(f.openingStock) })}
          className="px-3 py-1.5 rounded-md text-sm bg-[#2E6F95] text-white font-medium">Save Item</button>
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------------------- */
/*  Stock In                                                                */
/* ---------------------------------------------------------------------- */

function StockInPage({ data, setData, itemById, supplierById, pushToast, canEdit, currentUser }) {
  const [form, setForm] = useState(null);
  const [q, setQ] = useState("");
  const [dateFilter, setDateFilter] = useState("");
  const isLocked = (date) => !!data.closings[date]?.locked;

  const rows = data.stockIns.filter((s) =>
    (!q || itemById[s.itemId]?.name.toLowerCase().includes(q.toLowerCase()) || s.invoice.toLowerCase().includes(q.toLowerCase())) &&
    (!dateFilter || s.date === dateFilter)
  ).sort((a, b) => (a.date < b.date ? 1 : -1));

  const add = (f) => {
    if (isLocked(f.date)) return pushToast("This day is closed. Ask an admin to unlock it first.", "error");
    const entry = { id: uid(), createdBy: currentUser.name, createdAt: nowIso(), ...f };
    const activityLog = pushLog(data, currentUser.name, "Stock In Recorded", `${itemById[f.itemId]?.name} · ${f.qty} ${f.unit} · ${f.invoice}`);
    setData({ ...data, stockIns: [entry, ...data.stockIns], activityLog });
    pushToast("Stock In recorded — stock updated");
    setForm(null);
  };

  const exportExcel = () => {
    const rows2 = rows.map((s) => ({
      Date: s.date, Invoice: s.invoice, Item: itemById[s.itemId]?.name, Supplier: supplierById[s.supplier]?.name,
      Qty: s.qty, Unit: s.unit, Rate: s.rate, Total: s.qty * s.rate, Paid: s.paid, Pending: Math.max(0, s.qty * s.rate - s.paid),
    }));
    exportRowsToExcel("stock-in", "Stock In", rows2);
  };

  return (
    <div>
      <Toolbar>
        <SearchInput value={q} onChange={setQ} placeholder="Search item / invoice..." />
        <input type="date" value={dateFilter} onChange={(e) => setDateFilter(e.target.value)} className="px-2.5 py-1.5 rounded-md border border-slate-300 text-sm" />
        <div className="flex-1" />
        <ExportButton onClick={exportExcel} />
        {canEdit && <button onClick={() => setForm({})} className="flex items-center gap-1.5 bg-[#3E8266] text-white text-sm font-medium px-3 py-1.5 rounded-md hover:bg-[#336d53]"><Plus size={15} /> Record Stock In</button>}
      </Toolbar>
      {rows.length === 0 ? <EmptyState icon={ArrowDownCircle} title="No stock-in entries" sub="Record incoming purchases to increase warehouse stock." /> : (
        <div className="bg-white rounded-lg border border-slate-200 overflow-x-auto">
          <table className="w-full text-sm min-w-[880px]">
            <thead><tr className="text-left text-[11px] uppercase tracking-wide text-slate-400 border-b border-slate-200">
              <th className="px-4 py-2.5">Date</th><th className="px-4 py-2.5">Invoice</th><th className="px-4 py-2.5">Item</th><th className="px-4 py-2.5">Supplier</th>
              <th className="px-4 py-2.5 text-right">Qty</th><th className="px-4 py-2.5 text-right">Rate</th><th className="px-4 py-2.5 text-right">Total</th>
              <th className="px-4 py-2.5 text-right">Paid</th><th className="px-4 py-2.5 text-right">Pending</th><th className="px-4 py-2.5"></th>
            </tr></thead>
            <tbody>
              {rows.map((s) => {
                const total = s.qty * s.rate;
                const pending = Math.max(0, total - s.paid);
                return (
                  <tr key={s.id} className="border-b border-slate-100 hover:bg-slate-50">
                    <td className="px-4 py-2.5 text-slate-500">{fmtDate(s.date)}</td>
                    <td className="px-4 py-2.5 mono text-slate-600">{s.invoice}</td>
                    <td className="px-4 py-2.5 font-medium text-slate-700">{itemById[s.itemId]?.name}</td>
                    <td className="px-4 py-2.5 text-slate-500">{supplierById[s.supplier]?.name}</td>
                    <td className="px-4 py-2.5 text-right mono">{fmtQty(s.qty)} {s.unit}</td>
                    <td className="px-4 py-2.5 text-right mono">{fmtMoney(s.rate)}</td>
                    <td className="px-4 py-2.5 text-right mono font-medium">{fmtMoney(total)}</td>
                    <td className="px-4 py-2.5 text-right mono text-emerald-600">{fmtMoney(s.paid)}</td>
                    <td className="px-4 py-2.5 text-right mono">{pending > 0 ? <span className="text-amber-600 font-semibold">{fmtMoney(pending)}</span> : <span className="text-slate-300">—</span>}</td>
                    <td className="px-4 py-2.5">{isLocked(s.date) && <Lock size={13} className="text-slate-400" />}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
      {form !== null && <StockInForm items={data.items} suppliers={data.suppliers} onSave={add} onClose={() => setForm(null)} />}
    </div>
  );
}

function StockInForm({ items, suppliers, onSave, onClose }) {
  const [f, setF] = useState({ date: todayStr(), invoice: "", supplier: suppliers[0]?.id || "", itemId: items[0]?.id || "", qty: "", unit: items[0]?.unit || "", rate: items[0]?.purchaseRate || "", paid: "", method: PAY_METHODS[0], remarks: "" });
  const set = (k) => (e) => {
    const v = e.target.value;
    if (k === "itemId") {
      const it = items.find((i) => i.id === v);
      setF({ ...f, itemId: v, unit: it?.unit || "", rate: it?.purchaseRate || "" });
    } else setF({ ...f, [k]: v });
  };
  const total = (Number(f.qty) || 0) * (Number(f.rate) || 0);
  const pending = Math.max(0, total - (Number(f.paid) || 0));
  return (
    <Modal title="Record Stock In / Purchase" onClose={onClose} wide>
      <div className="grid grid-cols-2 gap-x-4">
        <Field label="Date"><input type="date" className={inputCls} value={f.date} onChange={set("date")} /></Field>
        <Field label="Invoice Number"><input className={inputCls} value={f.invoice} onChange={set("invoice")} placeholder="INV-####" /></Field>
        <Field label="Supplier"><select className={inputCls} value={f.supplier} onChange={set("supplier")}>{suppliers.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}</select></Field>
        <Field label="Item"><select className={inputCls} value={f.itemId} onChange={set("itemId")}>{items.map((i) => <option key={i.id} value={i.id}>{i.name}</option>)}</select></Field>
        <Field label={`Quantity (${f.unit})`}><input type="number" className={inputCls} value={f.qty} onChange={set("qty")} /></Field>
        <Field label="Purchase Rate"><input type="number" step="0.01" className={inputCls} value={f.rate} onChange={set("rate")} /></Field>
        <Field label="Paid Amount"><input type="number" step="0.01" className={inputCls} value={f.paid} onChange={set("paid")} /></Field>
        <Field label="Payment Method"><select className={inputCls} value={f.method} onChange={set("method")}>{PAY_METHODS.map((m) => <option key={m}>{m}</option>)}</select></Field>
      </div>
      <Field label="Remarks"><textarea className={inputCls} rows={2} value={f.remarks} onChange={set("remarks")} /></Field>
      <div className="flex justify-between text-sm bg-slate-50 rounded-md px-3 py-2 mb-3 mono">
        <span>Total: <b>{fmtMoney(total)}</b></span><span>Pending: <b className={pending > 0 ? "text-amber-600" : ""}>{fmtMoney(pending)}</b></span>
      </div>
      <div className="flex justify-end gap-2">
        <button onClick={onClose} className="px-3 py-1.5 rounded-md text-sm border border-slate-300 text-slate-600">Cancel</button>
        <button onClick={() => f.invoice.trim() && f.qty && onSave({ ...f, qty: Number(f.qty), rate: Number(f.rate), paid: Number(f.paid) || 0 })}
          className="px-3 py-1.5 rounded-md text-sm bg-[#3E8266] text-white font-medium">Save & Update Stock</button>
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------------------- */
/*  Stock Out                                                               */
/* ---------------------------------------------------------------------- */

function StockOutPage({ data, setData, itemById, branchById, currentStock, stockAsOf, pushToast, canEdit, currentUser }) {
  const [form, setForm] = useState(null);
  const [q, setQ] = useState("");
  const [branchFilter, setBranchFilter] = useState("");
  const isLocked = (date) => !!data.closings[date]?.locked;

  const rows = data.stockOuts.filter((s) =>
    (!q || itemById[s.itemId]?.name.toLowerCase().includes(q.toLowerCase())) &&
    (!branchFilter || s.destination === branchFilter)
  ).sort((a, b) => (a.date < b.date ? 1 : -1));

  const add = (f) => {
    if (isLocked(f.date)) return pushToast("This day is closed. Ask an admin to unlock it first.", "error");
    const availableAsOfYesterday = stockAsOf(f.itemId, addDays(f.date, -1));
    const alreadyOutSameDay = data.stockOuts.filter((s) => s.itemId === f.itemId && s.date === f.date).reduce((a, s) => a + s.qty, 0);
    const inSameDay = data.stockIns.filter((s) => s.itemId === f.itemId && s.date === f.date).reduce((a, s) => a + s.qty, 0);
    const availableToday = availableAsOfYesterday + inSameDay - alreadyOutSameDay;
    if (f.qty > availableToday) {
      pushToast(`Insufficient stock — only ${fmtQty(availableToday)} ${itemById[f.itemId]?.unit} available.`, "error");
      return;
    }
    const entry = { id: uid(), createdBy: currentUser.name, createdAt: nowIso(), ...f };
    const activityLog = pushLog(data, currentUser.name, "Stock Out Recorded", `${itemById[f.itemId]?.name} · ${f.qty} ${f.unit} → ${branchById[f.destination]?.name}`);
    setData({ ...data, stockOuts: [entry, ...data.stockOuts], activityLog });
    pushToast("Stock Out recorded — stock updated");
    setForm(null);
  };

  const exportExcel = () => {
    const rows2 = rows.map((s) => ({
      Date: s.date, Destination: branchById[s.destination]?.name, Item: itemById[s.itemId]?.name,
      Qty: s.qty, Unit: s.unit, Rate: s.rate, Value: s.qty * s.rate, Receiver: s.receiver,
    }));
    exportRowsToExcel("stock-out", "Stock Out", rows2);
  };

  return (
    <div>
      <Toolbar>
        <SearchInput value={q} onChange={setQ} placeholder="Search item..." />
        <SelectFilter value={branchFilter} onChange={setBranchFilter} placeholder="All Branches" options={data.branches.map((b) => ({ value: b.id, label: b.name }))} />
        <div className="flex-1" />
        <ExportButton onClick={exportExcel} />
        {canEdit && <button onClick={() => setForm({})} className="flex items-center gap-1.5 bg-[#B24A3A] text-white text-sm font-medium px-3 py-1.5 rounded-md hover:bg-[#9c3f31]"><Plus size={15} /> Record Dispatch</button>}
      </Toolbar>
      {rows.length === 0 ? <EmptyState icon={ArrowUpCircle} title="No dispatch records" sub="Send stock to a branch or restaurant to see it here." /> : (
        <div className="bg-white rounded-lg border border-slate-200 overflow-x-auto">
          <table className="w-full text-sm min-w-[760px]">
            <thead><tr className="text-left text-[11px] uppercase tracking-wide text-slate-400 border-b border-slate-200">
              <th className="px-4 py-2.5">Date</th><th className="px-4 py-2.5">Destination</th><th className="px-4 py-2.5">Item</th>
              <th className="px-4 py-2.5 text-right">Qty</th><th className="px-4 py-2.5 text-right">Value</th><th className="px-4 py-2.5">Receiver</th><th className="px-4 py-2.5"></th>
            </tr></thead>
            <tbody>
              {rows.map((s) => (
                <tr key={s.id} className="border-b border-slate-100 hover:bg-slate-50">
                  <td className="px-4 py-2.5 text-slate-500">{fmtDate(s.date)}</td>
                  <td className="px-4 py-2.5"><Badge tone="blue">{branchById[s.destination]?.name}</Badge></td>
                  <td className="px-4 py-2.5 font-medium text-slate-700">{itemById[s.itemId]?.name}</td>
                  <td className="px-4 py-2.5 text-right mono">{fmtQty(s.qty)} {s.unit}</td>
                  <td className="px-4 py-2.5 text-right mono">{fmtMoney(s.qty * s.rate)}</td>
                  <td className="px-4 py-2.5 text-slate-500">{s.receiver}</td>
                  <td className="px-4 py-2.5">{isLocked(s.date) && <Lock size={13} className="text-slate-400" />}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
      {form !== null && <StockOutForm items={data.items} branches={data.branches} currentStock={currentStock} onSave={add} onClose={() => setForm(null)} />}
    </div>
  );
}

function StockOutForm({ items, branches, currentStock, onSave, onClose }) {
  const [f, setF] = useState({ date: todayStr(), destination: branches[0]?.id || "", itemId: items[0]?.id || "", qty: "", unit: items[0]?.unit || "", rate: items[0]?.purchaseRate || "", receiver: "", remarks: "" });
  const set = (k) => (e) => {
    const v = e.target.value;
    if (k === "itemId") {
      const it = items.find((i) => i.id === v);
      setF({ ...f, itemId: v, unit: it?.unit || "", rate: it?.purchaseRate || "" });
    } else setF({ ...f, [k]: v });
  };
  const avail = currentStock(f.itemId);
  const over = Number(f.qty) > avail;
  return (
    <Modal title="Record Stock Out / Dispatch" onClose={onClose} wide>
      <div className="grid grid-cols-2 gap-x-4">
        <Field label="Date"><input type="date" className={inputCls} value={f.date} onChange={set("date")} /></Field>
        <Field label="Destination / Restaurant"><select className={inputCls} value={f.destination} onChange={set("destination")}>{branches.map((b) => <option key={b.id} value={b.id}>{b.name}</option>)}</select></Field>
        <Field label="Item"><select className={inputCls} value={f.itemId} onChange={set("itemId")}>{items.map((i) => <option key={i.id} value={i.id}>{i.name}</option>)}</select></Field>
        <Field label="Receiver Name"><input className={inputCls} value={f.receiver} onChange={set("receiver")} /></Field>
        <Field label={`Quantity (${f.unit})`} hint={`Available: ${fmtQty(avail)} ${f.unit}`}>
          <input type="number" className={cx(inputCls, over && "border-red-400 ring-1 ring-red-200")} value={f.qty} onChange={set("qty")} />
        </Field>
        <Field label="Rate"><input type="number" step="0.01" className={inputCls} value={f.rate} onChange={set("rate")} /></Field>
      </div>
      {over && <div className="flex items-center gap-2 text-xs text-red-600 bg-red-50 border border-red-200 rounded-md px-3 py-2 mb-3"><AlertTriangle size={13} /> Quantity exceeds available stock.</div>}
      <Field label="Remarks"><textarea className={inputCls} rows={2} value={f.remarks} onChange={set("remarks")} /></Field>
      <div className="flex justify-end gap-2 mt-1">
        <button onClick={onClose} className="px-3 py-1.5 rounded-md text-sm border border-slate-300 text-slate-600">Cancel</button>
        <button disabled={over || !f.qty || !f.receiver.trim()} onClick={() => onSave({ ...f, qty: Number(f.qty), rate: Number(f.rate) })}
          className="px-3 py-1.5 rounded-md text-sm bg-[#B24A3A] text-white font-medium disabled:opacity-40 disabled:cursor-not-allowed">Save & Dispatch</button>
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------------------- */
/*  Item Ledger                                                             */
/* ---------------------------------------------------------------------- */

function LedgerPage({ data, itemById }) {
  const [itemId, setItemId] = useState(data.items[0]?.id || "");
  const item = itemById[itemId];
  const events = useMemo(() => {
    if (!item) return [];
    const ins = data.stockIns.filter((s) => s.itemId === itemId).map((s) => ({ date: s.date, type: "Stock In", ref: s.invoice, qtyIn: s.qty, qtyOut: 0 }));
    const outs = data.stockOuts.filter((s) => s.itemId === itemId).map((s) => ({ date: s.date, type: "Stock Out", ref: s.receiver, qtyIn: 0, qtyOut: s.qty }));
    const merged = [...ins, ...outs].sort((a, b) => a.date < b.date ? -1 : a.date > b.date ? 1 : 0);
    let bal = Number(item.openingStock) || 0;
    return merged.map((e) => { bal += e.qtyIn - e.qtyOut; return { ...e, balance: bal }; });
  }, [data.stockIns, data.stockOuts, itemId, item]);

  const exportExcel = () => {
    const rows = events.map((e) => ({ Date: e.date, Type: e.type, Reference: e.ref, In: e.qtyIn, Out: e.qtyOut, Balance: e.balance }));
    exportRowsToExcel(`ledger-${item?.name || "item"}`, "Ledger", rows);
  };

  return (
    <div>
      <Toolbar>
        <SelectFilter value={itemId} onChange={setItemId} placeholder="Select item" options={data.items.map((i) => ({ value: i.id, label: i.name }))} />
        <div className="flex-1" />
        <ExportButton onClick={exportExcel} />
      </Toolbar>
      {item && (
        <div className="bg-white rounded-lg border border-slate-200 p-4 mb-4 flex flex-wrap gap-6">
          <div><div className="text-[11px] uppercase text-slate-400 font-semibold">Opening Stock</div><div className="mono font-semibold">{fmtQty(item.openingStock)} {item.unit}</div></div>
          <div><div className="text-[11px] uppercase text-slate-400 font-semibold">Total In</div><div className="mono font-semibold text-emerald-600">{fmtQty(events.reduce((a, e) => a + e.qtyIn, 0))}</div></div>
          <div><div className="text-[11px] uppercase text-slate-400 font-semibold">Total Out</div><div className="mono font-semibold text-red-500">{fmtQty(events.reduce((a, e) => a + e.qtyOut, 0))}</div></div>
          <div><div className="text-[11px] uppercase text-slate-400 font-semibold">Current Balance</div><div className="mono font-bold text-slate-800">{fmtQty(events.length ? events[events.length - 1].balance : item.openingStock)} {item.unit}</div></div>
        </div>
      )}
      {events.length === 0 ? <EmptyState icon={BookOpen} title="No ledger movements yet" /> : (
        <div className="bg-white rounded-lg border border-slate-200 overflow-x-auto">
          <table className="w-full text-sm min-w-[600px]">
            <thead><tr className="text-left text-[11px] uppercase tracking-wide text-slate-400 border-b border-slate-200">
              <th className="px-4 py-2.5">Date</th><th className="px-4 py-2.5">Type</th><th className="px-4 py-2.5">Reference</th>
              <th className="px-4 py-2.5 text-right">In</th><th className="px-4 py-2.5 text-right">Out</th><th className="px-4 py-2.5 text-right">Balance</th>
            </tr></thead>
            <tbody>
              {events.map((e, i) => (
                <tr key={i} className="border-b border-slate-100 hover:bg-slate-50">
                  <td className="px-4 py-2.5 text-slate-500">{fmtDate(e.date)}</td>
                  <td className="px-4 py-2.5">{e.type === "Stock In" ? <Badge tone="green">Stock In</Badge> : <Badge tone="red">Stock Out</Badge>}</td>
                  <td className="px-4 py-2.5 text-slate-500">{e.ref}</td>
                  <td className="px-4 py-2.5 text-right mono text-emerald-600">{e.qtyIn ? fmtQty(e.qtyIn) : "—"}</td>
                  <td className="px-4 py-2.5 text-right mono text-red-500">{e.qtyOut ? fmtQty(e.qtyOut) : "—"}</td>
                  <td className="px-4 py-2.5 text-right mono font-semibold">{fmtQty(e.balance)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/*  Suppliers                                                               */
/* ---------------------------------------------------------------------- */

function SuppliersPage({ data, setData, stockInPending, pushToast, canEdit, currentUser }) {
  const [editing, setEditing] = useState(null);
  const [expanded, setExpanded] = useState(null);

  const stats = (supId) => {
    const ins = data.stockIns.filter((s) => s.supplier === supId);
    const total = ins.reduce((a, s) => a + s.qty * s.rate, 0);
    const paid = ins.reduce((a, s) => a + s.paid, 0) + data.payments.filter((p) => ins.some((i) => i.id === p.stockInId)).reduce((a, p) => a + Number(p.amount), 0);
    return { total, paid, pending: Math.max(0, total - paid), count: ins.length };
  };

  const save = (f) => {
    const isNew = !editing.id;
    const suppliers = editing.id
      ? data.suppliers.map((s) => s.id === editing.id ? { ...editing, ...f, updatedBy: currentUser.name, updatedAt: nowIso() } : s)
      : [...data.suppliers, { id: uid(), createdBy: currentUser.name, createdAt: nowIso(), ...f }];
    setData({ ...data, suppliers, activityLog: pushLog(data, currentUser.name, isNew ? "Supplier Added" : "Supplier Updated", f.name) });
    pushToast(editing.id ? "Supplier updated" : "Supplier added");
    setEditing(null);
  };

  const exportExcel = () => {
    const rows = data.suppliers.map((s) => { const st = stats(s.id); return { Supplier: s.name, Contact: s.contact, Address: s.address, TotalPurchased: st.total, Paid: st.paid, Pending: st.pending }; });
    exportRowsToExcel("suppliers", "Suppliers", rows);
  };

  return (
    <div>
      <Toolbar><div className="flex-1" /><ExportButton onClick={exportExcel} />{canEdit && <button onClick={() => setEditing({})} className="flex items-center gap-1.5 bg-[#2E6F95] text-white text-sm font-medium px-3 py-1.5 rounded-md"><Plus size={15} /> Add Supplier</button>}</Toolbar>
      <div className="grid gap-3">
        {data.suppliers.map((s) => {
          const st = stats(s.id);
          const open = expanded === s.id;
          return (
            <div key={s.id} className="bg-white rounded-lg border border-slate-200">
              <div className="flex items-center justify-between p-4 cursor-pointer" onClick={() => setExpanded(open ? null : s.id)}>
                <div>
                  <div className="font-semibold text-slate-700">{s.name}</div>
                  <div className="text-xs text-slate-400">{s.contact} · {s.address}</div>
                </div>
                <div className="flex items-center gap-5">
                  <div className="text-right hidden sm:block"><div className="text-[10px] uppercase text-slate-400">Purchased</div><div className="mono text-sm">{fmtMoney(st.total)}</div></div>
                  <div className="text-right hidden sm:block"><div className="text-[10px] uppercase text-slate-400">Paid</div><div className="mono text-sm text-emerald-600">{fmtMoney(st.paid)}</div></div>
                  <div className="text-right"><div className="text-[10px] uppercase text-slate-400">Pending</div><div className="mono text-sm font-semibold text-amber-600">{fmtMoney(st.pending)}</div></div>
                  {canEdit && <button onClick={(e) => { e.stopPropagation(); setEditing(s); }} className="p-1.5 rounded hover:bg-slate-100 text-slate-500"><Pencil size={14} /></button>}
                  <ChevronDown size={16} className={cx("text-slate-400 transition-transform", open && "rotate-180")} />
                </div>
              </div>
              {open && (
                <div className="border-t border-slate-100 px-4 py-3">
                  {st.count === 0 ? <div className="text-xs text-slate-400">No purchase history yet.</div> : (
                    <table className="w-full text-xs">
                      <thead><tr className="text-left text-slate-400 uppercase text-[10px]"><th className="py-1">Date</th><th>Invoice</th><th className="text-right">Total</th><th className="text-right">Paid</th><th className="text-right">Pending</th></tr></thead>
                      <tbody>
                        {data.stockIns.filter((si) => si.supplier === s.id).map((si) => (
                          <tr key={si.id} className="border-t border-slate-100">
                            <td className="py-1.5">{fmtDate(si.date)}</td><td className="mono">{si.invoice}</td>
                            <td className="text-right mono">{fmtMoney(si.qty * si.rate)}</td>
                            <td className="text-right mono text-emerald-600">{fmtMoney(si.paid)}</td>
                            <td className="text-right mono text-amber-600">{fmtMoney(stockInPending(si))}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </div>
              )}
            </div>
          );
        })}
      </div>
      {editing !== null && <SupplierForm supplier={editing} onSave={save} onClose={() => setEditing(null)} />}
    </div>
  );
}

function SupplierForm({ supplier, onSave, onClose }) {
  const [f, setF] = useState({ name: supplier.name || "", contact: supplier.contact || "", address: supplier.address || "" });
  return (
    <Modal title={supplier.id ? "Edit Supplier" : "New Supplier"} onClose={onClose}>
      <Field label="Supplier Name"><input className={inputCls} value={f.name} onChange={(e) => setF({ ...f, name: e.target.value })} /></Field>
      <Field label="Contact"><input className={inputCls} value={f.contact} onChange={(e) => setF({ ...f, contact: e.target.value })} /></Field>
      <Field label="Address"><textarea className={inputCls} rows={2} value={f.address} onChange={(e) => setF({ ...f, address: e.target.value })} /></Field>
      <div className="flex justify-end gap-2">
        <button onClick={onClose} className="px-3 py-1.5 rounded-md text-sm border border-slate-300 text-slate-600">Cancel</button>
        <button onClick={() => f.name.trim() && onSave(f)} className="px-3 py-1.5 rounded-md text-sm bg-[#2E6F95] text-white font-medium">Save</button>
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------------------- */
/*  Payments                                                                */
/* ---------------------------------------------------------------------- */

function PaymentsPage({ data, setData, supplierById, stockInPending, pushToast, canEdit, currentUser }) {
  const [payFor, setPayFor] = useState(null);
  const [status, setStatus] = useState("");

  const rows = data.stockIns.map((si) => ({ ...si, pending: stockInPending(si), total: si.qty * si.rate }))
    .filter((si) => !status || (status === "pending" ? si.pending > 0 : si.pending === 0))
    .sort((a, b) => (a.date < b.date ? 1 : -1));

  const totals = data.stockIns.reduce((acc, si) => {
    const total = si.qty * si.rate;
    acc.payable += total; acc.pending += stockInPending(si); return acc;
  }, { payable: 0, pending: 0 });
  totals.paid = totals.payable - totals.pending;

  const recordPayment = (si, amount, method, date) => {
    const payments = [...data.payments, { id: uid(), stockInId: si.id, amount: Number(amount), method, date, createdBy: currentUser.name, createdAt: nowIso() }];
    setData({ ...data, payments, activityLog: pushLog(data, currentUser.name, "Payment Recorded", `${si.invoice} · ${fmtMoney(amount)}`) });
    pushToast("Payment recorded");
    setPayFor(null);
  };

  const exportExcel = () => {
    const rows2 = rows.map((si) => ({ Supplier: supplierById[si.supplier]?.name, Invoice: si.invoice, Date: si.date, Total: si.total, Paid: si.total - si.pending, Pending: si.pending }));
    exportRowsToExcel("payments", "Payments", rows2);
  };

  return (
    <div>
      <div className="grid grid-cols-3 gap-3 mb-4">
        <StatCard icon={Wallet} label="Total Payable" value={fmtMoney(totals.payable)} accent="#161E27" />
        <StatCard icon={CheckCircle2} label="Total Paid" value={fmtMoney(totals.paid)} accent="#3E8266" />
        <StatCard icon={AlertTriangle} label="Total Pending" value={fmtMoney(totals.pending)} accent="#B24A3A" />
      </div>
      <Toolbar>
        <SelectFilter value={status} onChange={setStatus} placeholder="All Payment Status" options={[{ value: "pending", label: "Pending" }, { value: "paid", label: "Fully Paid" }]} />
        <div className="flex-1" />
        <ExportButton onClick={exportExcel} />
      </Toolbar>
      <div className="bg-white rounded-lg border border-slate-200 overflow-x-auto">
        <table className="w-full text-sm min-w-[820px]">
          <thead><tr className="text-left text-[11px] uppercase tracking-wide text-slate-400 border-b border-slate-200">
            <th className="px-4 py-2.5">Supplier</th><th className="px-4 py-2.5">Invoice</th><th className="px-4 py-2.5">Date</th>
            <th className="px-4 py-2.5 text-right">Total</th><th className="px-4 py-2.5 text-right">Paid</th><th className="px-4 py-2.5 text-right">Pending</th><th className="px-4 py-2.5"></th>
          </tr></thead>
          <tbody>
            {rows.map((si) => (
              <tr key={si.id} className="border-b border-slate-100 hover:bg-slate-50">
                <td className="px-4 py-2.5 text-slate-700">{supplierById[si.supplier]?.name}</td>
                <td className="px-4 py-2.5 mono text-slate-500">{si.invoice}</td>
                <td className="px-4 py-2.5 text-slate-500">{fmtDate(si.date)}</td>
                <td className="px-4 py-2.5 text-right mono">{fmtMoney(si.total)}</td>
                <td className="px-4 py-2.5 text-right mono text-emerald-600">{fmtMoney(si.total - si.pending)}</td>
                <td className="px-4 py-2.5 text-right mono">{si.pending > 0 ? <span className="text-amber-600 font-semibold">{fmtMoney(si.pending)}</span> : <Badge tone="green">Settled</Badge>}</td>
                <td className="px-4 py-2.5 text-right">{canEdit && si.pending > 0 && <button onClick={() => setPayFor(si)} className="text-xs font-medium text-[#2E6F95] hover:underline">Record Payment</button>}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {payFor && <PaymentForm stockIn={payFor} onSave={recordPayment} onClose={() => setPayFor(null)} />}
    </div>
  );
}

function PaymentForm({ stockIn, onSave, onClose }) {
  const [amount, setAmount] = useState("");
  const [method, setMethod] = useState(PAY_METHODS[0]);
  const [date, setDate] = useState(todayStr());
  return (
    <Modal title={`Pay Invoice ${stockIn.invoice}`} onClose={onClose}>
      <div className="text-sm text-slate-500 mb-3">Pending amount: <b className="text-amber-600 mono">{fmtMoney(stockIn.pending)}</b></div>
      <Field label="Payment Date"><input type="date" className={inputCls} value={date} onChange={(e) => setDate(e.target.value)} /></Field>
      <Field label="Amount"><input type="number" step="0.01" className={inputCls} value={amount} onChange={(e) => setAmount(e.target.value)} /></Field>
      <Field label="Payment Method"><select className={inputCls} value={method} onChange={(e) => setMethod(e.target.value)}>{PAY_METHODS.map((m) => <option key={m}>{m}</option>)}</select></Field>
      <div className="flex justify-end gap-2 mt-1">
        <button onClick={onClose} className="px-3 py-1.5 rounded-md text-sm border border-slate-300 text-slate-600">Cancel</button>
        <button disabled={!amount || Number(amount) <= 0} onClick={() => onSave(stockIn, Math.min(Number(amount), stockIn.pending), method, date)}
          className="px-3 py-1.5 rounded-md text-sm bg-[#2E6F95] text-white font-medium disabled:opacity-40">Save Payment</button>
      </div>
    </Modal>
  );
}

/* ---------------------------------------------------------------------- */
/*  Branches                                                                */
/* ---------------------------------------------------------------------- */

function BranchesPage({ data, setData, itemById, pushToast, canEdit, currentUser }) {
  const [editing, setEditing] = useState(null);
  const [expanded, setExpanded] = useState(null);

  const save = (f) => {
    const isNew = !editing.id;
    const branches = editing.id
      ? data.branches.map((b) => b.id === editing.id ? { ...editing, ...f, updatedBy: currentUser.name, updatedAt: nowIso() } : b)
      : [...data.branches, { id: uid(), createdBy: currentUser.name, createdAt: nowIso(), ...f }];
    setData({ ...data, branches, activityLog: pushLog(data, currentUser.name, isNew ? "Branch Added" : "Branch Updated", f.name) });
    pushToast(editing.id ? "Branch updated" : "Branch added");
    setEditing(null);
  };

  return (
    <div>
      <Toolbar><div className="flex-1" />{canEdit && <button onClick={() => setEditing({})} className="flex items-center gap-1.5 bg-[#2E6F95] text-white text-sm font-medium px-3 py-1.5 rounded-md"><Plus size={15} /> Add Branch</button>}</Toolbar>
      <div className="grid gap-3">
        {data.branches.map((b) => {
          const dispatches = data.stockOuts.filter((s) => s.destination === b.id);
          const value = dispatches.reduce((a, s) => a + s.qty * s.rate, 0);
          const open = expanded === b.id;
          return (
            <div key={b.id} className="bg-white rounded-lg border border-slate-200">
              <div className="flex items-center justify-between p-4 cursor-pointer" onClick={() => setExpanded(open ? null : b.id)}>
                <div className="flex items-center gap-2"><Building2 size={16} className="text-[#2E6F95]" /><span className="font-semibold text-slate-700">{b.name}</span></div>
                <div className="flex items-center gap-5">
                  <div className="text-right"><div className="text-[10px] uppercase text-slate-400">Received (value)</div><div className="mono text-sm font-semibold">{fmtMoney(value)}</div></div>
                  <div className="text-right hidden sm:block"><div className="text-[10px] uppercase text-slate-400">Dispatches</div><div className="mono text-sm">{dispatches.length}</div></div>
                  {canEdit && <button onClick={(e) => { e.stopPropagation(); setEditing(b); }} className="p-1.5 rounded hover:bg-slate-100 text-slate-500"><Pencil size={14} /></button>}
                  <ChevronDown size={16} className={cx("text-slate-400 transition-transform", open && "rotate-180")} />
                </div>
              </div>
              {open && (
                <div className="border-t border-slate-100 px-4 py-3">
                  {dispatches.length === 0 ? <div className="text-xs text-slate-400">No dispatch history yet.</div> : (
                    <table className="w-full text-xs">
                      <thead><tr className="text-left text-slate-400 uppercase text-[10px]"><th className="py-1">Date</th><th>Item</th><th className="text-right">Qty</th><th className="text-right">Value</th><th>Receiver</th></tr></thead>
                      <tbody>
                        {dispatches.map((s) => (
                          <tr key={s.id} className="border-t border-slate-100">
                            <td className="py-1.5">{fmtDate(s.date)}</td><td>{itemById[s.itemId]?.name}</td>
                            <td className="text-right mono">{fmtQty(s.qty)} {s.unit}</td><td className="text-right mono">{fmtMoney(s.qty * s.rate)}</td><td>{s.receiver}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </div>
              )}
            </div>
          );
        })}
      </div>
      {editing !== null && (
        <Modal title={editing.id ? "Edit Branch" : "New Branch"} onClose={() => setEditing(null)}>
          <BranchFormBody branch={editing} onSave={save} onClose={() => setEditing(null)} />
        </Modal>
      )}
    </div>
  );
}

function BranchFormBody({ branch, onSave, onClose }) {
  const [name, setName] = useState(branch.name || "");
  return (
    <div>
      <Field label="Branch / Restaurant Name"><input className={inputCls} value={name} onChange={(e) => setName(e.target.value)} /></Field>
      <div className="flex justify-end gap-2">
        <button onClick={onClose} className="px-3 py-1.5 rounded-md text-sm border border-slate-300 text-slate-600">Cancel</button>
        <button onClick={() => name.trim() && onSave({ name })} className="px-3 py-1.5 rounded-md text-sm bg-[#2E6F95] text-white font-medium">Save</button>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/*  Daily Report                                                            */
/* ---------------------------------------------------------------------- */

function DailyReportPage({ data, stockAsOf, whMeta }) {
  const [date, setDate] = useState(todayStr());
  const prev = addDays(date, -1);

  const inRows = data.stockIns.filter((s) => s.date === date);
  const outRows = data.stockOuts.filter((s) => s.date === date);
  const totalPurchase = inRows.reduce((a, s) => a + s.qty * s.rate, 0);
  const totalPaid = inRows.reduce((a, s) => a + s.paid, 0);
  const totalPending = Math.max(0, totalPurchase - totalPaid);

  const itemRows = data.items.map((it) => {
    const opening = stockAsOf(it.id, prev);
    const inQty = inRows.filter((s) => s.itemId === it.id).reduce((a, s) => a + s.qty, 0);
    const outQty = outRows.filter((s) => s.itemId === it.id).reduce((a, s) => a + s.qty, 0);
    const closing = opening + inQty - outQty;
    return { ...it, opening, inQty, outQty, closing, value: closing * it.purchaseRate };
  });

  const stockOpening = itemRows.reduce((a, r) => a + r.opening * r.purchaseRate, 0);
  const stockIn = itemRows.reduce((a, r) => a + r.inQty * r.purchaseRate, 0);
  const stockOut = itemRows.reduce((a, r) => a + r.outQty * r.purchaseRate, 0);
  const stockClosing = itemRows.reduce((a, r) => a + r.value, 0);

  const exportExcel = () => {
    const rows = itemRows.map((r) => ({ Item: r.name, Opening: r.opening, In: r.inQty, Out: r.outQty, Closing: r.closing, Rate: r.purchaseRate, Value: r.value }));
    exportRowsToExcel(`daily-report-${date}`, "Daily Report", rows);
  };

  return (
    <div>
      <Toolbar>
        <input type="date" value={date} onChange={(e) => setDate(e.target.value)} className="px-2.5 py-1.5 rounded-md border border-slate-300 text-sm" />
        <div className="flex-1" />
        <button onClick={() => window.print()} className="flex items-center gap-1.5 text-sm border border-slate-300 rounded-md px-3 py-1.5 text-slate-600"><Printer size={14} /> Print / Save PDF</button>
        <ExportButton onClick={exportExcel} />
      </Toolbar>

      <div className="bg-white rounded-lg border border-slate-200 p-5 mb-4">
        <div className="text-center border-b border-dashed border-slate-300 pb-3 mb-4">
          <div className="disp text-lg font-semibold tracking-wide">{data.settings?.companyName}</div>
          <div className="text-xs text-slate-400">{whMeta?.name} · Daily Closing Report · {fmtDate(date)}</div>
        </div>
        <div className="grid grid-cols-2 md:grid-cols-4 gap-4 text-center">
          <div><div className="text-[11px] uppercase text-slate-400 font-semibold">Opening Stock</div><div className="mono font-bold">{fmtMoney(stockOpening)}</div></div>
          <div><div className="text-[11px] uppercase text-slate-400 font-semibold">Stock In</div><div className="mono font-bold text-emerald-600">{fmtMoney(stockIn)}</div></div>
          <div><div className="text-[11px] uppercase text-slate-400 font-semibold">Stock Out</div><div className="mono font-bold text-red-500">{fmtMoney(stockOut)}</div></div>
          <div><div className="text-[11px] uppercase text-slate-400 font-semibold">Closing Stock</div><div className="mono font-bold">{fmtMoney(stockClosing)}</div></div>
          <div><div className="text-[11px] uppercase text-slate-400 font-semibold">Total Purchase</div><div className="mono font-bold">{fmtMoney(totalPurchase)}</div></div>
          <div><div className="text-[11px] uppercase text-slate-400 font-semibold">Total Paid</div><div className="mono font-bold text-emerald-600">{fmtMoney(totalPaid)}</div></div>
          <div><div className="text-[11px] uppercase text-slate-400 font-semibold">Total Pending</div><div className="mono font-bold text-amber-600">{fmtMoney(totalPending)}</div></div>
        </div>
      </div>

      <div className="bg-white rounded-lg border border-slate-200 overflow-x-auto">
        <table className="w-full text-sm min-w-[760px]">
          <thead><tr className="text-left text-[11px] uppercase tracking-wide text-slate-400 border-b border-slate-200">
            <th className="px-4 py-2.5">Item</th><th className="px-4 py-2.5 text-right">Opening</th><th className="px-4 py-2.5 text-right">In</th>
            <th className="px-4 py-2.5 text-right">Out</th><th className="px-4 py-2.5 text-right">Closing</th><th className="px-4 py-2.5 text-right">Rate</th><th className="px-4 py-2.5 text-right">Value</th>
          </tr></thead>
          <tbody>
            {itemRows.map((r) => (
              <tr key={r.id} className="border-b border-slate-100">
                <td className="px-4 py-2.5 font-medium text-slate-700">{r.name}</td>
                <td className="px-4 py-2.5 text-right mono">{fmtQty(r.opening)}</td>
                <td className="px-4 py-2.5 text-right mono text-emerald-600">{r.inQty ? fmtQty(r.inQty) : "—"}</td>
                <td className="px-4 py-2.5 text-right mono text-red-500">{r.outQty ? fmtQty(r.outQty) : "—"}</td>
                <td className="px-4 py-2.5 text-right mono font-semibold">{fmtQty(r.closing)}</td>
                <td className="px-4 py-2.5 text-right mono text-slate-400">{fmtMoney(r.purchaseRate)}</td>
                <td className="px-4 py-2.5 text-right mono">{fmtMoney(r.value)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/*  Monthly Report                                                          */
/* ---------------------------------------------------------------------- */

function MonthlyReportPage({ data, stockAsOf }) {
  const [month, setMonth] = useState(todayStr().slice(0, 7));
  const start = month + "-01";
  const endDate = new Date(month + "-01T00:00:00"); endDate.setMonth(endDate.getMonth() + 1); endDate.setDate(0);
  const end = endDate.toISOString().slice(0, 10);
  const dayBefore = addDays(start, -1);

  const insInMonth = data.stockIns.filter((s) => s.date >= start && s.date <= end);
  const outsInMonth = data.stockOuts.filter((s) => s.date >= start && s.date <= end);

  const itemRows = data.items.map((it) => {
    const opening = stockAsOf(it.id, dayBefore);
    const inQty = insInMonth.filter((s) => s.itemId === it.id).reduce((a, s) => a + s.qty, 0);
    const outQty = outsInMonth.filter((s) => s.itemId === it.id).reduce((a, s) => a + s.qty, 0);
    return { ...it, opening, inQty, outQty, closing: opening + inQty - outQty };
  });

  const totalPurchase = insInMonth.reduce((a, s) => a + s.qty * s.rate, 0);
  const totalPaid = insInMonth.reduce((a, s) => a + s.paid, 0);
  const totalPending = Math.max(0, totalPurchase - totalPaid);
  const openingValue = itemRows.reduce((a, r) => a + r.opening * r.purchaseRate, 0);
  const closingValue = itemRows.reduce((a, r) => a + r.closing * r.purchaseRate, 0);
  const inValue = itemRows.reduce((a, r) => a + r.inQty * r.purchaseRate, 0);
  const outValue = itemRows.reduce((a, r) => a + r.outQty * r.purchaseRate, 0);

  const daysInMonth = endDate.getDate();
  const dailySeries = Array.from({ length: daysInMonth }, (_, i) => {
    const d = start.slice(0, 8) + String(i + 1).padStart(2, "0");
    return {
      day: i + 1,
      In: data.stockIns.filter((s) => s.date === d).reduce((a, s) => a + s.qty * s.rate, 0),
      Out: data.stockOuts.filter((s) => s.date === d).reduce((a, s) => a + s.qty * s.rate, 0),
      Purchase: data.stockIns.filter((s) => s.date === d).reduce((a, s) => a + s.qty * s.rate, 0),
      Paid: data.stockIns.filter((s) => s.date === d).reduce((a, s) => a + s.paid, 0),
    };
  });

  const topMoving = [...itemRows].sort((a, b) => (b.inQty + b.outQty) - (a.inQty + a.outQty)).slice(0, 6)
    .map((r) => ({ name: r.name.length > 14 ? r.name.slice(0, 13) + "…" : r.name, Movement: r.inQty + r.outQty }));

  const exportExcel = () => {
    const rows = itemRows.map((r) => ({ Item: r.name, Opening: r.opening, In: r.inQty, Out: r.outQty, Closing: r.closing }));
    exportRowsToExcel(`monthly-summary-${month}`, "Monthly Summary", rows);
  };

  return (
    <div>
      <Toolbar>
        <input type="month" value={month} onChange={(e) => setMonth(e.target.value)} className="px-2.5 py-1.5 rounded-md border border-slate-300 text-sm" />
        <div className="flex-1" />
        <button onClick={() => window.print()} className="flex items-center gap-1.5 text-sm border border-slate-300 rounded-md px-3 py-1.5 text-slate-600"><Printer size={14} /> Print / Save PDF</button>
        <ExportButton onClick={exportExcel} />
      </Toolbar>

      <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-4">
        <StatCard icon={Package} label="Opening Stock" value={fmtMoney(openingValue)} accent="#2E6F95" />
        <StatCard icon={ArrowDownCircle} label="Total Stock In" value={fmtMoney(inValue)} accent="#3E8266" />
        <StatCard icon={ArrowUpCircle} label="Total Stock Out" value={fmtMoney(outValue)} accent="#B24A3A" />
        <StatCard icon={Snowflake} label="Closing Stock" value={fmtMoney(closingValue)} accent="#161E27" />
        <StatCard icon={Wallet} label="Total Purchase" value={fmtMoney(totalPurchase)} accent="#C1832E" />
        <StatCard icon={CheckCircle2} label="Total Paid" value={fmtMoney(totalPaid)} accent="#3E8266" />
        <StatCard icon={AlertTriangle} label="Total Pending" value={fmtMoney(totalPending)} accent="#B24A3A" />
        <StatCard icon={TrendingUp} label="Stock Value" value={fmtMoney(closingValue)} accent="#161E27" />
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-4">
        <div className="bg-white rounded-lg border border-slate-200 p-4">
          <h3 className="text-sm font-semibold text-slate-700 disp tracking-wide mb-3">DAILY STOCK IN / OUT</h3>
          <ResponsiveContainer width="100%" height={220}>
            <LineChart data={dailySeries}>
              <CartesianGrid strokeDasharray="3 3" stroke="#eef1f4" /><XAxis dataKey="day" tick={{ fontSize: 10 }} /><YAxis tick={{ fontSize: 10 }} />
              <Tooltip formatter={(v) => fmtMoney(v)} /><Legend wrapperStyle={{ fontSize: 11 }} />
              <Line type="monotone" dataKey="In" stroke="#3E8266" dot={false} strokeWidth={2} />
              <Line type="monotone" dataKey="Out" stroke="#B24A3A" dot={false} strokeWidth={2} />
            </LineChart>
          </ResponsiveContainer>
        </div>
        <div className="bg-white rounded-lg border border-slate-200 p-4">
          <h3 className="text-sm font-semibold text-slate-700 disp tracking-wide mb-3">PAYMENT VS PENDING (BY DAY)</h3>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={dailySeries}>
              <CartesianGrid strokeDasharray="3 3" stroke="#eef1f4" /><XAxis dataKey="day" tick={{ fontSize: 10 }} /><YAxis tick={{ fontSize: 10 }} />
              <Tooltip formatter={(v) => fmtMoney(v)} /><Legend wrapperStyle={{ fontSize: 11 }} />
              <Bar dataKey="Paid" stackId="a" fill="#3E8266" /><Bar dataKey="Purchase" stackId="b" fill="#C1832E" fillOpacity={0.35} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div className="bg-white rounded-lg border border-slate-200 p-4 mb-4">
        <h3 className="text-sm font-semibold text-slate-700 disp tracking-wide mb-3">TOP MOVING ITEMS</h3>
        <ResponsiveContainer width="100%" height={200}>
          <BarChart data={topMoving} layout="vertical" margin={{ left: 10 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#eef1f4" />
            <XAxis type="number" tick={{ fontSize: 10 }} /><YAxis type="category" dataKey="name" width={100} tick={{ fontSize: 11 }} />
            <Tooltip /><Bar dataKey="Movement" fill="#2E6F95" radius={[0, 3, 3, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <div className="bg-white rounded-lg border border-slate-200 overflow-x-auto">
        <table className="w-full text-sm min-w-[600px]">
          <thead><tr className="text-left text-[11px] uppercase tracking-wide text-slate-400 border-b border-slate-200">
            <th className="px-4 py-2.5">Item</th><th className="px-4 py-2.5 text-right">Opening</th><th className="px-4 py-2.5 text-right">In</th><th className="px-4 py-2.5 text-right">Out</th><th className="px-4 py-2.5 text-right">Closing</th>
          </tr></thead>
          <tbody>
            {itemRows.map((r) => (
              <tr key={r.id} className="border-b border-slate-100">
                <td className="px-4 py-2.5 font-medium text-slate-700">{r.name}</td>
                <td className="px-4 py-2.5 text-right mono">{fmtQty(r.opening)}</td>
                <td className="px-4 py-2.5 text-right mono text-emerald-600">{r.inQty ? fmtQty(r.inQty) : "—"}</td>
                <td className="px-4 py-2.5 text-right mono text-red-500">{r.outQty ? fmtQty(r.outQty) : "—"}</td>
                <td className="px-4 py-2.5 text-right mono font-semibold">{fmtQty(r.closing)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/*  Daily Closing                                                           */
/* ---------------------------------------------------------------------- */

function ClosingPage({ data, setData, stockAsOf, totalStockValue, pushToast, canEdit, role, currentUser }) {
  const [date, setDate] = useState(todayStr());
  const prev = addDays(date, -1);
  const closing = data.closings[date];
  const isLocked = !!closing?.locked;

  const inRows = data.stockIns.filter((s) => s.date === date);
  const outRows = data.stockOuts.filter((s) => s.date === date);
  const openingValue = data.items.reduce((a, it) => a + stockAsOf(it.id, prev) * it.purchaseRate, 0);
  const inValue = inRows.reduce((a, s) => a + s.qty * s.rate, 0);
  const outValue = outRows.reduce((a, s) => a + s.qty * s.rate, 0);
  const purchase = inValue;
  const paid = inRows.reduce((a, s) => a + s.paid, 0);
  const pending = Math.max(0, purchase - paid);
  const closingValue = openingValue + inValue - outValue;

  const closeDay = () => {
    const closings = { ...data.closings, [date]: { opening: openingValue, in: inValue, out: outValue, purchase, paid, pending, closing: closingValue, locked: true, closedBy: currentUser.name, closedAt: nowIso() } };
    setData({ ...data, closings, activityLog: pushLog(data, currentUser.name, "Day Closed", fmtDate(date)) });
    pushToast(`Day closed for ${fmtDate(date)} — transactions locked`);
  };
  const unlockDay = () => {
    const closings = { ...data.closings, [date]: { ...closing, locked: false, unlockedBy: currentUser.name, unlockedAt: nowIso() } };
    setData({ ...data, closings, activityLog: pushLog(data, currentUser.name, "Day Unlocked (Admin)", fmtDate(date)) });
    pushToast("Day unlocked by admin — edits allowed", "warn");
  };

  return (
    <div>
      <Toolbar>
        <input type="date" value={date} onChange={(e) => setDate(e.target.value)} className="px-2.5 py-1.5 rounded-md border border-slate-300 text-sm" />
        {isLocked && <Badge tone="red"><Lock size={11} /> Day Closed</Badge>}
      </Toolbar>

      <div className="bg-white rounded-lg border border-slate-200 p-5">
        <h3 className="text-sm font-semibold text-slate-700 disp tracking-wide mb-4">CLOSING PREVIEW — {fmtDate(date)}</h3>
        <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-5">
          {[
            ["Opening Stock", openingValue, "#2E6F95"], ["Total Stock In", inValue, "#3E8266"], ["Total Stock Out", outValue, "#B24A3A"],
            ["Purchase Amount", purchase, "#C1832E"], ["Paid Amount", paid, "#3E8266"], ["Pending Amount", pending, "#B24A3A"],
            ["Closing Stock", closingValue, "#161E27"], ["Total Stock Value", totalStockValue, "#161E27"],
          ].map(([label, val, color]) => (
            <div key={label} className="border border-slate-100 rounded-md p-3">
              <div className="text-[10px] uppercase text-slate-400 font-semibold mb-1">{label}</div>
              <div className="mono font-bold" style={{ color }}>{fmtMoney(val)}</div>
            </div>
          ))}
        </div>

        {!canEdit ? (
          <div className="text-xs text-slate-400 flex items-center gap-2"><AlertTriangle size={13} /> Viewer role cannot close or unlock days.</div>
        ) : isLocked ? (
          role === "Super Admin" ? (
            <button onClick={unlockDay} className="flex items-center gap-1.5 text-sm border border-slate-300 rounded-md px-3 py-1.5 text-slate-600 hover:bg-slate-50"><Unlock size={14} /> Unlock Day (Admin)</button>
          ) : (
            <div className="text-xs text-slate-400 flex items-center gap-2"><Lock size={13} /> This day is closed. Only a Super Admin can reopen it.</div>
          )
        ) : (
          <button onClick={closeDay} className="flex items-center gap-1.5 bg-[#161E27] text-white text-sm font-medium px-4 py-2 rounded-md hover:bg-[#0f151c]"><CalendarClock size={15} /> Close Day & Generate Report</button>
        )}
      </div>

      <div className="mt-5">
        <h3 className="text-sm font-semibold text-slate-700 disp tracking-wide mb-3">CLOSING HISTORY</h3>
        {Object.keys(data.closings).length === 0 ? <EmptyState icon={CalendarClock} title="No days closed yet" /> : (
          <div className="bg-white rounded-lg border border-slate-200 overflow-x-auto">
            <table className="w-full text-sm min-w-[600px]">
              <thead><tr className="text-left text-[11px] uppercase tracking-wide text-slate-400 border-b border-slate-200">
                <th className="px-4 py-2.5">Date</th><th className="px-4 py-2.5 text-right">Closing Stock</th><th className="px-4 py-2.5 text-right">Purchase</th><th className="px-4 py-2.5 text-right">Pending</th><th className="px-4 py-2.5">Status</th><th className="px-4 py-2.5">Closed By</th>
              </tr></thead>
              <tbody>
                {Object.entries(data.closings).sort((a, b) => (a[0] < b[0] ? 1 : -1)).map(([d, c]) => (
                  <tr key={d} className="border-b border-slate-100">
                    <td className="px-4 py-2.5 text-slate-600">{fmtDate(d)}</td>
                    <td className="px-4 py-2.5 text-right mono">{fmtMoney(c.closing)}</td>
                    <td className="px-4 py-2.5 text-right mono">{fmtMoney(c.purchase)}</td>
                    <td className="px-4 py-2.5 text-right mono text-amber-600">{fmtMoney(c.pending)}</td>
                    <td className="px-4 py-2.5">{c.locked ? <Badge tone="red"><Lock size={10} /> Locked</Badge> : <Badge tone="green"><Unlock size={10} /> Open</Badge>}</td>
                    <td className="px-4 py-2.5 text-slate-400 text-xs">{c.closedBy || "-"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/*  Settings                                                                */
/* ---------------------------------------------------------------------- */

const SETTINGS_TABS = [
  { key: "profile", label: "Company Profile", icon: Building },
  { key: "categories", label: "Categories", icon: Tags },
  { key: "units", label: "Units", icon: Ruler },
  { key: "users", label: "Users & Roles", icon: ShieldCheck },
  { key: "activity", label: "Activity Log", icon: ClipboardList },
];

function SettingsPage({ data, setData, pushToast, isAdmin, currentUser, root, setRoot, currentWarehouseId, isPlatformOwner }) {
  const [sub, setSub] = useState("profile");

  return (
    <div>
      <div className="flex flex-wrap gap-2 mb-5 border-b border-slate-200 pb-3">
        {SETTINGS_TABS.map((t) => (
          <button key={t.key} onClick={() => setSub(t.key)}
            className={cx("flex items-center gap-1.5 text-sm font-medium px-3 py-1.5 rounded-md",
              sub === t.key ? "bg-[#161E27] text-white" : "text-slate-500 hover:bg-slate-100")}>
            <t.icon size={14} /> {t.label}
          </button>
        ))}
      </div>
      {sub === "profile" && <CompanyProfileTab data={data} setData={setData} pushToast={pushToast} isAdmin={isAdmin} currentUser={currentUser} root={root} setRoot={setRoot} currentWarehouseId={currentWarehouseId} isPlatformOwner={isPlatformOwner} />}
      {sub === "categories" && <ListManagerTab title="Categories" data={data} setData={setData} pushToast={pushToast} isAdmin={isAdmin} field="categories" currentUser={currentUser} />}
      {sub === "units" && <ListManagerTab title="Units" data={data} setData={setData} pushToast={pushToast} isAdmin={isAdmin} field="units" currentUser={currentUser} />}
      {sub === "users" && <UsersTab root={root} setRoot={setRoot} currentWarehouseId={currentWarehouseId} pushToast={pushToast} isAdmin={isAdmin} currentUser={currentUser} />}
      {sub === "activity" && <ActivityLogTab data={data} />}
    </div>
  );
}

function CompanyProfileTab({ data, setData, pushToast, isAdmin, currentUser, root, setRoot, currentWarehouseId, isPlatformOwner }) {
  const wh = root.warehouses[currentWarehouseId];
  const [f, setF] = useState({ companyName: data.settings?.companyName || "", warehouseName: wh?.name || "", address: wh?.address || "" });
  const save = () => {
    setData({ ...data, settings: { ...data.settings, companyName: f.companyName }, activityLog: pushLog(data, currentUser.name, "Company Profile Updated", f.companyName) });
    if (isPlatformOwner) {
      setRoot({ ...root, warehouses: { ...root.warehouses, [currentWarehouseId]: { ...wh, name: f.warehouseName, address: f.address } } });
    }
    pushToast("Company profile saved");
  };
  return (
    <div className="bg-white rounded-lg border border-slate-200 p-5 max-w-lg">
      <Field label="Company Name"><input disabled={!isAdmin} className={inputCls} value={f.companyName} onChange={(e) => setF({ ...f, companyName: e.target.value })} /></Field>
      <Field label="Warehouse Name" hint={!isPlatformOwner ? "Only the platform owner can rename a warehouse." : ""}>
        <input disabled={!isAdmin || !isPlatformOwner} className={inputCls} value={f.warehouseName} onChange={(e) => setF({ ...f, warehouseName: e.target.value })} />
      </Field>
      <Field label="Warehouse Address">
        <input disabled={!isAdmin || !isPlatformOwner} className={inputCls} value={f.address} onChange={(e) => setF({ ...f, address: e.target.value })} />
      </Field>
      {isAdmin ? (
        <button onClick={save} className="px-3 py-1.5 rounded-md text-sm bg-[#2E6F95] text-white font-medium">Save Profile</button>
      ) : (
        <div className="text-xs text-slate-400 flex items-center gap-2"><AlertTriangle size={13} /> Only Super Admin can edit the company profile.</div>
      )}
    </div>
  );
}

function ListManagerTab({ title, data, setData, pushToast, isAdmin, field, currentUser }) {
  const [val, setVal] = useState("");
  const list = data[field] || [];
  const add = () => {
    const v = val.trim();
    if (!v || list.includes(v)) return;
    setData({ ...data, [field]: [...list, v], activityLog: pushLog(data, currentUser.name, `${title.slice(0, -1)} Added`, v) });
    setVal("");
    pushToast(`${title.slice(0, -1)} added`);
  };
  const remove = (v) => {
    setData({ ...data, [field]: list.filter((x) => x !== v), activityLog: pushLog(data, currentUser.name, `${title.slice(0, -1)} Removed`, v) });
  };
  return (
    <div className="bg-white rounded-lg border border-slate-200 p-5 max-w-lg">
      {isAdmin && (
        <div className="flex gap-2 mb-4">
          <input className={inputCls} placeholder={`New ${title.slice(0, -1).toLowerCase()}...`} value={val} onChange={(e) => setVal(e.target.value)} onKeyDown={(e) => e.key === "Enter" && add()} />
          <button onClick={add} className="px-3 py-1.5 rounded-md text-sm bg-[#2E6F95] text-white font-medium shrink-0"><Plus size={15} /></button>
        </div>
      )}
      <div className="flex flex-wrap gap-2">
        {list.map((v) => (
          <span key={v} className="inline-flex items-center gap-1.5 bg-slate-100 text-slate-700 text-sm px-2.5 py-1 rounded-md">
            {v}
            {isAdmin && <button onClick={() => remove(v)} className="text-slate-400 hover:text-red-500"><X size={12} /></button>}
          </span>
        ))}
      </div>
      {!isAdmin && <div className="text-xs text-slate-400 flex items-center gap-2 mt-3"><AlertTriangle size={13} /> Only Super Admin can manage {title.toLowerCase()}.</div>}
    </div>
  );
}

function UsersTab({ root, setRoot, currentWarehouseId, pushToast, isAdmin, currentUser }) {
  const [form, setForm] = useState({ name: "", username: "", password: "", role: ROLES[1] });
  const usersHere = root.users.filter((u) => u.warehouseIds.includes(currentWarehouseId));

  const addUser = () => {
    if (!form.name.trim() || !form.username.trim() || !form.password.trim()) return pushToast("Name, username, and password are required.", "error");
    if (root.users.some((u) => u.username.toLowerCase() === form.username.trim().toLowerCase())) return pushToast("That username is already taken.", "error");
    const users = [...root.users, { id: uid(), name: form.name.trim(), username: form.username.trim(), password: form.password, role: form.role, warehouseIds: [currentWarehouseId] }];
    setRoot({ ...root, users });
    setForm({ name: "", username: "", password: "", role: ROLES[1] });
    pushToast("User added");
  };
  const changeRole = (u, role) => {
    const users = root.users.map((x) => x.id === u.id ? { ...x, role } : x);
    setRoot({ ...root, users });
  };
  const removeUser = (u) => {
    if (usersHere.length <= 1) return pushToast("At least one user must remain for this warehouse.", "error");
    setRoot({ ...root, users: root.users.map((x) => x.id === u.id ? { ...x, warehouseIds: x.warehouseIds.filter((w) => w !== currentWarehouseId) } : x) });
    pushToast("User removed from this warehouse");
  };
  return (
    <div className="bg-white rounded-lg border border-slate-200 p-5">
      <div className="text-xs text-slate-400 mb-3">Each login (username + password) can be given access to this warehouse. A user with no warehouse assigned loses access.</div>
      {isAdmin && (
        <div className="flex flex-wrap gap-2 mb-5 items-end">
          <div className="w-40"><Field label="Full Name"><input className={inputCls} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} /></Field></div>
          <div className="w-32"><Field label="Username"><input className={inputCls} value={form.username} onChange={(e) => setForm({ ...form, username: e.target.value })} /></Field></div>
          <div className="w-32"><Field label="Password"><input className={inputCls} value={form.password} onChange={(e) => setForm({ ...form, password: e.target.value })} /></Field></div>
          <div className="w-36"><Field label="Role"><select className={inputCls} value={form.role} onChange={(e) => setForm({ ...form, role: e.target.value })}>{ROLES.map((r) => <option key={r}>{r}</option>)}</select></Field></div>
          <button onClick={addUser} className="px-3 py-2 rounded-md text-sm bg-[#2E6F95] text-white font-medium mb-3 flex items-center gap-1.5"><Plus size={14} /> Add</button>
        </div>
      )}
      <table className="w-full text-sm">
        <thead><tr className="text-left text-[11px] uppercase tracking-wide text-slate-400 border-b border-slate-200">
          <th className="py-2">Name</th><th className="py-2">Username</th><th className="py-2">Role</th>{isAdmin && <th className="py-2"></th>}
        </tr></thead>
        <tbody>
          {usersHere.map((u) => (
            <tr key={u.id} className="border-b border-slate-100">
              <td className="py-2.5 font-medium text-slate-700">{u.name}</td>
              <td className="py-2.5 mono text-slate-500">{u.username}</td>
              <td className="py-2.5">
                {isAdmin ? (
                  <select className="text-xs border border-slate-300 rounded px-2 py-1" value={u.role} onChange={(e) => changeRole(u, e.target.value)}>
                    {ROLES.map((r) => <option key={r}>{r}</option>)}
                  </select>
                ) : <Badge>{u.role}</Badge>}
              </td>
              {isAdmin && <td className="py-2.5 text-right"><button onClick={() => removeUser(u)} className="p-1 rounded hover:bg-slate-100 text-slate-400"><Trash2 size={14} /></button></td>}
            </tr>
          ))}
        </tbody>
      </table>
      {!isAdmin && <div className="text-xs text-slate-400 flex items-center gap-2 mt-3"><AlertTriangle size={13} /> Only Super Admin can manage users and roles.</div>}
    </div>
  );
}

function ActivityLogTab({ data }) {
  const log = data.activityLog || [];
  return (
    <div className="bg-white rounded-lg border border-slate-200 overflow-x-auto">
      {log.length === 0 ? <EmptyState icon={ClipboardList} title="No activity recorded yet" /> : (
        <table className="w-full text-sm min-w-[600px]">
          <thead><tr className="text-left text-[11px] uppercase tracking-wide text-slate-400 border-b border-slate-200">
            <th className="px-4 py-2.5">Date / Time</th><th className="px-4 py-2.5">User</th><th className="px-4 py-2.5">Action</th><th className="px-4 py-2.5">Detail</th>
          </tr></thead>
          <tbody>
            {log.map((e) => (
              <tr key={e.id} className="border-b border-slate-100 hover:bg-slate-50">
                <td className="px-4 py-2.5 text-slate-500 mono text-xs">{fmtDateTime(e.ts)}</td>
                <td className="px-4 py-2.5 text-slate-700">{e.user}</td>
                <td className="px-4 py-2.5"><Badge tone="blue">{e.action}</Badge></td>
                <td className="px-4 py-2.5 text-slate-500">{e.detail}</td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  );
}
